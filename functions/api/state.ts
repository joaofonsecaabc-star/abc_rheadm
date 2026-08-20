import { currentUser } from './auth/_utils'

interface Env { DB: any }
type Context = { request: Request; env: Env }
type User = { id:number; role:string; modules:string; store_access:string }

const json = (data: unknown, status = 200) => Response.json(data, { status, headers: { 'Cache-Control': 'no-store' } })
const parseRows = (result: any) => (result?.results || []).map((row: any) => JSON.parse(row.data_json))
const csv = (value: unknown) => String(value || '').split(',').map(item => item.trim()).filter(Boolean)
const moduleAllowed = (user: User, module: string) => user.role === 'admin' || csv(user.modules).includes(module)
const storesFor = (user: User) => csv(user.store_access || '*')
const storeAllowed = (user: User, store: string) => storesFor(user).includes('*') || storesFor(user).includes(store)

const scopedQuery = (base: string, user: User, column = 'store_name') => {
  const stores = storesFor(user)
  if (stores.includes('*')) return { sql: base, binds: [] as string[] }
  if (!stores.length) return { sql: `${base} WHERE 1=0`, binds: [] as string[] }
  return { sql: `${base} WHERE ${column} IN (${stores.map(() => '?').join(',')})`, binds: stores }
}

const onRequestGet = async ({ env }: Context, user: User) => {
  try {
    const employeeQuery = scopedQuery('SELECT data_json FROM employees', user)
    const employees = await env.DB.prepare(`${employeeQuery.sql} ORDER BY employee_name`).bind(...employeeQuery.binds).all()
    const employeeIds = (employees.results || []).map((row: any) => Number(JSON.parse(row.data_json).id)).filter(Boolean)
    const idClause = employeeIds.length ? `IN (${employeeIds.map(() => '?').join(',')})` : 'IN (NULL)'
    const canTransit = moduleAllowed(user, 'transit')
    const canPeople = moduleAllowed(user, 'people')
    const canFinance = moduleAllowed(user, 'finance')
    const canAdministrative = moduleAllowed(user, 'administrative')
    const [events, occurrences, stores, positions, reasons, settings, revision] = await env.DB.batch([
      canTransit
        ? env.DB.prepare(`SELECT data_json FROM recharge_events WHERE employee_id ${idClause} ORDER BY completed_date DESC`).bind(...employeeIds)
        : env.DB.prepare('SELECT data_json FROM recharge_events WHERE 1=0'),
      canPeople || canFinance
        ? env.DB.prepare(`SELECT data_json FROM hr_occurrences WHERE employee_id ${idClause} ORDER BY occurrence_date DESC`).bind(...employeeIds)
        : env.DB.prepare('SELECT data_json FROM hr_occurrences WHERE 1=0'),
      storesFor(user).includes('*')
        ? env.DB.prepare('SELECT name FROM stores WHERE active=1 ORDER BY name')
        : env.DB.prepare(`SELECT name FROM stores WHERE active=1 AND name IN (${storesFor(user).map(() => '?').join(',') || "''"}) ORDER BY name`).bind(...storesFor(user)),
      env.DB.prepare('SELECT name FROM positions WHERE active=1 ORDER BY name'),
      canPeople ? env.DB.prepare('SELECT description FROM unregistered_reasons WHERE active=1 ORDER BY description') : env.DB.prepare('SELECT description FROM unregistered_reasons WHERE 1=0'),
      canFinance || user.role === 'admin'
        ? env.DB.prepare("SELECT key,value_json FROM app_settings WHERE key<>'state_revision'")
        : canAdministrative
          ? env.DB.prepare("SELECT key,value_json FROM app_settings WHERE key IN ('taxCompanies','companyCnpjs')")
          : env.DB.prepare("SELECT key,value_json FROM app_settings WHERE key IN ('advanceDays','advance_days')"),
      env.DB.prepare("SELECT value_json FROM app_settings WHERE key='state_revision'")
    ])
    const settingValues = Object.fromEntries((settings.results || []).map((row: any) => [row.key, JSON.parse(row.value_json)])) as Record<string,unknown>
    if (Array.isArray(settingValues.financialEntries))
      settingValues.financialEntries = settingValues.financialEntries.filter((entry: any) => employeeIds.includes(Number(entry.employeeId)))
    return json({
      employees: parseRows(employees), events: parseRows(events), occurrences: parseRows(occurrences),
      stores: (stores.results || []).map((row: any) => row.name),
      positions: (positions.results || []).map((row: any) => row.name),
      unregisteredReasons: (reasons.results || []).map((row: any) => row.description),
      settings: settingValues,
      revision: Number(revision.results?.[0]?.value_json || 0)
    })
  } catch (error) {
    return json({ error: 'Não foi possível carregar o banco D1.', detail: String(error) }, 500)
  }
}

const employeeUpsert = (env: Env, employee: any) => env.DB.prepare(`INSERT INTO employees(
  id,employee_name,cpf,role_name,store_name,phone,hired_at,birth_date,active,employment_status,
  formal_employment,unregistered_start_date,unregistered_reason,experience_days,experience_critical,
  notice_start,notice_end,termination_date,receives_transit,receives_cost_assistance,cost_assistance_amount,
  card_type,card_daily_fare,second_card_type,second_card_daily_fare,recharge_day,advance_days,
  schedule_type,schedule_start_date,work_days_json,data_json
) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
ON CONFLICT(id) DO UPDATE SET employee_name=excluded.employee_name,cpf=excluded.cpf,role_name=excluded.role_name,
  store_name=excluded.store_name,phone=excluded.phone,hired_at=excluded.hired_at,birth_date=excluded.birth_date,
  active=excluded.active,employment_status=excluded.employment_status,formal_employment=excluded.formal_employment,
  unregistered_start_date=excluded.unregistered_start_date,unregistered_reason=excluded.unregistered_reason,
  experience_days=excluded.experience_days,experience_critical=excluded.experience_critical,
  notice_start=excluded.notice_start,notice_end=excluded.notice_end,termination_date=excluded.termination_date,
  receives_transit=excluded.receives_transit,receives_cost_assistance=excluded.receives_cost_assistance,
  cost_assistance_amount=excluded.cost_assistance_amount,card_type=excluded.card_type,
  card_daily_fare=excluded.card_daily_fare,second_card_type=excluded.second_card_type,
  second_card_daily_fare=excluded.second_card_daily_fare,recharge_day=excluded.recharge_day,
  advance_days=excluded.advance_days,schedule_type=excluded.schedule_type,schedule_start_date=excluded.schedule_start_date,
  work_days_json=excluded.work_days_json,data_json=excluded.data_json,updated_at=CURRENT_TIMESTAMP`).bind(
  employee.id,employee.employee,employee.cpf || null,employee.role || '',employee.store || '',employee.phone || null,
  employee.hiredAt || null,employee.birthDate || null,employee.active === false ? 0 : 1,employee.employmentStatus || 'Ativo',
  employee.formalEmployment === false ? 0 : 1,employee.unregisteredStartDate || null,employee.unregisteredReason || null,
  employee.experienceDays || 90,employee.experienceCritical ? 1 : 0,employee.noticeStart || null,employee.noticeEnd || null,
  employee.terminationDate || null,employee.receivesTransit === false ? 0 : 1,employee.receivesCostAssistance ? 1 : 0,
  Number(employee.costAssistanceAmount || 0),employee.cardType || null,Number(employee.cardDailyFare || employee.dailyFare || 0),
  employee.secondCardType || null,Number(employee.secondCardDailyFare || 0),employee.rechargeDate ? new Date(employee.rechargeDate + 'T12:00:00').getDate() : null,
  Number(employee.advance || 3),employee.scheduleType || 'Personalizada',employee.scheduleStartDate || employee.hiredAt || null,
  JSON.stringify(employee.workDays || [1,2,3,4,5]),JSON.stringify(employee))

const onRequestPut = async ({ request, env }: Context, user: User) => {
  try {
    const body: any = await request.json()
    const currentRevisionRow = await env.DB.prepare("SELECT value_json FROM app_settings WHERE key='state_revision'").first()
    const currentRevision = Number(currentRevisionRow?.value_json || 0)
    if (Number(body.revision ?? -1) !== currentRevision)
      return json({ error: 'Os dados foram alterados em outra aba. Atualize a página antes de salvar.', code: 'STATE_CONFLICT', revision: currentRevision }, 409)

    const employees = Array.isArray(body.employees) ? body.employees : []
    const events = Array.isArray(body.events) ? body.events : []
    const occurrences = Array.isArray(body.occurrences) ? body.occurrences : []
    const employeeIds = employees.map((employee: any) => Number(employee.id)).filter(Boolean)
    const allowedEmployeeIds = new Set(employeeIds)
    if (employees.some((employee: any) => !storeAllowed(user, String(employee.store || ''))))
      return json({ error: 'Há funcionário de uma loja sem permissão nesta solicitação.' }, 403)
    if ([...events, ...occurrences].some((item: any) => !allowedEmployeeIds.has(Number(item.employeeId))))
      return json({ error: 'A solicitação contém dados de funcionário sem permissão.' }, 403)

    const statements: any[] = []
    if (user.role === 'admin') {
      for (const name of Array.isArray(body.stores) ? body.stores : []) statements.push(env.DB.prepare('INSERT OR IGNORE INTO stores(name) VALUES (?)').bind(String(name)))
      for (const name of Array.isArray(body.positions) ? body.positions : []) statements.push(env.DB.prepare('INSERT OR IGNORE INTO positions(name) VALUES (?)').bind(String(name)))
      for (const description of Array.isArray(body.unregisteredReasons) ? body.unregisteredReasons : []) statements.push(env.DB.prepare('INSERT OR IGNORE INTO unregistered_reasons(description) VALUES (?)').bind(String(description)))
      for (const employee of employees) statements.push(employeeUpsert(env, employee))
      if (body.settings) for (const [key, value] of Object.entries(body.settings)) statements.push(env.DB.prepare(`INSERT INTO app_settings(key,value_json,updated_at) VALUES (?,?,CURRENT_TIMESTAMP)
        ON CONFLICT(key) DO UPDATE SET value_json=excluded.value_json,updated_at=CURRENT_TIMESTAMP`).bind(key,JSON.stringify(value)))
    } else if (moduleAllowed(user, 'finance') && body.settings) {
      if (Array.isArray(body.settings.financialEntries)) {
        const incomingEntries = body.settings.financialEntries
        if (incomingEntries.some((entry: any) => !allowedEmployeeIds.has(Number(entry.employeeId))))
          return json({ error: 'Há lançamento financeiro de funcionário sem permissão.' }, 403)
        const savedSetting = await env.DB.prepare("SELECT value_json FROM app_settings WHERE key='financialEntries'").first()
        let savedEntries: any[] = []
        try { savedEntries = JSON.parse(savedSetting?.value_json || '[]') } catch { savedEntries = [] }
        const mergedEntries = [
          ...savedEntries.filter((entry: any) => !allowedEmployeeIds.has(Number(entry.employeeId))),
          ...incomingEntries,
        ]
        statements.push(env.DB.prepare(`INSERT INTO app_settings(key,value_json,updated_at) VALUES ('financialEntries',?,CURRENT_TIMESTAMP)
          ON CONFLICT(key) DO UPDATE SET value_json=excluded.value_json,updated_at=CURRENT_TIMESTAMP`).bind(JSON.stringify(mergedEntries)))
      }
      if (Array.isArray(body.settings.taxEntries))
        statements.push(env.DB.prepare(`INSERT INTO app_settings(key,value_json,updated_at) VALUES ('taxEntries',?,CURRENT_TIMESTAMP)
          ON CONFLICT(key) DO UPDATE SET value_json=excluded.value_json,updated_at=CURRENT_TIMESTAMP`).bind(JSON.stringify(body.settings.taxEntries)))
    }
    if (moduleAllowed(user, 'transit')) {
      if (employeeIds.length) statements.push(env.DB.prepare(`DELETE FROM recharge_events WHERE employee_id IN (${employeeIds.map(() => '?').join(',')})`).bind(...employeeIds))
      for (const event of events) statements.push(env.DB.prepare(`INSERT INTO recharge_events(
        id,employee_id,period,scheduled_date,completed_date,charged_days,card_amount,second_card_amount,total_amount,responsible_user_id,data_json
      ) VALUES (?,?,?,?,?,?,?,?,?,?,?)`).bind(event.id,event.employeeId,event.period,event.scheduledDate,event.completedDate,event.chargedDays ?? null,event.cardAmount ?? null,event.secondCardAmount ?? null,event.totalAmount ?? null,user.id,JSON.stringify(event)))
    }
    if (user.role === 'admin' && moduleAllowed(user, 'people')) {
      if (employeeIds.length) statements.push(env.DB.prepare(`DELETE FROM hr_occurrences WHERE employee_id IN (${employeeIds.map(() => '?').join(',')})`).bind(...employeeIds))
      for (const occurrence of occurrences) statements.push(env.DB.prepare(`INSERT INTO hr_occurrences(
        id,employee_id,occurrence_date,end_date,type,hours,minutes,days,note,data_json
      ) VALUES (?,?,?,?,?,?,?,?,?,?)`).bind(occurrence.id,occurrence.employeeId,occurrence.date,occurrence.endDate || null,occurrence.type,occurrence.hours || 0,occurrence.minutes || 0,occurrence.days || 0,occurrence.note || null,JSON.stringify(occurrence)))
    }
    const nextRevision = currentRevision + 1
    statements.push(env.DB.prepare(`INSERT INTO app_settings(key,value_json,updated_at) VALUES ('state_revision',?,CURRENT_TIMESTAMP)
      ON CONFLICT(key) DO UPDATE SET value_json=excluded.value_json,updated_at=CURRENT_TIMESTAMP`).bind(String(nextRevision)))
    statements.push(env.DB.prepare('INSERT INTO audit_logs(user_id,action,entity_type,entity_id,after_json) VALUES (?,?,?,?,?)').bind(user.id,'sync','state',String(nextRevision),JSON.stringify({employees:employees.length,events:events.length,occurrences:occurrences.length})))
    await env.DB.batch(statements)
    return json({ ok: true, savedAt: new Date().toISOString(), revision: nextRevision })
  } catch (error) {
    return json({ error: 'Não foi possível salvar no banco D1.', detail: String(error) }, 500)
  }
}

export const onRequest = async (context: Context) => {
  const user = await currentUser(context.request, context.env.DB) as User | null
  if (!user) return json({ error: 'Sessão expirada.' }, 401)
  if (context.request.method === 'GET') return onRequestGet(context, user)
  if (context.request.method === 'PUT') return onRequestPut(context, user)
  return json({ error: 'Método não permitido.' }, 405)
}
