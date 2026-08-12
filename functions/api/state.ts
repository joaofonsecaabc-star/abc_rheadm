import { currentUser } from './auth/_utils'

interface Env { DB: any }
type Context = { request: Request; env: Env }

const json = (data: unknown, status = 200) => Response.json(data, { status, headers: { 'Cache-Control': 'no-store' } })
const parseRows = (result: any) => (result?.results || []).map((row: any) => JSON.parse(row.data_json))

const onRequestGet = async ({ env }: Context) => {
  try {
    const [employees, events, occurrences, stores, positions, reasons, settings] = await env.DB.batch([
      env.DB.prepare('SELECT data_json FROM employees ORDER BY employee_name'),
      env.DB.prepare('SELECT data_json FROM recharge_events ORDER BY completed_date DESC'),
      env.DB.prepare('SELECT data_json FROM hr_occurrences ORDER BY occurrence_date DESC'),
      env.DB.prepare('SELECT name FROM stores WHERE active = 1 ORDER BY name'),
      env.DB.prepare('SELECT name FROM positions WHERE active = 1 ORDER BY name'),
      env.DB.prepare('SELECT description FROM unregistered_reasons WHERE active = 1 ORDER BY description'),
      env.DB.prepare('SELECT key, value_json FROM app_settings')
    ])
    return json({
      employees: parseRows(employees),
      events: parseRows(events),
      occurrences: parseRows(occurrences),
      stores: (stores.results || []).map((row: any) => row.name),
      positions: (positions.results || []).map((row: any) => row.name),
      unregisteredReasons: (reasons.results || []).map((row: any) => row.description),
      settings: Object.fromEntries((settings.results || []).map((row: any) => [row.key, JSON.parse(row.value_json)]))
    })
  } catch (error) {
    return json({ error: 'Não foi possível carregar o banco D1.', detail: String(error) }, 500)
  }
}

const onRequestPut = async ({ request, env }: Context) => {
  try {
    const body: any = await request.json()
    const employees = Array.isArray(body.employees) ? body.employees : []
    const events = Array.isArray(body.events) ? body.events : []
    const occurrences = Array.isArray(body.occurrences) ? body.occurrences : []
    const stores = Array.isArray(body.stores) ? body.stores : []
    const positions = Array.isArray(body.positions) ? body.positions : []
    const reasons = Array.isArray(body.unregisteredReasons) ? body.unregisteredReasons : []
    const statements = [
      env.DB.prepare('DELETE FROM recharge_events'),
      env.DB.prepare('DELETE FROM hr_occurrences'),
      env.DB.prepare('DELETE FROM employees'),
      env.DB.prepare('DELETE FROM stores'),
      env.DB.prepare('DELETE FROM positions'),
      env.DB.prepare('DELETE FROM unregistered_reasons')
    ]
    for (const name of stores) statements.push(env.DB.prepare('INSERT INTO stores(name) VALUES (?)').bind(String(name)))
    for (const name of positions) statements.push(env.DB.prepare('INSERT INTO positions(name) VALUES (?)').bind(String(name)))
    for (const description of reasons) statements.push(env.DB.prepare('INSERT INTO unregistered_reasons(description) VALUES (?)').bind(String(description)))
    for (const employee of employees) statements.push(env.DB.prepare(`INSERT INTO employees(
      id,employee_name,cpf,role_name,store_name,phone,hired_at,birth_date,active,employment_status,
      formal_employment,unregistered_start_date,unregistered_reason,experience_days,experience_critical,
      notice_start,notice_end,termination_date,receives_transit,receives_cost_assistance,cost_assistance_amount,
      card_type,card_daily_fare,second_card_type,second_card_daily_fare,recharge_day,advance_days,
      schedule_type,schedule_start_date,work_days_json,data_json
    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`).bind(
      employee.id, employee.employee, employee.cpf || null, employee.role || '', employee.store || '', employee.phone || null,
      employee.hiredAt || null, employee.birthDate || null, employee.active === false ? 0 : 1, employee.employmentStatus || 'Ativo',
      employee.formalEmployment === false ? 0 : 1, employee.unregisteredStartDate || null, employee.unregisteredReason || null,
      employee.experienceDays || 90, employee.experienceCritical ? 1 : 0, employee.noticeStart || null, employee.noticeEnd || null,
      employee.terminationDate || null, employee.receivesTransit === false ? 0 : 1, employee.receivesCostAssistance ? 1 : 0,
      Number(employee.costAssistanceAmount || 0), employee.cardType || null, Number(employee.cardDailyFare || employee.dailyFare || 0),
      employee.secondCardType || null, Number(employee.secondCardDailyFare || 0), employee.rechargeDate ? new Date(employee.rechargeDate + 'T12:00:00').getDate() : null,
      Number(employee.advance || 3), employee.scheduleType || 'Personalizada', employee.scheduleStartDate || employee.hiredAt || null,
      JSON.stringify(employee.workDays || [1,2,3,4,5]), JSON.stringify(employee)
    ))
    for (const event of events) statements.push(env.DB.prepare(`INSERT INTO recharge_events(
      id,employee_id,period,scheduled_date,completed_date,charged_days,card_amount,second_card_amount,total_amount,data_json
    ) VALUES (?,?,?,?,?,?,?,?,?,?)`).bind(event.id,event.employeeId,event.period,event.scheduledDate,event.completedDate,event.chargedDays ?? null,event.cardAmount ?? null,event.secondCardAmount ?? null,event.totalAmount ?? null,JSON.stringify(event)))
    for (const occurrence of occurrences) statements.push(env.DB.prepare(`INSERT INTO hr_occurrences(
      id,employee_id,occurrence_date,end_date,type,hours,minutes,days,note,data_json
    ) VALUES (?,?,?,?,?,?,?,?,?,?)`).bind(occurrence.id,occurrence.employeeId,occurrence.date,occurrence.endDate || null,occurrence.type,occurrence.hours || 0,occurrence.minutes || 0,occurrence.days || 0,occurrence.note || null,JSON.stringify(occurrence)))
    if (body.settings) for (const [key, value] of Object.entries(body.settings)) statements.push(env.DB.prepare(`INSERT INTO app_settings(key,value_json,updated_at) VALUES (?,?,CURRENT_TIMESTAMP)
      ON CONFLICT(key) DO UPDATE SET value_json=excluded.value_json, updated_at=CURRENT_TIMESTAMP`).bind(key,JSON.stringify(value)))
    await env.DB.batch(statements)
    return json({ ok: true, savedAt: new Date().toISOString() })
  } catch (error) {
    return json({ error: 'Não foi possível salvar no banco D1.', detail: String(error) }, 500)
  }
}

export const onRequest = async (context: Context) => {
  if (!await currentUser(context.request, context.env.DB)) return json({ error: 'Sessão expirada.' }, 401)
  if (context.request.method === 'GET') return onRequestGet(context)
  if (context.request.method === 'PUT') return onRequestPut(context)
  return json({ error: 'Método não permitido.' }, 405)
}
