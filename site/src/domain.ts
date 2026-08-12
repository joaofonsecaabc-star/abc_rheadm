import type { Recharge, Status } from './data'

export type RechargeEvent = {
  id: number
  employeeId: number
  period: string
  scheduledDate: string
  completedDate: string
  chargedDays?: number
  cardAmount?: number
  secondCardAmount?: number
  totalAmount?: number
}

export type LocalDatabase = { employees: Recharge[]; events: RechargeEvent[] }

const identity = (r: Recharge) => r.cpf?.replace(/\D/g, '') || `${r.employee.trim().toLowerCase()}|${r.store.trim().toLowerCase()}`
const iso = (d: Date) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
const isWorkDay = (employee: Recharge, date: Date) => {
  if (employee.scheduleType === '12x36') {
    const anchor = new Date((employee.scheduleStartDate || employee.hiredAt || iso(date)) + 'T12:00:00')
    const difference = Math.floor((date.getTime() - anchor.getTime()) / 86400000)
    return difference >= 0 && difference % 2 === 0
  }
  return (employee.workDays || [1, 2, 3, 4, 5]).includes(date.getDay())
}

export function loadDatabase(): LocalDatabase {
  try {
    const saved = localStorage.getItem('valefluxo_db_v3')
    if (saved) return JSON.parse(saved)
    const legacy: Recharge[] = JSON.parse(localStorage.getItem('valefluxo_rows_v2') || '[]')
    const grouped = new Map<string, Recharge[]>()
    legacy.forEach(row => grouped.set(identity(row), [...(grouped.get(identity(row)) || []), row]))
    const employees: Recharge[] = []
    const events: RechargeEvent[] = []
    for (const group of grouped.values()) {
      const leaf = group.filter(row => !group.some(next => next.sourceRechargeId === row.id)).sort((a, b) => b.id - a.id)[0] || group[0]
      employees.push({ ...leaf, sourceRechargeId: undefined, periodCompletionId: undefined, completedDate: undefined, status: 'Próximo' })
      group.filter(row => row.completedDate).forEach(row => {
        const period = row.completedDate!.slice(0, 7)
        if (!events.some(event => event.employeeId === leaf.id && event.period === period)) events.push({ id: row.id, employeeId: leaf.id, period, scheduledDate: row.rechargeDate, completedDate: row.completedDate! })
      })
    }
    const db = { employees, events }
    localStorage.setItem('valefluxo_db_v3', JSON.stringify(db))
    return db
  } catch { return { employees: [], events: [] } }
}

export function monthlyRows(employees: Recharge[], events: RechargeEvent[], period: string, referenceDate: string): Recharge[] {
  const [year, month] = period.split('-').map(Number)
  return employees.filter(employee => employee.receivesTransit !== false && (employee.terminationDate ? employee.terminationDate > referenceDate : employee.active !== false && employee.employmentStatus !== 'Desligado')).map(employee => {
    const original = new Date(employee.rechargeDate + 'T12:00:00')
    const day = Math.min(original.getDate(), new Date(year, month, 0).getDate())
    const scheduled = new Date(year, month - 1, day, 12)
    const credit = new Date(scheduled); credit.setDate(credit.getDate() + employee.advance)
    const rechargeDate = iso(scheduled)
    const event = events.filter(item => item.employeeId === employee.id && item.period === period).sort((a,b)=>b.completedDate.localeCompare(a.completedDate)||b.id-a.id)[0]
    // O planejamento mensal continua no dia fixo, mas o crédito de uma recarga
    // concluída é liberado a partir da data real informada na baixa.
    const effectiveCredit = event ? new Date(event.completedDate + 'T12:00:00') : credit
    if (event) effectiveCredit.setDate(effectiveCredit.getDate() + employee.advance)
    const creditDate = iso(effectiveCredit)
    let status: Status
    if (event) status = event.completedDate > rechargeDate ? 'Recarregado atrasado' : 'Recarregado'
    else status = rechargeDate < referenceDate ? 'Atrasado' : rechargeDate === referenceDate ? 'Pendente' : 'Próximo'
    const end = new Date(credit); end.setMonth(end.getMonth() + 1)
    let days = 0
    for (const d = new Date(credit); d < end; d.setDate(d.getDate() + 1)) if (isWorkDay(employee, d)) days++
    const cardAmount = days * (employee.cardDailyFare ?? employee.dailyFare ?? 0)
    const secondCardAmount = days * (employee.secondCardDailyFare ?? 0)
    return { ...employee, rechargeDate, creditDate, status, completedDate: event?.completedDate, periodCompletionId: event?.id, chargedDays:event?.chargedDays, cardAmount:event?.cardAmount??cardAmount, secondCardAmount:event?.secondCardAmount??(employee.secondCardType ? secondCardAmount : undefined), amount:event?.totalAmount??(cardAmount + secondCardAmount) }
  })
}
