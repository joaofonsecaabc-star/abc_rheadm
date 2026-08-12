import type { Recharge } from './data'
import type { RechargeEvent } from './domain'

export type CloudState = {
  employees: Recharge[]
  events: RechargeEvent[]
  occurrences: unknown[]
  stores: string[]
  positions: string[]
  unregisteredReasons: string[]
  settings?: Record<string, unknown>
}

export const cloudEnabled = () => !['localhost','127.0.0.1'].includes(location.hostname)

export async function cloudLogin(username:string,password:string){
  if(!cloudEnabled())return {ok:true}
  const response=await fetch('/api/auth/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username,password})})
  const data=await response.json() as {ok?:boolean;error?:string}
  if(!response.ok)throw new Error(data.error||'Falha ao entrar.')
  return data
}
export async function cloudSession(){if(!cloudEnabled())return true;return (await fetch('/api/auth/session')).ok}
export async function cloudLogout(){if(cloudEnabled())await fetch('/api/auth/logout',{method:'POST'})}

export async function loadCloudState(): Promise<CloudState | null> {
  if (!cloudEnabled()) return null
  const response = await fetch('/api/state', { headers: { Accept: 'application/json' } })
  if (!response.ok) throw new Error(`D1 respondeu ${response.status}`)
  return response.json()
}

export async function saveCloudState(state: CloudState): Promise<void> {
  if (!cloudEnabled()) return
  const response = await fetch('/api/state', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(state)
  })
  if (!response.ok) throw new Error(`D1 respondeu ${response.status}`)
}
