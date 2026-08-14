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

export type ModuleAccess = 'people'|'finance'|'transit'
export type CloudUser = { id:number; username:string; fullName:string; role:'admin'|'operator'; active:boolean; modules:ModuleAccess[]; storeAccess:string[]; password?:string; lastLoginAt?:string|null; createdAt?:string }
export type SessionUser = { id:number; username:string; fullName:string; role:'admin'|'operator'; modules:ModuleAccess[]; storeAccess:string[] }
const normalizeModules=(value:unknown):ModuleAccess[]=>{
  const values=Array.isArray(value)?value:String(value||'people,finance,transit').split(',')
  return values.filter((item):item is ModuleAccess=>item==='people'||item==='finance'||item==='transit')
}
const normalizeStores=(value:unknown):string[]=>{
  const values=Array.isArray(value)?value:String(value||'*').split(',')
  const normalized=values.map(String).map(item=>item.trim()).filter(Boolean)
  return normalized.includes('*')?['*']:normalized
}

async function apiJson<T>(url:string, options?:RequestInit):Promise<T>{
  const response=await fetch(url,options)
  const contentType=response.headers.get('content-type')||''
  const raw=await response.text()
  if(!contentType.includes('application/json'))throw new Error('A API do sistema ainda não está disponível nesta versão. Atualize a página com Ctrl+F5.')
  let data:T & {error?:string}
  try{data=JSON.parse(raw) as T & {error?:string}}catch{throw new Error('O servidor respondeu em um formato inválido. Atualize a página e tente novamente.')}
  if(!response.ok)throw new Error(data.error||'Não foi possível concluir a operação.')
  return data
}

export async function cloudSetupRequired(){if(!cloudEnabled())return false;return (await apiJson<{needsSetup:boolean}>('/api/auth/status')).needsSetup}
export async function cloudCreateAdmin(fullName:string,username:string,password:string){return apiJson('/api/auth/setup',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({fullName,username,password})})}
export async function listCloudUsers(){if(!cloudEnabled())return [];return (await apiJson<{users:Array<Omit<CloudUser,'modules'|'storeAccess'>&{modules:unknown;storeAccess:unknown}>}>('/api/users')).users.map(user=>({...user,active:Boolean(user.active),modules:normalizeModules(user.modules),storeAccess:normalizeStores(user.storeAccess)}))}
export async function createCloudUser(user:Omit<CloudUser,'id'>){return apiJson('/api/users',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(user)})}
export async function updateCloudUser(user:CloudUser){return apiJson('/api/users',{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(user)})}

export async function cloudLogin(username:string,password:string){
  if(!cloudEnabled())return {ok:true}
  return apiJson<{ok?:boolean;error?:string}>('/api/auth/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username,password})})
}
export async function cloudSession(){if(!cloudEnabled())return true;return (await fetch('/api/auth/session')).ok}
export async function cloudCurrentUser():Promise<SessionUser|null>{
  if(!cloudEnabled())return {id:0,username:'local',fullName:'João Fonseca',role:'admin',modules:['people','finance','transit'],storeAccess:['*']}
  const response=await fetch('/api/auth/session',{headers:{Accept:'application/json'}})
  if(!response.ok)return null
  const data=await response.json() as {user?:{id:number;username:string;fullName?:string;full_name?:string;role:'admin'|'operator';modules?:unknown;storeAccess?:unknown;store_access?:unknown}}
  if(!data.user)return null
  return {id:data.user.id,username:data.user.username,fullName:data.user.fullName||data.user.full_name||data.user.username,role:data.user.role,modules:normalizeModules(data.user.modules),storeAccess:normalizeStores(data.user.storeAccess||data.user.store_access)}
}
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
