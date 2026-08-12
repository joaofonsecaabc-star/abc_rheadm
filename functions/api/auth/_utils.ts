export const hex = (bytes: ArrayBuffer) => [...new Uint8Array(bytes)].map(value => value.toString(16).padStart(2,'0')).join('')
export const randomHex = (length = 32) => { const bytes = new Uint8Array(length); crypto.getRandomValues(bytes); return [...bytes].map(value => value.toString(16).padStart(2,'0')).join('') }
export async function passwordHash(password: string, salt: string) {
  const key = await crypto.subtle.importKey('raw',new TextEncoder().encode(password),'PBKDF2',false,['deriveBits'])
  return hex(await crypto.subtle.deriveBits({name:'PBKDF2',hash:'SHA-256',salt:new TextEncoder().encode(salt),iterations:150000},key,256))
}
export const cookieValue = (request: Request, name: string) => request.headers.get('Cookie')?.split(';').map(item=>item.trim()).find(item=>item.startsWith(name+'='))?.slice(name.length+1)
export async function currentUser(request: Request, db: any) {
  const session = cookieValue(request,'abc_session'); if(!session)return null
  return db.prepare(`SELECT users.id,users.username,users.full_name,users.role FROM sessions JOIN users ON users.id=sessions.user_id WHERE sessions.id=? AND sessions.expires_at>CURRENT_TIMESTAMP AND users.active=1`).bind(session).first()
}
