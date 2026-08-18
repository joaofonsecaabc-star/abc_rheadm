import { passwordHash, randomHex } from './_utils'

const json = (data: unknown, status = 200, headers: Record<string,string> = {}) =>
  Response.json(data, { status, headers: { 'Cache-Control': 'no-store', ...headers } })

export const onRequest = async ({request,env}:{request:Request;env:any}) => {
  if (request.method !== 'POST') return json({error:'Método não permitido.'},405)
  let body:any
  try { body=await request.json() } catch { return json({error:'Dados inválidos.'},400) }
  const username=String(body.username||'').trim()
  const password=String(body.password||'')
  const ip=request.headers.get('CF-Connecting-IP')||'unknown'
  const attemptKey=`${ip}|${username.toLocaleLowerCase()}`
  const attempt=await env.DB.prepare('SELECT attempts,blocked_until FROM login_attempts WHERE attempt_key=?').bind(attemptKey).first()
  if (attempt?.blocked_until && new Date(attempt.blocked_until).getTime()>Date.now())
    return json({error:'Muitas tentativas. Aguarde 15 minutos antes de tentar novamente.'},429,{'Retry-After':'900'})

  const user=await env.DB.prepare('SELECT * FROM users WHERE username=? COLLATE NOCASE AND active=1').bind(username).first()
  const valid=Boolean(user) && await passwordHash(password,user.password_salt)===user.password_hash
  if (!valid) {
    const attempts=Number(attempt?.attempts||0)+1
    const blockedUntil=attempts>=5?new Date(Date.now()+15*60*1000).toISOString():null
    await env.DB.prepare(`INSERT INTO login_attempts(attempt_key,attempts,blocked_until,updated_at) VALUES (?,?,?,CURRENT_TIMESTAMP)
      ON CONFLICT(attempt_key) DO UPDATE SET attempts=excluded.attempts,blocked_until=excluded.blocked_until,updated_at=CURRENT_TIMESTAMP`)
      .bind(attemptKey,attempts,blockedUntil).run()
    return json({error:attempts>=5?'Muitas tentativas. Acesso bloqueado por 15 minutos.':'Usuário ou senha inválidos.'},401)
  }

  const session=randomHex(32),expires=new Date(Date.now()+12*60*60*1000).toISOString()
  await env.DB.batch([
    env.DB.prepare('DELETE FROM login_attempts WHERE attempt_key=?').bind(attemptKey),
    env.DB.prepare('DELETE FROM sessions WHERE expires_at<=CURRENT_TIMESTAMP'),
    env.DB.prepare('INSERT INTO sessions(id,user_id,expires_at) VALUES (?,?,?)').bind(session,user.id,expires),
    env.DB.prepare('UPDATE users SET last_login_at=CURRENT_TIMESTAMP WHERE id=?').bind(user.id)
  ])
  return new Response(JSON.stringify({ok:true,user:{id:user.id,username:user.username,fullName:user.full_name,role:user.role,modules:user.modules,storeAccess:user.store_access}}),{
    headers:{'Content-Type':'application/json','Cache-Control':'no-store','Set-Cookie':`abc_session=${session}; Path=/; HttpOnly; Secure; SameSite=Strict; Max-Age=43200`}
  })
}
