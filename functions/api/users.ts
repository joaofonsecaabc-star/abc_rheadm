import { currentUser, passwordHash, randomHex } from './auth/_utils'

const json = (data: unknown, status = 200) => Response.json(data, { status })

export const onRequest = async ({ request, env }: { request: Request; env: any }) => {
  const current = await currentUser(request, env.DB)
  if (!current) return json({ error: 'Sessão expirada.' }, 401)
  if (current.role !== 'admin') return json({ error: 'Somente administradores podem gerenciar perfis.' }, 403)

  if (request.method === 'GET') {
    const result = await env.DB.prepare(`SELECT id, username, full_name AS fullName, role, active, modules,
      last_login_at AS lastLoginAt, created_at AS createdAt FROM users ORDER BY full_name`).all()
    return json({ users: result.results || [] })
  }

  const body: any = await request.json()
  const username = String(body.username || '').trim()
  const fullName = String(body.fullName || '').trim()
  const role = body.role === 'admin' ? 'admin' : 'operator'
  const validModules = ['people', 'finance', 'transit']
  const modules = [...new Set((Array.isArray(body.modules) ? body.modules : []).filter((item: unknown) => validModules.includes(String(item))))]
  if (!modules.length) return json({ error: 'Selecione pelo menos uma área para o usuário.' }, 400)
  const modulesValue = modules.join(',')

  try {
    if (request.method === 'POST') {
      const password = String(body.password || '')
      if (!username || !fullName || password.length < 8) return json({ error: 'Informe nome, usuário e senha com pelo menos 8 caracteres.' }, 400)
      const salt = randomHex(16)
      const hash = await passwordHash(password, salt)
      await env.DB.prepare(`INSERT INTO users(username,full_name,password_hash,password_salt,role,active,modules)
        VALUES (?,?,?,?,?,1,?)`).bind(username, fullName, hash, salt, role, modulesValue).run()
      return json({ ok: true }, 201)
    }

    if (request.method === 'PUT') {
      const id = Number(body.id)
      if (!id || !username || !fullName) return json({ error: 'Dados do perfil incompletos.' }, 400)
      if (id === Number(current.id) && body.active === false) return json({ error: 'Você não pode desativar o próprio perfil.' }, 400)
      const statements = [env.DB.prepare(`UPDATE users SET username=?, full_name=?, role=?, active=?, modules=?, updated_at=CURRENT_TIMESTAMP WHERE id=?`)
        .bind(username, fullName, role, body.active === false ? 0 : 1, modulesValue, id)]
      if (String(body.password || '').length) {
        if (String(body.password).length < 8) return json({ error: 'A nova senha deve ter pelo menos 8 caracteres.' }, 400)
        const salt = randomHex(16)
        const hash = await passwordHash(String(body.password), salt)
        statements.push(env.DB.prepare('UPDATE users SET password_hash=?, password_salt=? WHERE id=?').bind(hash, salt, id))
      }
      await env.DB.batch(statements)
      return json({ ok: true })
    }
    return json({ error: 'Método não permitido.' }, 405)
  } catch (error) {
    const duplicate = String(error).toLowerCase().includes('unique')
    return json({ error: duplicate ? 'Esse nome de usuário já está cadastrado.' : 'Não foi possível salvar o perfil.' }, 400)
  }
}
