import { currentUser } from './auth/_utils'
import { generateWithOpenAI } from './_openai'

const json = (data: unknown, status = 200) => Response.json(data, { status })

const parseResults = (value: unknown, allowedIds: Set<number>) => {
  const raw = String(value || '').trim()
  const start = raw.indexOf('['), end = raw.lastIndexOf(']')
  if (start < 0 || end < start) return []
  try {
    const parsed = JSON.parse(raw.slice(start, end + 1))
    if (!Array.isArray(parsed)) return []
    return parsed
      .map((item: any) => ({ id: Number(item.id), gender: String(item.gender || '') }))
      .filter((item) => allowedIds.has(item.id) && (item.gender === 'Masculino' || item.gender === 'Feminino'))
  } catch {
    return []
  }
}

export const onRequestPost = async ({ request, env }: { request: Request; env: any }) => {
  const user = await currentUser(request, env.DB)
  if (!user) return json({ error: 'Sessão expirada.' }, 401)
  if (user.role !== 'admin') return json({ error: 'Apenas administradores podem completar este cadastro.' }, 403)

  const body: any = await request.json().catch(() => ({}))
  const people = Array.isArray(body.people)
    ? body.people
      .slice(0, 100)
      .map((item: any) => ({ id: Number(item.id), name: String(item.name || '').trim(), role: String(item.role || '').trim() }))
      .filter((item: any) => item.id && item.name)
    : []
  if (!people.length) return json({ results: [] })
  if (!env.OPENAI_API_KEY && !env.AI) return json({ error: 'O recurso de IA ainda não está configurado.' }, 503)

  const allowedIds = new Set<number>(people.map((item: any) => item.id))
  const prompt = `Classifique o sexo cadastral mais provável a partir dos nomes brasileiros e das funções abaixo.
Devolva SOMENTE um array JSON válido no formato [{"id":1,"gender":"Feminino"}].
Use exclusivamente "Masculino" ou "Feminino". Preserve cada id e devolva todos os registros, sem explicações.
Nesta empresa, as funções de Auxiliar de Serviços Gerais, Frente de Caixa, Caixa/Operadora de Caixa, Suporte e Atendente são ocupadas por mulheres e devem ser classificadas como Feminino.
As funções de Repositor, Repositor de Mercadorias, Subgerente e Sub-Gerente são ocupadas por homens e devem ser classificadas como Masculino.

PESSOAS:
${JSON.stringify(people)}`

  if (env.AI) {
    try {
      const result: any = await env.AI.run('@cf/meta/llama-3.1-8b-instruct-fp8-fast', {
        messages: [
          { role: 'system', content: 'Responda somente com JSON válido e siga exatamente o formato solicitado.' },
          { role: 'user', content: prompt },
        ],
        max_tokens: Math.min(1800, people.length * 24 + 80),
        temperature: 0,
      })
      const results = parseResults(result?.response, allowedIds)
      if (results.length) return json({ results, provider: 'cloudflare' })
    } catch (error) {
      console.error('Cloudflare Workers AI gender inference error; trying OpenAI fallback', error)
    }
  }

  try {
    const text = await generateWithOpenAI(env, {
      instructions: 'Responda somente com JSON válido e siga exatamente o formato solicitado.',
      input: prompt,
      maxOutputTokens: Math.min(1800, people.length * 24 + 80),
    })
    const results = parseResults(text, allowedIds)
    if (results.length) return json({ results, provider: 'openai' })
  } catch (error) {
    console.error('OpenAI gender inference fallback error', error)
  }
  return json({ error: 'Não foi possível classificar os cadastros agora.' }, 502)
}
