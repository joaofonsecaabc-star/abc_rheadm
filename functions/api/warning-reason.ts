import { currentUser } from './auth/_utils'

const json = (data: unknown, status = 200) => Response.json(data, { status })

export const onRequestPost = async ({ request, env }: { request: Request; env: any }) => {
  const user = await currentUser(request, env.DB)
  if (!user) return json({ error: 'Sessão expirada.' }, 401)

  const body: any = await request.json().catch(() => ({}))
  const description = String(body.description || '').trim()
  if (description.length < 5) return json({ error: 'Descreva resumidamente o que aconteceu.' }, 400)
  if (description.length > 1000) return json({ error: 'A descrição deve ter no máximo 1.000 caracteres.' }, 400)
  if (!env.AI) return json({ error: 'O recurso de IA ainda não está configurado.' }, 503)

  const prompt = `Você auxilia um setor de Recursos Humanos brasileiro a redigir somente o motivo factual de uma advertência disciplinar.
Reescreva a descrição abaixo em português do Brasil, com linguagem profissional, objetiva, neutra e respeitosa.
Não invente fatos, datas, leis, nomes, punições ou circunstâncias. Não faça acusações além do que foi informado.
Produza apenas um parágrafo curto, sem título, saudação, aspas ou observações adicionais.

Descrição informada: ${description}`

  try {
    const result: any = await env.AI.run('@cf/meta/llama-3.1-8b-instruct', {
      messages: [
        { role: 'system', content: 'Siga rigorosamente as instruções e devolva somente o texto solicitado.' },
        { role: 'user', content: prompt },
      ],
      max_tokens: 220,
      temperature: 0.2,
    })
    const reason = String(result?.response || '').trim().replace(/^['"]|['"]$/g, '')
    if (!reason) return json({ error: 'A IA não conseguiu criar o texto. Tente novamente.' }, 502)
    return json({ reason })
  } catch (error) {
    console.error('warning reason AI error', error)
    return json({ error: 'Não foi possível gerar o texto com IA agora.' }, 502)
  }
}
