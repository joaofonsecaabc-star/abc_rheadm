import { currentUser } from './auth/_utils'
import { generateWithOpenAI } from './_openai'

const json = (data: unknown, status = 200) => Response.json(data, { status })

export const onRequestPost = async ({ request, env }: { request: Request; env: any }) => {
  const user = await currentUser(request, env.DB)
  if (!user) return json({ error: 'Sessão expirada.' }, 401)
  const body: any = await request.json().catch(() => ({}))
  const description = String(body.description || '').trim()
  if (description.length < 5) return json({ error: 'Descreva resumidamente o que aconteceu.' }, 400)
  if (description.length > 1000) return json({ error: 'A descrição deve ter no máximo 1.000 caracteres.' }, 400)
  if (!env.OPENAI_API_KEY && !env.AI) return json({ error: 'O recurso de IA ainda não está configurado.' }, 503)

  const prompt = `Você auxilia um setor de Recursos Humanos brasileiro a redigir somente o motivo factual de uma advertência disciplinar.
Reescreva a descrição abaixo em português do Brasil, com linguagem profissional, objetiva, neutra e respeitosa.
Não invente fatos, datas, leis, nomes, punições ou circunstâncias. Não faça acusações além do que foi informado.
Produza apenas uma oração curta que possa completar naturalmente a frase "foi apurado que, ...".
Comece diretamente pelo fato, preferencialmente com verbo em letra minúscula. Nunca inicie com "foi constatado", "foi apurado", "foi verificado", "ocorreu" ou expressões equivalentes.
Ao se referir à pessoa advertida, use sempre "o funcionário mencionado", nunca "um funcionário", "o funcionário" ou nomes próprios.
Não escreva "na data informada" e não inclua outra data, título, saudação, aspas, punição ou observações adicionais.

Descrição informada: ${description}`

  if (env.AI) {
    try {
      const result: any = await env.AI.run('@cf/meta/llama-3.1-8b-instruct-fp8-fast', {
        messages: [
          { role: 'system', content: 'Siga rigorosamente as instruções e devolva somente o texto solicitado.' },
          { role: 'user', content: prompt },
        ],
        max_tokens: 220,
        temperature: 0.2,
      })
      const reason = String(result?.response || '').trim().replace(/^["']|["']$/g, '').replace(/^(?:foi\s+(?:constatado|apurado|verificado)\s+que|constatou-se\s+que|ocorreu(?:\s+que)?)[,:\s]*/i, '').replace(/\bna data informada[,]?\s*/gi, '').replace(/\b(?:um|o) funcionário(?! mencionado)\b/gi, 'o funcionário mencionado')
      if (reason) return json({ reason, provider: 'cloudflare' })
      console.error('Cloudflare Workers AI returned an empty warning reason')
    } catch (error) {
      console.error('Cloudflare Workers AI warning reason error; trying OpenAI fallback', error)
    }
  }

  try {
    const openAIText = await generateWithOpenAI(env, {
      instructions: 'Você auxilia um setor de Recursos Humanos brasileiro. Siga rigorosamente as instruções, preserve os fatos e devolva somente o texto final solicitado.',
      input: prompt,
      maxOutputTokens: 220,
    })
    const reason = String(openAIText || '').trim().replace(/^["']|["']$/g, '').replace(/^(?:foi\s+(?:constatado|apurado|verificado)\s+que|constatou-se\s+que|ocorreu(?:\s+que)?)[,:\s]*/i, '').replace(/\bna data informada[,]?\s*/gi, '').replace(/\b(?:um|o) funcionário(?! mencionado)\b/gi, 'o funcionário mencionado')
    if (reason) return json({ reason, provider: 'openai' })
  } catch (error) {
    console.error('OpenAI warning reason fallback error', error)
  }
  return json({ error: 'Não foi possível gerar o texto com IA agora.' }, 502)
}
