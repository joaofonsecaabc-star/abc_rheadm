import { currentUser } from './auth/_utils'
import { generateWithOpenAI } from './_openai'

const json = (data: unknown, status = 200) => Response.json(data, { status })

export const onRequestPost = async ({ request, env }: { request: Request; env: any }) => {
  const user = await currentUser(request, env.DB)
  if (!user) return json({ error: 'Sessão expirada.' }, 401)
  const body: any = await request.json().catch(() => ({}))
  const text = String(body.text || '').trim()
  const context = String(body.context || 'reason')
  if (text.length < 2) return json({ error: 'Digite um texto para a IA melhorar.' }, 400)
  if (text.length > 1500) return json({ error: 'O texto deve ter no máximo 1.500 caracteres.' }, 400)
  if (!env.OPENAI_API_KEY && !env.AI) return json({ error: 'O recurso de IA ainda não está configurado.' }, 503)

  const instruction = context === 'receipt_reference'
    ? `Reescreva o texto como uma expressão curta, natural e profissional que complete a frase "referente a ..." em um recibo brasileiro. Evite dois-pontos, título, saudação, aspas e ponto final. Não invente informações. Exemplo: "Vale" pode se tornar "pagamento de vale".`
    : `Reescreva o motivo em português do Brasil, com linguagem profissional, objetiva, neutra e respeitosa. Preserve rigorosamente os fatos informados. Não invente nomes, datas, leis, punições ou circunstâncias. Produza apenas uma oração curta que complete naturalmente "foi apurado que, na data informada, ...". Comece diretamente pelo fato e nunca use como introdução "foi constatado", "foi apurado", "foi verificado" ou "ocorreu". Não inclua título, saudação, aspas ou comentários.`
  const prompt = `${instruction}\n\nTexto informado: ${text}`

  if (env.AI) {
    try {
      const result: any = await env.AI.run('@cf/meta/llama-3.1-8b-instruct-fp8-fast', {
        messages: [
          { role: 'system', content: 'Você revisa textos administrativos brasileiros. Responda somente com o texto final solicitado.' },
          { role: 'user', content: prompt },
        ],
        max_tokens: context === 'receipt_reference' ? 100 : 260,
        temperature: 0.25,
      })
      const improved = String(result?.response || '').trim().replace(/^["']|["']$/g, '').replace(/\.$/, '').replace(/^(?:foi\s+(?:constatado|apurado|verificado)\s+que|constatou-se\s+que|ocorreu(?:\s+que)?)[,:\s]*/i, '')
      if (improved) return json({ text: improved, provider: 'cloudflare' })
      console.error('Cloudflare Workers AI returned an empty improved text')
    } catch (error) {
      console.error('Cloudflare Workers AI improve text error; trying OpenAI fallback', error)
    }
  }

  try {
    const openAIText = await generateWithOpenAI(env, {
      instructions: 'Você revisa textos administrativos brasileiros. Preserve integralmente o sentido e os fatos. Responda somente com o texto final solicitado.',
      input: prompt,
      maxOutputTokens: context === 'receipt_reference' ? 100 : 260,
    })
    const improved = String(openAIText || '').trim().replace(/^["']|["']$/g, '').replace(/\.$/, '').replace(/^(?:foi\s+(?:constatado|apurado|verificado)\s+que|constatou-se\s+que|ocorreu(?:\s+que)?)[,:\s]*/i, '')
    if (improved) return json({ text: improved, provider: 'openai' })
  } catch (error) {
    console.error('OpenAI improve text fallback error', error)
  }
  return json({ error: 'Não foi possível melhorar o texto com IA agora.' }, 502)
}
