import { currentUser } from './auth/_utils'
import { generateWithOpenAI } from './_openai'

const json = (data: unknown, status = 200) => Response.json(data, { status })

const normalizeWarningText = (value: unknown) => {
  const fact = String(value || '').trim()
    .replace(/^(["'])|(["'])$/g, '')
    .replace(/[.]+$/, '')
    .replace(/^(?:foi\s+(?:constatado|apurado|verificado)\s+que|constatou-se\s+que|ocorreu(?:\s+que)?)[,:\s]*/i, '')
    .replace(/\bna data informada[,]?\s*/gi, '')
    .replace(/^(?:o\s+)?funcionário mencionado[,:]?\s*/i, '')
    .replace(/\b(?:um|o) funcionário(?! mencionado)\b/gi, 'o funcionário mencionado')
    .replace(/\s+/g, ' ')
    .trim()
  return fact ? `o funcionário mencionado ${fact.charAt(0).toLowerCase()}${fact.slice(1)}` : ''
}

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
    : `Reescreva em UMA frase factual para uma advertência de RH. Comece exatamente com "o funcionário mencionado". Preserve somente os fatos informados e não invente intenção, consequência, reincidência, testemunha, regra, punição ou lei. Use português brasileiro profissional, direto e natural, com no máximo 45 palavras. Não inclua data, título, saudação, aspas ou ponto final. Não use "foi constatado", "foi apurado", "na data informada", "um funcionário" ou o nome da pessoa. Exemplo: "faltou e não justificou" deve resultar em "o funcionário mencionado ausentou-se do trabalho sem apresentar justificativa".`
  const prompt = `${instruction}\n\nTexto informado: ${text}`

  if (env.AI) {
    try {
      const result: any = await env.AI.run('@cf/meta/llama-3.1-8b-instruct-fp8-fast', {
        messages: [
          { role: 'system', content: 'Você revisa textos administrativos brasileiros. Responda somente com o texto final solicitado.' },
          { role: 'user', content: prompt },
        ],
        max_tokens: context === 'receipt_reference' ? 60 : 90,
        temperature: context === 'receipt_reference' ? 0.15 : 0.05,
      })
      const improved = context === 'receipt_reference'
        ? String(result?.response || '').trim().replace(/^["']|["']$/g, '').replace(/\.$/, '')
        : normalizeWarningText(result?.response)
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
      maxOutputTokens: context === 'receipt_reference' ? 60 : 90,
    })
    const improved = context === 'receipt_reference'
      ? String(openAIText || '').trim().replace(/^["']|["']$/g, '').replace(/\.$/, '')
      : normalizeWarningText(openAIText)
    if (improved) return json({ text: improved, provider: 'openai' })
  } catch (error) {
    console.error('OpenAI improve text fallback error', error)
  }
  return json({ error: 'Não foi possível melhorar o texto com IA agora.' }, 502)
}
