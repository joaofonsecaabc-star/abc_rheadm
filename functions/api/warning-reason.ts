import { currentUser } from './auth/_utils'
import { generateWithOpenAI } from './_openai'

const json = (data: unknown, status = 200) => Response.json(data, { status })

const normalizeReason = (value: unknown, subject: string) => {
  const fact = String(value || '').trim()
    .replace(/^(["'])|(["'])$/g, '')
    .replace(/[.]+$/, '')
    .replace(/^(?:foi\s+(?:constatado|apurado|verificado)\s+que|constatou-se\s+que|ocorreu(?:\s+que)?)[,:\s]*/i, '')
    .replace(/\bna data informada[,]?\s*/gi, '')
    .replace(/^(?:(?:o|a)\s+)?(?:funcionári[oa]|colaborador[oa]) mencionad[oa][,:]?\s*/i, '')
    .replace(/\b(?:um|uma|o|a)\s+(?:funcionári[oa]|colaborador[oa])(?:\s+mencionad[oa])?\b/gi, subject)
    .replace(/\s+/g, ' ')
    .trim()
  return fact ? `${subject} ${fact.charAt(0).toLowerCase()}${fact.slice(1)}` : ''
}

export const onRequestPost = async ({ request, env }: { request: Request; env: any }) => {
  const user = await currentUser(request, env.DB)
  if (!user) return json({ error: 'Sessão expirada.' }, 401)
  const body: any = await request.json().catch(() => ({}))
  const description = String(body.description || '').trim()
  const subject = body.gender === 'Feminino'
    ? 'a colaboradora mencionada'
    : body.gender === 'Masculino'
      ? 'o colaborador mencionado'
      : 'a pessoa mencionada'
  if (description.length < 5) return json({ error: 'Descreva resumidamente o que aconteceu.' }, 400)
  if (description.length > 1000) return json({ error: 'A descrição deve ter no máximo 1.000 caracteres.' }, 400)
  if (!env.OPENAI_API_KEY && !env.AI) return json({ error: 'O recurso de IA ainda não está configurado.' }, 503)

  const prompt = `Transforme o relato em UMA frase factual para uma advertência de RH.
REGRAS OBRIGATÓRIAS:
- Comece exatamente com "${subject}".
- Preserve somente os fatos do relato; não invente intenção, consequência, reincidência, testemunha, regra, punição ou lei.
- Use português brasileiro profissional, direto e natural.
- Máximo de 45 palavras, sem data, título, saudação, aspas ou ponto final.
- Não use "foi constatado", "foi apurado", "na data informada", "um funcionário" ou o nome da pessoa.

EXEMPLOS:
Relato: faltou e não apresentou justificativa
Resposta: ${subject} ausentou-se do trabalho sem apresentar justificativa
Relato: abriu a gaveta pelo botão em vez de chamar a frente de caixa
Resposta: ${subject} utilizou o botão de abertura da gaveta do caixa, em vez de seguir o procedimento orientado e solicitar o apoio da frente de caixa
Relato: discutiu com uma colega no setor
Resposta: ${subject} envolveu-se em uma discussão com uma colega de trabalho no setor informado

RELATO: ${description}
RESPOSTA:`

  if (env.AI) {
    try {
      const result: any = await env.AI.run('@cf/meta/llama-3.1-8b-instruct-fp8-fast', {
        messages: [
          { role: 'system', content: 'Siga rigorosamente as instruções e devolva somente o texto solicitado.' },
          { role: 'user', content: prompt },
        ],
        max_tokens: 90,
        temperature: 0.05,
      })
      const reason = normalizeReason(result?.response, subject)
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
      maxOutputTokens: 90,
    })
    const reason = normalizeReason(openAIText, subject)
    if (reason) return json({ reason, provider: 'openai' })
  } catch (error) {
    console.error('OpenAI warning reason fallback error', error)
  }
  return json({ error: 'Não foi possível gerar o texto com IA agora.' }, 502)
}
