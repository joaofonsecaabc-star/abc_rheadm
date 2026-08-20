type GenerateOptions = {
  instructions: string
  input: string
  maxOutputTokens?: number
}

const extractOutputText = (payload: any) => {
  if (typeof payload?.output_text === 'string') return payload.output_text.trim()
  const parts: string[] = []
  for (const item of Array.isArray(payload?.output) ? payload.output : []) {
    for (const content of Array.isArray(item?.content) ? item.content : []) {
      if (content?.type === 'output_text' && typeof content?.text === 'string') parts.push(content.text)
    }
  }
  return parts.join('\n').trim()
}

export const generateWithOpenAI = async (env: any, options: GenerateOptions) => {
  const apiKey = String(env.OPENAI_API_KEY || '').trim()
  if (!apiKey) return null

  const response = await fetch('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${apiKey}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      model: String(env.OPENAI_MODEL || 'gpt-5.6-luna'),
      instructions: options.instructions,
      input: options.input,
      max_output_tokens: options.maxOutputTokens || 260,
      reasoning: { effort: 'none' },
      store: false,
    }),
  })

  const payload: any = await response.json().catch(() => ({}))
  if (!response.ok) {
    const message = String(payload?.error?.message || `OpenAI respondeu com o status ${response.status}.`)
    throw new Error(message)
  }

  const text = extractOutputText(payload)
  if (!text) throw new Error('A OpenAI não retornou um texto.')
  return text
}
