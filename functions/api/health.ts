interface Env { DB: any }
export const onRequestGet = async ({ env }: { env: Env }) => {
  try {
    const result = await env.DB.prepare('SELECT 1 AS ok').first()
    return Response.json({ ok: result?.ok === 1, database: 'D1' })
  } catch (error) {
    return Response.json({ ok: false, error: String(error) }, { status: 500 })
  }
}
