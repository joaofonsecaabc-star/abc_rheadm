export const onRequestGet = async ({ env }: { env: any }) => {
  const row = await env.DB.prepare('SELECT COUNT(*) AS total FROM users').first()
  return Response.json({ needsSetup: Number(row?.total || 0) === 0 })
}
