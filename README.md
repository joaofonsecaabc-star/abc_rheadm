# ABC — Pessoas e Benefícios

Aplicação React + TypeScript preparada para Cloudflare Pages, Pages Functions e banco Cloudflare D1.

## Estrutura

- `src/`: interface React.
- `functions/api/`: API executada no Cloudflare.
- `migrations/0001_initial.sql`: esquema completo do D1.
- `wrangler.jsonc`: configuração do Pages e binding `DB`.

## Desenvolvimento da interface

```bash
pnpm install
pnpm dev
```

Em `localhost`, o sistema usa o armazenamento local como contingência. No domínio publicado, usa `/api/state` e o D1.

## Criar e configurar o D1

```bash
npx wrangler login
npx wrangler d1 create abc_db
```

Copie o `database_id` retornado para `wrangler.jsonc` e aplique o SQL:

```bash
npx wrangler d1 migrations apply abc_db --remote
```

## Criar o primeiro administrador

Depois da primeira publicação, execute uma única vez (troque o domínio e a senha):

```bash
curl -X POST https://SEU-DOMINIO/api/auth/setup \
  -H "Content-Type: application/json" \
  -d '{"username":"admin","fullName":"Administrador ABC","password":"TROQUE-POR-UMA-SENHA-FORTE"}'
```

A rota deixa de aceitar novos cadastros assim que o primeiro usuário é criado.

## Publicar no Cloudflare Pages

1. Envie o projeto para um repositório GitHub.
2. No Cloudflare, crie um projeto em Workers & Pages e conecte o repositório.
3. Comando de build: `pnpm build`.
4. Diretório de saída: `dist`.
5. Confirme o binding D1 com nome `DB`.
6. Publique e crie o administrador inicial.

## IA da OpenAI

Os endpoints `/api/warning-reason` e `/api/improve-text` usam a Responses API da OpenAI quando o segredo `OPENAI_API_KEY` está configurado. A chave nunca é enviada ao navegador nem armazenada no GitHub.

Cadastre a chave diretamente como segredo do projeto Cloudflare Pages:

```bash
pnpm dlx wrangler pages secret put OPENAI_API_KEY --project-name abc-rheadm
```

O modelo padrão é `gpt-5.6-luna`. Para escolher outro modelo, cadastre também a variável `OPENAI_MODEL` no painel do Cloudflare. Se a chave da OpenAI não estiver configurada, o backend mantém o binding de IA do Cloudflare como contingência.

## Segurança e dados

- Senhas usam PBKDF2-SHA256 com salt e 150.000 iterações.
- Sessões usam cookie `HttpOnly`, `Secure` e `SameSite=Strict`.
- A API centraliza funcionários, recargas, ocorrências, lojas, cargos, configurações e históricos no D1.
- O `localStorage` permanece apenas como cache/contingência no ambiente local.
