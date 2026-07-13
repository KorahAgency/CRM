# Publicar o Korah CRM

## Vercel

1. Entre em https://vercel.com.
2. Clique em `Add New Project`.
3. Envie este projeto ou conecte o repositório.
4. Use:
   - Framework: `Vite`
   - Build command: `npm run build`
   - Output directory: `dist`
5. Cadastre as variáveis de ambiente:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
6. Publique.

## Netlify

1. Entre em https://app.netlify.com/drop.
2. Cadastre as variáveis de ambiente se for usar deploy por projeto.
3. Rode `npm run build`.
4. Arraste a pasta `dist` para publicar.

## Salvar dados no Supabase

1. Crie um projeto em https://supabase.com.
2. Abra `SQL Editor`.
3. Cole e rode o conteúdo do arquivo `supabase-schema.sql`.
4. Vá em `Project Settings` > `API`.
5. Copie:
   - Project URL
   - anon public key
6. Crie um arquivo `.env` na raiz do projeto usando o modelo `.env.example`.

```env
VITE_SUPABASE_URL=https://SEU-PROJETO.supabase.co
VITE_SUPABASE_ANON_KEY=SUA_CHAVE_ANON_PUBLIC
```

7. Rode:

```bash
npm run build
```

Quando o Supabase estiver respondendo corretamente, o cabeçalho do CRM mostra o selo `SUPABASE`.
Se as variáveis não existirem ou a tabela ainda não tiver sido criada, ele mostra `LOCAL`.

## Segurança

O SQL atual deixa o CRM público por link e permite leitura/escrita usando a chave anon.
Isso é prático para começar, mas qualquer pessoa com o link pode alterar os dados.
Para uma operação mais segura, o próximo passo é adicionar login e regras por usuário/equipe.
