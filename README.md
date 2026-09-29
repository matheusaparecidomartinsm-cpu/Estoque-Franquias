# Estoque Franquias — ProFood Embalagens

Painel de controle de estoque das franquias: dias de estoque (disponível ÷ consumo diário), alertas abaixo de 40 dias, acompanhamento de O.S. (número, última movimentação e prazo), histórico de posições e guia de montagem no Power BI.

## Arquivos
- `index.html` — o painel completo, com tela de login, em um único arquivo.
- `powerbi/01_PowerQuery_M.pq` — consultas Power Query (M).
- `powerbi/02_DAX.dax` — medidas DAX.
- `powerbi/03_Tema_ControleEstoque.json` — tema do relatório.

## Banco de dados
A versão web (`index.html`, publicada na Vercel) guarda tudo no Supabase:
- **Login:** Supabase Auth. O usuário digitado vira `usuario@profood.local` por baixo dos panos.
- **Tabelas:** `profiles` (usuário, nome e perfil), `snapshots` (posições diárias), `controle` (O.S., última movimentação e prazo por produto) e `config` (parâmetros).
- **Permissões (RLS):** só usuários cadastrados em `profiles` leem os dados; `editor` e `admin` alteram; `leitor` só consulta.
- **Gestão de usuários:** Edge Function `usuarios`, que só o `admin` pode chamar.

A chave que aparece no `index.html` é a chave pública (publishable) do Supabase. Ela sozinha não dá acesso aos dados: sem login de um usuário cadastrado, as regras do banco bloqueiam tudo.
