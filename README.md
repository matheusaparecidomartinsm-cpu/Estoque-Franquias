# Estoque Franquias — ProFood Embalagens

Painel de controle de estoque das franquias: dias de estoque (disponível ÷ consumo diário), alertas abaixo de 40 dias, acompanhamento de O.S. (número, última movimentação e prazo), histórico de posições e guia de montagem no Power BI.

## Arquivos
- `index.html` — o painel completo, com tela de login, em um único arquivo.
- `powerbi/01_PowerQuery_M.pq` — consultas Power Query (M).
- `powerbi/02_DAX.dax` — medidas DAX.
- `powerbi/03_Tema_ControleEstoque.json` — tema do relatório.

## Observação
Os dados (posições de estoque, O.S. e usuários) ficam no banco compartilhado do artifact publicado no Claude. Aberto fora dele (por exemplo, no GitHub Pages), o painel mostra a tela de login, mas não consegue conectar aos dados.
