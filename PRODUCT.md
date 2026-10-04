# Estoque Franquias — ProFood Embalagens

## O que é
Painel web interno da ProFood Embalagens para controlar, todo dia, o estoque de embalagens que a fábrica produz para as franquias clientes (Sodiê Doces, The Best Açaí, Nanica, Cachorrão do Claudião, TT Burger, Coco Bambu e outras). Ele mostra quantos dias o estoque de cada produto ainda dura e o que precisa de ordem de produção.

## Quem usa
- **Analista de PCP** (dono do painel): libera produção, abre O.S. e controla o estoque de todas as franquias. Trabalha sozinho e sob carga alta, então o painel precisa dizer em segundos o que resolver primeiro.
- **Equipe interna** com três perfis: administrador (edita e gerencia usuários), editor (edita O.S., etapa, prazo e importa a planilha) e leitor (só consulta).
- **Chefia / diretoria:** olha o resumo e os números gerais; o painel também serve para mostrar o trabalho do PCP.
- **Comercial / vendedores:** consultam o estoque de uma franquia para responder ao cliente.
- Uso principal no computador do escritório; consulta rápida no celular.

## Trabalho principal
1. Importar a planilha do dia (exportada do Google Sheets como .xlsx). Cada importação vira uma posição datada e as anteriores formam o histórico.
2. Ver quais produtos estão abaixo do limite e quais já faltam.
3. Registrar para cada produto a O.S., a etapa da produção (última movimentação) e o prazo.
4. Cobrar O.S. vencidas e exportar listas de críticos.

## Regras do negócio
- **Dias de estoque** = disponível ÷ (média mensal ÷ dias por mês).
- **Disponível** = estoque − empenhado. Disponível negativo = já em falta.
- **Situações:** crítico abaixo de 40 dias; atenção de 40 a 54 dias (margem de 15); seguro a partir de 55; sem consumo quando não há média de saída. Os limites são editáveis em Configurações.
- **O.S.:** número, etapa (Impressão, Corte e Vinco, Colagem, Formando etc.) e prazo. Situação do prazo: vencido, vencendo em até 5 dias, no prazo, sem prazo, concluída.
- **Franquia:** identificada pela primeira palavra-chave encontrada na descrição do produto.
- São 26 franquias atendidas e cerca de 330 produtos por posição.

## Vocabulário (usar exatamente estes termos)
Material/produto, franquia, estoque, empenhado, disponível, cobertura (dias de estoque), crítico, atenção, seguro, sem consumo, O.S., etapa, prazo, posição (a planilha de um dia), importar planilha.

## Superfícies
- **Login** — modo Experience: tela de marca. Fundo preto, ondas vermelhas em WebGL na parte de baixo, logo ProFood branca com leve 3D, cartão de vidro sem textos extras.
- **Painel (Visão geral, Estoque, Franquias, O.S., Histórico, Configurações, Guia Power BI)** — modo Operate: leitura rápida, prioridade clara, edição em painel lateral.
- **Parte de produtos** — etiquetas de caixa sobre papelão kraft, com abas de franquia por cima e cobertura em 12 caixinhas de 10 dias.

## Sensação desejada
Ao mesmo tempo: **calma e organizado** (tudo no lugar, fácil de achar), **controle e urgência** (o que está pegando fogo aparece primeiro), **orgulho da marca** (cara de ProFood, bom para mostrar à chefia) e **moderno e tecnológico** (vidro, animações, efeitos). Em conflito, vence a clareza: o efeito nunca pode atrasar a leitura de um número crítico.

## Marca e direção visual
- Cores da ProFood: preto, vermelho `#E3323B` e branco. O resto do painel segue essa paleta; a parte de produtos usa kraft e etiqueta branca por escolha do usuário.
- Vermelho é reservado para alerta e ação; atenção usa âmbar para não se confundir com crítico (daltonismo).
- O usuário gosta de animações: sequência de entrada nos gráficos, números contando, barras crescendo, linha se desenhando, e fundo com logo em vidro se movendo devagar.
- Referências já pedidas pelo usuário: painel escuro com cartões de vidro e barras hachuradas; vidro derretido vermelho embaixo da tela no login.
- Já rejeitado: paleta roxa; layout "bagunçado e difícil de entender" (três caixinhas por linha de produto, rótulos minúsculos em caixa alta).

## Restrições técnicas
- Arquivo único `index.html`, publicado na Vercel a partir do GitHub; dados e login no Supabase (tabelas `profiles`, `snapshots`, `controle`, `config`, com RLS por perfil).
- Também roda como artifact do Claude, que bloqueia vídeo e imagens externas: tudo visual precisa ser CSS, SVG ou WebGL embutido.
- Idioma: português do Brasil. Respeitar `prefers-reduced-motion` e funcionar em 390 px de largura sem rolagem lateral.
