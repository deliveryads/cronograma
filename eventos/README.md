# Eventos de curta duração (DAI Eventual)

Página estática (`eventos/index.html`) que lê os eventos **direto do Google Sheets**.
Qualquer alteração na planilha aparece na página automaticamente (a cada 60 s,
ao voltar para a aba ou pelo botão **Atualizar**). Não há dados embutidos.

- Planilha: https://docs.google.com/spreadsheets/d/1GJ_ShKOHqaLLfpgWJHIrXu44fPMYBaSSSrRiVEbFkYU/edit
- Página publicada: https://deliveryads.github.io/cronograma/eventos/

## Pré-requisito obrigatório

A planilha precisa estar em **Compartilhar → Acesso geral → Qualquer pessoa com o link → Leitor**.
Sem isso o navegador não consegue ler os dados. Atenção: com isso, quem tiver o link
da planilha consegue visualizá-la.

## Visualizações

Menu à esquerda (no celular, botão ☰) com Visualização, Filtros (Mês, Busca, Evento,
Status, Transmissão LF), Legenda e Dados (Atualizar, Imprimir / PDF).

- **Linha do tempo** (padrão): de hoje até 1, 2 ou 3 meses à frente, espaçamento
  proporcional às datas, nome e data de cada evento, marcos de mês e "Hoje".

- **Calendário**: mês a mês, cores da legenda, detalhes ao clicar.
- **Cards**: um card por evento, por mês ou todos.
- **Lista**: tabela com todas as colunas, agrupada por mês.

**Próximo evento** (topo de todas as visões, ignora os filtros): cards compactos com todos os
eventos em andamento hoje e todos os que começam na próxima data com evento. Cancelados não entram.

O filtro de Mês restringe todas as visões ao mês escolhido. Eventos passados aparecem esmaecidos.

Legenda (coluna Status): Confirmado = verde, A confirmar = amarelo, Cancelado = vermelho;
qualquer outro valor aparece em cinza como "Outros", com aviso.

## Estrutura da planilha

Cabeçalho (Mês, Data início, ...) em qualquer uma das 10 primeiras linhas — a página
localiza sozinha. Linhas acima dele (título, linha em branco) são ignoradas.
Dados logo abaixo do cabeçalho (ver `modelo-planilha.csv`).

| Coluna | Uso na página |
|---|---|
| Evento | Obrigatória. Título do card/barra e filtro "Evento" |
| Data início | Obrigatória. Posição no calendário |
| Data fim | Opcional (vazia = evento de 1 dia). Não pode ser anterior à Data início |
| Status | Cor do evento, legenda e filtro |
| Transmissão LF | Exibida no card e filtro |
| Observação | Exibida no card |
| Mês | Não exibida: o mês é calculado pelas datas |

Qualquer coluna nova adicionada à planilha aparece automaticamente nos cards.
Linhas sem Evento/Data início ou com datas inválidas não são exibidas e são listadas
num aviso acima dos cards, com o número da linha.

## Configuração

No topo do `<script>` em `index.html`, objeto `CONFIG`:
`SHEET_ID`, `SHEET_NAME` (vazio = primeira aba), `HEADER_ROW` (linha esperada do cabeçalho),
`REFRESH_SECONDS`. Para testar outra planilha sem editar: `?planilha=ID&aba=Nome`.

## Prévia sem publicar

Definir `window.__SNAPSHOT__ = { takenAt, headerRow, table }` antes do script principal faz a
página usar uma cópia dos dados (formato da resposta do Google Sheets) em vez de ler a planilha.
