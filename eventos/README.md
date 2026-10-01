# Calendário de Eventos (DAI Eventual)

Página estática (`eventos/index.html`) que lê os eventos **direto do Google Sheets**.
Qualquer alteração na planilha aparece na página automaticamente (a cada 60 s,
ao voltar para a aba ou pelo botão **Atualizar**). Não há dados embutidos.

- Planilha: https://docs.google.com/spreadsheets/d/1GJ_ShKOHqaLLfpgWJHIrXu44fPMYBaSSSrRiVEbFkYU/edit
- Página publicada (após merge na `main`): `https://<usuario>.github.io/cronograma/eventos/`

## Pré-requisito obrigatório

A planilha precisa estar em **Compartilhar → Acesso geral → Qualquer pessoa com o link → Leitor**.
Sem isso o navegador não consegue ler os dados. Atenção: com isso, quem tiver o link
da planilha consegue visualizá-la.

## Estrutura da planilha

Linha 1: título livre (ex.: "DAI Eventual"). Linha 2: cabeçalho. Dados a partir da linha 3
(ver `modelo-planilha.csv`).

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
`SHEET_ID`, `SHEET_NAME` (vazio = primeira aba), `HEADER_ROW` (linha do cabeçalho),
`REFRESH_SECONDS`. Para testar outra planilha sem editar: `?planilha=ID&aba=Nome`.
