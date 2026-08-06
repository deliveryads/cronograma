# Contrato de dados — Cronograma de Responsabilidades

Há dois jeitos de alimentar a página com dados:

1. **Estático (`data.json`)** — descrito neste arquivo. A página lê um único
   `data.json` (mesma pasta do `index.html`). Sem gravação: os botões Criar/Editar/Excluir
   funcionam, mas as alterações só existem enquanto a aba do navegador estiver aberta.
   Se `data.json` não for encontrado, a página usa dados de exemplo embutidos (demonstração).
2. **Ao vivo (Google Sheets)** — ver `APPS_SCRIPT.md`. A página lê e grava direto numa
   planilha do Google Sheets através de um Google Apps Script Web App (`Code.gs`).
   Criar/Editar/Excluir passam a persistir de verdade, compartilhado entre todos que
   acessarem a página.

O mapeamento abaixo (abas → campos) vale para os dois casos — a única diferença é
se os dados chegam via um arquivo `data.json` gerado a partir da planilha, ou se a
própria planilha é a fonte ao vivo.

## Mapeamento aba → `data.json`

### Aba "Analistas" → `analistas`

Cadastro de todos os analistas.

| Campo da aba | Campo no JSON | Tipo | Observação |
|---|---|---|---|
| ID / matrícula | `id` | string | Único, ex.: `"an-01"` |
| Nome | `nome` | string | |
| Squad | `squadId` | string | Deve bater com um `id` da lista `squads` |
| Função | `funcao` | string | Ex.: "Analista de Roteiros Pleno", "Coordenador de Squad" |

```json
{ "id": "an-01", "nome": "Camila Duarte", "squadId": "sq-alpha", "funcao": "Analista de Roteiros Pleno" }
```

### Aba "Squad" → `squads`

Divisão de squads, com os analistas (já cobertos por `analistas[].squadId`) e os roteiros/programas
pelos quais a squad é responsável.

| Campo da aba | Campo no JSON | Tipo | Observação |
|---|---|---|---|
| ID | `id` | string | Único, ex.: `"sq-alpha"` |
| Nome da squad | `nome` | string | |
| Roteiros responsáveis | `roteiros` | array de strings | Títulos dos roteiros/programas atendidos |

```json
{ "id": "sq-alpha", "nome": "Squad Alpha", "roteiros": ["Coração em Chamas", "Reencontro"] }
```

### Aba "Férias", "Cursos" e "Extraordinários" → `eventos`

As três abas viram itens da mesma lista `eventos`, diferenciados pelo campo `tipo`. Cada linha da
planilha vira um objeto:

| Campo da aba | Campo no JSON | Tipo | Observação |
|---|---|---|---|
| — | `id` | string | Único, ex.: `"ev-001"` |
| — | `tipo` | string | Ver tabela de tipos abaixo |
| Título / assunto | `titulo` | string | |
| Analista | `analistaId` | string (opcional) | `id` de `analistas` |
| Squad | `squadId` | string (opcional) | `id` de `squads`. Se ausente e houver `analistaId`, a squad é herdada do analista |
| Data início | `dataInicio` | string `AAAA-MM-DD` | |
| Data fim | `dataFim` | string `AAAA-MM-DD` | Igual à data início quando for evento de um dia só |
| Descrição / tema | `descricao` | string (opcional) | |

Tabela de `tipo` por aba de origem:

| Aba da planilha | valor de `tipo` |
|---|---|
| Férias | `ferias` |
| Cursos | `curso` |
| Extraordinários | `reuniao` |

Dois tipos adicionais já suportados pela página, sem aba própria ainda definida — usar o mesmo
formato de `eventos` quando a fonte de dados for definida:

| Categoria | valor de `tipo` |
|---|---|
| Data Importante | `data_importante` |
| Responsabilidade de Ponto | `ponto` |

```json
{ "id": "ev-001", "tipo": "ferias", "titulo": "Férias", "analistaId": "an-01", "dataInicio": "2026-01-12", "dataFim": "2026-01-26", "descricao": "Férias regulamentares" }
{ "id": "ev-009", "tipo": "reuniao", "titulo": "Revisão de Escopo", "squadId": "sq-alpha", "analistaId": "an-03", "dataInicio": "2026-02-11", "dataFim": "2026-02-11", "descricao": "Convocada pela coordenação" }
```

## Estrutura completa esperada

```json
{
  "squads": [ { "id", "nome", "roteiros": [] } ],
  "analistas": [ { "id", "nome", "squadId", "funcao" } ],
  "eventos": [ { "id", "tipo", "titulo", "analistaId", "squadId", "dataInicio", "dataFim", "descricao" } ]
}
```

`data.json` neste repositório já segue este formato com dados de exemplo — pode ser usado como
referência ao gerar o arquivo final a partir da planilha.
