# Eventos de curta duração — Power BI (versão simples)

Duas páginas, só com visuais nativos do Power BI e **nenhuma fórmula DAX**:

| Página | Visual |
|---|---|
| **Cards** | Cartão de várias linhas: um card por evento |
| **Lista** | Tabela, ordenada pela data de início |

Filtros à esquerda (sincronizados entre as páginas): **Mês**, **Situação** (Próximo,
Em andamento, Realizado), **Status**, **Transmissão LF** e **Evento**.

Todas as colunas de exibição são calculadas no Power Query:

| Coluna | Conteúdo |
|---|---|
| Período | `30/10/2026` ou `30/10 a 08/11/2026` |
| Dia da semana | `Sexta-feira` ou `Sexta-feira a domingo` |
| Status (legenda) | 🟢 Confirmado · 🟡 A confirmar · 🔴 Cancelado · ⚪ outros |
| Situação | Próximo / Em andamento / Realizado (pela data da última atualização) |
| Mês | Mês de início, ex.: `Outubro de 2026` |

## Como usar

1. Coloque o Excel no **OneDrive** ou **SharePoint** (pode ser `Eventos.xlsx` deste pacote ou a
   planilha original *DAI Eventual*; o cabeçalho é encontrado sozinho).
2. No Excel do computador: **Arquivo → Informações → Copiar caminho**.
3. Descompacte este pacote numa pasta nova e abra `projeto/Eventos.pbip` no Power BI Desktop.
4. **Página inicial → Transformar dados → Editar parâmetros** → cole o caminho em
   `CaminhoExcel` → **OK** → **Aplicar alterações**.
5. Credencial: **Conta da Microsoft / organizacional** (e-mail da empresa).
6. Para guardar como arquivo único: **Arquivo → Salvar como → Arquivo do Power BI (.pbix)**.
7. Publicar: **Página inicial → Publicar**; no serviço, configure as credenciais da fonte
   (OAuth) e a **atualização agendada diária**.

Se o Desktop não abrir o `.pbip`: **Arquivo → Opções → Recursos de visualização** → ative
*Opção de salvar Projeto do Power BI (.pbip)* e *formato PBIR*; reinicie.

## Colunas esperadas no Excel

Data início, Data fim, Evento, Status, Transmissão LF, Observação (a coluna Mês do Excel é
ignorada; o mês vem da Data início). Data fim vazia = evento de um dia. Linhas sem Evento ou
Data início, ou com Data fim anterior à Data início, não entram.

## Regerar

`python3 powerbi-simples/src/build.py` (edite `src/`, não `projeto/`).
