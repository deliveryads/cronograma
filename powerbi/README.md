# Eventos de curta duração — Power BI

Projeto Power BI (formato **PBIP**) com as mesmas visões da página web, no tema escuro
"azulterra": **Linha do tempo**, **Calendário**, **Cards** e **Lista**, com os cards de
próximo evento no topo de todas as páginas, filtros (Mês, Evento, Status, Transmissão LF,
Período) e legenda.

| Pasta/arquivo | O que é |
|---|---|
| `Eventos.xlsx` | Base de dados (tabela Excel chamada **Eventos**) |
| `projeto/` | Projeto Power BI pronto: abrir `projeto/Eventos.pbip` |
| `deneb/` | Gráficos da linha do tempo e do calendário (Vega-Lite, visual Deneb) |
| `src/` | Gerador do projeto (`build_pbip.py`), tema e estilo dos cards |

## Visuais gratuitos usados (AppSource)

- **Deneb** — linha do tempo e calendário.
- **HTML Content** — cards de próximo evento, cards, lista, legenda e títulos.

O Power BI Desktop baixa os dois automaticamente ao abrir o projeto. Se algum aparecer com
erro "visual não encontrado", em **Visualizações → … → Obter mais visuais**, procure o nome,
adicione e reabra o arquivo.

## 1. Preparar a base

1. Suba o Excel para o **OneDrive** ou **SharePoint** da empresa. Pode ser:
   - `Eventos.xlsx` deste pacote (tabela Excel chamada **Eventos**), ou
   - a planilha original (ex.: *DAI Eventual*): a primeira aba é lida e o cabeçalho
     (Mês, Data início, Data fim, Evento, Status, Transmissão LF, Observação) é localizado
     sozinho nas 10 primeiras linhas.
2. Abra o arquivo no **Excel do computador** → **Arquivo → Informações → Copiar caminho**.
   Esse é o valor do parâmetro `CaminhoExcel` (o final `?web=1` pode ficar; é ignorado).

O projeto usa o conector **SharePoint (pasta)**: a partir do caminho ele identifica o site
(`https://empresa.sharepoint.com/sites/NomeDoSite`) e procura o arquivo pelo nome. Por isso o
nome do arquivo deve ser único no site. Não funciona com link de compartilhamento
(`.../:x:/s/...`); use o "Copiar caminho".

## 2. Abrir o projeto

1. Power BI Desktop atualizado. Se ele recusar o projeto, ative em
   **Arquivo → Opções → Recursos de visualização**: *Opção de salvar Projeto do Power BI (.pbip)*
   e *Armazenar relatórios usando o formato de metadados aprimorado (PBIR)*; reinicie.
2. Abra `projeto/Eventos.pbip`.
3. **Página inicial → Transformar dados → Editar parâmetros** → em `CaminhoExcel`, cole o
   endereço do passo 1 → OK → **Aplicar alterações**.
4. Quando pedir credencial (SharePoint), escolha **Conta da Microsoft / organizacional** e entre
   com o e-mail da empresa.
5. Se aparecer "Objetos calculados precisam ser atualizados", clique em **Atualizar agora**.

### Erros comuns

| Mensagem | Causa e solução |
|---|---|
| `Web.Contents ... (404): Not Found` | Versão anterior do projeto. Use esta versão (lê pelo conector SharePoint). |
| `Arquivo não encontrado` | Nome do arquivo ou site diferente do parâmetro. Copie de novo pelo Excel → Copiar caminho. |
| Visuais com "problema de capacidade ou licença" | Os dados ainda não carregaram (corrija a fonte e clique em Atualizar) ou a empresa bloqueia visuais da AppSource: verifique se *Deneb* e *HTML Content* aparecem em Visualizações → … → Obter mais visuais. |
| Consulta extra (ex.: "DAI Eventual (Planilha1)") | Exclua consultas criadas à parte em Transformar dados; o relatório usa só **Eventos**. |

(Também funciona com arquivo local, ex.: `C:\Eventos\Eventos.xlsx`, mas aí a atualização
automática no serviço exige gateway.)

## 3. Publicar e compartilhar

1. **Página inicial → Publicar** → escolha o workspace.
2. No Power BI (web): conjunto de dados **Eventos → Configurações → Credenciais da fonte de
   dados** → OAuth2 / conta organizacional.
3. **Atualização agendada**: ligar, 1 vez por dia (ex.: 07:00).
4. Compartilhe o relatório ou o app do workspace só com quem deve ver.

## Regras (iguais à página web)

- **Próximo evento**: eventos em andamento hoje + todos os que começam na próxima data com
  evento; cancelados não entram; ignora os filtros.
- **Linha do tempo**: de hoje até 1, 2 ou 3 meses (filtro Período). Com o filtro **Mês**
  ativo, mostra o mês escolhido.
- **Calendário**: mês escolhido no filtro **Mês** (um só) ou o mês atual.
- **Mês**: filtra eventos que acontecem em qualquer dia do mês (inclui os que atravessam meses).
- **Legenda**: Confirmado (verde), A confirmar (amarelo), Cancelado (vermelho), outros (cinza).
- Linhas sem Evento ou Data início, ou com Data fim anterior à Data início, não entram.
- "Hoje" e os cards de destaque usam a data do momento da consulta; os dados são os da última
  atualização.

## Regerar o projeto

```
python3 powerbi/src/build_pbip.py
```
Edite `src/` ou `deneb/` antes; não edite `projeto/` à mão (é sobrescrito).
