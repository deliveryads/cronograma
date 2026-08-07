# Conectar a página ao Google Sheets (leitura e gravação ao vivo)

Sem isso, a página funciona só em modo leitura local (`data.json` ou dados de exemplo) —
os botões Criar/Editar/Excluir continuam funcionando, mas as alterações só existem
enquanto a aba do navegador estiver aberta. Com o Apps Script implantado, tudo é
lido e gravado direto na planilha.

## 1. Crie a planilha com as 7 abas

Nomes de aba **exatamente** como abaixo (acentos incluídos). Cada uma precisa da
linha de cabeçalho na primeira linha, com esses nomes de coluna exatos:

**Squad**

| id | nome | roteiros |
|---|---|---|
| sq-alpha | Squad Alpha | Coração em Chamas, Reencontro |

`roteiros` é uma única célula com os títulos separados por vírgula.

**Analistas**

| id | nome | squadId | funcao |
|---|---|---|---|
| an-01 | Camila Duarte | sq-alpha | Analista de Roteiros Pleno |

`squadId` precisa bater com um `id` da aba Squad.

**Férias**, **Cursos**, **Extraordinários**, **Datas Importantes** e **Ponto** — as
5 abas de eventos, todas com as mesmas colunas:

| id | titulo | analistaId | squadId | dataInicio | dataFim | descricao |
|---|---|---|---|---|---|---|
| ev-001 | Férias | an-01 | | 2026-01-12 | 2026-01-26 | Férias regulamentares |

`analistaId` e `squadId` são opcionais (podem ficar em branco); `dataInicio` e
`dataFim` no formato `AAAA-MM-DD` (ou célula de data do Sheets — o script converte).
O `id` pode ficar em branco em linhas novas: ao criar um evento pela página, o
Apps Script gera o `id` automaticamente. Se for popular a planilha manualmente
antes de usar a página, invente um `id` único por linha (ex.: `ev-001`, `ev-002`…).

Dica: use o `data.json` de exemplo deste repositório como referência de conteúdo —
os mesmos dados, só que organizados em abas em vez de um JSON único.

## 2. Cole o script

1. Na planilha, vá em **Extensões → Apps Script**.
2. Apague o conteúdo padrão de `Código.gs` e cole o conteúdo de `Code.gs`
   (raiz deste repositório).
3. Salve (ícone de disquete ou `Ctrl+S`).

## 3. Implante como Web App

1. No editor do Apps Script, clique em **Implantar → Nova implantação**.
2. Tipo: **App da Web**.
3. Configuração:
   - **Executar como:** Eu (sua conta — precisa ser dona ou editora da planilha)
   - **Quem pode acessar:** Qualquer pessoa
     (isso é necessário para a página estática conseguir chamar a API sem login;
     veja a nota de segurança abaixo)
4. Clique em **Implantar** e autorize as permissões pedidas (é a primeira vez
   que o script roda, o Google vai pedir para confirmar acesso à planilha).
5. Copie a **URL do app da Web** gerada (algo como
   `https://script.google.com/macros/s/AKfycb.../exec`).

## 4. Configure a URL na página

Abra `index.html` e edite a linha:

```js
var APPS_SCRIPT_URL = "";
```

para:

```js
var APPS_SCRIPT_URL = "https://script.google.com/macros/s/SEU_ID_AQUI/exec";
```

Salve e publique/suba o arquivo. A partir daí, ao abrir a página, o indicador no
topo passa a mostrar "Conectado à planilha (Google Sheets)" e os botões Criar
Evento, Editar e Excluir passam a gravar direto na planilha.

## Nota de segurança

Como decidido, o acesso está **livre**: qualquer pessoa com o link da página
publicada consegue criar, editar e excluir eventos na planilha, sem senha ou
login. Isso é aceitável para um cronograma interno de equipe, mas significa que
a URL da página (e a URL do Apps Script) não devem ser divulgadas publicamente.
Se depois quiser adicionar uma trava (senha simples ou login Google), isso pode
ser adicionado tanto no `index.html` quanto validado dentro do `doPost` do
`Code.gs` sem precisar redesenhar o restante do sistema.

## Toda vez que editar Code.gs

Reimplantar não atualiza a URL automaticamente. Depois de alterar o script:
**Implantar → Gerenciar implantações → (ícone de lápis) → Nova versão → Implantar**.
A URL existente continua funcionando; só o comportamento é atualizado.
