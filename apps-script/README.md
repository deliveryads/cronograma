# Eventos de curta duração — versão com acesso restrito (Google Apps Script)

A página é servida pelo Google (`script.google.com/...`) e lê a planilha **com a conta de quem
abre o link**. Só vê os eventos quem tiver a planilha compartilhada com o próprio e-mail.
A planilha deixa de ser pública.

Arquivos desta pasta:

| Arquivo | O que é |
|---|---|
| `Code.gs` | Servidor: entrega a página e lê a planilha |
| `Index.html` | A página (gerada de `eventos/index.html` por `tools/build_apps_script.py`) |
| `appsscript.json` | Manifesto: permissões e modo de implantação |

## 1. Criar o projeto (uma vez)

1. Abra a planilha → **Extensões → Apps Script**.
2. Em `Código.gs`, apague o conteúdo e cole o de `Code.gs`.
3. Clique em **+ → HTML**, nomeie o arquivo como `Index` (sem `.html`) e cole o conteúdo de `Index.html`.
4. **Configurações do projeto** (engrenagem) → marque **Mostrar arquivo de manifesto "appsscript.json"**.
   Volte ao Editor, abra `appsscript.json` e cole o conteúdo de `appsscript.json`.
5. Salve (Ctrl+S).

## 2. Implantar

1. **Implantar → Nova implantação → tipo: App da Web**.
2. **Executar como:** *Usuário que acessa o app da Web*.
3. **Quem pode acessar:** *Qualquer pessoa com uma Conta do Google*.
4. **Implantar** → autorize as permissões → copie a **URL do app da Web** (termina em `/exec`).

## 3. Fechar a planilha e liberar as pessoas

1. Na planilha: **Compartilhar → Acesso geral → Restrito**.
2. Adicione o e-mail de cada pessoa como **Leitor** (desmarque "Notificar", se preferir).
3. Envie a URL `/exec` para essas pessoas.

## Primeiro acesso de cada pessoa

1. Faz login numa Conta do Google (pode ser criada com o e-mail corporativo, sem Gmail).
2. O Google pede autorização para o app **ler planilhas** e **ver o e-mail**.
   Como o app não passou pela verificação do Google, aparece o aviso
   "O Google não verificou este app": clicar em **Avançado → Acessar (não seguro)** e **Permitir**.
   Isso acontece só uma vez por pessoa.
3. Quem não tem a planilha compartilhada vê "Você não tem acesso a este calendário".

Limite do Google para apps não verificados: até 100 pessoas diferentes.

## Atualizar a página depois de mudanças no código

1. Rode `python3 tools/build_apps_script.py` (ou pegue o `Index.html` atualizado no GitHub).
2. Cole o novo conteúdo em `Index` no Apps Script e salve.
3. **Implantar → Gerenciar implantações → lápis → Versão: Nova versão → Implantar.**
   A URL continua a mesma.

Alterações **nos dados** (planilha) não exigem nada disso: aparecem sozinhas em até 1 minuto.
