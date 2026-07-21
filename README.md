# Painel de dados - GAQ

Solucao local do painel de acompanhamento da GAQ. O painel e um `index.html`
estatico (React via Babel standalone) que carrega os dados de arquivos `.js`
gerados a partir das planilhas do OneDrive (KB_GAQ).

## Estrutura de pastas

```text
Sistema de compras/
├── index.html                  Painel (aplicacao completa)
├── styles.css                  Estilos
├── core.js                     Engine de dados (parsing, SLA, scores)
├── config.js                   Constantes e colunas do timeline
├── reports.js                  Relatorios PDF/Excel
├── auth.js                     Perfis e senhas (hash)
├── apikey.js                   Chave da IA (nao versionado)
├── dados.js                    Base principal (gerado, nao versionado)
├── tags.js                     Tags dos processos (gerado, nao versionado)
├── base_fracionamento.js       Base MXM (gerado, nao versionado)
├── chamados_servicedesk.js     Log do Service Desk (gerado, nao versionado)
├── chamados_servicedesk.json   Idem, formato JSON (gerado, nao versionado)
├── geradores/                  Scripts que geram/atualizam os dados
│   ├── gerar_dados_js.ps1            Excel -> dados.js (raiz)
│   ├── executar_agora.bat            Roda o gerador manualmente
│   ├── configurar_agendamento.ps1    Cria tarefa agendada (07h e 19h)
│   ├── configurar_agendamento.bat    Idem, por duplo clique
│   ├── gerar_base_fracionamento.ps1  Excel MXM -> base_fracionamento.js (raiz)
│   └── executar_base_fracionamento.bat
├── outros/                     Arquivos sem uso aguardando revisao/exclusao
└── Versões anteriores/         Backups (nao versionado)
```

## Execucao

Abra o `index.html` diretamente no navegador (file:// ou SharePoint/OneDrive).
Nao ha servidor: os dados sao carregados dos arquivos `.js` na mesma pasta.

## Atualizacao da base

Manual:

```powershell
.\geradores\executar_agora.bat
```

Agendada (07h e 19h, todos os dias) — execute UMA vez nesta maquina:

```powershell
.\geradores\configurar_agendamento.bat
```

IMPORTANTE: os scripts foram movidos para `geradores/` em 2026-07-09. Se a
tarefa agendada "Painel GAQ - Atualizar dados.js" foi criada antes disso, ela
aponta para o caminho antigo e vai falhar — rode `configurar_agendamento.bat`
de novo para recria-la apontando para o novo caminho.

O gerador localiza a planilha "Base Consolidada GAQ IA.xlsx" na pasta KB_GAQ do
OneDrive e grava `dados.js` na raiz do projeto. O log fica em
`geradores/gerar_dados_js.log`. Quando a automacao do Excel nao estiver
disponivel, o script usa leitura direta do `.xlsx`.

Base de fracionamento (MXM):

```powershell
.\geradores\executar_base_fracionamento.bat
```

## Historico de alteracoes

Consulte [CHANGELOG.md](./CHANGELOG.md).

## Repositorio publico

Este repositorio guarda apenas o codigo e a configuracao da solucao.

Os ficheiros de dados reais nao sao versionados aqui:

- `dados.js`
- `tags.js`
- `base_fracionamento.js`
- `chamados_servicedesk.js` / `.json`
- `apikey.js`

Para executar localmente, mantenha esses ficheiros fora do Git e apenas no
ambiente interno autorizado.

## Repositorio GitHub

```text
https://github.com/Claudiov13/Painel-de-dados---GAQ.git
```

Antes de publicar, confira se os arquivos de dados acima continuam ignorados
pelo Git.
