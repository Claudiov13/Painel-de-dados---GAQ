# Painel de dados - GAQ

Solucao local do painel de acompanhamento da GAQ. O painel e um `index.html`
estatico (React compilado para producao) que carrega os dados de arquivos `.js`
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

### Versao 6.1: distribuicao e manutencao

Os usuarios continuam abrindo o mesmo `index.html` na pasta local sincronizada
pelo OneDrive, com os mesmos perfis e senhas de `auth.js`. Aguarde a sincronizacao
completa antes de recarregar. A pasta `assets/` e obrigatoria. Nao e necessario
instalar Node nem manter acesso a uma CDN nas maquinas dos usuarios.

O uso diretamente no site do SharePoint depende das politicas de execucao de
HTML/JavaScript da organizacao; a verificacao desta versao foi feita por `file://`.

Codigo-fonte: `src/App.jsx`, `src/screens/` (24 telas), `src/components.jsx`,
`src/domain/service-desk.js`, `src/data-repository.js` e `src/access-log.js`.
As regras gerais continuam em `core.js`, a configuracao em `config.js` e as
exportacoes em `reports.js`. `index.html` e `assets/` sao saidas geradas.

Depois de editar codigo ou estilos, execute:

```powershell
.\geradores\compilar_painel.ps1
```

A compilacao requer Node 22.19+ apenas na maquina de manutencao. As dependencias
ficam em `%LOCALAPPDATA%\PainelGAQ\build`, fora da pasta compartilhada.
O build mantem as versoes das bibliotecas existentes e usa React de producao.
O primeiro preparo das ferramentas requer internet; o painel publicado funciona
com as bibliotecas locais. Nao edite os arquivos de `assets/` manualmente.

Depois de substituir manualmente qualquer base, execute:

```powershell
.\geradores\publicar_dados.ps1
```

Os geradores de Excel e MXM ja chamam essa publicacao. O publicador valida as
quatro bases, grava os arquivos por substituicao local e publica `publicacao.js`
por ultimo. Nao cria backups JSON permanentes. Mantenha uma unica maquina como
publicadora; o bloqueio de execucao e local, nao um bloqueio distribuido do OneDrive.
O manifesto identifica uma combinacao de arquivos, nao transforma fontes com
horarios distintos em uma transacao dos sistemas de origem.

O navegador confere as versoes antes de aplicar os dados. Durante uma sincronizacao
incompleta, mantem o conjunto anterior em memoria e informa a falha. Na primeira
abertura, e necessario concluir a sincronizacao para carregar o conjunto inteiro.
Sem alteracoes, a verificacao automatica le somente o manifesto; abas ocultas
pausam as verificacoes. Na virada do dia, os prazos sao recalculados.

Datas: "Base gerada" corresponde ao horario registrado pelo gerador, e
"Verificada" corresponde a leitura pelo navegador. Arquivos antigos sem metadados
mostram "Geracao nao informada". Importar uma planilha manualmente pausa a
sincronizacao automatica para nao substituir a importacao silenciosamente.

### Acessos (etapa 6)

O botao administrativo "Acessos locais" mostra entradas e saidas registradas
somente naquele navegador, permite exportar CSV e limita o historico a 500 eventos
dos ultimos 90 dias. O aviso aparece na tela de login. Nao registra senha, hash,
IP, conteudo de processos ou historico de navegacao. Se o armazenamento local
estiver bloqueado, o acesso ao painel continua funcionando.

Esse historico nao e uma auditoria central, pode ser apagado pelo usuario e nao
identifica individualmente pessoas que compartilham a mesma senha. Para consolidar
quem acessou e quando em todos os computadores, falta um receptor corporativo
autorizado e identidade individual (por exemplo, integracao autenticada com
SharePoint/Power Automate). Nenhum envio de telemetria foi ativado.

O Microsoft 365 registra atividades de arquivos e sincronizacao no Purview, mas
esses eventos nao equivalem a cada abertura do painel ja sincronizado no computador:
[atividades de auditoria Microsoft 365](https://learn.microsoft.com/en-us/purview/audit-log-activities).

### Verificacao tecnica

`npm test` executa testes de leitura coordenada, falhas, publicacao e registros
locais. `scripts/verify-browser.mjs` e `scripts/verify-profiles.mjs` usam uma sessao
descartavel do agent-browser; nao devem ser executados na sessao pessoal de um usuario.
Os perfis de teste sao adicionados apenas a memoria dessa aba, sem alterar `auth.js`.

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

Para executar localmente, mantenha esses ficheiros fora do Git e apenas no
ambiente interno autorizado.

## Repositorio GitHub

```text
https://github.com/Claudiov13/Painel-de-dados---GAQ.git
```

Antes de publicar, confira se os arquivos de dados acima continuam ignorados
pelo Git.
