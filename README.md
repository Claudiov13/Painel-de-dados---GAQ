# Painel de dados - GAQ

Solucao local do painel de acompanhamento da GAQ.

## Execucao local

1. Atualize os dados:

   ```powershell
   .\executar_agora.bat
   ```

2. Inicie o servidor:

   ```powershell
   .\INICIAR_SERVIDOR.bat
   ```

3. Acesse no navegador:

   ```text
   http://localhost:3000
   ```

O servidor tambem pode ser iniciado diretamente por PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File ".\iniciar_servidor.ps1"
```

## Node.js

O projeto procura o Node.js automaticamente nestes locais:

- `node` instalado no Windows e disponivel no PATH
- `node-portable\node.exe`
- `nodejs\node.exe`
- `tools\node\node.exe`
- instalacoes padrao em `Program Files`

Para usar Node portatil, consulte [INSTRUCOES_NODE_PORTATIL.txt](./INSTRUCOES_NODE_PORTATIL.txt).

## Servidor no boot do Windows

Para configurar o servidor para iniciar junto com o Windows:

```powershell
powershell -ExecutionPolicy Bypass -File ".\configurar_servidor_windows.ps1"
```

A tarefa agendada criada chama `iniciar_servidor.ps1` e grava logs em:

- `logs\server-out.log`
- `logs\server-err.log`

## Atualizacao da base

O arquivo [gerar_dados_js.ps1](./gerar_dados_js.ps1) localiza a planilha:

```text
C:\Users\cvduarte\OneDrive - Serviço Social do Comercio - Departamento Nacional\claudio – Pessoal\KB_GAQ\Base Consolidada GAQ IA.xlsx
```

Quando a automacao do Excel nao estiver disponivel, o script usa leitura direta do `.xlsx` e gera `dados.js` do mesmo jeito.

## Historico de alteracoes

Consulte [CHANGELOG.md](./CHANGELOG.md) para os registos das ultimas evolucoes da interface e das regras de negocio.

## Repositorio publico

Este repositorio guarda apenas o codigo e a configuracao da solucao.

Os ficheiros de dados reais nao sao versionados aqui:

- `dados.js`
- `tags.js`
- `dados.json`
- `tags.json`

Para executar localmente, mantenha esses ficheiros fora do Git e apenas no ambiente interno autorizado.

## Repositorio GitHub

Repositorio remoto configurado para este projeto:

```text
https://github.com/Claudiov13/Painel-de-dados---GAQ.git
```

Antes de publicar, confira se `dados.js`, `tags.js`, `node_modules`, `node-portable`, logs e arquivos `.zip` continuam ignorados pelo Git.
