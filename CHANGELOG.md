# Changelog

## 2026-04-30

### Visao Geral
- Separacao da regua global em `Critico`, `Em atencao`, `SLA vencido` e `SLA atendido`.
- Revisao dos cards para fechar a contagem entre `Com RC`, `SD sem RC`, criticidade e score da area.
- Inclusao do grafico percentual por subarea com distribuicao de processos por status da regua.
- Drill-downs mantidos e expandidos para os novos recortes.

### Diretoria
- Replicacao da `Visao Geral` do perfil Admin para o perfil Diretor.
- Manutencao do recorte por area do diretor, sem os blocos legados de NCL, CPL e Scont no rodape.
- Ajustes no diagnostico para abrir listas correctas de processos.

### Executivo
- Novo bloco de indicadores estrategicos com layout unificado ao restante sistema.
- Inclusao das caixas de `Processos abertos`, `Concluidos`, `Em andamento`, `Cancelados` e `Fracassados`.
- `Processos por modalidade` passou a considerar todo o recorte, e nao apenas os casos em andamento.
- `Ranking historico concluido` passou a mostrar tambem a distribuicao por modalidade de cada comprador.
- Remocao da caixa de mudancas na actualizacao e substituicao por leitura historica de compradores.

### Operacional e Processos
- Normalizacao da regra de criticidade para usar a mesma classificacao global de SLA e entrega.
- Ajuste de nomenclatura de `Parado` para `Ultimo Movimento`.
- Atualizacao da lista de processos para o novo padrao visual com score e aging percentual.

### Linha do Metro, Pantanal e Planejamento
- Revisao estrutural do scroll da aplicacao para evitar travamentos em abas longas.
- Ajustes de espacamento e apresentacao na `Linha do Metro`.
- Refinamento visual da aba `Planejamento` para o padrao mais clean da ferramenta.

### Seguranca do repositorio
- Mantida a exclusao de ficheiros sensiveis no Git:
  - `dados.js`
  - `tags.js`
  - `dados.json`
  - `tags.json`

### Admin Master e Qualidade de Dados
- Criacao do perfil `Admin Master`, com acesso total a base e destaque para a aba de qualidade.
- Inclusao da aba `Qualidade de Dados` com leitura aberta dos problemas de cadastro e integridade da base.
- Novo painel com cards de `Processos avaliados`, `Com pendencia`, `Erros criticos`, `Alertas de cadastro` e `Score da base`.
- Regras monitoradas com drill-down para:
  - `Datas incoerentes`
  - `RC sem recebimento`
  - `RC duplicada`
  - `Ativo sem responsavel`
  - `RC sem modalidade`
  - `Ativo com RC sem entrega`
  - `CPL sem responsavel`
  - `Ticket SD duplicado`
- Inclusao de leitura de pressao por subarea e exemplos prioritarios para triagem da base.
