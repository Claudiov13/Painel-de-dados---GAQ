// auth.js — Perfis de acesso (atualizar senhas aqui)
// null = administrador (vê tudo); string = filtra pelo nome exato do responsável
var PERFIS_ACESSO = {
  "GAQ#Visao26!@":  null,            // administrador — visão completa
  "Ale#26_Compr@":  "Alessandra",
  "Viv!Proc_26#":   "Viviane",
  "Wes@Ctrl#26!":   "Wesley",
  "Mau!Fluxo_26@":  "Mauricio",
  "Pam#Acesso26!":  "Pamela",
  "Bru@Painel_26#": "Bruna",
  "Glau!Dados#26":  "Gláucia",
  "Ina@Etapa_26!":  "Inacio",
  "Rafa#26_Compr@": "Raphael",
};

// Perfis Master — visão total com acesso destacado a Qualidade de Dados
var ADMIN_MASTER_PERFIS = {
  "AdminGAQ@": { nome: "Admin Master" },
};

// Senha do painel de Gestão/Performance (aba separada, acesso restrito)
var GESTAO_PASS = "GAQ@Gestao#26";

// Senhas de acesso por Área — filtram automaticamente processos da área
// Cada senha mapeia para o código exato da coluna ÁREA no dados.js
var AREA_PERFIS = {
  "GTI@Area26!#":   "GTI",
  "PLE@Area26!#":   "PLE",
  "ASCOM@Area26!#": "ASCOM",
  "PSC@Area26!#":   "PSC",
};

// Perfis de Diretoria — visão simplificada com áreas pré-definidas em destaque
// Cada senha mapeia para um objeto { nome, areasDestaque: [...] }
var DIRETOR_PERFIS = {
  "Alan@GAQ@": { nome: "Alan", areasDestaque: ["ASCOM", "TI", "GIN"] },
};
