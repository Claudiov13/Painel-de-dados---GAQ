/* ══════════════════════════════════════════════════════════════════════════════
   SEÇÃO 1 — CONSTANTES E CONFIGURAÇÃO
   ══════════════════════════════════════════════════════════════════════════════ */

var C_NCL = ["#1a5276","#2e86c1","#85c1e9","#aed6f1","#d6eaf8"];
var C_CPL = ["#6c3483","#8e44ad","#a569bd","#c39bd3","#d7bde2"];
var C_RSP = ["#1e8449","#27ae60","#52be80","#82e0aa","#abebc6"];
var IGNORAR = ["eduardo","renata"];

var TAG_COLORS = [
  "#e74c3c","#3498db","#2ecc71","#f39c12","#9b59b6",
  "#1abc9c","#e67e22","#34495e","#d35400","#2980b9",
  "#8e44ad","#16a085","#c0392b","#27ae60","#f1c40f",
];

var STOP = new Set([
  "para","com","por","dos","das","uma","uns","umas","que","nao","mais","como",
  "seu","sua","seus","suas","este","esta","esse","essa","esses","essas","sendo",
  "sera","pelo","pela","pelos","pelas","muito","outro","outra","outros","outras",
  "sobre","entre","todo","toda","todos","todas","isso","aqui","onde","quando",
  "tipo","item","conforme","processo","numero","data","prazo","valor","total",
  "referente","relativo","relativa","objeto","contrato","contratacao","aquisicao",
  "fornecimento","prestacao","servico","servicos","material","materiais","produto",
  "produtos","sistema","empresa","entrega","mediante","atraves","necessario",
  "necessaria","visando","tendo","sendo","inclusive","exceto","demais","demanda",
  "2020","2021","2022","2023","2024","2025","2026","janeiro","fevereiro","marco",
  "abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro",
  "setor","departamento","unidade","gerencia","diretoria","nacional","regional",
  "sesc","contratante","contratado","edital","licitacao","proposta","especificacao",
  "tecnica","qualidade","garantia","instalacao","manutencao","disposto",
  "decorrentes","eventual","eventuais","necessarios","forma","modo","termos",
  "condicoes","criterios","requisitos","minimos","comunicacao","informacao",
  "solicitacao","requisicao","pedido",
]);

var COL_HISTORICO_CPL = "HISTORICO_CPL_FINAL";
var COL_CPL_RECEBIDO_NCL = "RECEBIDO DO NCL EM";
var COL_CPL_ENVIADO_DJS_CHANCELA = "ENVIADO \u00C0 DJS PARA CHANCELA EM";

var SCONT_TIMELINE_KEYS = new Set([
  "DATA DE ENVIO DO PEDIDO (fornecedor) OU SUITE SESC (SCONT ou CPL)",
  "DATA DE RECEBIMENTO DA DEMANDA PELA ANALISTA DE CONTRATO",
  "DATA DE RECEBIMENTO NA DJ (QUANDO APLIC\u00C1VEL)",
  "DATA DE CONCLUSAO DE CONTRATO SCONT",
]);

function isCplTimelineKey(key) {
  return !!key && (
    key.startsWith("CPL_") ||
    key === COL_CPL_RECEBIDO_NCL ||
    key === COL_CPL_ENVIADO_DJS_CHANCELA
  );
}

function getTimelineSubareaByKey(key) {
  if (!key) return "";
  if (isCplTimelineKey(key)) return "CPL";
  if (SCONT_TIMELINE_KEYS.has(key)) return "Scont";
  return "NCL";
}

// Chave da nova fase final — vai começar a ser preenchida em processos de 2025+
var COL_CONCLUSAO_CONTRATO_SCONT = "DATA DE CONCLUSAO DE CONTRATO SCONT";

// ── Timeline LEGACY (ordem usada até 2024) ─────────────────────────────────
// Mantida para processos abertos em 2024 (anoSD === 2024 ou anoRC === 2024)
var TL_COLS_LEGACY = [
  ["Abertura SD","Data da abertura do SD"],
  ["Distribuição SD","Data da distribuição do SD"],
  ["Encerramento SD","Data do encerramento"],
  ["Recebimento RC","DATA DO RECEBIMENTO DA RC"],
  ["Planejamento RC","DATA DO PLANEJAMENTO DA RC"],
  ["Início propostas","Data inicial do envio de propostas"],
  ["Fim propostas","Data final do envio de propostas"],
  ["Envio aprovação","Data do envio para aprovação"],
  ["Última aprovação","Data da última aprovação"],
  ["Envio Pedido/Suite","DATA DE ENVIO DO PEDIDO (fornecedor) OU SUITE SESC (SCONT ou CPL)"],
  ["CPL: Recebido do NCL",COL_CPL_RECEBIDO_NCL],
  ["CPL: Enviado p/ DJS",COL_CPL_ENVIADO_DJS_CHANCELA],
  ["CPL: Recebimento","CPL_DATA_RECEBIMENTO_FINAL"],
  ["CPL: Publicação","CPL_PUBLICACAO_AVISO_FINAL"],
  ["CPL: Abertura Disputa","CPL_ABERTURA_DISPUTA_FINAL"],
  ["CPL: Fase ext.","CPL_FINALIZACAO_FASE_EXTERNA_FINAL"],
  ["CPL: Homologação","CPL_DATA_HOMOLOGACAO_FINAL"],
  ["Indicação Analista de Contrato","DATA DE RECEBIMENTO DA DEMANDA PELA ANALISTA DE CONTRATO"],
  ["Recebimento DJ","DATA DE RECEBIMENTO NA DJ (QUANDO APLICÁVEL)"],
];

// ── Timeline 2025+ (nova ordem) ────────────────────────────────────────────
// "Indicação Analista de Contrato" vem logo após "CPL: Recebido do NCL".
// "Recebimento DJ" vem logo após "CPL: Enviado p/ DJS".
// Nova fase final "Conclusão Contrato (Scont)".
var TL_COLS_2025 = [
  ["Abertura SD","Data da abertura do SD"],
  ["Distribuição SD","Data da distribuição do SD"],
  ["Encerramento SD","Data do encerramento"],
  ["Recebimento RC","DATA DO RECEBIMENTO DA RC"],
  ["Planejamento RC","DATA DO PLANEJAMENTO DA RC"],
  ["Início propostas","Data inicial do envio de propostas"],
  ["Fim propostas","Data final do envio de propostas"],
  ["Envio aprovação","Data do envio para aprovação"],
  ["Última aprovação","Data da última aprovação"],
  ["Envio Pedido/Suite","DATA DE ENVIO DO PEDIDO (fornecedor) OU SUITE SESC (SCONT ou CPL)"],
  ["CPL: Recebido do NCL",COL_CPL_RECEBIDO_NCL],
  ["Indicação Analista de Contrato","DATA DE RECEBIMENTO DA DEMANDA PELA ANALISTA DE CONTRATO"],
  ["CPL: Enviado p/ DJS",COL_CPL_ENVIADO_DJS_CHANCELA],
  ["Recebimento DJ","DATA DE RECEBIMENTO NA DJ (QUANDO APLICÁVEL)"],
  ["CPL: Recebimento","CPL_DATA_RECEBIMENTO_FINAL"],
  ["CPL: Publicação","CPL_PUBLICACAO_AVISO_FINAL"],
  ["CPL: Abertura Disputa","CPL_ABERTURA_DISPUTA_FINAL"],
  ["CPL: Fase ext.","CPL_FINALIZACAO_FASE_EXTERNA_FINAL"],
  ["CPL: Homologação","CPL_DATA_HOMOLOGACAO_FINAL"],
  ["Conclusão Contrato (Scont)", COL_CONCLUSAO_CONTRATO_SCONT],
];

// TL_COLS = default (usado por agregações que não têm proc específico).
// Sempre = TL_COLS_2025 para que novas fases entrem em cálculos globais.
var TL_COLS = TL_COLS_2025;

// Retorna a ordem correta de fases para um processo específico.
// Processos abertos em 2024 (SD ou RC) usam ordem LEGACY; demais usam 2025+.
function getTLColsForProc(proc) {
  if (proc && (proc.anoSD === 2024 || proc.anoRC === 2024)) return TL_COLS_LEGACY;
  return TL_COLS_2025;
}

var PAGE_SIZES = [20, 50, 100];

// Prazos gerais por tipo de modalidade (em dias úteis, a partir do Recebimento RC)
var PRAZO_LICITACAO = 90;      // Pregão, Concorrência
var PRAZO_COMPRA_DIRETA = 30;  // Dispensa, Inexigibilidade, Adesão ARP, Credenciamento
var PRAZO_SD = 10;             // Prazo da fase SD (Abertura SD → Encerramento SD) em dias úteis

// Determina o prazo geral (d.u.) com base na modalidade normalizada
function getPrazoGeral(modNrm) {
  if (modNrm.includes("pregao") || modNrm.includes("concorrencia") || modNrm.includes("dialogo")) return PRAZO_LICITACAO;
  return PRAZO_COMPRA_DIRETA;
}

// Etapas do timeline por tipo de modalidade (índices em TL_COLS do proc).
// CPL (Pregão/Concorrência): todas as etapas.
// Simples (Dispensa/Inexig/Adesão/Credenciamento): filtra as fases CPL pela key.
// Aceita proc opcional para usar ordem LEGACY/2025 correta; sem proc, usa default.
function getTLColsForMod(modNrm, proc) {
  var cols = proc ? getTLColsForProc(proc) : TL_COLS;
  if (modNrm.includes("pregao") || modNrm.includes("concorrencia") || modNrm.includes("dialogo")) {
    return cols.map(function (col, i) { return i; });
  }
  // Simples: filtra fases CPL identificadas pela key (não pelo índice — índices mudam entre LEGACY e 2025).
  return cols.map(function (col, i) { return { col: col, i: i }; })
             .filter(function (x) { return !isCplTimelineKey(x.col[1]); })
             .map(function (x) { return x.i; });
}

// Mapa de responsabilidade por fase (baseado no fluxo SESC GAQ)
var PHASE_RESP = {
  "Data da abertura do SD": "Área requisitante",
  "Data da distribuição do SD": "GAQ",
  "Data do encerramento": "Área requisitante / GAQ",
  "DATA DO RECEBIMENTO DA RC": "Área requisitante",
  "DATA DO PLANEJAMENTO DA RC": "GAQ",
  "Data inicial do envio de propostas": "GAQ",
  "Data final do envio de propostas": "Área requisitante / GAQ",
  "Data do envio para aprovação": "GAQ",
  "Data da última aprovação": "GAQ / Área requisitante / Diretoria / Presidente",
  "DATA DE ENVIO DO PEDIDO (fornecedor) OU SUITE SESC (SCONT ou CPL)": "GAQ / DJS",
  [COL_CPL_RECEBIDO_NCL]: "CPL",
  [COL_CPL_ENVIADO_DJS_CHANCELA]: "CPL / DJS",
  "CPL_DATA_RECEBIMENTO_FINAL": "GAQ / Diretoria / Presidente",
  "CPL_PUBLICACAO_AVISO_FINAL": "GAQ",
  "CPL_ABERTURA_DISPUTA_FINAL": "GAQ",
  "CPL_FINALIZACAO_FASE_EXTERNA_FINAL": "GAQ",
  "CPL_DATA_HOMOLOGACAO_FINAL": "GAQ / Diretoria / Presidente",
  "DATA DE RECEBIMENTO DA DEMANDA PELA ANALISTA DE CONTRATO": "GAQ",
  "DATA DE RECEBIMENTO NA DJ (QUANDO APLICÁVEL)": "DJ",
  "DATA DE CONCLUSAO DE CONTRATO SCONT": "Scont",
};

// ── FASE SD — PRAZO APARTADO (10 d.u. total da Abertura até Encerramento) ──
// Não faz parte do SLA de 30/90 dias. É o tempo de pré-compra.
var SLA_SD = {
  "Data da abertura do SD": 0,        // marco inicial
  "Data da distribuição do SD": 2,     // 2 d.u. para distribuir
  "Data do encerramento": 8,           // 8 d.u. para encerrar (total 10 d.u.)
};

// ── SLA_CPL / SLA_SIMPLES: valores FALLBACK (estáticos) ──
// Usados quando não há dados históricos. O sistema calcula SLA dinâmico
// a partir do histórico (phaseIntervals) quando disponível.
// IMPORTANTE: A contagem começa no Recebimento RC (marco zero do SLA).
// Fases SD não estão incluídas — têm prazo apartado de 10 d.u.

// Pregão e Concorrência — meta 90 d.u. a partir do Recebimento RC
var SLA_CPL = {
  "DATA DO RECEBIMENTO DA RC": 0,                                                     // Marco zero do SLA
  "DATA DO PLANEJAMENTO DA RC": 3,
  "Data inicial do envio de propostas": 7,
  "Data final do envio de propostas": 15,
  "Data do envio para aprovação": 3,
  "Data da última aprovação": 5,
  "DATA DE ENVIO DO PEDIDO (fornecedor) OU SUITE SESC (SCONT ou CPL)": 2,
  [COL_CPL_RECEBIDO_NCL]: null,
  [COL_CPL_ENVIADO_DJS_CHANCELA]: null,
  "CPL_DATA_RECEBIMENTO_FINAL": 2,
  "CPL_PUBLICACAO_AVISO_FINAL": 5,
  "CPL_ABERTURA_DISPUTA_FINAL": 12,
  "CPL_FINALIZACAO_FASE_EXTERNA_FINAL": 12,
  "CPL_DATA_HOMOLOGACAO_FINAL": 12,
  "DATA DE RECEBIMENTO DA DEMANDA PELA ANALISTA DE CONTRATO": 5,
  "DATA DE RECEBIMENTO NA DJ (QUANDO APLICÁVEL)": 7,
};
// Fallback total CPL: 3+7+15+3+5+2+2+5+12+12+12+5+7 = 90

// Demais modalidades (Dispensa, Inexigibilidade, Credenciamento, Adesão ARP) — meta 30 d.u. a partir do Recebimento RC
var SLA_SIMPLES = {
  "DATA DO RECEBIMENTO DA RC": 0,                                                     // Marco zero do SLA
  "DATA DO PLANEJAMENTO DA RC": 2,
  "Data inicial do envio de propostas": 4,
  "Data final do envio de propostas": 8,
  "Data do envio para aprovação": 3,
  "Data da última aprovação": 5,
  "DATA DE ENVIO DO PEDIDO (fornecedor) OU SUITE SESC (SCONT ou CPL)": 3,
  "DATA DE RECEBIMENTO DA DEMANDA PELA ANALISTA DE CONTRATO": 3,
  "DATA DE RECEBIMENTO NA DJ (QUANDO APLICÁVEL)": 2,
};
// Fallback total Simples: 2+4+8+3+5+3+3+2 = 30

// ── SLA DINÂMICO — calcula SLA por fase a partir do histórico real ──
// Pega as médias históricas de cada intervalo e redistribui proporcionalmente
// para que o total a partir do Recebimento RC = meta (90 ou 30 d.u.)
function buildDynamicSLA(phaseIntervals, modNrm) {
  if (!phaseIntervals || !phaseIntervals.global) return null;
  var isCPL = modNrm.includes("pregao") || modNrm.includes("concorrencia") || modNrm.includes("dialogo");
  var meta = isCPL ? PRAZO_LICITACAO : PRAZO_COMPRA_DIRETA;
  var baseSLA = isCPL ? SLA_CPL : SLA_SIMPLES;
  var keys = Object.keys(baseSLA);

  // Encontrar índice de "DATA DO RECEBIMENTO DA RC" em TL_COLS
  var rcIdx = -1;
  for (var i = 0; i < TL_COLS.length; i++) { if (TL_COLS[i][1] === "DATA DO RECEBIMENTO DA RC") { rcIdx = i; break; } }
  if (rcIdx < 0) return null;

  // Coletar médias históricas para cada transição entre fases (a partir da RC)
  var intervals = [];
  for (var i = 0; i < keys.length - 1; i++) {
    var keyA = keys[i], keyB = keys[i + 1];
    var pair = keyA + "\u2192" + keyB;
    // Tentar intervalo específico da modalidade, senão global
    var modData = phaseIntervals.byMod || {};
    var modIntervals = null;
    // Procurar por nome de modalidade que contenha o modNrm
    for (var mk in modData) {
      if (mk.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").includes(modNrm)) {
        modIntervals = modData[mk];
        break;
      }
    }
    var iv = (modIntervals && modIntervals[pair] && modIntervals[pair].avg) ? modIntervals[pair]
             : (phaseIntervals.global[pair] ? phaseIntervals.global[pair] : null);
    var fallback = baseSLA[keyB];
    intervals.push({ keyA: keyA, keyB: keyB, avg: iv ? iv.avg : fallback, n: iv ? iv.n : 0, fallback: fallback });
  }

  // Remover entradas com avg null (CPL null entries)
  var withAvg = intervals.filter(function(x) { return x.avg != null && x.avg > 0; });
  if (withAvg.length === 0) return null;

  // Soma das médias históricas
  var totalHist = withAvg.reduce(function(s, x) { return s + x.avg; }, 0);
  if (totalHist === 0) return null;

  // Redistribuir proporcionalmente para somar = meta
  var result = {};
  result[keys[0]] = 0; // Marco zero (Recebimento RC)
  for (var i = 0; i < intervals.length; i++) {
    var iv = intervals[i];
    if (iv.avg == null) {
      result[iv.keyB] = null; // CPL null entries mantêm null
    } else {
      result[iv.keyB] = Math.max(1, Math.round((iv.avg / totalHist) * meta));
    }
  }

  // Ajustar para que a soma exata = meta
  var valKeys = Object.keys(result).filter(function(k) { return result[k] != null && result[k] > 0; });
  var sum = valKeys.reduce(function(s, k) { return s + result[k]; }, 0);
  var diff = meta - sum;
  if (diff !== 0 && valKeys.length > 0) {
    // Ajusta na maior fase
    var maxKey = valKeys.reduce(function(a, b) { return result[a] >= result[b] ? a : b; });
    result[maxKey] = Math.max(1, result[maxKey] + diff);
  }

  return result;
}
