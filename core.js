/* ══════════════════════════════════════════════════════════════════════════════
   SEÇÃO 2 — ENGINE DE DADOS (parsing, cálculos, score, patterns)
   ══════════════════════════════════════════════════════════════════════════════ */

var nrm = s => s == null ? "" : s.toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim();

function pd(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  const n = Number(v);
  if (!isNaN(n) && n > 1000 && n < 100000) {
    const u = new Date((n - 25569) * 86400000);
    return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate());
  }
  if (typeof v === "string") {
    const m = v.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (m) { const y = m[3].length === 2 ? "20" + m[3] : m[3]; return new Date(+y, +m[2] - 1, +m[1]); }
    const iso = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return new Date(+iso[1], +iso[2] - 1, +iso[3]);
  }
  return null;
}

var fmt = v => { const d = pd(v); return d ? d.toLocaleDateString("pt-BR") : "—"; };

function du(a, b) {
  if (!a || !b) return 0;
  let n = 0; const c = new Date(a); c.setHours(0,0,0,0); const e = new Date(b); e.setHours(0,0,0,0);
  while (c < e) { const w = c.getDay(); if (w && w < 6) n++; c.setDate(c.getDate() + 1); }
  return n;
}

function duSigned(a, b) {
  if (!a || !b) return 0;
  const ai = new Date(a); ai.setHours(0,0,0,0);
  const bi = new Date(b); bi.setHours(0,0,0,0);
  if (ai.getTime() === bi.getTime()) return 0;
  return ai < bi ? du(ai, bi) : -du(bi, ai);
}

// OBS: Todo o painel trabalha em DIAS ÚTEIS (d.u.).
// Por consistência, a contagem "até a conclusão (a partir de hoje)" também deve ser em d.u.
function businessDaysFromToday(toDate) {
  if (!toDate) return 0;

  const hoje = new Date();
  hoje.setHours(0,0,0,0);

  const fim = new Date(toDate);
  fim.setHours(0,0,0,0);

  // usa a mesma lógica do du(): conta dias úteis EXCLUINDO o dia inicial e INCLUINDO o vencimento
  const n = du(hoje, fim);
  return n > 0 ? n : 0;
}

function businessDaysFromTodaySigned(toDate) {
  if (!toDate) return null;
  const hoje = new Date();
  hoje.setHours(0,0,0,0);
  return duSigned(hoje, toDate);
}

function kw(txt = "", minLen = 4) {
  const tokens = nrm(txt).split(/[\s,.\-\/\(\):;]+/).filter(p =>
    p.length >= minLen && !STOP.has(p) && !/^\d+$/.test(p) && !/^[a-z]{1,2}$/.test(p));
  const res = [];
  for (let i = 0; i < tokens.length - 2; i++)
    if (!STOP.has(tokens[i]) && !STOP.has(tokens[i+2]))
      res.push(tokens[i]+" "+tokens[i+1]+" "+tokens[i+2]);
  for (let i = 0; i < tokens.length - 3; i++)
    res.push(tokens[i]+" "+tokens[i+1]+" "+tokens[i+2]+" "+tokens[i+3]);
  for (let i = 0; i < tokens.length - 1; i++)
    if (tokens[i].length > 5 && tokens[i+1].length > 5 && !STOP.has(tokens[i]) && !STOP.has(tokens[i+1]))
      res.push(tokens[i]+" "+tokens[i+1]);
  return [...new Set(res)];
}

function matchObj(o, d) { return o && d ? kw(d, 4).some(p => nrm(o).includes(p)) : false; }

function parseBase(rows) {
  const hoje = new Date(); hoje.setHours(0,0,0,0);
  const DATE_COLS = [
    "Data da abertura do SD","Data da distribuição do SD","Data do encerramento",
    "DATA DO RECEBIMENTO DA RC","DATA DO PLANEJAMENTO DA RC",
    "Data inicial do envio de propostas","Data final do envio de propostas",
    "Data do envio para aprovação","Data da última aprovação",
    "DATA DE ENVIO DO PEDIDO (fornecedor) OU SUITE SESC (SCONT ou CPL)",
    COL_CPL_RECEBIDO_NCL,COL_CPL_ENVIADO_DJS_CHANCELA,
    "CPL_DATA_RECEBIMENTO_FINAL","CPL_PUBLICACAO_AVISO_FINAL",
    "CPL_ABERTURA_DISPUTA_FINAL","CPL_FINALIZACAO_FASE_EXTERNA_FINAL","CPL_DATA_HOMOLOGACAO_FINAL",
    "DATA DE RECEBIMENTO DA DEMANDA PELA ANALISTA DE CONTRATO","DATA DE RECEBIMENTO NA DJ (QUANDO APLICÁVEL)",
  ];
  return rows.map(r => {
    const row = { ...r };
    const findNormKey = (id) => Object.keys(row).find(k => nrm(k).replace(/[\s_\-]/g,"") === id);
    const recebNclKey = findNormKey("recebidodonclem") || findNormKey("cplrecebidonclfinal");
    const envDjsKey = findNormKey("enviadoadjsparachancelaem") || findNormKey("cplenviadodjschancelafinal");
    const histCplKey = findNormKey("historicocplfinal");
    if (recebNclKey && !(COL_CPL_RECEBIDO_NCL in row)) row[COL_CPL_RECEBIDO_NCL] = row[recebNclKey];
    if (envDjsKey && !(COL_CPL_ENVIADO_DJS_CHANCELA in row)) row[COL_CPL_ENVIADO_DJS_CHANCELA] = row[envDjsKey];
    if (histCplKey && !(COL_HISTORICO_CPL in row)) row[COL_HISTORICO_CPL] = row[histCplKey];
    r = row;
    // "HISTORICO DO PROCESSO" = sistema/NCL; "HISTÓRICO" = Scont (exato para não capturar HISTORICO_CPL_FINAL)
    const hk      = findNormKey("historicodoprocesso");
    const hkScont = findNormKey("historico");
    const pk = Object.keys(r).find(k => nrm(k).replace(/[\s_\-]/g,"").includes("pedido") && nrm(k).replace(/[\s_\-]/g,"").includes("suite"));
    const npk = Object.keys(r).find(k => { const kn = nrm(k); return kn.includes("cotacao") || (kn.includes("adesao") && kn.includes("arp")); });
    const Historico      = hk      ? (r[hk]      || "").toString().trim() : "";
    const HistoricoScont = hkScont ? (r[hkScont] || "").toString().trim() : "";
    const HistoricoCPL   = (r[COL_HISTORICO_CPL] || "").toString().trim();
    const NumPedidoSuite = pk ? (r[pk] || "").toString().trim() : "";
    const NumProcesso = npk ? (r[npk] || "").toString().trim() : "";
    const abertura = pd(r["Data da abertura do SD"]);
    const abRC = pd(r["DATA DO RECEBIMENTO DA RC"] || r["Data Recebimento RC"] || r["Recebimento RC"]);
    const receb = pd(r["DATA DO RECEBIMENTO DA RC"] || r["Data Recebimento RC"] || r["Recebimento RC"]);
    const encSD = pd(r["Data do encerramento"]);
    const ref = abertura || abRC;
    const leadTime = (abertura && receb) ? Math.max(0, du(abertura, receb)) : 0;
    const status = (r["STATUS SIMPLIFICADO"] || r["Status Simplificado"] || r.Status || "").toString().trim();
    const statusDet = (r["Key_StatusDetalhado"] || "").toString().trim();
    const sn = nrm(status);
    const emA = sn === "em andamento";
    const isCanceled = sn.includes("cancelad");
    const isFailed = sn.includes("fracass");
    const isConcluded = sn.includes("conclu");
    const isSuspended = sn.includes("suspens");
    const isEncerrado = isCanceled || isFailed || isConcluded || isSuspended;
    // AGING FIX: processos encerrados param na última data preenchida
    // ultimaData calculada para TODOS os processos (usada também para estagnação)
    let ultimaData = null;
    DATE_COLS.forEach(col => { const v = pd(r[col]); if (v && (!ultimaData || v > ultimaData)) ultimaData = v; });
    const rcAlt = pd(r["Data Recebimento RC"] || r["Recebimento RC"]);
    if (rcAlt && (!ultimaData || rcAlt > ultimaData)) ultimaData = rcAlt;
    let dataFimAging = hoje;
    if (isEncerrado && ref && ultimaData) dataFimAging = ultimaData;
    // diasTotaisCronograma: cálculo original desde o início (SD ou RC) — usado APENAS no Cronograma
    const diasTotaisCronograma = ref ? du(ref, dataFimAging) : 0;
    // diasTotais: AGING OFICIAL — começa sempre do Recebimento RC (nunca do SD)
    // processos sem RC têm diasTotais = 0 (ficam no Alerta SD)
    const diasTotais = abRC ? du(abRC, dataFimAging) : 0;
    const diasParado = (emA && ultimaData) ? du(ultimaData, hoje) : 0;
    const envRaw = r["DATA DE ENVIO DO PEDIDO (fornecedor) OU SUITE SESC (SCONT ou CPL)"];
    const envStr = (envRaw != null) ? envRaw.toString().trim() : "";
    const semEnvio = envStr === "" || envStr === "0" || envStr === "00/01/1900";
    // diasNcl: aging NCL operacional (RC → hoje) — igual ao diasRC para processos emA
    const diasNcl = (emA && abRC) ? du(abRC, hoje) : 0;
    const Modalidade = (r["MODALIDADE"] || r["Modalidade"] || "").toString().trim();
    const modNrm = nrm(Modalidade);
    const cplVal = nrm((r["CPL_ENCONTRADO"] || "").toString());
    const ehCPL = cplVal === "verdadeiro" || cplVal === "true" || cplVal === "1";
    const cplIni = pd(r[COL_CPL_RECEBIDO_NCL]) || pd(r[COL_CPL_ENVIADO_DJS_CHANCELA]) || pd(r["CPL_DATA_RECEBIMENTO_FINAL"]);
    const cplDisputa = pd(r["CPL_ABERTURA_DISPUTA_FINAL"]);
    const cplHomol = pd(r["CPL_DATA_HOMOLOGACAO_FINAL"]);
    const temFluxoCPL = ehCPL || modNrm.includes("pregao") || modNrm.includes("concorrencia") || !!(cplIni || cplDisputa || cplHomol);
    const diasCpl = (emA && ehCPL && cplIni) ? du(cplIni, cplDisputa || hoje) : 0;
    const diasCplTotal = (ehCPL && cplIni) ? du(cplIni, cplHomol || hoje) : 0;
    const anoSD = abertura ? abertura.getFullYear() : null;
    const anoRC = abRC ? abRC.getFullYear() : null;
    const Comprador = (r["COMPRADOR RESPONSÁVEL"] || r["Comprador Responsável"] || r.Comprador || "").toString().trim();
    const Pregoeiro = (r["CPL_PREGOEIRO_FINAL"] || "").toString().trim();
    const cplResp = (r["CPL_RESPONSAVEL_FINAL"] || "").toString().trim();
    const Avaliador = (r["Avaliador."] || r["Avaliador"] || "").toString().trim();
    const AnalistaContrato = (r["ANALISTA RESPONSÁVEL"] || "").toString().trim();
    // Duração específica da fase Scont: Envio Pedido/Suite → Indicação Analista de Contrato (ou hoje)
    const scontIni = pd(r["DATA DE ENVIO DO PEDIDO (fornecedor) OU SUITE SESC (SCONT ou CPL)"]);
    const scontFim = pd(r["DATA DE RECEBIMENTO NA DJ (QUANDO APLICÁVEL)"]);
    const analistaDt = pd(r["DATA DE RECEBIMENTO DA DEMANDA PELA ANALISTA DE CONTRATO"]);
    const diasScontFase = (AnalistaContrato && scontIni) ? du(scontIni, scontFim || hoje) : 0;
    const AdvogadoResp = (r["ADVOGADO RESPONSÁVEL"] || "").toString().trim();
    const EmpresaContratada = (r["EMPRESA CONTRATADA"] || "").toString().trim();
    // NCL: usa Comprador, se vazio faz fallback para Avaliador
    const respNCL = Comprador || Avaliador;
    const isNonCPL = modNrm.includes("credenciamento") || modNrm.includes("dispensa") || modNrm.includes("inexigibilidade");
    // NCL phase duration — data fim depende da modalidade:
    // Simples (DL, Inexig, Credenc, Adesão ARP): max(scontIni, analistaDt)
    // Complexas (Pregão, Concorrência): max(cplHomol, analistaDt, scontFim)
    const isSimpleModal = modNrm.includes("inexig") || modNrm.includes("dispensa") || modNrm.includes("credenci") || modNrm.includes("adesao");
    const _fimNCLCandidates = isSimpleModal
      ? [scontIni, analistaDt]
      : [cplHomol, analistaDt, scontFim];
    const fimNCL = _fimNCLCandidates.filter(Boolean).reduce((mx, d) => (!mx || d > mx) ? d : mx, null);
    const diasNCLFase = ref ? du(ref, fimNCL || (emA ? hoje : (ultimaData || hoje))) : 0;
    // REGRA OBJETO: se tem OBJETO RC usa essa, senão Objeto SD
    const objetoRC = (r["OBJETO RC"] || "").toString().trim();
    const objetoSD = (r["Objeto SD"] || r["Objeto"] || r["Descrição"] || r.Descricao || "").toString().trim();
    const Objeto = objetoRC || objetoSD;
    const NumRC = (r["Nº RC"] || r["Nº do Processo"] || "").toString().trim();
    const TicketSD = (r["Ticket#"] || "").toString().trim();
    const ProcessKey = (r["ProcessKey"] || NumRC || TicketSD || ("PK_" + Math.random().toString(36).slice(2,8))).toString().trim();
    const AreaReq = (r["Área_Normalizada"] || r["Area_Normalizada"] || r["Área Requisitante"] || r["Area Requisitante"] || "").toString().trim();
    const dataAbertura = abertura || abRC;
    const mesAbertura = dataAbertura ? dataAbertura.getMonth() : null;
    // Fase atual e responsável pela próxima etapa pendente
    // Apenas para processos em andamento — Fracassados, Cancelados e Suspensos ficam vazios
    let faseAtual = "—", respFase = "—", faseAtualIdx = -1;
    if (emA) {
      // Fases CPL são ignoradas em processos não-CPL (SCONT/NCL puro)
      let ultimaFasePreenchidaIdx = -1;
      for (let fi = TL_COLS.length - 1; fi >= 0; fi--) {
        if (!temFluxoCPL && isCplTimelineKey(TL_COLS[fi][1])) continue;
        if (pd(r[TL_COLS[fi][1]])) { ultimaFasePreenchidaIdx = fi; break; }
      }
      if (ultimaFasePreenchidaIdx >= 0) {
        let proximaFaseVaziaIdx = -1;
        for (let fi = ultimaFasePreenchidaIdx + 1; fi < TL_COLS.length; fi++) {
          if (!temFluxoCPL && isCplTimelineKey(TL_COLS[fi][1])) continue;
          if (!pd(r[TL_COLS[fi][1]])) { proximaFaseVaziaIdx = fi; break; }
        }
        const faseRefIdx = proximaFaseVaziaIdx >= 0 ? proximaFaseVaziaIdx : ultimaFasePreenchidaIdx;
        faseAtual = TL_COLS[faseRefIdx][0];
        respFase = PHASE_RESP[TL_COLS[faseRefIdx][1]] || "—";
        faseAtualIdx = faseRefIdx;
      }
    }
    const faseKey = faseAtualIdx < 0 ? "" : TL_COLS[faseAtualIdx][1];
    const faseSubarea = getTimelineSubareaByKey(faseKey);
    // Responsável atual: segue subárea em andamento; encerrados → "N/A"
    const respAtivo = !emA ? "N/A"
      : faseSubarea === "NCL"   ? (Comprador || Avaliador || "—")
      : faseSubarea === "CPL"   ? (cplResp || "—")
      : faseSubarea === "Scont" ? (AnalistaContrato || AdvogadoResp || "—")
      : "—";
    // Aging específico desde a RC
    const diasRC = (emA && abRC) ? du(abRC, hoje) : 0;
    // Aging específico desde a abertura do SD (para processos só com SD sem RC)
    const diasSD = (emA && abertura) ? du(abertura, hoje) : 0;
    // Processo só tem ticket (SD) sem RC
    const somenteSD = !!(TicketSD && !NumRC);
    // Previsão de entrega / início de prestação
    const dataEntrega = pd(r["PREVISÃO ENTREGA / INÍCIO PRESTAÇÃO DO SERVIÇO DA RC"] || r["Data prevista entrega / início prestação do serviço"] || r["PREVISAO ENTREGA"] || r["DATA ENTREGA"]);
    const diasParaEntrega = dataEntrega ? (function(){ const d = new Date(dataEntrega); d.setHours(0,0,0,0); const h2 = new Date(hoje); h2.setHours(0,0,0,0); return Math.round((d - h2) / 86400000); })() : null;
    const fonteAba = nrm((r["FonteAba"] || r["Fonte Aba"] || r["FONTE_ABA"] || r["fonte_aba"] || "").toString().trim());
    const isLegado = fonteAba.includes("antes de abril 2025") || fonteAba.includes("a partir de abril 2025");
    // GESTÃO: aging de gestão parte do Recebimento RC (não do SD)
    // Fim = última data preenchida para TODOS os processos (encerrados e em andamento)
    // Processos sem RC não geram diasTotaisGestao (ficam no Alerta SD)
    const dataFimAgingGestao = ultimaData || hoje;
    const diasTotaisGestao = abRC ? du(abRC, dataFimAgingGestao) : 0;
    const criticoNclGestao = diasTotaisGestao > 50;
    // Aging SD: Abertura SD → Encerramento SD (indicador NCL gestão de pessoas)
    const diasAgingSD = (abertura && encSD) ? du(abertura, encSD) : 0;
    // ── PRAZO GERAL por modalidade (30 d.u. compra direta, 90 d.u. licitação) ──
    // Conta a partir do Recebimento RC. Fase SD tem prazo separado de 10 d.u.
    const prazoGeral = getPrazoGeral(modNrm);
    // diasSDAberto: dias úteis desde abertura SD até hoje (SD ainda aberto, sem encerramento)
    const diasSDAberto = (emA && abertura && !encSD) ? du(abertura, hoje) : 0;
    // atrasoSD: SD aberto por mais de 10 d.u. (prazo apartado da fase SD)
    const atrasoSD = diasSDAberto > PRAZO_SD;
    // atrasoGeral: processo com RC que já ultrapassou o prazo geral da modalidade
    const atrasoGeral = emA && abRC && diasTotais > prazoGeral;
    // diasDesdeEncSD: d.u. desde o encerramento do SD sem RC aberta (responsabilidade da área requisitante)
    const diasDesdeEncSD = (encSD && !abRC) ? du(encSD, hoje) : 0;
    return {
      ...r, ProcessKey, Comprador, Pregoeiro, cplResp, Avaliador, AnalistaContrato, AdvogadoResp, EmpresaContratada, respNCL, NumPedidoSuite, NumProcesso, Historico, HistoricoScont, HistoricoCPL,
      "Área Requisitante": AreaReq, Modalidade, Objeto, NumRC, TicketSD, statusDet,
      LeadTime: leadTime, diasTotais, emA, status, isCanceled, isFailed, isConcluded, isSuspended, isEncerrado,
      diasNcl, criticoNcl: diasNcl > 50, semEnvio, ultimaData, diasParado, isNonCPL, mesAbertura,
      ehCPL, diasCpl, diasCplTotal, diasScontFase, diasNCLFase, criticoCpl: diasCpl > 50, anoSD, anoRC, _ref: ref,
      dataAbertura, aberturaSD: abertura, aberturaRC: abRC,
      faseAtual, faseSubarea, respAtivo, respFase, faseAtualIdx, diasRC, diasSD, somenteSD,
      dataEntrega, diasParaEntrega, isLegado,
      diasTotaisGestao, criticoNclGestao, encSD, diasAgingSD, diasTotaisCronograma,
      prazoGeral, diasSDAberto, atrasoSD, atrasoGeral, diasDesdeEncSD,
    };
  });
}

function calcularScore(comp, hist, backlog, dem, pool) {
  const h = hist.filter(r => pool === "cpl" ? (r.Pregoeiro === comp || r.cplResp === comp) : r.respNCL === comp);
  const tot = h.length || 1;
  const simArea = h.filter(r => r["Área Requisitante"] === dem.area).length;
  const simMod = h.filter(r => nrm(r.Modalidade || "") === nrm(dem.modalidade || "")).length;
  const simObj = h.filter(r => matchObj(r.Objeto, dem.objeto)).length;
  const sA = Math.min(35, Math.min(15, (simArea/tot)*45) + Math.min(10, (simMod/tot)*30) + Math.min(10, (simObj/tot)*30));
  const t = h.filter(r => r.LeadTime > 0).map(r => r.LeadTime);
  const med = t.length ? t.reduce((a,b) => a+b, 0) / t.length : 999;
  const vr = t.length > 1 ? t.reduce((a,b) => a + (b-med)**2, 0) / t.length : 999;
  const sB = (t.length ? Math.max(0, 20 - med/10) : 5) + (t.length > 1 ? Math.max(0, 10 - Math.sqrt(vr)/5) : 3);
  const bk = backlog.filter(r => pool === "cpl"
    ? (r.Pregoeiro === comp || r.cplResp === comp) && r.faseSubarea === "CPL"
    : r.respNCL === comp && r.faseSubarea === "NCL");
  const carga = bk.length;
  const crit = bk.filter(r => pool === "cpl" ? r.criticoCpl : r.criticoNcl).length;
  const sC = Math.max(0, 15 - carga*1.5) + Math.max(0, 10 - crit*2);
  const modN = nrm(dem.modalidade || "");
  const ehLic = ["pregao","concorrencia","dialogo competitivo"].some(m => modN.includes(m));
  const hMod = hist.filter(r => nrm(r.Modalidade || "") === modN);
  const taxaCF = hMod.length > 4 ? hMod.filter(r => r.isCanceled || r.isFailed).length / hMod.length : 0;
  const hCM = h.filter(r => nrm(r.Modalidade || "") === modN);
  let sD = ehLic ? (simMod > 0 ? 6 : 2) : 3;
  if (taxaCF > 0.25) sD += hCM.some(r => r.isCanceled || r.isFailed) ? 3 : -1;
  if (simObj > 0) sD += 2;
  return {
    total: Math.min(100, Math.round(sA+sB+sC+sD)), sA: Math.round(sA), sB: Math.round(sB), sC: Math.round(sC), sD,
    med: med === 999 ? null : Math.round(med), carga, crit, simArea, simMod, simObj, taxaCF: Math.round(taxaCF*100),
  };
}

function buildPatterns(procs, minOcorr) {
  const kwMap = new Map();
  procs.forEach(r => {
    if (!r.Objeto || r.Objeto.length < 15) return;
    [...new Set(kw(r.Objeto, 4))].forEach(w => {
      const wTokens = w.split(" ");
      if (wTokens.length < 2) return;
      // Requer ao menos 1 token com mais de 6 chars para evitar matches genéricos
      if (!wTokens.some(t => t.length > 6)) return;
      if (!kwMap.has(w)) kwMap.set(w, []);
      if (!kwMap.get(w).find(x => x.NumRC && x.NumRC === r.NumRC && x.NumRC !== "")) kwMap.get(w).push(r);
    });
  });
  const entries = [...kwMap.entries()].filter(([w, c]) => {
    const uniqCount = new Set(c.map(x => x.NumRC || x.Objeto + x.status)).size;
    return uniqCount >= minOcorr && w.split(" ").length >= 2;
  });
  return entries
    .map(([word, casos]) => {
      const uniq = [...new Map(casos.map(c => [(c.NumRC || c.Objeto + c.status), c])).values()];
      const lts = uniq.filter(c => c.LeadTime > 0).map(c => c.LeadTime);
      const wordCount = word.split(" ").length;
      return {
        word, count: uniq.length,
        cancelados: uniq.filter(c => c.isCanceled).length,
        fracassados: uniq.filter(c => c.isFailed).length,
        nums: [...new Set(uniq.map(c => c.NumRC).filter(Boolean))],
        areas: [...new Set(uniq.map(c => c["Área Requisitante"]).filter(Boolean))].slice(0, 3),
        mods: [...new Set(uniq.map(c => c.Modalidade).filter(Boolean))].slice(0, 3),
        exemplos: [...new Set(uniq.map(c => c.Objeto).filter(Boolean))].slice(0, 4),
        avgLT: lts.length ? Math.round(lts.reduce((a,b) => a+b, 0) / lts.length) : null,
        compradores: [...new Set(uniq.map(c => c.Comprador || c.Pregoeiro).filter(Boolean))].slice(0, 5),
        _score: uniq.length * 10 + (wordCount >= 4 ? 15 : wordCount >= 3 ? 8 : 0) + (word.length > 25 ? 5 : 0),
      };
    })
    .sort((a, b) => b._score - a._score)
    .slice(0, 25);
}

function calcMapInsights(base) {
  const canceled = base.filter(r => r.isCanceled), failed = base.filter(r => r.isFailed);
  const concl = base.filter(r => r.isConcluded);
  const prob = [...canceled, ...failed];
  if (!prob.length) return { padroes: [], padFracasso: [], porArea: [], porMod: [], taxas: [], reincidentes: [], tempoMedio: [], evolucaoAnual: [] };

  const areasAll = [...new Set(base.map(r => r["Área Requisitante"]).filter(Boolean))];
  const porArea = areasAll.map(a => {
    const tot = base.filter(r => r["Área Requisitante"] === a).length;
    const cf = base.filter(r => r["Área Requisitante"] === a && (r.isCanceled || r.isFailed)).length;
    return { name: a, cf, tot, taxa: tot > 3 ? Math.round((cf/tot)*100) : null };
  }).filter(x => x.cf > 0 && x.taxa !== null).sort((a,b) => b.taxa - a.taxa).slice(0, 10);

  const modsAll = [...new Set(base.map(r => nrm(r.Modalidade)).filter(Boolean))];
  const porMod = modsAll.map(mn => {
    const am = base.filter(r => nrm(r.Modalidade) === mn);
    const tot = am.length, cf = am.filter(r => r.isCanceled || r.isFailed).length;
    const lts = am.filter(r => r.LeadTime > 0).map(r => r.LeadTime);
    return {
      name: am[0]?.Modalidade || mn, cf, tot,
      taxa: tot > 3 ? Math.round((cf/tot)*100) : null,
      medLT: lts.length ? Math.round(lts.reduce((a,b) => a+b, 0) / lts.length) : null,
    };
  }).filter(x => x.cf > 0 && x.taxa !== null).sort((a,b) => b.taxa - a.taxa).slice(0, 8);

  // NOVO: Reincidentes — objetos que fracassaram/cancelaram mais de uma vez
  const objMap = new Map();
  prob.forEach(r => {
    if (!r.Objeto || r.Objeto.length < 20) return;
    const key = nrm(r.Objeto).slice(0, 80);
    if (!objMap.has(key)) objMap.set(key, []);
    objMap.get(key).push(r);
  });
  const reincidentes = [...objMap.entries()]
    .filter(([_, arr]) => arr.length >= 2)
    .map(([key, arr]) => ({
      objeto: arr[0].Objeto.slice(0, 120),
      count: arr.length,
      nums: arr.map(r => r.NumRC).filter(Boolean),
      areas: [...new Set(arr.map(r => r["Área Requisitante"]).filter(Boolean))],
      cancelados: arr.filter(r => r.isCanceled).length,
      fracassados: arr.filter(r => r.isFailed).length,
    }))
    .sort((a,b) => b.count - a.count)
    .slice(0, 10);

  // NOVO: Tempo médio até cancelamento/fracasso por modalidade
  const tempoMedio = modsAll.map(mn => {
    const cfProcs = base.filter(r => nrm(r.Modalidade) === mn && (r.isCanceled || r.isFailed) && r.diasTotais > 0);
    const conclProcs = base.filter(r => nrm(r.Modalidade) === mn && r.isConcluded && r.diasTotais > 0);
    if (cfProcs.length < 2) return null;
    return {
      name: cfProcs[0]?.Modalidade || mn,
      medCF: Math.round(cfProcs.reduce((a,b) => a + b.diasTotais, 0) / cfProcs.length),
      medConc: conclProcs.length ? Math.round(conclProcs.reduce((a,b) => a + b.diasTotais, 0) / conclProcs.length) : null,
      qtdCF: cfProcs.length,
    };
  }).filter(Boolean).sort((a,b) => b.medCF - a.medCF);

  // NOVO: Evolução anual de cancelamentos/fracassos
  const anos = [...new Set(base.map(r => r.anoSD || r.anoRC).filter(Boolean))].sort();
  const evolucaoAnual = anos.map(ano => {
    const total = base.filter(r => (r.anoSD || r.anoRC) === ano).length;
    const cf = base.filter(r => (r.anoSD || r.anoRC) === ano && (r.isCanceled || r.isFailed)).length;
    return { ano, total, cf, taxa: total > 0 ? Math.round((cf/total)*100) : 0 };
  });

  const totBase = base.length || 1;
  return {
    padroes: buildPatterns(prob, 3), padFracasso: buildPatterns(failed, 2), porArea, porMod,
    reincidentes, tempoMedio, evolucaoAnual,
    taxas: [
      { label: "Taxa geral cancelamento", val: Math.round((canceled.length/totBase)*100) + "%", num: canceled.length, color: "#e67e22" },
      { label: "Taxa geral fracasso", val: Math.round((failed.length/totBase)*100) + "%", num: failed.length, color: "#c0392b" },
      { label: "Processos problemáticos", val: prob.length, num: prob.length, color: "#8e44ad" },
      { label: "% sobre total", val: Math.round((prob.length/totBase)*100) + "%", num: base.length, color: "#1a5276" },
    ],
  };
}

/* ══════════════════════════════════════════════════════════════════════════════
   SEÇÃO 3 — GERENCIADOR DE TAGS (localStorage)
   ══════════════════════════════════════════════════════════════════════════════ */

var TagsManager = {
  KEY: "painel_gaq_tags",
  LIST_KEY: "painel_gaq_tag_list",

  getAll() { try { return JSON.parse(localStorage.getItem(this.KEY) || "{}"); } catch { return {}; } },
  // pk = ProcessKey principal, altPk = TicketSD (chave secundária para processos que eram só SD)
  getForProcess(pk, altPk) {
    if (!pk && !altPk) return [];
    const all = this.getAll();
    const tagsPk  = (pk && all[pk]) ? all[pk] : [];
    const tagsAlt = (altPk && altPk !== pk && all[altPk]) ? all[altPk] : [];
    // Merge sem duplicatas: prioriza pk, complementa com altPk
    if (tagsAlt.length > 0 && tagsPk.length === 0 && pk) {
      // Auto-migração: copia tags da chave SD para a chave ProcessKey atual
      all[pk] = [...tagsAlt];
      delete all[altPk];
      localStorage.setItem(this.KEY, JSON.stringify(all));
      return all[pk];
    }
    if (tagsAlt.length > 0 && tagsPk.length > 0) {
      const merged = [...new Set([...tagsPk, ...tagsAlt])];
      // Consolida tudo na chave principal
      if (pk) { all[pk] = merged; delete all[altPk]; localStorage.setItem(this.KEY, JSON.stringify(all)); }
      return merged;
    }
    return tagsPk;
  },
  setForProcess(pk, tags) {
    const all = this.getAll(); all[pk] = tags;
    localStorage.setItem(this.KEY, JSON.stringify(all));
  },
  getTagList() { try { return JSON.parse(localStorage.getItem(this.LIST_KEY) || "[]"); } catch { return []; } },
  addToTagList(tag) {
    const list = this.getTagList();
    if (!list.includes(tag)) { list.push(tag); localStorage.setItem(this.LIST_KEY, JSON.stringify(list)); }
  },
  removeFromTagList(tag) {
    const list = this.getTagList().filter(t => t !== tag);
    localStorage.setItem(this.LIST_KEY, JSON.stringify(list));
  },
  addTag(pk, tag, altPk) {
    if (!pk) return;
    const tags = this.getForProcess(pk, altPk);
    if (!tags.includes(tag)) { tags.push(tag); this.setForProcess(pk, tags); this.addToTagList(tag); }
  },
  removeTag(pk, tag, altPk) {
    if (!pk) return;
    const updated = this.getForProcess(pk, altPk).filter(t => t !== tag);
    this.setForProcess(pk, updated);
  },
  getTagColor(tag) {
    let h = 0; for (let i = 0; i < tag.length; i++) h = tag.charCodeAt(i) + ((h << 5) - h);
    return TAG_COLORS[Math.abs(h) % TAG_COLORS.length];
  },
  exportTags() { return { tags: this.getAll(), tagList: this.getTagList() }; },
  importTags(data) {
    if (data && data.tags) localStorage.setItem(this.KEY, JSON.stringify(data.tags));
    if (data && data.tagList) localStorage.setItem(this.LIST_KEY, JSON.stringify(data.tagList));
  },
};

/* ══════════════════════════════════════════════════════════════════════════════
   SEÇÃO 3b — WATCHLIST / ACOMPANHAMENTO (localStorage)
   ══════════════════════════════════════════════════════════════════════════════ */

var WatchlistManager = {
  KEY: "painel_gaq_watchlist",
  SNAP_KEY: "painel_gaq_watchlist_snaps",
  getList() { try { return JSON.parse(localStorage.getItem(this.KEY) || "[]"); } catch { return []; } },
  isWatched(pk) { return this.getList().includes(pk); },
  toggle(pk, proc) {
    const list = this.getList();
    if (list.includes(pk)) {
      localStorage.setItem(this.KEY, JSON.stringify(list.filter(x => x !== pk)));
      const snaps = this.getSnaps(); delete snaps[pk]; localStorage.setItem(this.SNAP_KEY, JSON.stringify(snaps));
      return false;
    } else {
      list.push(pk);
      localStorage.setItem(this.KEY, JSON.stringify(list));
      if (proc) this.saveSnap(pk, this.createSnap(proc));
      return true;
    }
  },
  getSnaps() { try { return JSON.parse(localStorage.getItem(this.SNAP_KEY) || "{}"); } catch { return {}; } },
  saveSnap(pk, snap) { const s = this.getSnaps(); s[pk] = snap; localStorage.setItem(this.SNAP_KEY, JSON.stringify(s)); },
  createSnap(proc) {
    const dates = {};
    TL_COLS.forEach(([, key]) => { dates[key] = pd(proc[key]) ? fmt(pd(proc[key])) : null; });
    return { status: proc.status || "", statusDet: proc.statusDet || "", diasTotais: proc.diasTotais || 0,
      semEnvio: proc.semEnvio, dates, savedAt: new Date().toISOString() };
  },
  detectChanges(base) {
    const snaps = this.getSnaps();
    const list = this.getList();
    const changes = [];
    list.forEach(pk => {
      const proc = base.find(r => r.ProcessKey === pk);
      const snap = snaps[pk];
      if (!proc) return; // processo removido da base
      if (!snap) { this.saveSnap(pk, this.createSnap(proc)); return; } // sem snapshot, criar
      const ch = { pk, proc, items: [] };
      if (proc.status !== snap.status) ch.items.push({ type: "status", icon: "🔄", desc: `Status: "${snap.status}" → "${proc.status}"` });
      if ((proc.statusDet || "") !== (snap.statusDet || "")) ch.items.push({ type: "statusDet", icon: "📋", desc: `Status Det.: "${snap.statusDet || "—"}" → "${proc.statusDet || "—"}"` });
      TL_COLS.forEach(([label, key]) => {
        const nowFilled = pd(proc[key]) ? fmt(pd(proc[key])) : null;
        const wasFilled = snap.dates[key] || null;
        if (nowFilled && !wasFilled) ch.items.push({ type: "date", icon: "📅", desc: `Nova data preenchida: ${label} = ${nowFilled}` });
      });
      if (proc.semEnvio !== snap.semEnvio && !proc.semEnvio) ch.items.push({ type: "pedido", icon: "📦", desc: "Pedido/Suite enviado!" });
      if (ch.items.length > 0) changes.push(ch);
    });
    return changes;
  },
  acknowledgeAll(base) {
    const list = this.getList();
    list.forEach(pk => {
      const proc = base.find(r => r.ProcessKey === pk);
      if (proc) this.saveSnap(pk, this.createSnap(proc));
    });
  },
  count() { return this.getList().length; },
};

/* ══════════════════════════════════════════════════════════════════════════════
   SEÇÃO 3c — CÁLCULO DE INTERVALOS MÉDIOS POR FASE (para previsão na Timeline)
   ══════════════════════════════════════════════════════════════════════════════ */

// Índice "Envio Pedido/Suite" em TL_COLS — modalidades sem CPL param aqui
var IDX_ENVIO_PEDIDO = 9;

function calcPhaseIntervals(base) {
  // FIX: usar TODOS os processos que tenham ambas as datas preenchidas (não apenas concluídos)
  // Isso aumenta drasticamente a amostragem para SLA por Fase e previsões
  const allProcs = base;
  const byMod = {}, global = {};
  for (let i = 0; i < TL_COLS.length - 1; i++) {
    const [, keyA] = TL_COLS[i];
    const [, keyB] = TL_COLS[i + 1];
    const pair = keyA + "\u2192" + keyB;
    allProcs.forEach(r => {
      const dA = pd(r[keyA]), dB = pd(r[keyB]);
      if (!dA || !dB || dB <= dA) return;
      const iv = du(dA, dB);
      if (iv <= 0 || iv > 500) return;
      const mod = r.Modalidade || "N/I";
      if (!byMod[mod]) byMod[mod] = {};
      if (!byMod[mod][pair]) byMod[mod][pair] = [];
      byMod[mod][pair].push(iv);
      if (!global[pair]) global[pair] = [];
      global[pair].push(iv);
    });
  }
  const avg = arr => arr.length >= 2 ? Math.round(arr.reduce((a,b) => a+b, 0) / arr.length) : null;
  const med = arr => { if (arr.length < 2) return null; const s = [...arr].sort((a,b) => a-b); return s[Math.floor(s.length/2)]; };
  const out = { byMod: {}, global: {} };
  Object.entries(byMod).forEach(([mod, pairs]) => {
    out.byMod[mod] = {};
    Object.entries(pairs).forEach(([p, vals]) => { out.byMod[mod][p] = { avg: avg(vals), med: med(vals), n: vals.length }; });
  });
  Object.entries(global).forEach(([p, vals]) => { out.global[p] = { avg: avg(vals), med: med(vals), n: vals.length }; });
  // Risco por modalidade para health score
  const _modRisk = {};
  const modTot = {}, modProb = {};
  base.forEach(r => { const m = r.Modalidade || "N/I"; modTot[m] = (modTot[m]||0)+1; if (r.isCanceled || r.isFailed) modProb[m] = (modProb[m]||0)+1; });
  Object.keys(modTot).forEach(m => { _modRisk[m] = modTot[m] > 3 ? Math.round(((modProb[m]||0)/modTot[m])*100) : 0; });
  out._modRisk = _modRisk;
  return out;
}

/* ── Score de Saúde do Processo (0-100) ─────────────────────────── */
function calcHealthScore(proc, phaseIntervals) {
  if (!proc.emA) return null;
  let score = 100;
  // Aging total (max -35)
  const d = proc.diasTotais || 0;
  if (d > 100) score -= 35; else if (d > 80) score -= 28; else if (d > 50) score -= 20; else if (d > 30) score -= 10; else if (d > 20) score -= 5;
  // Estagnação (max -30)
  const dp = proc.diasParado || 0;
  if (dp > 30) score -= 30; else if (dp > 20) score -= 22; else if (dp > 15) score -= 15; else if (dp > 10) score -= 8; else if (dp > 5) score -= 3;
  // Progresso de fase (max -20)
  const maxIdx = proc.isNonCPL ? IDX_ENVIO_PEDIDO : TL_COLS.length - 1;
  let filled = 0;
  for (let i = 0; i <= maxIdx; i++) { if (pd(proc[TL_COLS[i][1]])) filled++; }
  const pct = (maxIdx + 1) > 0 ? filled / (maxIdx + 1) : 0;
  if (d > 30 && pct < 0.3) score -= 20; else if (d > 20 && pct < 0.4) score -= 12; else if (d > 15 && pct < 0.5) score -= 6;
  // Risco modalidade (max -15)
  if (phaseIntervals && phaseIntervals._modRisk) {
    const risk = phaseIntervals._modRisk[proc.Modalidade || "N/I"] || 0;
    if (risk > 40) score -= 15; else if (risk > 25) score -= 10; else if (risk > 15) score -= 5;
  }
  return Math.max(0, Math.min(100, score));
}

/* ── Score de Complexidade do Processo (1–10) ────────────────────────────── */
function calcComplexidade(proc, phaseIntervals) {
  const modN = nrm(proc.Modalidade || "");
  const isLic     = modN.includes("pregao") || modN.includes("concorrencia") || modN.includes("dialogo");
  const isMed     = modN.includes("credenciamento");
  const isSimples = modN.includes("dispensa") || modN.includes("inexig") || modN.includes("adesao");
  let score = isLic ? 7 : isMed ? 5 : isSimples ? 3 : 4;
  const modRisk = (phaseIntervals?._modRisk?.[proc.Modalidade || "N/I"] || 0);
  if (modRisk > 30) score = Math.min(10, score + 2);
  else if (modRisk > 15) score = Math.min(10, score + 1);
  if ((proc.Objeto || "").length > 200) score = Math.min(10, score + 1);
  if (proc.ehCPL) score = Math.min(10, score + 1);
  return Math.min(10, Math.max(1, score));
}

/* ── Probabilidade de Atraso (0–100 %) ──────────────────────────────────── */
function calcDelayProb(proc, phaseIntervals) {
  if (!proc.emA) return null;
  const hs = calcHealthScore(proc, phaseIntervals);
  if (hs === null) return null;
  const modRisk = phaseIntervals?._modRisk?.[proc.Modalidade || "N/I"] || 0;
  const prob = Math.round((100 - hs) * 0.65 + modRisk * 0.35);
  return Math.min(95, Math.max(5, prob));
}

/* ── Classificação SLA Global ───────────────────────────────────────────── */
// Regra única para Visão Geral, Executivo, Timeline e Diagnóstico:
// - Crítico: SLA vencido E entrega já vencida ou projeção já fora do cronograma.
// - SLA Vencido: fora do SLA, mas ainda sem ruptura no cronograma de entrega.
// - Atenção: próximo de entrar em crítico pela margem de SLA e/ou de entrega.
// - SLA Atendido: dentro do SLA e sem pressão relevante de entrega.

var SLA_SCORE_RULES = {
  base: 1000,
  critico: -30,
  atencao: -15,
  sla_vencido: -10,
  sla_atendido: 30,
  attentionThresholdDU: 20,
  sdWorstCaseDU: PRAZO_LICITACAO,
};

function estimateRemainingBusinessDays(proc, phaseIntervals) {
  if (!proc || !proc.emA) return null;

  const hasRC = !!(proc.NumRC && String(proc.NumRC).trim() !== "");
  const prazoModalidade = proc.prazoGeral || getPrazoGeral(nrm(proc.Modalidade || ""));

  if (!hasRC) {
    return {
      total: SLA_SCORE_RULES.sdWorstCaseDU,
      source: "sd_worst_case",
      pairs: 0,
      usedFallback: true,
    };
  }

  const mod = proc.Modalidade || "N/I";
  const modData = phaseIntervals?.byMod?.[mod] || {};
  const globalData = phaseIntervals?.global || {};
  const maxPredIdx = proc.isNonCPL ? IDX_ENVIO_PEDIDO : TL_COLS.length - 1;

  let lastFilledIdx = -1;
  for (let i = 0; i <= maxPredIdx; i++) {
    if (pd(proc[TL_COLS[i][1]])) lastFilledIdx = i;
  }

  if (lastFilledIdx < 0) {
    return {
      total: Math.max(0, prazoModalidade - (proc.diasTotais || 0)),
      source: "sla_restante",
      pairs: 0,
      usedFallback: true,
    };
  }

  let total = 0;
  let pairs = 0;
  let pendingWithoutHistory = 0;

  for (let i = lastFilledIdx; i < maxPredIdx; i++) {
    const keyA = TL_COLS[i][1];
    const keyB = TL_COLS[i + 1][1];
    if (pd(proc[keyB])) continue;

    const pair = keyA + "→" + keyB;
    const sample = modData[pair] || globalData[pair];
    const days = sample ? (sample.med || sample.avg || null) : null;

    if (days != null) {
      total += days;
      pairs++;
    } else {
      pendingWithoutHistory++;
    }
  }

  if (pendingWithoutHistory > 0) {
    const remainingSla = Math.max(0, prazoModalidade - (proc.diasTotais || 0));
    const fallbackPerStep = pendingWithoutHistory > 0 ? Math.max(1, Math.round(remainingSla / pendingWithoutHistory)) : 0;
    total += fallbackPerStep * pendingWithoutHistory;
  }

  if (total <= 0) {
    total = Math.max(0, prazoModalidade - (proc.diasTotais || 0));
  }

  return {
    total: Math.max(0, Math.round(total)),
    source: pairs > 0 ? "historico_fases" : "sla_restante",
    pairs,
    usedFallback: pendingWithoutHistory > 0,
  };
}

function calcProcessSlaScore(sla) {
  if (!sla) return null;
  const points =
    sla.bucket === "critico" ? SLA_SCORE_RULES.critico :
    sla.bucket === "atencao" ? SLA_SCORE_RULES.atencao :
    sla.bucket === "sla_vencido" ? SLA_SCORE_RULES.sla_vencido :
    SLA_SCORE_RULES.sla_atendido;
  return Math.max(0, Math.min(100, Math.round(((points + 30) / 60) * 100)));
}

function calcAreaSlaScore(rows, phaseIntervals) {
  const classified = (rows || []).map(r => {
    if (!r) return null;
    return r._sla ? r : { ...r, _sla: calcSLAClassification(r, phaseIntervals) };
  }).filter(r => r && r._sla);

  const counts = { critico: 0, atencao: 0, sla_vencido: 0, no_prazo: 0 };
  classified.forEach(r => {
    if (counts[r._sla.bucket] == null) counts[r._sla.bucket] = 0;
    counts[r._sla.bucket]++;
  });

  const ativos = classified.length;
  if (!ativos) {
    return {
      score: 0,
      raw: SLA_SCORE_RULES.base,
      counts,
      weights: SLA_SCORE_RULES,
    };
  }

  const raw =
    SLA_SCORE_RULES.base +
    (counts.no_prazo * SLA_SCORE_RULES.sla_atendido) +
    (counts.sla_vencido * SLA_SCORE_RULES.sla_vencido) +
    (counts.atencao * SLA_SCORE_RULES.atencao) +
    (counts.critico * SLA_SCORE_RULES.critico);

  const minRaw = SLA_SCORE_RULES.base + (ativos * SLA_SCORE_RULES.critico);
  const maxRaw = SLA_SCORE_RULES.base + (ativos * SLA_SCORE_RULES.sla_atendido);
  const score = maxRaw === minRaw ? 100 : Math.round(((raw - minRaw) / (maxRaw - minRaw)) * 100);

  return {
    score: Math.max(0, Math.min(100, score)),
    raw,
    counts,
    weights: SLA_SCORE_RULES,
  };
}

function calcSLAClassification(proc, phaseIntervals, options) {
  if (!proc || !proc.emA) return null;

  options = options || {};
  const attentionThresholdDU = options.attentionThresholdDU || SLA_SCORE_RULES.attentionThresholdDU;
  const hasRC = !!(proc.NumRC && String(proc.NumRC).trim() !== "");
  const slaPrazo = hasRC ? (proc.prazoGeral || getPrazoGeral(nrm(proc.Modalidade || ""))) : PRAZO_SD;
  const diasConsumidos = hasRC ? (proc.diasTotais || 0) : (proc.diasSDAberto || proc.diasSD || 0);
  const pctSLA = slaPrazo > 0 ? (diasConsumidos / slaPrazo) : 0;
  const slaVencido = hasRC ? !!proc.atrasoGeral : !!proc.atrasoSD;

  const projection = estimateRemainingBusinessDays(proc, phaseIntervals);
  const diasRestantesProjetados = projection ? projection.total : (hasRC ? Math.max(0, slaPrazo - diasConsumidos) : SLA_SCORE_RULES.sdWorstCaseDU);

  const temEntrega = !!proc.dataEntrega;
  const diasEntregaUteis = temEntrega ? businessDaysFromTodaySigned(proc.dataEntrega) : null;
  const entregaVencida = diasEntregaUteis != null && diasEntregaUteis < 0;
  const margemEntregaDU = diasEntregaUteis == null ? null : (diasEntregaUteis - diasRestantesProjetados);
  const projecaoForaCronograma = margemEntregaDU != null && margemEntregaDU < 0;
  const margemSLADU = hasRC ? (slaPrazo - (diasConsumidos + diasRestantesProjetados)) : null;
  const janelaComprometida = !hasRC ? false : (margemEntregaDU != null && margemEntregaDU >= 0 && margemEntregaDU <= attentionThresholdDU);
  const atencaoPrazo = !hasRC ? false : (margemSLADU != null && margemSLADU >= 0 && margemSLADU <= attentionThresholdDU);
  const atencaoEntrega = !hasRC ? false : (
    entregaVencida ||
    projecaoForaCronograma ||
    janelaComprometida ||
    (diasEntregaUteis != null && diasEntregaUteis >= 0 && diasEntregaUteis <= attentionThresholdDU)
  );

  let bucket = "no_prazo";
  let detail = "Dentro do SLA e sem pressão relevante de entrega.";

  if (slaVencido && (entregaVencida || projecaoForaCronograma)) {
    bucket = "critico";
    detail = entregaVencida
      ? "SLA vencido e entrega já vencida."
      : "SLA vencido e projeção já fora do cronograma de entrega.";
  } else if (slaVencido) {
    bucket = "sla_vencido";
    detail = hasRC
      ? "Fora do SLA da modalidade, mas ainda dentro do cronograma de entrega."
      : "Pré-compra acima de 10 d.u., mas ainda sem ruptura de entrega pelo cenário de 90 d.u.";
  } else if (!hasRC) {
    bucket = "no_prazo";
    detail = "Pré-compra dentro do prazo de 10 d.u.";
  } else if (atencaoPrazo || atencaoEntrega) {
    bucket = "atencao";
    if (projecaoForaCronograma) detail = "Projeção de entrega já pressionada pelo ritmo atual.";
    else if (janelaComprometida) detail = `Margem de entrega em até ${attentionThresholdDU} d.u.`;
    else if (atencaoPrazo) detail = `Margem de SLA em até ${attentionThresholdDU} d.u.`;
    else detail = `Entrega prevista em até ${attentionThresholdDU} d.u.`;
  }

  const visual = bucket === "critico"
    ? { color: "#d35400", bg: "#fff3e8", label: "Crítico", rank: 4 }
    : bucket === "atencao"
      ? { color: "#ff9500", bg: "#fff7e6", label: "Atenção", rank: 3 }
      : bucket === "sla_vencido"
        ? { color: "#ff3b30", bg: "#fff5f5", label: "SLA Vencido", rank: 2 }
        : { color: "#34c759", bg: "#eafaf1", label: "SLA Atendido", rank: 1 };

  const procScore = calcProcessSlaScore({ bucket });

  return {
    bucket,
    label: visual.label,
    rank: visual.rank,
    color: visual.color,
    bg: visual.bg,
    detail,
    slaVencido,
    temEntrega,
    entregaVencida,
    projecaoForaCronograma,
    janelaComprometida,
    atencaoPrazo,
    atencaoEntrega,
    pctSLA,
    pctSLALabel: Math.round(pctSLA * 100),
    slaPrazo,
    diasConsumidos,
    diasRestantesProjetados,
    diasEntregaUteis,
    diasParaEntrega: proc.diasParaEntrega,
    margemEntregaDU,
    margemSLADU,
    attentionThresholdDU,
    hasRC,
    sdWorstCaseDU: !hasRC ? SLA_SCORE_RULES.sdWorstCaseDU : null,
    projection,
    processScore: procScore,
  };
}

// Labels e cores padronizadas para uso global
var SLA_BUCKETS = {
  critico:     { label: "Crítico",       color: "#d35400", bg: "#fff3e8" },
  atencao:     { label: "Atenção",       color: "#ff9500", bg: "#fff7e6" },
  sla_vencido: { label: "SLA Vencido",   color: "#ff3b30", bg: "#fff5f5" },
  no_prazo:    { label: "SLA Atendido",  color: "#34c759", bg: "#eafaf1" },
};

// Overrides da regua global de SLA/entrega.
function calcProcessSlaScore(sla) {
  if (!sla) return null;
  const points =
    sla.bucket === "critico" ? SLA_SCORE_RULES.critico :
    sla.bucket === "atencao" ? SLA_SCORE_RULES.atencao :
    sla.bucket === "sla_vencido" ? SLA_SCORE_RULES.sla_vencido :
    SLA_SCORE_RULES.sla_atendido;
  const minPoints = SLA_SCORE_RULES.critico;
  const maxPoints = SLA_SCORE_RULES.sla_atendido;
  return Math.max(0, Math.min(100, Math.round(((points - minPoints) / (maxPoints - minPoints)) * 100)));
}

function calcAreaSlaScore(rows, phaseIntervals) {
  const classificados = (rows || []).map(r => {
    if (!r) return null;
    return r._sla ? r : { ...r, _sla: calcSLAClassification(r, phaseIntervals) };
  }).filter(r => r && r._sla);

  const counts = { critico: 0, atencao: 0, sla_vencido: 0, no_prazo: 0 };
  classificados.forEach(r => {
    if (counts[r._sla.bucket] == null) counts[r._sla.bucket] = 0;
    counts[r._sla.bucket]++;
  });

  const ativos = classificados.length;
  if (!ativos) {
    return {
      score: 0,
      raw: SLA_SCORE_RULES.base,
      counts,
      weights: SLA_SCORE_RULES,
      detail: "Sem processos ativos no recorte.",
    };
  }

  const raw =
    SLA_SCORE_RULES.base +
    (counts.no_prazo * SLA_SCORE_RULES.sla_atendido) +
    (counts.sla_vencido * SLA_SCORE_RULES.sla_vencido) +
    (counts.atencao * SLA_SCORE_RULES.atencao) +
    (counts.critico * SLA_SCORE_RULES.critico);

  const minRaw = SLA_SCORE_RULES.base + (ativos * SLA_SCORE_RULES.critico);
  const maxRaw = SLA_SCORE_RULES.base + (ativos * SLA_SCORE_RULES.sla_atendido);
  const score = maxRaw === minRaw ? 100 : Math.round(((raw - minRaw) / (maxRaw - minRaw)) * 100);

  return {
    score: Math.max(0, Math.min(100, score)),
    raw,
    counts,
    weights: SLA_SCORE_RULES,
    detail: `Base 1000 + ${counts.no_prazo}x(${SLA_SCORE_RULES.sla_atendido}) + ${counts.sla_vencido}x(${SLA_SCORE_RULES.sla_vencido}) + ${counts.atencao}x(${SLA_SCORE_RULES.atencao}) + ${counts.critico}x(${SLA_SCORE_RULES.critico})`,
  };
}

function calcSLAClassification(proc, phaseIntervals, options) {
  if (!proc || !proc.emA) return null;

  options = options || {};
  const attentionThresholdDU = options.attentionThresholdDU || SLA_SCORE_RULES.attentionThresholdDU;
  const hasRC = !!(proc.NumRC && String(proc.NumRC).trim() !== "");
  const slaPrazo = hasRC ? (proc.prazoGeral || getPrazoGeral(nrm(proc.Modalidade || ""))) : PRAZO_SD;
  const diasConsumidos = hasRC ? (proc.diasTotais || 0) : (proc.diasSDAberto || proc.diasSD || 0);
  const pctSLA = slaPrazo > 0 ? (diasConsumidos / slaPrazo) : 0;
  const slaVencido = hasRC ? !!proc.atrasoGeral : !!proc.atrasoSD;
  const projection = estimateRemainingBusinessDays(proc, phaseIntervals);
  const diasRestantesProjetados = projection ? projection.total : (hasRC ? Math.max(0, slaPrazo - diasConsumidos) : SLA_SCORE_RULES.sdWorstCaseDU);
  const temEntrega = !!proc.dataEntrega;
  const diasEntregaUteis = temEntrega ? businessDaysFromTodaySigned(proc.dataEntrega) : null;
  const entregaVencida = diasEntregaUteis != null && diasEntregaUteis < 0;
  const margemEntregaDU = diasEntregaUteis == null ? null : (diasEntregaUteis - diasRestantesProjetados);
  const projecaoForaCronograma = margemEntregaDU != null && margemEntregaDU < 0;
  const margemSLADU = hasRC ? (slaPrazo - (diasConsumidos + diasRestantesProjetados)) : null;
  const projecaoEstouraSla = !hasRC ? false : (margemSLADU != null && margemSLADU < 0);
  const janelaComprometida = !hasRC ? false : (margemEntregaDU != null && margemEntregaDU <= attentionThresholdDU);
  const atencaoPrazo = !hasRC ? false : (margemSLADU != null && margemSLADU <= attentionThresholdDU);
  const atencaoEntrega = !hasRC ? false : (margemEntregaDU != null && margemEntregaDU <= attentionThresholdDU);
  const entregaProxima = !hasRC ? false : (diasEntregaUteis != null && diasEntregaUteis >= 0 && diasEntregaUteis <= attentionThresholdDU);

  let bucket = "no_prazo";
  let detail = temEntrega ? "Dentro do SLA e do cronograma de entrega." : "Dentro do SLA atual.";

  if (slaVencido && (entregaVencida || projecaoForaCronograma)) {
    bucket = "critico";
    detail = entregaVencida
      ? "SLA vencido e entrega ja vencida."
      : "SLA vencido e projecao de conclusao fora do cronograma de entrega.";
  } else if (slaVencido) {
    bucket = "sla_vencido";
    detail = hasRC
      ? "Fora do SLA da modalidade, mas ainda dentro do cronograma de entrega."
      : "Pré-compra acima de 10 d.u.; entrega avaliada no pior cenario de 90 d.u.";
  } else if (!hasRC) {
    bucket = "no_prazo";
    detail = "Pré-compra dentro do prazo de 10 d.u.";
  } else if (atencaoPrazo || atencaoEntrega || projecaoEstouraSla) {
    bucket = "atencao";
    if (projecaoEstouraSla && atencaoEntrega) detail = `Margem curta para SLA e entrega (ate ${attentionThresholdDU} d.u.).`;
    else if (projecaoEstouraSla) detail = "Ritmo atual projeta ruptura do SLA antes da conclusao.";
    else if (margemEntregaDU != null && margemEntregaDU < 0) detail = "Entrega pressionada: projecao de conclusao acima da data prevista.";
    else if (janelaComprometida && atencaoPrazo) detail = `Margem curta para SLA e entrega (ate ${attentionThresholdDU} d.u.).`;
    else if (janelaComprometida) detail = `Margem de entrega em ate ${attentionThresholdDU} d.u.`;
    else detail = `Margem de SLA em ate ${attentionThresholdDU} d.u.`;
  }

  const visual = bucket === "critico"
    ? { color: "#d35400", bg: "#fff3e8", label: "Crítico", rank: 4 }
    : bucket === "atencao"
      ? { color: "#ff9500", bg: "#fff7e6", label: "Atenção", rank: 3 }
      : bucket === "sla_vencido"
        ? { color: "#ff3b30", bg: "#fff5f5", label: "SLA Vencido", rank: 2 }
        : { color: "#34c759", bg: "#eafaf1", label: "SLA Atendido", rank: 1 };

  return {
    bucket,
    label: visual.label,
    rank: visual.rank,
    color: visual.color,
    bg: visual.bg,
    detail,
    slaVencido,
    temEntrega,
    entregaVencida,
    projecaoForaCronograma,
    projecaoEstouraSla,
    janelaComprometida,
    atencaoPrazo,
    atencaoEntrega,
    entregaProxima,
    pctSLA,
    pctSLALabel: Math.round(pctSLA * 100),
    prazo: slaPrazo,
    slaPrazo,
    diasConsumidos,
    diasRestantesProjetados,
    diasEntregaUteis,
    diasParaEntrega: proc.diasParaEntrega,
    margemEntregaDU,
    margemSLADU,
    attentionThresholdDU,
    hasRC,
    sdWorstCaseDU: !hasRC ? SLA_SCORE_RULES.sdWorstCaseDU : null,
    projection,
    processScore: calcProcessSlaScore({ bucket }),
  };
}

SLA_BUCKETS = {
  critico:     { label: "Crítico",      color: "#d35400", bg: "#fff3e8" },
  atencao:     { label: "Atenção",      color: "#ff9500", bg: "#fff7e6" },
  sla_vencido: { label: "SLA Vencido",  color: "#ff3b30", bg: "#fff5f5" },
  no_prazo:    { label: "SLA Atendido", color: "#34c759", bg: "#eafaf1" },
};

function calcDirectorPriority(proc) {
  if (!proc || !proc.emA || !proc.NumRC) return null;

  const prazo = proc.prazoGeral || getPrazoGeral(nrm(proc.Modalidade || ""));
  const diasConsumidos = proc.diasTotais || 0;
  const pctPrazo = prazo > 0 ? (diasConsumidos / prazo) : 0;
  const diasRestantesPrazo = prazo > 0 ? Math.max(0, prazo - diasConsumidos) : 0;
  const diasEntregaUteis = proc.dataEntrega ? businessDaysFromTodaySigned(proc.dataEntrega) : null;
  const entregaVencida = diasEntregaUteis != null && diasEntregaUteis < 0;
  const folgaEntrega = (diasEntregaUteis == null) ? null : (diasEntregaUteis - diasRestantesPrazo);

  let score = 0;

  if (pctPrazo >= 1.2) score += 55;
  else if (pctPrazo >= 1) score += 45;
  else if (pctPrazo >= 0.85) score += 35;
  else if (pctPrazo >= 0.7) score += 24;
  else if (pctPrazo >= 0.5) score += 12;
  else score += Math.round(pctPrazo * 10);

  if (diasEntregaUteis != null) {
    if (entregaVencida) score += 45;
    else if (diasEntregaUteis <= 5) score += 34;
    else if (diasEntregaUteis <= 10) score += 26;
    else if (diasEntregaUteis <= 20) score += 18;
    else if (diasEntregaUteis <= 40) score += 10;
    else score += 4;
  }

  if (folgaEntrega != null) {
    if (folgaEntrega < 0) score += 18;
    else if (folgaEntrega <= 5) score += 12;
    else if (folgaEntrega <= 10) score += 6;
  }

  if (proc.atrasoGeral) score += 10;
  if ((proc.diasParado || 0) > 10) score += 5;

  score = Math.max(0, Math.min(100, score));

  let bucket = "controlado";
  if (proc.atrasoGeral || entregaVencida || (folgaEntrega != null && folgaEntrega < 0) || score >= 80) bucket = "critico";
  else if (score >= 55 || pctPrazo >= 0.7 || (diasEntregaUteis != null && diasEntregaUteis <= 15)) bucket = "atencao";

  return {
    score,
    bucket,
    prazo,
    pctPrazo,
    pctPrazoLabel: Math.round(pctPrazo * 100),
    diasConsumidos,
    diasRestantesPrazo,
    diasEntregaUteis,
    folgaEntrega,
    entregaVencida,
    entregaCurta: diasEntregaUteis != null && diasEntregaUteis >= 0 && diasEntregaUteis <= 15,
    semDataEntrega: !proc.dataEntrega,
    janelaComprometida: folgaEntrega != null && folgaEntrega < 0,
  };
}
