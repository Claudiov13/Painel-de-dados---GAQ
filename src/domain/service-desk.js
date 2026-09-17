function getServiceDeskRecord(ticket, serviceDeskData) {
  const raw = String(ticket || "").trim();
  if (!raw || !serviceDeskData) return null;
  const compact = raw.replace(/\D/g, "");
  return serviceDeskData[raw] || serviceDeskData[compact] || null;
}

function fixServiceDeskText(value) {
  if (value == null) return "";
  const s = String(value);
  if (!/[\u00c3\u00c2\u00e2]/.test(s)) return s;
  try {
    return decodeURIComponent(escape(s));
  } catch (_) {
    const map = {
      "\u00c3\u00a7": "\u00e7", "\u00c3\u2021": "\u00c7",
      "\u00c3\u00a3": "\u00e3", "\u00c3\u00b5": "\u00f5",
      "\u00c3\u00a1": "\u00e1", "\u00c3\u00a9": "\u00e9", "\u00c3\u00ad": "\u00ed", "\u00c3\u00b3": "\u00f3", "\u00c3\u00ba": "\u00fa",
      "\u00c3\u00a0": "\u00e0", "\u00c3\u00aa": "\u00ea", "\u00c3\u00b4": "\u00f4", "\u00c3\u0160": "\u00ca",
      "\u00c3\u0081": "\u00c1", "\u00c3\u2030": "\u00c9", "\u00c3\u008d": "\u00cd", "\u00c3\u201c": "\u00d3", "\u00c3\u0161": "\u00da",
      "\u00e2\u20ac\u201c": "\u2013", "\u00e2\u20ac\u201d": "\u2014", "\u00e2\u20ac\u00a2": "\u2022",
      "\u00c2\u00ba": "\u00ba", "\u00c2\u00aa": "\u00aa"
    };
    return Object.keys(map).reduce((acc, key) => acc.split(key).join(map[key]), s);
  }
}

function cleanServiceDeskObject(obj) {
  const out = {};
  Object.entries(obj || {}).forEach(([key, value]) => {
    out[key] = typeof value === "string" ? fixServiceDeskText(value) : value;
  });
  return out;
}

function cleanServiceDeskParty(value) {
  const text = fixServiceDeskText(value || "")
    .replace(/<[^>]*>/g, " ")
    .replace(/["']/g, " ");
  const parts = text.split(/\n+/)
    .map(part => part.replace(/\s+/g, " ").trim())
    .filter(Boolean);
  const unique = [];
  parts.forEach(part => {
    if (!unique.some(item => nrm(item) === nrm(part))) unique.push(part);
  });
  return unique[0] || "";
}

function formatServiceDeskType(tipo) {
  const t = nrm(fixServiceDeskText(tipo || ""));
  if (t.includes("nota")) return "Nota interna";
  if (t.includes("email") || t.includes("e-mail")) return "E-mail externo";
  if (t.includes("cliente")) return "Cliente · Chamado web";
  if (t.includes("atendente")) return "Atendimento";
  return fixServiceDeskText(tipo || "Interação");
}

function formatServiceDeskSubject(assunto) {
  const s = fixServiceDeskText(assunto || "").trim();
  if (nrm(s) === "atualizacao de tecnico!") return "Atualização de técnico";
  return s;
}

function summarizeServiceDesk(ticket, record) {
  if (!record) return null;
  const info = cleanServiceDeskObject(record.informacoes_do_chamado);
  const cliente = cleanServiceDeskObject(record.informacoes_do_cliente);
  const interacoes = Array.isArray(record.interacoes) ? record.interacoes.map(cleanServiceDeskObject) : [];
  const last = interacoes.length ? interacoes[interacoes.length - 1] : {};
  const estado = info.estado || "Sem status";
  const estadoN = nrm(estado);
  const lastTipoN = nrm(last.tipo || "");
  const lastFromCliente = lastTipoN.includes("cliente");
  const aguardandoArea = estadoN.includes("usuario") || estadoN.includes("cliente") || estadoN.includes("pendente inf") || (!lastFromCliente && lastTipoN.includes("atendente") && !!last.para);
  const clienteNome = cleanServiceDeskParty([cliente.nome, cliente.sobrenome].filter(Boolean).join(" "));
  const lastPara = cleanServiceDeskParty(last.para);
  const tecnicoNome = cleanServiceDeskParty(info.tecnico);
  const filaNome = cleanServiceDeskParty(info.fila);
  const comQuem = aguardandoArea
    ? (lastPara || clienteNome || "Área requisitante")
    : (tecnicoNome || filaNome || "GAQ / Service Desk");
  const acao = aguardandoArea
    ? "A área requisitante precisa responder ou ajustar a pendência solicitada."
    : lastFromCliente
      ? "GAQ deve analisar o retorno da área e registrar o próximo encaminhamento."
      : "GAQ / Service Desk deve dar continuidade ao chamado.";
  const ultimaPendencia = (last.mensagem || last.assunto || info.servico || "").toString().replace(/\s+/g, " ").trim();
  const _parseSDDate = s => { if (!s) return null; const m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m ? new Date(+m[3], +m[2]-1, +m[1]) : null; };
  const ultimaInteracaoDate = _parseSDDate(last.criado);
  const _hojeSD = new Date(); _hojeSD.setHours(0,0,0,0);
  const diasDesdeUltimaInteracao = ultimaInteracaoDate ? du(ultimaInteracaoDate, _hojeSD) : null;
  return {
    numero: String(ticket || "").trim(),
    info,
    cliente,
    interacoes,
    last,
    estado,
    comQuem,
    acao,
    ultimaPendencia,
    ultimaInteracaoDate,
    diasDesdeUltimaInteracao,
  };
}

function getServiceDeskAlertSummary(proc, serviceDeskData) {
  const summary = summarizeServiceDesk(proc?.TicketSD, getServiceDeskRecord(proc?.TicketSD, serviceDeskData));
  if (!summary) return null;
  const semRC = !String(proc?.NumRC || "").trim();
  const encerrado = nrm(summary.estado).includes("finalizado") || nrm(summary.estado).includes("fechado") || nrm(summary.last?.assunto || "").includes("fechar");
  if (!semRC || !encerrado) return summary;
  const clienteNome = cleanServiceDeskParty([summary.cliente.nome, summary.cliente.sobrenome].filter(Boolean).join(" "));
  return {
    ...summary,
    comQuem: clienteNome ? `Área requisitante (${clienteNome})` : "Área requisitante",
    acao: "Abrir a RC com os documentos validados no Service Desk.",
  };
}

function isServiceDeskConcluido(proc) {
  const statusProc = nrm(proc?.statusDet || proc?.status || "");
  return statusProc.includes("service desk concluido");
}

function enrichWithSDHistory(proc, sdData) {
  if (!proc.TicketSD || !sdData) return proc;
  const record = getServiceDeskRecord(proc.TicketSD, sdData);
  if (!record || !Array.isArray(record.interacoes) || record.interacoes.length === 0) return proc;

  const hoje = new Date(); hoje.setHours(0,0,0,0);
  const interacoes = record.interacoes.map(cleanServiceDeskObject);

  const parseSDDate = s => {
    if (!s) return null;
    const m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    return m ? new Date(+m[3], +m[2]-1, +m[1]) : null;
  };

  const aberturaHist      = parseSDDate(interacoes[0]?.criado);
  const distribuicaoHist  = parseSDDate(interacoes[1]?.criado);
  const primRespostaHist  = parseSDDate(interacoes[2]?.criado);
  const encInteracao      = [...interacoes].reverse().find(i => /fechar/i.test(i.assunto || ""));
  const encHist           = encInteracao ? parseSDDate(encInteracao.criado) : null;

  if (!aberturaHist && !distribuicaoHist && !encHist && !primRespostaHist) return proc;

  const overrides = {};
  const histFields = {};

  if (aberturaHist) {
    overrides.aberturaSD = aberturaHist;
    overrides["Data da abertura do SD"] = aberturaHist;
    histFields["Data da abertura do SD"] = true;
  }
  if (distribuicaoHist) {
    overrides["Data da distribuição do SD"] = distribuicaoHist;
    histFields["Data da distribuição do SD"] = true;
  }
  if (primRespostaHist) {
    overrides[COL_SD_PRIM_RESPOSTA] = primRespostaHist;
    histFields[COL_SD_PRIM_RESPOSTA] = true;
  }
  if (encHist) {
    overrides.encSD = encHist;
    overrides["Data do encerramento"] = encHist;
    histFields["Data do encerramento"] = true;
  }

  const newAberturaSD = overrides.aberturaSD || proc.aberturaSD;
  const newEncSD      = overrides.encSD       || proc.encSD;
  const newDiasSDAberto  = (proc.emA && newAberturaSD && !newEncSD) ? du(newAberturaSD, hoje) : 0;
  const newAtrasoSD      = newDiasSDAberto > PRAZO_SD;
  const newDiasAgingSD   = (newAberturaSD && newEncSD) ? du(newAberturaSD, newEncSD) : 0;
  const newDiasDesdeEncSD = (newEncSD && !proc.aberturaRC) ? du(newEncSD, hoje) : 0;

  return {
    ...proc,
    ...overrides,
    diasSDAberto: newDiasSDAberto,
    atrasoSD: newAtrasoSD,
    diasAgingSD: newDiasAgingSD,
    diasDesdeEncSD: newDiasDesdeEncSD,
    _sdHistoricoDates: histFields,
  };
}

function enrichSD(proc, sdData) {
  const enriched = enrichWithSDHistory(proc, sdData);
  const novoStatus = statusEfetivoSD(enriched);
  if (nrm(novoStatus).includes("service desk concluido") &&
      nrm(enriched.statusDet || enriched.status || "") !== "service desk concluido") {
    return { ...enriched, statusDet: novoStatus, _statusAjustadoLog: true };
  }
  return enriched;
}

function gerarScriptExtratorOTRS(ticketIds, modo) {
  const tplEl = document.getElementById(modo === "diurno" ? "otrs-tpl-diurno" : "otrs-tpl-rapido");
  if (!tplEl) return "/* template não encontrado */";
  const lista = JSON.stringify(ticketIds, null, 4);
  return tplEl.textContent.replace("{{LISTA_CHAMADOS}}", lista);
}
export { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS };
