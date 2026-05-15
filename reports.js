/* ══════════════════════════════════════════════════════════════════════════════
   SEÇÃO 4 — UTILITÁRIOS DE EXPORTAÇÃO
   ══════════════════════════════════════════════════════════════════════════════ */

function exportAlertaSDToExcel(rows, sdSummaries, filename) {
  filename = filename || "alerta_precompra.xlsx";
  sdSummaries = sdSummaries || {};
  var fmtD = function(d) { return d instanceof Date ? d.toLocaleDateString("pt-BR") : (d ? String(d) : ""); };
  var clean = rows.map(function(r) {
    var sd = sdSummaries[String(r.TicketSD || "").trim()] || null;
    return {
      "ProcessKey"              : r.ProcessKey,
      "Nº RC"                   : r.NumRC,
      "Ticket Pré-compra"       : r.TicketSD,
      "Comprador"               : r.Comprador,
      "Resp. NCL"               : r.respNCL,
      "Pregoeiro"               : r.Pregoeiro,
      "Avaliador"               : r.Avaliador,
      "Analista (SCONT)"        : r.AnalistaContrato,
      "Advogado Responsável"    : r.AdvogadoResp,
      "Empresa Contratada"      : r.EmpresaContratada,
      "Modalidade"              : r.Modalidade,
      "Área Requisitante"       : r["Área Requisitante"],
      "Objeto"                  : r.Objeto,
      "Status"                  : r.status,
      "Status Detalhado"        : r.statusDet,
      "Data Abertura"           : fmtD(r.dataAbertura),
      "Tipo Abertura"           : r.aberturaSD ? "Pré-compra" : (r.aberturaRC ? "RC" : ""),
      "Dias Úteis (desde RC)"   : r.diasTotais,
      "Dias Úteis Pré-compra"   : r.diasTotaisCronograma || 0,
      "Dias Úteis CPL"          : r.diasCpl,
      "Crítico NCL (>50 d.u.)"  : r.criticoNclGestao ? "Sim" : "Não",
      "Crítico CPL"             : r.criticoCpl ? "Sim" : "Não",
      "CPL"                     : r.ehCPL ? "Sim" : "Não",
      "Fase Atual"              : r.faseAtual || "",
      "Resp. Fase"              : r.respFase || "",
      "Dias Parado"             : r.diasParado || 0,
      "Status SD"               : sd ? sd.estado : "",
      "Técnico SD"              : sd ? (sd.info && sd.info.tecnico) || "" : "",
      "Com quem está"           : sd ? sd.comQuem : "",
      "Ação necessária"         : sd ? sd.acao : "",
      "Última inter. SD"        : sd && sd.ultimaInteracaoDate ? fmtD(sd.ultimaInteracaoDate) : "",
      "d.u. últ. inter. SD"     : sd && sd.diasDesdeUltimaInteracao != null ? sd.diasDesdeUltimaInteracao : "",
      "Abertura SD"             : fmtD(r.aberturaSD),
      "Distribuição SD"         : fmtD(r["Data da distribuição do SD"]),
      "Início análise SD"       : fmtD(r[COL_SD_PRIM_RESPOSTA]),
      "Encerramento SD"         : fmtD(r.encSD),
      "d.u. SD aberto"          : r.diasSDAberto || 0,
      "d.u. s/ RC (enc.)"       : r.diasDesdeEncSD || 0,
      "Tags"                    : TagsManager.getForProcess(r.ProcessKey, r.TicketSD).join(", "),
    };
  });
  var ws = window.XLSX.utils.json_to_sheet(clean);
  var wb = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb, ws, "Alerta Pré-compra");
  window.XLSX.writeFile(wb, filename);
}

function exportToExcel(data, filename = "exportacao.xlsx") {
  const clean = data.map(r => ({
    "ProcessKey": r.ProcessKey, "Nº RC": r.NumRC, "Ticket Pré-compra": r.TicketSD, "Comprador": r.Comprador,
    "Resp. NCL": r.respNCL, "Pregoeiro": r.Pregoeiro, "Avaliador": r.Avaliador, "Analista (SCONT)": r.AnalistaContrato, "Advogado Responsável": r.AdvogadoResp, "Empresa Contratada": r.EmpresaContratada, "Modalidade": r.Modalidade,
    "Área Requisitante": r["Área Requisitante"], "Objeto": r.Objeto,
    "Status": r.status, "Status Detalhado": r.statusDet,
    "Data Abertura": r.dataAbertura ? r.dataAbertura.toLocaleDateString("pt-BR") : "",
    "Tipo Abertura": r.aberturaSD ? "Pré-compra" : (r.aberturaRC ? "RC" : ""),
    "Dias Úteis (desde RC)": r.diasTotais, "Dias Úteis Pré-compra (crono)": r.diasTotaisCronograma || 0, "Dias Úteis CPL": r.diasCpl,
    "Crítico NCL (>50 d.u. RC)": r.criticoNclGestao ? "Sim" : "Não",
    "Crítico CPL": r.criticoCpl ? "Sim" : "Não",
    "CPL": r.ehCPL ? "Sim" : "Não",
    "Fase Atual": r.faseAtual || "",
    "Resp. Fase": r.respFase || "",
    "Dias Parado": r.diasParado || 0,
    "Tags": TagsManager.getForProcess(r.ProcessKey, r.TicketSD).join(", "),
  }));
  const ws = window.XLSX.utils.json_to_sheet(clean);
  const wb = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb, ws, "Dados");
  window.XLSX.writeFile(wb, filename);
}

function exportAsJS(rawRows) {
  const json = JSON.stringify(rawRows, null, 2);
  const content = "window.__PAINEL_DADOS__ = " + json + ";";
  const blob = new Blob([content], { type: "application/javascript" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = "dados.js"; a.click();
  URL.revokeObjectURL(url);
}

function gerarRelatorioHTML({ execData, fBase, rankNCL, rankGeralNCL, rankCPLResp, alertNCL, alertCPL, meta }) {
  const hoje = new Date().toLocaleDateString("pt-BR", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  const hora = new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });
  const nomeBase = meta ? meta.nomeArq : "—";
  const atualizado = meta ? meta.atualizadoEm : "—";

  const modMap = {};
  fBase.forEach(r => { const m = r.Modalidade || "Não informado"; modMap[m] = (modMap[m] || 0) + 1; });
  const modRows = Object.entries(modMap).sort((a, b) => b[1] - a[1]).slice(0, 12);

  const areaMap = {};
  fBase.forEach(r => { const a = r["Área Requisitante"] || "Não informado"; areaMap[a] = (areaMap[a] || 0) + 1; });
  const areaRows = Object.entries(areaMap).sort((a, b) => b[1] - a[1]).slice(0, 12);

  const tdS = "padding:7px 10px;border-bottom:1px solid #e0e7ef;font-size:12px;word-wrap:break-word;overflow-wrap:break-word;";
  const thS = "padding:8px 10px;background:#082445;color:#fff;font-size:11px;font-weight:700;text-align:left;";

  const tableRow = (cells, bold) => "<tr>" + cells.map((c, i) =>
    "<td style=\"" + tdS + (i > 0 ? "text-align:center;" : "") + (bold ? "font-weight:700;" : "") + "\">" + c + "</td>"
  ).join("") + "</tr>";

  const kpiCard = (ico, label, value, color) =>
    "<div style=\"background:#fff;border-radius:10px;padding:16px 14px;border-top:4px solid " + color + ";text-align:center;box-shadow:0 2px 8px rgba(0,0,0,.07);flex:1;min-width:100px;\">" +
    "<div style=\"font-size:22px\">" + ico + "</div>" +
    "<div style=\"font-size:26px;font-weight:800;color:" + color + ";margin:4px 0\">" + value + "</div>" +
    "<div style=\"font-size:10px;color:#666;font-weight:600;text-transform:uppercase\">" + label + "</div>" +
    "</div>";

  const rankTable = (title, rows, col1, col2) => {
    if (!rows || rows.length === 0) return "";
    return "<div style=\"flex:1;min-width:240px\">" +
      "<div style=\"font-weight:700;font-size:13px;color:#082445;margin-bottom:8px\">" + title + "</div>" +
      "<table style=\"width:100%;border-collapse:collapse;border:1px solid #dbe3ee;border-radius:8px;overflow:hidden\">" +
      "<thead><tr><th style=\"" + thS + "\">#</th><th style=\"" + thS + "\">" + col1 + "</th><th style=\"" + thS + "text-align:center\">" + col2 + "</th></tr></thead>" +
      "<tbody>" + rows.slice(0, 8).map((x, i) => tableRow([i + 1, x.name, x.badge])).join("") + "</tbody>" +
      "</table></div>";
  };

  const critList = (list, tipo) => {
    if (!list || list.length === 0) return "<p style=\"color:#888;font-size:12px\">Nenhum processo crítico " + tipo + ".</p>";
    return "<table style=\"width:100%;border-collapse:collapse;border:1px solid #dbe3ee\">" +
      "<thead><tr>" +
      "<th style=\"" + thS + "\">Processo</th><th style=\"" + thS + "\">Responsável</th>" +
      "<th style=\"" + thS + "\">Área</th><th style=\"" + thS + "text-align:center\">d.u. RC</th><th style=\"" + thS + "text-align:center\">d.u. pré-compra</th><th style=\"" + thS + "\">Objeto</th>" +
      "</tr></thead><tbody>" +
      list.slice(0, 15).map(r => {
        const duSD = (r.diasAgingSD > 0) ? r.diasAgingSD : (r.diasSD > 0 ? r.diasSD : "—");
        return tableRow([
          r.NumRC || r.TicketSD || "—",
          r.respNCL || r.Pregoeiro || "—",
          r["Área Requisitante"] || "—",
          r.diasTotaisGestao || r.diasCpl || "—",
          duSD,
          (r.Objeto || "—").slice(0, 500) + ((r.Objeto || "").length > 500 ? "…" : "")
        ]);
      }).join("") +
      "</tbody></table>";
  };

  const html = "<!DOCTYPE html><html lang=\"pt-BR\"><head><meta charset=\"UTF-8\">" +
    "<title>Relatório GAQ — " + hoje + "</title>" +
    "<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:'Segoe UI',sans-serif;background:#eef3f8;color:#1f2937;padding:24px}" +
    "h2{font-size:14px;color:#082445;margin:20px 0 10px;border-left:4px solid #2e86c1;padding-left:10px}" +
    "@media print{body{background:#fff;padding:10px}}</style></head><body>" +

    "<div style=\"background:linear-gradient(135deg,#082445,#163a63);color:#fff;border-radius:12px;padding:22px 28px;margin-bottom:20px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px\">" +
    "<div><div style=\"font-size:20px;font-weight:800\">Sesc — Painel Inteligência de Compras GAQ</div>" +
    "<div style=\"font-size:11px;opacity:.7;margin-top:4px\">Relatório Mensal · " + hoje + " às " + hora + "</div>" +
    "<div style=\"font-size:10px;opacity:.55;margin-top:2px\">Base: " + nomeBase + " · Atualizado: " + atualizado + " · " + execData.tot + " processos</div></div>" +
    "<div style=\"font-size:11px;opacity:.7;text-align:right\">GAQ — Gerência de Aquisições - DN</div></div>" +

    "<h2>Resumo Geral</h2>" +
    "<div style=\"display:flex;gap:10px;flex-wrap:wrap;margin-bottom:20px\">" +
    kpiCard("📊", "Total", execData.tot, "#2e86c1") +
    kpiCard("⏳", "Em Andamento", execData.emA, "#3498db") +
    kpiCard("✅", "Taxa Conclusão", execData.taxaConc + "%", "#27ae60") +
    kpiCard("⚠️", "Taxa Problemáticos", execData.taxaProb + "%", execData.taxaProb > 20 ? "#c0392b" : "#e67e22") +
    kpiCard("🚨", "Críticos NCL", execData.alertNCLCount, execData.alertNCLCount > 0 ? "#c0392b" : "#27ae60") +
    kpiCard("🔴", "Críticos CPL", execData.alertCPLCount, execData.alertCPLCount > 0 ? "#8e44ad" : "#27ae60") +
    "</div>" +

    "<div style=\"display:flex;gap:16px;flex-wrap:wrap;margin-bottom:20px\">" +
    "<div style=\"flex:1;min-width:220px\"><h2 style=\"margin-top:0\">Modalidades</h2>" +
    "<table style=\"width:100%;border-collapse:collapse;border:1px solid #dbe3ee\">" +
    "<thead><tr><th style=\"" + thS + "\">Modalidade</th><th style=\"" + thS + "text-align:center\">Total</th></tr></thead>" +
    "<tbody>" + modRows.map(function(mr) { return tableRow([mr[0], mr[1]]); }).join("") + "</tbody></table></div>" +
    "<div style=\"flex:1;min-width:220px\"><h2 style=\"margin-top:0\">Áreas Requisitantes</h2>" +
    "<table style=\"width:100%;border-collapse:collapse;border:1px solid #dbe3ee\">" +
    "<thead><tr><th style=\"" + thS + "\">Área</th><th style=\"" + thS + "text-align:center\">Total</th></tr></thead>" +
    "<tbody>" + areaRows.map(function(ar) { return tableRow([ar[0], ar[1]]); }).join("") + "</tbody></table></div></div>" +

    "<h2>Desempenho NCL — Compradores (em andamento)</h2>" +
    "<div style=\"display:flex;gap:16px;flex-wrap:wrap;margin-bottom:20px\">" +
    rankTable("Críticos por Comprador", (rankNCL || []).map(function(x) { return { name: x.name, badge: x.criticos + " críticos / " + x.total + " total" }; }), "Comprador", "Críticos / Total") +
    rankTable("Média dias úteis", (rankGeralNCL || []).map(function(x) { return { name: x.name, badge: x.media + " d.u." }; }), "Comprador", "Média d.u.") +
    "</div>" +

    "<h2>Desempenho CPL — Responsáveis (em andamento)</h2>" +
    "<div style=\"display:flex;gap:16px;flex-wrap:wrap;margin-bottom:20px\">" +
    rankTable("Processos ativos", (rankCPLResp || []).map(function(x) { return { name: x.name, badge: x.total + " ativos" }; }), "Responsável", "Ativos") +
    rankTable("Média dias úteis", (rankCPLResp || []).slice().sort(function(a,b){ return b.media-a.media; }).map(function(x) { return { name: x.name, badge: x.media + " d.u." }; }), "Responsável", "Média d.u.") +
    "</div>" +

    "<h2>Processos Críticos NCL (&gt;50 d.u. em andamento)</h2>" +
    "<div style=\"margin-bottom:20px\">" + critList(alertNCL, "NCL") + "</div>" +

    "<h2>Processos Críticos CPL (&gt;50 d.u. em andamento)</h2>" +
    "<div style=\"margin-bottom:20px\">" + critList(alertCPL, "CPL") + "</div>" +

    "<div style=\"text-align:center;font-size:10px;color:#aaa;margin-top:24px;border-top:1px solid #e0e7ef;padding-top:12px\">" +
    "Gerado automaticamente pelo Painel GAQ · " + new Date().toLocaleString("pt-BR") + "</div>" +

    "<script>setTimeout(function(){ window.print(); }, 400);<\/script>" +
    "</body></html>";

  return html;
}

function gerarRelatorioOperacional({ processos, area, meta }) {
  const hoje = new Date().toLocaleDateString("pt-BR", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  const hora = new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });
  const nomeBase = meta ? meta.nomeArq : "—";
  const titulo = area === "TODAS" ? "Todas as Áreas" : area;

  const emA = processos.filter(function(r) { return r.emA; });
  const total = processos.length;
  const mediaAging = emA.length ? Math.round(emA.reduce(function(s, r) { return s + r.diasTotais; }, 0) / emA.length) : 0;
  const crit30 = emA.filter(function(r) { return r.diasTotais > 30; }).length;
  const crit60 = emA.filter(function(r) { return r.diasTotais > 60; }).length;
  const semPedido = emA.filter(function(r) { return r.semEnvio; }).length;

  // Por modalidade
  var modMap = {};
  emA.forEach(function(r) { var m = r.Modalidade || "N/I"; if (!modMap[m]) modMap[m] = 0; modMap[m]++; });
  var modRows = Object.entries(modMap).sort(function(a,b){ return b[1]-a[1]; });

  // Por comprador (NCL: respNCL, CPL: Pregoeiro)
  var compMap = {};
  emA.forEach(function(r) { var c = r.ehCPL ? (r.Pregoeiro || r.cplResp || "N/I") : (r.respNCL || "N/I"); if (!compMap[c]) compMap[c] = { total: 0, crit: 0 }; compMap[c].total++; if (r.diasTotais > 50) compMap[c].crit++; });
  var compRows = Object.entries(compMap).map(function(e) { return { name: e[0], total: e[1].total, crit: e[1].crit }; }).sort(function(a,b){ return b.crit-a.crit || b.total-a.total; });

  // Processos críticos (>50 d.u.) ordenados por aging
  var criticos = emA.filter(function(r) { return r.diasTotais > 50; }).sort(function(a,b){ return b.diasTotais-a.diasTotais; });

  // Todos os processos em andamento ordenados por aging
  var listaEmA = emA.slice().sort(function(a,b){ return b.diasTotais-a.diasTotais; });

  var thS = "padding:7px 10px;background:#2c3e50;color:#fff;font-size:11px;font-weight:700;text-align:left;";
  var tdS = "padding:6px 10px;border-bottom:1px solid #e0e7ef;font-size:11px;word-wrap:break-word;overflow-wrap:break-word;";

  function row(cells) {
    return "<tr>" + cells.map(function(c, i) {
      return "<td style=\"" + tdS + (i > 0 ? "text-align:center;" : "") + "\">" + (c === null || c === undefined ? "—" : c) + "</td>";
    }).join("") + "</tr>";
  }

  function kpi(label, value, color) {
    return "<div style=\"background:#fff;border-radius:8px;padding:14px 12px;border-top:4px solid " + color + ";text-align:center;flex:1;min-width:90px;box-shadow:0 2px 6px rgba(0,0,0,.07)\">" +
      "<div style=\"font-size:24px;font-weight:800;color:" + color + "\">" + value + "</div>" +
      "<div style=\"font-size:10px;color:#666;font-weight:600;text-transform:uppercase;margin-top:3px\">" + label + "</div>" +
      "</div>";
  }

  function table(headers, rowsHtml) {
    return "<table style=\"width:100%;border-collapse:collapse;border:1px solid #dbe3ee;margin-bottom:16px\">" +
      "<thead><tr>" + headers.map(function(h) { return "<th style=\"" + thS + "\">" + h + "</th>"; }).join("") + "</tr></thead>" +
      "<tbody>" + rowsHtml + "</tbody></table>";
  }

  var modTable = table(["Modalidade", "Qtd"],
    modRows.map(function(m) { return row([m[0], m[1]]); }).join(""));

  var compTable = table(["Responsável", "Em Andamento", "Críticos (>50 d.u.)"],
    compRows.map(function(c) { return row([c.name, c.total, c.crit > 0 ? "<span style=\"color:#c0392b;font-weight:700\">" + c.crit + "</span>" : c.crit]); }).join(""));

  var critTable = criticos.length === 0
    ? "<p style=\"color:#888;font-size:12px;margin-bottom:16px\">Nenhum processo crítico para este filtro.</p>"
    : table(["Processo", "Responsável", "Modalidade", "Status", "d.u. RC", "d.u. pré-compra", "Objeto"],
        criticos.slice(0, 30).map(function(r) {
          var duSD=r.diasAgingSD>0?r.diasAgingSD:(r.diasSD>0?r.diasSD:"—");
          return row([
            r.NumRC || r.TicketSD || "—",
            r.respNCL || r.Pregoeiro || "—",
            r.Modalidade || "—",
            r.statusDet || r.status || "—",
            "<span style=\"color:" + (r.diasTotais > 80 ? "#c0392b" : "#e67e22") + ";font-weight:700\">" + r.diasTotais + "</span>",
            duSD,
            (r.Objeto || "—").slice(0, 500) + ((r.Objeto || "").length > 500 ? "…" : "")
          ]);
        }).join(""));

  var listaTable = table(["Processo", "Responsável", "Área", "Modalidade", "Status", "d.u. RC", "d.u. pré-compra", "Objeto"],
    listaEmA.slice(0, 60).map(function(r) {
      var duSD=r.diasAgingSD>0?r.diasAgingSD:(r.diasSD>0?r.diasSD:"—");
      return row([
        r.NumRC || r.TicketSD || "—",
        r.respNCL || r.Pregoeiro || "—",
        r["Área Requisitante"] || "—",
        r.Modalidade || "—",
        r.statusDet || r.status || "—",
        r.diasTotais,
        duSD,
        (r.Objeto || "—").slice(0, 500) + ((r.Objeto || "").length > 500 ? "…" : "")
      ]);
    }).join(""));

  return "<!DOCTYPE html><html lang=\"pt-BR\"><head><meta charset=\"UTF-8\">" +
    "<title>Relatório Operacional — " + titulo + "</title>" +
    "<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:'Segoe UI',sans-serif;background:#eef3f8;color:#1f2937;padding:20px}" +
    "h2{font-size:13px;color:#2c3e50;margin:18px 0 8px;border-left:4px solid #34495e;padding-left:9px}" +
    "@media print{body{background:#fff;padding:8px}}</style></head><body>" +

    "<div style=\"background:linear-gradient(135deg,#2c3e50,#34495e);color:#fff;border-radius:10px;padding:18px 24px;margin-bottom:18px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px\">" +
    "<div><div style=\"font-size:18px;font-weight:800\">Sesc — Relatório Operacional</div>" +
    "<div style=\"font-size:12px;font-weight:700;margin-top:3px;opacity:.9\">Área: " + titulo + "</div>" +
    "<div style=\"font-size:10px;opacity:.6;margin-top:2px\">" + hoje + " às " + hora + " · Base: " + nomeBase + "</div></div>" +
    "<div style=\"font-size:10px;opacity:.6;text-align:right\">GAQ — Gerência de Aquisições - DN</div></div>" +

    "<h2>Indicadores</h2>" +
    "<div style=\"display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px\">" +
    kpi("Em Andamento", emA.length, "#3498db") +
    kpi("Média Aging", mediaAging + " d.u.", mediaAging > 50 ? "#c0392b" : "#2e86c1") +
    kpi(">30 d.u.", crit30, crit30 > 10 ? "#e67e22" : "#27ae60") +
    kpi(">60 d.u.", crit60, crit60 > 5 ? "#c0392b" : "#e67e22") +
    kpi("Sem Pedido", semPedido, semPedido > 5 ? "#e67e22" : "#27ae60") +
    kpi("Total Processos", total, "#2c3e50") +
    "</div>" +

    "<div style=\"display:flex;gap:16px;flex-wrap:wrap;margin-bottom:4px\">" +
    "<div style=\"flex:1;min-width:200px\"><h2 style=\"margin-top:0\">Por Modalidade</h2>" + modTable + "</div>" +
    "<div style=\"flex:2;min-width:280px\"><h2 style=\"margin-top:0\">Por Responsável</h2>" + compTable + "</div>" +
    "</div>" +

    "<h2>Processos Críticos (&gt;50 d.u.)</h2>" + critTable +

    "<h2>Todos os Processos em Andamento</h2>" + listaTable +

    "<div style=\"text-align:center;font-size:10px;color:#aaa;margin-top:20px;border-top:1px solid #e0e7ef;padding-top:10px\">" +
    "Gerado pelo Painel GAQ · " + new Date().toLocaleString("pt-BR") + "</div>" +

    "<script>setTimeout(function(){ window.print(); }, 400);<\/script>" +
    "</body></html>";
}

function _pdfCommon(accentColor) {
  var PRINT_COLOR = "-webkit-print-color-adjust:exact;print-color-adjust:exact;color-adjust:exact;";
  var thS = "padding:7px 10px;background:"+accentColor+";color:#fff;font-size:11px;font-weight:700;text-align:left;"+PRINT_COLOR;
  var tdS = "padding:6px 10px;border-bottom:1px solid #e0e7ef;font-size:11px;word-wrap:break-word;overflow-wrap:break-word;";
  function row(cells) { return "<tr>"+cells.map(function(c,i){return "<td style=\""+tdS+(i>0?"text-align:center;":"")+"\">"+(c===null||c===undefined?"—":c)+"</td>";}).join("")+"</tr>"; }
  function kpi(label,value,color){return "<div style=\"background:#fff;border-radius:8px;padding:14px 12px;border-top:4px solid "+color+";text-align:center;flex:1;min-width:90px;box-shadow:0 2px 6px rgba(0,0,0,.07);"+PRINT_COLOR+"\"><div style=\"font-size:24px;font-weight:800;color:"+color+"\">"+value+"</div><div style=\"font-size:10px;color:#666;font-weight:600;text-transform:uppercase;margin-top:3px\">"+label+"</div></div>";}
  function tbl(headers,rowsHtml){return "<table style=\"width:100%;border-collapse:collapse;border:1px solid #dbe3ee;margin-bottom:16px\"><thead><tr>"+headers.map(function(h){return "<th style=\""+thS+"\">"+h+"</th>";}).join("")+"</tr></thead><tbody>"+rowsHtml+"</tbody></table>";}
  function h2(txt,color){return "<h2 style=\"font-size:13px;color:"+(color||accentColor)+";margin:18px 0 8px;border-left:4px solid "+(color||accentColor)+";padding-left:9px;"+PRINT_COLOR+"\">"+txt+"</h2>";}
  function sazon(fBase,color){
    var MESES=["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
    var mc=Array(12).fill(0);
    var anoCount={};
    (fBase||[]).forEach(function(r){
      if(r.mesAbertura!==null&&r.mesAbertura!==undefined)mc[r.mesAbertura]++;
      var a=r.anoSD||r.anoRC; if(a){anoCount[a]=(anoCount[a]||0)+1;}
    });
    var anoLabel=Object.keys(anoCount).sort(function(a,b){return anoCount[b]-anoCount[a];})[0]||new Date().getFullYear();
    var maxM=Math.max.apply(null,mc),avgM=mc.reduce(function(a,b){return a+b;},0)/12;
    if(maxM<3)return "";
    var bars=mc.map(function(c,i){var h=maxM>0?Math.max(8,Math.round((c/maxM)*80)):8;var isPeak=c>avgM*1.3,isLow=c<avgM*0.7;var col=isPeak?"#c0392b":isLow?"#85c1e9":(color||"#2e86c1");return "<td style=\"text-align:center;vertical-align:bottom;padding:0 3px;width:7%\"><div style=\"font-size:10px;font-weight:700;color:"+col+";margin-bottom:2px\">"+c+"</div><div style=\"height:"+h+"px;background:"+col+";border-radius:3px 3px 0 0;"+PRINT_COLOR+"\"></div><div style=\"font-size:9px;color:#666;margin-top:3px\">"+MESES[i]+"</div></td>";}).join("");
    return h2("📅 Sazonalidade — Padrão de Aberturas por Mês - ("+anoLabel+")")+"<div style=\"background:#fff;border-radius:8px;padding:16px;border:1px solid #dbe3ee;margin-bottom:16px;"+PRINT_COLOR+"\"><table style=\"width:100%;border-collapse:collapse\"><tbody><tr style=\"vertical-align:bottom\">"+bars+"</tr></tbody></table><div style=\"font-size:10px;color:#999;margin-top:10px;border-top:1px solid #eee;padding-top:6px\">■ <span style=\"color:#c0392b;font-weight:700\">Pico (&gt;130% média)</span> &nbsp;■ <span style=\"color:#85c1e9;font-weight:700\">Baixa (&lt;70% média)</span> &nbsp;Média: "+Math.round(avgM)+" proc/mês</div></div>";
  }
  function funil(emA,ac){
    var FASES=[
      {label:"Sem pré-compra/RC",fn:function(r){return !r.aberturaSD&&!r.aberturaRC;}},
      {label:"Pré-compra em distribuição",fn:function(r){return !!(r.aberturaSD&&!r["Data da distribuição do SD"]);}},
      {label:"Pré-compra distribuída",fn:function(r){var v=r["Data da distribuição do SD"];return !!(v&&v!=="")&&!r["Data do encerramento"];}},
      {label:"Pré-compra encerrada",fn:function(r){var v=r["Data do encerramento"];return !!(v&&v!=="")&&!(r["DATA DO RECEBIMENTO DA RC"]&&r["DATA DO RECEBIMENTO DA RC"]!=="");}},
      {label:"RC Recebida",fn:function(r){var rc=r["DATA DO RECEBIMENTO DA RC"];return !!(rc&&rc!=="")&&!(r["Data inicial do envio de propostas"]&&r["Data inicial do envio de propostas"]!=="")&&!(r["Data do envio para aprovação"]&&r["Data do envio para aprovação"]!=="");}},
      {label:"Em Cotação",fn:function(r){var ini=r["Data inicial do envio de propostas"];return !!(ini&&ini!=="")&&!(r["Data final do envio de propostas"]&&r["Data final do envio de propostas"]!=="");}},
      {label:"Cotação Recebida",fn:function(r){var fim=r["Data final do envio de propostas"];return !!(fim&&fim!=="")&&!(r["Data do envio para aprovação"]&&r["Data do envio para aprovação"]!=="");}},
      {label:"Em Aprovação",fn:function(r){var env=r["Data do envio para aprovação"];return !!(env&&env!=="")&&!(r["Data da última aprovação"]&&r["Data da última aprovação"]!=="");}},
      {label:"Aguard. Pedido/Suite",fn:function(r){var ult=r["Data da última aprovação"];return !!(ult&&ult!=="")&&r.semEnvio;}},
    ];
    var fd=FASES.map(function(f){return{label:f.label,count:emA.filter(f.fn).length};}).filter(function(f){return f.count>0;});
    if(!fd.length)return "";
    var maxF=Math.max.apply(null,fd.map(function(f){return f.count;}));
    var rows=fd.map(function(f){var pct=maxF>0?Math.round((f.count/maxF)*100):0;return "<tr><td style=\"padding:5px 8px;font-size:11px;color:#1f2937;text-align:right;width:170px;white-space:nowrap\">"+f.label+"</td><td style=\"padding:5px 8px\"><div style=\"background:#e0e7ef;border-radius:3px;height:18px;overflow:hidden;"+PRINT_COLOR+"\"><div style=\"width:"+pct+"%;height:100%;background:"+ac+";border-radius:3px;"+PRINT_COLOR+"\"></div></div></td><td style=\"padding:5px 8px;font-size:12px;font-weight:700;color:"+ac+";width:40px;text-align:center\">"+f.count+"</td></tr>";}).join("");
    return h2("🔄 Funil de Fases — Processos em Andamento")+"<div style=\"background:#fff;border-radius:8px;padding:12px;border:1px solid #dbe3ee;margin-bottom:16px\"><table style=\"width:100%;border-collapse:collapse\"><tbody>"+rows+"</tbody></table></div>";
  }
  function statusDetSection(emA,total){
    var sm={};emA.forEach(function(r){var s=(r.statusDet||r.status||"Sem status").trim();sm[s]=(sm[s]||0)+1;});
    var sa=Object.entries(sm).sort(function(a,b){return b[1]-a[1];});
    var rowsHtml=sa.slice(0,25).map(function(e){return row([e[0],e[1],total>0?Math.round((e[1]/total)*100)+"%":"0%"]);}).join("");
    return rowsHtml?tbl(["Status Detalhado","Qtd","% Em And."],rowsHtml):"<p style=\"color:#888;font-size:12px\">Sem dados.</p>";
  }
  return {row:row,kpi:kpi,tbl:tbl,h2:h2,sazon:sazon,funil:funil,statusDetSection:statusDetSection,PRINT_COLOR:PRINT_COLOR};
}

function gerarRelatorioGAQSexta({ processos, area, meta, fBase }) {
  var hoje = new Date().toLocaleDateString("pt-BR", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  var hora = new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });
  var nomeBase = meta ? meta.nomeArq : "—";
  var titulo = (!area || area === "TODAS" || area === "Todas as Áreas") ? "Todas as Áreas" : area;
  var filtroInfo = (!area || area === "TODAS" || area === "Todas as Áreas") ? "" : " · Área filtrada: <strong style=\"color:#f39c12\">" + area + "</strong>";
  var accent = "#1a3a5c";
  var C = _pdfCommon(accent);

  var allProcs = processos; // todos os processos (fBase filtrado)
  var emA = allProcs.filter(function(r){ return r.emA; });
  var conc = allProcs.filter(function(r){ return r.isConcluded; });
  var canc = allProcs.filter(function(r){ return r.isCanceled; });
  var frac = allProcs.filter(function(r){ return r.isFailed; });
  var totalEmA = emA.length;
  var mediaAging = totalEmA ? Math.round(emA.reduce(function(s,r){ return s + r.diasTotais; }, 0) / totalEmA) : 0;
  var crit30 = emA.filter(function(r){ return r.diasTotais > 30; }).length;
  var crit60 = emA.filter(function(r){ return r.diasTotais > 60; }).length;
  var crit90 = emA.filter(function(r){ return r.diasTotais > 90; }).length;
  var kpiRow = "<div style=\"display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px\">" +
    C.kpi("Em Andamento", totalEmA, "#1a3a5c") +
    C.kpi("Concluídos", conc.length, "#27ae60") +
    C.kpi("Cancelados", canc.length, "#c0392b") +
    C.kpi("Fracassados", frac.length, "#e67e22") +
    C.kpi("Média Aging", mediaAging + " d.u.", mediaAging > 60 ? "#c0392b" : "#2e86c1") +
    C.kpi("> 30 d.u.", crit30, crit30 > 20 ? "#e67e22" : "#27ae60") +
    C.kpi("> 60 d.u.", crit60, crit60 > 10 ? "#c0392b" : "#e67e22") +
    C.kpi("> 90 d.u.", crit90, crit90 > 5 ? "#c0392b" : "#27ae60") +
    "</div>";

  var statusHTML = C.h2("📋 Status Detalhado") + C.statusDetSection(emA, totalEmA);

  // Por modalidade
  var modMap = {};
  emA.forEach(function(r){ var m = r.Modalidade || "N/I"; if (!modMap[m]) modMap[m] = { total: 0, crit: 0 }; modMap[m].total++; if (r.diasTotais > 50) modMap[m].crit++; });
  var modRows = Object.entries(modMap).sort(function(a,b){ return b[1].total - a[1].total; }).map(function(e){
    return C.row([e[0], e[1].total, e[1].crit > 0 ? "<span style=\"color:#c0392b;font-weight:700\">" + e[1].crit + "</span>" : "0"]);
  }).join("");
  var modTable = modRows ? C.tbl(["Modalidade", "Em Andamento", "Críticos (>50 d.u.)"], modRows) : "";

  // Por área
  var areaMap = {};
  emA.forEach(function(r){ var a = r["Área Requisitante"] || "N/I"; if (!areaMap[a]) areaMap[a] = { total: 0, crit: 0 }; areaMap[a].total++; if (r.diasTotais > 50) areaMap[a].crit++; });
  var areaRows = Object.entries(areaMap).sort(function(a,b){ return b[1].total - a[1].total; }).map(function(e){
    return C.row([e[0], e[1].total, e[1].crit > 0 ? "<span style=\"color:#c0392b;font-weight:700\">" + e[1].crit + "</span>" : "0"]);
  }).join("");
  var areaTable = areaRows ? C.tbl(["Área Requisitante", "Em Andamento", "Críticos (>50 d.u.)"], areaRows) : "";

  // Lista completa de processos em andamento ordenada por diasTotais desc
  var listaOrdenada = emA.slice().sort(function(a, b){ return b.diasTotais - a.diasTotais; });
  var listaRows = listaOrdenada.map(function(r, idx) {
    var corDu = r.diasTotais > 100 ? "#922b21" : r.diasTotais > 60 ? "#c0392b" : r.diasTotais > 30 ? "#e67e22" : "#27ae60";
    var duSD = r.diasAgingSD > 0 ? r.diasAgingSD : (r.diasSD > 0 ? r.diasSD : "—");
    var corSD = (r.diasAgingSD||r.diasSD) > 50 ? "#c0392b" : (r.diasAgingSD||r.diasSD) > 20 ? "#e67e22" : "#27ae60";
    return C.row([
      idx + 1,
      r.NumRC || "—",
      r.TicketSD || "—",
      r.NumProcesso || "—",
      r.Comprador || "—",
      r.Avaliador || "—",
      r.Modalidade || "—",
      r["Área Requisitante"] || "—",
      (r.Objeto || "—").slice(0, 500) + ((r.Objeto || "").length > 500 ? "…" : ""),
      r.status || "—",
      r.statusDet || "—",
      "<span style=\"color:" + corDu + ";font-weight:800\">" + r.diasTotais + "</span>",
      duSD !== "—" ? "<span style=\"color:" + corSD + ";font-weight:700\">" + duSD + "</span>" : "—"
    ]);
  }).join("");
  var listaTable = listaRows
    ? C.tbl(["#","RC","Pré-compra","Nº Processo","Comprador","Avaliador","Modalidade","Área","Objeto","Status","Status Det.","d.u. RC","d.u. pré-compra"], listaRows)
    : "<p style=\"color:#888;font-size:12px\">Nenhum processo em andamento.</p>";

  var css = "*{box-sizing:border-box;margin:0;padding:0;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important}" +
    "body{font-family:'Segoe UI',sans-serif;background:#eef3f8;color:#1f2937;padding:20px;font-size:12px}" +
    "table{font-size:11px;border-collapse:collapse}" +
    "@media print{body{background:#fff!important;padding:6px}@page{margin:1cm;size:A4 landscape}}";

  return "<!DOCTYPE html><html lang=\"pt-BR\"><head><meta charset=\"UTF-8\"><title>PDF GAQ Sexta — " + titulo + "</title><style>" + css + "</style></head><body>" +
    "<div style=\"background:linear-gradient(135deg,#1a3a5c,#2e6da4);color:#fff;border-radius:10px;padding:18px 24px;margin-bottom:20px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px\">" +
    "<div><div style=\"font-size:20px;font-weight:800\">📊 Sesc GAQ — Relatório Operacional Sexta</div>" +
    "<div style=\"font-size:12px;font-weight:700;margin-top:4px;opacity:.9\">Área: " + titulo + filtroInfo + "</div>" +
    "<div style=\"font-size:10px;opacity:.6;margin-top:3px\">" + hoje + " às " + hora + " · Base: " + nomeBase + "</div></div>" +
    "<div style=\"font-size:10px;opacity:.6;text-align:right\">GAQ — Gerência de Aquisições - DN</div></div>" +
    C.h2("📈 Big Numbers — Situação Geral") + kpiRow +
    C.sazon(fBase || allProcs, accent) +
    C.funil(emA, accent) +
    statusHTML +
    "<div style=\"display:flex;gap:16px;flex-wrap:wrap;margin-bottom:4px\">" +
    "<div style=\"flex:1;min-width:220px\">" + C.h2("🏷 Por Modalidade") + modTable + "</div>" +
    "<div style=\"flex:1;min-width:220px\">" + C.h2("📂 Por Área Requisitante") + areaTable + "</div></div>" +
    C.h2("📋 Lista Completa — Processos em Andamento (" + totalEmA + ") — Ordenados por Maior Aging") + listaTable +
    "<div style=\"text-align:center;font-size:10px;color:#aaa;margin-top:20px;border-top:1px solid #e0e7ef;padding-top:10px\">Gerado pelo Painel GAQ · " + new Date().toLocaleString("pt-BR") + "</div>" +
    "<script>setTimeout(function(){ window.print(); }, 400);<\/script></body></html>";
}

function gerarRelatorioGAQSextaPorAreas({ base, meta }) {
  var hoje = new Date().toLocaleDateString("pt-BR", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  var hora = new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });
  var nomeBase = meta ? meta.nomeArq : "—";
  var accent = "#1a3a5c";
  var C = _pdfCommon(accent);

  var emATodos = (base||[]).filter(function(r){ return r.emA; });
  var areas = [];
  var areaSeen = {};
  emATodos.forEach(function(r){ var a=r["Área Requisitante"]||"N/I"; if(!areaSeen[a]){areaSeen[a]=true;areas.push(a);} });
  areas.sort();

  var css = "*{box-sizing:border-box;margin:0;padding:0;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important}" +
    "body{font-family:'Segoe UI',sans-serif;background:#fff;color:#1f2937;padding:16px;font-size:11px}" +
    "table{font-size:10px;border-collapse:collapse}" +
    ".area-section{page-break-after:always;break-after:page;padding-bottom:20px;margin-bottom:20px}" +
    ".area-section:last-child{page-break-after:auto;break-after:auto}" +
    "@media print{body{padding:6px}@page{margin:1cm;size:A4 landscape}}";

  function gerarSecaoArea(area) {
    var emA = emATodos.filter(function(r){ return (r["Área Requisitante"]||"N/I") === area; });
    var total = emA.length;
    if(total === 0) return "";
    var mediaAging = Math.round(emA.reduce(function(s,r){return s+r.diasTotais;},0)/total);
    var crit30 = emA.filter(function(r){return r.diasTotais>30;}).length;
    var crit60 = emA.filter(function(r){return r.diasTotais>60;}).length;
    var crit90 = emA.filter(function(r){return r.diasTotais>90;}).length;

    var kpis = "<div style=\"display:flex;gap:6px;flex-wrap:wrap;margin-bottom:14px\">" +
      C.kpi("Em Andamento",total,accent) +
      C.kpi("Média Aging",mediaAging+" d.u.",mediaAging>60?"#c0392b":"#2e86c1") +
      C.kpi(">30 d.u.",crit30,crit30>10?"#e67e22":"#27ae60") +
      C.kpi(">60 d.u.",crit60,crit60>5?"#c0392b":"#e67e22") +
      C.kpi(">90 d.u.",crit90,crit90>2?"#c0392b":"#27ae60") +
      "</div>";

    var sorted = emA.slice().sort(function(a,b){return b.diasTotais-a.diasTotais;});
    var listaRows = sorted.map(function(r,idx){
      var corDu=r.diasTotais>100?"#922b21":r.diasTotais>60?"#c0392b":r.diasTotais>30?"#e67e22":"#27ae60";
      var duSD=r.diasAgingSD>0?r.diasAgingSD:(r.diasSD>0?r.diasSD:"—");
      var corSD=(r.diasAgingSD||r.diasSD)>50?"#c0392b":(r.diasAgingSD||r.diasSD)>20?"#e67e22":"#27ae60";
      return C.row([
        idx+1, r.NumRC||"—", r.TicketSD||"—", r.NumProcesso||"—",
        r.Comprador||"—", r.Avaliador||"—", r.Modalidade||"—",
        (r.Objeto||"—").slice(0,500)+((r.Objeto||"").length>500?"…":""),
        r.status||"—", r.statusDet||"—",
        "<span style=\"color:"+corDu+";font-weight:800\">"+r.diasTotais+"</span>",
        duSD!=="—"?"<span style=\"color:"+corSD+";font-weight:700\">"+duSD+"</span>":"—"
      ]);
    }).join("");
    var listaTable = listaRows ? C.tbl(["#","RC","Pré-compra","Nº Processo","Comprador","Avaliador","Modalidade","Objeto","Status","Status Det.","d.u. RC","d.u. pré-compra"],listaRows)
      : "<p style=\"color:#888\">Nenhum processo.</p>";

    return "<div class=\"area-section\">" +
      "<div style=\"background:linear-gradient(135deg,"+accent+",#2e6da4);color:#fff;border-radius:8px;padding:14px 20px;margin-bottom:14px;display:flex;justify-content:space-between;align-items:center\">" +
      "<div><div style=\"font-size:16px;font-weight:800\">📂 "+area+"</div>" +
      "<div style=\"font-size:10px;opacity:.7;margin-top:2px\">"+hoje+" às "+hora+" · "+total+" processo(s) em andamento · Base: "+nomeBase+"</div></div>" +
      "<div style=\"font-size:10px;opacity:.6\">GAQ</div></div>" +
      kpis +
      C.funil(emA,accent) +
      C.h2("📋 Status Detalhado")+C.statusDetSection(emA,total) +
      C.h2("📋 Lista de Processos em Andamento — Ordenados por Maior Aging") + listaTable +
      "</div>";
  }

  // Capa / sumário
  var totalGeral = emATodos.length;
  var mediaGeral = totalGeral ? Math.round(emATodos.reduce(function(s,r){return s+r.diasTotais;},0)/totalGeral) : 0;
  var resumoRows = areas.map(function(a){
    var emA=emATodos.filter(function(r){return (r["Área Requisitante"]||"N/I")===a;});
    var c60=emA.filter(function(r){return r.diasTotais>60;}).length;
    return C.row([a, emA.length, c60>0?"<span style=\"color:#c0392b;font-weight:700\">"+c60+"</span>":"0",
      Math.round(emA.reduce(function(s,r){return s+r.diasTotais;},0)/(emA.length||1))+" d.u."]);
  }).join("");

  var capa = "<div class=\"area-section\">" +
    "<div style=\"background:linear-gradient(135deg,"+accent+",#2e6da4);color:#fff;border-radius:10px;padding:22px 28px;margin-bottom:20px\">" +
    "<div style=\"font-size:22px;font-weight:800\">📊 Sesc GAQ — Relatório Operacional por Área</div>" +
    "<div style=\"font-size:13px;font-weight:700;margin-top:4px;opacity:.9\">Todas as Áreas — Processos em Andamento</div>" +
    "<div style=\"font-size:11px;opacity:.6;margin-top:3px\">"+hoje+" às "+hora+" · Base: "+nomeBase+"</div></div>" +
    "<div style=\"display:flex;gap:8px;flex-wrap:wrap;margin-bottom:18px\">" +
    C.kpi("Total em Andamento",totalGeral,accent)+C.kpi("Áreas com processos",areas.length,"#8e44ad")+C.kpi("Média Aging Geral",mediaGeral+" d.u.",mediaGeral>60?"#c0392b":"#2e86c1")+"</div>" +
    C.sazon(base,accent) +
    C.h2("📂 Resumo por Área") + C.tbl(["Área","Em Andamento","Críticos >60 d.u.","Média Aging"],resumoRows) +
    "</div>";

  var sections = areas.map(gerarSecaoArea).filter(function(s){return s!=="";});

  return "<!DOCTYPE html><html lang=\"pt-BR\"><head><meta charset=\"UTF-8\"><title>GAQ Sexta — Por Área — "+hoje+"</title><style>"+css+"</style></head><body>" +
    capa + sections.join("") +
    "<div style=\"text-align:center;font-size:10px;color:#aaa;margin-top:20px;border-top:1px solid #e0e7ef;padding-top:10px\">Gerado pelo Painel GAQ · "+new Date().toLocaleString("pt-BR")+"</div>" +
    "<script>setTimeout(function(){ window.print(); }, 600);<\/script></body></html>";
}

function gerarRelatorioAlertaRC({ processos, area, meta, fBase }) {
  var hoje = new Date().toLocaleDateString("pt-BR", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  var hora = new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });
  var nomeBase = meta ? meta.nomeArq : "—";
  var titulo = (!area || area === "TODAS" || area === "Todas as Áreas") ? "Todas as Áreas" : area;
  var filtroInfo = (!area || area === "TODAS" || area === "Todas as Áreas") ? "" : " · Área: <strong>" + area + "</strong>";
  var accent = "#1a5276";
  var C = _pdfCommon(accent);

  // Métricas Crítico
  var total = processos.length;
  var slaVencido = processos.filter(function(r){return r.atrasoGeral;}).length;
  var crit100 = processos.filter(function(r){return (r.diasTotaisGestao||0)>100;}).length;

  // Métricas Atenção
  var atenc60_100 = processos.filter(function(r){var d=r.diasTotaisGestao||0;return d>60&&d<=100;}).length;
  var parados = processos.filter(function(r){return (r.diasParado||0)>15;}).length;
  var mediaGestao = total ? Math.round(processos.reduce(function(s,r){return s+(r.diasTotaisGestao||0);},0)/total) : 0;

  var areaMap={};processos.forEach(function(r){var a=r["Área Requisitante"]||"N/I";areaMap[a]=(areaMap[a]||0)+1;});
  var areaTable=C.tbl(["Área Requisitante","Qtd"],Object.entries(areaMap).sort(function(a,b){return b[1]-a[1];}).map(function(e){return C.row([e[0],e[1]]);}).join(""));
  var modMap={};processos.forEach(function(r){var m=r.Modalidade||"N/I";modMap[m]=(modMap[m]||0)+1;});
  var modTable=C.tbl(["Modalidade","Qtd"],Object.entries(modMap).sort(function(a,b){return b[1]-a[1];}).map(function(e){return C.row([e[0],e[1]]);}).join(""));
  var sorted=processos.slice().sort(function(a,b){return (b.diasTotaisGestao||0)-(a.diasTotaisGestao||0);});
  var listaRows=sorted.map(function(r){var dg=r.diasTotaisGestao||0;var corDu=dg>100?"#922b21":dg>50?"#c0392b":dg>30?"#e67e22":"#27ae60";var duSD=r.diasAgingSD>0?r.diasAgingSD:(r.diasSD>0?r.diasSD:"—");var corSD=(r.diasAgingSD||r.diasSD)>50?"#c0392b":(r.diasAgingSD||r.diasSD)>20?"#e67e22":"#555";return C.row([r.NumRC||"—",r.TicketSD?"Pré-compra: "+r.TicketSD:"—",r.respNCL||r.Pregoeiro||"—",r["Área Requisitante"]||"—",r.Modalidade||"—",r.statusDet||r.status||"—",r.faseAtual||"—","<span style=\"color:"+corDu+";font-weight:700\">"+dg+"</span>",duSD!=="—"?"<span style=\"color:"+corSD+";font-weight:700\">"+duSD+"</span>":"—",r.NumProcesso||"—",(r.Objeto||"—").slice(0,500)+((r.Objeto||"").length>500?"…":"")]);}).join("");

  // KPI translúcido (glass) sobre o gradient
  function gKpi(label, value, valColor) {
    return "<div style=\"background:rgba(255,255,255,.10);border:1px solid rgba(255,255,255,.18);border-radius:12px;padding:14px 16px;-webkit-print-color-adjust:exact;print-color-adjust:exact\">"+
      "<div style=\"font-size:10px;font-weight:600;opacity:.78;letter-spacing:.06em;text-transform:uppercase;margin-bottom:6px;color:#fff\">"+label+"</div>"+
      "<div style=\"font-size:30px;font-weight:700;line-height:1;color:"+(valColor||"#fff")+";letter-spacing:-.02em;font-variant-numeric:tabular-nums\">"+value+"</div>"+
      "</div>";
  }

  var css="*{box-sizing:border-box;margin:0;padding:0}"+
    "body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f7;color:#1d1d1f;padding:20px}"+
    "@media print{body{background:#fff;padding:8px}*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important}}";

  return "<!DOCTYPE html><html lang=\"pt-BR\"><head><meta charset=\"UTF-8\"><title>Alerta RC — "+titulo+"</title><style>"+css+"</style></head><body>"+
    // ═════════ Header gradient (mesmo padrão da tela) ═════════
    "<div style=\"background:linear-gradient(135deg,#0a1f3d 0%,#2980b9 100%);color:#fff;border-radius:16px;padding:24px 28px;margin-bottom:18px;box-shadow:0 4px 16px rgba(10,31,61,.18);-webkit-print-color-adjust:exact;print-color-adjust:exact\">"+
      "<div style=\"display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:12px;margin-bottom:18px\">"+
        "<div>"+
          "<div style=\"font-size:11px;font-weight:600;opacity:.72;letter-spacing:.05em;text-transform:uppercase;margin-bottom:4px\">Alerta RC · Processos com RC em andamento</div>"+
          "<div style=\"font-size:22px;font-weight:700;letter-spacing:-.02em\">Acompanhamento de RCs</div>"+
          "<div style=\"font-size:12px;opacity:.82;margin-top:2px\">"+titulo+filtroInfo+" · Aging desde Recebimento RC · SLA por modalidade</div>"+
          "<div style=\"font-size:10px;opacity:.6;margin-top:2px\">"+hoje+" às "+hora+" · Base: "+nomeBase+"</div>"+
        "</div>"+
        "<div style=\"font-size:10px;opacity:.7;text-align:right\">GAQ — Gerência de Aquisições · DN</div>"+
      "</div>"+
      "<div style=\"display:grid;grid-template-columns:1fr 1fr;gap:18px\">"+
        "<div>"+
          "<div style=\"font-size:10px;font-weight:700;opacity:.82;letter-spacing:.06em;text-transform:uppercase;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid rgba(255,255,255,.18)\">Crítico — SLA Vencido / Aging Alto</div>"+
          "<div style=\"display:grid;grid-template-columns:repeat(3,1fr);gap:10px\">"+
            gKpi("Total RC em Andamento",total,"#fff")+
            gKpi("SLA Vencido",slaVencido,slaVencido>0?"#fecaca":"#fff")+
            gKpi("Crítico > 100 d.u.",crit100,crit100>0?"#fecaca":"#fff")+
          "</div>"+
        "</div>"+
        "<div>"+
          "<div style=\"font-size:10px;font-weight:700;opacity:.82;letter-spacing:.06em;text-transform:uppercase;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid rgba(255,255,255,.18)\">Atenção — Monitoramento</div>"+
          "<div style=\"display:grid;grid-template-columns:repeat(3,1fr);gap:10px\">"+
            gKpi("Atenção 60–100 d.u.",atenc60_100,atenc60_100>0?"#fed7aa":"#fff")+
            gKpi("Parados > 15 d",parados,parados>0?"#fed7aa":"#fff")+
            gKpi("Média d.u. (RC)",mediaGestao+" d.u.",mediaGestao>60?"#fecaca":"#fff")+
          "</div>"+
        "</div>"+
      "</div>"+
    "</div>"+

    C.sazon(fBase,accent)+C.funil(processos,accent)+
    C.h2("Status Detalhado")+C.statusDetSection(processos,total)+
    "<div style=\"display:flex;gap:16px;flex-wrap:wrap;margin-bottom:4px\"><div style=\"flex:1;min-width:200px\">"+C.h2("Por Área")+areaTable+"</div><div style=\"flex:1;min-width:200px\">"+C.h2("Por Modalidade")+modTable+"</div></div>"+
    C.h2("Lista Completa RC em Andamento ("+total+") — d.u. desde Recebimento RC")+C.tbl(["RC","Pré-compra","Responsável","Área","Modalidade","Status","Fase","d.u. RC","d.u. pré-compra","Nº Processo","Objeto"],listaRows)+
    "<div style=\"text-align:center;font-size:10px;color:#aaa;margin-top:20px;border-top:1px solid #e0e7ef;padding-top:10px\">Gerado pelo Painel GAQ · "+new Date().toLocaleString("pt-BR")+"</div>"+
    "<script>setTimeout(function(){ window.print(); }, 400);<\/script></body></html>";
}

function gerarRelatorioAlertaSD({ sdAberto, sdEncerrado, area, meta, fBase, sdSummaries }) {
  sdAberto = sdAberto || [];
  sdEncerrado = sdEncerrado || [];
  sdSummaries = sdSummaries || {};
  var hoje = new Date().toLocaleDateString("pt-BR", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  var hora = new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });
  var nomeBase = meta ? meta.nomeArq : "—";
  var titulo = (!area || area === "TODAS" || area === "Todas as Áreas") ? "Todas as Áreas" : area;
  var filtroInfo = (!area || area === "TODAS" || area === "Todas as Áreas") ? "" : " · Área: <strong>" + area + "</strong>";
  var prazoSD = typeof PRAZO_SD !== "undefined" ? PRAZO_SD : 10;
  var accent = "#1e8449";
  var C = _pdfCommon(accent);

  // ── Métricas SD em Andamento (SLA 10 d.u.) ──
  var totalAb = sdAberto.length;
  var critAb = sdAberto.filter(function(r){ return (r.diasSDAberto||0) > prazoSD; }).length;
  var atencAb = sdAberto.filter(function(r){ var d=r.diasSDAberto||0; return d > 7 && d <= prazoSD; }).length;
  var mediaAb = totalAb ? Math.round(sdAberto.reduce(function(s,r){ return s+(r.diasSDAberto||0); },0)/totalAb) : 0;
  var areaMapAb={};sdAberto.forEach(function(r){var a=r["Área Requisitante"]||"N/I";areaMapAb[a]=(areaMapAb[a]||0)+1;});
  var compMapAb={};sdAberto.forEach(function(r){var c=r.respNCL||r.Comprador||"N/I";compMapAb[c]=(compMapAb[c]||0)+1;});

  // ── Métricas SD Encerrado (aguardando RC pela área) ──
  var totalEnc = sdEncerrado.length;
  var urgEnc = sdEncerrado.filter(function(r){ return (r.diasDesdeEncSD||0) > 7; }).length;
  var mediaEnc = totalEnc ? Math.round(sdEncerrado.reduce(function(s,r){ return s+(r.diasDesdeEncSD||0); },0)/totalEnc) : 0;
  var areaMapEnc={};sdEncerrado.forEach(function(r){var a=r["Área Requisitante"]||"N/I";areaMapEnc[a]=(areaMapEnc[a]||0)+1;});

  // ── Tabelas de distribuição ──
  var areaTableAb = C.tbl(["Área Requisitante","Qtd"],Object.entries(areaMapAb).sort(function(a,b){return b[1]-a[1];}).map(function(e){return C.row([e[0],e[1]]);}).join(""));
  var compTableAb = C.tbl(["Responsável","Qtd"],Object.entries(compMapAb).sort(function(a,b){return b[1]-a[1];}).map(function(e){return C.row([e[0],e[1]]);}).join(""));
  var areaTableEnc = C.tbl(["Área Requisitante","Qtd"],Object.entries(areaMapEnc).sort(function(a,b){return b[1]-a[1];}).map(function(e){return C.row([e[0],e[1]]);}).join(""));

  // ── Lista SD em Andamento ──
  var listaAbRows = sdAberto.slice().sort(function(a,b){return (b.diasSDAberto||0)-(a.diasSDAberto||0);}).map(function(r){
    var d=r.diasSDAberto||0;
    var cor=d>prazoSD?"#c0392b":d>7?"#e67e22":"#27ae60";
    var sd=sdSummaries[String(r.TicketSD||"").trim()];
    var duI=sd&&sd.diasDesdeUltimaInteracao!=null?sd.diasDesdeUltimaInteracao:null;
    var duICor=duI==null?"#6b7280":duI>10?"#c0392b":duI>5?"#e67e22":"#27ae60";
    return C.row([
      r.TicketSD||"—",
      r.respNCL||r.Comprador||"—",
      r["Área Requisitante"]||"—",
      r.statusDet||r.status||"—",
      "<span style=\"color:"+cor+";font-weight:700\">"+d+" d.u.</span>",
      r.aberturaSD?r.aberturaSD.toLocaleDateString("pt-BR"):"—",
      sd?sd.comQuem:"—",
      duI!=null?"<span style=\"color:"+duICor+";font-weight:700\">"+duI+" d.u.</span>":"—",
      sd?(sd.acao||"—").slice(0,80):"—",
      (r.Objeto||"—").slice(0,120)+((r.Objeto||"").length>120?"…":"")
    ]);
  }).join("");

  // ── Lista SD Encerrado ──
  var listaEncRows = sdEncerrado.slice().sort(function(a,b){return (b.diasDesdeEncSD||0)-(a.diasDesdeEncSD||0);}).map(function(r){
    var d=r.diasDesdeEncSD||0;
    var cor=d>7?"#c0392b":d>3?"#e67e22":"#b45309";
    var sd=sdSummaries[String(r.TicketSD||"").trim()];
    var duI=sd&&sd.diasDesdeUltimaInteracao!=null?sd.diasDesdeUltimaInteracao:null;
    var duICor=duI==null?"#6b7280":duI>10?"#c0392b":duI>5?"#e67e22":"#27ae60";
    return C.row([
      r.TicketSD||"—",
      r["Área Requisitante"]||"—",
      "<span style=\"color:"+cor+";font-weight:700\">"+d+" d.u.</span>",
      r.encSD?r.encSD.toLocaleDateString("pt-BR"):"—",
      sd?sd.comQuem:"—",
      duI!=null?"<span style=\"color:"+duICor+";font-weight:700\">"+duI+" d.u.</span>":"—",
      (r.Objeto||"—").slice(0,120)+((r.Objeto||"").length>120?"…":"")
    ]);
  }).join("");

  // KPI translúcido (glass) sobre o gradient
  function gKpi(label, value, valColor) {
    return "<div style=\"background:rgba(255,255,255,.10);border:1px solid rgba(255,255,255,.18);border-radius:12px;padding:14px 16px;-webkit-print-color-adjust:exact;print-color-adjust:exact\">"+
      "<div style=\"font-size:10px;font-weight:600;opacity:.78;letter-spacing:.06em;text-transform:uppercase;margin-bottom:6px;color:#fff\">"+label+"</div>"+
      "<div style=\"font-size:30px;font-weight:700;line-height:1;color:"+(valColor||"#fff")+";letter-spacing:-.02em;font-variant-numeric:tabular-nums\">"+value+"</div>"+
      "</div>";
  }

  var css="*{box-sizing:border-box;margin:0;padding:0}"+
    "body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f5f7;color:#1d1d1f;padding:20px}"+
    "@media print{body{background:#fff;padding:8px}*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important}}"+
    ".section-header{display:flex;align-items:center;gap:8px;margin:20px 0 10px}"+
    ".section-bar{width:4px;border-radius:2px;height:18px;flex-shrink:0}"+
    ".section-title{font-size:14px;font-weight:700}"+
    ".section-badge{border-radius:20px;padding:2px 10px;font-size:11px;font-weight:700}"+
    ".notice{border-radius:10px;padding:10px 14px;margin-bottom:12px;font-size:11px}";

  return "<!DOCTYPE html><html lang=\"pt-BR\"><head><meta charset=\"UTF-8\"><title>Alerta Pré-compra — "+titulo+"</title><style>"+css+"</style></head><body>"+
    // ═════════ Header gradient ═════════
    "<div style=\"background:linear-gradient(135deg,#064e3b 0%,#059669 100%);color:#fff;border-radius:16px;padding:24px 28px;margin-bottom:18px;box-shadow:0 4px 16px rgba(6,78,59,.18);-webkit-print-color-adjust:exact;print-color-adjust:exact\">"+
      "<div style=\"display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:12px;margin-bottom:18px\">"+
        "<div>"+
          "<div style=\"font-size:11px;font-weight:600;opacity:.72;letter-spacing:.05em;text-transform:uppercase;margin-bottom:4px\">Alerta Pré-compra · Tickets sem RC</div>"+
          "<div style=\"font-size:22px;font-weight:700;letter-spacing:-.02em\">Acompanhamento Pré-compra</div>"+
          "<div style=\"font-size:12px;opacity:.82;margin-top:2px\">"+titulo+filtroInfo+" · SLA pré-compra: "+prazoSD+" d.u.</div>"+
          "<div style=\"font-size:10px;opacity:.6;margin-top:2px\">"+hoje+" às "+hora+" · Base: "+nomeBase+"</div>"+
        "</div>"+
        "<div style=\"font-size:10px;opacity:.7;text-align:right\">GAQ — Gerência de Aquisições · DN</div>"+
      "</div>"+
      "<div style=\"display:grid;grid-template-columns:1fr 1fr;gap:18px\">"+
        "<div>"+
          "<div style=\"font-size:10px;font-weight:700;opacity:.82;letter-spacing:.06em;text-transform:uppercase;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid rgba(255,255,255,.18)\">Pré-compra em andamento — Acompanhamento necessário GAQ</div>"+
          "<div style=\"display:grid;grid-template-columns:repeat(3,1fr);gap:10px\">"+
            gKpi("Em Andamento",totalAb,"#fff")+
            gKpi("Atraso > "+prazoSD+" d.u.",critAb,critAb>0?"#fecaca":"#fff")+
            gKpi("Atenção 7–"+prazoSD+" d.u.",atencAb,atencAb>0?"#fed7aa":"#fff")+
          "</div>"+
          "<div style=\"display:grid;grid-template-columns:1fr;gap:10px;margin-top:10px\">"+
            gKpi("Média d.u. (andamento)",mediaAb+" d.u.",mediaAb>prazoSD?"#fecaca":"#fff")+
          "</div>"+
        "</div>"+
        "<div>"+
          "<div style=\"font-size:10px;font-weight:700;opacity:.82;letter-spacing:.06em;text-transform:uppercase;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid rgba(255,255,255,.18)\">Pré-compra encerrada — Aguardando RC da Área</div>"+
          "<div style=\"display:grid;grid-template-columns:repeat(3,1fr);gap:10px\">"+
            gKpi("Enc. sem RC",totalEnc,"#fff")+
            gKpi("Urgente > 7 d.u.",urgEnc,urgEnc>0?"#fecaca":"#fff")+
            gKpi("Média d.u. s/ RC",mediaEnc+" d.u.",mediaEnc>7?"#fecaca":"#fff")+
          "</div>"+
        "</div>"+
      "</div>"+
    "</div>"+

    // ═════════ SEÇÃO 1: PRÉ-COMPRA EM ANDAMENTO ═════════
    (totalAb > 0 ? (
      "<div class=\"section-header\"><div class=\"section-bar\" style=\"background:#22c55e\"></div><div class=\"section-title\">Pré-compra em andamento</div><span class=\"section-badge\" style=\"background:#dcfce7;color:#15803d\">"+totalAb+"</span><span style=\"font-size:11px;color:#6b7280\">SLA: "+prazoSD+" d.u. · Comprador responsável pelo encerramento</span></div>"+
      "<div style=\"display:flex;gap:16px;flex-wrap:wrap;margin-bottom:8px\">"+
        "<div style=\"flex:1;min-width:200px\">"+C.h2("Por Área")+areaTableAb+"</div>"+
        "<div style=\"flex:1;min-width:200px\">"+C.h2("Por Responsável")+compTableAb+"</div>"+
      "</div>"+
      C.h2("Lista pré-compra em andamento ("+totalAb+")")+
      C.tbl(["Pré-compra","Responsável","Área","Status","d.u. SD","Abertura SD","Com quem está","Últ. inter.","Ação necessária","Objeto"],listaAbRows)
    ) : "")+

    // ═════════ SEÇÃO 2: PRÉ-COMPRA ENCERRADA AGUARDANDO RC ═════════
    (totalEnc > 0 ? (
      "<div class=\"section-header\" style=\"margin-top:28px\"><div class=\"section-bar\" style=\"background:#f59e0b\"></div><div class=\"section-title\">Pré-compra encerrada — Aguardando RC</div><span class=\"section-badge\" style=\"background:#fef3c7;color:#b45309\">"+totalEnc+"</span><span style=\"font-size:11px;color:#6b7280\">Responsabilidade da Área Requisitante</span></div>"+
      "<div class=\"notice\" style=\"background:#fffbeb;border:1px solid #fde68a;color:#92400e\"><b>Cobrança à Área:</b> A pré-compra foi encerrada pelo GAQ. A Área Requisitante deve abrir a RC para dar continuidade ao processo.</div>"+
      "<div style=\"min-width:200px;margin-bottom:8px\">"+C.h2("Por Área (pré-compra encerrada)")+areaTableEnc+"</div>"+
      C.h2("Lista pré-compra encerrada aguardando RC ("+totalEnc+")")+
      C.tbl(["Pré-compra","Área","d.u. s/ RC","Enc. SD","Com quem está","Últ. inter.","Objeto"],listaEncRows)
    ) : "")+

    "<div style=\"text-align:center;font-size:10px;color:#aaa;margin-top:24px;border-top:1px solid #e0e7ef;padding-top:10px\">Gerado pelo Painel GAQ · "+new Date().toLocaleString("pt-BR")+"</div>"+
    "<script>setTimeout(function(){ window.print(); }, 400);<\/script></body></html>";
}

function exportTagsAsJS() {
  const data = TagsManager.exportTags();
  const content = "window.__PAINEL_TAGS__ = " + JSON.stringify(data, null, 2) + ";";
  const blob = new Blob([content], { type: "application/javascript" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = "tags.js"; a.click();
  URL.revokeObjectURL(url);
}

// ── Cronograma de Processo ──────────────────────────────────────────────────

function addBD(date, n) {
  if (!date || n <= 0) return new Date(date);
  const d = new Date(date); d.setHours(0,0,0,0);
  let added = 0;
  while (added < n) { d.setDate(d.getDate()+1); const w=d.getDay(); if(w&&w<6) added++; }
  return d;
}

function buildCronograma(proc, phaseIntervals) {
  const mn = (proc.Modalidade||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"");
  const isCPL = mn.includes("pregao") || mn.includes("concorrencia") || mn.includes("dialogo");
  const meta = isCPL ? PRAZO_LICITACAO : PRAZO_COMPRA_DIRETA;

  // SLA dinâmico baseado no histórico (ou fallback estático)
  const dynamicSLA = phaseIntervals ? buildDynamicSLA(phaseIntervals, mn) : null;
  const slaRC = dynamicSLA || (isCPL ? SLA_CPL : SLA_SIMPLES);

  // Fase SD: prazo apartado de 10 d.u. (SLA_SD)
  const sdKeys = Object.keys(SLA_SD);
  // Fase RC: SLA principal (30/90 d.u.)
  const rcKeys = Object.keys(slaRC);

  // Montar colunas combinadas: SD + RC (sem duplicar chaves)
  const allCols = [];
  // Primeiro as fases SD
  sdKeys.forEach(key => {
    const tlCol = TL_COLS.find(([,k]) => k === key);
    if (tlCol) allCols.push({ label: tlCol[0], key, sla: SLA_SD[key], zone: "SD" });
  });
  // Depois as fases RC (a partir do Recebimento RC)
  rcKeys.forEach(key => {
    // Evitar duplicatas (SD já adicionou Abertura/Distribuição/Encerramento)
    if (sdKeys.includes(key)) return;
    const tlCol = TL_COLS.find(([,k]) => k === key);
    if (tlCol) allCols.push({ label: tlCol[0], key, sla: slaRC[key], zone: "RC" });
  });

  let firstIdx = -1;
  for (let i=0; i<allCols.length; i++) { if(pd(proc[allCols[i].key])){firstIdx=i;break;} }
  if (firstIdx < 0) return null;

  const rows = [];
  let prevBase = null;
  let prevZone = null;

  for (let i=0; i<allCols.length; i++) {
    const { label, key, sla: slaDu, zone } = allCols[i];
    const actual = pd(proc[key]);

    // Quando transiciona de SD para RC, resetar prevBase para a data real do Recebimento RC
    const isZoneChange = prevZone === "SD" && zone === "RC";

    let planned;
    if (i < firstIdx) {
      planned = null;
    } else if (i === firstIdx || isZoneChange) {
      planned = actual || (prevBase && slaDu != null ? addBD(prevBase, slaDu) : null) || new Date();
    } else if (prevBase && slaDu != null) {
      planned = addBD(prevBase, slaDu);
    } else {
      planned = null;
    }

    let status="—", variacao="—";
    if (planned) {
      if (actual) {
        const late = du(planned, actual);
        const early = du(actual, planned);
        if (actual <= planned) {
          status = "No prazo"; variacao = early > 0 ? "-"+early+" d.u." : "0 d.u.";
        } else {
          status = "Atrasado"; variacao = "+"+late+" d.u.";
        }
      } else {
        status = "Pendente";
      }
    } else if (actual) {
      status = "Registrado";
    } else if (i >= firstIdx) {
      status = "Pendente";
    }

    // SLA label diferente para zona SD vs RC
    let slaLabel;
    if (slaDu === 0) slaLabel = "Marco inicial";
    else if (slaDu == null) slaLabel = "A definir";
    else slaLabel = slaDu + " d.u.";

    rows.push({ num:i+1, key, etapa:label, responsavel:PHASE_RESP[key]||"—",
      sla: slaLabel, zone,
      planejada: planned ? planned.toLocaleDateString("pt-BR") : "—",
      real: actual ? actual.toLocaleDateString("pt-BR") : "—",
      status, variacao, _ok:!!actual });

    prevBase = (slaDu == null && !actual) ? null : (actual || planned || prevBase);
    prevZone = zone;
  }

  // Reposicionar fases voláteis (Indicação Analista) na ordem cronológica do processo.
  if (typeof VOLATILE_TIMELINE_KEYS !== "undefined") {
    VOLATILE_TIMELINE_KEYS.forEach(volKey => {
      const idxRow = rows.findIndex(r => r.key === volKey);
      if (idxRow < 0) return;
      const actualDate = pd(proc[volKey]);
      if (!actualDate) return; // sem data → mantém posição default
      const entry = rows[idxRow];
      rows.splice(idxRow, 1);
      // Inserir antes do primeiro row com data real posterior
      let insertAt = rows.length;
      for (let i = 0; i < rows.length; i++) {
        const oDate = pd(proc[rows[i].key]);
        if (oDate && oDate > actualDate) { insertAt = i; break; }
      }
      rows.splice(insertAt, 0, entry);
    });
    // Renumera após a reordenação
    rows.forEach((r, i) => { r.num = i + 1; });
  }

  return { rows, isCPL, metaTotal: meta, metaSD: PRAZO_SD, proc, dynamicSLA: !!dynamicSLA };
}

function exportCronogramaExcel(proc, phaseIntervals) {
  const cron = buildCronograma(proc, phaseIntervals);
  if (!cron) { alert("Nenhuma data preenchida para gerar o cronograma."); return; }
  const wb = window.XLSX.utils.book_new();
  const info = [
    ["Campo","Valor"],
    ["No RC", proc.NumRC||"—"], ["Ticket pré-compra", proc.TicketSD||"—"],
    ["Modalidade", proc.Modalidade||"—"],
    ["Area Requisitante", proc["Área Requisitante"]||proc.AreaReq||"—"],
    ["Objeto", proc.Objeto||"—"], ["Status", proc.status||"—"],
    ["Status Detalhado", proc.statusDet||"—"],
    ["Comprador NCL", proc.Comprador||"—"], ["Avaliador", proc.Avaliador||"—"],
    ["Resp CPL", proc.cplResp||"—"], ["Pregoeiro", proc.Pregoeiro||"—"],
    ["Analista SCONT", proc.AnalistaContrato||"—"],
    ["Prazo pré-compra (apartado)", PRAZO_SD+" d.u."],
    ["SLA Meta Total (a partir RC)", cron.metaTotal+" d.u."],
    ["SLA baseado em historico", cron.dynamicSLA ? "Sim" : "Nao (fallback)"],
    ["Gerado em", new Date().toLocaleDateString("pt-BR")],
  ];
  const ws1 = window.XLSX.utils.aoa_to_sheet(info);
  ws1["!cols"] = [{wch:30},{wch:60}];
  window.XLSX.utils.book_append_sheet(wb, ws1, "Processo");
  const header = ["No","Etapa","Responsavel","SLA Meta","Data Planejada","Data Real","Status","Variacao"];
  const data = cron.rows.map(r => [r.num,r.etapa,r.responsavel,r.sla,r.planejada,r.real,r.status,r.variacao]);
  const ws2 = window.XLSX.utils.aoa_to_sheet([header,...data]);
  ws2["!cols"] = [{wch:4},{wch:36},{wch:28},{wch:14},{wch:16},{wch:16},{wch:12},{wch:14}];
  window.XLSX.utils.book_append_sheet(wb, ws2, "Cronograma");
  const nome = (proc.NumRC||proc.TicketSD||"processo").replace(/[^a-z0-9]/gi,"_");
  window.XLSX.writeFile(wb, "cronograma_"+nome+".xlsx");
}

function exportCronogramaPDF(proc, phaseIntervals) {
  const cron = buildCronograma(proc, phaseIntervals);
  if (!cron) { alert("Nenhuma data preenchida para gerar o cronograma."); return; }
  const corBg = (r) => r.status==="No prazo"?"#e8f8f0":r.status==="Atrasado"?"#fdf0f0":r.status==="Registrado"?"#eef3fd":r.status==="Pendente"?"#fefaf0":"#f5f5f5";
  const corVar = (v) => v.startsWith("+")?"#c0392b":v.startsWith("-")?"#1e8449":"#555";
  let prevZone = null;
  const linhas = cron.rows.map(r => {
    let sep = "";
    if (prevZone === "SD" && r.zone === "RC") {
      sep = "<tr><td colspan='9' style='background:#1a5276;color:#fff;text-align:center;font-weight:700;padding:6px;font-size:11px'>SLA: "+cron.metaTotal+" d.u. a partir do Recebimento RC</td></tr>";
    }
    prevZone = r.zone;
    const zoneBadge = r.zone === "SD" ? "<span style='background:#e67e22;color:#fff;border-radius:3px;padding:0 5px;font-size:9px;margin-left:4px'>Pré-compra</span>" : "";
    return sep+"<tr style='background:"+corBg(r)+"'><td style='text-align:center;color:#555'>"+r.num+"</td><td style='font-weight:600'>"+r.etapa+zoneBadge+"</td><td style='color:#555;font-size:11px'>"+r.responsavel+"</td><td style='text-align:center;color:#1a5276'>"+r.sla+"</td><td style='text-align:center'>"+r.planejada+"</td><td style='text-align:center;font-weight:"+(r._ok?700:400)+";color:"+(r._ok?"#1e8449":"#999")+"'>"+r.real+"</td><td style='text-align:center;font-size:12px'>"+r.status+"</td><td style='text-align:center;color:"+corVar(r.variacao)+";font-weight:600'>"+r.variacao+"</td></tr>";
  }).join("");
  const gerado = new Date().toLocaleDateString("pt-BR");
  const html = "<!DOCTYPE html><html><head><meta charset='utf-8'><title>Cronograma "+( proc.NumRC||proc.TicketSD||"")+"</title><style>body{font-family:Arial,sans-serif;margin:0;padding:24px;color:#222;font-size:13px}.hdr{background:#1a3a5c;color:#fff;padding:16px 24px;border-radius:8px;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center}.info{background:#f0f4f8;border-radius:6px;padding:12px 18px;margin-bottom:14px;display:grid;grid-template-columns:1fr 1fr;gap:5px 24px;font-size:12px}.info .lbl{color:#555;font-weight:600;min-width:130px}.meta{background:#1a5276;color:#fff;border-radius:5px;padding:5px 14px;display:inline-block;font-size:12px;font-weight:700;margin-bottom:12px}table{width:100%;border-collapse:collapse;font-size:12px}th{background:#1a3a5c;color:#fff;padding:8px 10px;text-align:left;font-size:11px}td{padding:7px 10px;border-bottom:1px solid #ddd}.ftr{margin-top:16px;font-size:10px;color:#999;text-align:right;border-top:1px solid #ddd;padding-top:8px}@media print{.np{display:none}body{padding:12px}}</style></head><body>"
    +"<div class='np' style='margin-bottom:14px'><button onclick='window.print()' style='background:#1a3a5c;color:#fff;border:none;padding:10px 24px;border-radius:6px;cursor:pointer;font-size:14px;font-weight:700'>Imprimir / Salvar PDF</button></div>"
    +"<div class='hdr'><div><div style='font-size:18px;font-weight:800'>Cronograma de Processo</div><div style='font-size:11px;opacity:.8'>Painel de Dados de Compras GAQ · Sesc Departamento Nacional</div></div><div style='text-align:right;font-size:11px;opacity:.8'>Gerado em "+gerado+"<br/>SLA: "+(cron.isCPL?"Pregao/Concorrencia":"Demais modalidades")+"</div></div>"
    +"<div class='info'>"
    +"<div style='display:flex;gap:8px'><span class='lbl'>No RC:</span><span>"+(proc.NumRC||"—")+"</span></div>"
    +"<div style='display:flex;gap:8px'><span class='lbl'>Ticket pré-compra:</span><span>"+(proc.TicketSD||"—")+"</span></div>"
    +"<div style='display:flex;gap:8px'><span class='lbl'>Modalidade:</span><span><b>"+(proc.Modalidade||"—")+"</b></span></div>"
    +"<div style='display:flex;gap:8px'><span class='lbl'>Area Requisitante:</span><span>"+(proc["Área Requisitante"]||proc.AreaReq||"—")+"</span></div>"
    +"<div style='display:flex;gap:8px;grid-column:1/-1'><span class='lbl'>Objeto:</span><span>"+(proc.Objeto||"—")+"</span></div>"
    +"<div style='display:flex;gap:8px'><span class='lbl'>Status:</span><span>"+(proc.status||"—")+"</span></div>"
    +"<div style='display:flex;gap:8px'><span class='lbl'>Comprador NCL:</span><span>"+(proc.Comprador||"—")+"</span></div>"
    +"<div style='display:flex;gap:8px'><span class='lbl'>Resp. CPL:</span><span>"+(proc.cplResp||"—")+"</span></div>"
    +"<div style='display:flex;gap:8px'><span class='lbl'>Analista SCONT:</span><span>"+(proc.AnalistaContrato||"—")+"</span></div>"
    +"</div>"
    +"<div class='meta'>Pré-compra: "+PRAZO_SD+" d.u. (apartado) | SLA Meta: "+cron.metaTotal+" d.u. a partir do Recebimento RC | "+(cron.isCPL?"Pregao / Concorrencia":"Demais Modalidades")+(cron.dynamicSLA?" | SLA baseado no historico":"")+"</div>"
    +"<table><thead><tr><th style='width:30px'>No</th><th>Etapa</th><th>Responsavel</th><th style='width:90px;text-align:center'>SLA Meta</th><th style='width:110px;text-align:center'>Data Planejada</th><th style='width:110px;text-align:center'>Data Real</th><th style='width:100px;text-align:center'>Status</th><th style='width:100px;text-align:center'>Variacao</th></tr></thead><tbody>"+linhas+"</tbody></table>"
    +"<div class='ftr'>Painel de Dados de Compras GAQ · Sesc Departamento Nacional · "+new Date().getFullYear()+"</div>"
    +"</body></html>";
  const w = window.open("","_blank","width=1100,height=800");
  if (w) { w.document.write(html); w.document.close(); }
}
