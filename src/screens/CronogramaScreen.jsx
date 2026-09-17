import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function CronogramaScreen({ calDiaSel, calMes, fArea, fBack, futurosTodosOpen, loginUser, meta, setCalDiaSel, setCalMes, setFuturosTodosOpen, setSelProc, setVencidosOpen, vencidosOpen }) {
  return ((() => {
          const hoje0 = new Date(); hoje0.setHours(0,0,0,0);
          // procEntregas = todos do fBack (filtros aplicados, só em andamento) que têm dataEntrega
          const procEntregas = fBack.filter(r => r.dataEntrega);
          // Badge cor por dias restantes
          function corEntrega(dias) {
            if (dias === null) return "#8a94a6";
            if (dias < 0)  return "#e74c3c";   // vencido
            if (dias <= 7)  return "#ff6b6b";   // urgente ≤7 dias
            if (dias <= 30) return "#fd9644";   // próximo ≤30 dias
            return "#27ae60"; // futuro — verde
          }
          function labelEntrega(dias) {
            if (dias === null) return "";
            if (dias < 0)  return `Vencido há ${Math.abs(dias)} dia(s)`;
            if (dias === 0) return "Entrega HOJE";
            if (dias === 1) return "Entrega AMANHÃ";
            return `${dias} dia(s)`;
          }
          // Estado local do mês navegado
          const primeiroDia = new Date(calMes.y, calMes.m, 1);
          const ultimoDia   = new Date(calMes.y, calMes.m + 1, 0);
          const totalDias   = ultimoDia.getDate();
          const inicioSem   = primeiroDia.getDay(); // 0=dom

          // Map: dia -> processos do mês
          const mapDia = {};
          procEntregas.forEach(r => {
            const d = r.dataEntrega;
            if (d.getFullYear() === calMes.y && d.getMonth() === calMes.m) {
              const k = d.getDate();
              if (!mapDia[k]) mapDia[k] = [];
              mapDia[k].push(r);
            }
          });

          // Alertas: vencidos + próximos 30 dias
          const alertas = procEntregas.filter(r => r.diasParaEntrega !== null && r.diasParaEntrega <= 30).sort((a,b) => a.diasParaEntrega - b.diasParaEntrega);
          const vencidos = procEntregas.filter(r => r.diasParaEntrega !== null && r.diasParaEntrega < 0);

          // Processos do dia selecionado
          const procDiaSel = calDiaSel ? (mapDia[calDiaSel] || []) : [];

          const nomeMes = primeiroDia.toLocaleDateString("pt-BR", { month: "long", year: "numeric" });
          const diasSemana = ["Dom","Seg","Ter","Qua","Qui","Sex","Sáb"];

          // Processos do mês corrente com entrega >= hoje (não vencidos)
          const mesCorrente = procEntregas.filter(r =>
            r.dataEntrega.getFullYear() === calMes.y &&
            r.dataEntrega.getMonth()    === calMes.m &&
            r.diasParaEntrega !== null  && r.diasParaEntrega >= 0
          ).sort((a, b) => a.diasParaEntrega - b.diasParaEntrega);

          // Todos os processos com entrega no mês selecionado (inclui vencidos)
          const procDoMes = procEntregas.filter(r =>
            r.dataEntrega.getFullYear() === calMes.y &&
            r.dataEntrega.getMonth()    === calMes.m
          ).sort((a, b) => a.dataEntrega - b.dataEntrega);

          return <div className="anim-fade">
            {/* Header */}
            <div style={{ marginBottom: 18 }}>
              <div className="gaq-h1" style={{ marginBottom: 4 }}>Cronograma de Entregas</div>
              <div className="gaq-body" style={{ color: "var(--gaq-text-3)" }}>Previsao de entrega / inicio de prestacao · {procEntregas.length} processo(s) com data cadastrada{loginUser ? ` · ${loginUser}` : ""}</div>
            </div>

            <div className="gaq-kpi-grid" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", marginBottom: 18 }}>
              {[
                { label: "Vencidos",  count: vencidos.length, color: "#e74c3c" },
                { label: "≤ 7 dias",  count: procEntregas.filter(r => r.diasParaEntrega >= 0 && r.diasParaEntrega <= 7).length, color: "#ff6b6b" },
                { label: "≤ 30 dias", count: procEntregas.filter(r => r.diasParaEntrega > 7 && r.diasParaEntrega <= 30).length, color: "#fd9644" },
                { label: "Futuros",   count: procEntregas.filter(r => r.diasParaEntrega > 30).length, color: "#27ae60" },
              ].map(({ label, count, color }) => (
                <div key={label} className="gaq-kpi" style={{ borderTop: `3px solid ${color}` }}>
                  <span className="lbl">{label}</span>
                  <div className="val" style={{ color }}>{count}</div>
                </div>
              ))}
            </div>

            {/* ── 1. CALENDARIO ── */}
            <div className="gaq-card" style={{ padding: 22, marginBottom: 16 }}>
              {/* Navegacao de mes */}
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <button onClick={() => { setCalMes(p => { const d = new Date(p.y, p.m - 1, 1); return { y: d.getFullYear(), m: d.getMonth() }; }); setCalDiaSel(null); }}
                  className="gaq-btn" style={{ fontSize: 14 }}>‹</button>
                <div className="gaq-h3" style={{ textTransform: "capitalize" }}>{nomeMes}</div>
                <button onClick={() => { setCalMes(p => { const d = new Date(p.y, p.m + 1, 1); return { y: d.getFullYear(), m: d.getMonth() }; }); setCalDiaSel(null); }}
                  className="gaq-btn" style={{ fontSize: 14 }}>›</button>
              </div>
              {/* Toolbar do mês selecionado: total + ações (PDF / e-mail) */}
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 16, padding: "10px 14px", borderRadius: 10, background: "var(--card2)", border: "1px solid var(--border2)" }}>
                <div style={{ display: "flex", alignItems: "baseline", gap: 8 }}>
                  <span style={{ fontSize: 11, color: "var(--text3)", textTransform: "uppercase", fontWeight: 600, letterSpacing: ".04em" }}>Total no mês</span>
                  <span style={{ fontSize: 22, fontWeight: 800, color: "#e67e22", lineHeight: 1 }}>{procDoMes.length}</span>
                  <span style={{ fontSize: 11, color: "var(--text3)" }}>processo(s) com entrega prevista</span>
                </div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  <button className="gaq-btn no-print" disabled={procDoMes.length === 0} onClick={() => {
                    const html = gerarRelatorioCronograma({ processos: procDoMes, mes: calMes.m, ano: calMes.y, area: fArea || "Todas as Áreas", meta });
                    const win = window.open("", "_blank");
                    win.document.write(html);
                    win.document.close();
                  }} title="Gera PDF do cronograma do mês selecionado">
                    <Icon name="doc" size={14}/> Gerar Relatório
                  </button>
                  <button className="gaq-btn no-print" disabled={procDoMes.length === 0} onClick={() => {
                    const area = fArea || "Todas as Áreas";
                    const MESES_PT = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
                    const nomeMesPt = MESES_PT[calMes.m] + " de " + calMes.y;
                    const html = gerarRelatorioCronograma({ processos: procDoMes, mes: calMes.m, ano: calMes.y, area, meta });
                    const win = window.open("", "_blank");
                    win.document.write(html);
                    win.document.close();
                    const venc = procDoMes.filter(r => r.diasParaEntrega !== null && r.diasParaEntrega < 0).length;
                    const ate7 = procDoMes.filter(r => r.diasParaEntrega !== null && r.diasParaEntrega >= 0 && r.diasParaEntrega <= 7).length;
                    const ate30 = procDoMes.filter(r => r.diasParaEntrega !== null && r.diasParaEntrega > 7 && r.diasParaEntrega <= 30).length;
                    const futs = procDoMes.filter(r => r.diasParaEntrega !== null && r.diasParaEntrega > 30).length;
                    const subject = `Cronograma de Entregas — ${nomeMesPt} — ${area}`;
                    const linhasProc = procDoMes.slice(0, 30).map(r => {
                      const d = r.dataEntrega.toLocaleDateString("pt-BR");
                      const id = r.NumRC || r.TicketSD || "—";
                      const ar = r["Área Requisitante"] || "—";
                      const obj = (r.Objeto || "").slice(0, 80);
                      const sit = r.diasParaEntrega < 0 ? `vencido há ${Math.abs(r.diasParaEntrega)}d` : `em ${r.diasParaEntrega}d`;
                      return `- ${d} · ${id} · ${ar} · ${sit} · ${obj}`;
                    }).join("\n");
                    const body =
`Prezados,

Segue o cronograma de entregas previstas para ${nomeMesPt}.

Área: ${area}
Total no mês: ${procDoMes.length} processo(s)
- Vencidos: ${venc}
- Próximos 7 dias: ${ate7}
- 8 a 30 dias: ${ate30}
- Mais de 30 dias: ${futs}

${procDoMes.length > 30 ? "Primeiros 30 processos (relatório completo em anexo PDF):" : "Processos:"}
${linhasProc}
${procDoMes.length > 30 ? `\n... e mais ${procDoMes.length - 30} processo(s) no relatório anexo.` : ""}

Relatório completo em anexo (PDF — gerado em nova aba).

Atenciosamente,
GAQ — Gerência de Aquisições`;
                    setTimeout(() => {
                      const a = document.createElement("a");
                      a.href = "mailto:?subject=" + encodeURIComponent(subject) + "&body=" + encodeURIComponent(body);
                      a.click();
                    }, 800);
                  }} title="Gera o PDF e abre o e-mail com o resumo do mês">
                    <Icon name="external" size={14}/> Enviar por E-mail
                  </button>
                </div>
              </div>
              {/* Cabeçalho dias da semana */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 3, marginBottom: 3 }}>
                {diasSemana.map(d => <div key={d} style={{ textAlign: "center", fontSize: 10, fontWeight: 700, color: "var(--text3)", padding: "4px 0", textTransform: "uppercase" }}>{d}</div>)}
              </div>
              {/* Grid de dias */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 3 }}>
                {Array.from({ length: inicioSem }).map((_, i) => <div key={"e"+i} />)}
                {Array.from({ length: totalDias }).map((_, i) => {
                  const dia = i + 1;
                  const diaDate = new Date(calMes.y, calMes.m, dia);
                  diaDate.setHours(0,0,0,0);
                  const ehHoje = diaDate.getTime() === hoje0.getTime();
                  const procs = mapDia[dia] || [];
                  const temProc = procs.length > 0;
                  const corDom = temProc ? corEntrega(Math.round((diaDate - hoje0) / 86400000)) : null;
                  const selecionado = calDiaSel === dia;
                  return (
                    <div key={dia} onClick={() => temProc ? setCalDiaSel(selecionado ? null : dia) : null}
                      style={{ minHeight: 48, borderRadius: 8, padding: "4px 2px", textAlign: "center", cursor: temProc ? "pointer" : "default",
                        background: selecionado ? (corDom + "22") : ehHoje ? "rgba(46,134,193,0.10)" : "var(--card2)",
                        border: selecionado ? `2px solid ${corDom}` : ehHoje ? "2px solid #2e86c1" : "1px solid var(--border2)",
                        transition: "all .15s" }}>
                      <div style={{ fontSize: 11, fontWeight: ehHoje ? 800 : 500, color: ehHoje ? "#2e86c1" : "var(--text)" }}>{dia}</div>
                      {temProc && <div style={{ marginTop: 3, display: "flex", flexWrap: "wrap", gap: 2, justifyContent: "center" }}>
                        {procs.slice(0, 3).map((_, pi) => <div key={pi} style={{ width: 7, height: 7, borderRadius: "50%", background: corDom }} />)}
                        {procs.length > 3 && <div style={{ fontSize: 8, fontWeight: 800, color: corDom }}>+{procs.length-3}</div>}
                      </div>}
                    </div>
                  );
                })}
              </div>
              {/* Legenda */}
              <div style={{ display: "flex", gap: 12, marginTop: 14, flexWrap: "wrap", justifyContent: "center" }}>
                {[["#e74c3c","Vencido"],["#ff6b6b","≤ 7 dias"],["#fd9644","≤ 30 dias"],["#27ae60","Futuro"]].map(([c,l]) => (
                  <div key={l} style={{ display: "flex", alignItems: "center", gap: 4 }}>
                    <div style={{ width: 10, height: 10, borderRadius: "50%", background: c }} />
                    <span style={{ fontSize: 10, color: "var(--text3)" }}>{l}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Detalhe do dia selecionado */}
            {calDiaSel && procDiaSel.length > 0 && <div style={{ background: "var(--card)", borderRadius: 12, padding: 16, boxShadow: "var(--shadow)", marginBottom: 16, border: "1.5px solid #2e86c130" }}>
              <div style={{ fontWeight: 700, fontSize: 13, color: "var(--text)", marginBottom: 10 }}>
                📌 Entregas em {String(calDiaSel).padStart(2,"0")}/{String(calMes.m+1).padStart(2,"0")}/{calMes.y} — {procDiaSel.length} processo(s)
              </div>
              {procDiaSel.map((r, i) => {
                const cor = corEntrega(r.diasParaEntrega);
                const rAtivoColor = r.faseSubarea === "CPL" ? "#8e44ad" : r.faseSubarea === "Scont" ? "#16a085" : "#2e86c1";
                return <div key={i} onClick={() => setSelProc(r)} style={{ background: "var(--card2)", borderRadius: 9, padding: "11px 14px", marginBottom: 7, border: `1.5px solid ${cor}40`, cursor: "pointer", display: "flex", gap: 12, alignItems: "flex-start" }}>
                  <div style={{ width: 10, height: 10, borderRadius: "50%", background: cor, flexShrink: 0, marginTop: 3 }} />
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center", marginBottom: 3 }}>
                      <span style={{ fontWeight: 700, color: "#2e86c1", fontSize: 12 }}>{r.NumRC || r.TicketSD || "—"}</span>
                      {r["Área Requisitante"] && <span style={{ fontSize: 9, background: "#2e86c118", color: "#2e86c1", borderRadius: 3, padding: "0 5px", fontWeight: 600 }}>{r["Área Requisitante"]}</span>}
                      <span style={{ fontSize: 9, color: "var(--text3)", background: "var(--card)", borderRadius: 4, padding: "0 5px" }}>{r.Modalidade || "—"}</span>
                      <span style={{ fontSize: 10, fontWeight: 700, color: cor }}>{labelEntrega(r.diasParaEntrega)}</span>
                    </div>
                    <div style={{ display: "flex", gap: 5, alignItems: "center", flexWrap: "wrap", marginBottom: 3 }}>
                      <span style={{ fontWeight: 800, fontSize: 12, color: rAtivoColor }}>◉ {r.respAtivo || "—"}</span>
                      {r.faseSubarea && <span style={{ fontSize: 8, background: `${rAtivoColor}18`, color: rAtivoColor, borderRadius: 3, padding: "0 5px", fontWeight: 600 }}>{r.faseSubarea}</span>}
                      {r.statusDet && <span style={{ fontSize: 9, color: "#8e44ad", background: "#8e44ad11", borderRadius: 3, padding: "1px 5px" }}>{r.statusDet}</span>}
                    </div>
                    <div style={{ fontSize: 11, color: "var(--text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", marginBottom: 2 }}>{(r.Objeto || "").slice(0, 100) || "—"}</div>
                    {(r.Comprador || r.cplResp || r.AnalistaContrato) && <div style={{ fontSize: 9, color: "var(--text3)" }}>
                      {[r.Comprador && `Comp: ${r.Comprador}`, r.cplResp && `CPL: ${r.cplResp}`, r.AnalistaContrato && `Scont: ${r.AnalistaContrato}`, r.AdvogadoResp && `Adv: ${r.AdvogadoResp}`].filter(Boolean).join(" · ")}
                    </div>}
                  </div>
                </div>;
              })}
            </div>}

            {/* ── 2. ENTREGAS DO MÊS CORRENTE ───────────────────────── */}
            {mesCorrente.length > 0 && (() => {
              const rAtivoColor = r => r.faseSubarea === "CPL" ? "#8e44ad" : r.faseSubarea === "Scont" ? "#16a085" : "#2e86c1";
              const entregaCard = (r, i, borderColor) => (
                <div key={i} onClick={() => setSelProc(r)} style={{ display: "flex", alignItems: "flex-start", gap: 10, padding: "10px 12px", borderRadius: 8, background: "var(--card2)", border: `1.5px solid ${borderColor}40`, cursor: "pointer", transition: "background .15s" }}>
                  <div style={{ width: 10, height: 10, borderRadius: "50%", background: borderColor, flexShrink: 0, marginTop: 3 }} />
                  <div style={{ flex: 1, minWidth: 0 }}>
                    {/* Linha 1: RC + área + modalidade */}
                    <div style={{ display: "flex", gap: 5, alignItems: "center", flexWrap: "wrap" }}>
                      <span style={{ fontWeight: 700, fontSize: 12, color: "#2e86c1" }}>{r.NumRC || r.TicketSD || "—"}</span>
                      {r["Área Requisitante"] && <span style={{ fontSize: 9, background: "#2e86c118", color: "#2e86c1", borderRadius: 3, padding: "0 5px", fontWeight: 600 }}>{r["Área Requisitante"]}</span>}
                      <span style={{ fontSize: 9, color: "var(--text2)", background: "var(--card)", borderRadius: 4, padding: "0 5px" }}>{r.Modalidade || "—"}</span>
                    </div>
                    {/* Linha 2: responsável ativo + subárea + statusDet */}
                    <div style={{ display: "flex", gap: 5, alignItems: "center", flexWrap: "wrap", marginTop: 3 }}>
                      <span style={{ fontWeight: 700, fontSize: 11, color: rAtivoColor(r) }}>◉ {r.respAtivo || "—"}</span>
                      {r.faseSubarea && <span style={{ fontSize: 8, color: "var(--text3)", background: "var(--card)", borderRadius: 3, padding: "0 4px", fontWeight: 600 }}>{r.faseSubarea}</span>}
                      {r.statusDet && <span style={{ fontSize: 9, color: "#8e44ad", background: "#8e44ad11", borderRadius: 3, padding: "1px 5px", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.statusDet}</span>}
                    </div>
                    {/* Linha 3: demais responsáveis */}
                    {(r.Comprador || r.cplResp || r.AnalistaContrato) && <div style={{ fontSize: 9, color: "var(--text3)", marginTop: 2 }}>
                      {[r.Comprador && `Comp: ${r.Comprador}`, r.cplResp && `CPL: ${r.cplResp}`, r.AnalistaContrato && `Scont: ${r.AnalistaContrato}`].filter(Boolean).join(" · ")}
                    </div>}
                    {/* Linha 4: objeto */}
                    <div style={{ fontSize: 10, color: "var(--text2)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", marginTop: 2 }}>{(r.Objeto || "").slice(0, 85) || "—"}</div>
                  </div>
                  <div style={{ flexShrink: 0, textAlign: "right" }}>
                    <div style={{ fontSize: 11, fontWeight: 800, color: borderColor }}>{labelEntrega(r.diasParaEntrega)}</div>
                    <div style={{ fontSize: 10, color: "var(--text3)" }}>{r.dataEntrega.toLocaleDateString("pt-BR")}</div>
                  </div>
                </div>
              );
              return <div style={{ background: "var(--card)", borderRadius: 12, padding: 16, marginBottom: 16, boxShadow: "var(--shadow)", border: "1.5px solid #e67e2230" }}>
                <div style={{ fontWeight: 700, fontSize: 13, color: "#e67e22", marginBottom: 10 }}>📋 Entregas deste mês — {mesCorrente.length} processo(s)</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  {mesCorrente.map((r, i) => entregaCard(r, i, corEntrega(r.diasParaEntrega)))}
                </div>
              </div>;
            })()}

            {/* ── 3. VENCIDOS — colapsável ──────────────────────────── */}
            {vencidos.length > 0 && <div style={{ background: "var(--card)", borderRadius: 12, boxShadow: "var(--shadow)", border: "1.5px solid #c0392b30", marginBottom: 16, overflow: "hidden" }}>
              <div onClick={() => setVencidosOpen(v => !v)}
                style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 16px", cursor: "pointer", userSelect: "none" }}>
                <div style={{ fontWeight: 700, fontSize: 13, color: "#c0392b" }}>
                  ⚠️ Vencidos — {vencidos.length} processo(s)
                </div>
                <div style={{ fontSize: 18, color: "#c0392b", transition: "transform .2s", transform: vencidosOpen ? "rotate(180deg)" : "rotate(0deg)" }}>▾</div>
              </div>
              {vencidosOpen && <div style={{ padding: "0 16px 16px" }}>
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  {vencidos.sort((a,b) => b.diasParaEntrega - a.diasParaEntrega).map((r, i) => {
                    const rAtivoColor = r.faseSubarea === "CPL" ? "#8e44ad" : r.faseSubarea === "Scont" ? "#16a085" : "#2e86c1";
                    return <div key={i} onClick={() => setSelProc(r)} style={{ display: "flex", alignItems: "flex-start", gap: 10, padding: "10px 12px", borderRadius: 8, background: "var(--card2)", border: "1.5px solid #c0392b40", cursor: "pointer" }}>
                      <div style={{ width: 10, height: 10, borderRadius: "50%", background: "#c0392b", flexShrink: 0, marginTop: 3 }} />
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ display: "flex", gap: 5, alignItems: "center", flexWrap: "wrap" }}>
                          <span style={{ fontWeight: 700, fontSize: 12, color: "#2e86c1" }}>{r.NumRC || r.TicketSD || "—"}</span>
                          {r["Área Requisitante"] && <span style={{ fontSize: 9, background: "#2e86c118", color: "#2e86c1", borderRadius: 3, padding: "0 5px", fontWeight: 600 }}>{r["Área Requisitante"]}</span>}
                          <span style={{ fontSize: 9, color: "var(--text2)", background: "var(--card)", borderRadius: 4, padding: "0 5px" }}>{r.Modalidade || "—"}</span>
                        </div>
                        <div style={{ display: "flex", gap: 5, alignItems: "center", flexWrap: "wrap", marginTop: 3 }}>
                          <span style={{ fontWeight: 700, fontSize: 11, color: rAtivoColor }}>◉ {r.respAtivo || "—"}</span>
                          {r.faseSubarea && <span style={{ fontSize: 8, color: "var(--text3)", background: "var(--card)", borderRadius: 3, padding: "0 4px", fontWeight: 600 }}>{r.faseSubarea}</span>}
                          {r.statusDet && <span style={{ fontSize: 9, color: "#8e44ad", background: "#8e44ad11", borderRadius: 3, padding: "1px 5px", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.statusDet}</span>}
                        </div>
                        {(r.Comprador || r.cplResp || r.AnalistaContrato) && <div style={{ fontSize: 9, color: "var(--text3)", marginTop: 2 }}>
                          {[r.Comprador && `Comp: ${r.Comprador}`, r.cplResp && `CPL: ${r.cplResp}`, r.AnalistaContrato && `Scont: ${r.AnalistaContrato}`].filter(Boolean).join(" · ")}
                        </div>}
                        <div style={{ fontSize: 10, color: "var(--text2)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", marginTop: 2 }}>{(r.Objeto || "").slice(0, 85) || "—"}</div>
                      </div>
                      <div style={{ flexShrink: 0, textAlign: "right" }}>
                        <div style={{ fontSize: 11, fontWeight: 800, color: "#c0392b" }}>{labelEntrega(r.diasParaEntrega)}</div>
                        <div style={{ fontSize: 10, color: "var(--text3)" }}>{r.dataEntrega.toLocaleDateString("pt-BR")}</div>
                      </div>
                    </div>;
                  })}
                </div>
              </div>}
            </div>}

            {/* ── 4. PRÓXIMAS ENTREGAS — todos, collapsível ─────────── */}
            {(() => {
              const futuros = procEntregas
                .filter(r => r.diasParaEntrega !== null && r.diasParaEntrega >= 0)
                .sort((a, b) => a.diasParaEntrega - b.diasParaEntrega);
              if (futuros.length === 0) return null;
              return <div style={{ background: "var(--card)", borderRadius: 12, boxShadow: "var(--shadow)", border: "1.5px solid #2e86c130", marginBottom: 16, overflow: "hidden" }}>
                <div onClick={() => setFuturosTodosOpen(v => !v)}
                  style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 16px", cursor: "pointer", userSelect: "none" }}>
                  <div style={{ fontWeight: 700, fontSize: 13, color: "#2e86c1" }}>
                    📅 Próximas Entregas — {futuros.length} processo(s)
                  </div>
                  <div style={{ fontSize: 18, color: "#2e86c1", transition: "transform .2s", transform: futurosTodosOpen ? "rotate(180deg)" : "rotate(0deg)" }}>▾</div>
                </div>
                {futurosTodosOpen && <div style={{ padding: "0 16px 16px" }}>
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    {futuros.map((r, i) => {
                      const cor = corEntrega(r.diasParaEntrega);
                      const rAtivoColor = r.faseSubarea === "CPL" ? "#8e44ad" : r.faseSubarea === "Scont" ? "#16a085" : "#2e86c1";
                      return <div key={i} onClick={() => setSelProc(r)} style={{ display: "flex", alignItems: "flex-start", gap: 10, padding: "10px 12px", borderRadius: 8, background: "var(--card2)", border: `1.5px solid ${cor}40`, cursor: "pointer", transition: "background .15s" }}
                        onMouseEnter={e => e.currentTarget.style.background = "var(--hover)"}
                        onMouseLeave={e => e.currentTarget.style.background = "var(--card2)"}>
                        <div style={{ width: 10, height: 10, borderRadius: "50%", background: cor, flexShrink: 0, marginTop: 3 }} />
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ display: "flex", gap: 5, alignItems: "center", flexWrap: "wrap" }}>
                            <span style={{ fontWeight: 700, fontSize: 12, color: "#2e86c1" }}>{r.NumRC || r.TicketSD || "—"}</span>
                            {r["Área Requisitante"] && <span style={{ fontSize: 9, background: "#2e86c118", color: "#2e86c1", borderRadius: 3, padding: "0 5px", fontWeight: 600 }}>{r["Área Requisitante"]}</span>}
                            <span style={{ fontSize: 9, color: "var(--text2)", background: "var(--card)", borderRadius: 4, padding: "0 5px" }}>{r.Modalidade || "—"}</span>
                          </div>
                          <div style={{ display: "flex", gap: 5, alignItems: "center", flexWrap: "wrap", marginTop: 3 }}>
                            <span style={{ fontWeight: 700, fontSize: 11, color: rAtivoColor }}>◉ {r.respAtivo || "—"}</span>
                            {r.faseSubarea && <span style={{ fontSize: 8, color: "var(--text3)", background: "var(--card)", borderRadius: 3, padding: "0 4px", fontWeight: 600 }}>{r.faseSubarea}</span>}
                            {r.statusDet && <span style={{ fontSize: 9, color: "#8e44ad", background: "#8e44ad11", borderRadius: 3, padding: "1px 5px", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.statusDet}</span>}
                          </div>
                          {(r.Comprador || r.cplResp || r.AnalistaContrato) && <div style={{ fontSize: 9, color: "var(--text3)", marginTop: 2 }}>
                            {[r.Comprador && `Comp: ${r.Comprador}`, r.cplResp && `CPL: ${r.cplResp}`, r.AnalistaContrato && `Scont: ${r.AnalistaContrato}`].filter(Boolean).join(" · ")}
                          </div>}
                          <div style={{ fontSize: 10, color: "var(--text2)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", marginTop: 2 }}>{(r.Objeto || "").slice(0, 85) || "—"}</div>
                        </div>
                        <div style={{ flexShrink: 0, textAlign: "right" }}>
                          <div style={{ fontSize: 11, fontWeight: 800, color: cor }}>{labelEntrega(r.diasParaEntrega)}</div>
                          <div style={{ fontSize: 10, color: "var(--text3)" }}>{r.dataEntrega.toLocaleDateString("pt-BR")}</div>
                        </div>
                      </div>;
                    })}
                  </div>
                </div>}
              </div>;
            })()}

            {procEntregas.length === 0 && <div style={{ textAlign: "center", padding: 60, color: "var(--text3)" }}>
              <div style={{ fontSize: 38 }}>📭</div>
              <div style={{ fontWeight: 600, marginTop: 8 }}>Nenhum processo em andamento com data de entrega cadastrada.</div>
              <div style={{ fontSize: 12, marginTop: 6 }}>Preencha a coluna "PREVISÃO ENTREGA / INÍCIO PRESTAÇÃO DO SERVIÇO DA RC" na planilha.</div>
            </div>}
          </div>;
        })());
}
