import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function DirTimelineScreen({ base, dirTLDrill, phaseIntervals, setDirTLDrill, setSelProc }) {
  return ((() => {
          // Processos em andamento COM RC (aging oficial a partir da RC)
          const tlBack = base.filter(r => r.emA && r.NumRC && r.NumRC.trim() !== "");
          // Agrupar por modalidade normalizada
          const modMap = new Map();
          tlBack.forEach(r => {
            const mn = nrm(r.Modalidade || "N/I");
            if (!modMap.has(mn)) modMap.set(mn, { label: r.Modalidade || "N/I", procs: [] });
            modMap.get(mn).procs.push(r);
          });
          const modRows = [...modMap.values()].sort((a, b) => b.procs.length - a.procs.length);

          // Cores por modalidade
          const MOD_COLORS = { "pregao": "#1a5276", "concorrencia": "#6c3483", "dispensa": "#e67e22", "inexigibilidade": "#c0392b", "adesao": "#16a085", "credenciamento": "#2e86c1", "dialogo": "#8e44ad" };
          const getModColor = (mn) => {
            for (const [k, v] of Object.entries(MOD_COLORS)) { if (mn.includes(k)) return v; }
            return "#34495e";
          };

          // Função para abrir drill-down (com contexto de fase opcional)
          const openDrill = (title, procs, phaseCtx) => { if (procs.length > 0) setDirTLDrill({ title, procs, phaseCtx: phaseCtx || null }); };

          return (
            <div className="anim-fade">
              {/* Header */}
              <div style={{ marginBottom: 18 }}>
                <div className="gaq-h1" style={{ marginBottom: 4 }}>Timeline por Modalidade</div>
                <div className="gaq-body" style={{ color: "var(--gaq-text-3)" }}>Visao consolidada das etapas por modalidade · {tlBack.length} processos em andamento com RC</div>
              </div>

              {/* KPIs resumo */}
              <div className="gaq-kpi-grid" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", marginBottom: 18 }}>
                <div className="gaq-kpi" style={{ borderTop: "3px solid #2e86c1", cursor: "pointer" }} onClick={() => openDrill("Todos os Processos em Andamento", tlBack)}>
                  <span className="lbl">Total em Andamento</span>
                  <div className="val" style={{ color: "#2e86c1" }}>{tlBack.length}</div>
                </div>
                <div className="gaq-kpi" style={{ borderTop: "3px solid #27ae60", cursor: "pointer" }} onClick={() => openDrill("Processos com SLA Atendido", tlBack.filter(r => !r.atrasoGeral))}>
                  <span className="lbl">SLA Atendido</span>
                  <div className="val" style={{ color: "#27ae60" }}>{tlBack.filter(r => !r.atrasoGeral).length}</div>
                </div>
                <div className="gaq-kpi" style={{ borderTop: "3px solid #e74c3c", cursor: "pointer" }} onClick={() => openDrill("Todos os Processos com SLA Vencido", tlBack.filter(r => r.atrasoGeral))}>
                  <span className="lbl">SLA Vencido</span>
                  <div className="val" style={{ color: "#e74c3c" }}>{tlBack.filter(r => r.atrasoGeral).length}</div>
                </div>
                <div className="gaq-kpi" style={{ borderTop: "3px solid #f39c12" }}>
                  <span className="lbl">Modalidades</span>
                  <div className="val" style={{ color: "#f39c12" }}>{modRows.length}</div>
                </div>
              </div>

              {/* Timeline por modalidade */}
              {modRows.map(({ label: modLabel, procs: modProcs }) => {
                const mn = nrm(modLabel);
                const modColor = getModColor(mn);
                const prazo = getPrazoGeral(mn);
                const colIdxs = getTLColsForMod(mn);
                const atrasados = modProcs.filter(r => r.atrasoGeral);

                // SLA por fase: dinâmico (se disponível) ou estático
                const dynamicSLA = buildDynamicSLA(phaseIntervals, mn);
                const isCPLMod = mn.includes("pregao") || mn.includes("concorrencia") || mn.includes("dialogo");
                const baseSLA = isCPLMod ? SLA_CPL : SLA_SIMPLES;
                const faseSLA = dynamicSLA || baseSLA;
                const hoje = new Date();

                // 1) Determinar posição de cada processo pela ÚLTIMA DATA PREENCHIDA.
                // Ignora fases voláteis (ex.: Indicação Analista) — sua data não representa
                // o avanço do processo no trilho, pois pode acontecer fora de sequência.
                const procLastIdx = new Map();
                modProcs.forEach(r => {
                  let lastIdx = -1;
                  for (let ci = colIdxs.length - 1; ci >= 0; ci--) {
                    const cIdx = colIdxs[ci];
                    const k = TL_COLS[cIdx][1];
                    if (VOLATILE_TIMELINE_KEYS.has(k)) continue;
                    if (pd(r[k])) { lastIdx = cIdx; break; }
                  }
                  // Se nenhuma data preenchida, colocar na primeira fase
                  if (lastIdx < 0) lastIdx = colIdxs[0];
                  procLastIdx.set(r, lastIdx);
                });

                // 2) Para cada etapa, agrupar processos cuja última data preenchida = esse idx
                // SLA por fase: dias desde a última data preenchida vs SLA da PRÓXIMA fase
                const phaseData = colIdxs.map((idx, posInList) => {
                  const [phaseLabel, phaseKey] = TL_COLS[idx];
                  const procsAtPhase = modProcs.filter(r => procLastIdx.get(r) === idx);

                  // Próxima fase no fluxo (para saber o SLA de transição)
                  const nextPos = posInList + 1;
                  const nextIdx = nextPos < colIdxs.length ? colIdxs[nextPos] : -1;
                  const nextPhaseKey = nextIdx >= 0 ? TL_COLS[nextIdx][1] : null;
                  let nextSLA = null;
                  if (nextPhaseKey) {
                    if (SLA_SD.hasOwnProperty(nextPhaseKey)) nextSLA = SLA_SD[nextPhaseKey];
                    else if (faseSLA.hasOwnProperty(nextPhaseKey)) nextSLA = faseSLA[nextPhaseKey];
                  }

                  // Verificar atraso por fase: dias desde a última data preenchida vs SLA da próxima fase
                  const procsComDelay = procsAtPhase.map(r => {
                    const lastDate = pd(r[phaseKey]);
                    const diasNaFase = lastDate ? du(lastDate, hoje) : 0;
                    const atrasoFase = (nextSLA != null && nextSLA > 0 && diasNaFase > nextSLA);
                    return { proc: r, diasNaFase, slaFase: nextSLA, atrasoFase };
                  });
                  const foraPrazo = procsComDelay.filter(x => x.atrasoFase).length;
                  const noPrazo = procsAtPhase.length - foraPrazo;

                  return { idx, phaseLabel, phaseKey, total: procsAtPhase.length, noPrazo, foraPrazo, procs: procsAtPhase, procsComDelay, nextSLA, nextPhaseKey };
                });

                return (
                  <div key={mn} style={{ background: "var(--card)", borderRadius: 14, marginBottom: 18, boxShadow: "var(--shadow)", overflow: "hidden", border: `1px solid ${modColor}22` }}>
                    {/* Header da modalidade */}
                    <div style={{ background: `linear-gradient(135deg, ${modColor} 0%, ${modColor}cc 100%)`, padding: "14px 20px", color: "#fff", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div>
                        <div style={{ fontSize: 16, fontWeight: 800 }}>{modLabel}</div>
                        <div style={{ fontSize: 11, opacity: .8 }}>Prazo: {prazo} d.u. (a partir do Recebimento RC) · {modProcs.length} processo(s) em andamento</div>
                      </div>
                      <div style={{ display: "flex", gap: 10 }}>
                        <div onClick={() => openDrill(`Todos — ${modLabel}`, modProcs)}
                          style={{ background: "rgba(255,255,255,0.2)", borderRadius: 8, padding: "6px 12px", textAlign: "center", cursor: "pointer" }}>
                          <div style={{ fontSize: 18, fontWeight: 800 }}>{modProcs.length}</div>
                          <div style={{ fontSize: 8, opacity: .8 }}>total</div>
                        </div>
                        <div onClick={() => openDrill(`SLA Atendido — ${modLabel}`, modProcs.filter(r => !r.atrasoGeral))}
                          style={{ background: "rgba(39,174,96,0.3)", borderRadius: 8, padding: "6px 12px", textAlign: "center", cursor: "pointer" }}>
                          <div style={{ fontSize: 18, fontWeight: 800 }}>{modProcs.length - atrasados.length}</div>
                          <div style={{ fontSize: 8, opacity: .9 }}>SLA atendido</div>
                        </div>
                        <div onClick={() => openDrill(`SLA Vencido — ${modLabel}`, atrasados)}
                          style={{ background: atrasados.length > 0 ? "rgba(231,76,60,0.4)" : "rgba(255,255,255,0.15)", borderRadius: 8, padding: "6px 12px", textAlign: "center", cursor: atrasados.length > 0 ? "pointer" : "default", border: atrasados.length > 0 ? "1px solid rgba(231,76,60,0.6)" : "none" }}>
                          <div style={{ fontSize: 18, fontWeight: 800 }}>{atrasados.length}</div>
                          <div style={{ fontSize: 8, opacity: .9 }}>atraso</div>
                        </div>
                      </div>
                    </div>

                    {/* Linha do timeline — scroll horizontal */}
                    <div style={{ overflowX: "auto", padding: "16px 12px 12px" }}>
                      <div style={{ display: "flex", alignItems: "flex-start", gap: 0, minWidth: "max-content" }}>
                        {phaseData.map((ph, pi) => {
                          const isLast = pi === phaseData.length - 1;
                          const hasProcs = ph.total > 0;
                          const bgColor = ph.foraPrazo > 0 ? "#fdecea" : ph.total > 0 ? "#eafaf1" : "var(--card2)";
                          const borderColor = ph.foraPrazo > 0 ? "#e74c3c" : ph.total > 0 ? "#27ae60" : "var(--border2)";
                          return (
                            <React.Fragment key={ph.idx}>
                              {/* Etapa */}
                              <div onClick={() => openDrill(`${modLabel} — ${ph.phaseLabel}`, ph.procs, { procsComDelay: ph.procsComDelay, phaseLabel: ph.phaseLabel, nextSLA: ph.nextSLA })}
                                style={{ display: "flex", flexDirection: "column", alignItems: "center", minWidth: 90, maxWidth: 110, cursor: hasProcs ? "pointer" : "default", transition: "transform .15s" }}
                                onMouseEnter={e => { if (hasProcs) e.currentTarget.style.transform = "scale(1.05)"; }}
                                onMouseLeave={e => e.currentTarget.style.transform = "scale(1)"}>
                                {/* Bolha com contadores */}
                                <div style={{ background: bgColor, border: `2px solid ${borderColor}`, borderRadius: 12, padding: "8px 6px", textAlign: "center", width: "100%", minHeight: 60, display: "flex", flexDirection: "column", justifyContent: "center", transition: "box-shadow .15s", boxShadow: hasProcs ? "0 2px 8px rgba(0,0,0,0.08)" : "none" }}>
                                  <div style={{ fontSize: 20, fontWeight: 800, color: ph.total > 0 ? (ph.foraPrazo > 0 ? "#c0392b" : "#1a5276") : "var(--text3)" }}>{ph.total}</div>
                                  {ph.total > 0 && (
                                    <div style={{ display: "flex", justifyContent: "center", gap: 6, marginTop: 3 }}>
                                      <span style={{ fontSize: 9, color: "#27ae60", fontWeight: 700 }}>{ph.noPrazo} ok</span>
                                      {ph.foraPrazo > 0 && <span style={{ fontSize: 9, color: "#e74c3c", fontWeight: 700 }}>{ph.foraPrazo} atr</span>}
                                    </div>
                                  )}
                                </div>
                                {/* Nome da etapa */}
                                <div style={{ fontSize: 9, color: "var(--text2)", textAlign: "center", marginTop: 5, lineHeight: 1.2, fontWeight: 600, maxWidth: 100, wordBreak: "break-word" }}>{ph.phaseLabel}</div>
                              </div>
                              {/* Conector → */}
                              {!isLast && (
                                <div style={{ display: "flex", alignItems: "center", paddingTop: 28, minWidth: 20 }}>
                                  <div style={{ height: 2, width: 20, background: `linear-gradient(90deg, ${modColor}66, ${modColor}33)` }} />
                                </div>
                              )}
                            </React.Fragment>
                          );
                        })}

                        {/* Coluna final: Atraso Geral */}
                        <div style={{ display: "flex", alignItems: "center", paddingTop: 28, minWidth: 20 }}>
                          <div style={{ height: 2, width: 24, background: `linear-gradient(90deg, ${modColor}33, #e74c3c66)` }} />
                        </div>
                        <div onClick={() => openDrill(`Atraso Geral — ${modLabel} (>${prazo} d.u.)`, atrasados)}
                          style={{ display: "flex", flexDirection: "column", alignItems: "center", minWidth: 100, cursor: atrasados.length > 0 ? "pointer" : "default", transition: "transform .15s" }}
                          onMouseEnter={e => { if (atrasados.length > 0) e.currentTarget.style.transform = "scale(1.05)"; }}
                          onMouseLeave={e => e.currentTarget.style.transform = "scale(1)"}>
                          <div style={{ background: atrasados.length > 0 ? "#fdecea" : "#eafaf1", border: `2px solid ${atrasados.length > 0 ? "#e74c3c" : "#27ae60"}`, borderRadius: 12, padding: "8px 10px", textAlign: "center", width: "100%", minHeight: 60, display: "flex", flexDirection: "column", justifyContent: "center", boxShadow: atrasados.length > 0 ? "0 2px 12px rgba(231,76,60,0.15)" : "none" }}>
                            <div style={{ fontSize: 22, fontWeight: 800, color: atrasados.length > 0 ? "#c0392b" : "#27ae60" }}>{atrasados.length}</div>
                            <div style={{ fontSize: 9, fontWeight: 700, color: atrasados.length > 0 ? "#c0392b" : "#27ae60" }}>
                              {atrasados.length > 0 ? `>${prazo} d.u.` : "Todos ok"}
                            </div>
                          </div>
                          <div style={{ fontSize: 9, color: "#e74c3c", textAlign: "center", marginTop: 5, fontWeight: 700 }}>Atraso Geral</div>
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })}

              {/* Modal Drill-down: lista de processos */}
              {dirTLDrill && <>
                <div onClick={() => setDirTLDrill(null)} style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.5)", zIndex: 9998 }} />
                <div style={{ position: "fixed", top: "5vh", left: "5vw", right: "5vw", bottom: "5vh", background: "var(--bg, #fff)", borderRadius: 16, boxShadow: "0 12px 48px rgba(0,0,0,0.3)", zIndex: 9999, display: "flex", flexDirection: "column", overflow: "hidden" }}>
                  {/* Header do modal */}
                  <div style={{ padding: "16px 24px", borderBottom: "2px solid var(--border2)", display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0, background: "var(--card)" }}>
                    <div>
                      <div style={{ fontWeight: 800, fontSize: 16, color: "#1a5276" }}>{dirTLDrill.title}</div>
                      <div style={{ fontSize: 12, color: "var(--text3)", marginTop: 2 }}>{dirTLDrill.procs.length} processo(s) · Clique em uma linha para ver detalhes</div>
                    </div>
                    <button onClick={() => setDirTLDrill(null)} style={{ background: "#e74c3c15", border: "1px solid #e74c3c33", color: "#e74c3c", cursor: "pointer", fontSize: 14, fontWeight: 700, borderRadius: 8, padding: "6px 16px", transition: "background .15s" }}
                      onMouseEnter={e => e.currentTarget.style.background = "#e74c3c25"}
                      onMouseLeave={e => e.currentTarget.style.background = "#e74c3c15"}>Fechar</button>
                  </div>
                  {/* Corpo com scroll */}
                  <div style={{ flex: 1, overflowY: "auto", padding: "0 4px" }}>
                    {dirTLDrill.procs.length === 0
                      ? <div style={{ textAlign: "center", padding: 40, color: "var(--text3)", fontSize: 14 }}>Nenhum processo nesta seleção.</div>
                      : (() => {
                        const hasPhaseCtx = !!dirTLDrill.phaseCtx;
                        const delayMap = new Map();
                        if (hasPhaseCtx && dirTLDrill.phaseCtx.procsComDelay) {
                          dirTLDrill.phaseCtx.procsComDelay.forEach(x => delayMap.set(x.proc, x));
                        }
                        const sortedProcs = hasPhaseCtx
                          ? [...dirTLDrill.procs].sort((a, b) => {
                              const da = delayMap.get(a), db = delayMap.get(b);
                              // Atrasados na fase primeiro, depois por dias na fase desc
                              if (da && db) { if (da.atrasoFase !== db.atrasoFase) return da.atrasoFase ? -1 : 1; return (db.diasNaFase || 0) - (da.diasNaFase || 0); }
                              return (b.diasTotais || 0) - (a.diasTotais || 0);
                            })
                          : [...dirTLDrill.procs].sort((a, b) => (b.diasTotais || 0) - (a.diasTotais || 0));
                        const slaFaseLabel = hasPhaseCtx && dirTLDrill.phaseCtx.nextSLA ? `${dirTLDrill.phaseCtx.nextSLA} d.u.` : "—";
                        const totalAtrasoFase = hasPhaseCtx ? [...delayMap.values()].filter(x => x.atrasoFase).length : 0;
                        const totalOkFase = hasPhaseCtx ? dirTLDrill.procs.length - totalAtrasoFase : 0;
                        return <>
                          {/* Resumo por fase (só quando veio de uma bolha de fase) */}
                          {hasPhaseCtx && (
                            <div style={{ display: "flex", gap: 12, padding: "12px 16px", borderBottom: "1px solid var(--border2)", background: "var(--card2)", flexWrap: "wrap", alignItems: "center" }}>
                              <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text2)" }}>SLA da etapa: <span style={{ color: "#1a5276" }}>{slaFaseLabel}</span></div>
                              <div style={{ fontSize: 11 }}><span style={{ background: "#27ae6022", color: "#27ae60", borderRadius: 4, padding: "2px 8px", fontWeight: 700 }}>{totalOkFase} no prazo da etapa</span></div>
                              <div style={{ fontSize: 11 }}><span style={{ background: totalAtrasoFase > 0 ? "#e74c3c" : "#27ae6022", color: totalAtrasoFase > 0 ? "#fff" : "#27ae60", borderRadius: 4, padding: "2px 8px", fontWeight: 700 }}>{totalAtrasoFase} em atraso na etapa</span></div>
                            </div>
                          )}
                          <div className="gaq-table-scroll"><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                            <thead>
                              <tr style={{ background: "var(--card2)", borderBottom: "2px solid var(--border2)", position: "sticky", top: 0, zIndex: 1 }}>
                                <th style={{ padding: "10px 8px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase" }}>Nº RC</th>
                                <th style={{ padding: "10px 8px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase" }}>Ticket pré-compra</th>
                                <th style={{ padding: "10px 8px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", maxWidth: 220 }}>Objeto</th>
                                <th style={{ padding: "10px 8px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase" }}>Área</th>
                                <th style={{ padding: "10px 8px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase" }}>Comprador</th>
                                {hasPhaseCtx && <>
                                  <th style={{ padding: "10px 8px", textAlign: "center", fontSize: 10, fontWeight: 700, color: "#8e44ad", textTransform: "uppercase" }}>Dias na Fase</th>
                                  <th style={{ padding: "10px 8px", textAlign: "center", fontSize: 10, fontWeight: 700, color: "#8e44ad", textTransform: "uppercase" }}>SLA Fase</th>
                                  <th style={{ padding: "10px 8px", textAlign: "center", fontSize: 10, fontWeight: 700, color: "#8e44ad", textTransform: "uppercase" }}>Status Fase</th>
                                </>}
                                <th style={{ padding: "10px 8px", textAlign: "center", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase" }}>Dias (RC)</th>
                                <th style={{ padding: "10px 8px", textAlign: "center", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase" }}>Prazo Geral</th>
                                <th style={{ padding: "10px 8px", textAlign: "center", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase" }}>Status Geral</th>
                              </tr>
                            </thead>
                            <tbody>
                              {sortedProcs.map((r, i) => {
                                const late = r.atrasoGeral;
                                const cor = late ? "#c0392b" : (r.diasTotais || 0) > r.prazoGeral * 0.7 ? "#e67e22" : "#27ae60";
                                const phDelay = delayMap.get(r);
                                return (
                                  <tr key={r.ProcessKey || i} onClick={() => { setDirTLDrill(null); setSelProc(r); }} style={{ cursor: "pointer", borderBottom: "1px solid var(--border2)", transition: "background .15s", background: phDelay && phDelay.atrasoFase ? "#fdecea44" : "transparent" }}
                                    onMouseEnter={e => e.currentTarget.style.background = "var(--hover)"}
                                    onMouseLeave={e => e.currentTarget.style.background = phDelay && phDelay.atrasoFase ? "#fdecea44" : "transparent"}>
                                    <td style={{ padding: "9px 8px", fontWeight: 700, color: "#2e86c1" }}>{r.NumRC || "—"}</td>
                                    <td style={{ padding: "9px 8px", color: "#27ae60", fontSize: 11 }}>{r.TicketSD || "—"}</td>
                                    <td style={{ padding: "9px 8px", color: "var(--text)", maxWidth: 220, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={r.Objeto}>{(r.Objeto || "").slice(0, 60) || "—"}</td>
                                    <td style={{ padding: "9px 8px", color: "var(--text2)" }}>{r["Área Requisitante"] || "—"}</td>
                                    <td style={{ padding: "9px 8px", color: "var(--text2)" }}>{r.Comprador || r.respNCL || "—"}</td>
                                    {hasPhaseCtx && phDelay && <>
                                      <td style={{ padding: "9px 8px", textAlign: "center", fontWeight: 800, color: phDelay.atrasoFase ? "#c0392b" : "#1a5276" }}>{phDelay.diasNaFase}</td>
                                      <td style={{ padding: "9px 8px", textAlign: "center", fontSize: 11, color: "#8e44ad" }}>{phDelay.slaFase != null ? `${phDelay.slaFase} d.u.` : "—"}</td>
                                      <td style={{ padding: "9px 8px", textAlign: "center" }}>
                                        {phDelay.atrasoFase
                                          ? <span style={{ background: "#e74c3c", color: "#fff", borderRadius: 4, padding: "2px 8px", fontSize: 10, fontWeight: 700 }}>ATRASO</span>
                                          : <span style={{ background: "#27ae6022", color: "#27ae60", borderRadius: 4, padding: "2px 8px", fontSize: 10, fontWeight: 600 }}>Ok</span>}
                                      </td>
                                    </>}
                                    {hasPhaseCtx && !phDelay && <>
                                      <td style={{ padding: "9px 8px", textAlign: "center", color: "var(--text3)" }}>—</td>
                                      <td style={{ padding: "9px 8px", textAlign: "center", color: "var(--text3)" }}>—</td>
                                      <td style={{ padding: "9px 8px", textAlign: "center", color: "var(--text3)" }}>—</td>
                                    </>}
                                    <td style={{ padding: "9px 8px", textAlign: "center", fontWeight: 800, color: cor }}>{r.diasTotais || 0}</td>
                                    <td style={{ padding: "9px 8px", textAlign: "center", fontSize: 11, color: "var(--text3)" }}>{r.prazoGeral} d.u.</td>
                                    <td style={{ padding: "9px 8px", textAlign: "center" }}>
                                      {late
                                        ? <span style={{ background: "#e74c3c", color: "#fff", borderRadius: 4, padding: "2px 8px", fontSize: 10, fontWeight: 700 }}>ATRASO</span>
                                        : <span style={{ background: "#27ae6022", color: "#27ae60", borderRadius: 4, padding: "2px 8px", fontSize: 10, fontWeight: 600 }}>No prazo</span>}
                                    </td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table></div>
                        </>;
                      })()}
                  </div>
                </div>
              </>}

              {/* Legenda */}
              <div style={{ background: "var(--card)", borderRadius: 12, padding: "14px 20px", boxShadow: "var(--shadow)", marginTop: 10 }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text2)", marginBottom: 8 }}>Legenda</div>
                <div style={{ display: "flex", gap: 20, flexWrap: "wrap", fontSize: 11, color: "var(--text2)" }}>
                  <span><span style={{ display: "inline-block", width: 12, height: 12, background: "#eafaf1", border: "2px solid #27ae60", borderRadius: 4, marginRight: 4, verticalAlign: "middle" }} /> No prazo</span>
                  <span><span style={{ display: "inline-block", width: 12, height: 12, background: "#fdecea", border: "2px solid #e74c3c", borderRadius: 4, marginRight: 4, verticalAlign: "middle" }} /> Fora do prazo</span>
                  <span style={{ color: "var(--text3)" }}>Licitações (Pregão, Concorrência): <b>{PRAZO_LICITACAO} d.u.</b> a partir do Recebimento RC</span>
                  <span style={{ color: "var(--text3)" }}>Compras Diretas (Dispensa, Inexig., Adesão ARP, Credenc.): <b>{PRAZO_COMPRA_DIRETA} d.u.</b> a partir do Recebimento RC</span>
                  <span style={{ color: "var(--text3)" }}>Pré-compra (Abertura → Encerramento): <b>{PRAZO_SD} d.u.</b> (prazo apartado)</span>
                </div>
              </div>
            </div>
          );
        })());
}
