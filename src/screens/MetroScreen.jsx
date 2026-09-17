import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function MetroScreen({ base, fBase, metroAgrupado, metroBusca, metroComp, pgMetro, pgMetroSz, phaseIntervals, setAba, setMetroAgrupado, setMetroBusca, setMetroComp, setPgMetro, setPgMetroSz, setSelProc }) {
  return (<div className="anim-fade" style={{ minHeight: "max-content", overflow: "visible", paddingBottom: 40 }}>
          {base.length === 0 ? (
            <div style={{ textAlign: "center", padding: 60, color: "var(--text3)" }}>
              <div style={{ fontSize: 38 }}>🚇</div>
              <div style={{ fontWeight: 600, marginTop: 8 }}>Carregue a base para acessar a Linha de Metro.</div>
              <button onClick={() => setAba("upload")} style={{ marginTop: 12, background: "#1a5276", color: "#fff", border: "none", borderRadius: 8, padding: "9px 22px", cursor: "pointer", fontSize: 13 }}>Upload</button>
            </div>
          ) : (() => {
            const stColor = r => r.emA ? "#1e8449" : r.isConcluded ? "#2e86c1" : r.isCanceled ? "#e67e22" : r.isFailed ? "#c0392b" : r.isSuspended ? "#7f8c8d" : "#95a5a6";
            const getResp = r => (r.respAtivo && r.respAtivo !== "N/A") ? r.respAtivo : (r.respNCL || "Sem responsável");
            const compList = [...new Set(fBase.map(getResp).filter(Boolean))].sort();
            const metroBuscaNrm = nrm(metroBusca);
            const filteredMetro = fBase.filter(r => {
              if (metroComp && getResp(r) !== metroComp) return false;
              if (metroBuscaNrm.length >= 2 &&
                  !nrm(r.Objeto || "").includes(metroBuscaNrm) &&
                  !nrm(r.TicketSD || "").includes(metroBuscaNrm) &&
                  !nrm(r.NumRC || "").includes(metroBuscaNrm) &&
                  !nrm(r.NumPedidoSuite || "").includes(metroBuscaNrm) &&
                  !nrm(r.NumProcesso || "").includes(metroBuscaNrm)) return false;
              return true;
            });
            const METRO_COLORS = ["#1565c0","#1e8449","#8e44ad","#c0392b","#e67e22","#16a085","#b7950b","#2c3e50","#117a65","#6c3483"];

            // Paginação — aplicada ANTES de montar os grupos
            const metroTotal = filteredMetro.length;
            const metroPage = Math.min(pgMetro, Math.ceil(metroTotal / pgMetroSz) || 1);
            const metroSlice = filteredMetro.slice((metroPage - 1) * pgMetroSz, metroPage * pgMetroSz);
            const forwardMetroWheel = (e) => {
              if (Math.abs(e.deltaY) <= Math.abs(e.deltaX)) return;
              const scroller = e.currentTarget.closest("[data-metro-scroll='true']");
              if (!scroller) return;
              scroller.scrollTop += e.deltaY;
              e.preventDefault();
            };

            const renderStation = (r, i, total, lineColor) => {
              const sc = stColor(r);
              const resp = getResp(r);
              const cron = buildCronograma(r, phaseIntervals);

              // Mini SLA track — versão aprimorada para gestores
              const SHORT_LBL = {
                "Abertura SD": "Aber. pré-compra", "Distribuição SD": "Dist. pré-compra",
                "Encerramento SD": "Enc. pré-compra", "Recebimento RC": "Receb. RC",
                "Planejamento RC": "Plan. RC", "Início propostas": "Prop. Iníc.",
                "Fim propostas": "Prop. Fim", "Envio aprovação": "Env. Apr.",
                "Última aprovação": "Últ. Apr.", "Envio Pedido/Suite": "Env. Pedido",
                "CPL: Recebido do NCL": "CPL NCL",
                "CPL: Enviado p/ DJS": "CPL DJS",
                "CPL: Recebimento": "CPL Receb.", "CPL: Publicação": "CPL Pub.",
                "CPL: Abertura Disputa": "CPL Disp.", "CPL: Fase ext.": "CPL F.Ext.",
                "CPL: Homologação": "CPL Homol.", "Indicação Analista de Contrato": "Ind. Analista",
                "Recebimento DJ": "Receb. DJ",
              };

              const miniTrack = cron ? (() => {
                const rows = cron.rows;
                const nOk   = rows.filter(x => x.status === "No prazo").length;
                const nLate = rows.filter(x => x.status === "Atrasado").length;
                const nLogged = rows.filter(x => x.status === "Registrado").length;
                const nPend = rows.filter(x => x.status === "Pendente").length;
                const nDone = rows.filter(x => x._ok).length;

                // current stage = first Pendente after last completed
                let currentIdx = -1;
                for (let ci = rows.length - 1; ci >= 0; ci--) {
                  if (rows[ci]._ok) { currentIdx = ci + 1; break; }
                }
                if (currentIdx < 0 && rows.some(x => x.status === "Pendente")) currentIdx = 0;

                // SLA bar: percent of total SLA dias used
                const pct = Math.min(100, Math.round(((r.diasTotais || r.diasTotaisCronograma) / cron.metaTotal) * 100));
                const barColor = pct > 120 ? "#c0392b" : pct > 80 ? "#e67e22" : "#1e8449";

                // Determine step icon + colors
                const stepInfo = rows.map((row, ri) => {
                  const isCurrent = ri === currentIdx && r.emA;
                  let bg, border, icon, textColor;
                  if (row._ok) {
                    bg = row.status === "Atrasado" ? "#c0392b" : row.status === "Registrado" ? "#8e44ad" : "#1e8449";
                    border = bg;
                    icon = row.status === "Atrasado" ? "!" : row.status === "Registrado" ? "R" : "✓";
                    textColor = "#fff";
                  } else if (isCurrent) {
                    bg = "#fff7ed";
                    border = "#e67e22";
                    icon = String(ri + 1);
                    textColor = "#e67e22";
                  } else {
                    bg = "var(--card2)";
                    border = "var(--border)";
                    icon = String(ri + 1);
                    textColor = "var(--text3)";
                  }
                  return { ...row, isCurrent, bg, border, icon, textColor };
                });

                return (
                  <div style={{ marginBottom: 10 }}>
                    {/* SLA progress bar */}
                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                      <div style={{ fontSize: 9, color: "var(--text3)", fontWeight: 700, whiteSpace: "nowrap" }}>SLA {cron.metaTotal} d.u.</div>
                      <div style={{ flex: 1, height: 6, background: "var(--border2)", borderRadius: 4, overflow: "hidden" }}>
                        <div style={{ height: "100%", width: pct + "%", background: barColor, borderRadius: 4, transition: "width .3s" }} />
                      </div>
                      <div style={{ fontSize: 9, fontWeight: 800, color: barColor, whiteSpace: "nowrap" }}>{r.diasTotaisCronograma} / {cron.metaTotal} d.u. ({pct}%)</div>
                    </div>

                    {/* Steps track — scrollable horizontally */}
                    <div onWheel={forwardMetroWheel} style={{ overflowX: "auto", overflowY: "hidden", overscrollBehaviorX: "contain", paddingBottom: 4 }}>
                      <div style={{ display: "flex", alignItems: "flex-start", gap: 0, minWidth: "max-content" }}>
                        {stepInfo.map((step, ri) => (
                          <React.Fragment key={ri}>
                            {/* Connector line */}
                            {ri > 0 && (
                              <div style={{
                                width: 18, height: 3, marginTop: 18, flexShrink: 0,
                                background: (stepInfo[ri-1]._ok && step._ok)
                                  ? (stepInfo[ri-1].status === "Atrasado" || step.status === "Atrasado" ? "#c0392b" : stepInfo[ri-1].status === "Registrado" || step.status === "Registrado" ? "#8e44ad" : "#1e8449")
                                  : step.isCurrent ? "#e67e2255" : "var(--border2)"
                              }} />
                            )}
                            {/* Step column */}
                            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", width: 42 }}
                              title={`${step.etapa}\nResponsável: ${step.responsavel}\nPlanejado: ${step.planejada}\nReal: ${step.real}\nStatus: ${step.status}${step.variacao !== "—" ? "\nVariação: " + step.variacao : ""}`}>
                              {/* Circle */}
                              <div style={{
                                width: 38, height: 38, borderRadius: "50%",
                                background: step.bg, border: `2.5px solid ${step.border}`,
                                display: "flex", alignItems: "center", justifyContent: "center",
                                fontSize: step._ok ? 14 : 11, fontWeight: 800, color: step.textColor,
                                flexShrink: 0,
                                boxShadow: step.isCurrent ? `0 0 0 3px #e67e2233, 0 2px 8px #e67e2244` : step._ok ? `0 1px 4px ${step.bg}55` : "none",
                                transition: "box-shadow .2s"
                              }}>
                                {step.icon}
                              </div>
                              {/* Label */}
                              <div style={{
                                fontSize: 8, fontWeight: step.isCurrent ? 800 : 500,
                                color: step.isCurrent ? "#e67e22" : step._ok ? (step.status === "Atrasado" ? "#c0392b" : step.status === "Registrado" ? "#8e44ad" : "#1e8449") : "var(--text3)",
                                textAlign: "center", marginTop: 4, lineHeight: 1.2,
                                maxWidth: 42, wordBreak: "break-word"
                              }}>
                                {SHORT_LBL[step.etapa] || step.etapa}
                              </div>
                              {/* Variação badge (só se atrasado ou no prazo) */}
                              {step._ok && step.variacao !== "—" && (
                                <div style={{ fontSize: 7, fontWeight: 700, color: step.variacao.startsWith("+") ? "#c0392b" : "#1e8449", marginTop: 2 }}>
                                  {step.variacao}
                                </div>
                              )}
                              {/* "Atual" badge */}
                              {step.isCurrent && (
                                <div style={{ fontSize: 7, fontWeight: 800, color: "#e67e22", marginTop: 2, background: "#fff7ed", borderRadius: 3, padding: "1px 4px", border: "1px solid #e67e2244" }}>
                                  ATUAL
                                </div>
                              )}
                            </div>
                          </React.Fragment>
                        ))}
                      </div>
                    </div>

                    {/* Summary row */}
                    <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 8, paddingTop: 6, borderTop: "1px solid var(--border2)", alignItems: "center" }}>
                      <span style={{ fontSize: 10, color: "var(--text3)" }}>{nDone}/{rows.length} etapas</span>
                      {nOk   > 0 && <span style={{ background: "#eafaf1", color: "#1e8449", border: "1px solid #a9dfbf", borderRadius: 5, padding: "2px 8px", fontSize: 10, fontWeight: 700 }}>✓ {nOk} no prazo</span>}
                      {nLate > 0 && <span style={{ background: "#fdf0f0", color: "#c0392b", border: "1px solid #f1948a", borderRadius: 5, padding: "2px 8px", fontSize: 10, fontWeight: 700 }}>⚠ {nLate} atrasada{nLate > 1 ? "s" : ""}</span>}
                      {nLogged > 0 && <span style={{ background: "#f5eef8", color: "#8e44ad", border: "1px solid #d2b4de", borderRadius: 5, padding: "2px 8px", fontSize: 10, fontWeight: 700 }}>~ {nLogged} sem SLA</span>}
                      {nPend > 0 && <span style={{ background: "var(--card2)", color: "var(--text3)", border: "1px solid var(--border)", borderRadius: 5, padding: "2px 8px", fontSize: 10 }}>○ {nPend} pendente{nPend > 1 ? "s" : ""}</span>}
                      {currentIdx >= 0 && currentIdx < rows.length && r.emA && (
                        <span style={{ marginLeft: "auto", fontSize: 10, color: "#e67e22", fontWeight: 700 }}>▶ {SHORT_LBL[rows[currentIdx]?.etapa] || rows[currentIdx]?.etapa}</span>
                      )}
                    </div>
                  </div>
                );
              })() : null;

              return (
                <div key={r.ProcessKey + i} style={{ display: "flex", alignItems: "stretch", minHeight: 0 }}>
                  {/* Trilho esquerdo */}
                  <div style={{ display: "flex", flexDirection: "column", alignItems: "center", width: 36, flexShrink: 0, marginRight: 14 }}>
                    <div style={{ width: 3, flex: i === 0 ? "0 0 16px" : "1 1 16px", background: lineColor, opacity: i === 0 ? 0 : 0.35 }} />
                    <div style={{ width: 22, height: 22, borderRadius: "50%", background: sc, border: "3px solid var(--card)", boxShadow: `0 0 0 2.5px ${sc}55`, flexShrink: 0, zIndex: 1 }} />
                    <div style={{ width: 3, flex: "1 1 16px", background: lineColor, opacity: i === total - 1 ? 0 : 0.35 }} />
                  </div>
                  {/* Card */}
                  <div style={{ flex: 1, marginBottom: 10, marginTop: 2 }}>
                    <div className="gaq-card" style={{ padding: "14px 18px", borderRadius: 16, borderLeft: `3px solid ${sc}`, cursor: "pointer" }}
                      onClick={() => setSelProc(r)}>
                      <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 6, flexWrap: "wrap" }}>
                        <span style={{ fontWeight: 600, color: "var(--gaq-text)", fontSize: 13 }}>{r.TicketSD || r.NumRC || "—"}</span>
                        <span className="gaq-meta">{r["Área Requisitante"] || "—"}</span>
                        <span style={{ fontSize: 12, fontWeight: 600, color: sc }}>{resp}</span>
                      </div>
                      <div style={{ fontSize: 12, color: "var(--gaq-text-2)", lineHeight: 1.5, marginBottom: 10 }}>{(r.Objeto || "Sem descricao").slice(0, 220)}{(r.Objeto || "").length > 220 ? "..." : ""}</div>
                      {miniTrack}
                      <div style={{ display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center", borderTop: miniTrack ? "1px solid var(--border2)" : "none", paddingTop: miniTrack ? 8 : 0 }}>
                        {r.Modalidade && <span style={{ background: "#1565c015", color: "#1565c0", borderRadius: 6, padding: "2px 8px", fontSize: 10, fontWeight: 700 }}>{r.Modalidade}</span>}
                        <span style={{ background: sc + "20", color: sc, borderRadius: 6, padding: "2px 8px", fontSize: 10, fontWeight: 700 }}>{r.status || "—"}</span>
                        {r.faseAtual && r.faseAtual !== "—" && <span style={{ background: "var(--card2)", color: "var(--text2)", borderRadius: 6, padding: "2px 8px", fontSize: 10 }}>{r.faseAtual}</span>}
                        {r.diasParado > 10 && <span style={{ background: "#c0392b15", color: "#c0392b", borderRadius: 6, padding: "2px 8px", fontSize: 10, fontWeight: 700 }}>{r.diasParado} d.u. sem movimento</span>}
                        {r.diasTotais > 0 && <span style={{ marginLeft: "auto", fontWeight: 800, fontSize: 12, color: r.diasTotais > 100 ? "#c0392b" : r.diasTotais > 50 ? "#e67e22" : "#27ae60" }}>{r.diasTotais} d.u.</span>}
                      </div>
                    </div>
                  </div>
                </div>
              );
            };

            const groups = metroAgrupado
              ? compList
                  .map((c, gi) => ({ name: c, color: METRO_COLORS[gi % METRO_COLORS.length], items: metroSlice.filter(r => getResp(r) === c) }))
                  .filter(g => g.items.length > 0)
              : [{ name: null, color: "#1565c0", items: metroSlice }];

            return (
              <div style={{ minHeight: "max-content", paddingLeft: 20, paddingRight: 10 }}>
                {/* Cabecalho */}
                <div style={{ marginBottom: 18 }}>
                  <div className="gaq-h1" style={{ marginBottom: 4 }}>Linha do Metro</div>
                  <div className="gaq-body" style={{ color: "var(--gaq-text-3)" }}>Trilha visual de cada processo pelas estacoes da jornada de compra · {metroTotal.toLocaleString("pt-BR")} processos</div>
                </div>

                <div className="gaq-card" style={{ padding: "14px 20px", marginBottom: 18, display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                  <input type="text" value={metroBusca} onChange={e => { setMetroBusca(e.target.value); setPgMetro(1); }}
                    placeholder="Buscar pré-compra, RC, Suite, Processo, Objeto..."
                    style={{ padding: "8px 14px", borderRadius: 10, border: "1px solid var(--gaq-line)", background: "var(--gaq-surface)", color: "var(--gaq-text)", fontSize: 12, width: 260, outline: "none" }} />
                  <select value={metroComp} onChange={e => { setMetroComp(e.target.value); setPgMetro(1); }}
                    style={{ padding: "8px 12px", borderRadius: 10, border: "1px solid var(--gaq-line)", background: "var(--gaq-surface)", color: "var(--gaq-text)", fontSize: 12, cursor: "pointer" }}>
                    <option value="">Todos os responsaveis</option>
                    {compList.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                  <button onClick={() => { setMetroAgrupado(v => !v); setPgMetro(1); }} className="gaq-btn"
                    style={{ fontSize: 12 }}>
                    {metroAgrupado ? "Agrupado" : "Agrupar por responsavel"}
                  </button>
                  <select value={pgMetroSz} onChange={e => { setPgMetroSz(+e.target.value); setPgMetro(1); }}
                    style={{ padding: "8px 12px", borderRadius: 10, border: "1px solid var(--gaq-line)", background: "var(--gaq-surface)", color: "var(--gaq-text)", fontSize: 12, cursor: "pointer" }}>
                    {[50, 100].map(n => <option key={n} value={n}>{n} / pag</option>)}
                  </select>
                </div>

                <div style={{ paddingRight: 6, paddingBottom: 28 }}>
                  {filteredMetro.length === 0 ? (
                    <div style={{ textAlign: "center", padding: 60, color: "var(--text3)" }}>
                      <div style={{ fontSize: 32 }}>🚉</div>
                      <div style={{ marginTop: 8 }}>Nenhum processo encontrado com os filtros atuais.</div>
                    </div>
                  ) : (
                    <div>
                      <div style={{ display: "flex", flexDirection: "column", gap: 28 }}>
                        {groups.map((group, gi) => (
                          <div key={group.name || "all"}>
                            {group.name && (
                              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
                                <div style={{ width: 13, height: 13, borderRadius: "50%", background: group.color, flexShrink: 0 }} />
                                <div style={{ fontWeight: 800, fontSize: 15, color: "var(--text)" }}>{group.name}</div>
                                <div style={{ height: 2, flex: 1, background: group.color, opacity: .22, borderRadius: 2 }} />
                                <span style={{ background: group.color + "22", color: group.color, borderRadius: 10, padding: "2px 10px", fontSize: 11, fontWeight: 700 }}>{group.items.length}</span>
                              </div>
                            )}
                            <div style={{ paddingLeft: group.name ? 6 : 0 }}>
                              {group.items.map((r, i) => renderStation(r, i, group.items.length, group.color))}
                            </div>
                          </div>
                        ))}
                      </div>
                      {/* Paginação */}
                      {metroTotal > pgMetroSz && (
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 4px", flexWrap: "wrap", gap: 8, borderTop: "1px solid var(--border2)", marginTop: 8 }}>
                          <div style={{ fontSize: 12, color: "var(--text3)" }}>
                            {((metroPage-1)*pgMetroSz)+1}–{Math.min(metroPage*pgMetroSz, metroTotal)} de {metroTotal.toLocaleString("pt-BR")}
                          </div>
                          <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
                            <select value={pgMetroSz} onChange={e => { setPgMetroSz(+e.target.value); setPgMetro(1); }}
                              style={{ padding: "4px 8px", borderRadius: 6, border: "1px solid var(--input-bd)", fontSize: 12, background: "var(--input-bg)", color: "var(--text)" }}>
                              {[50, 100].map(n => <option key={n} value={n}>{n}/pág</option>)}
                            </select>
                            <button disabled={metroPage <= 1} onClick={() => setPgMetro(p => p - 1)}
                              style={{ padding: "4px 10px", borderRadius: 5, border: "1px solid var(--border)", background: "var(--card)", color: "var(--text)", cursor: metroPage <= 1 ? "default" : "pointer", fontSize: 12, fontWeight: 600, opacity: metroPage <= 1 ? .4 : 1 }}>◀</button>
                            <span style={{ fontSize: 12, fontWeight: 700, color: "var(--text2)", minWidth: 70, textAlign: "center" }}>
                              {metroPage} / {Math.ceil(metroTotal / pgMetroSz)}
                            </span>
                            <button disabled={metroPage >= Math.ceil(metroTotal / pgMetroSz)} onClick={() => setPgMetro(p => p + 1)}
                              style={{ padding: "4px 10px", borderRadius: 5, border: "1px solid var(--border)", background: "var(--card)", color: "var(--text)", cursor: metroPage >= Math.ceil(metroTotal / pgMetroSz) ? "default" : "pointer", fontSize: 12, fontWeight: 600, opacity: metroPage >= Math.ceil(metroTotal / pgMetroSz) ? .4 : 1 }}>▶</button>
                          </div>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              </div>
            );
          })()}
        </div>);
}
