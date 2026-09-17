import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function ProcessosScreen({ allPage, allProcesses, fBase, hideSDConcluido, pgAll, pgAllSz, phaseIntervals, renderBuscaPanel, setHideSDConcluido, setPgAll, setPgAllSz, setSelProc, sortAll, toggleSort }) {
  return (<div>
          {renderBuscaPanel()}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
              <div style={{ fontWeight: 700, fontSize: 15, color: "var(--gaq-text)" }}>
                Todos os Processos ({allProcesses.length}{hideSDConcluido && fBase.length !== allProcesses.length ? ` de ${fBase.length}` : ""})
              </div>
              <button onClick={() => { setHideSDConcluido(v => !v); setPgAll(1); }}
                style={{ display:"flex", alignItems:"center", gap:6, padding:"5px 12px", borderRadius:20, cursor:"pointer", fontSize:11, fontWeight:600, border:`1.5px solid ${hideSDConcluido ? "#dc2626" : "var(--gaq-line)"}`, background: hideSDConcluido ? "#fee2e2" : "var(--gaq-bg-2)", color: hideSDConcluido ? "#dc2626" : "var(--gaq-text-3)", transition:"all .15s" }}>
                {hideSDConcluido ? "✕ Ocultando SD Concluído" : "Ocultar SD Concluído"}
              </button>
            </div>
            <button className="gaq-btn" onClick={() => exportToExcel(fBase, "processos.xlsx")}>
              <Icon name="download" size={13}/> Exportar Excel
            </button>
          </div>
          <div className="gaq-card" style={{ padding: 0, overflow: "auto" }}>
            <div className="gaq-table-scroll"><table className="gaq-table" style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  {[
                    { key: "ProcessKey", label: "Processo", w: 130 },
                    { key: "Objeto", label: "Objeto", w: 220 },
                    { key: "Modalidade", label: "Modalidade", w: 110 },
                    { key: "_tags", label: "Tags", w: 100 },
                    { key: "status", label: "Status", w: 100 },
                    { key: "diasParado", label: "Último Movimento", w: 80 },
                    { key: "diasTotais", label: "Aging", w: 100 },
                    { key: "_score", label: "Score", w: 64 },
                  ].map(({ key, label, w }) => (
                    <th key={key} onClick={() => key !== "_score" && key !== "_tags" && toggleSort(key)}
                      style={{ padding: "10px 10px", textAlign: "left", cursor: key !== "_score" ? "pointer" : "default", color: "var(--gaq-text-3)", fontWeight: 600, fontSize: 10.5, textTransform: "uppercase", letterSpacing: "0.04em", whiteSpace: "nowrap", minWidth: w, userSelect: "none" }}>
                      {label} {sortAll.col === key && (sortAll.dir === "desc" ? "▼" : "▲")}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {allPage.map((r, i) => {
                  const sla = calcSLAClassification(r, phaseIntervals);
                  const bk = sla ? (SLA_BUCKETS[sla.bucket] || SLA_BUCKETS.no_prazo) : null;
                  const pctAging = sla ? Math.min(100, sla.pctSLALabel) : 0;
                  const agingColor = sla ? sla.color : "var(--gaq-text-3)";
                  const subareaColor = (sa) => sa === "NCL" ? "var(--gaq-ncl)" : sa === "CPL" ? "var(--gaq-cpl)" : sa === "Scont" ? "var(--gaq-rsp)" : "var(--gaq-text-3)";
                  const subareaColorBg = (sa) => sa === "NCL" ? "rgba(10,132,255,0.08)" : sa === "CPL" ? "rgba(94,92,230,0.08)" : sa === "Scont" ? "rgba(52,199,89,0.08)" : "transparent";
                  const stDot = r.emA ? "#34c759" : r.isConcluded ? "#007aff" : r.isCanceled ? "#ff9500" : r.isFailed ? "#ff3b30" : "#8e8e93";
                  const statusLabel = r.statusDet || r.status || "—";
                  return (
                    <tr key={i} style={{ borderBottom: "1px solid var(--gaq-line)", cursor: "pointer", transition: "background .15s" }}
                      onMouseEnter={e => e.currentTarget.style.background = "var(--gaq-bg-2)"} onMouseLeave={e => e.currentTarget.style.background = "transparent"}
                      onClick={() => setSelProc(r)}>
                      <td style={{ padding: "12px 10px" }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          <div>
                            <div style={{ fontSize: 13, fontWeight: 600, color: "var(--gaq-text)", fontFamily: "var(--gaq-mono)" }}>{r.NumRC || r.TicketSD || r.ProcessKey || "—"}</div>
                            <div style={{ display: "flex", gap: 4, marginTop: 3 }}>
                              {r.faseSubarea && <span style={{ fontSize: 9, fontWeight: 600, background: subareaColorBg(r.faseSubarea), color: subareaColor(r.faseSubarea), borderRadius: 4, padding: "1px 6px" }}>{r.faseSubarea}</span>}
                              {r.ehCPL && r.faseSubarea !== "CPL" && <span style={{ fontSize: 9, fontWeight: 600, background: "rgba(94,92,230,0.08)", color: "var(--gaq-cpl)", borderRadius: 4, padding: "1px 6px" }}>CPL</span>}
                              {r.isLegado && <span style={{ fontSize: 8, background: "var(--gaq-bg-2)", color: "var(--gaq-text-3)", borderRadius: 3, padding: "1px 5px" }}>Legado</span>}
                            </div>
                          </div>
                        </div>
                      </td>
                      <td style={{ padding: "12px 10px", maxWidth: 260 }}>
                        <div style={{ fontSize: 12.5, color: "var(--gaq-text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: 500, lineHeight: 1.35 }} title={r.Objeto || ""}>{r.Objeto || "Sem objeto"}</div>
                        <div style={{ fontSize: 11, color: "var(--gaq-text-3)", marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                          {r["Área Requisitante"] || "Área n/i"} · {r.respAtivo && r.respAtivo !== "N/A" ? r.respAtivo : (r.Comprador || r.Pregoeiro || "—")}
                        </div>
                      </td>
                      <td style={{ padding: "12px 10px", fontSize: 12, color: "var(--gaq-text-2)" }}>{r.Modalidade || "—"}</td>
                      <td style={{ padding: "8px 10px", maxWidth: 140 }}>
                        {(() => { const tags = TagsManager.getForProcess(r.ProcessKey, r.TicketSD); return tags.length > 0 ? <div style={{ display:"flex", flexWrap:"wrap", gap:3 }}>{tags.slice(0,3).map(t => <TagBadge key={t} tag={t} small />)}{tags.length > 3 && <span style={{ fontSize:9, color:"var(--gaq-text-3)" }}>+{tags.length-3}</span>}</div> : <span style={{ fontSize:11, color:"var(--gaq-text-3)" }}>—</span>; })()}
                      </td>
                      <td style={{ padding: "12px 10px" }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                          <span style={{ width: 7, height: 7, borderRadius: 999, background: stDot, flex: "none" }} />
                          <span style={{ fontSize: 12, color: "var(--gaq-text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 120 }} title={statusLabel}>{statusLabel}</span>
                        </div>
                      </td>
                      <td style={{ padding: "12px 10px", fontVariantNumeric: "tabular-nums" }}>
                        <span style={{ fontSize: 12, fontWeight: 500, color: r.diasParado > 15 ? "var(--gaq-orange)" : "var(--gaq-text-2)" }}>
                          {r.emA && r.diasParado > 0 ? `${r.diasParado} d.u.` : "—"}
                        </span>
                      </td>
                      <td style={{ padding: "12px 10px", minWidth: 100 }}>
                        {r.emA ? (
                          <div>
                            <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 3 }}>
                              <span style={{ fontSize: 12, fontWeight: 600, color: agingColor, fontVariantNumeric: "tabular-nums" }}>{r.diasTotais || 0} d.u.</span>
                            </div>
                            <div style={{ height: 5, background: "rgba(0,0,0,0.06)", borderRadius: 3, overflow: "hidden", width: "100%" }}>
                              <div style={{ width: `${Math.min(100, pctAging)}%`, height: "100%", background: agingColor, borderRadius: 3, transition: "width .3s" }} />
                            </div>
                            <div style={{ fontSize: 9, color: "var(--gaq-text-3)", marginTop: 2 }}>{pctAging}% do SLA ({sla?.prazo || 0} d.u.)</div>
                          </div>
                        ) : (
                          <span style={{ fontSize: 12, color: "var(--gaq-text-3)" }}>{r.diasTotais || "—"}</span>
                        )}
                      </td>
                      <td style={{ padding: "12px 10px", textAlign: "center" }}>
                        {sla ? (
                          <div style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", width: 38, height: 38, borderRadius: 999, background: bk.bg, border: `2px solid ${bk.color}` }}>
                            <span style={{ fontSize: 12, fontWeight: 700, color: bk.color, fontVariantNumeric: "tabular-nums" }}>{sla.processScore}</span>
                          </div>
                        ) : (
                          <span style={{ fontSize: 11, color: "var(--gaq-text-3)" }}>—</span>
                        )}
                      </td>
                    </tr>
                  );
                })}
                {allPage.length === 0 && <tr><td colSpan={8} style={{ textAlign: "center", padding: 30, color: "var(--gaq-text-3)" }}>Nenhum processo encontrado.</td></tr>}
              </tbody>
            </table></div>
          </div>
          <Pagination total={allProcesses.length} page={pgAll} pageSize={pgAllSz} onPage={setPgAll} onSize={setPgAllSz} />
        </div>);
}
