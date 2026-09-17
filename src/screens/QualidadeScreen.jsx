import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function QualidadeScreen({ baseVis, drillDown, pgDrillDown, pgDrillDownSz, phaseIntervals, qualityData, setAba, setDrillDown, setPgDrillDown, setPgDrillDownSz, setSelProc }) {
  return (qualityData ? <div>
          {drillDown && drillDown.sourceAba === "qualidade" ? <DrillDownPanel
            drillDown={drillDown}
            defaultSourceAba="qualidade"
            setDrillDown={setDrillDown}
            setAba={setAba}
            setSelProc={setSelProc}
            pgDrillDown={pgDrillDown}
            pgDrillDownSz={pgDrillDownSz}
            setPgDrillDown={setPgDrillDown}
            setPgDrillDownSz={setPgDrillDownSz}
            phaseIntervals={phaseIntervals}
          /> : <>
            <div style={{ background: "linear-gradient(135deg, #1f2937 0%, #0f172a 45%, #1d4ed8 100%)", borderRadius: "var(--gaq-r-xl)", padding: "26px 30px", marginBottom: 18, color: "#fff", boxShadow: "var(--gaq-shadow-2)" }}>
              <div style={{ display: "flex", justifyContent: "space-between", gap: 16, flexWrap: "wrap", alignItems: "flex-start" }}>
                <div>
                  <div className="gaq-eyebrow" style={{ color: "rgba(255,255,255,0.68)", marginBottom: 8 }}>Admin Master · Qualidade de Dados</div>
                  <div style={{ fontSize: 28, fontWeight: 800, letterSpacing: "-0.03em", marginBottom: 6 }}>Painel de Qualidade da Base</div>
                  <div style={{ fontSize: 13, color: "rgba(255,255,255,0.76)", maxWidth: 780, lineHeight: 1.55 }}>
                    Monitoriza campos críticos, duplicidades e incoerências de datas para reduzir erros de leitura, cronograma e SLA em todo o sistema.
                  </div>
                </div>
                <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(110px, 1fr))", gap: 10, minWidth: 250 }}>
                  {[
                    { label: "Registos", value: qualityData.totalRows, color: "#bfdbfe" },
                    { label: "Pendências", value: qualityData.totalIssues, color: "#fed7aa" },
                    { label: "Críticas", value: qualityData.criticalCount, color: "#fecaca" },
                    { label: "Score", value: qualityData.score, color: "#bbf7d0" },
                  ].map(item => (
                    <div key={item.label} style={{ background: "rgba(255,255,255,0.12)", border: "1px solid rgba(255,255,255,0.12)", borderRadius: 18, padding: "14px 16px" }}>
                      <div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", color: "rgba(255,255,255,0.64)", marginBottom: 4 }}>{item.label}</div>
                      <div className="gaq-num" style={{ fontSize: 24, fontWeight: 800, color: item.color }}>{item.value.toLocaleString ? item.value.toLocaleString("pt-BR") : item.value}</div>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 12, marginBottom: 18 }}>
              {[
                { label: "Processos avaliados", value: qualityData.totalRows, color: "#1f6feb", data: baseVis, title: "Qualidade — Processos avaliados" },
                { label: "Com pendência", value: qualityData.totalIssues, color: "#ea580c", data: qualityData.allIssueRows, title: "Qualidade — Processos com pendência" },
                { label: "Erros críticos", value: qualityData.criticalCount, color: "#dc2626", data: qualityData.criticalRows, title: "Qualidade — Erros críticos" },
                { label: "Alertas de cadastro", value: qualityData.attentionCount, color: "#f59e0b", data: qualityData.attentionRows, title: "Qualidade — Alertas de cadastro" },
                { label: "Score da base", value: qualityData.score, color: qualityData.score >= 85 ? "#16a34a" : qualityData.score >= 70 ? "#f59e0b" : "#dc2626", data: qualityData.allIssueRows, title: "Qualidade — Base classificada" },
              ].map(card => (
                <div key={card.label} onClick={() => setDrillDown({ title: card.title, data: card.data, color: card.color, sourceAba: "qualidade" })}
                  className="gaq-card" style={{ padding: 18, borderRadius: 22, cursor: "pointer" }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "var(--gaq-text-3)", textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: 8 }}>{card.label}</div>
                  <div className="gaq-num" style={{ fontSize: 30, fontWeight: 800, color: card.color }}>{card.value.toLocaleString ? card.value.toLocaleString("pt-BR") : card.value}</div>
                </div>
              ))}
            </div>

            <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "minmax(0, 1.3fr) minmax(0, 1fr)", gap: 18, marginBottom: 18 }}>
              <div className="gaq-card" style={{ padding: 22 }}>
                <div className="gaq-section-h">
                  <div>
                    <div className="gaq-h3">Regras monitoradas</div>
                    <div className="sub">Clique numa regra para ver os processos afectados.</div>
                  </div>
                </div>
                <div style={{ display: "grid", gap: 10, marginTop: 8 }}>
                  {qualityData.issueDefs.map(issue => (
                    <button
                      key={issue.key}
                      onClick={() => setDrillDown({ title: `Qualidade — ${issue.label}`, data: issue.rows, color: issue.color, sourceAba: "qualidade" })}
                      style={{ textAlign: "left", border: `1px solid ${issue.color}22`, borderLeft: `4px solid ${issue.color}`, background: issue.count ? `${issue.color}10` : "var(--gaq-bg-2)", borderRadius: 18, padding: "14px 16px", cursor: issue.count ? "pointer" : "default", opacity: issue.count ? 1 : 0.72 }}
                    >
                      <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center", marginBottom: 6 }}>
                        <div style={{ fontSize: 13, fontWeight: 700, color: "var(--gaq-text)" }}>{issue.label}</div>
                        <div className="gaq-num" style={{ fontSize: 22, fontWeight: 800, color: issue.color }}>{issue.count}</div>
                      </div>
                      <div style={{ fontSize: 12, color: "var(--gaq-text-2)", lineHeight: 1.5 }}>{issue.desc}</div>
                      <div style={{ fontSize: 10, color: issue.color, fontWeight: 700, textTransform: "uppercase", marginTop: 8 }}>
                        {issue.severity === "critical" ? "Erro crítico" : "Alerta de cadastro"}
                      </div>
                    </button>
                  ))}
                </div>
              </div>

              <div className="gaq-card" style={{ padding: 22 }}>
                <div className="gaq-section-h">
                  <div>
                    <div className="gaq-h3">Pressão por subárea</div>
                    <div className="sub">Percentual do recorte com pendências por subárea actual.</div>
                  </div>
                </div>
                <div style={{ display: "grid", gap: 12, marginTop: 8 }}>
                  {qualityData.bySubarea.map(item => (
                    <div key={item.label}
                      onClick={() => setDrillDown({ title: `Qualidade — ${item.label}`, data: item.rows, color: item.pct >= 40 ? "#dc2626" : item.pct >= 20 ? "#f59e0b" : "#16a34a", sourceAba: "qualidade" })}
                      style={{ cursor: "pointer" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", gap: 8, marginBottom: 6 }}>
                        <span style={{ fontSize: 12.5, fontWeight: 600, color: "var(--gaq-text)" }}>{item.label}</span>
                        <span className="gaq-meta gaq-num">{item.count}/{item.total} · {item.pct}%</span>
                      </div>
                      <div style={{ height: 10, background: "rgba(15,23,42,0.06)", borderRadius: 999, overflow: "hidden" }}>
                        <div style={{ width: `${Math.max(item.pct, item.count > 0 ? 6 : 0)}%`, height: "100%", background: item.pct >= 40 ? "#dc2626" : item.pct >= 20 ? "#f59e0b" : "#16a34a", borderRadius: 999 }} />
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            <div className="gaq-card" style={{ padding: 22 }}>
              <div className="gaq-section-h">
                <div>
                  <div className="gaq-h3">Exemplos prioritários</div>
                  <div className="sub">Primeiros processos afectados para triagem rápida da base.</div>
                </div>
                <button className="gaq-btn is-ghost" style={{ fontSize: 12 }} onClick={() => setDrillDown({ title: "Qualidade — Todos os processos com pendência", data: qualityData.allIssueRows, color: "#ea580c", sourceAba: "qualidade" })}>
                  Ver todos <Icon name="chevR" size={12}/>
                </button>
              </div>
              {qualityData.topExamples.length === 0
                ? <div style={{ color: "var(--gaq-text-3)", fontSize: 13, paddingTop: 8 }}>Nenhuma pendência identificada no recorte actual.</div>
                : qualityData.topExamples.map((r, idx) => (
                  <div className="gaq-responsive-grid" key={`${r.ProcessKey || idx}-dq`} onClick={() => setSelProc(r)}
                    style={{ display: "grid", gridTemplateColumns: "34px 130px 1fr 140px 26px", gap: 12, alignItems: "center", padding: "12px 0", borderTop: idx ? "1px solid var(--gaq-line)" : "none", cursor: "pointer" }}>
                    <span style={{ fontSize: 11, fontWeight: 700, color: "var(--gaq-text-3)" }}>#{idx + 1}</span>
                    <span style={{ fontFamily: "var(--gaq-mono)", fontSize: 12, fontWeight: 700, color: "var(--gaq-text)" }}>{r.NumRC || r.TicketSD || r.ProcessKey || "—"}</span>
                    <div style={{ minWidth: 0 }}>
                      <div style={{ fontSize: 13, fontWeight: 600, color: "var(--gaq-text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.Objeto || "Sem objecto"}</div>
                      <div className="gaq-meta" style={{ marginTop: 3, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r._dqDetail || "Pendência identificada"} · {r["Área Requisitante"] || "Área não informada"}</div>
                    </div>
                    <span style={{ justifySelf: "start", fontSize: 11, fontWeight: 700, color: r._dqIssue && r._dqIssue.includes("dup") ? "#e74c3c" : r._dqIssue && r._dqIssue.includes("sem_") ? "#ff9500" : "#d35400", background: r._dqIssue && r._dqIssue.includes("dup") ? "#fff1f0" : r._dqIssue && r._dqIssue.includes("sem_") ? "#fff7e6" : "#fff3e8", borderRadius: 999, padding: "6px 10px", whiteSpace: "nowrap" }}>
                      {r._dqIssue ? r._dqIssue.replaceAll("_", " ") : "pendência"}
                    </span>
                    <button className="gaq-icon-btn" onClick={(e) => { e.stopPropagation(); setSelProc(r); }}><Icon name="chevR" size={14}/></button>
                  </div>
                ))}
            </div>
          </>}
        </div> : <div className="gaq-card" style={{ padding: 28, textAlign: "center" }}>Carregue a base para analisar a qualidade dos dados.</div>);
}
