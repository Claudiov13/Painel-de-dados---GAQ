import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function PantanalScreen({ base, pantanalBusca, pgPantanal, pgPantanalSz, setAba, setPantanalBusca, setPantanalSel, setPgPantanal, setPgPantanalSz }) {
  return (<div className="anim-fade" style={{ minHeight: "max-content", overflow: "visible", paddingBottom: 40 }}>
          {base.length === 0 ? (
            <div style={{ textAlign: "center", padding: 60, color: "var(--text3)" }}>
              <div style={{ fontSize: 38 }}>🌿</div>
              <div style={{ fontWeight: 600, marginTop: 8 }}>Carregue a base para ver as oportunidades.</div>
              <button onClick={() => setAba("upload")} style={{ marginTop: 12, background: "#1e8449", color: "#fff", border: "none", borderRadius: 8, padding: "9px 22px", cursor: "pointer", fontSize: 13 }}>Upload</button>
            </div>
          ) : (() => {
            const hoje = new Date(); hoje.setHours(0,0,0,0);
            const buscaNrm = nrm(pantanalBusca);
            const pantanalProcs = base
              .filter(r =>
                r.emA &&
                r.aberturaSD &&
                !pd(r["Data do encerramento"]) &&
                (buscaNrm.length < 2 ||
                  nrm(r.Objeto).includes(buscaNrm) ||
                  nrm(r.NumRC).includes(buscaNrm) ||
                  nrm(r["Área Requisitante"]).includes(buscaNrm) ||
                  nrm(r.Comprador).includes(buscaNrm))
              )
              .sort((a, b) => (b.aberturaSD || 0) - (a.aberturaSD || 0));

            const panTotal = pantanalProcs.length;
            const panPage  = Math.min(pgPantanal, Math.ceil(panTotal / pgPantanalSz) || 1);
            const panSlice = pantanalProcs.slice((panPage - 1) * pgPantanalSz, panPage * pgPantanalSz);

            const ageInfo = dias => dias < 20
              ? { color: "#1e8449", bg: "#eafaf1", label: "Muito recente", icon: "🟢" }
              : dias < 50
              ? { color: "#e67e22", bg: "#fef9e7", label: "Recente",       icon: "🟡" }
              : { color: "#c0392b", bg: "#fdf0f0", label: "Atenção",       icon: "🟠" };

            const muito  = pantanalProcs.filter(r => du(r.aberturaSD, hoje) < 20).length;
            const recente = pantanalProcs.filter(r => { const d = du(r.aberturaSD, hoje); return d >= 20 && d < 50; }).length;
            const antigo  = pantanalProcs.filter(r => du(r.aberturaSD, hoje) >= 50).length;

            return (
              <div style={{ minHeight: "max-content" }}>
                {/* Cabecalho */}
                <div style={{ marginBottom: 18 }}>
                  <div className="gaq-h1" style={{ marginBottom: 4 }}>Pantanal</div>
                  <div className="gaq-body" style={{ color: "var(--gaq-text-3)" }}>Padroes emergentes na base · keywords e clusters de objetos com comportamento atipico</div>
                </div>

                <div className="gaq-kpi-grid gaq-responsive-grid" style={{ gridTemplateColumns: "repeat(3, 1fr)", marginBottom: 18 }}>
                  <div className="gaq-kpi" style={{ borderTop: "3px solid #1e8449" }}>
                    <span className="lbl">Muito recentes (&lt;20 d.u.)</span>
                    <div className="val" style={{ color: "#1e8449" }}>{muito}</div>
                  </div>
                  <div className="gaq-kpi" style={{ borderTop: "3px solid #e67e22" }}>
                    <span className="lbl">Recentes (20-49 d.u.)</span>
                    <div className="val" style={{ color: "#e67e22" }}>{recente}</div>
                  </div>
                  <div className="gaq-kpi" style={{ borderTop: "3px solid #c0392b" }}>
                    <span className="lbl">Atencao (≥50 d.u.)</span>
                    <div className="val" style={{ color: "#c0392b" }}>{antigo}</div>
                  </div>
                </div>

                <div style={{ marginBottom: 16 }}>
                  <input value={pantanalBusca} onChange={e => { setPantanalBusca(e.target.value); setPgPantanal(1); }}
                    placeholder="Buscar por objeto, RC, area ou comprador..."
                    className="gaq-input"
                    style={{ width: "100%", padding: "10px 16px", borderRadius: 12, border: "1px solid var(--gaq-line)", background: "var(--gaq-surface)", color: "var(--gaq-text)", fontSize: 13, boxSizing: "border-box", outline: "none" }} />
                </div>

                {panTotal === 0 ? (
                  <div style={{ textAlign: "center", padding: 60, color: "var(--text3)" }}>
                    <div style={{ fontSize: 32 }}>🌱</div>
                    <div style={{ marginTop: 8 }}>Nenhum processo encontrado com os filtros atuais.</div>
                  </div>
                ) : (
                  <div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                      {panSlice.map((r, i) => {
                        const dias = du(r.aberturaSD, hoje);
                        const ai   = ageInfo(dias);
                        const resp = r.Comprador || r.Avaliador || r.cplResp || "—";
                        const dtAb = r.aberturaSD ? r.aberturaSD.toLocaleDateString("pt-BR") : "—";
                        return (
                          <div key={r.ProcessKey + i} className="gaq-card" style={{ padding: "16px 20px", marginBottom: 0, borderRadius: 18, borderLeft: `3px solid ${ai.color}`, display: "flex", gap: 14, alignItems: "flex-start", cursor: "pointer" }}>
                            {/* Aging badge */}
                            <div style={{ flexShrink: 0, textAlign: "center", minWidth: 48 }}>
                              <div style={{ width: 44, height: 44, borderRadius: 12, background: ai.bg || `${ai.color}15`, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 14, color: ai.color }}>{dias}</div>
                              <div style={{ fontSize: 9, color: "var(--gaq-text-3)", marginTop: 3 }}>{ai.label}</div>
                            </div>
                            {/* Conteudo principal */}
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 6, flexWrap: "wrap" }}>
                                <span style={{ fontWeight: 600, color: "var(--gaq-text)", fontSize: 13 }}>{r.NumRC || r.TicketSD || "—"}</span>
                                {r.TicketSD && r.NumRC && <span style={{ fontSize: 11, color: "var(--gaq-green)" }}>Pré-compra {r.TicketSD}</span>}
                                <span className="gaq-meta">{r["Área Requisitante"] || "—"}</span>
                                <span className="gaq-meta">Abertura: {dtAb}</span>
                              </div>
                              <div style={{ fontSize: 13, fontWeight: 400, color: "var(--gaq-text-2)", marginBottom: 8, lineHeight: 1.45 }}>
                                {(r.Objeto || "Sem descricao").slice(0, 160)}{(r.Objeto||"").length > 160 ? "..." : ""}
                              </div>
                              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                                {r.Modalidade && <span style={{ background: "rgba(21,101,192,0.08)", color: "#1565c0", borderRadius: 999, padding: "3px 10px", fontSize: 11, fontWeight: 600 }}>{r.Modalidade}</span>}
                                {resp !== "—" && <span style={{ background: "rgba(30,132,73,0.08)", color: "#1e8449", borderRadius: 999, padding: "3px 10px", fontSize: 11, fontWeight: 600 }}>{resp}</span>}
                                {r.faseAtual && r.faseAtual !== "—" && <span style={{ background: "var(--gaq-bg-2)", color: "var(--gaq-text-2)", borderRadius: 999, padding: "3px 10px", fontSize: 11 }}>{r.faseAtual}</span>}
                              </div>
                            </div>
                            {/* Botao */}
                            <button onClick={() => setPantanalSel(r)} className="gaq-btn is-primary"
                              style={{ flexShrink: 0, fontSize: 11, whiteSpace: "nowrap" }}>
                              Solicitar Adesao
                            </button>
                          </div>
                        );
                      })}
                    </div>
                    {/* Paginação */}
                    {panTotal > pgPantanalSz && (
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "14px 4px", flexWrap: "wrap", gap: 8, borderTop: "1px solid var(--border2)", marginTop: 12 }}>
                        <div style={{ fontSize: 12, color: "var(--text3)" }}>
                          {((panPage-1)*pgPantanalSz)+1}–{Math.min(panPage*pgPantanalSz, panTotal)} de {panTotal.toLocaleString("pt-BR")} processos
                        </div>
                        <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
                          <select value={pgPantanalSz} onChange={e => { setPgPantanalSz(+e.target.value); setPgPantanal(1); }}
                            style={{ padding: "4px 8px", borderRadius: 6, border: "1px solid var(--input-bd)", fontSize: 12, background: "var(--input-bg)", color: "var(--text)" }}>
                            {[50, 100].map(n => <option key={n} value={n}>{n}/pág</option>)}
                          </select>
                          <button disabled={panPage <= 1} onClick={() => setPgPantanal(p => p - 1)}
                            style={{ padding: "4px 10px", borderRadius: 5, border: "1px solid var(--border)", background: "var(--card)", color: "var(--text)", cursor: panPage <= 1 ? "default" : "pointer", fontSize: 12, fontWeight: 600, opacity: panPage <= 1 ? .4 : 1 }}>◀</button>
                          <span style={{ fontSize: 12, fontWeight: 700, color: "var(--text2)", minWidth: 70, textAlign: "center" }}>
                            {panPage} / {Math.ceil(panTotal / pgPantanalSz)}
                          </span>
                          <button disabled={panPage >= Math.ceil(panTotal / pgPantanalSz)} onClick={() => setPgPantanal(p => p + 1)}
                            style={{ padding: "4px 10px", borderRadius: 5, border: "1px solid var(--border)", background: "var(--card)", color: "var(--text)", cursor: panPage >= Math.ceil(panTotal / pgPantanalSz) ? "default" : "pointer", fontSize: 12, fontWeight: 600, opacity: panPage >= Math.ceil(panTotal / pgPantanalSz) ? .4 : 1 }}>▶</button>
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            );
          })()}
        </div>);
}
