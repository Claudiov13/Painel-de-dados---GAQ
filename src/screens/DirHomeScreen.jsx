import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function DirHomeScreen({ base, dirAreaSel, dirExpandNCL, dirExpandNCLScont, dirScopedBase, dirShowSDList, loginDiretor, setDirAreaSel, setDirExpandNCL, setDirExpandNCLScont, setDirShowSDList, setSelProc }) {
  return ((() => {
          // Dados filtrados por área selecionada (ou todas se nenhuma)
          const dirBase = dirScopedBase;
          const dirBack = dirBase.filter(r => r.emA);
          // SD em andamento (tem SD, sem RC, em andamento)
          const sdEmAndamento = dirBack.filter(r => r.TicketSD && r.TicketSD.trim() !== "" && (!r.NumRC || r.NumRC.trim() === ""));
          // SD concluídos pela GAQ sem RC aberta (encerramento SD preenchido, sem RC, em andamento)
          const sdConcluidosSemRC = dirBack.filter(r => r.encSD && (!r.NumRC || r.NumRC.trim() === ""));
          // RC em andamento
          const rcEmAndamento = dirBack.filter(r => r.NumRC && r.NumRC.trim() !== "");
          // Top 10 RC mais antigas na NCL
          const top10NCL = dirBack
            .filter(r => r.NumRC && r.NumRC.trim() !== "" && r.faseSubarea === "NCL")
            .sort((a, b) => (b.diasTotaisGestao || 0) - (a.diasTotaisGestao || 0))
            .slice(0, 10);
          // Top 10 RC mais antigas NCL ou Scont (processos que já saíram/estão em NCL ou Scont)
          const top10NCLScont = dirBack
            .filter(r => r.NumRC && r.NumRC.trim() !== "" && (r.faseSubarea === "NCL" || r.faseSubarea === "Scont"))
            .sort((a, b) => (b.diasTotaisGestao || 0) - (a.diasTotaisGestao || 0))
            .slice(0, 10);
          // Áreas disponíveis na base
          const areasDisp = [...new Set(base.map(r => r["Área Requisitante"]).filter(Boolean))].sort();
          // Áreas em destaque (do perfil do diretor)
          const areasDestaque = loginDiretor.areasDestaque || [];
          // Expandir bullets
          const expandNCL = dirExpandNCL;
          const setExpandNCL = setDirExpandNCL;
          const expandNCLScont = dirExpandNCLScont;
          const setExpandNCLScont = setDirExpandNCLScont;
          const showSDList = dirShowSDList;
          const setShowSDList = setDirShowSDList;

          const thS = { padding: "10px 8px", textAlign: "left", fontSize: 10.5, fontWeight: 600, color: "var(--gaq-text-3)", textTransform: "uppercase", letterSpacing: "0.04em" };
          const tdS = { padding: "10px 8px", fontSize: 12.5, color: "var(--gaq-text-2)" };
          const renderTopTable = (items) => (
            <div className="gaq-table-scroll"><table className="gaq-table" style={{ width: "100%", borderCollapse: "collapse", marginTop: 8 }}>
              <thead>
                <tr>
                  <th style={thS}>#</th>
                  <th style={thS}>N RC</th>
                  <th style={{ ...thS, maxWidth: 200 }}>Objeto</th>
                  <th style={thS}>Area</th>
                  <th style={thS}>Abertura RC</th>
                  <th style={thS}>Comprador</th>
                  <th style={thS}>Prazo Entrega</th>
                  <th style={{ ...thS, textAlign: "center" }}>Dias (RC)</th>
                  <th style={{ ...thS, textAlign: "center" }}>Dias (pré-compra)</th>
                </tr>
              </thead>
              <tbody>
                {items.map((r, i) => {
                  const cor = (r.diasTotaisGestao || 0) > 100 ? "var(--gaq-red)" : (r.diasTotaisGestao || 0) > 60 ? "var(--gaq-orange)" : "var(--gaq-green)";
                  return (
                    <tr key={r.ProcessKey || i} onClick={() => setSelProc(r)} style={{ cursor: "pointer", borderBottom: "1px solid var(--gaq-line)", transition: "background .15s" }}
                      onMouseEnter={e => e.currentTarget.style.background = "var(--gaq-bg-2)"}
                      onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                      <td style={{ ...tdS, fontWeight: 600, color: "var(--gaq-text-3)" }}>{i + 1}</td>
                      <td style={{ ...tdS, fontWeight: 600, color: "var(--gaq-blue)" }}>{r.NumRC || "—"}</td>
                      <td style={{ ...tdS, color: "var(--gaq-text)", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={r.Objeto}>{(r.Objeto || "").slice(0, 60) || "—"}</td>
                      <td style={tdS}>{r["Área Requisitante"] || "—"}</td>
                      <td style={tdS}>{r.aberturaRC ? r.aberturaRC.toLocaleDateString("pt-BR") : "—"}</td>
                      <td style={tdS}>{r.Comprador || r.respNCL || "—"}</td>
                      <td style={tdS}>{r.dataEntrega ? r.dataEntrega.toLocaleDateString("pt-BR") : "—"}</td>
                      <td style={{ ...tdS, textAlign: "center", fontWeight: 700, color: cor, fontVariantNumeric: "tabular-nums" }}>{r.diasTotaisGestao || 0}</td>
                      <td style={{ ...tdS, textAlign: "center", fontWeight: 600, color: "var(--gaq-cpl)", fontVariantNumeric: "tabular-nums" }}>{r.diasSD || 0}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table></div>
          );

          return (
            <div>
              {/* Header do diretor */}
              <div style={{ background: "var(--gaq-surface)", borderRadius: "var(--gaq-r-xl)", padding: "24px 28px", marginBottom: 20, border: "1px solid var(--gaq-line)", boxShadow: "var(--gaq-shadow-2)" }}>
                <div style={{ fontSize: 22, fontWeight: 700, color: "var(--gaq-text)", letterSpacing: "-0.02em", marginBottom: 4 }}>Painel da Diretoria</div>
                <div style={{ fontSize: 13, color: "var(--gaq-text-3)" }}>Visao consolidada · {loginDiretor.nome} · {dirBase.length.toLocaleString("pt-BR")} processos{dirAreaSel ? ` · Area: ${dirAreaSel}` : ""}</div>
              </div>

              {/* Areas em destaque */}
              <div style={{ marginBottom: 20 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: "var(--gaq-text)", marginBottom: 10 }}>Areas em Destaque</div>
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 10 }}>
                  <button onClick={() => setDirAreaSel(null)}
                    className={`gaq-chip${!dirAreaSel ? " is-active" : ""}`}
                    style={{ padding: "10px 20px", borderRadius: "var(--gaq-r-md)", fontSize: 14, fontWeight: 600 }}>
                    Todas <span style={{ fontSize: 11, opacity: .7, marginLeft: 4 }}>{base.filter(r => r.emA).length}</span>
                  </button>
                  {areasDestaque.map(area => {
                    const active = dirAreaSel === area;
                    const count = base.filter(r => r.emA && nrm(r["Área Requisitante"] || "").includes(nrm(area))).length;
                    return (
                      <button key={area} onClick={() => setDirAreaSel(active ? null : area)}
                        className={`gaq-chip${active ? " is-active" : ""}`}
                        style={{ padding: "10px 20px", borderRadius: "var(--gaq-r-md)", fontSize: 14, fontWeight: 600 }}>
                        {area} <span style={{ fontSize: 11, opacity: .7, marginLeft: 4 }}>{count}</span>
                      </button>
                    );
                  })}
                  <select value={dirAreaSel && !areasDestaque.includes(dirAreaSel) ? dirAreaSel : ""} onChange={e => setDirAreaSel(e.target.value || null)}
                    className="gaq-input" style={{ padding: "8px 12px", fontSize: 12, width: "auto", minWidth: 140 }}>
                    <option value="">Outra area...</option>
                    {areasDisp.filter(a => !areasDestaque.includes(a)).map(a => <option key={a} value={a}>{a}</option>)}
                  </select>
                </div>
              </div>

              {/* KPIs principais */}
              <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14, marginBottom: 20 }}>
                {[
                  { label: "Pré-compras em andamento", value: sdEmAndamento.length, sub: "Tickets de pré-compra sem RC aberta", color: "var(--gaq-blue)", click: null },
                  { label: "Pré-compras concluídas sem RC", value: sdConcluidosSemRC.length, sub: "Clique para ver a lista", color: "var(--gaq-orange)", click: () => setShowSDList(s => !s) },
                  { label: "RCs em Andamento", value: rcEmAndamento.length, sub: "Processos com RC aberta", color: "var(--gaq-green)", click: null },
                ].map(kpi => (
                  <div key={kpi.label} onClick={kpi.click} style={{
                    background: "var(--gaq-surface)", borderRadius: "var(--gaq-r-lg)", padding: "20px 22px",
                    boxShadow: "var(--gaq-shadow-1)", border: "1px solid var(--gaq-line)",
                    borderTop: `3px solid ${kpi.color}`, cursor: kpi.click ? "pointer" : "default",
                    transition: "transform .15s, box-shadow .15s",
                  }}
                    onMouseEnter={e => { if(kpi.click) { e.currentTarget.style.transform = "translateY(-2px)"; e.currentTarget.style.boxShadow = "var(--gaq-shadow-2)"; }}}
                    onMouseLeave={e => { e.currentTarget.style.transform = "translateY(0)"; e.currentTarget.style.boxShadow = "var(--gaq-shadow-1)"; }}>
                    <div style={{ fontSize: 10.5, fontWeight: 600, color: kpi.color, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 8 }}>{kpi.label}</div>
                    <div style={{ fontSize: 36, fontWeight: 700, color: kpi.color, fontVariantNumeric: "tabular-nums", lineHeight: 1 }}>{kpi.value}</div>
                    <div style={{ fontSize: 11, color: "var(--gaq-text-3)", marginTop: 6 }}>{kpi.sub}</div>
                  </div>
                ))}
              </div>

              {/* Lista pré-compras concluídas sem RC (expandível) */}
              {showSDList && sdConcluidosSemRC.length > 0 && (
                <div className="gaq-card" style={{ marginBottom: 20, borderTop: "3px solid var(--gaq-orange)" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                    <div style={{ fontWeight: 600, fontSize: 14, color: "var(--gaq-orange)" }}>Pré-compras concluídas pela GAQ sem RC ({sdConcluidosSemRC.length})</div>
                    <button className="gaq-icon-btn" onClick={() => setShowSDList(false)}><Icon name="x" size={14}/></button>
                  </div>
                  <div className="gaq-table-scroll"><table className="gaq-table" style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr>
                        <th style={thS}>Ticket pré-compra</th>
                        <th style={{ ...thS, maxWidth: 250 }}>Objeto</th>
                        <th style={thS}>Area</th>
                        <th style={thS}>Encerramento pré-compra</th>
                        <th style={{ padding: "8px 6px", textAlign: "center", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase" }}>Dias (pré-compra)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {sdConcluidosSemRC.sort((a, b) => (b.diasSD || 0) - (a.diasSD || 0)).map((r, i) => (
                        <tr key={r.ProcessKey || i} onClick={() => setSelProc(r)} style={{ cursor: "pointer", borderBottom: "1px solid var(--border2)" }}
                          onMouseEnter={e => e.currentTarget.style.background = "var(--hover)"}
                          onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                          <td style={{ padding: "8px 6px", fontWeight: 700, color: "#2e86c1" }}>{r.TicketSD || "—"}</td>
                          <td style={{ padding: "8px 6px", color: "var(--text)", maxWidth: 250, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={r.Objeto}>{(r.Objeto || "").slice(0, 70) || "—"}</td>
                          <td style={{ padding: "8px 6px", color: "var(--text2)" }}>{r["Área Requisitante"] || "—"}</td>
                          <td style={{ padding: "8px 6px", color: "var(--text2)" }}>{r.encSD ? r.encSD.toLocaleDateString("pt-BR") : "—"}</td>
                          <td style={{ padding: "8px 6px", textAlign: "center", fontWeight: 700, color: (r.diasSD || 0) > 30 ? "#c0392b" : "#e67e22" }}>{r.diasSD || 0}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table></div>
                </div>
              )}

              {/* BULLET: Top 10 RC mais antigas - NCL */}
              <div className="gaq-card" style={{ marginBottom: 16, borderTop: "3px solid var(--gaq-ncl)", padding: 0, overflow: "hidden" }}>
                <div onClick={() => setExpandNCL(s => !s)} style={{ padding: "16px 20px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <div style={{ fontWeight: 600, fontSize: 14, color: "var(--gaq-ncl)" }}>Top 10 RCs mais antigas — Dentro da NCL</div>
                    <div style={{ fontSize: 11, color: "var(--gaq-text-3)", marginTop: 3 }}>Processos em andamento na subarea NCL, ordenados por dias em aberto</div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span className="gaq-badge is-soft" style={{ fontSize: 12 }}>{top10NCL.length}</span>
                    <Icon name={expandNCL ? "chevU" : "chevD"} size={14} color="var(--gaq-text-3)" />
                  </div>
                </div>
                {expandNCL && (
                  <div style={{ padding: "0 20px 16px" }}>
                    {top10NCL.length === 0
                      ? <div style={{ color: "var(--gaq-text-3)", fontSize: 12, padding: 10 }}>Nenhuma RC em NCL encontrada.</div>
                      : renderTopTable(top10NCL)}
                  </div>
                )}
              </div>

              {/* BULLET: Top 10 RC mais antigas - NCL ou Scont (Lic + Suica) */}
              <div className="gaq-card" style={{ marginBottom: 16, borderTop: "3px solid var(--gaq-rsp)", padding: 0, overflow: "hidden" }}>
                <div onClick={() => setExpandNCLScont(s => !s)} style={{ padding: "16px 20px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <div style={{ fontWeight: 600, fontSize: 14, color: "var(--gaq-rsp)" }}>Top 10 RCs mais antigas — NCL + Scont</div>
                    <div style={{ fontSize: 11, color: "var(--gaq-text-3)", marginTop: 3 }}>Processos em andamento nas subareas NCL e Scont, ordenados por dias em aberto</div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span className="gaq-badge is-soft" style={{ fontSize: 12 }}>{top10NCLScont.length}</span>
                    <Icon name={expandNCLScont ? "chevU" : "chevD"} size={14} color="var(--gaq-text-3)" />
                  </div>
                </div>
                {expandNCLScont && (
                  <div style={{ padding: "0 20px 16px" }}>
                    {top10NCLScont.length === 0
                      ? <div style={{ color: "var(--gaq-text-3)", fontSize: 12, padding: 10 }}>Nenhuma RC em NCL/Scont encontrada.</div>
                      : renderTopTable(top10NCLScont)}
                  </div>
                )}
              </div>
            </div>
          );
        })());
}
