import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function ProjetosScreen({ base, dirAreaSel, isDiretor, loginArea, meta, projetoTag, renderBuscaPanel, setProjetoTag, setSelProc }) {
  return ((() => {
          const hoje0 = new Date(); hoje0.setHours(0,0,0,0);
          // Recorte por perfil: perfil de Área vê só processos da sua área;
          // perfil de Diretor com área selecionada idem. Admin/Comprador veem tudo.
          const areaFilter = loginArea || (isDiretor && dirAreaSel) || null;
          const baseScope = areaFilter
            ? base.filter(r => areaMatch(r["Área Requisitante"] || r.Area || r["ÁREA"] || "", areaFilter))
            : base;
          // Agrupar processos por tag (campo PROJETO / EVENTO / AÇÃO)
          const buckets = {};
          baseScope.forEach(r => {
            const t = (r["PROJETO / EVENTO / AÇÃO"] || "").toString().trim();
            if (!t) return;
            if (!buckets[t]) buckets[t] = [];
            buckets[t].push(r);
          });
          // Lista de tags presente NO RECORTE atual (não em toda a base)
          const tags = Object.keys(buckets).sort((a, b) => a.localeCompare(b, "pt-BR"));
          // Cor por proximidade da entrega
          const corPorEntrega = dias => {
            if (dias === null || dias === undefined) return "var(--gaq-text-3)";
            if (dias < 0)   return "var(--gaq-red)";      // vencida
            if (dias <= 7)  return "#ff3b30";             // urgente
            if (dias <= 30) return "var(--gaq-orange)";   // próxima
            return "var(--gaq-green)";                    // futura
          };
          const lista = tags.map(t => {
            const procs = buckets[t] || [];
            const emA = procs.filter(r => r.emA).length;
            const ca = procs.filter(r => r.isCanceled || r.isFailed).length;
            const co = procs.filter(r => r.isConcluded).length;
            const cor = TagsManager.getTagColor(t);
            // Próxima entrega = menor data de entrega FUTURA entre os em andamento;
            // se nenhuma futura, usa a menor entre as vencidas; senão null.
            const datasEntrega = procs
              .filter(r => r.emA && r.dataEntrega instanceof Date)
              .map(r => r.dataEntrega)
              .sort((a, b) => a - b);
            const proxEntrega = datasEntrega[0] || null;
            const diasProx = proxEntrega ? Math.round((proxEntrega - hoje0) / 86400000) : null;
            return { tag: t, procs, total: procs.length, emA, ca, co, cor, proxEntrega, diasProx };
          }).sort((a, b) => {
            // Ordena por data de entrega mais próxima (ascendente). Sem data vai pro final.
            if (a.proxEntrega && b.proxEntrega) return a.proxEntrega - b.proxEntrega;
            if (a.proxEntrega) return -1;
            if (b.proxEntrega) return 1;
            return b.total - a.total;
          });
          const tagSel = projetoTag && buckets[projetoTag] ? projetoTag : (lista[0] ? lista[0].tag : "");
          const sel = lista.find(x => x.tag === tagSel) || lista[0] || null;

          if (lista.length === 0) {
            return (
              <div className="anim-fade">
                {renderBuscaPanel()}
                <div className="gaq-card" style={{ padding: 30, textAlign: "center", color: "var(--gaq-text-3)" }}>
                  <div style={{ fontSize: 30, marginBottom: 8 }}>📁</div>
                  <div style={{ fontWeight: 600 }}>
                    {areaFilter ? `Nenhum projeto/evento informado na área ${areaFilter}.` : "Nenhum projeto/evento informado na base."}
                  </div>
                  <div style={{ fontSize: 12, marginTop: 6 }}>
                    Preencha o campo <b>PROJETO / EVENTO / AÇÃO</b> {areaFilter ? `nos processos da área ${areaFilter}` : "no dados.js"} para acompanhar processos agrupados por iniciativa.
                  </div>
                </div>
              </div>
            );
          }

          return (
            <div className="anim-fade">
              {renderBuscaPanel()}
              <div style={{ marginBottom: 14 }}>
                <div className="gaq-h1" style={{ marginBottom: 4 }}>Acompanhamento de Projetos</div>
                <div className="gaq-body" style={{ color: "var(--gaq-text-3)" }}>
                  Processos agrupados pelo campo <b>PROJETO / EVENTO / AÇÃO</b>
                  {areaFilter ? <> · recorte da área <b>{areaFilter}</b></> : <> · todas as áreas</>}
                  {" · "}{lista.length} projeto(s)
                </div>
              </div>

              <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "minmax(260px, 340px) 1fr", gap: 16, alignItems: "start" }}>

                {/* Lista lateral de projetos */}
                <div className="gaq-card" style={{ padding: 14, maxHeight: "calc(100vh - 200px)", overflowY: "auto" }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "var(--gaq-text-3)", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 8, padding: "0 4px", display: "flex", justifyContent: "space-between" }}>
                    <span>Projetos / Eventos</span>
                    <span style={{ textTransform: "none", letterSpacing: 0, fontWeight: 500 }}>por entrega ↑</span>
                  </div>
                  {lista.map(x => {
                    const corEnt = corPorEntrega(x.diasProx);
                    const labelEnt = x.proxEntrega
                      ? `${x.proxEntrega.toLocaleDateString("pt-BR")} · ${x.diasProx < 0 ? `${-x.diasProx}d atrasada` : x.diasProx === 0 ? "hoje" : `em ${x.diasProx}d`}`
                      : "sem data de entrega";
                    return (
                      <button key={x.tag} onClick={() => setProjetoTag(x.tag)}
                        className={`gaq-side-item${tagSel === x.tag ? " is-active" : ""}`}
                        style={{ width: "100%", marginBottom: 4, textAlign: "left", flexDirection: "column", alignItems: "stretch", gap: 4, padding: "8px 10px" }}>
                        <span style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
                          <span style={{ display: "flex", alignItems: "center", gap: 8, minWidth: 0 }}>
                            <span style={{ width: 8, height: 8, borderRadius: 999, background: x.cor, flexShrink: 0 }}/>
                            <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontWeight: tagSel === x.tag ? 700 : 500 }}>{x.tag}</span>
                          </span>
                          <span className="gaq-badge is-soft">{x.total}</span>
                        </span>
                        <span style={{ fontSize: 10.5, color: corEnt, fontWeight: 600, paddingLeft: 16 }}>
                          {x.proxEntrega ? "⏱" : "—"} {labelEnt}
                        </span>
                      </button>
                    );
                  })}
                </div>

                {/* Detalhe do projeto selecionado */}
                {sel && <div>
                  <div className="gaq-card" style={{ padding: 18, marginBottom: 14 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, flexWrap: "wrap" }}>
                      <div style={{ minWidth: 0 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                          <span style={{ width: 10, height: 10, borderRadius: 999, background: sel.cor }}/>
                          <div className="gaq-h2" style={{ margin: 0 }}>{sel.tag}</div>
                        </div>
                        <div className="gaq-meta">{sel.total} processo(s) · {sel.emA} em andamento · {sel.ca} cancelado(s) · {sel.co} concluído(s)</div>
                        {sel.proxEntrega && (
                          <div style={{ marginTop: 6, fontSize: 12, fontWeight: 600, color: corPorEntrega(sel.diasProx) }}>
                            ⏱ Entrega mais próxima: {sel.proxEntrega.toLocaleDateString("pt-BR")} ({sel.diasProx < 0 ? `${-sel.diasProx}d atrasada` : sel.diasProx === 0 ? "hoje" : `em ${sel.diasProx} dias`})
                          </div>
                        )}
                      </div>
                      <button className="gaq-btn is-primary" onClick={() => {
                        const html = gerarRelatorioPorTag({ tag: sel.tag, processos: sel.procs, meta });
                        const win = window.open("", "_blank");
                        if (win) { win.document.write(html); win.document.close(); }
                      }}>
                        <Icon name="doc" size={14}/> Exportar PDF
                      </button>
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(120px, 1fr))", gap: 10, marginTop: 14 }}>
                      <div style={{ background: "var(--gaq-bg-2)", borderRadius: 8, padding: "10px 12px", borderTop: "3px solid var(--gaq-blue)" }}>
                        <div style={{ fontSize: 22, fontWeight: 800, color: "var(--gaq-blue)" }}>{sel.total}</div>
                        <div style={{ fontSize: 10, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase" }}>Total</div>
                      </div>
                      <div style={{ background: "var(--gaq-bg-2)", borderRadius: 8, padding: "10px 12px", borderTop: "3px solid var(--gaq-ncl)" }}>
                        <div style={{ fontSize: 22, fontWeight: 800, color: "var(--gaq-ncl)" }}>{sel.emA}</div>
                        <div style={{ fontSize: 10, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase" }}>Em andamento</div>
                      </div>
                      <div style={{ background: "var(--gaq-bg-2)", borderRadius: 8, padding: "10px 12px", borderTop: "3px solid var(--gaq-red)" }}>
                        <div style={{ fontSize: 22, fontWeight: 800, color: "var(--gaq-red)" }}>{sel.ca}</div>
                        <div style={{ fontSize: 10, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase" }}>Cancelados</div>
                      </div>
                      <div style={{ background: "var(--gaq-bg-2)", borderRadius: 8, padding: "10px 12px", borderTop: "3px solid var(--gaq-green)" }}>
                        <div style={{ fontSize: 22, fontWeight: 800, color: "var(--gaq-green)" }}>{sel.co}</div>
                        <div style={{ fontSize: 10, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase" }}>Concluídos</div>
                      </div>
                    </div>
                  </div>

                  {/* Lista dos processos da tag */}
                  <div className="gaq-card" style={{ padding: 0, overflow: "hidden" }}>
                    <div style={{ padding: "12px 18px", borderBottom: "1px solid var(--gaq-line)", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ fontWeight: 700, fontSize: 13 }}>Processos vinculados</div>
                      <span className="gaq-meta">Clique para ver detalhes</span>
                    </div>
                    <div style={{ maxHeight: "calc(100vh - 420px)", overflowY: "auto" }}>
                      <div className="gaq-table-scroll"><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                        <thead style={{ position: "sticky", top: 0, background: "var(--gaq-bg-2)", zIndex: 1 }}>
                          <tr>
                            <th style={{ padding: "8px 10px", textAlign: "left", fontSize: 10.5, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase", borderBottom: "1px solid var(--gaq-line)" }}>Entrega ↑</th>
                            <th style={{ padding: "8px 10px", textAlign: "left", fontSize: 10.5, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase", borderBottom: "1px solid var(--gaq-line)" }}>Nº RC</th>
                            <th style={{ padding: "8px 10px", textAlign: "left", fontSize: 10.5, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase", borderBottom: "1px solid var(--gaq-line)" }}>Ticket</th>
                            <th style={{ padding: "8px 10px", textAlign: "left", fontSize: 10.5, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase", borderBottom: "1px solid var(--gaq-line)" }}>Área</th>
                            <th style={{ padding: "8px 10px", textAlign: "left", fontSize: 10.5, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase", borderBottom: "1px solid var(--gaq-line)" }}>Responsável</th>
                            <th style={{ padding: "8px 10px", textAlign: "left", fontSize: 10.5, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase", borderBottom: "1px solid var(--gaq-line)" }}>Objeto</th>
                            <th style={{ padding: "8px 10px", textAlign: "left", fontSize: 10.5, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase", borderBottom: "1px solid var(--gaq-line)" }}>Status</th>
                            <th style={{ padding: "8px 10px", textAlign: "right", fontSize: 10.5, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase", borderBottom: "1px solid var(--gaq-line)" }}>d.u.</th>
                          </tr>
                        </thead>
                        <tbody>
                          {sel.procs
                            .slice()
                            .sort((a, b) => {
                              // Entrega mais próxima primeiro; sem data vai pro fim.
                              const aD = a.dataEntrega instanceof Date ? a.dataEntrega.getTime() : Infinity;
                              const bD = b.dataEntrega instanceof Date ? b.dataEntrega.getTime() : Infinity;
                              if (aD !== bD) return aD - bD;
                              return (b.diasTotais || 0) - (a.diasTotais || 0);
                            })
                            .map(r => {
                              const resp = r.Comprador || r.Avaliador || r.Pregoeiro || r.cplResp || "—";
                              const stCol = r.isCanceled || r.isFailed ? "var(--gaq-red)" : r.isConcluded ? "var(--gaq-green)" : r.emA ? "var(--gaq-blue)" : "var(--gaq-text-3)";
                              const ent = r.dataEntrega instanceof Date ? r.dataEntrega : null;
                              const dE = ent ? Math.round((ent - hoje0) / 86400000) : null;
                              const corE = corPorEntrega(dE);
                              const labelE = ent
                                ? `${ent.toLocaleDateString("pt-BR")}${r.emA ? ` · ${dE < 0 ? `${-dE}d atrasada` : dE === 0 ? "hoje" : `em ${dE}d`}` : ""}`
                                : "—";
                              return (
                                <tr key={r.ProcessKey} onClick={() => setSelProc(r)}
                                  style={{ cursor: "pointer", borderBottom: "1px solid var(--gaq-line)" }}>
                                  <td style={{ padding: "9px 10px", borderLeft: ent && r.emA ? `3px solid ${corE}` : "3px solid transparent" }}>
                                    <span style={{ fontSize: 11.5, fontWeight: 700, color: ent ? corE : "var(--gaq-text-3)" }}>{labelE}</span>
                                  </td>
                                  <td style={{ padding: "9px 10px", fontFamily: "var(--gaq-mono)", fontWeight: 700 }}>{r.NumRC || "—"}</td>
                                  <td style={{ padding: "9px 10px", fontFamily: "var(--gaq-mono)", fontSize: 11 }}>{r.TicketSD || "—"}</td>
                                  <td style={{ padding: "9px 10px" }}>{r["Área Requisitante"] || "—"}</td>
                                  <td style={{ padding: "9px 10px" }}>{resp}</td>
                                  <td style={{ padding: "9px 10px", maxWidth: 320, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.Objeto || "—"}</td>
                                  <td style={{ padding: "9px 10px" }}>
                                    <span style={{ fontSize: 11, fontWeight: 700, color: stCol, background: stCol + "18", borderRadius: 999, padding: "2px 8px" }}>{r.status || "—"}</span>
                                  </td>
                                  <td style={{ padding: "9px 10px", textAlign: "right", fontWeight: 700, color: (r.diasTotais || 0) > 50 ? "var(--gaq-red)" : "var(--gaq-text)" }}>{r.diasTotais || 0}</td>
                                </tr>
                              );
                            })}
                        </tbody>
                      </table></div>
                    </div>
                  </div>
                </div>}
              </div>
            </div>
          );
        })());
}
