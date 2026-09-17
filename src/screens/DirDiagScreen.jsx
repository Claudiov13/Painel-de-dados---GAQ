import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function DirDiagScreen({ dirAreaSel, diretorDiag, setDirTLDrill, setSelProc }) {
  return ((() => {
          const diag = diretorDiag;
          const scopeLabel = dirAreaSel ? `Área: ${dirAreaSel}` : "Todas as áreas";
          const topPressao = diag.topCriticos.length > 0 ? diag.topCriticos : diag.classificados.slice(0, 8);
          const flowMax = Math.max(1, ...diag.flowByYear.map(r => Math.max(r.entradas, r.concluidos, r.encerrados)));
          const openDrill = (title, procs) => { if (procs && procs.length > 0) setDirTLDrill({ title, procs }); };
          const formatSignedDays = (n, empty = "—") => {
            if (n == null) return empty;
            if (n === 0) return "0 d.u.";
            return `${n > 0 ? "+" : ""}${n} d.u.`;
          };
          const modChips = (mods, color, emptyText) => (
            mods.length === 0
              ? <div style={{ fontSize: 10, color: "var(--text3)" }}>{emptyText}</div>
              : <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {mods.map(m => (
                    <span key={m.name} style={{ background: color + "15", color, border: `1px solid ${color}33`, borderRadius: 999, padding: "3px 9px", fontSize: 10, fontWeight: 700 }}>
                      {m.name} · {m.count}
                    </span>
                  ))}
                </div>
          );

          const branches = [
            {
              key: "critico",
              icon: "🔴",
              title: "Casos Críticos",
              color: "#c0392b",
              bg: "#fff5f4",
              rows: diag.criticos,
              subtitle: "Atraso geral, entrega vencida ou janela de entrega menor que o prazo restante.",
              mods: diag.criticosTopMods,
              items: [
                        { label: "SLA Vencido", value: diag.atrasoPrazo.length, rows: diag.atrasoPrazo, drillTitle: `Diagnóstico Diretoria — SLA Vencido (${scopeLabel})` },
                { label: "Entrega vencida", value: diag.entregaVencida.length, rows: diag.entregaVencida, drillTitle: `Diagnóstico Diretoria — Entrega vencida (${scopeLabel})` },
                        { label: "Janela comprometida", value: diag.janelaComprometida.length, rows: diag.janelaComprometida, drillTitle: `Diagnóstico Diretoria — Janela comprometida (${scopeLabel})` },
              ],
              drillTitle: `Diagnóstico Diretoria — Casos Críticos (${scopeLabel})`,
            },
            {
              key: "atencao",
              icon: "🟠",
              title: "Casos em Atenção",
              color: "#e67e22",
              bg: "#fff8f1",
              rows: diag.atencao,
              subtitle: "Quase-críticos. Disparam quando alguma margem (SLA ou entrega) está ≤ 3 d.u., OU quando SLA e entrega estão ambos ≤ 8 d.u. de margem, OU quando a projeção atual já estoura o SLA.",
              mods: diag.atencaoTopMods,
              items: [
                { label: "Margem crítica (≤ 3 d.u.)", value: diag.atencaoForte.length, rows: diag.atencaoForte, drillTitle: `Diretoria — Atenção · margem crítica (${scopeLabel})` },
                { label: "Pressão dupla (≤ 8 d.u.)", value: diag.atencaoCombinada.length, rows: diag.atencaoCombinada, drillTitle: `Diretoria — Atenção · pressão dupla (${scopeLabel})` },
                { label: "Score da área", value: diag.mediaScore },
              ],
              drillTitle: `Diagnóstico Diretoria — Casos em Atenção (${scopeLabel})`,
            },
            {
              key: "controlado",
              icon: "🟢",
              title: "Casos Controlados",
              color: "#1e8449",
              bg: "#f3fbf5",
              rows: diag.controlados,
              subtitle: "Pipeline sem pressão imediata entre SLA da modalidade e data prevista de entrega.",
              mods: diag.controladosTopMods,
              items: [
                { label: "Base monitorada", value: diag.classificados.length },
                { label: "Consumo médio do prazo", value: `${diag.mediaConsumoPrazo}%` },
                        { label: "Score da área", value: diag.mediaScore },
              ],
              drillTitle: `Diagnóstico Diretoria — Casos Controlados (${scopeLabel})`,
            },
          ];

          const branchCard = (branch) => (
            <div key={branch.key} style={{ position: "relative", marginBottom: 14 }}>
              <div style={{ position: "absolute", left: -16, top: 28, width: 16, height: 2, background: "#d6dde8" }} />
              <div style={{ background: branch.bg, border: `1px solid ${branch.color}22`, borderLeft: `5px solid ${branch.color}`, borderRadius: 16, padding: "16px 18px", boxShadow: "var(--shadow)" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, flexWrap: "wrap" }}>
                  <div>
                    <div style={{ fontSize: 15, fontWeight: 800, color: branch.color }}>{branch.icon} {branch.title}</div>
                    <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 4, maxWidth: 520 }}>{branch.subtitle}</div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontSize: 30, fontWeight: 900, color: branch.color, lineHeight: 1 }}>{branch.rows.length}</div>
                    <button onClick={() => openDrill(branch.drillTitle, branch.rows)} disabled={branch.rows.length === 0}
                      style={{ marginTop: 8, padding: "6px 12px", borderRadius: 8, border: `1px solid ${branch.color}44`, background: branch.rows.length > 0 ? branch.color : "#d0d7de", color: "#fff", cursor: branch.rows.length > 0 ? "pointer" : "not-allowed", fontSize: 11, fontWeight: 700 }}>
                      Ver processos
                    </button>
                  </div>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 8, marginTop: 14 }}>
                  {branch.items.map(item => (
                    <button key={item.label} onClick={() => openDrill(item.drillTitle || branch.drillTitle, item.rows || branch.rows)} disabled={!((item.rows || branch.rows) && (item.rows || branch.rows).length > 0)}
                      style={{ textAlign: "left", background: "#fff", borderRadius: 10, border: "1px solid var(--border2)", padding: "10px 12px", cursor: ((item.rows || branch.rows) && (item.rows || branch.rows).length > 0) ? "pointer" : "default" }}>
                      <div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase", letterSpacing: .35 }}>{item.label}</div>
                      <div style={{ fontSize: 22, fontWeight: 800, color: branch.color, marginTop: 2 }}>{item.value}</div>
                    </button>
                  ))}
                </div>
                <div style={{ marginTop: 12 }}>
                  {modChips(branch.mods, branch.color, "Sem predominância por modalidade neste grupo.")}
                </div>
              </div>
            </div>
          );

          return (
            <div className="anim-fade">
              <div style={{ background: "linear-gradient(135deg, #154360 0%, #1f618d 55%, #2874a6 100%)", borderRadius: 14, padding: "22px 28px", marginBottom: 20, color: "#fff", boxShadow: "0 6px 24px rgba(21,67,96,0.28)" }}>
                <div style={{ fontSize: 20, fontWeight: 800, marginBottom: 4 }}>Diagnóstico da Diretoria</div>
                <div style={{ fontSize: 12, opacity: .82 }}>Escopo: {scopeLabel} · Criticidade = consumo do prazo da modalidade + pressão da data de entrega prevista.</div>
              </div>

              <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 20 }}>
                <div onClick={() => openDrill(`Diagnóstico Diretoria — Em andamento (${scopeLabel})`, diag.emA)}
                  style={{ flex: 1, minWidth: 150, background: "var(--card)", borderRadius: 12, padding: "16px 18px", boxShadow: "var(--shadow)", borderTop: "4px solid #2e86c1", cursor: diag.emA.length > 0 ? "pointer" : "default" }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "#2e86c1", textTransform: "uppercase", letterSpacing: .5 }}>Em andamento</div>
                  <div style={{ fontSize: 32, fontWeight: 800, color: "#2e86c1" }}>{diag.emA.length}</div>
                  <div style={{ fontSize: 10, color: "var(--text3)" }}>Pipeline total da área</div>
                </div>
                <div onClick={() => openDrill(`Diagnóstico Diretoria — Processos em andamento com RC (${scopeLabel})`, diag.ativosComRC)}
                  style={{ flex: 1, minWidth: 150, background: "var(--card)", borderRadius: 12, padding: "16px 18px", boxShadow: "var(--shadow)", borderTop: "4px solid #154360", cursor: diag.ativosComRC.length > 0 ? "pointer" : "default" }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "#154360", textTransform: "uppercase", letterSpacing: .5 }}>Com RC no radar</div>
                  <div style={{ fontSize: 32, fontWeight: 800, color: "#154360" }}>{diag.ativosComRC.length}</div>
                  <div style={{ fontSize: 10, color: "var(--text3)" }}>Base com SLA por modalidade</div>
                </div>
                <div onClick={() => openDrill(`Diagnóstico Diretoria — Casos Críticos (${scopeLabel})`, diag.criticos)}
                  style={{ flex: 1, minWidth: 150, background: "var(--card)", borderRadius: 12, padding: "16px 18px", boxShadow: "var(--shadow)", borderTop: "4px solid #c0392b", cursor: diag.criticos.length > 0 ? "pointer" : "default" }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "#c0392b", textTransform: "uppercase", letterSpacing: .5 }}>Críticos</div>
                  <div style={{ fontSize: 32, fontWeight: 800, color: "#c0392b" }}>{diag.criticos.length}</div>
                  <div style={{ fontSize: 10, color: "var(--text3)" }}>Ação prioritária da Diretoria</div>
                </div>
                <div onClick={() => openDrill(`Diagnóstico Diretoria — Base classificada (${scopeLabel})`, diag.classificados)}
                  style={{ flex: 1, minWidth: 150, background: "var(--card)", borderRadius: 12, padding: "16px 18px", boxShadow: "var(--shadow)", borderTop: "4px solid #8e44ad", cursor: diag.classificados.length > 0 ? "pointer" : "default" }}>
                        <div style={{ fontSize: 10, fontWeight: 700, color: "#8e44ad", textTransform: "uppercase", letterSpacing: .5 }}>Score da área</div>
                  <div style={{ fontSize: 32, fontWeight: 800, color: "#8e44ad" }}>{diag.mediaScore}</div>
                  <div style={{ fontSize: 10, color: "var(--text3)" }}>Pressão combinada do pipeline</div>
                </div>
                <div onClick={() => openDrill(`Diagnóstico Diretoria — Pré-compra em triagem (${scopeLabel})`, diag.sdTriagem)}
                  style={{ flex: 1, minWidth: 150, background: "var(--card)", borderRadius: 12, padding: "16px 18px", boxShadow: "var(--shadow)", borderTop: "4px solid #16a085", cursor: diag.sdTriagem.length > 0 ? "pointer" : "default" }}>
                  <div style={{ fontSize: 10, fontWeight: 700, color: "#16a085", textTransform: "uppercase", letterSpacing: .5 }}>Pré-compra sem RC</div>
                  <div style={{ fontSize: 32, fontWeight: 800, color: "#16a085" }}>{diag.sdTriagem.length}</div>
                  <div style={{ fontSize: 10, color: "var(--text3)" }}>Triagem fora do SLA da modalidade</div>
                </div>
              </div>

              <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "minmax(0, 1.35fr) minmax(280px, .9fr)", gap: 18, alignItems: "start", marginBottom: 20 }}>
                <div style={{ background: "var(--card)", borderRadius: 16, padding: "18px 20px", boxShadow: "var(--shadow)" }}>
                  <div style={{ fontWeight: 800, fontSize: 15, color: "#154360", marginBottom: 4 }}>Árvore de Diagnóstico</div>
                  <div style={{ fontSize: 11, color: "var(--text3)", marginBottom: 14 }}>Leitura hierárquica dos processos com RC, comparando prazo da modalidade com a entrega prevista.</div>
                  <div onClick={() => openDrill(`Diagnóstico Diretoria — Processos em andamento com RC (${scopeLabel})`, diag.ativosComRC)}
                    style={{ background: "#f4f8fc", border: "1px solid #dce6f0", borderRadius: 14, padding: "14px 16px", marginBottom: 14, cursor: diag.ativosComRC.length > 0 ? "pointer" : "default" }}>
                    <div style={{ fontSize: 10, fontWeight: 700, color: "#1f618d", textTransform: "uppercase", letterSpacing: .45 }}>Raiz</div>
                    <div style={{ fontSize: 18, fontWeight: 800, color: "#154360", marginTop: 4 }}>Processos em andamento com RC</div>
                    <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 4 }}>{diag.ativosComRC.length} processo(s) avaliados · consumo médio do prazo {diag.mediaConsumoPrazo}%.</div>
                  </div>
                  <div style={{ position: "relative", paddingLeft: 26 }}>
                    <div style={{ position: "absolute", left: 10, top: 0, bottom: 8, width: 2, background: "linear-gradient(180deg, #d6dde8 0%, #eef3f8 100%)" }} />
                    {branches.map(branchCard)}
                  </div>
                </div>

                <div style={{ display: "grid", gap: 14 }}>
                  <div style={{ background: "var(--card)", borderRadius: 16, padding: "16px 18px", boxShadow: "var(--shadow)", borderTop: "4px solid #e67e22" }}>
                    <div style={{ fontSize: 13, fontWeight: 800, color: "#e67e22", marginBottom: 10 }}>Radar de Entrega</div>
                    <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                      {[
                        { label: "Entrega vencida", value: diag.entregaVencida.length, color: "#c0392b", onClick: () => openDrill(`Diagnóstico Diretoria — Entrega vencida (${scopeLabel})`, diag.entregaVencida) },
                        { label: "Janela comprometida", value: diag.janelaComprometida.length, color: "#d35400", onClick: () => openDrill(`Diagnóstico Diretoria — Janela comprometida (${scopeLabel})`, diag.janelaComprometida) },
                        { label: "Entrega ≤ 20 d.u.", value: diag.entregaCurta.length, color: "#f39c12", onClick: () => openDrill(`Diagnóstico Diretoria — Entrega curta (${scopeLabel})`, diag.entregaCurta) },
                        { label: "Sem data de entrega", value: diag.semDataEntrega.length, color: "#7f8c8d", onClick: () => openDrill(`Diagnóstico Diretoria — Sem data de entrega (${scopeLabel})`, diag.semDataEntrega) },
                      ].map(item => (
                        <button key={item.label} onClick={item.onClick} disabled={item.value === 0}
                          style={{ textAlign: "left", borderRadius: 12, border: `1px solid ${item.color}33`, background: item.value > 0 ? item.color + "12" : "#f7f9fb", padding: "10px 12px", cursor: item.value > 0 ? "pointer" : "not-allowed" }}>
                          <div style={{ fontSize: 10, color: item.color, fontWeight: 700, textTransform: "uppercase", letterSpacing: .35 }}>{item.label}</div>
                          <div style={{ fontSize: 24, fontWeight: 800, color: item.color, marginTop: 4 }}>{item.value}</div>
                        </button>
                      ))}
                    </div>
                  </div>

                  <div style={{ background: "var(--card)", borderRadius: 16, padding: "16px 18px", boxShadow: "var(--shadow)", borderTop: "4px solid #16a085" }}>
                    <div style={{ fontSize: 13, fontWeight: 800, color: "#16a085", marginBottom: 10 }}>Triagem pré-compra</div>
                    <div style={{ fontSize: 11, color: "var(--text3)", marginBottom: 10 }}>Processos ainda sem RC não entram no score executivo, mas ficam visíveis para não perder o SLA da pré-compra.</div>
                    <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                      <div onClick={() => openDrill(`Diagnóstico Diretoria — Pré-compra sem RC (${scopeLabel})`, diag.sdTriagem)}
                        style={{ borderRadius: 12, border: "1px solid #16a08533", background: "#16a08512", padding: "10px 12px", cursor: diag.sdTriagem.length > 0 ? "pointer" : "default" }}>
                        <div style={{ fontSize: 10, color: "#16a085", fontWeight: 700, textTransform: "uppercase" }}>Pré-compra sem RC</div>
                        <div style={{ fontSize: 24, fontWeight: 800, color: "#16a085", marginTop: 4 }}>{diag.sdTriagem.length}</div>
                      </div>
                      <div onClick={() => openDrill(`Diagnóstico Diretoria — Pré-compra acima do prazo (${scopeLabel})`, diag.sdTriagem.filter(r => r.atrasoSD))}
                        style={{ borderRadius: 12, border: "1px solid #c0392b33", background: "#c0392b12", padding: "10px 12px", cursor: diag.sdTriagem.filter(r => r.atrasoSD).length > 0 ? "pointer" : "default" }}>
                        <div style={{ fontSize: 10, color: "#c0392b", fontWeight: 700, textTransform: "uppercase" }}>Pré-compra &gt; {PRAZO_SD} d.u.</div>
                        <div style={{ fontSize: 24, fontWeight: 800, color: "#c0392b", marginTop: 4 }}>{diag.sdTriagem.filter(r => r.atrasoSD).length}</div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              <div style={{ background: "var(--card)", borderRadius: 16, padding: "18px 20px", boxShadow: "var(--shadow)", marginBottom: 20 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
                  <div>
                    <div style={{ fontWeight: 800, fontSize: 15, color: "#154360" }}>Top Casos Mais Pressionados</div>
                    <div style={{ fontSize: 11, color: "var(--text3)", marginTop: 3 }}>Ordenação pelo score executivo da Diretoria, já combinando SLA da modalidade com a data de entrega.</div>
                  </div>
                  {topPressao.length > 0 && <button onClick={() => openDrill(`Diagnóstico Diretoria — Top Pressão (${scopeLabel})`, topPressao)}
                    style={{ padding: "7px 12px", borderRadius: 8, border: "1px solid #15436033", background: "#154360", color: "#fff", cursor: "pointer", fontSize: 11, fontWeight: 700 }}>
                    Abrir lista
                  </button>}
                </div>
                {topPressao.length === 0 ? <div style={{ fontSize: 12, color: "var(--text3)", padding: "12px 0" }}>Nenhum processo com RC em andamento para este âmbito.</div> : (
                  <div style={{ overflowX: "auto" }}>
                    <div className="gaq-table-scroll"><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                      <thead>
                        <tr style={{ background: "var(--card2)", borderBottom: "2px solid var(--border2)" }}>
                          <th style={{ padding: "8px 6px", textAlign: "left", fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Score</th>
                          <th style={{ padding: "8px 6px", textAlign: "left", fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Processo</th>
                          <th style={{ padding: "8px 6px", textAlign: "left", fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Modalidade</th>
                          <th style={{ padding: "8px 6px", textAlign: "left", fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Área</th>
                          <th style={{ padding: "8px 6px", textAlign: "center", fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>% Prazo</th>
                          <th style={{ padding: "8px 6px", textAlign: "center", fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Entrega</th>
                          <th style={{ padding: "8px 6px", textAlign: "center", fontSize: 10, color: "var(--text3)", textTransform: "uppercase" }}>Folga</th>
                        </tr>
                      </thead>
                      <tbody>
                        {topPressao.map((r, i) => (
                          <tr key={r.ProcessKey || i} onClick={() => setSelProc(r)} style={{ borderBottom: "1px solid var(--border2)", cursor: "pointer" }}
                            onMouseEnter={e => e.currentTarget.style.background = "var(--hover)"}
                            onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                            <td style={{ padding: "8px 6px" }}>
                              <span style={{ background: (r._diag.score >= 80 ? "#c0392b" : r._diag.score >= 55 ? "#e67e22" : "#1e8449"), color: "#fff", borderRadius: 999, padding: "4px 10px", fontWeight: 800 }}>{r._diag.score}</span>
                            </td>
                            <td style={{ padding: "8px 6px", fontWeight: 700, color: "#2e86c1" }}>{r.NumRC || r.TicketSD || "—"}</td>
                            <td style={{ padding: "8px 6px", color: "var(--text)" }}>{r.Modalidade || "—"}</td>
                            <td style={{ padding: "8px 6px", color: "var(--text2)" }}>{r["Área Requisitante"] || "—"}</td>
                            <td style={{ padding: "8px 6px", textAlign: "center", fontWeight: 800, color: r._diag.pctPrazo >= 1 ? "#c0392b" : r._diag.pctPrazo >= 0.7 ? "#e67e22" : "#1e8449" }}>{r._diag.pctPrazoLabel}%</td>
                            <td style={{ padding: "8px 6px", textAlign: "center", fontWeight: 700, color: r._diag.entregaVencida ? "#c0392b" : (r._diag.entregaCurta ? "#e67e22" : "#1e8449") }}>{formatSignedDays(r._diag.diasEntregaUteis, "Sem data")}</td>
                            <td style={{ padding: "8px 6px", textAlign: "center", fontWeight: 700, color: r._diag.janelaComprometida ? "#c0392b" : "#1e8449" }}>{formatSignedDays(r._diag.folgaEntrega, "Sem data")}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table></div>
                  </div>
                )}
              </div>

              <div style={{ background: "var(--card)", borderRadius: 16, padding: "18px 20px", boxShadow: "var(--shadow)" }}>
                <div style={{ fontWeight: 800, fontSize: 15, color: "#154360", marginBottom: 4 }}>Entradas x Conclusões</div>
                <div style={{ fontSize: 11, color: "var(--text3)", marginBottom: 14 }}>Leitura anual do que está entrando na área e do que efetivamente está sendo concluído/encerrado.</div>
                <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 14 }}>
                  <div style={{ flex: 1, minWidth: 140, borderRadius: 12, padding: "12px 14px", background: "#2e86c112", border: "1px solid #2e86c133" }}>
                    <div style={{ fontSize: 10, color: "#2e86c1", fontWeight: 700, textTransform: "uppercase" }}>Entradas {diag.flowAtual.ano}</div>
                    <div style={{ fontSize: 28, fontWeight: 800, color: "#2e86c1", marginTop: 2 }}>{diag.flowAtual.entradas}</div>
                  </div>
                  <div style={{ flex: 1, minWidth: 140, borderRadius: 12, padding: "12px 14px", background: "#27ae6012", border: "1px solid #27ae6033" }}>
                    <div style={{ fontSize: 10, color: "#27ae60", fontWeight: 700, textTransform: "uppercase" }}>Concluídos {diag.flowAtual.ano}</div>
                    <div style={{ fontSize: 28, fontWeight: 800, color: "#27ae60", marginTop: 2 }}>{diag.flowAtual.concluidos}</div>
                  </div>
                  <div style={{ flex: 1, minWidth: 140, borderRadius: 12, padding: "12px 14px", background: "#8e44ad12", border: "1px solid #8e44ad33" }}>
                    <div style={{ fontSize: 10, color: "#8e44ad", fontWeight: 700, textTransform: "uppercase" }}>Encerrados {diag.flowAtual.ano}</div>
                    <div style={{ fontSize: 28, fontWeight: 800, color: "#8e44ad", marginTop: 2 }}>{diag.flowAtual.encerrados}</div>
                  </div>
                  <div style={{ flex: 1, minWidth: 140, borderRadius: 12, padding: "12px 14px", background: (diag.flowAtual.saldo > 0 ? "#f39c1212" : "#1e844912"), border: `1px solid ${diag.flowAtual.saldo > 0 ? "#f39c1233" : "#1e844933"}` }}>
                    <div style={{ fontSize: 10, color: diag.flowAtual.saldo > 0 ? "#d68910" : "#1e8449", fontWeight: 700, textTransform: "uppercase" }}>Saldo {diag.flowAtual.ano}</div>
                    <div style={{ fontSize: 28, fontWeight: 800, color: diag.flowAtual.saldo > 0 ? "#d68910" : "#1e8449", marginTop: 2 }}>{diag.flowAtual.saldo > 0 ? `+${diag.flowAtual.saldo}` : diag.flowAtual.saldo}</div>
                  </div>
                </div>
                {diag.flowByYear.length === 0 ? <div style={{ fontSize: 12, color: "var(--text3)" }}>Sem histórico suficiente para este escopo.</div> : (
                  <div style={{ display: "grid", gap: 10 }}>
                    {diag.flowByYear.map(row => (
                      <div key={row.ano} style={{ border: "1px solid var(--border2)", borderRadius: 14, padding: "12px 14px" }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap", marginBottom: 10 }}>
                          <div style={{ fontSize: 15, fontWeight: 800, color: "#154360" }}>{row.ano}</div>
                          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                            <button onClick={() => openDrill(`Entradas ${row.ano} (${scopeLabel})`, row.entradasList)} disabled={row.entradas === 0}
                              style={{ padding: "5px 10px", borderRadius: 999, border: "1px solid #2e86c133", background: "#2e86c112", color: "#2e86c1", cursor: row.entradas > 0 ? "pointer" : "not-allowed", fontSize: 11, fontWeight: 700 }}>
                              Entradas: {row.entradas}
                            </button>
                            <button onClick={() => openDrill(`Concluídos ${row.ano} (${scopeLabel})`, row.concluidosList)} disabled={row.concluidos === 0}
                              style={{ padding: "5px 10px", borderRadius: 999, border: "1px solid #27ae6033", background: "#27ae6012", color: "#27ae60", cursor: row.concluidos > 0 ? "pointer" : "not-allowed", fontSize: 11, fontWeight: 700 }}>
                              Concluídos: {row.concluidos}
                            </button>
                            <button onClick={() => openDrill(`Encerrados ${row.ano} (${scopeLabel})`, row.encerradosList)} disabled={row.encerrados === 0}
                              style={{ padding: "5px 10px", borderRadius: 999, border: "1px solid #8e44ad33", background: "#8e44ad12", color: "#8e44ad", cursor: row.encerrados > 0 ? "pointer" : "not-allowed", fontSize: 11, fontWeight: 700 }}>
                              Encerrados: {row.encerrados}
                            </button>
                          </div>
                        </div>
                        <div style={{ display: "grid", gap: 6 }}>
                          {[
                            { label: "Entradas", value: row.entradas, color: "#2e86c1" },
                            { label: "Concluídos", value: row.concluidos, color: "#27ae60" },
                            { label: "Encerrados", value: row.encerrados, color: "#8e44ad" },
                          ].map(bar => (
                            <div className="gaq-responsive-grid" key={bar.label} style={{ display: "grid", gridTemplateColumns: "92px 1fr 64px", gap: 8, alignItems: "center" }}>
                              <div style={{ fontSize: 11, color: "var(--text2)", fontWeight: 700 }}>{bar.label}</div>
                              <div style={{ background: "#eef3f8", borderRadius: 999, height: 10, overflow: "hidden" }}>
                                <div style={{ width: `${(bar.value / flowMax) * 100}%`, minWidth: bar.value > 0 ? 8 : 0, height: "100%", background: bar.color, borderRadius: 999 }} />
                              </div>
                              <div style={{ fontSize: 11, color: bar.color, fontWeight: 800, textAlign: "right" }}>{bar.value}</div>
                            </div>
                          ))}
                        </div>
                        <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginTop: 10, fontSize: 11, color: "var(--text3)" }}>
                          <span>Taxa de conclusão: <b style={{ color: "#27ae60" }}>{row.taxaConclusao}%</b></span>
                          <span>Taxa de encerramento: <b style={{ color: "#8e44ad" }}>{row.taxaEncerramento}%</b></span>
                          <span>Saldo: <b style={{ color: row.saldo > 0 ? "#d68910" : "#1e8449" }}>{row.saldo > 0 ? `+${row.saldo}` : row.saldo}</b></span>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          );
        })());
}
