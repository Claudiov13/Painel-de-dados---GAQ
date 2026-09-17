import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function ExecutivoScreen({ alertCPL, alertNCL, base, baseVis, drillDown, execData, executiveStrategic, fBack, fBase, loginArea, loginAreaDisp, loginUser, meta, notifs, pgDrillDown, pgDrillDownSz, phaseIntervals, rankCPLResp, rankGeralNCL, rankNCL, setAba, setDrillDown, setPgDrillDown, setPgDrillDownSz, setSelProc, showNotifPanel }) {
  return ((execData && executiveStrategic) ? <div>
          {drillDown && drillDown.sourceAba === "executivo" ? <DrillDownPanel
            drillDown={drillDown}
            defaultSourceAba="executivo"
            setDrillDown={setDrillDown}
            setAba={setAba}
            setSelProc={setSelProc}
            pgDrillDown={pgDrillDown}
            pgDrillDownSz={pgDrillDownSz}
            setPgDrillDown={setPgDrillDown}
            setPgDrillDownSz={setPgDrillDownSz}
            phaseIntervals={phaseIntervals}
          /> : <>
          <div style={{ background: "var(--gaq-surface)", borderRadius: "var(--gaq-r-xl)", padding: "24px 28px", marginBottom: 18, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, border: "1px solid var(--gaq-line)", boxShadow: "var(--gaq-shadow-2)" }}>
            <div>
              <div style={{ fontSize: 22, fontWeight: 700, color: "var(--gaq-text)", letterSpacing: "-0.02em" }}>Painel Executivo</div>
              <div style={{ fontSize: 12, color: "var(--gaq-text-3)", marginTop: 2 }}>Atualizado em {new Date().toLocaleString("pt-BR")} · {execData.tot} processos na base</div>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button className="gaq-btn no-print" onClick={() => window.print()} title="Imprimir / Salvar como PDF">
                <Icon name="download" size={14}/> Imprimir
              </button>
              <button className="gaq-btn is-primary no-print" onClick={() => {
                const html = gerarRelatorioHTML({ execData, fBase, rankNCL, rankGeralNCL, rankCPLResp, alertNCL, alertCPL, meta });
                const win = window.open("", "_blank");
                win.document.write(html);
                win.document.close();
              }} title="Gerar relatorio completo em nova janela">
                <Icon name="doc" size={14}/> PDF Mensal
              </button>
            </div>
          </div>

          {/* Print header (only visible when printing) */}
          <div className="print-only" style={{ textAlign: "center", marginBottom: 20, borderBottom: "3px solid #1a5276", paddingBottom: 12 }}>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#1a5276" }}>Sesc — Painel de Dados de Compras GAQ</div>
            <div style={{ fontSize: 11, color: "#666" }}>Relatório Executivo · Gerado em {new Date().toLocaleString("pt-BR")} · {execData.tot} processos</div>
          </div>

          <div style={{ marginBottom: 18 }}>
            <div className="gaq-eyebrow" style={{ marginBottom: 6 }}>Painel Executivo · recorte actual</div>
            <div className="gaq-h1" style={{ marginBottom: 4 }}>Indicadores estratégicos</div>
            <div className="sub">{executiveStrategic.ativos.toLocaleString("pt-BR")} processos em andamento · {executiveStrategic.problematicosRows.length.toLocaleString("pt-BR")} sob pressão estratégica</div>
          </div>

          <div className="gaq-card" style={{ padding: 28, marginBottom: 18, background: "linear-gradient(135deg, #0a1f3d 0%, #0071e3 100%)", color: "#fff", border: "none" }}>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 24 }}>
              {[
                {
                  label: "Total de processos",
                  value: executiveStrategic.total.toLocaleString("pt-BR"),
                  sub: "base filtrada do recorte atual",
                  onClick: () => setDrillDown({ title: "Executivo — Total de processos", data: fBase, color: "#85c1e9", sourceAba: "executivo" }),
                },
                {
                  label: "Processos ativos",
                  value: executiveStrategic.ativos.toLocaleString("pt-BR"),
                  sub: `${(executiveStrategic.slaVencidos || []).length.toLocaleString("pt-BR")} com SLA vencido`,
                  onClick: () => setDrillDown({ title: "Executivo — Processos em andamento", data: fBack, color: "#2e86c1", sourceAba: "executivo" }),
                },
                {
                  label: "Taxa problemática",
                  value: `${executiveStrategic.taxaProblematica}%`,
                  sub: `${executiveStrategic.problematicosRows.length.toLocaleString("pt-BR")} casos únicos sob pressão`,
                  onClick: () => setDrillDown({ title: "Executivo — Casos problemáticos", data: executiveStrategic.problematicosRows, color: "#ff9500", sourceAba: "executivo" }),
                },
                {
                  label: "Score da área",
                  value: executiveStrategic.scoreArea,
                  sub: executiveStrategic.scoreLabel,
                  onClick: () => setDrillDown({ title: "Executivo — Base classificada", data: executiveStrategic.classificados, color: executiveStrategic.scoreColor, sourceAba: "executivo" }),
                },
              ].map(card => (
                <div key={card.label} onClick={card.onClick} style={{ cursor: "pointer" }}>
                  <div style={{ fontSize: 11, fontWeight: 600, opacity: 0.72, letterSpacing: "0.04em", textTransform: "uppercase", marginBottom: 6 }}>{card.label}</div>
                  <div style={{ fontSize: 36, fontWeight: 700, letterSpacing: "-0.03em", lineHeight: 1, fontVariantNumeric: "tabular-nums" }}>{card.value}</div>
                  <div style={{ fontSize: 12, opacity: 0.8, marginTop: 8 }}>{card.sub}</div>
                </div>
              ))}
            </div>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))", gap: 12, marginBottom: 18 }}>
            {[
              { label: "Processos abertos", value: executiveStrategic.abertos, color: "#1f6feb", data: fBase, title: "Executivo — Processos abertos" },
              { label: "Concluídos", value: executiveStrategic.concluidos, color: "#16a34a", data: fBase.filter(r => r.isConcluded), title: "Executivo — Processos concluídos" },
              { label: "Em andamento", value: executiveStrategic.emAndamento, color: "#0f766e", data: fBack, title: "Executivo — Processos em andamento" },
              { label: "Cancelados", value: executiveStrategic.cancelados, color: "#ea580c", data: fBase.filter(r => r.isCanceled), title: "Executivo — Processos cancelados" },
              { label: "Fracassados", value: executiveStrategic.fracassados, color: "#7c3aed", data: fBase.filter(r => r.isFailed), title: "Executivo — Processos fracassados" },
            ].map(card => (
              <div
                key={card.label}
                onClick={() => setDrillDown({ title: card.title, data: card.data, color: card.color, sourceAba: "executivo" })}
                style={{
                  cursor: "pointer",
                  background: `linear-gradient(180deg, ${card.color}1A 0%, var(--gaq-surface) 100%)`,
                  border: `1px solid ${card.color}33`,
                  borderTop: `3px solid ${card.color}`,
                  borderRadius: 22,
                  padding: "16px 18px",
                  boxShadow: "var(--gaq-shadow-1)",
                }}
              >
                <div style={{ fontSize: 11, fontWeight: 700, color: "var(--gaq-text-3)", textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: 8 }}>{card.label}</div>
                <div className="gaq-num" style={{ fontSize: 30, fontWeight: 800, lineHeight: 1, color: card.color }}>{card.value.toLocaleString("pt-BR")}</div>
              </div>
            ))}
          </div>

          <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "minmax(0, 1.2fr) minmax(0, 1fr)", gap: 18, marginBottom: 18 }}>
            <div className="gaq-card" style={{ padding: 22 }}>
              <div className="gaq-section-h">
                <div>
                  <div className="gaq-h3">Processos por modalidade</div>
                  <div className="sub">{executiveStrategic.total.toLocaleString("pt-BR")} processos no recorte actual</div>
                </div>
                <button className="gaq-btn is-ghost" style={{ fontSize: 12 }} onClick={() => setDrillDown({ title: "Executivo — Todos os processos", data: fBase, color: "#2e86c1", sourceAba: "executivo" })}>Ver todos <Icon name="chevR" size={12}/></button>
              </div>
              {executiveStrategic.modalidades.length === 0
                ? <div style={{ color: "var(--gaq-text-3)", fontSize: 13, paddingTop: 8 }}>Sem processos para o recorte actual.</div>
                : (() => {
                    const maxModalidade = executiveStrategic.modalidades[0]?.count || 1;
                    return (
                      <div style={{ display: "flex", flexDirection: "column", gap: 7, marginTop: 8 }}>
                        {executiveStrategic.modalidades.map(item => {
                          const pct = executiveStrategic.total ? Math.round((item.count / executiveStrategic.total) * 100) : 0;
                          return (
                            <div className="gaq-responsive-grid" key={item.label} onClick={() => setDrillDown({ title: `Executivo — ${item.label}`, data: item.rows, color: "#2e86c1", sourceAba: "executivo" })}
                              style={{ display: "grid", gridTemplateColumns: "150px 1fr 80px", alignItems: "center", gap: 10, cursor: "pointer" }}>
                              <span style={{ fontSize: 12, fontWeight: 500, color: "var(--gaq-text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{item.label}</span>
                              <div style={{ height: 14, background: "rgba(0,0,0,0.04)", borderRadius: 4, position: "relative", overflow: "hidden" }}>
                                <div style={{ width: `${Math.max(6, (item.count / maxModalidade) * 100)}%`, height: "100%", background: "linear-gradient(90deg, #5dade2 0%, #2e86c1 100%)", borderRadius: 4 }} />
                              </div>
                              <span className="gaq-num" style={{ fontSize: 12.5, textAlign: "right", color: "var(--gaq-text-2)" }}><span style={{ fontWeight: 600 }}>{item.count.toLocaleString("pt-BR")}</span> <span style={{ fontWeight: 400, color: "var(--gaq-text-3)" }}>({pct}%)</span></span>
                            </div>
                          );
                        })}
                      </div>
                    );
                  })()}
            </div>

            <div className="gaq-card" style={{ padding: 22 }}>
              <div className="gaq-section-h">
                <div>
                  <div className="gaq-h3">Matriz de risco</div>
                  <div className="sub">{executiveStrategic.ativos.toLocaleString("pt-BR")} processos avaliados pela régua SLA + entrega</div>
                </div>
              </div>
              {(() => {
                const riskData = executiveStrategic.riskSegments.filter(seg => seg.value > 0);
                const totalRisk = executiveStrategic.riskSegments.reduce((sum, seg) => sum + seg.value, 0);
                if (!totalRisk) {
                  return <div style={{ color: "var(--gaq-text-3)", fontSize: 13, paddingTop: 8 }}>Sem processos em andamento para classificar.</div>;
                }
                const C = 2 * Math.PI * 60;
                let acc = 0;
                return (
                  <div style={{ display: "flex", alignItems: "center", gap: 18, marginTop: 6 }}>
                    <svg width="148" height="148" viewBox="0 0 148 148" style={{ flex: "none" }}>
                      {riskData.map((seg, idx) => {
                        const len = (seg.value / totalRisk) * C;
                        const dash = `${len} ${C - len}`;
                        const off = -acc;
                        acc += len;
                        return <circle key={`${seg.label}-${idx}`} cx="74" cy="74" r="60" fill="none" stroke={seg.color} strokeWidth="20" strokeDasharray={dash} strokeDashoffset={off} transform="rotate(-90 74 74)" strokeLinecap="butt" />;
                      })}
                      <text x="74" y="68" textAnchor="middle" fontSize="11" fill="var(--gaq-text-3)" fontWeight="600">Activos</text>
                      <text x="74" y="88" textAnchor="middle" fontSize="22" fill="var(--gaq-text)" fontWeight="700">{executiveStrategic.ativos.toLocaleString("pt-BR")}</text>
                    </svg>
                    <div style={{ flex: 1 }}>
                      {executiveStrategic.riskSegments.map(seg => {
                        const pct = totalRisk ? Math.round((seg.value / totalRisk) * 100) : 0;
                        const displayLabel = seg.label;
                        return (
                          <div key={seg.label} onClick={() => setDrillDown({ title: `Executivo — ${displayLabel}`, data: seg.rows, color: seg.color, sourceAba: "executivo" })}
                            style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 0", borderTop: "1px solid var(--gaq-line)", cursor: "pointer" }}>
                            <span style={{ width: 9, height: 9, borderRadius: 999, background: seg.color, flex: "none" }} />
                            <span style={{ fontSize: 12.5, flex: 1, color: "var(--gaq-text)" }}>{displayLabel}</span>
                            <span className="gaq-num" style={{ fontSize: 12.5, fontWeight: 700 }}>{seg.value.toLocaleString("pt-BR")}</span>
                            <span className="gaq-meta gaq-num" style={{ width: 36, textAlign: "right" }}>{pct}%</span>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })()}
            </div>
          </div>

          <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1.15fr)", gap: 18, marginBottom: 18 }}>
            <div className="gaq-card" style={{ padding: 22 }}>
              <div className="gaq-section-h">
                <div>
                  <div className="gaq-h3">Semáforo SLA e entrega</div>
                  <div className="sub">Pré-compra: 10 d.u. · Negociações diretas: 30 d.u. · Licitações: 90 d.u.</div>
                </div>
              </div>
              <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 12, marginTop: 10 }}>
                {executiveStrategic.semaforo.map(item => (
                  <div key={item.label} onClick={() => setDrillDown({ title: `Executivo — ${item.label}`, data: item.rows, color: item.color, sourceAba: "executivo" })}
                    style={{ border: `1px solid ${item.color}22`, borderTop: `3px solid ${item.color}`, borderRadius: 12, padding: "14px 14px 12px", background: "var(--gaq-bg-2)", cursor: "pointer" }}>
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginBottom: 6 }}>
                      <div style={{ fontSize: 12, fontWeight: 700, color: item.color, textTransform: "uppercase", letterSpacing: "0.03em" }}>{item.label}</div>
                      <div className="gaq-num" style={{ fontSize: 24, fontWeight: 800, color: item.color }}>{item.rows.length.toLocaleString("pt-BR")}</div>
                    </div>
                    <div style={{ fontSize: 12, color: "var(--gaq-text-2)", lineHeight: 1.45 }}>{item.detail}</div>
                  </div>
                ))}
              </div>
            </div>

            <div className="gaq-card" style={{ padding: 22 }}>
              <div className="gaq-section-h">
                <div>
                  <div className="gaq-h3">Top 5 processos prioritários</div>
                  <div className="sub">Ordenados por criticidade, atenção e pressão de entrega</div>
                </div>
                <button className="gaq-btn is-ghost" style={{ fontSize: 12 }} onClick={() => setDrillDown({ title: "Executivo — Processos prioritários", data: executiveStrategic.criticos.length ? executiveStrategic.criticos : executiveStrategic.problematicosRows, color: "#d35400", sourceAba: "executivo" })}>Ver todos <Icon name="chevR" size={12}/></button>
              </div>
              {executiveStrategic.topCriticos.length === 0
                ? <div style={{ color: "var(--gaq-text-3)", fontSize: 13, paddingTop: 8 }}>Sem casos prioritários no recorte actual.</div>
                : executiveStrategic.topCriticos.map((r, idx) => {
                    const sla = r._sla || {};
                    const bk = SLA_BUCKETS[sla.bucket] || SLA_BUCKETS.no_prazo;
                    const reason = { label: bk.label, color: bk.color, bg: bk.bg };
                    return (
                      <div className="gaq-responsive-grid" key={r.ProcessKey || `${r.NumRC || ""}-${idx}`} onClick={() => setSelProc(r)}
                        style={{ display: "grid", gridTemplateColumns: "26px 120px 1fr 128px 72px 26px", gap: 12, alignItems: "center", padding: "12px 0", borderTop: idx ? "1px solid var(--gaq-line)" : "none", cursor: "pointer" }}>
                        <span style={{ fontSize: 12, color: "var(--gaq-text-3)", fontWeight: 700 }}>#{idx + 1}</span>
                        <span style={{ fontFamily: "var(--gaq-mono)", fontSize: 12, fontWeight: 700, color: "var(--gaq-text)" }}>{r.NumRC || r.TicketSD || r.ProcessKey || "—"}</span>
                        <div style={{ minWidth: 0 }}>
                          <div style={{ fontSize: 13, fontWeight: 600, color: "var(--gaq-text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.Objeto || "Sem objecto"}</div>
                          <div className="gaq-meta" style={{ marginTop: 3, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r["Área Requisitante"] || "Área não informada"} · {r.Modalidade || "Modalidade não informada"}</div>
                        </div>
                        <span style={{ justifySelf: "start", fontSize: 11, fontWeight: 700, color: reason.color, background: reason.bg, borderRadius: 999, padding: "6px 10px", whiteSpace: "nowrap" }}>{reason.label}</span>
                        <span className="gaq-num" style={{ textAlign: "right", fontWeight: 700, color: reason.color }}>{(r.diasTotais || 0).toLocaleString("pt-BR")} d.u.</span>
                        <button className="gaq-icon-btn" onClick={(e) => { e.stopPropagation(); setSelProc(r); }}><Icon name="chevR" size={14}/></button>
                      </div>
                    );
                  })}
            </div>
          </div>

          {!loginArea && <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1fr)", gap: 18, marginBottom: 16 }}>
            <div className="gaq-card" style={{ padding: 22 }}>
              <div className="gaq-section-h">
                <div>
                  <div className="gaq-h3">Ranking de compradores</div>
                  <div className="sub">Volume de demandas em andamento e média de d.u. NCL</div>
                </div>
                <span className="gaq-meta">{(executiveStrategic.buyerRanking || []).length} comprador(es)</span>
              </div>
              {(!executiveStrategic.buyerRanking || executiveStrategic.buyerRanking.length === 0)
                ? <div style={{ color: "var(--gaq-text-3)", fontSize: 13, paddingTop: 8 }}>Sem compradores no recorte actual.</div>
                : <div style={{ marginTop: 4 }}>
                    <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "28px 1fr 70px 70px", gap: 8, padding: "6px 0", borderBottom: "1px solid var(--gaq-line)" }}>
                      <span style={{ fontSize: 10, fontWeight: 600, color: "var(--gaq-text-3)", textTransform: "uppercase" }}>#</span>
                      <span style={{ fontSize: 10, fontWeight: 600, color: "var(--gaq-text-3)", textTransform: "uppercase" }}>Comprador</span>
                      <span style={{ fontSize: 10, fontWeight: 600, color: "var(--gaq-text-3)", textTransform: "uppercase", textAlign: "right" }}>Demandas</span>
                      <span style={{ fontSize: 10, fontWeight: 600, color: "var(--gaq-text-3)", textTransform: "uppercase", textAlign: "right" }}>Média d.u.</span>
                    </div>
                    {executiveStrategic.buyerRanking.map((b, idx) => (
                      <div className="gaq-responsive-grid" key={b.nome}
                        onClick={() => setDrillDown({ title: `Executivo — Demandas NCL em andamento de ${b.nome}`, data: b.rows || [], color: "#1f6feb", sourceAba: "executivo" })}
                        style={{ display: "grid", gridTemplateColumns: "28px 1fr 70px 70px", gap: 8, padding: "8px 0", borderBottom: "1px solid var(--gaq-line)", alignItems: "center", cursor: "pointer" }}>
                        <span style={{ fontSize: 11, fontWeight: 700, color: "var(--gaq-text-3)" }}>{idx + 1}</span>
                        <span style={{ fontSize: 12.5, fontWeight: 500, color: "var(--gaq-text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{b.nome}</span>
                        <span className="gaq-num" style={{ fontSize: 13, fontWeight: 700, textAlign: "right", color: "var(--gaq-blue)" }}>{b.demandas}</span>
                        <span className="gaq-num" style={{ fontSize: 12, fontWeight: 500, textAlign: "right", color: b.mediaDU > 50 ? "var(--gaq-red)" : b.mediaDU > 30 ? "var(--gaq-orange)" : "var(--gaq-green)" }}>{b.mediaDU} d.u.</span>
                      </div>
                    ))}
                  </div>}
            </div>

            <div className="gaq-card" style={{ padding: 22 }}>
              <div className="gaq-section-h">
                <div>
                  <div className="gaq-h3">Ranking histórico concluído</div>
                  <div className="sub">Quem mais concluiu compras e com menor tempo médio</div>
                </div>
                <span className="gaq-meta">{(executiveStrategic.buyerRankingConcluded || []).length} comprador(es)</span>
              </div>
              {(!executiveStrategic.buyerRankingConcluded || executiveStrategic.buyerRankingConcluded.length === 0)
                ? <div style={{ color: "var(--gaq-text-3)", fontSize: 13, paddingTop: 8 }}>Sem histórico concluído suficiente no recorte actual.</div>
                : <div style={{ marginTop: 4 }}>
                    <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "28px minmax(0,1fr) minmax(160px, 220px) 78px 78px", gap: 8, padding: "6px 0", borderBottom: "1px solid var(--gaq-line)" }}>
                      <span style={{ fontSize: 10, fontWeight: 600, color: "var(--gaq-text-3)", textTransform: "uppercase" }}>#</span>
                      <span style={{ fontSize: 10, fontWeight: 600, color: "var(--gaq-text-3)", textTransform: "uppercase" }}>Comprador</span>
                      <span style={{ fontSize: 10, fontWeight: 600, color: "var(--gaq-text-3)", textTransform: "uppercase" }}>Modalidades</span>
                      <span style={{ fontSize: 10, fontWeight: 600, color: "var(--gaq-text-3)", textTransform: "uppercase", textAlign: "right" }}>Concluídos</span>
                      <span style={{ fontSize: 10, fontWeight: 600, color: "var(--gaq-text-3)", textTransform: "uppercase", textAlign: "right" }}>Média d.u.</span>
                    </div>
                    {executiveStrategic.buyerRankingConcluded.map((b, idx) => (
                      <div className="gaq-responsive-grid" key={b.nome}
                        onClick={() => setDrillDown({ title: `Executivo — Concluídos por ${b.nome}`, data: b.rows || [], color: "#1f6feb", sourceAba: "executivo" })}
                        style={{ display: "grid", gridTemplateColumns: "28px minmax(0,1fr) minmax(160px, 220px) 78px 78px", gap: 8, padding: "10px 0", borderBottom: "1px solid var(--gaq-line)", alignItems: "center", cursor: "pointer" }}>
                        <span style={{ fontSize: 11, fontWeight: 700, color: "var(--gaq-text-3)" }}>{idx + 1}</span>
                        <span style={{ fontSize: 12.5, fontWeight: 500, color: "var(--gaq-text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{b.nome}</span>
                        <div style={{ minWidth: 0 }}>
                          <div style={{ height: 10, background: "rgba(15,23,42,0.06)", borderRadius: 999, overflow: "hidden", display: "flex" }}>
                            {(b.modParts || []).map(part => (
                              <div
                                key={`${b.nome}-${part.label}`}
                                title={`${part.label}: ${part.count} concluído(s) · ${part.pctBuyer}% da carteira de ${b.nome}`}
                                style={{ width: `${Math.max(part.pctBuyer, part.count > 0 ? 6 : 0)}%`, background: part.color, minWidth: part.count > 0 ? 8 : 0 }}
                              />
                            ))}
                          </div>
                          <div style={{ marginTop: 6, display: "flex", gap: 6, flexWrap: "wrap" }}>
                            {(b.modParts || []).slice(0, 3).map(part => (
                              <span key={`${b.nome}-${part.label}-chip`} style={{ fontSize: 10, color: "var(--gaq-text-3)", background: "var(--gaq-bg-2)", borderRadius: 999, padding: "3px 7px" }}>
                                <span style={{ color: part.color, fontWeight: 700 }}>{part.count}</span> {part.label}
                              </span>
                            ))}
                          </div>
                        </div>
                        <span className="gaq-num" style={{ fontSize: 13, fontWeight: 700, textAlign: "right", color: "var(--gaq-blue)" }}>{b.concluidos}</span>
                        <span className="gaq-num" style={{ fontSize: 12, fontWeight: 500, textAlign: "right", color: b.mediaDU > 90 ? "var(--gaq-red)" : b.mediaDU > 45 ? "var(--gaq-orange)" : "var(--gaq-green)" }}>{b.mediaDU} d.u.</span>
                      </div>
                    ))}
                  </div>}
            </div>
          </div>}

          {/* Pré-compra (SD sem RC) por Responsável NCL — gráfico à parte */}
          {!loginArea && (() => {
            const byResp = executiveStrategic.preCompraByResp || [];
            const total = executiveStrategic.preCompraTotal || 0;
            const maxT = byResp.reduce((m, p) => Math.max(m, p.total), 0) || 1;
            const COR_AN = "#e67e22", COR_OK = "#27ae60", COR_OU = "#c0392b";
            const temOutros = byResp.some(p => p.outros > 0);
            const drill = (title, rows, color) => { if (rows && rows.length) setDrillDown({ title, data: rows, color, sourceAba: "executivo" }); };
            const dot = (cor) => ({ width: 9, height: 9, borderRadius: 2, background: cor, display: "inline-block", marginRight: 5 });
            return (
              <div className="gaq-card" style={{ padding: 22, marginBottom: 16 }}>
                <div className="gaq-section-h">
                  <div>
                    <div className="gaq-h3" style={{ display: "flex", alignItems: "center", gap: 7 }}>
                      Pré-compra · Service Desk sem RC — por Responsável (NCL)
                      <span className="kpi-tooltip" style={{ display: "inline-flex" }}>
                        <span style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", width: 16, height: 16, borderRadius: "50%", background: "var(--gaq-blue)", color: "#fff", fontSize: 11, fontWeight: 700, fontStyle: "italic", lineHeight: 1, cursor: "help" }}>i</span>
                        <span className="kpi-tip" style={{ width: 320, textAlign: "left" }}>
                          <b>Regras deste gráfico:</b><br/>
                          • <b>O que entra:</b> processos com nº de SD e ainda SEM RC vinculada (pré-compra na mão do comprador), agrupados pelo responsável NCL (Comprador, ou Avaliador quando não há Comprador).<br/>
                          • <b>Em análise:</b> SD ainda em aberto. <b>Concluído:</b> SD encerrado mas sem RC aberta (pelo status do JSON ou pela data de encerramento no log).<br/>
                          • A soma de todos os responsáveis = total da carteira sem RC. Clique no nome/barra ou em cada número para abrir a lista.
                        </span>
                      </span>
                    </div>
                    <div className="sub">Carteira de pré-compra (SD sem RC) distribuída por responsável. Clique para detalhar.</div>
                  </div>
                  <span className="gaq-meta">{total} processo(s)</span>
                </div>
                <div style={{ display: "flex", gap: 16, margin: "6px 0 12px", fontSize: 11, color: "var(--gaq-text-3)", flexWrap: "wrap" }}>
                  <span style={{ display: "inline-flex", alignItems: "center" }}><span style={dot(COR_AN)} />Em análise</span>
                  <span style={{ display: "inline-flex", alignItems: "center" }}><span style={dot(COR_OK)} />Concluído (sem RC)</span>
                  {temOutros && <span style={{ display: "inline-flex", alignItems: "center" }}><span style={dot(COR_OU)} />Outros status</span>}
                </div>
                {byResp.length === 0
                  ? <div style={{ color: "var(--gaq-text-3)", fontSize: 13, paddingTop: 8 }}>Nenhum processo em pré-compra (SD sem RC) no recorte actual.</div>
                  : <div style={{ display: "flex", flexDirection: "column", gap: 11 }}>
                      {byResp.map(p => {
                        const wTotal = Math.max((p.total / maxT) * 100, 6);
                        return (
                          <div key={p.nome}>
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 8, marginBottom: 4 }}>
                              <span onClick={() => drill(`Executivo — ${p.nome}: Pré-compra (SD sem RC)`, p.rows, COR_AN)}
                                style={{ fontSize: 12.5, fontWeight: 600, color: "var(--gaq-text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", cursor: "pointer" }}>{p.nome}</span>
                              <span className="gaq-num" style={{ fontSize: 11.5, whiteSpace: "nowrap" }}>
                                <span onClick={() => drill(`Executivo — ${p.nome}: SD em análise`, p.emAnaliseRows, COR_AN)}
                                  style={{ color: COR_AN, fontWeight: 700, cursor: p.emAnalise ? "pointer" : "default" }}>{p.emAnalise} análise</span>
                                <span style={{ color: "var(--gaq-text-3)" }}> · </span>
                                <span onClick={() => drill(`Executivo — ${p.nome}: SD concluído (sem RC)`, p.concluidoRows, COR_OK)}
                                  style={{ color: COR_OK, fontWeight: 700, cursor: p.concluido ? "pointer" : "default" }}>{p.concluido} concl.</span>
                                {p.outros > 0 && <>
                                  <span style={{ color: "var(--gaq-text-3)" }}> · </span>
                                  <span onClick={() => drill(`Executivo — ${p.nome}: SD outros status`, p.outrosRows, COR_OU)}
                                    style={{ color: COR_OU, fontWeight: 700, cursor: "pointer" }}>{p.outros} outros</span>
                                </>}
                                <span style={{ color: "var(--gaq-text-3)", fontWeight: 500 }}>&nbsp;&nbsp;({p.total})</span>
                              </span>
                            </div>
                            <div onClick={() => drill(`Executivo — ${p.nome}: Pré-compra (SD sem RC)`, p.rows, COR_AN)}
                              title={`${p.nome}: ${p.total} processo(s)`}
                              style={{ display: "flex", width: `${wTotal}%`, minWidth: 60, height: 12, background: "rgba(15,23,42,0.06)", borderRadius: 999, overflow: "hidden", cursor: "pointer" }}>
                              {p.emAnalise > 0 && <div style={{ width: `${(p.emAnalise / p.total) * 100}%`, height: "100%", background: COR_AN, flexShrink: 0 }} />}
                              {p.concluido > 0 && <div style={{ width: `${(p.concluido / p.total) * 100}%`, height: "100%", background: COR_OK, flexShrink: 0 }} />}
                              {p.outros > 0 && <div style={{ width: `${(p.outros / p.total) * 100}%`, height: "100%", background: COR_OU, flexShrink: 0 }} />}
                            </div>
                          </div>
                        );
                      })}
                    </div>}
              </div>
            );
          })()}

          {false && <>

          {/* KPIs com tooltips e trends */}
          <div className="grid-5 gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 12, marginBottom: 18 }}>
            {[
              { t: "Total", v: execData.tot, c: "var(--gaq-blue)", tip: "Quantidade total de processos cadastrados na base de dados." },
              { t: "Em Andamento", v: execData.emA, c: "var(--gaq-ncl)", tip: "Processos atualmente em andamento em todas as areas (NCL + CPL)." },
              { t: "Taxa Conclusao", v: execData.taxaConc + "%", c: "var(--gaq-green)", tip: "Percentual de processos concluidos sobre o total.", trend: execData.trendConc, trendV: `${execData.curConc} concluidos este mes vs ${execData.prevConc} mes ant.` },
              { t: "Taxa Problematicos", v: execData.taxaProb + "%", c: execData.taxaProb > 20 ? "var(--gaq-red)" : "var(--gaq-orange)", tip: "Percentual de cancelados + fracassados sobre o total.", trend: execData.trendProb, trendV: `${execData.curProb} problematicos este mes vs ${execData.prevProb} mes ant.` },
              { t: "Criticos (>50 d.u.)", v: execData.criticalTotal, c: execData.criticalTotal > 0 ? "var(--gaq-red)" : "var(--gaq-green)", tip: "Processos em andamento ha mais de 50 dias uteis. NCL: " + execData.alertNCLCount + " · CPL: " + execData.alertCPLCount },
            ].map((k, idx) => (
              <div key={k.t} className="kpi-tooltip" style={{ background: "var(--gaq-surface)", borderRadius: "var(--gaq-r-lg)", padding: "18px 16px", boxShadow: "var(--gaq-shadow-1)", border: "1px solid var(--gaq-line)", borderTop: "3px solid " + k.c, textAlign: "center", position: "relative", transition: "transform .15s, box-shadow .15s" }}
                onMouseEnter={e => { e.currentTarget.style.transform = "translateY(-2px)"; e.currentTarget.style.boxShadow = "var(--gaq-shadow-2)"; }}
                onMouseLeave={e => { e.currentTarget.style.transform = "translateY(0)"; e.currentTarget.style.boxShadow = "var(--gaq-shadow-1)"; }}>
                <div className="kpi-tip">{k.tip}</div>
                <div style={{ fontSize: 28, fontWeight: 700, color: k.c, fontVariantNumeric: "tabular-nums", margin: "6px 0" }}>{k.v}</div>
                <div style={{ fontSize: 10.5, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.04em" }}>{k.t}</div>
                {k.trend && <div className={"trend-" + k.trend} style={{ marginTop: 6, fontSize: 10 }}>
                  {k.trend === "up" ? "+" : k.trend === "down" ? "-" : "="} {k.trendV}
                </div>}
              </div>
            ))}
          </div>

          {/* Tempo medio por Modalidade (comparativo anual) */}
          {execData.tempoPorModalidadeAno && execData.tempoPorModalidadeAno.data && execData.tempoPorModalidadeAno.data.length > 0 && (
            <div className="gaq-card" style={{ marginBottom: 16 }}>
              <div style={{ fontWeight: 600, fontSize: 14, color: "var(--gaq-text)", marginBottom: 14, lineHeight: 1.4 }}>
                Tempo medio por Modalidade (d.u.)
                <span style={{ display: "block", fontSize: 12, fontWeight: 400, color: "var(--gaq-text-3)", marginTop: 2 }}>comparativo {execData.tempoPorModalidadeAno.years.join(" vs ")}</span>
              </div>
              <Recharts.ResponsiveContainer width="100%" height={Math.max(260, execData.tempoPorModalidadeAno.data.length * 48)}>
                <Recharts.BarChart data={execData.tempoPorModalidadeAno.data} layout="vertical" margin={{ top: 5, right: 30, left: 10, bottom: 5 }}>
                  <Recharts.YAxis type="category" dataKey="name" fontSize={10} width={150} tick={{ fill: "var(--text3)" }} tickFormatter={v => v.length > 22 ? v.slice(0,20) + "…" : v} />
                  <Recharts.XAxis type="number" fontSize={10} tick={{ fill: "var(--text3)" }} />
                  <Recharts.Tooltip formatter={(v, n, p) => (v==null ? ["—", n] : [v + " d.u.", n])} />
                  <Recharts.Legend verticalAlign="top" height={28} />
                  {execData.tempoPorModalidadeAno.years.length === 1 ? (
                    <Recharts.Bar dataKey={String(execData.tempoPorModalidadeAno.years[0])} fill="#2e86c1" radius={[0,3,3,0]} barSize={16} />
                  ) : (
                    <>
                      <Recharts.Bar dataKey={String(execData.tempoPorModalidadeAno.years[0])} fill="#85c1e9" radius={[0,3,3,0]} barSize={14} />
                      <Recharts.Bar dataKey={String(execData.tempoPorModalidadeAno.years[1])} fill="#2e86c1" radius={[0,3,3,0]} barSize={14} />
                    </>
                  )}
                </Recharts.BarChart>
              </Recharts.ResponsiveContainer>
              <div style={{ fontSize: 10, color: "var(--text3)", marginTop: 8 }}>
                Regra: início = data de abertura da pré-compra (fallback: DATA DO RECEBIMENTO DA RC). Fim = CPL_DATA_HOMOLOGACAO_FINAL (fallback: Data de envio do pedido/Suite; se vazio, usa hoje).
              </div>
            </div>
          )}

          {/* Semaforo de Aging */}
          <div style={{ marginBottom: 18 }}>
            <div style={{ fontWeight: 600, fontSize: 14, color: "var(--gaq-text)", marginBottom: 10 }}>Semaforo de Aging — Processos em Andamento</div>
            <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
              {execData.aging.map(a => (
                <div key={a.faixa} onClick={() => {
                  const ranges = {"0-30 d.u.":[0,30],"31-50 d.u.":[31,50],"51-100 d.u.":[51,100],">100 d.u.":[101,99999]};
                  const [lo,hi] = ranges[a.faixa] || [0,99999];
                  const dd = baseVis.filter(r => r.emA && r.diasTotais >= lo && r.diasTotais <= hi);
                  setDrillDown({ title: "Aging: " + a.faixa, data: dd, color: a.color }); setAba("overview");
                }}
                  style={{ background: "var(--gaq-surface)", border: "1px solid var(--gaq-line)", borderRadius: "var(--gaq-r-lg)", padding: "18px 16px", textAlign: "center", cursor: "pointer", transition: "transform .15s, box-shadow .15s", borderTop: "3px solid " + a.color }}
                  onMouseEnter={e => { e.currentTarget.style.transform = "translateY(-2px)"; e.currentTarget.style.boxShadow = "var(--gaq-shadow-2)"; }}
                  onMouseLeave={e => { e.currentTarget.style.transform = "translateY(0)"; e.currentTarget.style.boxShadow = "none"; }}>
                  <div style={{ fontSize: 28, fontWeight: 700, color: a.color, fontVariantNumeric: "tabular-nums" }}>{a.count}<span style={{ fontSize: 13, fontWeight: 500, color: "var(--gaq-text-3)" }}>/{execData.emA > 0 ? Math.round((a.count/execData.emA)*100) : 0}%</span></div>
                  <div style={{ fontSize: 11, fontWeight: 600, color: a.color, marginTop: 4 }}>{a.faixa}</div>
                  <div style={{ fontSize: 9, color: "var(--gaq-text-3)", marginTop: 2 }}>ver detalhes</div>
                </div>
              ))}
            </div>
          </div>

          {/* Alertas Rápidos */}
          <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 16 }}>
            <div style={{ background: execData.risco > 0 ? "#c0392b18" : "#27ae6018", border: "1px solid " + (execData.risco > 0 ? "#c0392b" : "#27ae60"), borderRadius: 10, padding: 14, textAlign: "center" }}>
              <div style={{ fontSize: 20, fontWeight: 800, color: execData.risco > 0 ? "#c0392b" : "#27ae60" }}>{execData.risco}</div>
              <div style={{ fontSize: 11, fontWeight: 600 }}>🔴 Risco Alto (&gt;100 d.u.)</div>
            </div>
            <div style={{ background: "#1a527618", border: "1px solid #1a5276", borderRadius: 10, padding: 14, textAlign: "center" }}>
              <div style={{ fontSize: 20, fontWeight: 800, color: "#1a5276" }}>{execData.alertNCLCount}</div>
              <div style={{ fontSize: 11, fontWeight: 600 }}>🔵 NCL Críticos</div>
            </div>
            <div style={{ background: "#6c348318", border: "1px solid #6c3483", borderRadius: 10, padding: 14, textAlign: "center" }}>
              <div style={{ fontSize: 20, fontWeight: 800, color: "#6c3483" }}>{execData.alertCPLCount}</div>
              <div style={{ fontSize: 11, fontWeight: 600 }}>🟣 CPL Críticos</div>
            </div>
          </div>

          {/* Gráficos */}
          <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 16 }}>
            <div style={{ background: "var(--card)", borderRadius: 12, padding: 16, boxShadow: "var(--shadow)" }}>
              <div style={{ fontWeight: 700, fontSize: 13, color: "var(--text)", marginBottom: 10 }}>📈 Evolução Mensal (últimos 12 meses)</div>
              <Recharts.ResponsiveContainer width="100%" height={220}>
                <Recharts.BarChart data={execData.meses} margin={{ top: 5, right: 10, left: 0, bottom: 5 }}>
                  <Recharts.XAxis dataKey="label" fontSize={9} tick={{ fill: "var(--text3)" }} />
                  <Recharts.YAxis fontSize={9} tick={{ fill: "var(--text3)" }} />
                  <Recharts.Tooltip />
                  <Recharts.Bar dataKey="abertos" fill="#3498db" name="Abertos" radius={[3,3,0,0]} />
                  <Recharts.Bar dataKey="concluidos" fill="#27ae60" name="Concluídos" radius={[3,3,0,0]} />
                </Recharts.BarChart>
              </Recharts.ResponsiveContainer>
            </div>
            <div style={{ background: "var(--card)", borderRadius: 12, padding: 16, boxShadow: "var(--shadow)" }}>
              <div style={{ fontWeight: 700, fontSize: 13, color: "var(--text)", marginBottom: 10 }}>📊 Distribuição de Status</div>
              <Recharts.ResponsiveContainer width="100%" height={220}>
                <Recharts.BarChart data={[
                  { name: "Em And.", value: execData.emA, fill: "#3498db" },
                  { name: "Concluído", value: execData.conc, fill: "#27ae60" },
                  { name: "Cancelado", value: execData.canc, fill: "#e67e22" },
                  { name: "Fracassado", value: execData.frac, fill: "#c0392b" },
                ]} margin={{ top: 5, right: 10, left: 0, bottom: 5 }}>
                  <Recharts.XAxis dataKey="name" fontSize={10} tick={{ fill: "var(--text3)" }} />
                  <Recharts.YAxis fontSize={9} tick={{ fill: "var(--text3)" }} />
                  <Recharts.Tooltip />
                  <Recharts.Bar dataKey="value" name="Qtd" radius={[4,4,0,0]}>
                    {[
                      { name: "Em And.", value: execData.emA, fill: "#3498db" },
                      { name: "Concluído", value: execData.conc, fill: "#27ae60" },
                      { name: "Cancelado", value: execData.canc, fill: "#e67e22" },
                      { name: "Fracassado", value: execData.frac, fill: "#c0392b" },
                    ].map((entry, idx) => <Recharts.Cell key={idx} fill={entry.fill} />)}
                  </Recharts.Bar>
                </Recharts.BarChart>
              </Recharts.ResponsiveContainer>
            </div>
          </div>

          {/* Top Áreas e Modalidades */}
          <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 16 }}>
            <div style={{ background: "var(--card)", borderRadius: 12, padding: 16, boxShadow: "var(--shadow)" }}>
              <div style={{ fontWeight: 700, fontSize: 13, color: "var(--text)", marginBottom: 10 }}>🏢 Top Áreas Problemáticas (Canc+Frac)</div>
              {execData.topAreasProb.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Nenhuma.</div>
                : execData.topAreasProb.map(([area, cnt], i) => (
                  <div key={area} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", borderBottom: "1px solid var(--border2)" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <div style={{ width: 22, height: 22, borderRadius: "50%", background: "#c0392b", color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: 700 }}>{i+1}</div>
                      <span style={{ fontSize: 12, color: "var(--text)" }}>{area}</span>
                    </div>
                    <span style={{ fontWeight: 700, color: "#c0392b", fontSize: 13 }}>{cnt}</span>
                  </div>
                ))}
            </div>
            <div style={{ background: "var(--card)", borderRadius: 12, padding: 16, boxShadow: "var(--shadow)" }}>
              <div style={{ fontWeight: 700, fontSize: 13, color: "var(--text)", marginBottom: 10 }}>🐢 Modalidades Mais Lentas — Concluídos (Tempo médio d.u.)</div>
              {execData.topModLentas.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Insuficiente (mín. 3 processos concluídos).</div>
                : execData.topModLentas.map((m, i) => (
                  <div key={m.name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", borderBottom: "1px solid var(--border2)" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <div style={{ width: 22, height: 22, borderRadius: "50%", background: "#e67e22", color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: 700 }}>{i+1}</div>
                      <span style={{ fontSize: 12, color: "var(--text)" }}>{m.name}</span>
                    </div>
                    <div style={{ textAlign: "right" }}>
                      <span style={{ fontWeight: 700, color: "#e67e22", fontSize: 13 }}>{m.media} d.u.</span>
                      <span style={{ fontSize: 10, color: "var(--text3)", marginLeft: 6 }}>({m.count} proc.)</span>
                    </div>
                  </div>
                ))}
            </div>
          </div>

          {/* Responsáveis Sobrecarregados — admin only, nunca para acesso por área */}
          {!loginUser && !loginArea && execData.topSobrecarregados.length > 0 && (() => {
            const semaforo = execData.topSobrecarregados.map(c => {
              const procs = baseVis.filter(r => r.emA && ((r.ehCPL ? (r.Pregoeiro || r.cplResp) : r.respNCL) === c.name));
              const agingMedio = procs.length ? Math.round(procs.reduce((s, r) => s + r.diasTotaisGestao, 0) / procs.length) : 0;
              const cor = (c.criticos > 3 || agingMedio > 60) ? "#c0392b" : (c.criticos > 0 || agingMedio > 30) ? "#e67e22" : "#27ae60";
              const label = cor === "#c0392b" ? "Crítico" : cor === "#e67e22" ? "Atenção" : "Saudável";
              const dot = cor === "#c0392b" ? "🔴" : cor === "#e67e22" ? "🟡" : "🟢";
              return { ...c, agingMedio, cor, label, dot, procs };
            });
            return (
              <div style={{ marginBottom: 16 }}>
                <div style={{ fontWeight: 700, fontSize: 14, color: "var(--text)", marginBottom: 8 }}>👥 Semáforo de Saúde — Compradores</div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", gap: 10 }}>
                  {semaforo.map(c => (
                    <div key={c.name} onClick={() => { setDrillDown({ title: c.name + " — Em Andamento", data: c.procs, color: c.cor }); setAba("overview"); }}
                      style={{ background: "var(--card)", borderRadius: 10, padding: 12, boxShadow: "var(--shadow)", cursor: "pointer", borderTop: "4px solid " + c.cor, transition: "transform .15s" }}
                      onMouseEnter={e => e.currentTarget.style.transform = "translateY(-2px)"}
                      onMouseLeave={e => e.currentTarget.style.transform = "none"}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
                        <div style={{ fontSize: 12, fontWeight: 700, color: "var(--text)" }}>{c.name.split(" ")[0]}</div>
                        <span style={{ fontSize: 14 }}>{c.dot}</span>
                      </div>
                      <div style={{ fontSize: 10, fontWeight: 700, color: c.cor, marginBottom: 6, background: c.cor + "18", borderRadius: 4, padding: "2px 6px", display: "inline-block" }}>{c.label}</div>
                      <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 4 }}>
                        <div style={{ background: "var(--card2)", borderRadius: 6, padding: "5px 6px", textAlign: "center" }}>
                          <div style={{ fontSize: 16, fontWeight: 800, color: "#2e86c1" }}>{c.total}</div>
                          <div style={{ fontSize: 8, color: "var(--text3)", textTransform: "uppercase" }}>em and.</div>
                        </div>
                        <div style={{ background: "var(--card2)", borderRadius: 6, padding: "5px 6px", textAlign: "center" }}>
                          <div style={{ fontSize: 16, fontWeight: 800, color: c.agingMedio > 50 ? "#c0392b" : c.agingMedio > 30 ? "#e67e22" : "#27ae60" }}>{c.agingMedio}</div>
                          <div style={{ fontSize: 8, color: "var(--text3)", textTransform: "uppercase" }}>d.u. méd.</div>
                        </div>
                      </div>
                      {c.criticos > 0 && <div style={{ fontSize: 10, fontWeight: 700, color: "#c0392b", marginTop: 6, textAlign: "center" }}>⚠ {c.criticos} crítico(s) &gt;50 d.u.</div>}
                    </div>
                  ))}
                </div>
              </div>
            );
          })()}

          {/* Dashboard Pessoal — comprador */}
          {loginUser && (() => {
            const sorted = [...fBack].sort((a,b) => b.diasTotais - a.diasTotais);
            const criticos = sorted.filter(r => r.diasTotais > 50).length;
            const noPrazo = sorted.filter(r => r.diasTotais <= 30).length;
            const entregas = sorted.filter(r => r.dataEntrega && r.diasParaEntrega !== null && r.diasParaEntrega <= 14 && r.diasParaEntrega >= 0).length;
            const agingMedio = sorted.length ? Math.round(sorted.reduce((s,r) => s + r.diasTotais, 0) / sorted.length) : 0;
            const saudeGeral = criticos === 0 && agingMedio <= 30 ? "#27ae60" : criticos > 3 || agingMedio > 60 ? "#c0392b" : "#e67e22";
            const saudeLabel = saudeGeral === "#27ae60" ? "Saudável" : saudeGeral === "#c0392b" ? "Atenção Crítica" : "Atenção";
            return (
              <div style={{ marginBottom: 16 }}>
                {/* Painel de notificações */}
                {showNotifPanel && notifs.length > 0 && (
                  <div style={{ background: "#fef9e7", border: "1px solid #f39c12", borderRadius: 10, padding: 12, marginBottom: 12 }}>
                    <div style={{ fontWeight: 700, fontSize: 13, color: "#e67e22", marginBottom: 8 }}>🔔 Novidades desde seu último acesso ({notifs.length})</div>
                    {notifs.slice(0, 5).map(r => (
                      <div key={r.ProcessKey} onClick={() => setSelProc(r)} style={{ display: "flex", gap: 8, alignItems: "center", padding: "6px 8px", borderRadius: 6, background: "#fff", marginBottom: 4, cursor: "pointer", border: "1px solid #f39c1233" }}>
                        <span style={{ fontWeight: 700, color: "#2e86c1", fontSize: 11, whiteSpace: "nowrap" }}>{r.NumRC || r.TicketSD || "—"}</span>
                        <span style={{ fontSize: 10, color: "var(--text2)", flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{(r.Objeto || "").slice(0, 60)}</span>
                        <span style={{ fontSize: 10, color: "#e67e22", whiteSpace: "nowrap" }}>{r.faseAtual || "—"}</span>
                      </div>
                    ))}
                    {notifs.length > 5 && <div style={{ fontSize: 10, color: "var(--text3)", textAlign: "right" }}>+ {notifs.length - 5} outros</div>}
                  </div>
                )}
                {/* Boas-vindas + saúde geral */}
                <div style={{ background: "linear-gradient(135deg, #1a5276 0%, #2e86c1 100%)", borderRadius: 12, padding: "14px 18px", marginBottom: 12, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
                  <div>
                    <div style={{ fontSize: 16, fontWeight: 800, color: "#fff" }}>Olá, {loginUser}!</div>
                    <div style={{ fontSize: 11, color: "rgba(255,255,255,0.75)", marginTop: 2 }}>Você tem <b style={{ color: "#fff" }}>{sorted.length}</b> processo(s) em andamento</div>
                  </div>
                  <div style={{ background: saudeGeral + "33", border: "2px solid " + saudeGeral, borderRadius: 10, padding: "8px 16px", textAlign: "center" }}>
                    <div style={{ fontSize: 10, color: "#fff", fontWeight: 700, textTransform: "uppercase", letterSpacing: .5 }}>Saúde Geral</div>
                    <div style={{ fontSize: 18, fontWeight: 800, color: saudeGeral === "#27ae60" ? "#2ecc71" : saudeGeral === "#c0392b" ? "#e74c3c" : "#f39c12" }}>{saudeLabel}</div>
                  </div>
                </div>
                {/* KPIs resumo */}
                <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8, marginBottom: 12 }}>
                  {[
                    { label: "Em Andamento", val: sorted.length, color: "#2e86c1" },
                    { label: "Críticos >50 d.u.", val: criticos, color: criticos > 0 ? "#c0392b" : "#27ae60" },
                    { label: "No Prazo ≤30 d.u.", val: noPrazo, color: "#27ae60" },
                    { label: "Entregas ≤14 dias", val: entregas, color: entregas > 0 ? "#e67e22" : "#2e86c1" },
                  ].map(k => (
                    <div key={k.label} style={{ background: "var(--card)", borderRadius: 10, padding: "10px 12px", boxShadow: "var(--shadow)", textAlign: "center", borderTop: "3px solid " + k.color }}>
                      <div style={{ fontSize: 22, fontWeight: 800, color: k.color }}>{k.val}</div>
                      <div style={{ fontSize: 9, color: "var(--text3)", textTransform: "uppercase", letterSpacing: .4, marginTop: 2 }}>{k.label}</div>
                    </div>
                  ))}
                </div>
                {/* Lista de processos com mini-timeline */}
                {sorted.length > 0 && (
                  <div>
                    <div style={{ fontWeight: 700, fontSize: 13, color: "var(--text)", marginBottom: 8 }}>🎯 Meus Processos — Ordenados por Urgência</div>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(210px, 1fr))", gap: 10 }}>
                      {sorted.map(r => {
                        const cor = r.diasTotais > 100 ? "#c0392b" : r.diasTotais > 50 ? "#e67e22" : r.diasTotais > 30 ? "#f39c12" : "#27ae60";
                        const isNovo = notifs.some(n => n.ProcessKey === r.ProcessKey);
                        return (
                          <div key={r.ProcessKey} onClick={() => setSelProc(r)}
                            style={{ background: "var(--card)", borderRadius: 10, padding: 12, boxShadow: "var(--shadow)", cursor: "pointer", borderTop: "3px solid " + cor, transition: "transform .15s", outline: isNovo ? "2px solid #f39c12" : "none" }}
                            onMouseEnter={e => e.currentTarget.style.transform="translateY(-2px)"}
                            onMouseLeave={e => e.currentTarget.style.transform="none"}>
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 3 }}>
                              <div style={{ fontWeight: 700, color: "#2e86c1", fontSize: 12 }}>{r.NumRC || r.TicketSD || "—"}</div>
                              {isNovo && <span style={{ fontSize: 8, background: "#f39c12", color: "#fff", borderRadius: 3, padding: "1px 5px", fontWeight: 700 }}>NOVO</span>}
                            </div>
                            <div style={{ fontSize: 10, color: "var(--text2)", marginBottom: 6, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{(r.Objeto || "").slice(0, 55) || "—"}</div>
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 4 }}>
                              <span style={{ fontSize: 13, fontWeight: 800, color: cor }}>{r.diasTotais} d.u.</span>
                              <span style={{ fontSize: 9, color: "var(--text3)", background: "var(--card2)", borderRadius: 4, padding: "1px 6px" }}>{r.faseSubarea || "—"}</span>
                            </div>
                            {r.dataEntrega && <div style={{ fontSize: 9, color: "#2e86c1", marginTop: 4 }}>📅 Entrega: {r.dataEntrega.toLocaleDateString("pt-BR")}</div>}
                            <MiniTimeline r={r} />
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
            );
          })()}

          {/* Comparativo Período a Período */}
          {execData.comparativo && <div style={{ background: "var(--card)", borderRadius: 12, padding: 16, boxShadow: "var(--shadow)", marginBottom: 16 }}>
            <div style={{ fontWeight: 700, fontSize: 14, color: "var(--text)", marginBottom: 12 }}>📊 Comparativo — Mês Atual vs. Mês Anterior</div>
            <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 10 }}>
              {[
                { label: "Aberturas", ...execData.comparativo.aberturas, color: "#2e86c1", unit: "" },
                { label: "Conclusões", ...execData.comparativo.conclusoes, color: "#27ae60", unit: "" },
                { label: "Problemáticos", ...execData.comparativo.problematicos, color: "#e67e22", unit: "", invert: true },
                { label: "Aging Médio", ...execData.comparativo.agingMedio, color: "#8e44ad", unit: " d.u.", invert: true },
                { label: "Críticos (>50)", ...execData.comparativo.criticos, color: "#c0392b", unit: "", invert: true },
              ].map(k => {
                const d = Number(k.delta);
                const good = k.invert ? d <= 0 : d >= 0;
                const arrow = d > 0 ? "▲" : d < 0 ? "▼" : "●";
                const arrowColor = d === 0 ? "var(--text3)" : good ? "#27ae60" : "#c0392b";
                return (
                  <div key={k.label} style={{ background: "var(--card2)", borderRadius: 10, padding: "12px 10px", textAlign: "center", borderTop: "3px solid " + k.color }}>
                    <div style={{ fontSize: 10, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase", marginBottom: 4 }}>{k.label}</div>
                    <div style={{ fontSize: 22, fontWeight: 800, color: k.color }}>{k.cur}{k.unit}</div>
                    <div style={{ fontSize: 10, color: "var(--text3)" }}>vs {k.prev}{k.unit} ant.</div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: arrowColor, marginTop: 4 }}>{arrow} {Math.abs(d)}%</div>
                  </div>
                );
              })}
            </div>
          </div>}

          {/* Painel de Ação — apenas para compradores */}
          {loginUser && <div style={{ background: "var(--card)", borderRadius: 12, padding: 16, boxShadow: "var(--shadow)" }}>
            <div style={{ fontWeight: 700, fontSize: 14, color: "var(--text)", marginBottom: 10 }}>📋 Meu Painel de Ação</div>
            {(() => {
              const items = [];
              const criticos    = fBack.filter(r => r.diasTotais > 100);
              const risco       = fBack.filter(r => r.diasTotais > 50 && r.diasTotais <= 100);
              const atencao     = fBack.filter(r => r.diasTotais > 30 && r.diasTotais <= 50);
              const emDia       = fBack.filter(r => r.diasTotais <= 30);
              const semEnvio30  = fBack.filter(r => r.semEnvio && r.diasTotais > 30);
              if (criticos.length > 0)   items.push({ icon: "🔴", text: `${criticos.length} processo(s) com mais de 100 d.u. sem conclusão — ação urgente necessária.`, color: "#c0392b" });
              if (risco.length > 0)      items.push({ icon: "🟠", text: `${risco.length} processo(s) entre 51 e 100 d.u. — acompanhar de perto para não virar crítico.`, color: "#e67e22" });
              if (atencao.length > 0)    items.push({ icon: "🟡", text: `${atencao.length} processo(s) entre 31 e 50 d.u. — manter cadência e não perder prazo.`, color: "#f39c12" });
              if (semEnvio30.length > 0) items.push({ icon: "📦", text: `${semEnvio30.length} processo(s) com mais de 30 d.u. ainda sem envio de pedido/Suite. Verificar possível travamento.`, color: "#c0392b" });
              if (emDia.length > 0)      items.push({ icon: "✅", text: `${emDia.length} processo(s) dentro do prazo (0–30 d.u.). Continue acompanhando.`, color: "#27ae60" });
              if (items.length === 0)    items.push({ icon: "✅", text: "Nenhum alerta no momento. Todos os seus processos estão dentro do prazo.", color: "#27ae60" });
              return items.map((ins, i) => (
                <div key={i} style={{ display: "flex", gap: 10, alignItems: "flex-start", padding: "8px 0", borderBottom: "1px solid var(--border2)" }}>
                  <span style={{ fontSize: 18 }}>{ins.icon}</span>
                  <span style={{ fontSize: 12, color: ins.color, fontWeight: 500 }}>{ins.text}</span>
                </div>
              ));
            })()}
          </div>}

          {/* Insights Automáticos — apenas para administrador */}
          {!loginUser && <div style={{ background: "var(--card)", borderRadius: 12, padding: 16, boxShadow: "var(--shadow)" }}>
            <div style={{ fontWeight: 700, fontSize: 14, color: "var(--text)", marginBottom: 10 }}>💡 Insights Automáticos</div>
            {(() => {
              const insights = [];
              // Insight: Taxa de problemáticos
              if (execData.taxaProb > 20) insights.push({ icon: "🔴", text: `Taxa de processos problemáticos está em ${execData.taxaProb}% — acima do aceitável (20%). Investigar causas de cancelamento e fracasso.`, color: "#c0392b" });
              // Insight: Risco alto
              if (execData.risco > 0) insights.push({ icon: "⚠️", text: `${execData.risco} processo(s) com mais de 100 d.u. sem conclusão. Risco de impacto operacional e perda de prazos.`, color: "#e67e22" });
              // Insight: Áreas problemáticas
              if (execData.topAreasProb.length > 0 && execData.topAreasProb[0][1] > 3) insights.push({ icon: "🏢", text: `Área "${execData.topAreasProb[0][0]}" lidera cancelamentos/fracassos com ${execData.topAreasProb[0][1]} ocorrências. Avaliar qualidade da especificação técnica e planejamento.`, color: "#8e44ad" });
              // Insight: Sobrecarga (apenas admin — nunca para acesso por área)
              if (!loginArea && execData.topSobrecarregados.length > 0 && execData.topSobrecarregados[0].criticos > 3) insights.push({ icon: "👤", text: `${execData.topSobrecarregados[0].name} possui ${execData.topSobrecarregados[0].criticos} processos críticos de ${execData.topSobrecarregados[0].total} em andamento. Avaliar redistribuição urgente.`, color: "#c0392b" });
              if (loginArea && execData.topSobrecarregados.reduce((s,c) => s + c.criticos, 0) > 3) insights.push({ icon: "🔴", text: `${execData.topSobrecarregados.reduce((s,c) => s + c.criticos, 0)} processo(s) da área ${loginAreaDisp} estão em estado crítico (>50 d.u.). Acionar o GAQ para verificar prioridade.`, color: "#c0392b" });
              // Insight: Gargalo de cotação (fase entre propostas e aprovação)
              const procsCotacao = base.filter(r => r.emA && pd(r["Data inicial do envio de propostas"]) && !pd(r["Data final do envio de propostas"]));
              if (procsCotacao.length > 5) insights.push({ icon: "📋", text: `${procsCotacao.length} processos em andamento estão na fase de cotação (aguardando retorno de propostas). Esta etapa representa um gargalo significativo do fluxo — acompanhar prazos de resposta dos fornecedores.`, color: "#d35400" });
              // Insight: Processos sem envio de pedido
              const semPedido = base.filter(r => r.emA && r.semEnvio && r.diasTotais > 30);
              if (semPedido.length > 5) insights.push({ icon: "📦", text: `${semPedido.length} processos em andamento há mais de 30 d.u. ainda sem envio de pedido/Suite. Verificar se há travamento no fluxo de aprovação.`, color: "#c0392b" });
              // Insight: Tempo médio de conclusão
              if (execData.mediaLTConc > 0) {
                const meta = 50;
                if (execData.mediaLTConc > meta) insights.push({ icon: "⏱", text: `Tempo médio de conclusão é ${execData.mediaLTConc} d.u. — acima da meta de ${meta} d.u. Modalidades como ${execData.topModLentas.length > 0 ? execData.topModLentas[0].name + " (" + execData.topModLentas[0].media + " d.u.)" : "—"} puxam a média para cima.`, color: "#e67e22" });
                else insights.push({ icon: "⏱", text: `Tempo médio de conclusão de ${execData.mediaLTConc} d.u. está dentro de parâmetros aceitáveis.`, color: "#27ae60" });
              }
              // Insight: Processos parados em análise de SD
              const emAnaliseSD = base.filter(r => r.emA && nrm(r.statusDet || "").includes("analise service desk"));
              if (emAnaliseSD.length > 10) insights.push({ icon: "🔍", text: `${emAnaliseSD.length} processos parados em "Análise Service Desk". Avaliar se há backlog na pré-compra que impacta o início efetivo dos processos.`, color: "#d35400" });
              // Insight: Processos aguardando aprovação
              const emAprov = base.filter(r => r.emA && (nrm(r.statusDet || "").includes("aprovac") || nrm(r.statusDet || "").includes("alcada")));
              if (emAprov.length > 5) insights.push({ icon: "✍️", text: `${emAprov.length} processos aguardando aprovação de alçada. Gargalo potencial no fluxo de aprovações — agilizar validação para destravar processos.`, color: "#e67e22" });
              // Insight: SLA positivo
              if (execData.taxaSLA > 70) insights.push({ icon: "✅", text: `${execData.taxaSLA}% dos processos concluídos ficaram dentro do SLA (≤50 d.u.). Desempenho saudável.`, color: "#27ae60" });
              // Insight: Taxa de conclusão boa
              if (execData.taxaConc > 70) insights.push({ icon: "✅", text: `Taxa de conclusão de ${execData.taxaConc}% está saudável. Manter boas práticas.`, color: "#27ae60" });
              if (insights.length === 0) insights.push({ icon: "✅", text: "Nenhum alerta crítico no momento. Indicadores dentro dos parâmetros.", color: "#27ae60" });
              return insights.map((ins, i) => (
                <div key={i} style={{ display: "flex", gap: 10, alignItems: "flex-start", padding: "8px 0", borderBottom: "1px solid var(--border2)" }}>
                  <span style={{ fontSize: 18 }}>{ins.icon}</span>
                  <span style={{ fontSize: 12, color: ins.color, fontWeight: 500 }}>{ins.text}</span>
                </div>
              ));
            })()}
          </div>}
          </>}
          </>}
        </div> : <div style={{ textAlign: "center", padding: 60, color: "var(--text3)" }}>
          <div style={{ fontSize: 38 }}>📭</div><div style={{ fontWeight: 600, marginTop: 8 }}>Carregue a base para acessar o painel executivo.</div>
          <button onClick={() => setAba("upload")} style={{ marginTop: 12, background: "#1a5276", color: "#fff", border: "none", borderRadius: 8, padding: "9px 22px", cursor: "pointer", fontSize: 13 }}>Upload</button>
        </div>);
}
