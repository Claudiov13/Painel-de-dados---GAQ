import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function OverviewScreen({ areaAnoData, base, drillDown, fBack, fBase, loginArea, loginAreaDisp, loginUser, ocultarProblematicosOverview, overviewAdmin, overviewInfoCard, overviewInfoDefs, pgDrillDown, pgDrillDownSz, phaseIntervals, renderBuscaPanel, setAba, setDrillDown, setOcultarProblematicosOverview, setOverviewInfoCard, setPgDrillDown, setPgDrillDownSz, setSelProc }) {
  return (base.length === 0
          ? <div style={{ textAlign: "center", padding: 80, color: "var(--gaq-text-3)" }}>
              <Icon name="upload" size={40} color="var(--gaq-text-3)" style={{ marginBottom: 12 }} />
              <div style={{ fontWeight: 600, fontSize: 15, marginTop: 8 }}>Nenhuma base carregada.</div>
              <button onClick={() => setAba("upload")} className="gaq-btn is-primary" style={{ marginTop: 16, padding: "10px 24px", fontSize: 13 }}>Upload</button>
            </div>
          : <>
              {renderBuscaPanel()}
              {drillDown ? <DrillDownPanel
                drillDown={drillDown}
                defaultSourceAba="overview"
                setDrillDown={setDrillDown}
                setAba={setAba}
                setSelProc={setSelProc}
                pgDrillDown={pgDrillDown}
                pgDrillDownSz={pgDrillDownSz}
                setPgDrillDown={setPgDrillDown}
                setPgDrillDownSz={setPgDrillDownSz}
                phaseIntervals={phaseIntervals}
              />
              : <>
              {!loginUser && !loginArea && overviewAdmin && <>
              <div className="anim-card anim-d1" style={{ marginBottom: 18 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 14, flexWrap: "wrap", marginBottom: 18 }}>
                  <div>
                    <div className="gaq-eyebrow" style={{ marginBottom: 6 }}>Resumo</div>
                    <h2 className="gaq-h1" style={{ margin: 0 }}>{overviewAdmin.saudacao}, GAQ.</h2>
                    <p className="gaq-body" style={{ margin: "4px 0 0", color: "var(--gaq-text-3)" }}>
                      <b style={{ color: "var(--gaq-text-2)", fontWeight: 600 }}>{overviewAdmin.emAndamento.toLocaleString("pt-BR")} processos</b> em andamento ·
                      <span style={{ color: "var(--gaq-red)", fontWeight: 600 }}> {overviewAdmin.criticos.length} críticos</span>
                      {overviewAdmin.atencao.length > 0 && <span style={{ color: "var(--gaq-orange)", fontWeight: 600 }}> e {overviewAdmin.atencao.length} em atenção</span>}
                      {" "}requerem atenção hoje
                    </p>
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 4, alignItems: "flex-end" }}>
                    <span className="gaq-meta gaq-num"><b style={{ color: "var(--gaq-text)" }}>{fBase.length.toLocaleString("pt-BR")}</b> registos filtrados</span>
                    {overviewAdmin.updatedAt && <span className="gaq-meta">Atualizado: {overviewAdmin.updatedAt}</span>}
                  </div>
                </div>

                <div className="gaq-kpi-grid" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", marginBottom: 12 }}>
                  {[
                    {
                      label: "Em andamento",
                      value: overviewAdmin.emAndamento,
                      sub: `Com RC: ${overviewAdmin.ativosComRC.length} · Pré-compra: ${overviewAdmin.sdTriagem.length}`,
                      color: "var(--gaq-blue)",
                      rows: fBack,
                      title: "Processos em andamento",
                      spark: overviewAdmin.sparkEmAndamento,
                      infoKey: "em_andamento",
                    },
                    {
                      label: "Crítico",
                      value: overviewAdmin.criticos.length,
                      sub: `SLA + entrega rompidos: ${overviewAdmin.criticos.length} · Em atenção: ${overviewAdmin.atencao.length}`,
                      color: "#d35400",
                      rows: overviewAdmin.criticos,
                      title: "Casos críticos",
                      spark: overviewAdmin.sparkCriticos,
                      infoKey: "criticos",
                    },
                    {
                      label: "SLA Vencido",
                      value: overviewAdmin.slaVencidos.length,
                      sub: `RC acima do SLA: ${overviewAdmin.atrasadosRC.length} · Pré-compra acima do SLA: ${overviewAdmin.sdAtrasado.length}`,
                      color: "var(--gaq-red)",
                      rows: overviewAdmin.slaVencidos,
                      title: "Casos com SLA vencido",
                      spark: overviewAdmin.sparkAtrasados,
                      infoKey: "atrasados_sla",
                    },
                    {
                      label: "SLA Atendido",
                      value: overviewAdmin.slaAtendidos.length,
                      sub: `RC dentro do SLA: ${overviewAdmin.slaAtendidosRC.length}${overviewAdmin.atencaoDentroSlaRC.length ? ` (−${overviewAdmin.atencaoDentroSlaRC.length} em atenção)` : ""} · Pré-compra no prazo: ${overviewAdmin.slaAtendidosSD.length}`,
                      color: "var(--gaq-green)",
                      rows: overviewAdmin.slaAtendidos,
                      title: "Processos com SLA atendido (excl. atenção)",
                      spark: overviewAdmin.sparkSla,
                      infoKey: "sla_cumprido",
                    },
                    {
                      label: "Score da área",
                      value: overviewAdmin.mediaScore,
                      sub: `${overviewAdmin.criticos.length} críticos · ${overviewAdmin.atencao.length} atenção · ${overviewAdmin.slaVencidosScore.length} SLA vencido · ${overviewAdmin.slaAtendidosScore.length} SLA atendido`,
                      color: overviewAdmin.scoreColor,
                      rows: fBack,
                      title: "Score da carteira ativa",
                      spark: overviewAdmin.sparkScore,
                      infoKey: "score_medio",
                    },
                  ].map(card => (
                    <div key={card.label} className="gaq-kpi" onClick={() => setDrillDown({ title: card.title, data: card.rows, color: card.color })}
                      style={{ cursor: "pointer" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, marginBottom: 4 }}>
                        <span className="lbl">{card.label}</span>
                        <button onClick={(e) => { e.stopPropagation(); setOverviewInfoCard(v => v === card.infoKey ? "" : card.infoKey); }}
                          style={{ background: overviewInfoCard === card.infoKey ? card.color : "var(--gaq-line)", border: "none", borderRadius: "50%", width: 20, height: 20, cursor: "pointer", fontSize: 11, fontWeight: 800, color: overviewInfoCard === card.infoKey ? "#fff" : "var(--gaq-text-3)", flexShrink: 0, lineHeight: "20px", textAlign: "center" }}
                          title={`Ver regra de ${card.label}`}>ℹ</button>
                      </div>
                      <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between", gap: 10 }}>
                        <span className="val gaq-num" style={{ color: card.color }}>{card.value}</span>
                        <Spark data={card.spark.length > 1 ? card.spark : [0, card.rows.length || 0]} color={card.color} w={64} h={22} />
                      </div>
                      <span className="sub">{card.sub}</span>
                    </div>
                  ))}
                </div>
                {overviewInfoCard && overviewInfoDefs[overviewInfoCard] && (
                  <div style={{ marginBottom: 18, background: overviewInfoDefs[overviewInfoCard].bg, border: `1px solid ${overviewInfoDefs[overviewInfoCard].color}33`, borderRadius: 12, padding: "14px 16px" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center", marginBottom: 8 }}>
                      <div style={{ fontSize: 12, fontWeight: 800, color: overviewInfoDefs[overviewInfoCard].color, textTransform: "uppercase", letterSpacing: 0.4 }}>
                        Regras — {overviewInfoDefs[overviewInfoCard].title}
                      </div>
                      <button onClick={() => setOverviewInfoCard("")}
                        style={{ background: "transparent", border: "none", color: "var(--gaq-text-3)", cursor: "pointer", fontSize: 16, lineHeight: 1 }}>×</button>
                    </div>
                    <div style={{ display: "grid", gap: 5 }}>
                      {overviewInfoDefs[overviewInfoCard].lines.map((line, idx) => (
                        <div key={idx} style={{ fontSize: 12, color: "var(--gaq-text-2)", lineHeight: 1.5 }}>• {line}</div>
                      ))}
                    </div>
                  </div>
                )}

                <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "minmax(0, 1.6fr) minmax(320px, 1fr)", gap: 18, marginBottom: 18 }}>
                  <div className="gaq-card" style={{ padding: 22 }}>
                    <div className="gaq-section-h">
                      <div>
                        <div className="gaq-h3">Distribuição por subárea</div>
                        <div className="sub">{overviewAdmin.emAndamento.toLocaleString("pt-BR")} processos em andamento</div>
                      </div>
                      <button className="gaq-btn is-ghost" style={{ fontSize: 12 }}
                        onClick={() => setDrillDown({ title: "Processos em andamento por subárea", data: fBack, color: "var(--gaq-blue)" })}>
                        Ver detalhes <Icon name="chevR" size={12}/>
                      </button>
                    </div>
                    {overviewAdmin.subareas.map(item => (
                      <div key={item.key} style={{ marginTop: 16, cursor: "pointer" }}
                        onClick={() => setDrillDown({ title: `${item.label} — Em andamento`, data: item.rows, color: item.color })}>
                        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6, fontSize: 13, gap: 8 }}>
                          <span style={{ display: "flex", gap: 8, alignItems: "center", minWidth: 0 }}>
                            <span className="gaq-dot" style={{ background: item.color }} />
                            <span style={{ fontWeight: 500, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{item.label}</span>
                          </span>
                          <span className="gaq-num" style={{ color: "var(--gaq-text-3)", flexShrink: 0 }}>
                            <b style={{ color: "var(--gaq-text)" }}>{item.rows.length}</b> · {item.pct}%
                          </span>
                        </div>
                        <div style={{ height: 6, background: "rgba(0,0,0,0.05)", borderRadius: 4, overflow: "hidden" }}>
                          <div style={{ width: `${item.pct}%`, height: "100%", background: item.color, borderRadius: 4 }} />
                        </div>
                      </div>
                    ))}

                    <div style={{ marginTop: 22, paddingTop: 16, borderTop: "1px solid var(--gaq-line)" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "flex-start", flexWrap: "wrap", marginBottom: 12 }}>
                        <div>
                          <div className="gaq-eyebrow" style={{ fontSize: 10, marginBottom: 4 }}>Régua SLA e entrega (%)</div>
                          <div className="gaq-meta">Percentual dos processos em andamento por subárea em cada situação da régua global.</div>
                        </div>
                        <button onClick={() => setOverviewInfoCard(v => v === "atencao" ? "" : "atencao")}
                          style={{ background: overviewInfoCard === "atencao" ? "#e67e22" : "var(--gaq-line)", border: "none", borderRadius: 999, padding: "5px 12px", cursor: "pointer", fontSize: 11, fontWeight: 700, color: overviewInfoCard === "atencao" ? "#fff" : "var(--gaq-text-2)", display: "inline-flex", alignItems: "center", gap: 6 }}
                          title="Ver regra de classificação SLA e Atenção">
                          ℹ Regra de classificação
                        </button>
                      </div>
                      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 14, alignItems: "center" }}>
                        {[
                          { label: "Crítico", color: "#d35400", info: "criticos" },
                          { label: "Em atenção", color: "#ff9500", info: "atencao" },
                          { label: "SLA vencido", color: "#ff3b30", info: "atrasados_sla" },
                          { label: "No prazo", color: "#34c759", info: "sla_cumprido" },
                        ].map(legend => (
                          <button key={legend.label}
                            onClick={() => setOverviewInfoCard(v => v === legend.info ? "" : legend.info)}
                            style={{ display: "inline-flex", alignItems: "center", gap: 6, padding: "4px 10px", borderRadius: 999, background: overviewInfoCard === legend.info ? legend.color : `${legend.color}12`, color: overviewInfoCard === legend.info ? "#fff" : legend.color, fontSize: 11, fontWeight: 600, border: `1px solid ${legend.color}33`, cursor: "pointer" }}
                            title={`Ver regra de ${legend.label}`}>
                            <span className="gaq-dot" style={{ background: overviewInfoCard === legend.info ? "#fff" : legend.color, width: 7, height: 7 }} />
                            {legend.label}
                          </button>
                        ))}
                      </div>
                      <div style={{ display: "grid", gap: 14 }}>
                        {overviewAdmin.subareas.map(item => (
                          <div key={`mix-${item.key}`}>
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, marginBottom: 6 }}>
                              <span style={{ fontSize: 12.5, fontWeight: 600, color: "var(--gaq-text)" }}>{item.label}</span>
                              <span className="gaq-meta gaq-num">{item.rows.length.toLocaleString("pt-BR")} proc.</span>
                            </div>
                            <div style={{ height: 12, display: "flex", borderRadius: 999, overflow: "hidden", background: "rgba(0,0,0,0.05)", border: "1px solid var(--gaq-line)" }}>
                              {item.statusSegments.map(seg => (
                                seg.pct > 0
                                  ? <button
                                      key={`${item.key}-${seg.key}`}
                                      onClick={() => setDrillDown({ title: `${item.label} — ${seg.label}`, data: seg.rows, color: seg.color })}
                                      title={`${seg.label}: ${seg.count} processo(s) · ${seg.pct}%`}
                                      style={{ width: `${seg.pct}%`, minWidth: seg.count > 0 ? 10 : 0, border: "none", background: seg.color, padding: 0, cursor: "pointer" }}
                                    />
                                  : <React.Fragment key={`${item.key}-${seg.key}`} />
                              ))}
                            </div>
                            <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(4, minmax(0, 1fr))", gap: 8, marginTop: 8 }}>
                              {item.statusSegments.map(seg => (
                                <button
                                  key={`${item.key}-${seg.key}-meta`}
                                  onClick={() => setDrillDown({ title: `${item.label} — ${seg.label}`, data: seg.rows, color: seg.color })}
                                  disabled={seg.count === 0}
                                  style={{ textAlign: "left", border: `1px solid ${seg.count ? `${seg.color}33` : "var(--gaq-line)"}`, background: seg.count ? `${seg.color}10` : "var(--gaq-bg-2)", color: seg.count ? seg.color : "var(--gaq-text-3)", borderRadius: 12, padding: "8px 10px", cursor: seg.count ? "pointer" : "default", opacity: seg.count ? 1 : 0.72 }}>
                                  <div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: 3 }}>{seg.label}</div>
                                  <div className="gaq-num" style={{ fontSize: 16, fontWeight: 800 }}>{seg.pct}%</div>
                                  <div style={{ fontSize: 10, marginTop: 2 }}>{seg.count} proc.</div>
                                </button>
                              ))}
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>

                    <div style={{ marginTop: 22, paddingTop: 16, borderTop: "1px solid var(--gaq-line)", display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 18 }}>
                      {[
                        { label: "Tempo médio NCL", value: `${overviewAdmin.mediaNCL} d.u.`, color: "var(--gaq-ncl)", rows: overviewAdmin.subareas.find(x => x.key === "NCL")?.rows || [] },
                        { label: "Tempo médio CPL", value: `${overviewAdmin.mediaCPL} d.u.`, color: "var(--gaq-cpl)", rows: overviewAdmin.subareas.find(x => x.key === "CPL")?.rows || [] },
                        { label: "Acima de 60 d.u.", value: overviewAdmin.above60.length, color: "var(--gaq-red)", rows: overviewAdmin.above60 },
                      ].map(stat => (
                        <div key={stat.label} onClick={() => setDrillDown({ title: stat.label, data: stat.rows, color: stat.color })} style={{ cursor: "pointer" }}>
                          <div className="gaq-eyebrow" style={{ fontSize: 10, marginBottom: 3 }}>{stat.label}</div>
                          <div className="gaq-h2 gaq-num" style={{ color: stat.color }}>{stat.value}</div>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div style={{ display: "grid", gap: 18 }}>
                    <div className="gaq-card" style={{ padding: 22 }}>
                      <div className="gaq-section-h">
                        <div>
                          <div className="gaq-h3">Atividade recente</div>
                          <div className="sub">Últimas movimentações da base atual</div>
                        </div>
                        <span className="gaq-meta">{overviewAdmin.recentActivity.length} itens</span>
                      </div>
                      {overviewAdmin.recentActivity.length === 0
                        ? <div style={{ color: "var(--gaq-text-3)", fontSize: 13, paddingTop: 8 }}>Sem movimentações recentes.</div>
                        : overviewAdmin.recentActivity.map((item, idx) => (
                          <div key={`${item.label}-${idx}`} onClick={() => setSelProc(item.proc)}
                            style={{ display: "flex", gap: 12, padding: "10px 0", borderTop: idx ? "1px solid var(--gaq-line)" : "none", cursor: "pointer" }}>
                            <div style={{ width: 28, height: 28, borderRadius: "50%", background: `${item.color}22`, color: item.color, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 11, flex: "none" }}>
                              {item.who.split(" ").map(s => s[0]).join("").slice(0, 2).toUpperCase()}
                            </div>
                            <div style={{ flex: 1, minWidth: 0, fontSize: 13, lineHeight: 1.4 }}>
                              <span style={{ fontWeight: 600 }}>{item.who}</span>
                              <span style={{ color: "var(--gaq-text-3)" }}> {item.verb} </span>
                              <span style={{ fontFamily: "var(--gaq-mono)", fontSize: 12, color: "var(--gaq-blue)" }}>{item.label}</span>
                              <div className="gaq-meta" style={{ marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{item.detail}</div>
                            </div>
                            <span className="gaq-meta gaq-num" style={{ flex: "none" }}>{item.stamp}</span>
                          </div>
                        ))}
                    </div>

                    <div className="gaq-card" style={{ padding: 22 }}>
                      <div className="gaq-section-h">
                        <div>
                          <div className="gaq-h3">Mudanças na atualização</div>
                          <div className="sub">
                            {overviewAdmin.baseDiffInfo.hasPrevious
                              ? `Comparado com ${overviewAdmin.baseDiffInfo.previousAt || "a base anterior"}`
                              : "Sem base anterior para comparar"}
                          </div>
                        </div>
                        {overviewAdmin.baseDiffInfo.currentAt && <span className="gaq-meta">{overviewAdmin.baseDiffInfo.currentAt}</span>}
                      </div>
                      {!overviewAdmin.baseDiffInfo.hasPrevious
                        ? <div style={{ color: "var(--gaq-text-3)", fontSize: 13, paddingTop: 8 }}>A lista de mudanças passa a ser preenchida a partir da próxima atualização da base.</div>
                        : overviewAdmin.baseDiffFeed.length === 0
                          ? <div style={{ color: "var(--gaq-text-3)", fontSize: 13, paddingTop: 8 }}>Nenhuma mudança relevante foi identificada face à base anterior.</div>
                          : overviewAdmin.baseDiffFeed.map((change, idx) => (
                            <div key={change.key} onClick={() => change.proc && setSelProc(change.proc)}
                              style={{ display: "flex", gap: 12, padding: "10px 0", borderTop: idx ? "1px solid var(--gaq-line)" : "none", cursor: change.proc ? "pointer" : "default" }}>
                              <div style={{ width: 28, height: 28, borderRadius: "50%", background: `${change.color}22`, color: change.color, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 10, flex: "none" }}>
                                {change.icon}
                              </div>
                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ fontSize: 13, fontWeight: 600, color: "var(--gaq-text)" }}>{change.title}</div>
                                <div className="gaq-meta" style={{ marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{change.detail}</div>
                              </div>
                              <span className="gaq-meta gaq-num" style={{ flex: "none" }}>{change.stamp}</span>
                            </div>
                          ))}
                    </div>
                  </div>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 18, marginBottom: 12 }}>
                  {[
                    { title: "Modalidade por volume", border: "var(--gaq-blue)", color: "var(--gaq-blue)", rows: overviewAdmin.modalidades, titlePrefix: "Modalidade" },
                    { title: "Top áreas requisitantes", border: "var(--gaq-green)", color: "var(--gaq-green)", rows: overviewAdmin.areasAtivas, titlePrefix: "Área requisitante" },
                    { title: "Top 10 status atuais", border: "var(--gaq-orange)", color: "var(--gaq-orange)", rows: overviewAdmin.statusAtivos, titlePrefix: "Status atual" },
                  ].map(card => {
                    const max = card.rows.length ? card.rows[0].count : 1;
                    return (
                      <WCard key={card.title} title={card.title} border={card.border}>
                        {card.rows.length === 0
                          ? <div style={{ color: "var(--gaq-text-3)", fontSize: 12 }}>Sem dados no recorte atual.</div>
                          : card.rows.map(row => (
                            <div key={row.label} onClick={() => setDrillDown({ title: `${card.titlePrefix} — ${row.label}`, data: row.rows, color: card.color })}
                              style={{ padding: "7px 0", borderBottom: "1px solid var(--gaq-line)", cursor: "pointer" }}>
                              <div style={{ display: "flex", justifyContent: "space-between", gap: 8, fontSize: 12.5, marginBottom: 6 }}>
                                <span style={{ color: "var(--gaq-text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={row.label}>{row.label}</span>
                                <b style={{ color: card.color, fontVariantNumeric: "tabular-nums", flexShrink: 0 }}>{row.count}</b>
                              </div>
                              <div style={{ height: 6, background: "rgba(0,0,0,0.05)", borderRadius: 999, overflow: "hidden" }}>
                                <div style={{ width: `${max ? Math.max(8, Math.round((row.count / max) * 100)) : 0}%`, height: "100%", background: card.color, borderRadius: 999 }} />
                              </div>
                            </div>
                          ))}
                      </WCard>
                    );
                  })}
                </div>

                <div className="gaq-card" style={{ padding: 22, marginBottom: 12 }}>
                  <div className="gaq-section-h">
                    <div>
                      <div className="gaq-h3">Total de processos</div>
                      <div className="sub">Todos os processos do recorte atual — por situação e por modalidade</div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                      <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: "var(--gaq-text-3)", cursor: "pointer" }}>
                        <input type="checkbox" checked={ocultarProblematicosOverview}
                          onChange={e => setOcultarProblematicosOverview(e.target.checked)} />
                        Ocultar cancelados/fracassados
                      </label>
                      <span className="gaq-h2 gaq-num" style={{ fontSize: 22 }}>
                        {(ocultarProblematicosOverview
                          ? overviewAdmin.totalPorSituacao.filter(s => !s.problematico)
                          : overviewAdmin.totalPorSituacao).reduce((s, i) => s + i.count, 0)}
                      </span>
                    </div>
                  </div>
                  {(() => {
                    const rows = ocultarProblematicosOverview
                      ? overviewAdmin.totalPorSituacao.filter(s => !s.problematico)
                      : overviewAdmin.totalPorSituacao;
                    if (!rows.length) return <div style={{ color: "var(--gaq-text-3)", fontSize: 12 }}>Sem dados no recorte atual.</div>;
                    return (
                      <Recharts.ResponsiveContainer width="100%" height={240}>
                        <Recharts.BarChart data={rows} margin={{ top: 20, right: 10, left: 0, bottom: 5 }}>
                          <Recharts.CartesianGrid vertical={false} stroke="var(--gaq-line)" />
                          <Recharts.XAxis dataKey="label" fontSize={11} tick={{ fill: "var(--gaq-text-3)" }} axisLine={{ stroke: "var(--gaq-line)" }} tickLine={false} />
                          <Recharts.YAxis fontSize={10} tick={{ fill: "var(--gaq-text-3)" }} axisLine={false} tickLine={false} allowDecimals={false} />
                          <Recharts.Tooltip formatter={(v) => [v, "Processos"]} labelStyle={{ color: "#1d1d1f" }} cursor={{ fill: "rgba(0,0,0,0.04)" }} />
                          <Recharts.Bar dataKey="count" name="Processos" radius={[4, 4, 0, 0]} maxBarSize={64} isAnimationActive={false}>
                            {rows.map(row => (
                              <Recharts.Cell key={row.key} fill={row.color} style={{ cursor: "pointer" }}
                                onClick={() => setDrillDown({ title: `Total de processos — ${row.label}`, data: row.rows, color: row.color })} />
                            ))}
                            <Recharts.LabelList dataKey="count" position="top" style={{ fontSize: 12, fontWeight: 600, fill: "var(--gaq-text)" }} />
                          </Recharts.Bar>
                        </Recharts.BarChart>
                      </Recharts.ResponsiveContainer>
                    );
                  })()}

                  <div style={{ fontWeight: 600, fontSize: 13, color: "var(--gaq-text)", margin: "18px 0 4px", paddingTop: 14, borderTop: "1px solid var(--gaq-line)" }}>
                    Por modalidade
                  </div>
                  {(() => {
                    const rows = overviewAdmin.totalPorModalidade
                      .map(m => ocultarProblematicosOverview
                        ? { label: m.label, count: m.countAtivo, rows: m.rowsAtivo }
                        : { label: m.label, count: m.count, rows: m.rows })
                      .filter(m => m.count > 0)
                      .sort((a, b) => b.count - a.count);
                    if (!rows.length) return <div style={{ color: "var(--gaq-text-3)", fontSize: 12 }}>Sem dados no recorte atual.</div>;
                    return (
                      <Recharts.ResponsiveContainer width="100%" height={240}>
                        <Recharts.BarChart data={rows} margin={{ top: 20, right: 10, left: 0, bottom: 5 }}>
                          <Recharts.CartesianGrid vertical={false} stroke="var(--gaq-line)" />
                          <Recharts.XAxis dataKey="label" fontSize={11} tick={{ fill: "var(--gaq-text-3)" }} axisLine={{ stroke: "var(--gaq-line)" }} tickLine={false} />
                          <Recharts.YAxis fontSize={10} tick={{ fill: "var(--gaq-text-3)" }} axisLine={false} tickLine={false} allowDecimals={false} />
                          <Recharts.Tooltip formatter={(v) => [v, "Processos"]} labelStyle={{ color: "#1d1d1f" }} cursor={{ fill: "rgba(0,0,0,0.04)" }} />
                          <Recharts.Bar dataKey="count" name="Processos" fill="var(--gaq-blue)" radius={[4, 4, 0, 0]} maxBarSize={64} isAnimationActive={false}
                            onClick={(_, i) => rows[i] && setDrillDown({ title: `Total de processos — ${rows[i].label}`, data: rows[i].rows, color: "var(--gaq-blue)" })}
                            style={{ cursor: "pointer" }}>
                            <Recharts.LabelList dataKey="count" position="top" style={{ fontSize: 12, fontWeight: 600, fill: "var(--gaq-text)" }} />
                          </Recharts.Bar>
                        </Recharts.BarChart>
                      </Recharts.ResponsiveContainer>
                    );
                  })()}
                </div>
              </div>
              </>}
              {/* ── Visão da Área (perfil área bloqueado) ── */}
              {loginArea && areaAnoData.length > 0 && (() => {
                const areaAllProcs = base.filter(r => areaMatch(r["Área Requisitante"] || r.Area || r["ÁREA"] || "", loginArea));
                const areaActiveRows = areaAllProcs.filter(r => r.emA);
                const areaDoneRows = areaAllProcs.filter(r => !r.emA);
                const totalAnd = areaActiveRows.length;
                const totalConc = areaDoneRows.length;
                const criticos = areaActiveRows.filter(r => calcSLAClassification(r, phaseIntervals)?.bucket === "critico").length;
                const concComDias = areaDoneRows.filter(r => r.diasTotais > 0);
                const mediaGeral = concComDias.length
                  ? Math.round(concComDias.reduce((s,r) => s + r.diasTotais, 0) / concComDias.length)
                  : null;
                const anoAtual = new Date().getFullYear();
                const dAtual = areaAnoData.find(d => d.ano == anoAtual);
                const dAnterior = areaAnoData.find(d => d.ano == anoAtual - 1);
                const tendencia = (dAtual?.media && dAnterior?.media)
                  ? (dAtual.media < dAnterior.media ? { txt: `↓ ${dAnterior.media - dAtual.media} d.u. mais rápido`, cor: "#27ae60" }
                     : dAtual.media > dAnterior.media ? { txt: `↑ ${dAtual.media - dAnterior.media} d.u. mais lento`, cor: "#c0392b" }
                     : { txt: "→ Estável", cor: "#f39c12" }) : null;
                return (
                  <div className="anim-card anim-d4" style={{ marginBottom: 12 }}>
                    {/* Header da área */}
                    <div style={{ background: "linear-gradient(135deg, #4a235a 0%, #8e44ad 100%)", borderRadius: 12, padding: "16px 22px", marginBottom: 12, color: "#fff", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                        <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0, opacity: 0.95 }}>
                          <path d="M3 21h18"/><path d="M5 21V8l7-4 7 4v13"/><path d="M9 21v-6h6v6"/><path d="M9 11h.01"/><path d="M15 11h.01"/>
                        </svg>
                        <div>
                          <div style={{ fontSize: 16, fontWeight: 800, letterSpacing: "-0.01em" }}>Área {loginAreaDisp} — Acompanhamento de Processos</div>
                          <div style={{ fontSize: 10, opacity: .75, marginTop: 3 }}>Visualização restrita · Processos da área {loginAreaDisp} · Painel de Dados de Compras GAQ</div>
                        </div>
                      </div>
                      {tendencia && <div style={{ background: "rgba(255,255,255,0.15)", borderRadius: 8, padding: "6px 14px", fontSize: 12, fontWeight: 700 }}>{tendencia.txt}</div>}
                    </div>
                    {/* KPIs resumo */}
                    <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 10, marginBottom: 12 }}>
                      {[
                        { label: "Em andamento", val: totalAnd, color: "#2e86c1",
                          icon: <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#2e86c1" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="9"/><polyline points="12 7 12 12 15.5 14"/></svg> },
                        { label: "Concluídos (total)", val: totalConc, color: "#27ae60",
                          icon: <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#27ae60" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="9"/><polyline points="8 12.5 11 15.5 16 9.5"/></svg> },
                        { label: "Críticos SLA + entrega", val: criticos, color: "#c0392b",
                          icon: <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#c0392b" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"><path d="M12 3 L21.5 20 L2.5 20 Z"/><line x1="12" y1="10" x2="12" y2="14"/><line x1="12" y1="17" x2="12" y2="17.01"/></svg> },
                        { label: "Média d.u. (concluídos)", val: mediaGeral !== null ? mediaGeral + " d.u." : "—", color: "#8e44ad",
                          icon: <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#8e44ad" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="13" r="8"/><path d="M9 2h6"/><path d="M12 9v4l2.5 1.5"/></svg> },
                      ].map(k => (
                        <div key={k.label} style={{ background: "var(--card)", borderRadius: 10, padding: "14px 16px", boxShadow: "var(--shadow)", borderTop: `3px solid ${k.color}`, textAlign: "center" }}>
                          <div style={{ display: "flex", justifyContent: "center", marginBottom: 6, height: 22 }}>{k.icon}</div>
                          <div style={{ fontSize: 22, fontWeight: 800, color: k.color }}>{k.val}</div>
                          <div style={{ fontSize: 10, color: "var(--text3)", marginTop: 2 }}>{k.label}</div>
                        </div>
                      ))}
                    </div>
                    {/* Histórico por ano */}
                    <div style={{ background: "var(--card)", borderRadius: 10, padding: "16px 18px", boxShadow: "var(--shadow)" }}>
                      <div style={{ fontWeight: 700, fontSize: 13, color: "#8e44ad", marginBottom: 12, display: "flex", alignItems: "center", gap: 8 }}>
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#8e44ad" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round"><line x1="3" y1="20" x2="21" y2="20"/><rect x="5" y="11" width="3.5" height="9"/><rect x="10.25" y="6" width="3.5" height="14"/><rect x="15.5" y="3" width="3.5" height="17"/></svg>
                        Histórico de Processos por Ano — Área {loginAreaDisp}
                      </div>
                      <div style={{ overflowX: "auto" }}>
                        <div className="gaq-table-scroll"><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                          <thead>
                            <tr style={{ background: "#8e44ad11" }}>
                              {["Ano","Total","Concluídos","Em andamento","Média d.u. (concluídos)","Modalidade principal","Tendência"].map(h => (
                                <th key={h} style={{ padding: "8px 12px", textAlign: "left", fontWeight: 700, color: "var(--text2)", fontSize: 11, borderBottom: "2px solid #8e44ad33" }}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {areaAnoData.map((d, i) => {
                              const prev = areaAnoData[i - 1];
                              let tend = null;
                              if (prev && d.media !== null && prev.media !== null) {
                                if (d.media < prev.media) tend = { txt: `↓ ${prev.media - d.media} d.u.`, cor: "#27ae60" };
                                else if (d.media > prev.media) tend = { txt: `↑ ${d.media - prev.media} d.u.`, cor: "#c0392b" };
                                else tend = { txt: "→", cor: "#f39c12" };
                              }
                              const isAtual = d.ano == anoAtual;
                              return (
                                <tr key={d.ano} style={{ background: isAtual ? "#8e44ad08" : "transparent", borderBottom: "1px solid var(--border2)" }}>
                                  <td style={{ padding: "8px 12px", fontWeight: 800, color: isAtual ? "#8e44ad" : "var(--text)" }}>{d.ano}{isAtual ? " ★" : ""}</td>
                                  <td style={{ padding: "8px 12px", fontWeight: 700 }}>{d.total}</td>
                                  <td style={{ padding: "8px 12px", color: "#27ae60", fontWeight: 600 }}>{d.conc}</td>
                                  <td style={{ padding: "8px 12px", color: "#2e86c1" }}>{d.andamento}</td>
                                  <td style={{ padding: "8px 12px" }}>
                                    {d.media !== null
                                      ? <span style={{ fontWeight: 700, color: d.media > 100 ? "#c0392b" : d.media > 60 ? "#e67e22" : "#27ae60" }}>{d.media} d.u.</span>
                                      : <span style={{ color: "var(--text3)" }}>—</span>}
                                  </td>
                                  <td style={{ padding: "8px 12px", color: "var(--text2)" }}>{d.topMod}</td>
                                  <td style={{ padding: "8px 12px" }}>{tend ? <span style={{ fontWeight: 700, color: tend.cor }}>{tend.txt}</span> : <span style={{ color: "var(--text3)" }}>—</span>}</td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table></div>
                      </div>
                      <div style={{ fontSize: 10, color: "var(--text3)", marginTop: 8 }}>★ Ano atual · Tendência: variação da média d.u. em relação ao ano anterior (↓ melhorou · ↑ piorou)</div>
                    </div>
                  </div>
                );
              })()}

              {/* NCL / CPL / Scont removidos — informação disponível na aba Gestão de Processos */}

              {loginUser && (() => {
                const agingFaixas = [
                  { faixa: "0–30 d.u.", icon: "🟢", data: fBack.filter(r => r.diasTotais <= 30), color: "#27ae60", bg: "#eafaf1", desc: "No prazo" },
                  { faixa: "31–50 d.u.", icon: "🟡", data: fBack.filter(r => r.diasTotais > 30 && r.diasTotais <= 50), color: "#f39c12", bg: "#fef9e7", desc: "Atenção" },
                  { faixa: "51–100 d.u.", icon: "🟠", data: fBack.filter(r => r.diasTotais > 50 && r.diasTotais <= 100), color: "#e67e22", bg: "#fdf2e9", desc: "Risco" },
                  { faixa: "> 100 d.u.", icon: "🔴", data: fBack.filter(r => r.diasTotais > 100), color: "#c0392b", bg: "#fdedec", desc: "Crítico" },
                ];
                const faseMap = {};
                fBack.forEach(r => { const f = r.faseAtual || "—"; if (!faseMap[f]) faseMap[f] = []; faseMap[f].push(r); });
                const faseEntries = Object.entries(faseMap).sort((a,b) => b[1].length - a[1].length);
                const faseIcons = { "Recebimento DJ": "📥", "CPL: Recebido do NCL": "📥", "CPL: Enviado p/ DJS": "📤", "CPL: Recebimento": "🏛", "Encerramento SD": "📋", "Última aprovação": "✅", "Planejamento RC": "📝", "Início propostas": "📨", "Recebimento RC": "📬", "Fim propostas": "📊", "Envio aprovação": "📤" };
                const maxFase = faseEntries.length > 0 ? faseEntries[0][1].length : 1;
                return (
                  <div className="anim-card anim-d4 gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
                    {/* Card Aging */}
                    <div style={{ background: "var(--card)", borderRadius: 14, boxShadow: "var(--shadow)", overflow: "hidden" }}>
                      <div style={{ background: "linear-gradient(135deg, #1a5276 0%, #2e86c1 100%)", padding: "14px 18px 12px" }}>
                        <div style={{ fontSize: 13, fontWeight: 800, color: "#fff", letterSpacing: 0.2 }}>⏱ Meu Aging</div>
                        <div style={{ fontSize: 10, color: "rgba(255,255,255,0.75)", marginTop: 2 }}>Clique em uma faixa para ver os processos</div>
                      </div>
                      <div style={{ padding: "10px 14px 14px" }}>
                        {agingFaixas.map(({ faixa, icon, data, color, bg, desc }) => (
                          <div key={faixa}
                            onClick={() => { setDrillDown({ title: "Aging: " + faixa + " — " + desc, data, color, sourceAba: "overview" }); }}
                            style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 10px", marginBottom: 6, borderRadius: 9, background: bg, border: "1.5px solid " + color + "33", cursor: "pointer", transition: "box-shadow .15s, transform .15s" }}
                            onMouseEnter={e => { e.currentTarget.style.boxShadow = "0 3px 12px " + color + "44"; e.currentTarget.style.transform = "translateY(-1px)"; }}
                            onMouseLeave={e => { e.currentTarget.style.boxShadow = "none"; e.currentTarget.style.transform = "none"; }}>
                            <div style={{ fontSize: 16, flexShrink: 0 }}>{icon}</div>
                            <div style={{ flex: 1 }}>
                              <div style={{ fontSize: 12, fontWeight: 700, color: "var(--text)" }}>{faixa}</div>
                              <div style={{ height: 5, borderRadius: 3, background: "#e0e0e0", marginTop: 4 }}>
                                <div style={{ height: "100%", borderRadius: 3, background: color, width: fBack.length > 0 ? (data.length / fBack.length * 100) + "%" : "0%", transition: "width .4s" }} />
                              </div>
                            </div>
                            <div style={{ textAlign: "right", flexShrink: 0 }}>
                              <div style={{ fontSize: 20, fontWeight: 900, color, lineHeight: 1 }}>{data.length}</div>
                              <div style={{ fontSize: 9, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase" }}>{desc}</div>
                            </div>
                            <div style={{ color: color, fontSize: 14, opacity: .6 }}>›</div>
                          </div>
                        ))}
                        <div style={{ marginTop: 4, padding: "7px 12px", background: "var(--card2)", borderRadius: 8, fontSize: 11, color: "var(--text2)", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                          <span>Total em andamento</span>
                          <b style={{ fontSize: 15, color: "#2e86c1" }}>{fBack.length}</b>
                        </div>
                      </div>
                    </div>
                    {/* Card Fases */}
                    <div style={{ background: "var(--card)", borderRadius: 14, boxShadow: "var(--shadow)", overflow: "hidden" }}>
                      <div style={{ background: "linear-gradient(135deg, #1e8449 0%, #27ae60 100%)", padding: "14px 18px 12px" }}>
                        <div style={{ fontSize: 13, fontWeight: 800, color: "#fff", letterSpacing: 0.2 }}>📋 Meus Processos por Fase</div>
                        <div style={{ fontSize: 10, color: "rgba(255,255,255,0.75)", marginTop: 2 }}>Clique em uma fase para ver os processos</div>
                      </div>
                      <div style={{ padding: "10px 14px 14px", maxHeight: 320, overflowY: "auto" }}>
                        {faseEntries.length === 0
                          ? <div style={{ color: "var(--text3)", fontSize: 12, padding: 10 }}>Nenhum processo em andamento.</div>
                          : faseEntries.slice(0, 12).map(([fase, procs]) => (
                            <div key={fase}
                              onClick={() => { setDrillDown({ title: "Fase: " + fase, data: procs, color: "#1e8449", sourceAba: "overview" }); }}
                              style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 10px", marginBottom: 5, borderRadius: 9, background: "var(--card2)", border: "1px solid var(--border2)", cursor: "pointer", transition: "box-shadow .15s, transform .15s" }}
                              onMouseEnter={e => { e.currentTarget.style.boxShadow = "0 3px 10px #27ae6033"; e.currentTarget.style.transform = "translateY(-1px)"; e.currentTarget.style.background = "#eafaf1"; }}
                              onMouseLeave={e => { e.currentTarget.style.boxShadow = "none"; e.currentTarget.style.transform = "none"; e.currentTarget.style.background = "var(--card2)"; }}>
                              <div style={{ fontSize: 14, flexShrink: 0 }}>{faseIcons[fase] || "📌"}</div>
                              <div style={{ flex: 1 }}>
                                <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{fase}</div>
                                <div style={{ height: 4, borderRadius: 2, background: "#e0e0e0", marginTop: 3 }}>
                                  <div style={{ height: "100%", borderRadius: 2, background: "#27ae60", width: (procs.length / maxFase * 100) + "%", transition: "width .4s" }} />
                                </div>
                              </div>
                              <div style={{ fontSize: 14, fontWeight: 800, color: "#1e8449", flexShrink: 0, background: "#d5f5e3", borderRadius: 7, padding: "2px 10px", minWidth: 28, textAlign: "center" }}>{procs.length}</div>
                              <div style={{ color: "#27ae60", fontSize: 14, opacity: .6 }}>›</div>
                            </div>
                          ))
                        }
                      </div>
                    </div>
                  </div>
                );
              })()}
            </>}
            </>);
}
