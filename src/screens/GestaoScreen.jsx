import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function GestaoScreen({ alertCPL, alertNCL, alertScont, base, cplByMod, fBack, gestaoPassErr, gestaoPassInput, gestaoUnlocked, nclByMod, onClickComp, phaseIntervals, rankCPLConc, rankCPLPerf, rankGeralNCL, rankNCL, rankNCLAgingSD, rankNCLAgingSDAtivo, rankNCLConc, rankScontConc, rankScontPerf, scontBack, scontByMod, scontPerfBack, setAba, setDrillDown, setGestaoPassErr, setGestaoPassInput, setGestaoUnlocked, setSelProc }) {
  return (<div className="anim-fade">
          {!gestaoUnlocked ? (
            <div style={{ maxWidth: 400, margin: "80px auto", background: "var(--card)", borderRadius: 16, padding: 36, boxShadow: "var(--shadow)", textAlign: "center" }}>
              <div style={{ fontSize: 32, marginBottom: 12 }}>📊</div>
              <div style={{ fontSize: 18, fontWeight: 800, color: "var(--text)", marginBottom: 6 }}>Painel de Gestão</div>
              <div style={{ fontSize: 12, color: "var(--text3)", marginBottom: 24 }}>Acesso restrito. Digite a senha do painel de gestão.</div>
              <input type="password" placeholder="Senha do painel" value={gestaoPassInput} onChange={e => { setGestaoPassInput(e.target.value); setGestaoPassErr(false); }}
                onKeyDown={async e => { if (e.key === "Enter") { const h = await gaqHashSenha(gestaoPassInput); if (typeof GESTAO_PASS !== "undefined" && h === GESTAO_PASS) { setGestaoUnlocked(true); setGestaoPassInput(""); } else setGestaoPassErr(true); }}}
                style={{ width: "100%", padding: "10px 14px", borderRadius: 8, border: gestaoPassErr ? "2px solid #c0392b" : "1.5px solid var(--input-bd)", background: "var(--input-bg)", color: "var(--text)", fontSize: 14, marginBottom: 8, boxSizing: "border-box" }} />
              {gestaoPassErr && <div style={{ color: "#c0392b", fontSize: 12, marginBottom: 8 }}>Senha incorreta.</div>}
              <button onClick={async () => { const h = await gaqHashSenha(gestaoPassInput); if (typeof GESTAO_PASS !== "undefined" && h === GESTAO_PASS) { setGestaoUnlocked(true); setGestaoPassInput(""); } else setGestaoPassErr(true); }}
                style={{ background: "#2e86c1", color: "#fff", border: "none", borderRadius: 8, padding: "10px 28px", fontWeight: 700, fontSize: 14, cursor: "pointer", width: "100%" }}>Entrar</button>
            </div>
          ) : (
            <div>
              <div style={{ background: "linear-gradient(135deg, #1a2a1a 0%, #1e8449 100%)", borderRadius: 14, padding: "20px 28px", marginBottom: 18, color: "#fff", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <div style={{ fontSize: 20, fontWeight: 800 }}>📊 Painel de Gestão — Performance por Subárea</div>
                  <div style={{ fontSize: 11, opacity: .7, marginTop: 4 }}>NCL · CPL · Scont · Acesso restrito · v5.0</div>
                </div>
                <button onClick={() => setGestaoUnlocked(false)} style={{ background: "rgba(255,255,255,0.15)", border: "1px solid rgba(255,255,255,0.3)", borderRadius: 8, padding: "6px 14px", color: "#fff", cursor: "pointer", fontSize: 11, fontWeight: 700 }}>🔒 Sair</button>
              </div>

              {/* KPIs resumo */}
              <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12, marginBottom: 18 }}>
                {[
                  { label: "Em fase NCL", count: fBack.filter(r => r.faseSubarea === "NCL").length, criticos: alertNCL.length, color: "#2e86c1" },
                  { label: "Em fase CPL", count: fBack.filter(r => r.faseSubarea === "CPL").length, criticos: alertCPL.length, color: "#8e44ad" },
                  { label: "Em fase Scont", count: scontBack.length, criticos: alertScont.length, color: "#16a085" },
                ].map(k => (
                  <div key={k.label} style={{ background: "var(--card)", borderRadius: 12, padding: "16px 18px", boxShadow: "var(--shadow)", borderTop: `4px solid ${k.color}` }}>
                    <div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase", letterSpacing: .4, marginBottom: 4 }}>{k.label}</div>
                    <div style={{ fontSize: 28, fontWeight: 800, color: k.color }}>{k.count}</div>
                    <div style={{ fontSize: 11, color: "#c0392b", fontWeight: 700 }}>⚠ {k.criticos} crítico(s) &gt;50 d.u.</div>
                  </div>
                ))}
              </div>

              {/* Performance NCL */}
              <div style={{ background: "var(--card)", borderRadius: 12, padding: 18, boxShadow: "var(--shadow)", marginBottom: 14, borderLeft: "4px solid #2e86c1" }}>
                <div style={{ fontWeight: 800, fontSize: 14, color: "#2e86c1", marginBottom: 4 }}>📋 Performance NCL — Compradores / Avaliadores</div>
                <div style={{ fontSize: 10, color: "#2e86c1", opacity: .7, marginBottom: 4, fontStyle: "italic" }}>Processos em Andamento (fase NCL) · Aging contado a partir do Recebimento RC</div>
                <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 16, marginBottom: 14 }}>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#c0392b", marginBottom: 6 }}>Críticos (&gt;50 d.u. desde RC)</div>
                    {rankNCL.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Nenhum.</div>
                      : rankNCL.map((x, i) => <RankRow key={x.name} rank={i+1} name={x.name} badge={x.criticos} badgeColor="#c0392b" extra={`/${x.total}`} onClick={() => onClickComp(x.name, "ncl")} />)}
                  </div>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#2e86c1", marginBottom: 6 }}>Média d.u. RC — Andamento</div>
                    {rankGeralNCL.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Sem dados.</div>
                      : rankGeralNCL.map((x, i) => <RankRow key={x.name} rank={i+1} name={x.name} badge={`${x.media} d.u.`} badgeColor={x.media > 100 ? "#c0392b" : x.media > 60 ? "#e67e22" : "#1e8449"} onClick={() => onClickComp(x.name, "ncl")} />)}
                  </div>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#1a5276", marginBottom: 6 }}>Média d.u. RC — Concluídos</div>
                    {rankNCLConc.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Sem dados.</div>
                      : rankNCLConc.map((x, i) => <RankRow key={x.name} rank={i+1} name={x.name} badge={`${x.media} d.u.`} extra={x.total < 15 ? `/${x.total} ⚠` : `/${x.total}`} badgeColor={x.media > 100 ? "#c0392b" : x.media > 60 ? "#e67e22" : "#1e8449"} onClick={() => onClickComp(x.name, "ncl")} />)}
                    <div style={{ fontSize: 9, color: "var(--text3)", marginTop: 4 }}>⚠ = amostra abaixo de 15 processos</div>
                  </div>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#5d6d7e", marginBottom: 6 }}>Por Modalidade — RC (Concluídos)</div>
                    {nclByMod.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Sem dados.</div>
                      : nclByMod.map((x, i) => <RankRow key={x.name} rank={i+1} name={x.name} badge={`${x.media} d.u.`} extra={`/${x.total}`} badgeColor={x.media > 100 ? "#c0392b" : x.media > 60 ? "#e67e22" : "#1e8449"} />)}
                  </div>
                </div>
                {/* Aging pré-compra — Abertura → Encerramento */}
                <div style={{ borderTop: "1px dashed #b2d4ef", paddingTop: 12, marginTop: 4 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#1a5276", marginBottom: 8 }}>🎫 Aging pré-compra — Abertura → Encerramento (por comprador)</div>
                  <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                    <div>
                      <div style={{ fontSize: 10, fontWeight: 700, color: "#7f8c8d", marginBottom: 6 }}>Média d.u. pré-compra — Concluídos</div>
                      {rankNCLAgingSD.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Sem dados.</div>
                        : rankNCLAgingSD.slice(0, 8).map((x, i) => <RankRow key={x.name} rank={i+1} name={x.name} badge={`${x.media} d.u.`} extra={`/${x.total}`} badgeColor={x.media > 15 ? "#c0392b" : x.media > 8 ? "#e67e22" : "#1e8449"} onClick={() => onClickComp(x.name, "ncl")} />)}
                      <div style={{ fontSize: 9, color: "var(--text3)", marginTop: 4 }}>Contagem: abertura até encerramento da pré-compra</div>
                    </div>
                    <div>
                      <div style={{ fontSize: 10, fontWeight: 700, color: "#7f8c8d", marginBottom: 6 }}>Média d.u. pré-compra — Em andamento</div>
                      {rankNCLAgingSDAtivo.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Sem dados.</div>
                        : rankNCLAgingSDAtivo.slice(0, 8).map((x, i) => <RankRow key={x.name} rank={i+1} name={x.name} badge={`${x.media} d.u.`} extra={`/${x.total}`} badgeColor={x.media > 15 ? "#c0392b" : x.media > 8 ? "#e67e22" : "#1e8449"} onClick={() => onClickComp(x.name, "ncl")} />)}
                      <div style={{ fontSize: 9, color: "var(--text3)", marginTop: 4 }}>Contagem: abertura da pré-compra até hoje</div>
                    </div>
                  </div>
                </div>
              </div>

              {/* Performance CPL */}
              <div style={{ background: "var(--card)", borderRadius: 12, padding: 18, boxShadow: "var(--shadow)", marginBottom: 14, borderLeft: "4px solid #8e44ad" }}>
                <div style={{ fontWeight: 800, fontSize: 14, color: "#8e44ad", marginBottom: 4 }}>🏛 Performance CPL — Responsáveis CPL</div>
                <div style={{ fontSize: 10, color: "#8e44ad", opacity: .7, marginBottom: 12, fontStyle: "italic" }}>Processos em Andamento (CPL_ENCONTRADO = Verdadeiro)</div>
                <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 16 }}>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#c0392b", marginBottom: 6 }}>Críticos (&gt;50 d.u.)</div>
                    {rankCPLPerf.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Sem dados CPL.</div>
                      : [...rankCPLPerf].sort((a,b) => b.criticos - a.criticos).slice(0, 10)
                          .map((x, i) => <RankRow key={x.name} rank={i+1} name={x.name} badge={x.criticos} badgeColor={x.criticos > 0 ? "#c0392b" : "#7f8c8d"} extra={`/${x.total}`} onClick={() => onClickComp(x.name, "cpl")} />)}
                  </div>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#8e44ad", marginBottom: 6 }}>Média d.u. — Andamento</div>
                    {rankCPLPerf.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Sem dados.</div>
                      : [...rankCPLPerf].sort((a,b) => b.media - a.media).map((x, i) => <RankRow key={x.name} rank={i+1} name={x.name} badge={`${x.media} d.u.`} badgeColor={x.media > 100 ? "#c0392b" : x.media > 60 ? "#e67e22" : "#1e8449"} onClick={() => onClickComp(x.name, "cpl")} />)}
                  </div>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#6c3483", marginBottom: 6 }}>Média d.u. — Concluídos</div>
                    {rankCPLConc.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Sem dados.</div>
                      : rankCPLConc.map((x, i) => <RankRow key={x.name} rank={i+1} name={x.name} badge={`${x.media} d.u.`} extra={x.total < 15 ? `/${x.total} ⚠` : `/${x.total}`} badgeColor={x.media > 100 ? "#c0392b" : x.media > 60 ? "#e67e22" : "#1e8449"} onClick={() => onClickComp(x.name, "cpl")} />)}
                    <div style={{ fontSize: 9, color: "var(--text3)", marginTop: 4 }}>⚠ = amostra abaixo de 15 processos</div>
                  </div>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#5d6d7e", marginBottom: 6 }}>Por Modalidade (Concluídos)</div>
                    {cplByMod.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Sem dados.</div>
                      : cplByMod.map((x, i) => <RankRow key={x.name} rank={i+1} name={x.name} badge={`${x.media} d.u.`} extra={`/${x.total}`} badgeColor={x.media > 100 ? "#c0392b" : x.media > 60 ? "#e67e22" : "#1e8449"} />)}
                  </div>
                </div>
              </div>

              {/* Performance SCONT */}
              <div style={{ background: "var(--card)", borderRadius: 12, padding: 18, boxShadow: "var(--shadow)", marginBottom: 14, borderLeft: "4px solid #16a085" }}>
                <div style={{ fontWeight: 800, fontSize: 14, color: "#16a085", marginBottom: 4 }}>📝 Performance Scont — Analistas de Contrato</div>
                <div style={{ fontSize: 10, color: "#16a085", opacity: .7, marginBottom: 12, fontStyle: "italic" }}>Processos em Andamento com Analista de Contrato</div>
                <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 16 }}>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#c0392b", marginBottom: 6 }}>Críticos (&gt;50 d.u.)</div>
                    {rankScontPerf.filter(x => x.criticos > 0).length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Nenhum.</div>
                      : rankScontPerf.filter(x => x.criticos > 0).sort((a,b) => b.criticos - a.criticos)
                          .map((x, i) => <RankRow key={x.name} rank={i+1} name={x.name} badge={x.criticos} badgeColor="#c0392b" extra={`/${x.total}`} onClick={() => { setAba("overview"); setDrillDown({ title: "Scont — " + x.name, data: scontPerfBack.filter(r => r.AnalistaContrato === x.name), color: "#16a085", sourceAba: "gestao" }); }} />)}
                  </div>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#16a085", marginBottom: 6 }}>Média d.u. — Andamento</div>
                    {rankScontPerf.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Sem dados.</div>
                      : [...rankScontPerf].sort((a,b) => b.media - a.media).map((x, i) => <RankRow key={x.name} rank={i+1} name={x.name} badge={`${x.media} d.u.`} badgeColor={x.media > 100 ? "#c0392b" : x.media > 60 ? "#e67e22" : "#1e8449"} onClick={() => { setAba("overview"); setDrillDown({ title: "Scont — " + x.name, data: scontPerfBack.filter(r => r.AnalistaContrato === x.name), color: "#16a085", sourceAba: "gestao" }); }} />)}
                  </div>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#0e6655", marginBottom: 6 }}>Média d.u. — Concluídos</div>
                    {rankScontConc.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Sem dados.</div>
                      : rankScontConc.map((x, i) => <RankRow key={x.name} rank={i+1} name={x.name} badge={`${x.media} d.u.`} extra={x.total < 15 ? `/${x.total} ⚠` : `/${x.total}`} badgeColor={x.media > 100 ? "#c0392b" : x.media > 60 ? "#e67e22" : "#1e8449"} />)}
                    <div style={{ fontSize: 9, color: "var(--text3)", marginTop: 4 }}>⚠ = amostra abaixo de 15 processos</div>
                  </div>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#5d6d7e", marginBottom: 6 }}>Por Modalidade (Concluídos)</div>
                    {scontByMod.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 12 }}>Sem dados.</div>
                      : scontByMod.map((x, i) => <RankRow key={x.name} rank={i+1} name={x.name} badge={`${x.media} d.u.`} extra={`/${x.total}`} badgeColor={x.media > 100 ? "#c0392b" : x.media > 60 ? "#e67e22" : "#1e8449"} />)}
                  </div>
                </div>
              </div>

              {/* ══════════════════════════════════════════════════════
                  FASE 3 - VISAO ESTRATEGICA (admin only)
                  ══════════════════════════════════════════════════════ */}
              {(() => {
                const MESES = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];

                // ── 1. Mapa de Calor: Área × Mês ──────────────────────
                const areaVol = {};
                base.forEach(r => { const a = r["Área Requisitante"]; if (a) areaVol[a] = (areaVol[a]||0)+1; });
                const topAreas = Object.entries(areaVol).sort((a,b)=>b[1]-a[1]).slice(0,8).map(([a])=>a);
                const heatRows = topAreas.map(area => ({
                  area,
                  counts: MESES.map((_,m) => base.filter(r => r["Área Requisitante"]===area && r.mesAbertura===m).length)
                }));
                const heatMax = Math.max(1, ...heatRows.flatMap(r=>r.counts));

                // ── 2. Score de Complexidade — top processos ativos ───
                const complexAtivos = base.filter(r => r.emA).map(r => ({
                  ...r,
                  complexidade: calcComplexidade(r, phaseIntervals),
                  delayProb: calcDelayProb(r, phaseIntervals),
                })).sort((a,b)=>b.complexidade-a.complexidade || b.diasTotais-a.diasTotais).slice(0,15);

                // ── 3. Alerta Preditivo — processos com alto risco ────
                const alertas = base.filter(r => r.emA).map(r => ({
                  ...r,
                  delayProb: calcDelayProb(r, phaseIntervals),
                  complexidade: calcComplexidade(r, phaseIntervals),
                })).filter(r => r.delayProb !== null && r.delayProb >= 55)
                   .sort((a,b)=>b.delayProb-a.delayProb).slice(0,20);

                return (
                  <div style={{ marginTop: 28 }}>
                    {/* Header estratégico */}
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20, paddingBottom: 12, borderBottom: "2px solid var(--border)" }}>
                      <div style={{ fontSize: 22 }}>🔭</div>
                      <div>
                        <div style={{ fontSize: 15, fontWeight: 800, color: "var(--text)" }}>Visão Estratégica — Fase 3</div>
                        <div style={{ fontSize: 11, color: "var(--text3)" }}>Mapa de calor · Complexidade · Alertas preditivos · Apenas administrador</div>
                      </div>
                    </div>

                    {/* 1 — MAPA DE CALOR */}
                    <div style={{ background: "var(--card)", borderRadius: 12, padding: 18, boxShadow: "var(--shadow)", marginBottom: 20 }}>
                      <div style={{ fontWeight: 700, fontSize: 13, color: "var(--text)", marginBottom: 4 }}>🌡 Mapa de Calor — Demanda por Área e Mês</div>
                      <div style={{ fontSize: 11, color: "var(--text3)", marginBottom: 14 }}>Intensidade = nº de processos abertos naquele mês · Top 8 áreas por volume</div>
                      <div style={{ overflowX: "auto" }}>
                        <div className="gaq-table-scroll"><table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11 }}>
                          <thead>
                            <tr>
                              <th style={{ padding: "6px 10px", textAlign: "left", color: "var(--text3)", fontWeight: 700, fontSize: 10, textTransform: "uppercase", whiteSpace: "nowrap", minWidth: 100 }}>Área</th>
                              {MESES.map(m => (
                                <th key={m} style={{ padding: "6px 6px", textAlign: "center", color: "var(--text3)", fontWeight: 700, fontSize: 10, textTransform: "uppercase", minWidth: 38 }}>{m}</th>
                              ))}
                              <th style={{ padding: "6px 8px", textAlign: "center", color: "var(--text3)", fontWeight: 700, fontSize: 10, textTransform: "uppercase" }}>Total</th>
                            </tr>
                          </thead>
                          <tbody>
                            {heatRows.map(row => {
                              const rowTotal = row.counts.reduce((s,c)=>s+c,0);
                              return (
                                <tr key={row.area}>
                                  <td style={{ padding: "5px 10px", fontWeight: 600, color: "var(--text)", fontSize: 11, whiteSpace: "nowrap" }}>{row.area}</td>
                                  {row.counts.map((c, mi) => {
                                    const intensity = heatMax > 0 ? c / heatMax : 0;
                                    const bg = c === 0 ? "var(--card2)" : `rgba(46,134,193,${0.15 + intensity * 0.85})`;
                                    const color = intensity > 0.6 ? "#fff" : "var(--text)";
                                    return (
                                      <td key={mi} style={{ padding: "5px 4px", textAlign: "center", background: bg, color, fontWeight: c > 0 ? 700 : 400, borderRadius: 4, fontSize: 11 }}>
                                        {c > 0 ? c : "·"}
                                      </td>
                                    );
                                  })}
                                  <td style={{ padding: "5px 8px", textAlign: "center", fontWeight: 800, color: "#2e86c1", fontSize: 12 }}>{rowTotal}</td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table></div>
                      </div>
                      <div style={{ fontSize: 10, color: "var(--text3)", marginTop: 10 }}>
                        Escala: <span style={{ display: "inline-block", width: 60, height: 8, background: "linear-gradient(90deg,rgba(46,134,193,0.15),rgba(46,134,193,1))", borderRadius: 4, verticalAlign: "middle", margin: "0 4px" }} /> baixo → alto volume
                      </div>
                    </div>

                    {/* 2 — ALERTA PREDITIVO */}
                    <div style={{ background: "var(--card)", borderRadius: 12, padding: 18, boxShadow: "var(--shadow)", marginBottom: 20 }}>
                      <div style={{ fontWeight: 700, fontSize: 13, color: "var(--text)", marginBottom: 4 }}>⚡ Alerta Preditivo — Processos com Alto Risco de Atraso</div>
                      <div style={{ fontSize: 11, color: "var(--text3)", marginBottom: 14 }}>
                        Probabilidade calculada com base em: aging atual · estagnação · progresso de fases · risco histórico da modalidade · apenas processos ≥ 55%
                      </div>
                      {alertas.length === 0
                        ? <div style={{ textAlign: "center", padding: 24, color: "var(--text3)", fontSize: 13 }}>✅ Nenhum processo com risco elevado de atraso no momento.</div>
                        : <div style={{ overflowX: "auto" }}>
                            <div className="gaq-table-scroll"><table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12 }}>
                              <thead>
                                <tr style={{ background: "var(--card2)", borderBottom: "2px solid var(--border)" }}>
                                  {["Processo","Comprador","Área","Modalidade","Aging","Fase Atual","Complexidade","Risco Atraso"].map(h => (
                                    <th key={h} style={{ padding: "8px 10px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "var(--text2)", textTransform: "uppercase", letterSpacing: .3, whiteSpace: "nowrap" }}>{h}</th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody>
                                {alertas.map((r,i) => {
                                  const probColor = r.delayProb >= 80 ? "#c0392b" : r.delayProb >= 65 ? "#e67e22" : "#f39c12";
                                  const cxColor = r.complexidade >= 8 ? "#8e44ad" : r.complexidade >= 6 ? "#2e86c1" : "var(--text3)";
                                  return (
                                    <tr key={i} onClick={() => setSelProc(r)} style={{ borderBottom: "1px solid var(--border2)", cursor: "pointer", transition: "background .15s" }}
                                      onMouseEnter={e => e.currentTarget.style.background="var(--hover)"}
                                      onMouseLeave={e => e.currentTarget.style.background="transparent"}>
                                      <td style={{ padding: "8px 10px", fontWeight: 700, color: "#2e86c1", whiteSpace: "nowrap" }}>{r.NumRC || r.TicketSD || "—"}</td>
                                      <td style={{ padding: "8px 10px", color: "var(--text)" }}>{r.respNCL || r.Pregoeiro || "—"}</td>
                                      <td style={{ padding: "8px 10px", color: "var(--text2)", fontSize: 11 }}>{r["Área Requisitante"] || "—"}</td>
                                      <td style={{ padding: "8px 10px", color: "var(--text2)", fontSize: 11 }}>{r.Modalidade || "—"}</td>
                                      <td style={{ padding: "8px 10px", fontWeight: 700, color: r.diasTotais > 100 ? "#c0392b" : r.diasTotais > 50 ? "#e67e22" : "#27ae60", whiteSpace: "nowrap" }}>{r.diasTotais} d.u.</td>
                                      <td style={{ padding: "8px 10px", fontSize: 10, color: "var(--text2)", maxWidth: 120, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.faseAtual || "—"}</td>
                                      <td style={{ padding: "8px 10px" }}>
                                        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                          <div style={{ height: 6, width: 60, background: "var(--border2)", borderRadius: 3, overflow: "hidden" }}>
                                            <div style={{ height: "100%", width: (r.complexidade * 10) + "%", background: cxColor, borderRadius: 3 }} />
                                          </div>
                                          <span style={{ fontSize: 11, fontWeight: 700, color: cxColor }}>{r.complexidade}/10</span>
                                        </div>
                                      </td>
                                      <td style={{ padding: "8px 10px" }}>
                                        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                          <div style={{ height: 8, width: 70, background: "var(--border2)", borderRadius: 4, overflow: "hidden" }}>
                                            <div style={{ height: "100%", width: r.delayProb + "%", background: probColor, borderRadius: 4 }} />
                                          </div>
                                          <span style={{ fontSize: 12, fontWeight: 800, color: probColor, whiteSpace: "nowrap" }}>{r.delayProb}%</span>
                                        </div>
                                      </td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table></div>
                          </div>
                      }
                    </div>

                    {/* 3 — SCORE DE COMPLEXIDADE */}
                    <div style={{ background: "var(--card)", borderRadius: 12, padding: 18, boxShadow: "var(--shadow)", marginBottom: 20 }}>
                      <div style={{ fontWeight: 700, fontSize: 13, color: "var(--text)", marginBottom: 4 }}>🧩 Score de Complexidade — Processos em Andamento</div>
                      <div style={{ fontSize: 11, color: "var(--text3)", marginBottom: 14 }}>
                        Score 1–10 baseado em: modalidade · risco histórico de C/F · tamanho do objeto · presença de CPL · top 15 por complexidade
                      </div>
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 10 }}>
                        {complexAtivos.map((r,i) => {
                          const cxColor = r.complexidade >= 8 ? "#8e44ad" : r.complexidade >= 6 ? "#2e86c1" : r.complexidade >= 4 ? "#27ae60" : "#7f8c8d";
                          const probColor = (r.delayProb||0) >= 75 ? "#c0392b" : (r.delayProb||0) >= 55 ? "#e67e22" : "#27ae60";
                          return (
                            <div key={i} onClick={() => setSelProc(r)} style={{ background: "var(--card2)", borderRadius: 10, padding: "12px 14px", cursor: "pointer", borderLeft: "4px solid " + cxColor, transition: "transform .15s" }}
                              onMouseEnter={e => e.currentTarget.style.transform="translateY(-2px)"}
                              onMouseLeave={e => e.currentTarget.style.transform="none"}>
                              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 6 }}>
                                <div>
                                  <div style={{ fontWeight: 700, color: "#2e86c1", fontSize: 12 }}>{r.NumRC || r.TicketSD || "—"}</div>
                                  <div style={{ fontSize: 10, color: "var(--text3)" }}>{r["Área Requisitante"] || "—"} · {r.Modalidade || "—"}</div>
                                </div>
                                <div style={{ textAlign: "right" }}>
                                  <div style={{ fontSize: 18, fontWeight: 800, color: cxColor }}>{r.complexidade}<span style={{ fontSize: 10, color: "var(--text3)" }}>/10</span></div>
                                  <div style={{ fontSize: 9, color: "var(--text3)", textTransform: "uppercase" }}>complexidade</div>
                                </div>
                              </div>
                              <div style={{ fontSize: 10, color: "var(--text2)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", marginBottom: 8 }}>{(r.Objeto||"").slice(0,70)||"—"}</div>
                              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
                                <div>
                                  <div style={{ fontSize: 9, color: "var(--text3)", marginBottom: 2 }}>Complexidade</div>
                                  <div style={{ height: 5, width: 80, background: "var(--border)", borderRadius: 3, overflow: "hidden" }}>
                                    <div style={{ height: "100%", width: (r.complexidade*10)+"%", background: cxColor, borderRadius: 3 }} />
                                  </div>
                                </div>
                                {r.delayProb !== null && (
                                  <div style={{ textAlign: "right" }}>
                                    <div style={{ fontSize: 9, color: "var(--text3)", marginBottom: 2 }}>Risco atraso</div>
                                    <span style={{ fontSize: 11, fontWeight: 800, color: probColor }}>{r.delayProb}%</span>
                                  </div>
                                )}
                                <span style={{ fontSize: 10, fontWeight: 700, color: r.diasTotais > 50 ? "#c0392b" : "#27ae60" }}>{r.diasTotais} d.u.</span>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>

                  </div>
                );
              })()}

            </div>
          )}
        </div>);
}
