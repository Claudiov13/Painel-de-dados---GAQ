import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function PlanejamentoScreen({ base, fBack, recMesAberto, setAba, setRecMesAberto, setSelProc }) {
  return (<div className="anim-fade" style={{ paddingBottom: 16 }}>
          {base.length === 0 ? (
            <div className="gaq-card" style={{ textAlign: "center", padding: 80, color: "#6b7280", background: "#f9f8f6", borderRadius: 28, border: "1px dashed #d1d5db" }}>
              <div style={{ fontSize: 48, marginBottom: 12, opacity: .4 }}>◈</div>
              <div style={{ fontWeight: 700, fontSize: 16, color: "#111", marginBottom: 6 }}>Planejamento de Compras</div>
              <div style={{ fontSize: 13, marginBottom: 20 }}>Carregue a base de dados para gerar forecasts e insights.</div>
              <button onClick={() => setAba("upload")} style={{ background: "#111", color: "#fff", border: "none", borderRadius: 999, padding: "10px 28px", cursor: "pointer", fontSize: 13, fontWeight: 600, letterSpacing: .5 }}>Carregar Base</button>
            </div>
          ) : (() => {
            /* ── Design system: clean, light and Apple-like ── */
            const PL = {
              bg:      "#f5f5f7",
              bg2:     "#ffffff",
              ink:     "#101828",
              ink2:    "#344054",
              ink3:    "#667085",
              line:    "#e5e7eb",
              accent:  "#0a84ff",
              accentL: "#e8f2ff",
              green:   "#138a58",
              greenL:  "#def7e8",
              amber:   "#b26a00",
              amberL:  "#fff3d6",
              red:     "#c2410c",
              redL:    "#ffede5",
              mono:    "'SF Mono','Fira Mono','Consolas',monospace",
            };
            const MESES = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];

            const hoje = new Date(); hoje.setHours(0,0,0,0);
            const anoAtual = hoje.getFullYear();
            const anoBase = anoAtual - 1;

            const concAno = base.filter(r => r.isConcluded && ((r.anoRC === anoBase) || (!r.anoRC && r.anoSD === anoBase)));
            const totalConc = concAno.length;

            const mesesCount = Array(12).fill(0);
            concAno.forEach(r => { if (r.mesAbertura !== null && r.mesAbertura !== undefined) mesesCount[r.mesAbertura]++; });
            const maxMes = Math.max(...mesesCount, 1);
            const avgMes = totalConc / 12;
            const peakMesIdx = mesesCount.indexOf(maxMes);

            const areaMap = {};
            concAno.forEach(r => {
              const a = r["Área Requisitante"] || "N/I";
              if (!areaMap[a]) areaMap[a] = { total: 0, leads: [], meses: Array(12).fill(0) };
              areaMap[a].total++;
              if (r.diasTotais > 0) areaMap[a].leads.push(r.diasTotais);
              if (r.mesAbertura !== null) areaMap[a].meses[r.mesAbertura]++;
            });
            const topAreas = Object.entries(areaMap)
              .sort((a, b) => b[1].total - a[1].total).slice(0, 10)
              .map(([area, d]) => ({
                area, total: d.total,
                avgLead: d.leads.length ? Math.round(d.leads.reduce((s,v) => s+v, 0) / d.leads.length) : 0,
                peakMes: d.meses.indexOf(Math.max(...d.meses)),
                projecao: Math.round(d.total * 1.05),
              }));

            const modMap = {};
            concAno.forEach(r => {
              const m = r.Modalidade || "N/I";
              if (!modMap[m]) modMap[m] = { total: 0, leads: [] };
              modMap[m].total++;
              if (r.diasTotais > 0) modMap[m].leads.push(r.diasTotais);
            });
            const topMods = Object.entries(modMap)
              .sort((a, b) => b[1].total - a[1].total).slice(0, 7)
              .map(([m, d]) => ({
                mod: m, total: d.total,
                avg: d.leads.length ? Math.round(d.leads.reduce((s,x) => s+x, 0) / d.leads.length) : 0,
                min: d.leads.length ? Math.min(...d.leads) : 0,
                max: d.leads.length ? Math.max(...d.leads) : 0,
              }));
            const avgLeadGeral = topMods.reduce((s,m) => s + m.avg * m.total, 0) / (topMods.reduce((s,m) => s + m.total, 0) || 1) | 0;
            const recorrentes = fBack.filter(r => r.emA && !r.aberturaRC);
            const maxBarMod = Math.max(...topMods.map(m => m.max), 1);

            return (
              <div style={{ background: PL.bg, minHeight: 600, borderRadius: 32, overflow: "hidden", boxShadow: "0 24px 60px rgba(15,23,42,0.08)", border: "1px solid rgba(15,23,42,0.06)", fontFamily: "-apple-system, BlinkMacSystemFont, 'SF Pro Display', 'Segoe UI', sans-serif" }}>

                {/* ── HERO: soft gradient + clean glass stats ── */}
                <div style={{ background: "linear-gradient(135deg, #09111f 0%, #153d6f 52%, #0a84ff 100%)", position: "relative", overflow: "hidden", padding: "42px 42px 34px" }}>
                  <div style={{ position: "absolute", inset: 0, backgroundImage: "radial-gradient(circle, rgba(255,255,255,0.08) 1px, transparent 1px)", backgroundSize: "30px 30px", pointerEvents: "none" }} />
                  <div style={{ position: "absolute", right: -60, top: "50%", transform: "translateY(-50%)", width: 520, height: 520, background: "radial-gradient(ellipse at center, rgba(10,132,255,0.28) 0%, transparent 68%)", pointerEvents: "none" }} />
                  <div style={{ position: "relative", zIndex: 1 }}>
                    <div style={{ fontSize: 10, color: "rgba(255,255,255,.52)", letterSpacing: 4, textTransform: "uppercase", marginBottom: 18, fontWeight: 600 }}>Sesc GAQ — Inteligência de Compras</div>
                    <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between", flexWrap: "wrap", gap: 20 }}>
                      <div>
                        <div style={{ fontSize: 52, fontWeight: 900, color: "#fff", lineHeight: 1, letterSpacing: -2, marginBottom: 10 }}>
                          PLANEJAMENTO<br/><span style={{ color: "rgba(255,255,255,.25)" }}>{anoAtual}</span>
                        </div>
                        <div style={{ fontSize: 13, color: "rgba(255,255,255,.72)", fontWeight: 400, maxWidth: 460, lineHeight: 1.55 }}>
                          Forecast baseado em {totalConc} processos concluídos em {anoBase}. Sazonalidade, lead time e abertura recomendada por área.
                        </div>
                      </div>
                      <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, minWidth: 300 }}>
                        {[
                          { label: "Processos " + anoBase, val: totalConc.toLocaleString("pt-BR") },
                          { label: "Lead médio (d.u.)", val: avgLeadGeral + " d.u." },
                          { label: "Mês pico", val: MESES[peakMesIdx] },
                          { label: "Pré-compra sem RC (carryover)", val: recorrentes.length },
                        ].map((s, i) => (
                          <div key={i} style={{ background: "rgba(255,255,255,.12)", backdropFilter: "blur(10px)", border: "1px solid rgba(255,255,255,.14)", borderRadius: 18, padding: "16px 18px" }}>
                            <div style={{ fontSize: 9, color: "rgba(255,255,255,.55)", textTransform: "uppercase", letterSpacing: 1.5, marginBottom: 4 }}>{s.label}</div>
                            <div style={{ fontSize: 22, fontWeight: 800, color: "#fff", letterSpacing: -.5 }}>{s.val}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>

                {totalConc === 0 ? (
                  <div style={{ padding: 40, textAlign: "center", color: PL.ink3 }}>
                    <div style={{ fontSize: 14, fontWeight: 700, color: PL.ink, marginBottom: 4 }}>Nenhum processo concluído em {anoBase}</div>
                    <div style={{ fontSize: 12 }}>Verifique as datas na base de dados.</div>
                  </div>
                ) : (<>

                  {/* ── SAZONALIDADE: banda creme com barras editoriais ── */}
                  <div style={{ padding: "36px 48px", borderBottom: `1px solid ${PL.line}` }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 24, flexWrap: "wrap", gap: 8 }}>
                      <div>
                        <div style={{ fontSize: 10, color: PL.ink3, textTransform: "uppercase", letterSpacing: 3, fontWeight: 600, marginBottom: 4 }}>Análise · {anoBase}</div>
                        <div style={{ fontSize: 24, fontWeight: 800, color: PL.ink, letterSpacing: -1 }}>Sazonalidade de Aberturas</div>
                      </div>
                      <div style={{ fontSize: 11, color: PL.ink3 }}>Média: <strong style={{ color: PL.ink }}>{Math.round(avgMes)}</strong> proc/mês &nbsp;·&nbsp; Pico: <strong style={{ color: PL.ink }}>{MESES[peakMesIdx]} ({maxMes})</strong></div>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(12, 1fr)", gap: 6 }}>
                      {mesesCount.map((c, i) => {
                        const pct = maxMes > 0 ? (c / maxMes) : 0;
                        const isPeak = c >= maxMes;
                        const isHigh = c > avgMes * 1.2 && !isPeak;
                        const barH = Math.max(6, Math.round(pct * 110));
                        const barBg = isPeak ? PL.ink : isHigh ? "#3d3d3d" : PL.line;
                        const numColor = isPeak ? PL.ink : isHigh ? "#3d3d3d" : PL.ink3;
                        return (
                          <div key={i} style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
                            <div style={{ fontSize: 12, fontWeight: isPeak ? 900 : 600, color: numColor, marginBottom: 6, fontVariantNumeric: "tabular-nums" }}>{c}</div>
                            <div style={{ width: "100%", height: barH, background: barBg, borderRadius: 999, transition: "height .3s" }} />
                            <div style={{ fontSize: 9, color: isPeak ? PL.ink : PL.ink3, marginTop: 6, fontWeight: isPeak ? 800 : 500, textTransform: "uppercase", letterSpacing: .5 }}>{MESES[i]}</div>
                            {isPeak && <div style={{ fontSize: 8, background: PL.ink, color: "#fff", borderRadius: 2, padding: "1px 5px", marginTop: 3, letterSpacing: .5, fontWeight: 700 }}>PICO</div>}
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {/* ── LEAD TIME POR MODALIDADE: horizontal bars editoriais ── */}
                  <div style={{ padding: "36px 48px", background: PL.bg2, borderBottom: `1px solid ${PL.line}` }}>
                    <div style={{ marginBottom: 24 }}>
                      <div style={{ fontSize: 10, color: PL.ink3, textTransform: "uppercase", letterSpacing: 3, fontWeight: 600, marginBottom: 4 }}>Benchmark · Dias Úteis</div>
                      <div style={{ fontSize: 24, fontWeight: 800, color: PL.ink, letterSpacing: -1 }}>Lead Time por Modalidade</div>
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
                      {topMods.map((m, i) => {
                        const avgPct = maxBarMod > 0 ? Math.round((m.avg / maxBarMod) * 100) : 0;
                        const maxPct = maxBarMod > 0 ? Math.round((m.max / maxBarMod) * 100) : 0;
                        const risk = m.avg > 60 ? { bg: PL.redL, text: PL.red, label: "Alto" } : m.avg > 30 ? { bg: PL.amberL, text: PL.amber, label: "Médio" } : { bg: PL.greenL, text: PL.green, label: "Baixo" };
                        return (
                          <div className="gaq-responsive-grid" key={i} style={{ display: "grid", gridTemplateColumns: "200px 1fr 80px 72px", alignItems: "center", gap: 16, padding: "14px 0", borderBottom: `1px solid ${PL.line}` }}>
                            <div style={{ fontSize: 12, fontWeight: 700, color: PL.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={m.mod}>{m.mod}</div>
                            <div style={{ position: "relative", height: 10, background: PL.line, borderRadius: 999 }}>
                              <div style={{ position: "absolute", left: 0, width: maxPct + "%", height: "100%", background: "#cbd5e1", borderRadius: 999 }} />
                              <div style={{ position: "absolute", left: 0, width: avgPct + "%", height: "100%", background: PL.accent, borderRadius: 999 }} />
                            </div>
                            <div style={{ textAlign: "right", fontFamily: PL.mono, fontSize: 13, fontWeight: 800, color: PL.ink }}>
                              {m.avg}<span style={{ fontSize: 9, fontWeight: 400, color: PL.ink3, marginLeft: 2 }}>d.u.</span>
                            </div>
                            <div style={{ textAlign: "center" }}>
                              <span style={{ background: risk.bg, color: risk.text, fontSize: 9, fontWeight: 700, borderRadius: 999, padding: "4px 9px", textTransform: "uppercase", letterSpacing: .5 }}>{risk.label}</span>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                    <div style={{ marginTop: 16, fontSize: 10, color: PL.ink3, display: "flex", gap: 20 }}>
                      <span><span style={{ display: "inline-block", width: 12, height: 4, background: PL.ink, borderRadius: 1, verticalAlign: "middle", marginRight: 5 }} />Média</span>
                      <span><span style={{ display: "inline-block", width: 12, height: 4, background: "#d5d0c8", borderRadius: 1, verticalAlign: "middle", marginRight: 5 }} />Máximo histórico</span>
                    </div>
                  </div>

                  {/* ── FORECAST POR ÁREA: lista editorial grande ── */}
                  <div style={{ padding: "36px 48px", borderBottom: `1px solid ${PL.line}` }}>
                    <div style={{ marginBottom: 28 }}>
                      <div style={{ fontSize: 10, color: PL.ink3, textTransform: "uppercase", letterSpacing: 3, fontWeight: 600, marginBottom: 4 }}>Forecast · {anoAtual}</div>
                      <div style={{ fontSize: 24, fontWeight: 800, color: PL.ink, letterSpacing: -1 }}>Projeção por Área Requisitante</div>
                      <div style={{ fontSize: 12, color: PL.ink3, marginTop: 4 }}>Recomendação de abertura calculada: Mês pico − lead time médio − 20 d.u. buffer</div>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(340px, 1fr))", gap: 12 }}>
                      {topAreas.map((r, i) => {
                        const recMes = (r.peakMes - Math.ceil((r.avgLead + 20) / 21) + 12) % 12;
                        const riskColor = r.avgLead > 60 ? PL.red : r.avgLead > 30 ? PL.amber : PL.green;
                        const riskBg   = r.avgLead > 60 ? PL.redL : r.avgLead > 30 ? PL.amberL : PL.greenL;
                        return (
                          <div key={i} style={{ background: "#fff", padding: "20px 22px", borderRadius: 24, border: `1px solid ${PL.line}`, boxShadow: "0 8px 24px rgba(15,23,42,0.04)" }}>
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
                              <div style={{ fontSize: 13, fontWeight: 700, color: PL.ink, flex: 1, paddingRight: 8, lineHeight: 1.3 }}>{r.area}</div>
                              <div style={{ background: PL.accent, color: "#fff", fontSize: 9, fontWeight: 700, borderRadius: 999, padding: "5px 10px", letterSpacing: 1, textTransform: "uppercase", flexShrink: 0 }}>
                                Abrir em {MESES[recMes]}
                              </div>
                            </div>
                            <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
                              {[
                                { l: anoBase, v: r.total, unit: "proc" },
                                { l: "Forecast", v: r.projecao, unit: "proj.", accent: true },
                                { l: "Lead médio", v: r.avgLead || "—", unit: "d.u.", color: riskColor, bg: riskBg },
                              ].map((s, j) => (
                                <div key={j} style={{ background: s.bg || PL.bg2, borderRadius: 16, padding: "10px 12px" }}>
                                  <div style={{ fontSize: 8, color: s.color || PL.ink3, textTransform: "uppercase", letterSpacing: 1, marginBottom: 3, fontWeight: 600 }}>{s.l}</div>
                                  <div style={{ fontSize: 20, fontWeight: 900, color: s.color || (s.accent ? PL.ink : PL.ink), letterSpacing: -1, fontVariantNumeric: "tabular-nums" }}>
                                    {s.v}<span style={{ fontSize: 9, fontWeight: 400, marginLeft: 2, opacity: .6 }}>{s.unit}</span>
                                  </div>
                                </div>
                              ))}
                            </div>
                            <div style={{ marginTop: 10, fontSize: 10, color: PL.ink3 }}>
                              Pico histórico: <strong style={{ color: PL.ink }}>{MESES[r.peakMes]}</strong>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {/* ── CARRYOVER: pré-compra sem RC ── */}
                  {recorrentes.length > 0 && (
                    <div style={{ padding: "36px 48px", background: PL.bg2, borderBottom: `1px solid ${PL.line}` }}>
                      <div style={{ marginBottom: 20 }}>
                        <div style={{ fontSize: 10, color: PL.ink3, textTransform: "uppercase", letterSpacing: 3, fontWeight: 600, marginBottom: 4 }}>Carryover · {anoAtual}</div>
                        <div style={{ fontSize: 24, fontWeight: 800, color: PL.ink, letterSpacing: -1 }}>Pré-compra em aberto — Demandas pendentes</div>
                        <div style={{ fontSize: 12, color: PL.ink3, marginTop: 4 }}>{recorrentes.length} demandas com Service Desk aberto sem RC criada. Potencial recorrência ou pendência.</div>
                      </div>
                      <div style={{ display: "flex", flexDirection: "column", gap: 0, borderTop: `1px solid ${PL.line}` }}>
                        {recorrentes.slice(0, 12).map((r, i) => {
                          const diasColor = r.diasSD > 50 ? PL.red : r.diasSD > 20 ? PL.amber : PL.green;
                          const diasBg    = r.diasSD > 50 ? PL.redL : r.diasSD > 20 ? PL.amberL : PL.greenL;
                          return (
                            <div className="gaq-responsive-grid" key={i} onClick={() => setSelProc(r)} style={{ display: "grid", gridTemplateColumns: "90px 160px 1fr 80px", gap: 16, alignItems: "center", padding: "14px 0", borderBottom: `1px solid ${PL.line}`, cursor: "pointer", transition: "background .15s" }}
                              onMouseEnter={e => e.currentTarget.style.background = PL.bg}
                              onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                              <div style={{ background: diasBg, borderRadius: 16, padding: "8px 10px", textAlign: "center" }}>
                                <div style={{ fontSize: 18, fontWeight: 900, color: diasColor, letterSpacing: -1, lineHeight: 1 }}>{r.diasSD || 0}</div>
                                <div style={{ fontSize: 8, color: diasColor, textTransform: "uppercase", letterSpacing: 1 }}>d.u. pré-compra</div>
                              </div>
                              <div>
                                <div style={{ fontSize: 11, fontWeight: 700, color: PL.ink }}>{r.TicketSD || "—"}</div>
                                <div style={{ fontSize: 10, color: PL.ink3, marginTop: 2 }}>{r.respNCL || "—"}</div>
                              </div>
                              <div>
                                <div style={{ fontSize: 12, color: PL.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{(r.Objeto || "Sem descrição").slice(0, 80)}</div>
                                <div style={{ fontSize: 10, color: PL.ink3, marginTop: 2 }}>{r["Área Requisitante"] || "—"} · {r.Modalidade || "—"}</div>
                              </div>
                              <div style={{ textAlign: "right", fontSize: 10, color: PL.ink3 }}>
                                {r.aberturaSD ? r.aberturaSD.toLocaleDateString("pt-BR") : "—"}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                      {recorrentes.length > 12 && <div style={{ fontSize: 11, color: PL.ink3, marginTop: 12 }}>+ {recorrentes.length - 12} demandas adicionais na aba Processos</div>}
                    </div>
                  )}

                  {/* ── CRONOGRAMA ANUAL INTELIGENTE (NLP + clustering) ── */}
                  {(() => {
                    const MESES_FULL = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
                    const recur = detectRecurrentes(base, { anos: [anoAtual - 2, anoBase, anoAtual] });
                    if (recur.recurrentes.length === 0) {
                      return (
                        <div style={{ padding: "36px 48px", borderBottom: `1px solid ${PL.line}` }}>
                          <div style={{ marginBottom: 12 }}>
                            <div style={{ fontSize: 10, color: PL.ink3, textTransform: "uppercase", letterSpacing: 3, fontWeight: 600, marginBottom: 4 }}>Pattern Detection · NLP</div>
                            <div style={{ fontSize: 24, fontWeight: 800, color: PL.ink, letterSpacing: -1 }}>Cronograma Anual Inteligente</div>
                          </div>
                          <div style={{ background: "#fff", padding: 24, borderRadius: 18, border: `1px dashed ${PL.line}`, color: PL.ink3, fontSize: 12 }}>
                            Ainda não há sinal suficiente para detectar demandas recorrentes — é preciso pelo menos 2 anos de dados com objetos descritos.
                          </div>
                        </div>
                      );
                    }
                    const maxEv = Math.max(...recur.calendario.map(c => c.eventos.length), 1);
                    const totalEventosAno = recur.recurrentes.length;
                    const eventosNoCalendario = recur.calendario.reduce((s, c) => s + c.eventos.length, 0);
                    const anosLabel = recur.meta.anosCobertos.join(", ");
                    return (
                      <div style={{ padding: "36px 48px", background: PL.bg2, borderBottom: `1px solid ${PL.line}` }}>
                        {/* Header */}
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 16, marginBottom: 24 }}>
                          <div>
                            <div style={{ fontSize: 10, color: PL.ink3, textTransform: "uppercase", letterSpacing: 3, fontWeight: 600, marginBottom: 4 }}>Pattern Detection · NLP + Clustering</div>
                            <div style={{ fontSize: 24, fontWeight: 800, color: PL.ink, letterSpacing: -1 }}>Cronograma Anual Inteligente</div>
                            <div style={{ fontSize: 12, color: PL.ink3, marginTop: 4, maxWidth: 640 }}>
                              <strong>{totalEventosAno}</strong> demanda(s) recorrente(s) detectada(s) em <strong>{recur.meta.totalProcs}</strong> processo(s) de {anosLabel}. Clique em um mês para ver as demandas previstas e a sugestão de quando abrir a pré-compra.
                            </div>
                          </div>
                          <div style={{ display: "flex", gap: 8, fontSize: 10, color: PL.ink3 }}>
                            <span style={{ background: PL.bg, padding: "8px 12px", borderRadius: 12, border: `1px solid ${PL.line}` }}>{eventosNoCalendario} demandas com mês previsto</span>
                            <span style={{ background: PL.bg, padding: "8px 12px", borderRadius: 12, border: `1px solid ${PL.line}` }}>{totalEventosAno - eventosNoCalendario} sem data clara</span>
                          </div>
                        </div>

                        {/* Linha do tempo anual */}
                        <div style={{ display: "grid", gridTemplateColumns: "repeat(12, 1fr)", gap: 6, marginBottom: 18 }}>
                          {recur.calendario.map(({ mes, eventos }) => {
                            const ativo = recMesAberto === mes;
                            const intensity = eventos.length / maxEv;
                            const isEmpty = eventos.length === 0;
                            const bg = isEmpty
                              ? PL.bg
                              : ativo
                                ? PL.ink
                                : `rgba(10, 132, 255, ${(0.08 + intensity * 0.30).toFixed(3)})`;
                            const color = ativo ? "#fff" : (isEmpty ? PL.ink3 : PL.ink);
                            return (
                              <div key={mes}
                                onClick={() => !isEmpty && setRecMesAberto(ativo ? null : mes)}
                                style={{ background: bg, color, borderRadius: 12, padding: "12px 8px", textAlign: "center", cursor: isEmpty ? "default" : "pointer", border: `1px solid ${ativo ? PL.ink : PL.line}`, transition: "all .15s" }}>
                                <div style={{ fontSize: 9, fontWeight: 600, textTransform: "uppercase", letterSpacing: 1, opacity: .8 }}>{MESES[mes]}</div>
                                <div style={{ fontSize: 22, fontWeight: 900, marginTop: 4, letterSpacing: -1, fontVariantNumeric: "tabular-nums" }}>{eventos.length}</div>
                                <div style={{ fontSize: 8, marginTop: 2, opacity: .7 }}>{eventos.length === 1 ? "demanda" : "demandas"}</div>
                              </div>
                            );
                          })}
                        </div>

                        {/* Painel do mês selecionado */}
                        {recMesAberto !== null && recur.calendario[recMesAberto].eventos.length > 0 && (
                          <div style={{ background: "#fff", borderRadius: 24, padding: "20px 24px", border: `1px solid ${PL.line}`, marginBottom: 8 }}>
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
                              <div style={{ fontSize: 12, color: PL.ink, textTransform: "uppercase", letterSpacing: 2, fontWeight: 700 }}>
                                {MESES_FULL[recMesAberto]} · {recur.calendario[recMesAberto].eventos.length} demanda(s) prevista(s)
                              </div>
                              <button onClick={() => setRecMesAberto(null)} style={{ background: "transparent", border: "none", color: PL.ink3, fontSize: 11, cursor: "pointer" }}>fechar ×</button>
                            </div>
                            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(320px, 1fr))", gap: 12 }}>
                              {recur.calendario[recMesAberto].eventos.map((ev, i) => {
                                const confColor = ev.confidence >= 75 ? PL.green : ev.confidence >= 50 ? PL.amber : PL.red;
                                const confBg = ev.confidence >= 75 ? PL.greenL : ev.confidence >= 50 ? PL.amberL : PL.redL;
                                const anosOrd = Object.entries(ev.anos).sort((a, b) => +a[0] - +b[0]);
                                return (
                                  <div key={i} style={{ background: PL.bg, padding: "16px 18px", borderRadius: 18, border: `1px solid ${PL.line}`, display: "flex", flexDirection: "column", gap: 9 }}>
                                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
                                      <div style={{ fontSize: 13, fontWeight: 800, color: PL.ink, lineHeight: 1.25 }}>{ev.label}</div>
                                      <div style={{ background: confBg, color: confColor, fontSize: 9, fontWeight: 800, padding: "3px 8px", borderRadius: 999, letterSpacing: .5, textTransform: "uppercase", flexShrink: 0 }}>
                                        {ev.confidence}%
                                      </div>
                                    </div>
                                    <div style={{ fontSize: 10, color: PL.ink3, display: "flex", gap: 10, flexWrap: "wrap" }}>
                                      {anosOrd.map(([a, c]) => (
                                        <span key={a}><strong style={{ color: PL.ink2 }}>{a}:</strong> {c} proc</span>
                                      ))}
                                      <span>· lead {ev.avgLead || "—"} d.u.</span>
                                    </div>
                                    {ev.topAreas.length > 0 && (
                                      <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
                                        {ev.topAreas.map(([a, c]) => (
                                          <span key={a} style={{ background: "#fff", border: `1px solid ${PL.line}`, fontSize: 9, padding: "2px 7px", borderRadius: 999, color: PL.ink2 }}>{a} <strong style={{ color: PL.ink }}>{c}</strong></span>
                                        ))}
                                      </div>
                                    )}
                                    {ev.mesSDSugerido !== null && (
                                      <div style={{ background: PL.accentL, padding: "8px 10px", borderRadius: 10, fontSize: 10, color: PL.ink, lineHeight: 1.4 }}>
                                        <strong>Sugestão:</strong> abrir pré-compra em <strong style={{ color: PL.accent }}>{MESES_FULL[ev.mesSDSugerido]}</strong>
                                        {ev.aindaEsteAno != null && (
                                          <div style={{ marginTop: 3, color: PL.amber, fontSize: 9, fontWeight: 700 }}>
                                            ⚠ {ev.aindaEsteAno} ainda sem demanda registrada
                                          </div>
                                        )}
                                      </div>
                                    )}
                                    <div style={{ fontSize: 9, color: PL.ink3, fontStyle: "italic", overflow: "hidden", textOverflow: "ellipsis", display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical" }} title={ev.objetoExemplo}>
                                      "{(ev.objetoExemplo || "").slice(0, 140)}{(ev.objetoExemplo || "").length > 140 ? "…" : ""}"
                                    </div>
                                    <div style={{ display: "flex", gap: 4, flexWrap: "wrap", borderTop: `1px solid ${PL.line}`, paddingTop: 8 }}>
                                      {ev.keywords.slice(0, 3).map((k, j) => (
                                        <span key={j} style={{ fontFamily: PL.mono, fontSize: 8, color: PL.ink3, background: "#fff", padding: "1px 6px", borderRadius: 4, border: `1px solid ${PL.line}` }}>{k}</span>
                                      ))}
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        )}

                        {/* Top recorrências (sem agrupar por mês) */}
                        <div style={{ marginTop: 8 }}>
                          <div style={{ fontSize: 10, color: PL.ink3, textTransform: "uppercase", letterSpacing: 2, fontWeight: 700, marginBottom: 10 }}>
                            Top demandas por confiança
                          </div>
                          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(260px, 1fr))", gap: 8 }}>
                            {recur.recurrentes.slice(0, 8).map((ev, i) => {
                              const confColor = ev.confidence >= 75 ? PL.green : ev.confidence >= 50 ? PL.amber : PL.red;
                              return (
                                <div key={i}
                                  onClick={() => ev.mesPrevisto != null && setRecMesAberto(ev.mesPrevisto)}
                                  style={{ background: "#fff", padding: "10px 12px", borderRadius: 14, border: `1px solid ${PL.line}`, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, cursor: ev.mesPrevisto != null ? "pointer" : "default", transition: "background .15s" }}
                                  onMouseEnter={e => e.currentTarget.style.background = PL.bg}
                                  onMouseLeave={e => e.currentTarget.style.background = "#fff"}>
                                  <div style={{ flex: 1, minWidth: 0 }}>
                                    <div style={{ fontSize: 11, fontWeight: 700, color: PL.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{ev.label}</div>
                                    <div style={{ fontSize: 9, color: PL.ink3, marginTop: 2 }}>
                                      {ev.total} proc · {ev.yearsSet.join("/")} {ev.mesPrevisto != null ? `· pico ${MESES[ev.mesPrevisto]}` : ""}
                                    </div>
                                  </div>
                                  <div style={{ fontSize: 11, fontWeight: 800, color: confColor, fontFamily: PL.mono }}>{ev.confidence}%</div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      </div>
                    );
                  })()}

                  {/* ── RODAPÉ ── */}
                  <div style={{ padding: "20px 48px", display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 8, alignItems: "center" }}>
                    <div style={{ fontSize: 10, color: PL.ink3, letterSpacing: .5 }}>
                      Sesc GAQ — Planejamento de Compras · {new Date().toLocaleDateString("pt-BR")} · Base {anoBase} → Forecast {anoAtual}
                    </div>
                    <div style={{ display: "flex", gap: 16, fontSize: 10, color: PL.ink3 }}>
                      <span>{totalConc} processos analisados</span>
                      <span>{topMods.length} modalidades</span>
                      <span>{topAreas.length} áreas</span>
                    </div>
                  </div>
                </>)}
              </div>
            );
          })()}
        </div>);
}
