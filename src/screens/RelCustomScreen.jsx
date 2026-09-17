import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function RelCustomScreen({ areas, baseEnriched, meta, phaseIntervals, rcuAreas, rcuDias, rcuDiasModo, rcuExcl, rcuInds, rcuStatus, setRcuAreas, setRcuDias, setRcuDiasModo, setRcuExcl, setRcuInds, setRcuStatus, setSelProc }) {
  return ((() => {
          const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
          const diasN = parseInt(rcuDias, 10);
          const temDias = !isNaN(diasN) && diasN > 0;
          const stDefs = [
            { id: "andamento", label: "Em andamento", test: r => r.emA, color: "#2e86c1" },
            { id: "concluidos", label: "Concluídos", test: r => r.isConcluded, color: "#27ae60" },
            { id: "cancelfrac", label: "Cancelados/Fracassados", test: r => r.isCanceled || r.isFailed, color: "#c0392b" },
            { id: "suspensos", label: "Suspensos", test: r => r.isSuspended, color: "#e67e22" },
          ];
          // Id estável do processo (para exclusão manual)
          const procId = (r) => r.ProcessKey || r.NumRC || r.TicketSD || "";
          // Campos Suite Sesc (mesma busca tolerante do modal de timeline)
          const suiteInfo = (r) => {
            const fk = (id) => Object.keys(r).find(k => nrm(k).replace(/[\s_\-]/g, "") === id);
            const kId = fk("identificadorsuitesesc"), kSt = fk("statussuitesesc"), kHab = fk("habilitadaemsuitesesc"), kResp = fk("responsabilidadesuitesesc");
            const txt = (v) => (v == null ? "" : v.toString().trim());
            const vId = txt(kId && r[kId]), vSt = txt(kSt && r[kSt]), vResp = txt(kResp && r[kResp]);
            const dHab = kHab ? pd(r[kHab]) : null;
            const vHab = dHab ? fmt(dHab) : txt(kHab && r[kHab]);
            return (vId || vSt || vHab || vResp) ? { id: vId, st: vSt, hab: vHab, resp: vResp } : null;
          };
          const procs = baseEnriched
            .filter(r => {
              const pid = procId(r);
              if (pid && rcuExcl.includes(pid)) return false;
              if (rcuAreas.length && !rcuAreas.includes(r["Área Requisitante"])) return false;
              if (rcuStatus.length && !rcuStatus.some(id => { const d = stDefs.find(s => s.id === id); return d && d.test(r); })) return false;
              if (temDias) {
                if (!r.aberturaRC) return false;
                const d = du(r.aberturaRC, hoje);
                if (rcuDiasModo === "ate" ? d > diasN : d <= diasN) return false;
              }
              return true;
            })
            .map(r => ({ ...r, _diasRC: r.aberturaRC ? du(r.aberturaRC, hoje) : null }))
            .sort((a, b) => (b._diasRC ?? -1) - (a._diasRC ?? -1));

          // ── Indicadores do resumo ──
          const emAList = procs.filter(r => r.emA);
          const conclList = procs.filter(r => r.isConcluded);
          // SLA só é computado se algum indicador que depende dele estiver selecionado
          const needSla = rcuInds.includes("slaok") || rcuInds.includes("criticos");
          const slaList = needSla ? procs.map(r => calcSLAClassification(r, phaseIntervals)).filter(Boolean) : [];
          const slaVenc = slaList.filter(s => s.bucket === "critico" || s.bucket === "sla_vencido").length;
          const slaOkPct = slaList.length ? Math.round(((slaList.length - slaVenc) / slaList.length) * 100) : null;
          const criticosN = slaList.filter(s => s.bucket === "critico").length;
          const leads = conclList.map(r => r.LeadTime).filter(v => v > 0);
          const leadMedio = leads.length ? Math.round(leads.reduce((a, b) => a + b, 0) / leads.length) : null;
          const agings = emAList.filter(r => r.aberturaRC).map(r => du(r.aberturaRC, hoje));
          const agingMedio = agings.length ? Math.round(agings.reduce((a, b) => a + b, 0) / agings.length) : null;
          const IND_DEFS = [
            { id: "total", label: "Processos", val: procs.length, color: "#1a5276" },
            { id: "andamento", label: "Em andamento", val: emAList.length, color: "#2e86c1" },
            { id: "concluidos", label: "Concluídos", val: conclList.length, color: "#27ae60" },
            { id: "slaok", label: "SLA em dia", val: slaOkPct == null ? "—" : slaOkPct + "%", color: "#16a085" },
            { id: "criticos", label: "Críticos", val: criticosN, color: "#c0392b" },
            { id: "lead", label: "Lead médio (d.u.)", val: leadMedio ?? "—", color: "#8e44ad" },
            { id: "aging", label: "Aging médio RC (d.u.)", val: agingMedio ?? "—", color: "#e67e22" },
            { id: "parados", label: "Parados >15 d.u.", val: emAList.filter(r => r.diasParado > 15).length, color: "#d35400" },
          ];
          const kpisSel = IND_DEFS.filter(d => rcuInds.includes(d.id));

          const toggle = (arr, set, v) => set(arr.includes(v) ? arr.filter(x => x !== v) : [...arr, v]);
          const Chip = ({ on, color = "#1a5276", onClick, children }) => (
            <button type="button" onClick={onClick}
              style={{ background: on ? color : "var(--gaq-surface)", color: on ? "#fff" : "var(--gaq-text-2)", border: `1px solid ${on ? color : "var(--gaq-line)"}`, borderRadius: 999, padding: "4px 12px", fontSize: 11, fontWeight: 600, cursor: "pointer", transition: "all .15s" }}>
              {children}
            </button>
          );

          const diasLabel = temDias
            ? (rcuDiasModo === "ate" ? `RC aberta nos últimos ${diasN} d.u.` : `RC aberta há mais de ${diasN} d.u.`)
            : "sem filtro de dias";

          const exportarPDF = () => {
            const payload = {
              areasLabel: rcuAreas.length ? rcuAreas.join(" · ") : "Todas as áreas",
              diasLabel,
              statusLabel: rcuStatus.length ? rcuStatus.map(id => (stDefs.find(s => s.id === id) || {}).label).filter(Boolean).join(" + ") : "Todas as situações",
              kpis: kpisSel.map(d => ({ label: d.label, value: d.val, color: d.color })),
              procs: procs.slice(0, 400).map(r => ({
                rc: r.NumRC || "", ticket: r.TicketSD || "", area: r["Área Requisitante"] || "",
                mod: r.Modalidade || "", status: r.statusDet || r.status || "",
                resp: r.respAtivo || r.respNCL || r.Pregoeiro || "",
                objeto: (r.Objeto || "").slice(0, 220),
                diasRC: r._diasRC,
                processo: r.NumProcesso || "", pedidoSuite: r.NumPedidoSuite || "",
                suite: suiteInfo(r),
                entries: getOrderedTimelineEntries(r).filter(e => e.date).map((e, i, arr) => ({
                  label: e.label, data: fmt(e.date), du: i > 0 ? duSigned(arr[i - 1].date, e.date) : null, isCPL: e.isCPL,
                })),
              })),
              meta,
            };
            const html = gerarRelatorioCustom(payload);
            const win = window.open("", "_blank");
            win.document.write(html);
            win.document.close();
          };

          const MAX_LIST = 300;
          const lista = procs.slice(0, MAX_LIST);

          return <div className="anim-fade">
            {/* Header gradient */}
            <div style={{ background: "linear-gradient(135deg, #0f2e2a 0%, #16a085 100%)", color: "#fff", borderRadius: "var(--gaq-r-xl, 16px)", padding: "24px 28px", marginBottom: 18, boxShadow: "var(--gaq-shadow-2)" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
                <div>
                  <div style={{ fontSize: 11, fontWeight: 600, opacity: .72, letterSpacing: ".05em", textTransform: "uppercase", marginBottom: 4 }}>Administração · Admin Master</div>
                  <div style={{ fontSize: 22, fontWeight: 700, letterSpacing: "-.02em" }}>Relatório Customizado</div>
                  <div style={{ fontSize: 12, opacity: .82, marginTop: 2 }}>Selecione áreas, janela de dias a contar da RC e indicadores — e gere o relatório com a linha de datas de cada processo.</div>
                </div>
                <button type="button" onClick={exportarPDF} disabled={procs.length === 0}
                  style={{ background: "rgba(255,255,255,.16)", border: "1px solid rgba(255,255,255,.35)", color: "#fff", borderRadius: 10, padding: "9px 18px", fontSize: 12, fontWeight: 700, cursor: procs.length ? "pointer" : "not-allowed", opacity: procs.length ? 1 : .5, display: "flex", alignItems: "center", gap: 7 }}>
                  <Icon name="doc" size={14}/> Gerar Relatório (PDF)
                </button>
              </div>
            </div>

            {/* ── Filtros ── */}
            <div className="gaq-card" style={{ padding: 18, marginBottom: 14 }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 8, marginBottom: 10 }}>
                <div className="gaq-eyebrow">Áreas requisitantes {rcuAreas.length > 0 ? `(${rcuAreas.length} selecionada(s))` : "(todas)"}</div>
                <div style={{ display: "flex", gap: 6 }}>
                  <button type="button" className="gaq-btn" style={{ fontSize: 10, padding: "3px 10px" }} onClick={() => setRcuAreas([...areas])}>Todas</button>
                  <button type="button" className="gaq-btn" style={{ fontSize: 10, padding: "3px 10px" }} onClick={() => setRcuAreas([])}>Limpar</button>
                </div>
              </div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 16 }}>
                {areas.map(a => (
                  <Chip key={a} on={rcuAreas.includes(a)} color="#16a085" onClick={() => toggle(rcuAreas, setRcuAreas, a)}>{a}</Chip>
                ))}
              </div>

              <div style={{ display: "flex", gap: 24, flexWrap: "wrap", alignItems: "flex-start" }}>
                <div>
                  <div className="gaq-eyebrow" style={{ marginBottom: 8 }}>Dias a contar da RC (d.u.)</div>
                  <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                    <select value={rcuDiasModo} onChange={e => setRcuDiasModo(e.target.value)}
                      style={{ background: "var(--gaq-surface)", color: "var(--gaq-text)", border: "1px solid var(--gaq-line)", borderRadius: 8, padding: "6px 10px", fontSize: 12 }}>
                      <option value="ate">RC aberta nos últimos…</option>
                      <option value="acima">RC aberta há mais de…</option>
                    </select>
                    <input type="number" min="1" placeholder="ex.: 90" value={rcuDias} onChange={e => setRcuDias(e.target.value)}
                      style={{ width: 90, background: "var(--gaq-surface)", color: "var(--gaq-text)", border: "1px solid var(--gaq-line)", borderRadius: 8, padding: "6px 10px", fontSize: 12 }} />
                    <span style={{ fontSize: 11, color: "var(--gaq-text-3)" }}>d.u.</span>
                    {temDias && <button type="button" className="gaq-btn" style={{ fontSize: 10, padding: "3px 10px" }} onClick={() => setRcuDias("")}>×</button>}
                  </div>
                  <div style={{ fontSize: 10, color: "var(--gaq-text-3)", marginTop: 5 }}>{temDias ? diasLabel + " · processos sem RC ficam fora" : "Vazio = sem filtro de dias (inclui processos sem RC)"}</div>
                </div>
                <div>
                  <div className="gaq-eyebrow" style={{ marginBottom: 8 }}>Situação</div>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                    {stDefs.map(s => (
                      <Chip key={s.id} on={rcuStatus.includes(s.id)} color={s.color} onClick={() => toggle(rcuStatus, setRcuStatus, s.id)}>{s.label}</Chip>
                    ))}
                  </div>
                  <div style={{ fontSize: 10, color: "var(--gaq-text-3)", marginTop: 5 }}>Nenhuma selecionada = todas as situações</div>
                </div>
                <div>
                  <div className="gaq-eyebrow" style={{ marginBottom: 8 }}>Indicadores do resumo</div>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap", maxWidth: 420 }}>
                    {IND_DEFS.map(d => (
                      <Chip key={d.id} on={rcuInds.includes(d.id)} color={d.color} onClick={() => toggle(rcuInds, setRcuInds, d.id)}>{d.label}</Chip>
                    ))}
                  </div>
                </div>
              </div>
            </div>

            {/* ── Resumo dos indicadores ── */}
            {kpisSel.length > 0 && (
              <div className="gaq-kpi-grid" style={{ gridTemplateColumns: `repeat(auto-fit, minmax(130px, 1fr))`, marginBottom: 14 }}>
                {kpisSel.map(d => (
                  <div key={d.id} className="gaq-kpi" style={{ borderTop: `3px solid ${d.color}` }}>
                    <span className="lbl">{d.label}</span>
                    <div className="val" style={{ color: d.color }}>{d.val}</div>
                  </div>
                ))}
              </div>
            )}

            {/* ── Lista de processos com linha de datas ── */}
            <div className="gaq-card" style={{ padding: 18 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8, marginBottom: 12 }}>
                <div className="gaq-eyebrow">Processos ({procs.length}){procs.length > MAX_LIST ? ` · exibindo os ${MAX_LIST} mais antigos por RC` : ""}</div>
                <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                  {rcuExcl.length > 0 && (
                    <button type="button" onClick={() => setRcuExcl([])} title="Devolve ao relatório todos os processos excluídos manualmente"
                      style={{ background: "#c0392b18", border: "1px solid #c0392b55", color: "#c0392b", borderRadius: 999, padding: "3px 12px", fontSize: 10, fontWeight: 700, cursor: "pointer" }}>
                      {rcuExcl.length} excluído(s) manualmente · restaurar todos
                    </button>
                  )}
                  <div style={{ fontSize: 10, color: "var(--gaq-text-3)" }}>{diasLabel} · clique num processo para abrir a timeline · × exclui do relatório</div>
                </div>
              </div>
              {procs.length === 0 && (
                <div style={{ padding: "26px 10px", textAlign: "center", color: "var(--gaq-text-3)", fontSize: 12 }}>Nenhum processo com os filtros atuais.</div>
              )}
              {lista.map((r, ri) => {
                const entries = getOrderedTimelineEntries(r).filter(e => e.date);
                const stChip = r.emA ? { c: "#2e86c1", t: "Em andamento" } : r.isConcluded ? { c: "#27ae60", t: "Concluído" } : (r.isCanceled || r.isFailed) ? { c: "#c0392b", t: r.status } : { c: "#e67e22", t: r.status || "—" };
                const suite = suiteInfo(r);
                return (
                  <div key={(r.ProcessKey || "") + ri} onClick={() => setSelProc(r)}
                    style={{ border: "1px solid var(--gaq-line)", borderRadius: 12, padding: "12px 14px", marginBottom: 10, cursor: "pointer", background: "var(--gaq-surface)" }}
                    onMouseEnter={e => e.currentTarget.style.background = "var(--gaq-bg-2)"}
                    onMouseLeave={e => e.currentTarget.style.background = "var(--gaq-surface)"}>
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", marginBottom: 6 }}>
                      <span style={{ fontWeight: 700, fontSize: 13, color: "var(--gaq-blue)" }}>{r.NumRC || r.TicketSD || r.ProcessKey || "—"}</span>
                      <span style={{ fontSize: 10, fontWeight: 700, background: "#16a08522", color: "#16a085", borderRadius: 999, padding: "2px 9px" }}>{r["Área Requisitante"] || "—"}</span>
                      <span style={{ fontSize: 10, color: "var(--gaq-text-3)" }}>{r.Modalidade || "—"}</span>
                      <span style={{ fontSize: 10, fontWeight: 700, background: stChip.c + "22", color: stChip.c, borderRadius: 999, padding: "2px 9px" }}>{stChip.t}</span>
                      {r.statusDet && nrm(r.statusDet) !== nrm(stChip.t) && (
                        <span style={{ fontSize: 10, fontWeight: 800, background: stChip.c, color: "#fff", borderRadius: 999, padding: "2px 11px", letterSpacing: ".02em", textTransform: "uppercase" }}>{r.statusDet}</span>
                      )}
                      <span style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center", flexShrink: 0 }}>
                        {r._diasRC != null && <span style={{ fontSize: 10, fontWeight: 700, color: r._diasRC > 90 ? "#c0392b" : r._diasRC > 50 ? "#e67e22" : "var(--gaq-text-3)" }}>{r._diasRC} d.u. desde a RC</span>}
                        <button type="button" title="Excluir este processo do relatório"
                          onClick={(e) => { e.stopPropagation(); setRcuExcl(prev => [...prev, procId(r)]); }}
                          style={{ background: "none", border: "1px solid var(--gaq-line)", color: "var(--gaq-text-3)", borderRadius: 6, width: 22, height: 22, fontSize: 12, lineHeight: 1, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}
                          onMouseEnter={e => { e.currentTarget.style.background = "#c0392b"; e.currentTarget.style.color = "#fff"; e.currentTarget.style.borderColor = "#c0392b"; }}
                          onMouseLeave={e => { e.currentTarget.style.background = "none"; e.currentTarget.style.color = "var(--gaq-text-3)"; e.currentTarget.style.borderColor = "var(--gaq-line)"; }}>×</button>
                      </span>
                    </div>
                    {/* Códigos do processo (copiáveis) */}
                    <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 6 }}>
                      {r.NumRC && <CopyChip label="RC" value={r.NumRC} color="#2e86c1" solid />}
                      {r.TicketSD && <CopyChip label="Pré-compra" value={r.TicketSD} color="#27ae60" solid />}
                      {r.NumProcesso && <CopyChip label="Nº Processo" value={r.NumProcesso} color="#8e44ad" solid />}
                      {r.NumPedidoSuite && <CopyChip label="Pedido Suite" value={r.NumPedidoSuite} color="#d35400" solid />}
                    </div>
                    {/* Suite Sesc (quando houver informação) */}
                    {suite && (
                      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", background: "#16a08511", border: "1px solid #16a08533", borderRadius: 8, padding: "4px 10px", marginBottom: 8, fontSize: 10.5, color: "var(--gaq-text-2)" }}>
                        <b style={{ color: "#16a085", fontSize: 9, textTransform: "uppercase", letterSpacing: ".04em" }}>Suite Sesc</b>
                        {suite.id && <CopyChip label="Identificador" value={suite.id} color="#16a085" solid />}
                        {suite.st && <span>Status: <b style={{ color: "var(--gaq-text)" }}>{suite.st}</b></span>}
                        {suite.hab && <span>Habilitada em: <b style={{ color: "var(--gaq-text)" }}>{suite.hab}</b></span>}
                        {suite.resp && <span>Resp.: <b style={{ color: "var(--gaq-text)" }}>{suite.resp}</b></span>}
                      </div>
                    )}
                    <div style={{ fontSize: 11, color: "var(--gaq-text-2)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", marginBottom: 8 }}>
                      {(r.respAtivo || r.respNCL || r.Pregoeiro) ? <b>{r.respAtivo || r.respNCL || r.Pregoeiro} · </b> : null}{(r.Objeto || "").slice(0, 160) || "—"}
                    </div>
                    {/* Linha de datas do processo */}
                    {entries.length === 0
                      ? <div style={{ fontSize: 10, color: "var(--gaq-text-3)", fontStyle: "italic" }}>Sem datas registradas.</div>
                      : <div style={{ display: "flex", alignItems: "stretch", gap: 0, flexWrap: "wrap", rowGap: 10 }}>
                          {entries.map((e, i) => {
                            const gap = i > 0 ? duSigned(entries[i - 1].date, e.date) : null;
                            const cor = e.isCPL ? "#8e44ad" : "#2e86c1";
                            return (
                              <React.Fragment key={e.key + "_" + i}>
                                {i > 0 && (
                                  <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: "0 6px", minWidth: 34 }}>
                                    <span style={{ fontSize: 9, fontWeight: 700, color: gap > 15 ? "#e67e22" : "var(--gaq-text-3)" }}>+{gap} d.u.</span>
                                    <span style={{ color: "var(--gaq-line)", fontSize: 11, lineHeight: 1 }}>─▶</span>
                                  </div>
                                )}
                                <div style={{ background: "var(--gaq-bg-2)", border: `1px solid ${cor}33`, borderTop: `2px solid ${cor}`, borderRadius: 8, padding: "5px 10px", minWidth: 86 }}>
                                  <div style={{ fontSize: 8, fontWeight: 700, color: cor, textTransform: "uppercase", letterSpacing: ".03em", whiteSpace: "nowrap" }}>{e.label}</div>
                                  <div style={{ fontSize: 11.5, fontWeight: 700, color: "var(--gaq-text)", fontVariantNumeric: "tabular-nums" }}>{fmt(e.date)}</div>
                                </div>
                              </React.Fragment>
                            );
                          })}
                        </div>}
                  </div>
                );
              })}
            </div>
          </div>;
        })());
}
