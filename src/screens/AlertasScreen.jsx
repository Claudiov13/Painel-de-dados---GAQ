import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function AlertasScreen({ alertCPL, alertNCL, fBack, onClickComp, onClickProc, onOpenTags, pgAlert, pgAlertSz, pgCpl, pgCplSz, pgHS, pgHSSz, pgStag, pgStagSz, phaseIntervals, renderBuscaPanel, setPgAlert, setPgAlertSz, setPgCpl, setPgCplSz, setPgHS, setPgHSSz, setPgStag, setPgStagSz, setSelProc, setShowHS, setShowSLA, setShowStag, showHS, showSLA, showStag, tagVersion, toggleWatch, watchVersion }) {
  return (<div className="anim-fade">
          {renderBuscaPanel()}

          {/* Header */}
          <div style={{ marginBottom: 18 }}>
            <div className="gaq-h1" style={{ marginBottom: 4 }}>Central de Alertas</div>
            <div className="gaq-body" style={{ color: "var(--gaq-text-3)" }}>Alertas ativos baseados em SLA, estagnacao e lead time por fase</div>
          </div>

          {/* KPIs resumo */}
          {(() => {
            const scored = fBack.map(r => { const sla = calcSLAClassification(r, phaseIntervals); return sla ? { ...r, _sla: sla } : null; }).filter(r => r && (r._sla.bucket === "critico" || r._sla.bucket === "sla_vencido")).sort((a,b) => (b._sla.pctSLA || 0) - (a._sla.pctSLA || 0));
            const stagnados = fBack.filter(r => r.diasParado > 15);
            const kpis = [
              { label: "Critico", value: scored.filter(r => r._sla.bucket === "critico").length, color: "var(--gaq-red)" },
              { label: "SLA Vencido", value: scored.filter(r => r._sla.bucket === "sla_vencido").length, color: "var(--gaq-orange)" },
              { label: "Estagnados", value: stagnados.length, color: "#e67e22" },
              { label: "NCL Criticos", value: alertNCL.length, color: "#2e86c1" },
              { label: "CPL Criticos", value: alertCPL.length, color: "#8e44ad" },
            ];
            return (
              <div className="gaq-kpi-grid" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", marginBottom: 18 }}>
                {kpis.map(k => (
                  <div key={k.label} className="gaq-kpi" style={{ borderTop: `3px solid ${k.color}` }}>
                    <span className="lbl">{k.label}</span>
                    <div className="val" style={{ color: k.color }}>{k.value}</div>
                  </div>
                ))}
              </div>
            );
          })()}

          {/* ── PAINEL: Classificação SLA (Crítico + SLA Vencido) ── */}
          {(() => {
          const scored = fBack.map(r => { const sla = calcSLAClassification(r, phaseIntervals); return sla ? { ...r, _sla: sla } : null; }).filter(r => r && (r._sla.bucket === "critico" || r._sla.bucket === "sla_vencido")).sort((a,b) => (b._sla.pctSLA || 0) - (a._sla.pctSLA || 0));
            const hsTotalPages = Math.ceil(scored.length / pgHSSz);
            const hsPage = scored.slice((pgHS-1)*pgHSSz, pgHS*pgHSSz);

            return scored.length > 0 && (
              <div className="gaq-card" style={{ marginBottom: 14, padding: 0, overflow: "hidden" }}>
                <div onClick={() => setShowHS(s => !s)} style={{ padding: "16px 22px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: showHS ? "1px solid var(--gaq-line)" : "none" }}>
                  <div>
                    <div className="gaq-h3" style={{ color: "var(--gaq-red)", marginBottom: 2 }}>SLA Vencido e Critico</div>
                    <div className="gaq-meta">{scored.length} processos fora do SLA da modalidade</div>
                  </div>
                  <Icon name={showHS ? "chevU" : "chevD"} size={14} color="var(--gaq-text-3)" />
                </div>
                {showHS && <div style={{ padding: "12px 22px 18px" }}>
                  {hsPage.map((r, i) => {
                    const bk = SLA_BUCKETS[r._sla.bucket];
                    return (
                    <div key={r.ProcessKey} onClick={() => setSelProc(r)} className="gaq-card" style={{ padding: "14px 18px", marginBottom: 10, borderRadius: 16, cursor: "pointer", display: "flex", gap: 14, alignItems: "center" }}>
                      <div style={{ width: 40, height: 40, borderRadius: "50%", background: bk.bg, border: `2px solid ${bk.color}`, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 11, flexShrink: 0, color: bk.color }}>{r._sla.pctSLALabel}%</div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center", marginBottom: 4 }}>
                          <span style={{ fontWeight: 600, fontSize: 13, color: "var(--gaq-text)" }}>{r.NumRC || r.TicketSD || "—"}</span>
                          {r.TicketSD && r.NumRC && <span style={{ fontSize: 11, color: "var(--gaq-green)" }}>Pré-compra {r.TicketSD}</span>}
                          <span style={{ fontSize: 11, color: "var(--gaq-text-3)" }}>{r.Modalidade || "—"}</span>
                          <span style={{ fontSize: 11, color: "var(--gaq-text-3)" }}>{r.respNCL || r.Pregoeiro || "—"}</span>
                        </div>
                        <div style={{ fontSize: 12, color: "var(--gaq-text-2)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{(r.Objeto||"").slice(0,100)}</div>
                      </div>
                      <div style={{ textAlign: "right", flexShrink: 0, display: "grid", gap: 4, justifyItems: "end" }}>
                        <span style={{ fontSize: 11, fontWeight: 700, color: bk.color, background: bk.bg, borderRadius: 999, padding: "4px 10px" }}>{bk.label}</span>
                        <span className="gaq-num" style={{ fontSize: 14, fontWeight: 700, color: bk.color }}>{r.diasTotais} d.u.</span>
                      </div>
                    </div>
                  );})}
                  {hsTotalPages > 1 && <Pagination total={scored.length} page={pgHS} pageSize={pgHSSz} onPage={setPgHS} onSize={setPgHSSz} />}
                </div>}
              </div>
            );
          })()}

          {/* ── PAINEL: Estagnação ── */}
          {(() => {
            const stagnados = fBack.filter(r => r.diasParado > 15).sort((a,b) => b.diasParado - a.diasParado);
            const stagTotalPages = Math.ceil(stagnados.length / pgStagSz);
            const stagPage = stagnados.slice((pgStag-1)*pgStagSz, pgStag*pgStagSz);

            return stagnados.length > 0 && (
              <div className="gaq-card" style={{ marginBottom: 14, padding: 0, overflow: "hidden" }}>
                <div onClick={() => setShowStag(s => !s)} style={{ padding: "16px 22px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: showStag ? "1px solid var(--gaq-line)" : "none" }}>
                  <div>
                    <div className="gaq-h3" style={{ color: "var(--gaq-orange)", marginBottom: 2 }}>Processos Estagnados</div>
                    <div className="gaq-meta">{stagnados.length} processos sem movimentacao ha mais de 15 d.u.</div>
                  </div>
                  <Icon name={showStag ? "chevU" : "chevD"} size={14} color="var(--gaq-text-3)" />
                </div>
                {showStag && <div style={{ padding: "12px 22px 18px" }}>
                  {stagPage.map((r, i) => (
                    <div key={r.ProcessKey} onClick={() => setSelProc(r)} className="gaq-card" style={{ padding: "14px 18px", marginBottom: 10, borderRadius: 16, cursor: "pointer", display: "flex", gap: 14, alignItems: "center" }}>
                      <div style={{ width: 40, height: 40, borderRadius: 12, background: r.diasParado > 30 ? "#c0392b" : "#e67e22", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 800, fontSize: 13, flexShrink: 0 }}>{r.diasParado}</div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center", marginBottom: 4 }}>
                          <span style={{ fontWeight: 600, fontSize: 13, color: "var(--gaq-text)" }}>{r.NumRC || r.TicketSD || "—"}</span>
                          {r.TicketSD && r.NumRC && <span style={{ fontSize: 11, color: "var(--gaq-green)" }}>Pré-compra {r.TicketSD}</span>}
                          <span style={{ fontSize: 11, color: "var(--gaq-text-3)" }}>{r.Modalidade || "—"}</span>
                          <span style={{ fontSize: 11, color: "var(--gaq-text-3)" }}>{r.respNCL || r.Pregoeiro || "—"}</span>
                        </div>
                        <div style={{ fontSize: 12, color: "var(--gaq-text-2)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{(r.Objeto||"").slice(0,100)}</div>
                      </div>
                      <div style={{ fontSize: 12, color: "var(--gaq-text-3)", textAlign: "right", flexShrink: 0 }}>total {r.diasTotais} d.u.</div>
                    </div>
                  ))}
                  {stagTotalPages > 1 && <Pagination total={stagnados.length} page={pgStag} pageSize={pgStagSz} onPage={setPgStag} onSize={setPgStagSz} />}
                </div>}
              </div>
            );
          })()}

          {/* ── PAINEL: SLA por Fase ── */}
          {(() => {
            const phases = [];
            for (let i = 0; i < TL_COLS.length - 1; i++) {
              const [lA] = TL_COLS[i], [lB, keyB] = TL_COLS[i+1];
              const [, keyA] = TL_COLS[i];
              const pair = keyA + "\u2192" + keyB;
              const gd = phaseIntervals.global[pair];
              if (!gd || !gd.avg) continue;
              phases.push({ label: lA + " → " + lB, avg: gd.avg, med: gd.med, n: gd.n, isCPL: isCplTimelineKey(keyB) });
            }
            return phases.length > 0 && (
              <div className="gaq-card" style={{ marginBottom: 14, padding: 0, overflow: "hidden" }}>
                <div onClick={() => setShowSLA(s => !s)} style={{ padding: "16px 22px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: showSLA ? "1px solid var(--gaq-line)" : "none" }}>
                  <div>
                    <div className="gaq-h3" style={{ color: "var(--gaq-blue)", marginBottom: 2 }}>SLA por Fase</div>
                    <div className="gaq-meta">Lead time medio entre etapas do processo</div>
                  </div>
                  <Icon name={showSLA ? "chevU" : "chevD"} size={14} color="var(--gaq-text-3)" />
                </div>
                {showSLA && <div style={{ padding: "12px 22px 18px" }}>
                  <div className="gaq-meta" style={{ marginBottom: 12 }}>Verde ≤10 d.u. · Amarelo ≤25 d.u. · Vermelho &gt;25 d.u.</div>
                  <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    {phases.map((p, i) => {
                      const cor = p.avg <= 10 ? "#27ae60" : p.avg <= 25 ? "#f39c12" : "#c0392b";
                      return (
                        <div key={i} className="gaq-card" style={{ padding: "12px 16px", borderRadius: 14, borderLeft: `3px solid ${cor}` }}>
                          <div style={{ fontSize: 11, fontWeight: 600, color: p.isCPL ? "#8e44ad" : "var(--gaq-text-2)", textTransform: "uppercase", marginBottom: 4, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p.label}</div>
                          <div style={{ display: "flex", gap: 10, alignItems: "baseline" }}>
                            <span className="gaq-num" style={{ fontSize: 20, fontWeight: 800, color: cor }}>{p.avg}</span>
                            <span className="gaq-meta">d.u. media · {p.n} amostras</span>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>}
              </div>
            );
          })()}

          {/* ── NCL Críticos ── */}
          <div className="gaq-card" style={{ padding: 22, marginBottom: 14 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
              <div>
                <div className="gaq-h3" style={{ color: "#2e86c1" }}>NCL — Criticos (&gt;50 d.u.)</div>
                <div className="gaq-meta">Em andamento · d.u. desde abertura &gt; 50</div>
              </div>
              {alertNCL.length > 0 && <button onClick={() => exportToExcel(alertNCL, "alertas_ncl.xlsx")} className="gaq-btn is-primary" style={{ fontSize: 11 }}>Exportar NCL</button>}
            </div>
            {alertNCL.length === 0 ? <div style={{ color: "var(--gaq-text-3)", fontSize: 13 }}>Nenhum critico NCL.</div>
              : alertNCL.slice((pgAlert-1)*pgAlertSz, pgAlert*pgAlertSz).map((r, i) => <ProcCard key={i} r={r} tipo="ncl" onClickComp={onClickComp} onClickProc={onClickProc} onOpenTags={onOpenTags} tagVersion={tagVersion} onToggleWatch={toggleWatch} watchVersion={watchVersion} />)}
            <Pagination total={alertNCL.length} page={pgAlert} pageSize={pgAlertSz} onPage={setPgAlert} onSize={setPgAlertSz} />
          </div>

          {/* ── CPL Críticos ── */}
          <div className="gaq-card" style={{ padding: 22 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
              <div>
                <div className="gaq-h3" style={{ color: "#8e44ad" }}>CPL — Criticos (&gt;50 d.u.)</div>
                <div className="gaq-meta">CPL=VERDADEIRO · em andamento · d.u. &gt; 50</div>
              </div>
              {alertCPL.length > 0 && <button onClick={() => exportToExcel(alertCPL, "alertas_cpl.xlsx")} className="gaq-btn is-primary" style={{ fontSize: 11, background: "#8e44ad" }}>Exportar CPL</button>}
            </div>
            {alertCPL.length === 0 ? <div style={{ color: "var(--gaq-text-3)", fontSize: 13 }}>Nenhum critico CPL.</div>
              : alertCPL.slice((pgCpl-1)*pgCplSz, pgCpl*pgCplSz).map((r, i) => <ProcCard key={i} r={r} tipo="cpl" onClickComp={onClickComp} onClickProc={onClickProc} onOpenTags={onOpenTags} tagVersion={tagVersion} onToggleWatch={toggleWatch} watchVersion={watchVersion} />)}
            <Pagination total={alertCPL.length} page={pgCpl} pageSize={pgCplSz} onPage={setPgCpl} onSize={setPgCplSz} />
          </div>
        </div>);
}
