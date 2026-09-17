import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function MapeamentoScreen({ mapInsights, renderBuscaPanel, setAba, setFArea }) {
  return (<div className="anim-fade">
          {renderBuscaPanel()}
          <div style={{ marginBottom: 18 }}>
            <div className="gaq-h1" style={{ marginBottom: 4 }}>Mapeamento</div>
            <div className="gaq-body" style={{ color: "var(--gaq-text-3)" }}>Cancelamentos e fracassos · padroes recorrentes nas RCs encerradas sem efetivacao</div>
          </div>

          {/* KPIs */}
          <div className="gaq-kpi-grid" style={{ gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", marginBottom: 18 }}>
            {mapInsights.taxas.map(t => (
              <div key={t.label} className="gaq-kpi" style={{ borderTop: `3px solid ${t.color}` }}>
                <span className="lbl">{t.label}</span>
                <div className="val" style={{ color: t.color }}>{t.val}</div>
              </div>
            ))}
          </div>

          {/* NOVO: Evolução anual */}
          {mapInsights.evolucaoAnual.length > 1 && <WCard title="Evolução Anual — Cancelamentos + Fracassos" border="#e67e22">
            <ResponsiveContainer width="100%" height={180}>
              <BarChart data={mapInsights.evolucaoAnual}>
                <XAxis dataKey="ano" fontSize={11} />
                <YAxis fontSize={11} />
                <Tooltip formatter={(v, name) => [name === "cf" ? `${v} problemáticos` : `${v} total`, name === "cf" ? "Canc./Frac." : "Total"]} />
                <Bar dataKey="total" fill="#85c1e9" radius={[3,3,0,0]} name="Total" />
                <Bar dataKey="cf" fill="#e74c3c" radius={[3,3,0,0]} name="Canc./Frac." />
              </BarChart>
            </ResponsiveContainer>
            <div style={{ display: "flex", gap: 16, marginTop: 8, fontSize: 11, color: "var(--text2)" }}>
              {mapInsights.evolucaoAnual.map(e => (
                <span key={e.ano}><b>{e.ano}:</b> {e.taxa}% ({e.cf}/{e.total})</span>
              ))}
            </div>
          </WCard>}

          <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
            <div className="gaq-card" style={{ padding: 22 }}>
              <div className="gaq-h3" style={{ marginBottom: 12 }}>Taxa por Area (min. 4 proc.)</div>
              {mapInsights.porArea.length === 0 ? <div className="gaq-meta">Sem dados.</div> : mapInsights.porArea.map((x, i) => (
                <div key={x.name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", borderBottom: "1px solid var(--gaq-line)", fontSize: 12, cursor: "pointer" }}
                  onClick={() => { setFArea(x.name); setAba("processos"); }}>
                  <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 200, color: "var(--gaq-text)" }}><span className="gaq-meta" style={{ marginRight: 6 }}>#{i+1}</span>{x.name}</span>
                  <span style={{ display: "flex", gap: 6, alignItems: "center" }}>
                    <span style={{ background: x.taxa > 40 ? "#c0392b" : x.taxa > 20 ? "#e67e22" : "#f39c12", color: "#fff", borderRadius: 999, padding: "2px 10px", fontSize: 11, fontWeight: 700 }}>{x.taxa}%</span>
                    <span className="gaq-meta">{x.cf}/{x.tot}</span>
                  </span>
                </div>
              ))}
            </div>
            <div className="gaq-card" style={{ padding: 22 }}>
              <div className="gaq-h3" style={{ marginBottom: 12 }}>Taxa por Modalidade</div>
              {mapInsights.porMod.length === 0 ? <div className="gaq-meta">Sem dados.</div> : mapInsights.porMod.map((x, i) => (
                <div key={x.name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", borderBottom: "1px solid var(--gaq-line)", fontSize: 12 }}>
                  <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 180, color: "var(--gaq-text)" }}><span className="gaq-meta" style={{ marginRight: 6 }}>#{i+1}</span>{x.name}</span>
                  <span style={{ display: "flex", gap: 6, alignItems: "center" }}>
                    <span style={{ background: x.taxa > 40 ? "#c0392b" : x.taxa > 20 ? "#e67e22" : "#f39c12", color: "#fff", borderRadius: 999, padding: "2px 10px", fontSize: 11, fontWeight: 700 }}>{x.taxa}%</span>
                    {x.medLT && <span className="gaq-meta">{x.medLT} d.u.</span>}
                  </span>
                </div>
              ))}
            </div>
          </div>

          {/* NOVO: Tempo médio até cancelamento vs conclusão */}
          {mapInsights.tempoMedio.length > 0 && <WCard title="Tempo Médio até Cancelamento/Fracasso vs. Conclusão (por Modalidade)" border="#2e86c1">
            <div style={{ marginBottom: 4 }}>
              {mapInsights.tempoMedio.map(t => (
                <div key={t.name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "5px 0", borderBottom: "1px solid var(--border2)", fontSize: 12 }}>
                  <span style={{ color: "var(--text)", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.name} <span style={{ color: "var(--text3)", fontSize: 10 }}>({t.qtdCF} proc.)</span></span>
                  <span style={{ display: "flex", gap: 10, alignItems: "center" }}>
                    <span style={{ color: "#c0392b", fontWeight: 700 }}>{t.medCF} d.u. <span style={{ fontWeight: 400, fontSize: 10 }}>canc./frac.</span></span>
                    {t.medConc && <span style={{ color: "#27ae60", fontWeight: 700 }}>{t.medConc} d.u. <span style={{ fontWeight: 400, fontSize: 10 }}>concluído</span></span>}
                  </span>
                </div>
              ))}
            </div>
            <div style={{ fontSize: 10, color: "var(--text3)", marginTop: 4 }}>Insight: se o tempo até cancelamento é alto, há custo de oportunidade significativo.</div>
          </WCard>}

          {/* Reincidentes */}
          {mapInsights.reincidentes.length > 0 && <div className="gaq-card" style={{ padding: 22, marginBottom: 16 }}>
            <div className="gaq-h3" style={{ marginBottom: 12 }}>Objetos com reincidencia</div>
            <div className="gaq-meta" style={{ marginBottom: 14 }}>2+ tentativas de contratacao no ano</div>
            {mapInsights.reincidentes.map((r, i) => (
              <div key={i} className="gaq-card" style={{ padding: "14px 18px", marginBottom: 10, borderRadius: 16, borderLeft: `3px solid ${r.count >= 3 ? "#c0392b" : "#e67e22"}` }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 13, fontWeight: 500, color: "var(--gaq-text)", marginBottom: 6 }}>{r.objeto}</div>
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap", fontSize: 11 }}>
                      <span style={{ color: "#e67e22", background: "rgba(230,126,34,0.1)", borderRadius: 999, padding: "2px 10px", fontWeight: 600 }}>{r.cancelados} cancelados</span>
                      <span style={{ color: "#8e44ad", background: "rgba(142,68,173,0.1)", borderRadius: 999, padding: "2px 10px", fontWeight: 600 }}>{r.fracassados} fracassados</span>
                      <span className="gaq-meta">Areas: {r.areas.join(", ") || "—"}</span>
                    </div>
                    <div className="gaq-meta" style={{ marginTop: 4 }}>RCs: {r.nums.join(" · ")}</div>
                  </div>
                  <div style={{ fontWeight: 800, fontSize: 24, color: r.count >= 3 ? "#c0392b" : "#e67e22", flexShrink: 0 }}>{r.count}x</div>
                </div>
              </div>
            ))}
          </div>}

          {/* Padroes fracasso */}
          <div className="gaq-card" style={{ padding: 22, marginBottom: 16 }}>
            <div className="gaq-h3" style={{ color: "#8e44ad", marginBottom: 4 }}>Padroes exclusivos de fracassados (≥2)</div>
            <div className="gaq-meta" style={{ marginBottom: 14 }}>Foco em objetos com historico de disputa fracassada.</div>
            {mapInsights.padFracasso.length === 0 ? <div className="gaq-meta" style={{ textAlign: "center", padding: 16 }}>Nenhum padrao.</div>
              : mapInsights.padFracasso.map(p => <PatternCard key={p.word} p={p} accent="#8e44ad" />)}
          </div>

          {/* Padroes gerais */}
          <div className="gaq-card" style={{ padding: 22, marginBottom: 16 }}>
            <div className="gaq-h3" style={{ marginBottom: 4 }}>Padroes Gerais — Cancelados + Fracassados (≥3)</div>
            <div className="gaq-meta" style={{ marginBottom: 14 }}>Padroes recorrentes na base de dados.</div>
            {mapInsights.padroes.length === 0 ? <div className="gaq-meta" style={{ textAlign: "center", padding: 16 }}>Nenhum padrao.</div>
              : mapInsights.padroes.map(p => <PatternCard key={p.word} p={p} />)}
          </div>
        </div>);
}
