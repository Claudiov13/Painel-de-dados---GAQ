import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function AlertaRCScreen({ alertRC, fArea, fBase, meta, onOpenTags, pgAlertRC, pgAlertRCSz, renderBuscaPanel, setPgAlertRC, setPgAlertRCSz, setSelProc, toggleWatch }) {
  return (<div>
          {renderBuscaPanel()}
          {(() => {
            const slaVencido = alertRC.filter(r => r.atrasoGeral).length;
            const critRC = alertRC.filter(r => (r.diasTotaisGestao||0) > 100).length;
            const atencRC = alertRC.filter(r => { const d = r.diasTotaisGestao||0; return d > 60 && d <= 100; }).length;
            const parados = alertRC.filter(r => (r.diasParado||0) > 15).length;
            const mediaRC = alertRC.length ? Math.round(alertRC.reduce((s,r)=>s+(r.diasTotaisGestao||0),0)/alertRC.length) : 0;
            const cardKpi = (label, value, color) => (
              <div key={label} style={{ background: "rgba(255,255,255,.10)", backdropFilter: "blur(8px)", borderRadius: 12, padding: "14px 16px", border: "1px solid rgba(255,255,255,.14)" }}>
                <div style={{ fontSize: 10, fontWeight: 600, opacity: .78, letterSpacing: ".06em", textTransform: "uppercase", marginBottom: 6, color: "#fff" }}>{label}</div>
                <div style={{ fontSize: 30, fontWeight: 700, lineHeight: 1, color: color || "#fff", letterSpacing: "-.02em", fontVariantNumeric: "tabular-nums" }}>{value}</div>
              </div>
            );
            return (
              <div style={{ background: "linear-gradient(135deg, #0a1f3d 0%, #2980b9 100%)", color: "#fff", borderRadius: "var(--gaq-r-xl, 16px)", padding: "24px 28px", marginBottom: 18, boxShadow: "var(--gaq-shadow-2, 0 4px 16px rgba(10,31,61,.18))", border: "none" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12, marginBottom: 18 }}>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 600, opacity: .72, letterSpacing: ".05em", textTransform: "uppercase", marginBottom: 4 }}>Alerta RC · Processos com RC em andamento</div>
                    <div style={{ fontSize: 22, fontWeight: 700, letterSpacing: "-.02em" }}>Acompanhamento de RCs</div>
                    <div style={{ fontSize: 12, opacity: .82, marginTop: 2 }}>Aging desde Recebimento RC · SLA por modalidade · Filtros globais aplicados</div>
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    {alertRC.length > 0 && (
                      <button onClick={() => exportToExcel(alertRC, "alerta_rc.xlsx")}
                        style={{ background: "rgba(255,255,255,.16)", color: "#fff", border: "1px solid rgba(255,255,255,.28)", borderRadius: 10, padding: "8px 16px", cursor: "pointer", fontSize: 12, fontWeight: 600, backdropFilter: "blur(8px)" }}>
                        Exportar Excel
                      </button>
                    )}
                    {alertRC.length > 0 && (
                      <button onClick={() => { const html = gerarRelatorioAlertaRC({ processos: alertRC, area: fArea || "Todas as Áreas", meta, fBase }); const win = window.open("", "_blank"); win.document.write(html); win.document.close(); }}
                        style={{ background: "#fff", color: "#1a5276", border: "none", borderRadius: 10, padding: "8px 16px", cursor: "pointer", fontSize: 12, fontWeight: 700 }}>
                        PDF Alerta RC
                      </button>
                    )}
                  </div>
                </div>
                <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 18 }}>
                  <div>
                    <div style={{ fontSize: 10, fontWeight: 700, opacity: .82, letterSpacing: ".06em", textTransform: "uppercase", marginBottom: 10, paddingBottom: 6, borderBottom: "1px solid rgba(255,255,255,.18)" }}>
                      Crítico — SLA Vencido / Aging Alto
                    </div>
                    <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10 }}>
                      {cardKpi("Total RC em Andamento", alertRC.length, "#fff")}
                      {cardKpi("SLA Vencido", slaVencido, slaVencido > 0 ? "#fecaca" : "#fff")}
                      {cardKpi("Crítico > 100 d.u.", critRC, critRC > 0 ? "#fecaca" : "#fff")}
                    </div>
                  </div>
                  <div>
                    <div style={{ fontSize: 10, fontWeight: 700, opacity: .82, letterSpacing: ".06em", textTransform: "uppercase", marginBottom: 10, paddingBottom: 6, borderBottom: "1px solid rgba(255,255,255,.18)" }}>
                      Atenção — Monitoramento
                    </div>
                    <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10 }}>
                      {cardKpi("Atenção 60–100 d.u.", atencRC, atencRC > 0 ? "#fed7aa" : "#fff")}
                      {cardKpi("Parados > 15 d", parados, parados > 0 ? "#fed7aa" : "#fff")}
                      {cardKpi("Média d.u. (RC)", mediaRC + " d.u.", mediaRC > 60 ? "#fecaca" : "#fff")}
                    </div>
                  </div>
                </div>
              </div>
            );
          })()}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, flexWrap: "wrap", gap: 8 }}>
            <div style={{ fontWeight: 700, fontSize: 14, color: "var(--text)" }}>Processos com RC ({alertRC.length})</div>
          </div>
          {alertRC.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 13, textAlign: "center", padding: 30 }}>Nenhum processo com RC em andamento.</div>
            : alertRC.slice((pgAlertRC-1)*pgAlertRCSz, pgAlertRC*pgAlertRCSz).map((r, i) => {
              const cor = r.diasTotaisGestao > 150 ? "#922b21" : r.diasTotaisGestao > 100 ? "#c0392b" : r.diasTotaisGestao > 60 ? "#e67e22" : "#27ae60";
              const tags = TagsManager.getForProcess(r.ProcessKey, r.TicketSD);
              const watched = WatchlistManager.isWatched(r.ProcessKey);
              const rAtivoColor = r.faseSubarea === "CPL" ? "#8e44ad" : r.faseSubarea === "Scont" ? "#16a085" : "#2e86c1";
              return (
                <div key={r.ProcessKey + i} style={{ background: "var(--card)", borderRadius: 9, padding: "11px 14px", boxShadow: "var(--shadow)", marginBottom: 7, borderLeft: `5px solid ${cor}`, outline: watched ? "2px solid #f39c1288" : "none", cursor: "pointer" }}
                  onClick={() => setSelProc(r)}>
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 10 }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      {/* Linha 1: chips de identificação */}
                      <div style={{ display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center", marginBottom: 4 }}>
                        {watched && <span style={{ fontSize: 11, color: "#f39c12" }}>👁</span>}
                        <span style={{ fontSize: 9, color: "var(--text3)", background: "var(--card2)", borderRadius: 3, padding: "0 4px" }}>{r.ProcessKey}</span>
                        <span style={{ fontWeight: 700, color: "#2e86c1", fontSize: 12 }}>{r.NumRC}</span>
                        {r.TicketSD && <span style={{ fontSize: 10, color: "#27ae60" }}>Pré-compra: {r.TicketSD}</span>}
                        {r.atrasoGeral && <span style={{ fontSize: 9, background: "#e74c3c", color: "#fff", borderRadius: 4, padding: "1px 6px", fontWeight: 700 }}>ATRASO &gt;{r.prazoGeral} d.u.</span>}
                        {r["Área Requisitante"] && <span style={{ fontSize: 9, background: "#2e86c118", color: "#2e86c1", borderRadius: 3, padding: "0 5px", fontWeight: 600 }}>{r["Área Requisitante"]}</span>}
                        <span style={{ fontSize: 10, color: "#2e86c1" }}>{r.Modalidade || "—"}</span>
                        <span onClick={(e) => { e.stopPropagation(); onOpenTags(r); }} style={{ fontSize: 10, cursor: "pointer", color: "#34495e", textDecoration: "underline" }}>+tag</span>
                        <span onClick={(e) => { e.stopPropagation(); toggleWatch(r); }} style={{ fontSize: 10, cursor: "pointer", color: watched ? "#f39c12" : "#bbb", fontWeight: watched ? 700 : 400 }}>{watched ? "★" : "☆"}</span>
                      </div>
                      {/* Linha 2: responsável ativo em destaque */}
                      <div style={{ display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center", marginBottom: 3 }}>
                        <span style={{ fontWeight: 800, fontSize: 12, color: rAtivoColor }}>◉ {r.respAtivo || "—"}</span>
                        {r.faseSubarea && <span style={{ fontSize: 9, background: `${rAtivoColor}18`, color: rAtivoColor, borderRadius: 4, padding: "1px 6px", fontWeight: 600 }}>{r.faseSubarea}</span>}
                        {r.statusDet && <span style={{ fontSize: 9, color: "#8e44ad", background: "#8e44ad11", borderRadius: 4, padding: "1px 6px" }}>{r.statusDet}</span>}
                        {r.faseAtual && r.faseAtual !== "—" && <span style={{ fontSize: 9, background: "#2e86c115", borderRadius: 4, padding: "0 5px", color: "#2e86c1" }}>{r.faseAtual}</span>}
                      </div>
                      {/* Linha 3: objeto */}
                      <div style={{ fontSize: 11, color: "var(--text)", marginBottom: 3, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{(r.Objeto || "").slice(0, 100) || "—"}</div>
                      {/* Linha 4: todos os responsáveis */}
                      <div style={{ display: "flex", gap: 4, flexWrap: "wrap", fontSize: 9, color: "var(--text3)" }}>
                        {r.Comprador && <span>Comp: <b style={{ color: "var(--text2)" }}>{r.Comprador}</b></span>}
                        {r.Avaliador && <span>· Aval: <b style={{ color: "var(--text2)" }}>{r.Avaliador}</b></span>}
                        {r.cplResp && <span>· CPL: <b style={{ color: "#8e44ad" }}>{r.cplResp}</b></span>}
                        {r.Pregoeiro && r.Pregoeiro !== r.cplResp && <span>· Preg: <b style={{ color: "#8e44ad" }}>{r.Pregoeiro}</b></span>}
                        {r.AnalistaContrato && <span>· Scont: <b style={{ color: "#16a085" }}>{r.AnalistaContrato}</b></span>}
                        {r.AdvogadoResp && <span>· Adv: <b style={{ color: "#16a085" }}>{r.AdvogadoResp}</b></span>}
                      </div>
                      <div style={{ marginTop: 3 }}>{tags.map(t => <TagBadge key={t} tag={t} small />)}</div>
                    </div>
                    <div style={{ textAlign: "center", minWidth: 60, flexShrink: 0 }}>
                      <div style={{ fontSize: 22, fontWeight: 800, color: cor }}>{r.diasTotaisGestao}</div>
                      <div style={{ fontSize: 9, color: "var(--text3)" }}>d.u. (RC)</div>
                      {r.diasTotais !== r.diasTotaisGestao && r.diasTotais > 0 && <div style={{ fontSize: 8, color: "var(--text3)", marginTop: 1 }}>Total: {r.diasTotais} d.u.</div>}
                      {r.diasParado > 15 && <div style={{ fontSize: 8, fontWeight: 700, color: "#e67e22", marginTop: 2 }}>últ. mov. {r.diasParado}d</div>}
                      {r.dataEntrega && (() => {
                        const dc = r.diasParaEntrega < 0 ? "#e74c3c" : r.diasParaEntrega <= 7 ? "#ff6b6b" : r.diasParaEntrega <= 30 ? "#fd9644" : "#27ae60";
                        const dl = r.diasParaEntrega < 0 ? `Venc. ${Math.abs(r.diasParaEntrega)}d` : r.diasParaEntrega === 0 ? "Hoje!" : `+${r.diasParaEntrega}d`;
                        return <div style={{ marginTop: 4, borderTop: "1px solid var(--border2)", paddingTop: 3 }}>
                          <div style={{ fontSize: 9, fontWeight: 800, color: dc }}>📅 {dl}</div>
                          <div style={{ fontSize: 7, color: "var(--text3)" }}>{r.dataEntrega.toLocaleDateString("pt-BR")}</div>
                        </div>;
                      })()}
                    </div>
                  </div>
                </div>
              );
            })}
          <Pagination total={alertRC.length} page={pgAlertRC} pageSize={pgAlertRCSz} onPage={setPgAlertRC} onSize={setPgAlertRCSz} />
        </div>);
}
