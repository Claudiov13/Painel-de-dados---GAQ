import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function AlertaSDScreen({ alertSD, alertSDAberto, alertSDEncerrado, fArea, fBase, meta, onOpenTags, pgAlertSD, pgAlertSDEnc, pgAlertSDEncSz, pgAlertSDSz, renderBuscaPanel, sdCopyStatus, serviceDeskData, setPgAlertSD, setPgAlertSDEnc, setPgAlertSDEncSz, setPgAlertSDSz, setSdCopyStatus, setSelProc, toggleWatch }) {
  return (<div>
          {renderBuscaPanel()}

          {/* ── Header gradient (estilo Painel Executivo) ── */}
          {(() => {
            const sdAtraso = alertSDAberto.filter(r => r.atrasoSD).length;
            const sdAtenc = alertSDAberto.filter(r => (r.diasSDAberto||0) > 7 && !r.atrasoSD).length;
            const encUrg = alertSDEncerrado.filter(r => (r.diasDesdeEncSD||0) > 7).length;
            const encAtenc = alertSDEncerrado.filter(r => { const d = r.diasDesdeEncSD||0; return d > 3 && d <= 7; }).length;
            const cardKpi = (label, value, color) => (
              <div key={label} style={{ background: "rgba(255,255,255,.10)", backdropFilter: "blur(8px)", borderRadius: 12, padding: "14px 16px", border: "1px solid rgba(255,255,255,.14)" }}>
                <div style={{ fontSize: 10, fontWeight: 600, opacity: .78, letterSpacing: ".06em", textTransform: "uppercase", marginBottom: 6, color: "#fff" }}>{label}</div>
                <div style={{ fontSize: 30, fontWeight: 700, lineHeight: 1, color: color || "#fff", letterSpacing: "-.02em", fontVariantNumeric: "tabular-nums" }}>{value}</div>
              </div>
            );
            return (
              <div style={{ background: "linear-gradient(135deg, #064e3b 0%, #059669 100%)", color: "#fff", borderRadius: "var(--gaq-r-xl, 16px)", padding: "24px 28px", marginBottom: 18, boxShadow: "var(--gaq-shadow-2, 0 4px 16px rgba(6,78,59,.18))", border: "none" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12, marginBottom: 18 }}>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 600, opacity: .72, letterSpacing: ".05em", textTransform: "uppercase", marginBottom: 4 }}>Alerta Pré-compra · Tickets sem RC</div>
                    <div style={{ fontSize: 22, fontWeight: 700, letterSpacing: "-.02em" }}>Acompanhamento Pré-compra</div>
                    <div style={{ fontSize: 12, opacity: .82, marginTop: 2 }}>SLA pré-compra: {PRAZO_SD} d.u. · Filtros globais aplicados</div>
                  </div>
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    {alertSD.length > 0 && (
                      <button onClick={() => { const sdSummaries = {}; [...alertSDAberto, ...alertSDEncerrado].forEach(r => { if (r.TicketSD) sdSummaries[String(r.TicketSD).trim()] = getServiceDeskAlertSummary(r, serviceDeskData); }); exportAlertaSDToExcel(alertSD, sdSummaries); }}
                        style={{ background: "rgba(255,255,255,.16)", color: "#fff", border: "1px solid rgba(255,255,255,.28)", borderRadius: 10, padding: "8px 16px", cursor: "pointer", fontSize: 12, fontWeight: 600, backdropFilter: "blur(8px)" }}>
                        Exportar Excel
                      </button>
                    )}
                    {alertSD.length > 0 && (
                      <button onClick={() => { const sdSummaries = {}; [...alertSDAberto, ...alertSDEncerrado].forEach(r => { if (r.TicketSD) sdSummaries[String(r.TicketSD).trim()] = getServiceDeskAlertSummary(r, serviceDeskData); }); const html = gerarRelatorioAlertaSD({ sdAberto: alertSDAberto, sdEncerrado: alertSDEncerrado, area: fArea || "Todas as Áreas", meta, fBase, sdSummaries }); const win = window.open("", "_blank"); win.document.write(html); win.document.close(); }}
                        style={{ background: "#fff", color: "#065f46", border: "none", borderRadius: 10, padding: "8px 16px", cursor: "pointer", fontSize: 12, fontWeight: 700 }}>
                        PDF Alerta Pré-compra
                      </button>
                    )}
                    {alertSDAberto.length > 0 && (
                      <button onClick={async () => { const ids = alertSDAberto.filter(r => { if (!r.TicketSD) return false; const rec = getServiceDeskRecord(String(r.TicketSD).trim(), serviceDeskData); if (!rec) return true; const en = nrm((rec.informacoes_do_chamado || {}).estado || ""); return !en.includes("finalizado") && !en.includes("fechado") && !en.includes("concluido"); }).map(r => String(r.TicketSD || "").trim()).filter(Boolean); const script = gerarScriptExtratorOTRS(ids, "rapido"); try { await navigator.clipboard.writeText(script); setSdCopyStatus("rapido"); setTimeout(() => setSdCopyStatus(""), 2500); } catch { setSdCopyStatus(""); } }}
                        style={{ background: sdCopyStatus === "rapido" ? "#d1fae5" : "rgba(255,255,255,.13)", color: sdCopyStatus === "rapido" ? "#065f46" : "#fff", border: "1px solid rgba(255,255,255,.28)", borderRadius: 10, padding: "8px 14px", cursor: "pointer", fontSize: 11, fontWeight: 600, backdropFilter: "blur(8px)", transition: "all .2s", whiteSpace: "nowrap" }}>
                        {sdCopyStatus === "rapido" ? "✓ Copiado!" : `Script Rápido (${alertSDAberto.length})`}
                      </button>
                    )}
                    {alertSDAberto.length > 0 && (
                      <button onClick={async () => { const ids = alertSDAberto.filter(r => { if (!r.TicketSD) return false; const rec = getServiceDeskRecord(String(r.TicketSD).trim(), serviceDeskData); if (!rec) return true; const en = nrm((rec.informacoes_do_chamado || {}).estado || ""); return !en.includes("finalizado") && !en.includes("fechado") && !en.includes("concluido"); }).map(r => String(r.TicketSD || "").trim()).filter(Boolean); const script = gerarScriptExtratorOTRS(ids, "diurno"); try { await navigator.clipboard.writeText(script); setSdCopyStatus("diurno"); setTimeout(() => setSdCopyStatus(""), 2500); } catch { setSdCopyStatus(""); } }}
                        style={{ background: sdCopyStatus === "diurno" ? "#d1fae5" : "rgba(255,255,255,.13)", color: sdCopyStatus === "diurno" ? "#065f46" : "#fff", border: "1px solid rgba(255,255,255,.28)", borderRadius: 10, padding: "8px 14px", cursor: "pointer", fontSize: 11, fontWeight: 600, backdropFilter: "blur(8px)", transition: "all .2s", whiteSpace: "nowrap" }}>
                        {sdCopyStatus === "diurno" ? "✓ Copiado!" : `Script Diurno (${alertSDAberto.length})`}
                      </button>
                    )}
                  </div>
                </div>

                <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(2, 1fr)", gap: 18 }}>
                  {/* Bloco 1: Pré-compra em andamento */}
                  <div>
                    <div style={{ fontSize: 10, fontWeight: 700, opacity: .82, letterSpacing: ".06em", textTransform: "uppercase", marginBottom: 10, paddingBottom: 6, borderBottom: "1px solid rgba(255,255,255,.18)" }}>
                      Pré-compra em andamento — Acompanhamento necessário GAQ
                    </div>
                    <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10 }}>
                      {cardKpi("Em Andamento", alertSDAberto.length, "#fff")}
                      {cardKpi(`Atraso > ${PRAZO_SD} d.u.`, sdAtraso, sdAtraso > 0 ? "#fecaca" : "#fff")}
                      {cardKpi(`Atenção 7–${PRAZO_SD} d.u.`, sdAtenc, sdAtenc > 0 ? "#fed7aa" : "#fff")}
                    </div>
                  </div>

                  {/* Bloco 2: Pré-compra encerrada aguardando RC */}
                  <div>
                    <div style={{ fontSize: 10, fontWeight: 700, opacity: .82, letterSpacing: ".06em", textTransform: "uppercase", marginBottom: 10, paddingBottom: 6, borderBottom: "1px solid rgba(255,255,255,.18)" }}>
                      Pré-compra encerrada — Aguardando RC da Área
                    </div>
                    <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10 }}>
                      {cardKpi("Enc. sem RC", alertSDEncerrado.length, "#fff")}
                      {cardKpi("Urgente > 7 d.u.", encUrg, encUrg > 0 ? "#fecaca" : "#fff")}
                      {cardKpi("Atenção 3–7 d.u.", encAtenc, encAtenc > 0 ? "#fed7aa" : "#fff")}
                    </div>
                  </div>
                </div>
              </div>
            );
          })()}

          {alertSD.length === 0
            ? <div style={{ color: "var(--text3)", fontSize: 13, textAlign: "center", padding: 40, background: "var(--card)", borderRadius: 12, border: "1px solid var(--border2)" }}>
                Nenhum ticket de pré-compra sem RC em andamento.
              </div>
            : <>
              {/* ══ SEÇÃO 1: PRÉ-COMPRA EM ANDAMENTO (sem encerramento) ══ */}
              {alertSDAberto.length > 0 && <>
                <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
                  <div style={{ width: 4, height: 20, background: "#22c55e", borderRadius: 2, flexShrink: 0 }} />
                  <div style={{ fontSize: 14, fontWeight: 700, color: "var(--text)" }}>Pré-compra em andamento</div>
                  <span style={{ background: "#dcfce7", color: "#15803d", borderRadius: 20, padding: "2px 10px", fontSize: 11, fontWeight: 700 }}>{alertSDAberto.length}</span>
                  <div style={{ fontSize: 11, color: "var(--text3)" }}>SLA: {PRAZO_SD} d.u. · Sem data de encerramento</div>
                </div>
                {alertSDAberto.slice((pgAlertSD-1)*pgAlertSDSz, pgAlertSD*pgAlertSDSz).map((r, i) => {
                  const du_ = r.diasSDAberto || 0;
                  const cor = du_ > PRAZO_SD ? "#dc2626" : du_ > 7 ? "#ea580c" : "#16a34a";
                  const borderCor = du_ > PRAZO_SD ? "#fca5a5" : du_ > 7 ? "#fed7aa" : "#bbf7d0";
                  const bgBadge = du_ > PRAZO_SD ? "#fff1f2" : du_ > 7 ? "#fff7ed" : "#f0fdf4";
                  const tags = TagsManager.getForProcess(r.ProcessKey, r.TicketSD);
                  const watched = WatchlistManager.isWatched(r.ProcessKey);
                  const sdAlert = getServiceDeskAlertSummary(r, serviceDeskData);
                  return (
                    <div key={r.ProcessKey + i}
                      style={{ background: "var(--card)", borderRadius: 12, padding: "12px 16px", boxShadow: "0 1px 4px rgba(0,0,0,.06)", marginBottom: 8, border: `1px solid ${borderCor}`, borderLeft: `4px solid ${cor}`, cursor: "pointer", outline: watched ? "2px solid #f59e0b44" : "none" }}
                      onClick={() => setSelProc(r)}>
                      <div style={{ display: "flex", justifyContent: "space-between", gap: 12 }}>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center", marginBottom: 5 }}>
                            {watched && <span style={{ fontSize: 11, color: "#f59e0b" }}>⭐</span>}
                            <span style={{ fontWeight: 700, color: "var(--gaq-text)", fontSize: 13 }}>Pré-compra {r.TicketSD}</span>
                            <span style={{ fontSize: 10, background: "transparent", borderRadius: 6, padding: "1px 7px", color: "#dc2626", fontWeight: 600, border: "1px solid #fca5a5" }}>sem RC</span>
                            {r.atrasoSD && <span style={{ fontSize: 10, background: "#dc2626", color: "#fff", borderRadius: 6, padding: "2px 9px", fontWeight: 700 }}>ATRASO &gt;{PRAZO_SD} d.u.</span>}
                            {!r.atrasoSD && du_ > 7 && <span style={{ fontSize: 10, background: "transparent", color: "#c2410c", borderRadius: 6, padding: "2px 9px", fontWeight: 600, border: "1px solid #fed7aa" }}>ATENÇÃO {du_} d.u.</span>}
                            {r["Área Requisitante"] && <span style={{ fontSize: 10, background: "#eff6ff", color: "#2563eb", borderRadius: 6, padding: "1px 7px", fontWeight: 600 }}>{r["Área Requisitante"]}</span>}
                            {r.Modalidade && <span style={{ fontSize: 10, color: "var(--text3)" }}>{r.Modalidade}</span>}
                            <span onClick={(e) => { e.stopPropagation(); onOpenTags(r); }} style={{ fontSize: 10, cursor: "pointer", color: "var(--text3)", textDecoration: "underline" }}>+tag</span>
                            <span onClick={(e) => { e.stopPropagation(); toggleWatch(r); }} style={{ fontSize: 10, cursor: "pointer", color: watched ? "#f59e0b" : "var(--border2)", fontWeight: watched ? 700 : 400 }}>{watched ? "★" : "☆"}</span>
                          </div>
                          {r.statusDet && <div style={{ display: "inline-block", fontSize: 10, color: "#7c3aed", background: "#f5f3ff", borderRadius: 5, padding: "1px 7px", marginBottom: 4 }}>{r.statusDet}</div>}
                          <div style={{ fontSize: 12, color: "var(--text)", marginBottom: 4, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{(r.Objeto || "").slice(0, 100) || "—"}</div>
                          <ServiceDeskAlertStrip summary={sdAlert} />
                          <div style={{ display: "flex", gap: 5, flexWrap: "wrap", fontSize: 10, color: "var(--text3)" }}>
                            {r.Comprador && <span>Comp: <b style={{ color: "var(--text2)" }}>{r.Comprador}</b></span>}
                            {r.Avaliador && <span>· Aval: <b style={{ color: "var(--text2)" }}>{r.Avaliador}</b></span>}
                            {r.aberturaSD && <span>· Abertura: <b style={{ color: "var(--text2)" }}>{r.aberturaSD.toLocaleDateString("pt-BR")}</b></span>}
                          </div>
                          {tags.length > 0 && <div style={{ marginTop: 4 }}>{tags.map(t => <TagBadge key={t} tag={t} small />)}</div>}
                        </div>
                        <div style={{ textAlign: "center", minWidth: 64, flexShrink: 0, background: bgBadge, borderRadius: 10, padding: "8px 10px", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
                          <div style={{ fontSize: 26, fontWeight: 800, color: cor, lineHeight: 1 }}>{du_}</div>
                          <div style={{ fontSize: 9, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase", marginTop: 2 }}>d.u. pré-compra</div>
                          <div style={{ fontSize: 9, color: cor, fontWeight: 600, marginTop: 3 }}>SLA {PRAZO_SD} d.u.</div>
                        </div>
                      </div>
                    </div>
                  );
                })}
                <Pagination total={alertSDAberto.length} page={pgAlertSD} pageSize={pgAlertSDSz} onPage={setPgAlertSD} onSize={setPgAlertSDSz} />
              </>}

              {/* ══ SEÇÃO 2: PRÉ-COMPRA ENCERRADA — AGUARDANDO RC DA ÁREA ══ */}
              {alertSDEncerrado.length > 0 && <>
                <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: alertSDAberto.length > 0 ? 24 : 0, marginBottom: 10 }}>
                  <div style={{ width: 4, height: 20, background: "#f59e0b", borderRadius: 2, flexShrink: 0 }} />
                  <div style={{ fontSize: 14, fontWeight: 700, color: "var(--text)" }}>Pré-compra encerrada — Aguardando RC</div>
                  <span style={{ background: "#fef3c7", color: "#b45309", borderRadius: 20, padding: "2px 10px", fontSize: 11, fontWeight: 700 }}>{alertSDEncerrado.length}</span>
                  <div style={{ fontSize: 11, color: "var(--text3)" }}>Responsabilidade da Área Requisitante</div>
                </div>
                <div style={{ background: "#fffbeb", borderRadius: 10, padding: "10px 14px", marginBottom: 14, border: "1px solid #fde68a", fontSize: 11, color: "#92400e" }}>
                  <b>Cobrança à Área:</b> A pré-compra foi encerrada pelo GAQ. A Área Requisitante deve abrir a RC para dar continuidade ao processo de compras.
                </div>
                {alertSDEncerrado.slice((pgAlertSDEnc-1)*pgAlertSDEncSz, pgAlertSDEnc*pgAlertSDEncSz).map((r, i) => {
                  const duEnc = r.diasDesdeEncSD || 0;
                  const cor = duEnc > 7 ? "#dc2626" : duEnc > 3 ? "#ea580c" : "#b45309";
                  const borderCor = duEnc > 7 ? "#fca5a5" : duEnc > 3 ? "#fed7aa" : "#fde68a";
                  const bgBadge = duEnc > 7 ? "#fff1f2" : duEnc > 3 ? "#fff7ed" : "#fffbeb";
                  const tags = TagsManager.getForProcess(r.ProcessKey, r.TicketSD);
                  const watched = WatchlistManager.isWatched(r.ProcessKey);
                  const sdAlert = getServiceDeskAlertSummary(r, serviceDeskData);
                  return (
                    <div key={r.ProcessKey + i}
                      style={{ background: "var(--card)", borderRadius: 12, padding: "12px 16px", boxShadow: "0 1px 4px rgba(0,0,0,.06)", marginBottom: 8, border: `1px solid ${borderCor}`, borderLeft: `4px solid ${cor}`, cursor: "pointer", outline: watched ? "2px solid #f59e0b44" : "none" }}
                      onClick={() => setSelProc(r)}>
                      <div style={{ display: "flex", justifyContent: "space-between", gap: 12 }}>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center", marginBottom: 5 }}>
                            {watched && <span style={{ fontSize: 11, color: "#f59e0b" }}>⭐</span>}
                            <span style={{ fontWeight: 700, color: "var(--gaq-text)", fontSize: 13 }}>Pré-compra {r.TicketSD}</span>
                            <span style={{ fontSize: 10, background: "transparent", borderRadius: 6, padding: "1px 7px", color: "#b45309", fontWeight: 600, border: "1px solid #fde68a" }}>enc. sem RC</span>
                            {duEnc > 7 && <span style={{ fontSize: 10, background: "#dc2626", color: "#fff", borderRadius: 6, padding: "2px 9px", fontWeight: 700 }}>URGENTE {duEnc} d.u.</span>}
                            {r["Área Requisitante"] && <span style={{ fontSize: 10, background: "#eff6ff", color: "#2563eb", borderRadius: 6, padding: "1px 7px", fontWeight: 700 }}>{r["Área Requisitante"]}</span>}
                            {r.Modalidade && <span style={{ fontSize: 10, color: "var(--text3)" }}>{r.Modalidade}</span>}
                            <span onClick={(e) => { e.stopPropagation(); onOpenTags(r); }} style={{ fontSize: 10, cursor: "pointer", color: "var(--text3)", textDecoration: "underline" }}>+tag</span>
                            <span onClick={(e) => { e.stopPropagation(); toggleWatch(r); }} style={{ fontSize: 10, cursor: "pointer", color: watched ? "#f59e0b" : "var(--border2)", fontWeight: watched ? 700 : 400 }}>{watched ? "★" : "☆"}</span>
                          </div>
                          {r.statusDet && <div style={{ display: "inline-block", fontSize: 10, color: "#7c3aed", background: "#f5f3ff", borderRadius: 5, padding: "1px 7px", marginBottom: 4 }}>{r.statusDet}</div>}
                          <div style={{ fontSize: 12, color: "var(--text)", marginBottom: 4, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{(r.Objeto || "").slice(0, 100) || "—"}</div>
                          <ServiceDeskAlertStrip summary={sdAlert} showLastInteraction={!isServiceDeskConcluido(r)} />
                          <div style={{ display: "flex", gap: 5, flexWrap: "wrap", fontSize: 10, color: "var(--text3)" }}>
                            {r["Área Requisitante"] && <span>Área: <b style={{ color: "var(--text2)" }}>{r["Área Requisitante"]}</b></span>}
                            {r.aberturaSD && <span>· Abertura pré-compra: <b style={{ color: "var(--text2)" }}>{r.aberturaSD.toLocaleDateString("pt-BR")}</b></span>}
                            {r.encSD && <span>· Encerramento pré-compra: <b style={{ color: "#b45309" }}>{r.encSD.toLocaleDateString("pt-BR")}</b></span>}
                            {r.diasAgingSD > 0 && <span>· Aging pré-compra: <b style={{ color: "var(--text2)" }}>{r.diasAgingSD} d.u.</b></span>}
                          </div>
                          {tags.length > 0 && <div style={{ marginTop: 4 }}>{tags.map(t => <TagBadge key={t} tag={t} small />)}</div>}
                        </div>
                        <div style={{ textAlign: "center", minWidth: 64, flexShrink: 0, background: bgBadge, borderRadius: 10, padding: "8px 10px", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
                          <div style={{ fontSize: 26, fontWeight: 800, color: cor, lineHeight: 1 }}>{duEnc}</div>
                          <div style={{ fontSize: 9, color: "var(--text3)", fontWeight: 600, textTransform: "uppercase", marginTop: 2 }}>d.u. s/ RC</div>
                          <div style={{ fontSize: 9, color: "#b45309", fontWeight: 600, marginTop: 3 }}>Pend. Área</div>
                        </div>
                      </div>
                    </div>
                  );
                })}
                <Pagination total={alertSDEncerrado.length} page={pgAlertSDEnc} pageSize={pgAlertSDEncSz} onPage={setPgAlertSDEnc} onSize={setPgAlertSDEncSz} />
              </>}
            </>
          }
        </div>);
}
