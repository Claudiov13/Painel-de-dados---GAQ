import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function NovaScreen({ adicionarQuickTemp, alertMapDem, areas, base, comprIgnorados, compsCPLPool, compsNCLPool, contarCarteira, dem, gerarRecs, limparDemanda, limparTempDemandas, mods, quickTemp, recsCPL, recsNCL, redistForm, redistribuirCarteira, registrarTempDemanda, removerTempDemanda, renderBuscaPanel, setComprIgnorados, setDem, setQuickTemp, setRedistForm, setVerRankingCompleto, tempDemandas, toggleIgnorado, verRankingCompleto }) {
  return (<div style={{ maxWidth: 920, margin: "0 auto" }}>
          {renderBuscaPanel()}
          <div style={{ background: "var(--card)", borderRadius: 10, padding: 22, boxShadow: "var(--shadow)", marginBottom: 16 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
              <div style={{ fontWeight: 700, fontSize: 15, color: "#1a5276" }}>Dados da Nova Demanda</div>
              {(recsNCL.length > 0 || recsCPL.length > 0) && (
                <button onClick={limparDemanda}
                  style={{ background: "none", border: "1px solid #e74c3c", color: "#e74c3c", borderRadius: 6, padding: "5px 14px", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
                  ✕ Limpar e Nova Consulta
                </button>
              )}
            </div>
            <div className="grid-3 gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 2fr", gap: 10 }}>
              <div><label style={{ fontSize: 11, color: "var(--text3)" }}>Área Requisitante</label>
                <select value={dem.area} onChange={e => setDem({ ...dem, area: e.target.value })} style={{ width: "100%", padding: "7px", borderRadius: 6, border: "1px solid var(--input-bd)", marginTop: 3, fontSize: 12, background: "var(--input-bg)", color: "var(--text)" }}>
                  <option value="">Selecione...</option>{areas.map(a => <option key={a} value={a}>{a}</option>)}</select></div>
              <div>
                <label style={{ fontSize: 11, color: "var(--text3)" }}>Modalidade</label>
                <select value={dem.modalidade} onChange={e => setDem({ ...dem, modalidade: e.target.value })}
                  disabled={dem.ignorarMod}
                  style={{ width: "100%", padding: "7px", borderRadius: 6, border: "1px solid var(--input-bd)", marginTop: 3, fontSize: 12, background: "var(--input-bg)", color: "var(--text)", opacity: dem.ignorarMod ? 0.5 : 1 }}>
                  <option value="">Selecione...</option>{mods.map(m => <option key={m} value={m}>{m}</option>)}</select>
                <label style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 5, fontSize: 11, color: "var(--text3)", cursor: "pointer" }}>
                  <input type="checkbox" checked={!!dem.ignorarMod}
                    onChange={e => setDem({ ...dem, ignorarMod: e.target.checked, modalidade: e.target.checked ? "" : dem.modalidade })} />
                  Não sei a modalidade (ignorar indicador)
                </label>
              </div>
              <div><label style={{ fontSize: 11, color: "var(--text3)" }}>Objeto</label>
                <input value={dem.objeto} onChange={e => setDem({ ...dem, objeto: e.target.value })} placeholder="Ex: Aquisição de materiais de escritório..."
                  style={{ width: "100%", padding: "7px", borderRadius: 6, border: "1px solid var(--input-bd)", marginTop: 3, fontSize: 12, background: "var(--input-bg)", color: "var(--text)" }} /></div>
            </div>
            {/* Compradores indisponíveis (férias, atestado etc.) */}
            {(compsNCLPool.length > 0 || compsCPLPool.length > 0) && (
              <div style={{ marginTop: 14, padding: "12px 14px", background: "var(--card2)", border: "1px solid var(--input-bd)", borderRadius: 8 }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8, gap: 8, flexWrap: "wrap" }}>
                  <div>
                    <div style={{ fontSize: 12, fontWeight: 700, color: "var(--text)" }}>Indisponíveis (férias, atestado etc.)</div>
                    <div style={{ fontSize: 10.5, color: "var(--text3)", marginTop: 2 }}>
                      Clique para excluir alguém da sugestão. {comprIgnorados.size > 0 && <b style={{ color: "#c0392b" }}>{comprIgnorados.size} pessoa(s) marcada(s)</b>}
                    </div>
                  </div>
                  {comprIgnorados.size > 0 && (
                    <button onClick={() => setComprIgnorados(new Set())}
                      style={{ background: "none", border: "1px solid var(--input-bd)", color: "var(--text2)", borderRadius: 6, padding: "4px 10px", cursor: "pointer", fontSize: 11, fontWeight: 600 }}>
                      Limpar marcações
                    </button>
                  )}
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                  {[...new Set([...compsNCLPool, ...compsCPLPool])].sort((a, b) => a.localeCompare(b, "pt-BR")).map(nome => {
                    const ign = comprIgnorados.has(nome);
                    return (
                      <button key={nome} onClick={() => toggleIgnorado(nome)}
                        title={ign ? "Clique para reincluir" : "Clique para marcar como indisponível"}
                        style={{
                          background: ign ? "#fde2e2" : "var(--card)",
                          color: ign ? "#c0392b" : "var(--text2)",
                          border: ign ? "1px solid #c0392b" : "1px solid var(--input-bd)",
                          borderRadius: 999, padding: "4px 12px", fontSize: 11, fontWeight: 600,
                          cursor: "pointer", textDecoration: ign ? "line-through" : "none",
                        }}>
                        {ign ? "✕ " : ""}{nome}
                      </button>
                    );
                  })}
                </div>
              </div>
            )}
            {/* Redistribuir carteira de quem está de férias/indisponível (opcional) */}
            {comprIgnorados.size > 0 && (() => {
              const inp = { padding: "6px 7px", borderRadius: 6, border: "1px solid var(--input-bd)", fontSize: 11.5, background: "var(--input-bg)", color: "var(--text)" };
              const ignorados = [...comprIgnorados].sort((a, b) => a.localeCompare(b, "pt-BR"));
              const setRF = (key, patch) => setRedistForm(prev => ({ ...prev, [key]: { ...(prev[key] || {}), ...patch } }));
              const linhas = [];
              ignorados.forEach(nome => {
                const c = contarCarteira(nome);
                [["ncl", c.ncl, compsNCLPool], ["cpl", c.cpl, compsCPLPool]].forEach(([pool, qtd, poolMembers]) => {
                  if (qtd <= 0) return;
                  const key = `${nome}|${pool}`;
                  const f = redistForm[key] || {};
                  const receivers = poolMembers.filter(p => p !== nome && !comprIgnorados.has(p)).sort((a, b) => a.localeCompare(b, "pt-BR"));
                  linhas.push(
                    <div key={key} style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center", padding: "8px 10px", background: "var(--card)", borderRadius: 7, border: "1px solid var(--input-bd)" }}>
                      <div style={{ flex: "1 1 180px", fontSize: 12, color: "var(--text)" }}>
                        <b>{nome}</b> <span style={{ color: pool === "cpl" ? "#8e44ad" : "#1a5276", fontSize: 10, fontWeight: 700 }}>{pool.toUpperCase()}</span>
                        <span style={{ color: "var(--text3)", marginLeft: 6 }}>{qtd} ativo(s)</span>
                      </div>
                      <span style={{ fontSize: 11, color: "var(--text3)" }}>→ para</span>
                      <select value={f.receiver || ""} onChange={e => setRF(key, { receiver: e.target.value })} style={{ ...inp, flex: "1 1 150px" }}>
                        <option value="">Selecione quem assume...</option>
                        {receivers.map(p => <option key={p} value={p}>{p}</option>)}
                      </select>
                      <label style={{ fontSize: 10.5, color: "var(--text3)" }}>Qtd</label>
                      <input type="number" min={1} max={99} value={f.qty != null ? f.qty : qtd}
                        onChange={e => setRF(key, { qty: e.target.value })} style={{ ...inp, width: 60 }} />
                      <button type="button" onClick={() => redistribuirCarteira(nome, pool)} disabled={!f.receiver}
                        style={{ background: f.receiver ? "#d35400" : "#bbb", color: "#fff", border: "none", borderRadius: 6, padding: "6px 12px", fontSize: 11.5, fontWeight: 700, cursor: f.receiver ? "pointer" : "not-allowed", whiteSpace: "nowrap" }}>
                        Redistribuir
                      </button>
                    </div>
                  );
                });
              });
              return (
                <div style={{ marginTop: 12, padding: "12px 14px", background: "var(--card2)", border: "1px solid #e8b88a", borderRadius: 8 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#d35400" }}>Redistribuir carteira de indisponíveis (temporário)</div>
                  <div style={{ fontSize: 10.5, color: "var(--text3)", margin: "2px 0 10px" }}>
                    Opcional: transfira a carteira ativa de quem está de férias para outra pessoa. Cria demandas temporárias na conta de quem assume e não altera os dados (JSON).
                  </div>
                  {linhas.length > 0
                    ? <div style={{ display: "grid", gap: 6 }}>{linhas}</div>
                    : <div style={{ fontSize: 11.5, color: "var(--text3)" }}>As pessoas marcadas não têm carteira ativa para redistribuir.</div>}
                </div>
              );
            })()}
            <button onClick={() => gerarRecs()} disabled={base.length === 0}
              style={{ marginTop: 14, background: base.length === 0 ? "#aaa" : "#1a5276", color: "#fff", padding: "9px 26px", borderRadius: 8, border: "none", cursor: base.length === 0 ? "not-allowed" : "pointer", fontWeight: 700, fontSize: 13 }}>
              Gerar Recomendações</button>
          </div>
          {/* Distribuição rápida temporária — só a pessoa é obrigatória */}
          {base.length > 0 && (() => {
            const lbl = { fontSize: 10.5, color: "var(--text3)", display: "block", marginBottom: 3 };
            const inp = { width: "100%", padding: "7px", borderRadius: 6, border: "1px solid var(--input-bd)", fontSize: 12, background: "var(--input-bg)", color: "var(--text)" };
            const pessoas = [...new Set([...compsNCLPool, ...compsCPLPool])].sort((a, b) => a.localeCompare(b, "pt-BR"));
            return (
              <div style={{ background: "var(--card)", borderRadius: 10, padding: 18, boxShadow: "var(--shadow)", marginBottom: 14, borderLeft: "4px solid #16a085" }}>
                <div style={{ fontSize: 13, fontWeight: 800, color: "#117a65" }}>Distribuição rápida (temporária)</div>
                <div style={{ fontSize: 10.5, color: "var(--text3)", margin: "2px 0 12px" }}>
                  Atribua uma demanda direto a alguém — só a <b>pessoa</b> é obrigatória; área, modalidade e objeto são opcionais. Não altera os dados (JSON): vale só nesta sessão e impacta as próximas sugestões.
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "flex-end" }}>
                  <div style={{ flex: "2 1 190px" }}>
                    <label style={lbl}>Pessoa *</label>
                    <select value={quickTemp.responsavel}
                      onChange={e => { const v = e.target.value; setQuickTemp(q => ({ ...q, responsavel: v, pool: compsCPLPool.includes(v) && !compsNCLPool.includes(v) ? "cpl" : "ncl" })); }}
                      style={inp}>
                      <option value="">Selecione a pessoa...</option>
                      {pessoas.map(n => <option key={n} value={n}>{n}</option>)}
                    </select>
                  </div>
                  <div style={{ flex: "0 1 100px" }}>
                    <label style={lbl}>Pool</label>
                    <select value={quickTemp.pool || "ncl"} onChange={e => setQuickTemp(q => ({ ...q, pool: e.target.value }))} style={inp}>
                      <option value="ncl">NCL</option>
                      <option value="cpl">CPL</option>
                    </select>
                  </div>
                  <div style={{ flex: "0 1 80px" }}>
                    <label style={lbl}>Qtd</label>
                    <input type="number" min={1} max={99} value={quickTemp.quantidade}
                      onChange={e => setQuickTemp(q => ({ ...q, quantidade: e.target.value }))} style={inp} />
                  </div>
                  <div style={{ flex: "1 1 150px" }}>
                    <label style={lbl}>Área (opcional)</label>
                    <select value={quickTemp.area} onChange={e => setQuickTemp(q => ({ ...q, area: e.target.value }))} style={inp}>
                      <option value="">—</option>{areas.map(a => <option key={a} value={a}>{a}</option>)}
                    </select>
                  </div>
                  <div style={{ flex: "1 1 150px" }}>
                    <label style={lbl}>Modalidade (opcional)</label>
                    <select value={quickTemp.modalidade} onChange={e => setQuickTemp(q => ({ ...q, modalidade: e.target.value }))} style={inp}>
                      <option value="">—</option>{mods.map(m => <option key={m} value={m}>{m}</option>)}
                    </select>
                  </div>
                  <div style={{ flex: "2 1 180px" }}>
                    <label style={lbl}>Objeto (opcional)</label>
                    <input value={quickTemp.objeto} onChange={e => setQuickTemp(q => ({ ...q, objeto: e.target.value }))} placeholder="Demanda temporária" style={inp} />
                  </div>
                  <button type="button" onClick={adicionarQuickTemp} disabled={!quickTemp.responsavel}
                    style={{ background: quickTemp.responsavel ? "#16a085" : "#aaa", color: "#fff", border: "none", borderRadius: 8, padding: "9px 18px", fontWeight: 700, fontSize: 12.5, cursor: quickTemp.responsavel ? "pointer" : "not-allowed", whiteSpace: "nowrap" }}>
                    + Adicionar
                  </button>
                </div>
              </div>
            );
          })()}
          {tempDemandas.length > 0 && <div style={{ background: "var(--card)", borderRadius: 10, padding: 14, boxShadow: "var(--shadow)", marginBottom: 14, borderLeft: "4px solid #d35400" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, marginBottom: 10, flexWrap: "wrap" }}>
              <div>
                <div style={{ fontSize: 13, fontWeight: 800, color: "#d35400" }}>Demandas temporárias da sessão</div>
                <div style={{ fontSize: 10.5, color: "var(--text3)", marginTop: 2 }}>{tempDemandas.length} ajuste(s) impactando as próximas sugestões</div>
              </div>
              <button type="button" onClick={limparTempDemandas}
                style={{ background: "transparent", border: "1px solid #d35400", color: "#d35400", borderRadius: 6, padding: "4px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
                Limpar temporárias
              </button>
            </div>
            <div style={{ display: "grid", gap: 6 }}>
              {tempDemandas.map(t => (
                <div key={t.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, padding: "7px 10px", background: "var(--card2)", borderRadius: 7 }}>
                  <div style={{ minWidth: 0 }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: "var(--text)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                      {t.responsavel} <span style={{ color: t.pool === "cpl" ? "#8e44ad" : "#1a5276", fontSize: 10, textTransform: "uppercase" }}>{t.pool}</span>
                      {t.origem && <span style={{ marginLeft: 6, background: "#d3540018", color: "#d35400", borderRadius: 4, padding: "1px 6px", fontSize: 9.5, fontWeight: 700, textTransform: "none" }}>férias de {t.origem}</span>}
                    </div>
                    <div style={{ fontSize: 10.5, color: "var(--text3)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                      {t.area || "Área n/i"} · {t.modalidade || "Modalidade n/i"} · {(t.objeto || "Demanda temporária").slice(0, 110)}
                    </div>
                  </div>
                  <button type="button" onClick={() => removerTempDemanda(t.id)}
                    style={{ background: "transparent", border: "none", color: "#c0392b", fontSize: 12, fontWeight: 800, cursor: "pointer", padding: "3px 6px" }}>
                    Remover
                  </button>
                </div>
              ))}
            </div>
          </div>}
          {alertMapDem.length > 0 && <div style={{ background: "#fef9e7", border: "2px solid #e67e22", borderRadius: 10, padding: 14, marginBottom: 14 }}>
            <div style={{ fontWeight: 700, color: "#d35400", marginBottom: 6, fontSize: 13 }}>Alerta: Objeto com histórico problemático</div>
            {alertMapDem.map(p => (
              <div key={p.word} style={{ background: "var(--card)", borderRadius: 6, padding: "8px 12px", marginBottom: 6, border: "1px solid #f0c080" }}>
                <span style={{ background: "#e67e22", color: "#fff", borderRadius: 4, padding: "1px 8px", fontSize: 11, fontWeight: 700, marginRight: 8 }}>"{p.word}"</span>
                <span style={{ fontSize: 12, color: "var(--text2)" }}>{p.count} casos: {p.cancelados} cancelados · {p.fracassados} fracassados{p.avgLT ? ` · Lead médio: ${p.avgLT} d.u.` : ""}</span>
              </div>
            ))}
          </div>}
          {(recsNCL.length > 0 || recsCPL.length > 0) && (
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, padding: "10px 14px", background: "var(--card2)", borderRadius: 8, flexWrap: "wrap", gap: 8 }}>
              <div style={{ fontSize: 11.5, color: "var(--text3)" }}>
                <b style={{ color: "var(--text)" }}>{recsNCL.length}</b> NCL considerado(s) · <b style={{ color: "var(--text)" }}>{recsCPL.length}</b> CPL considerado(s)
                {comprIgnorados.size > 0 && <> · <b style={{ color: "#c0392b" }}>{comprIgnorados.size} ignorado(s)</b></>}
                {tempDemandas.length > 0 && <> · <b style={{ color: "#d35400" }}>{tempDemandas.length} temporária(s)</b></>}
              </div>
              <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 11.5, color: "var(--text2)", cursor: "pointer", fontWeight: 600 }}>
                <input type="checkbox" checked={verRankingCompleto} onChange={e => setVerRankingCompleto(e.target.checked)} />
                Ver ranking completo (todos os candidatos)
              </label>
            </div>
          )}
          {/* NCL recommendations */}
          {recsNCL.length > 0 && <div style={{ marginBottom: 20 }}>
            <div style={{ fontWeight: 700, fontSize: 14, color: "#1a5276", marginBottom: 10, padding: "8px 14px", background: "#1a527611", borderRadius: 8, border: "2px solid #1a5276" }}>
              {verRankingCompleto ? `Ranking NCL — ${recsNCL.length} candidato(s)` : "Top 3 Compradores NCL Recomendados"}
              <div style={{ fontSize: 10, fontWeight: 400, color: "var(--text3)", marginTop: 2 }}>Núcleo de Compras e Licitações — pool exclusivo NCL</div>
            </div>
            {(verRankingCompleto ? recsNCL : recsNCL.slice(0, 3)).map((r, i) => <RecCard key={r.comp} r={r} i={i} pool="ncl" dem={dem} onTempAssign={registrarTempDemanda} />)}
          </div>}
          {/* CPL recommendations */}
          {recsCPL.length > 0 && <div>
            <div style={{ fontWeight: 700, fontSize: 14, color: "#8e44ad", marginBottom: 10, padding: "8px 14px", background: "#8e44ad11", borderRadius: 8, border: "2px solid #8e44ad" }}>
              {verRankingCompleto ? `Ranking CPL — ${recsCPL.length} candidato(s)` : "Top 3 Responsáveis CPL Recomendados"}
              <div style={{ fontSize: 10, fontWeight: 400, color: "var(--text3)", marginTop: 2 }}>Comissão Permanente de Licitações — pool exclusivo CPL</div>
            </div>
            {(verRankingCompleto ? recsCPL : recsCPL.slice(0, 3)).map((r, i) => <RecCard key={r.comp} r={r} i={i} pool="cpl" dem={dem} onTempAssign={registrarTempDemanda} />)}
          </div>}
        </div>);
}
