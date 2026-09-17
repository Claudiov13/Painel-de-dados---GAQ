import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function OperacionalScreen({ base, distAnosFilter, fArea, fBack, fBase, loginArea, loginAreaDisp, meta, phaseIntervals, setAba, setDistAnosFilter, setDrillDown, setSelProc, setShowEmailDiario, setShowFunilInfo, setShowGaqSextaEditor, showFunilInfo }) {
  return (base.length > 0 ? <div className="anim-fade">
          <div style={{ background: "var(--gaq-surface)", borderRadius: "var(--gaq-r-xl)", padding: "24px 28px", marginBottom: 18, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, border: "1px solid var(--gaq-line)", boxShadow: "var(--gaq-shadow-2)" }}>
            <div>
              <div style={{ fontSize: 22, fontWeight: 700, color: "var(--gaq-text)", letterSpacing: "-0.02em" }}>Painel Operacional</div>
              <div style={{ fontSize: 12, color: "var(--gaq-text-3)", marginTop: 2 }}>Aging, fases e gargalos da carteira activa · {fBase.filter(r => r.emA).length} processos em andamento</div>
            </div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <button className="gaq-btn no-print" onClick={() => {
                const tituloArea = fArea || "Todas as Áreas";
                const html = gerarRelatorioOperacional({ processos: fBack, area: tituloArea, meta });
                const win = window.open("", "_blank");
                win.document.write(html);
                win.document.close();
              }} title="Gerar PDF com os filtros ativos">
                <Icon name="doc" size={14}/> PDF Operacional
              </button>
              <button className="gaq-btn no-print" onClick={() => setShowGaqSextaEditor(true)} title="Revisar status/observações e gerar o relatório GAQ do filtro atual">
                <Icon name="doc" size={14}/> GAQ Sexta
              </button>
              <button className="gaq-btn no-print" onClick={() => {
                const area = fArea || "Todas";
                const hoje = new Date();
                const dataFmt = hoje.toLocaleDateString("pt-BR");
                const diasSemana = ["Domingo","Segunda","Terça","Quarta","Quinta","Sexta","Sábado"];
                const diaNome = diasSemana[hoje.getDay()];
                const emA = fBase.filter(r => r.emA);
                const totalEmA = emA.length;
                const criticos = emA.filter(r => {
                  const sla = calcSLAClassification(r, phaseIntervals);
                  return sla && sla.bucket === "critico";
                }).length;
                const parados = emA.filter(r => r.diasParado > 15).length;
                const html = gerarRelatorioGAQSexta({ processos: fBase, area, meta, fBase });
                const win = window.open("", "_blank");
                win.document.write(html);
                win.document.close();
                const subject = `Relatório GAQ — ${area} — ${dataFmt} (${diaNome})`;
                const body =
`Prezados,

Segue em anexo o Relatório GAQ referente à ${diaNome}-feira, ${dataFmt}.

Área filtrada: ${area}
Processos em andamento: ${totalEmA}
Críticos (SLA + entrega): ${criticos}
Estagnados (>15 d.u. sem movimentação): ${parados}

O relatório completo está em anexo (PDF GAQ Sexta).

Em caso de dúvidas, fico à disposição.

Atenciosamente,
GAQ — Gerência de Aquisições`;
                setTimeout(() => {
                  const a = document.createElement("a");
                  a.href = "mailto:?subject=" + encodeURIComponent(subject) + "&body=" + encodeURIComponent(body);
                  a.click();
                }, 800);
              }} title="Gera o PDF GAQ Sexta e abre o Outlook com e-mail pré-preenchido">
                <Icon name="external" size={14}/> Outlook
              </button>
              <button className="gaq-btn no-print" onClick={() => {
                const html = gerarRelatorioGAQSextaPorAreas({ base, meta });
                const win = window.open("", "_blank");
                win.document.write(html);
                win.document.close();
              }} title="Gerar um PDF separado por área">
                <Icon name="download" size={14}/> PDF por Áreas
              </button>
              <button className="gaq-btn no-print" onClick={() => setShowEmailDiario(true)} title="Gerar e-mail diário para o Marcio com SD e RCs de hoje">
                <Icon name="external" size={14}/> Email Diário Marcio
              </button>
            </div>
          </div>

          {/* KPIs Operacionais */}
          {(() => {
            const emAF = fBase.filter(r => r.emA);
            const concF = fBase.filter(r => r.isConcluded);
            const totalF = fBase.length;
            const mediaAging = emAF.length ? Math.round(emAF.reduce((s,r) => s + r.diasTotais, 0) / emAF.length) : 0;
            const semPedido = emAF.filter(r => r.semEnvio).length;
            const opRowKey = (r) => r.ProcessKey || `${r.NumRC || ""}|${r.TicketSD || ""}|${r.Objeto || ""}`;
            const opClassificados = emAF.map(r => {
              const sla = calcSLAClassification(r, phaseIntervals);
              return sla ? { ...r, _sla: sla } : null;
            }).filter(Boolean);
            const opSlaMap = new Map(opClassificados.map(r => [opRowKey(r), r._sla]));
            const getOpSla = (r) => opSlaMap.get(opRowKey(r)) || calcSLAClassification(r, phaseIntervals);
            const isOpCritico = (r) => getOpSla(r)?.bucket === "critico";
            const isOpAtencao = (r) => getOpSla(r)?.bucket === "atencao";
            const isOpSlaVencido = (r) => getOpSla(r)?.bucket === "sla_vencido";
            const opsCriticos = opClassificados.filter(r => r._sla.bucket === "critico");
            const opsAtencao = opClassificados.filter(r => r._sla.bucket === "atencao");
            const opsSlaVencido = opClassificados.filter(r => r._sla.bucket === "sla_vencido");
            const opsSlaAtendido = opClassificados.filter(r => r._sla.bucket === "no_prazo");
            const opTone = (r) => {
              const bucket = getOpSla(r)?.bucket;
              return bucket === "critico"
                ? "#d35400"
                : bucket === "atencao"
                  ? "#ff9500"
                  : bucket === "sla_vencido"
                    ? "#ff3b30"
                    : "#34c759";
            };

            // Status detalhado breakdown
            const statusDetMap = {};
            emAF.forEach(r => { const s = (r.statusDet || r.status || "Sem status").trim(); statusDetMap[s] = (statusDetMap[s] || 0) + 1; });
            const statusDetArr = Object.entries(statusDetMap).sort((a,b) => b[1] - a[1]);

            // Por modalidade (em andamento filtrado)
            const modMap = {};
            emAF.forEach(r => { const m = r.Modalidade || "N/I"; if (!modMap[m]) modMap[m] = { total: 0, crit: 0, sumDias: 0 }; modMap[m].total++; if (isOpCritico(r)) modMap[m].crit++; modMap[m].sumDias += r.diasTotais; });
            const modArr = Object.entries(modMap).map(([name, v]) => ({ name, ...v, media: Math.round(v.sumDias / v.total) })).sort((a,b) => b.total - a.total);

            // Por comprador (em andamento filtrado) — NCL: respNCL (Comprador||Avaliador), CPL: Pregoeiro
            const compMap = {};
            emAF.forEach(r => { const c = r.ehCPL ? (r.Pregoeiro || r.cplResp || "N/I") : (r.respNCL || "N/I"); if (!compMap[c]) compMap[c] = { total: 0, crit: 0, sumDias: 0 }; compMap[c].total++; if (isOpCritico(r)) compMap[c].crit++; compMap[c].sumDias += r.diasTotais; });
            const compArr = Object.entries(compMap).map(([name, v]) => ({ name, ...v, media: Math.round(v.sumDias / v.total) })).sort((a,b) => b.crit - a.crit || b.total - a.total);

            // Funil expandido — filtros explícitos por data para NCL, CPL e Contratos
            // Processos CPL (para fases CPL)
            const emAF_CPL = emAF.filter(r => r.ehCPL || pd(r[COL_CPL_RECEBIDO_NCL]) || pd(r[COL_CPL_ENVIADO_DJS_CHANCELA]) || pd(r["CPL_ABERTURA_DISPUTA_FINAL"]) || pd(r["CPL_DATA_HOMOLOGACAO_FINAL"]));
            const fasesExpandidas = [
              // ── NCL ──
              { label: "Sem pré-compra/RC",      area: "NCL", color: "#95a5a6", grad: "linear-gradient(90deg, #95a5a6, #bdc3c7)",
                procs: emAF.filter(r => !r.aberturaSD && !r.aberturaRC) },
              { label: "Pré-compra em distribuição", area: "NCL", color: "#2e86c1", grad: "linear-gradient(90deg, #2e86c1, #85c1e9)",
                procs: emAF.filter(r => r.aberturaSD && !pd(r["Data da distribuição do SD"])) },
              { label: "Pré-compra distribuída", area: "NCL", color: "#2e86c1", grad: "linear-gradient(90deg, #2e86c1, #85c1e9)",
                procs: emAF.filter(r => pd(r["Data da distribuição do SD"]) && !pd(r["Data do encerramento"])) },
              { label: "Pré-compra encerrada",     area: "NCL", color: "#2e86c1", grad: "linear-gradient(90deg, #2e86c1, #85c1e9)",
                procs: emAF.filter(r => pd(r["Data do encerramento"]) && !pd(r["DATA DO RECEBIMENTO DA RC"])) },
              { label: "RC Recebida",              area: "NCL", color: "#2e86c1", grad: "linear-gradient(90deg, #2e86c1, #85c1e9)",
                procs: emAF.filter(r => pd(r["DATA DO RECEBIMENTO DA RC"]) && !pd(r["Data inicial do envio de propostas"]) && !pd(r["Data do envio para aprovação"]) && !pd(r[COL_CPL_RECEBIDO_NCL])) },
              { label: "Em Cotação",               area: "NCL", color: "#2e86c1", grad: "linear-gradient(90deg, #2e86c1, #85c1e9)",
                procs: emAF.filter(r => pd(r["Data inicial do envio de propostas"]) && !pd(r["Data final do envio de propostas"])) },
              { label: "Cotação Recebida",         area: "NCL", color: "#2e86c1", grad: "linear-gradient(90deg, #2e86c1, #85c1e9)",
                procs: emAF.filter(r => pd(r["Data final do envio de propostas"]) && !pd(r["Data do envio para aprovação"])) },
              { label: "Em Aprovação",             area: "NCL", color: "#2e86c1", grad: "linear-gradient(90deg, #2e86c1, #85c1e9)",
                procs: emAF.filter(r => pd(r["Data do envio para aprovação"]) && !pd(r["Data da última aprovação"])) },
              { label: "Última Aprovação",         area: "NCL", color: "#2e86c1", grad: "linear-gradient(90deg, #2e86c1, #85c1e9)",
                procs: emAF.filter(r => pd(r["Data da última aprovação"]) && r.semEnvio && !pd(r[COL_CPL_RECEBIDO_NCL])) },
              { label: "Envio Pedido/Suite",        area: "NCL", color: "#2e86c1", grad: "linear-gradient(90deg, #2e86c1, #85c1e9)",
                procs: emAF.filter(r => pd(r["DATA DE ENVIO DO PEDIDO (fornecedor) OU SUITE SESC (SCONT ou CPL)"]) && !pd(r["DATA DE RECEBIMENTO DA DEMANDA PELA ANALISTA DE CONTRATO"]) && !pd(r[COL_CPL_RECEBIDO_NCL])) },
              // ── CPL ──
              { label: "CPL: Recebido do NCL",     area: "CPL", color: "#8e44ad", grad: "linear-gradient(90deg, #8e44ad, #c39bd3)",
                procs: emAF_CPL.filter(r => pd(r[COL_CPL_RECEBIDO_NCL]) && !pd(r[COL_CPL_ENVIADO_DJS_CHANCELA])) },
              { label: "CPL: Enviado p/ DJS",      area: "CPL", color: "#8e44ad", grad: "linear-gradient(90deg, #8e44ad, #c39bd3)",
                procs: emAF_CPL.filter(r => pd(r[COL_CPL_ENVIADO_DJS_CHANCELA]) && !pd(r["CPL_DATA_RECEBIMENTO_FINAL"])) },
              { label: "CPL: Recebimento",         area: "CPL", color: "#8e44ad", grad: "linear-gradient(90deg, #8e44ad, #c39bd3)",
                procs: emAF_CPL.filter(r => pd(r["CPL_DATA_RECEBIMENTO_FINAL"]) && !pd(r["CPL_PUBLICACAO_AVISO_FINAL"])) },
              { label: "CPL: Publicação",           area: "CPL", color: "#8e44ad", grad: "linear-gradient(90deg, #8e44ad, #c39bd3)",
                procs: emAF_CPL.filter(r => pd(r["CPL_PUBLICACAO_AVISO_FINAL"]) && !pd(r["CPL_ABERTURA_DISPUTA_FINAL"])) },
              { label: "CPL: Abertura Disputa",    area: "CPL", color: "#8e44ad", grad: "linear-gradient(90deg, #8e44ad, #c39bd3)",
                procs: emAF_CPL.filter(r => pd(r["CPL_ABERTURA_DISPUTA_FINAL"]) && !pd(r["CPL_FINALIZACAO_FASE_EXTERNA_FINAL"])) },
              { label: "CPL: Fase ext.",            area: "CPL", color: "#8e44ad", grad: "linear-gradient(90deg, #8e44ad, #c39bd3)",
                procs: emAF_CPL.filter(r => pd(r["CPL_FINALIZACAO_FASE_EXTERNA_FINAL"]) && !pd(r["CPL_DATA_HOMOLOGACAO_FINAL"])) },
              { label: "CPL: Homologação",          area: "CPL", color: "#8e44ad", grad: "linear-gradient(90deg, #8e44ad, #c39bd3)",
                procs: emAF_CPL.filter(r => pd(r["CPL_DATA_HOMOLOGACAO_FINAL"]) && !pd(r["DATA DE RECEBIMENTO DA DEMANDA PELA ANALISTA DE CONTRATO"])) },
              // ── Contratos (SCONT / DJ) ──
              { label: "Indicação Analista",       area: "Contratos", color: "#16a085", grad: "linear-gradient(90deg, #16a085, #76d7c4)",
                procs: emAF.filter(r => pd(r["DATA DE RECEBIMENTO DA DEMANDA PELA ANALISTA DE CONTRATO"]) && !pd(r["DATA DE RECEBIMENTO NA DJ (QUANDO APLICÁVEL)"])) },
              { label: "Recebimento DJ",            area: "Contratos", color: "#16a085", grad: "linear-gradient(90deg, #16a085, #76d7c4)",
                procs: emAF.filter(r => pd(r["DATA DE RECEBIMENTO NA DJ (QUANDO APLICÁVEL)"])) },
            ].map(f => ({ ...f, count: f.procs.length }));
            // Totais por grupo (para decidir se mostra o grupo)
            const totalPorGrupo = {};
            fasesExpandidas.forEach(f => { totalPorGrupo[f.area] = (totalPorGrupo[f.area] || 0) + f.count; });
            const maxFase = Math.max(...fasesExpandidas.map(f => f.count), 1);

            return <>
            {/* Funil de Fases Expandido — NCL + CPL + Contratos */}
            <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 16 }}>
              <div className="gaq-card" style={{ padding: 22, borderRadius: 24, position: "relative" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
                  <div>
                    <div style={{ fontWeight: 700, fontSize: 15, color: "var(--gaq-text)" }}>Funil de fases</div>
                    <div style={{ fontSize: 11, color: "var(--gaq-text-3)", marginTop: 2 }}>Onde cada processo está parado agora</div>
                  </div>
                  <button onClick={() => setShowFunilInfo(!showFunilInfo)}
                    style={{ background: showFunilInfo ? "#2e86c1" : "var(--border2)", border: "none", borderRadius: "50%", width: 22, height: 22, cursor: "pointer", fontSize: 12, fontWeight: 800, color: showFunilInfo ? "#fff" : "var(--text3)", flexShrink: 0, lineHeight: "22px", textAlign: "center" }}
                    title="Ver legenda das subáreas">?</button>
                </div>
                {showFunilInfo && (
                  <div style={{ background: "var(--bg)", border: "1px solid var(--border2)", borderRadius: 8, padding: 12, marginBottom: 12, fontSize: 10, lineHeight: 1.6, color: "var(--text2)" }}>
                    <div style={{ fontWeight: 700, fontSize: 11, marginBottom: 6, color: "var(--text)" }}>Legenda das subáreas</div>
                    <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
                      <span><span style={{ color: "#2e86c1", fontWeight: 700 }}>■</span> NCL — Núcleo de Compras e Licitações</span>
                      <span><span style={{ color: "#8e44ad", fontWeight: 700 }}>■</span> CPL — Comissão Permanente de Licitação</span>
                      <span><span style={{ color: "#16a085", fontWeight: 700 }}>■</span> Contratos — SCONT / DJ</span>
                    </div>
                    <div style={{ marginTop: 6, fontSize: 9, color: "var(--text3)", fontStyle: "italic" }}>Cada fase mostra quantos processos ativos estão naquela etapa. Clique para ver a lista.</div>
                  </div>
                )}
                {(() => {
                  let lastArea = "";
                  return fasesExpandidas.map((f, i) => {
                    const pct = maxFase > 0 ? Math.round((f.count / maxFase) * 100) : 0;
                    const showHeader = f.area !== lastArea;
                    const grupoTemProcessos = (totalPorGrupo[f.area] || 0) > 0;
                    lastArea = f.area;
                    // Esconder grupo inteiro só se não tem nenhum processo (ex: CPL quando não há pregões)
                    if (!grupoTemProcessos && f.area !== "NCL") return null;
                    return (
                      <React.Fragment key={f.label}>
                        {showHeader && (
                          <div style={{ fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".06em", color: f.color, marginTop: i > 0 ? 10 : 0, marginBottom: 4, paddingLeft: 148, opacity: .8 }}>
                            {f.area} {f.area !== "NCL" && <span style={{ fontWeight: 400, opacity: .7 }}>({totalPorGrupo[f.area] || 0})</span>}
                          </div>
                        )}
                        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4, cursor: f.count > 0 ? "pointer" : "default", borderRadius: 6, padding: "3px 4px", transition: "background .15s", opacity: f.count === 0 ? .45 : 1 }}
                          onMouseEnter={e => { if (f.count > 0) e.currentTarget.style.background = "var(--hover)"; }} onMouseLeave={e => e.currentTarget.style.background = "transparent"}
                          onClick={() => { if (f.count > 0) { setDrillDown({ title: "Fase: " + f.label, data: f.procs, color: f.color, sourceAba: "operacional" }); setAba("overview"); } }}>
                          <div style={{ fontSize: 10, color: "var(--text2)", width: 140, flexShrink: 0, textAlign: "right", lineHeight: 1.2 }}>{f.label}</div>
                          <div style={{ flex: 1, background: "var(--border2)", borderRadius: 4, height: 18, overflow: "hidden" }}>
                            {f.count > 0 && <div style={{ width: Math.max(pct, 3) + "%", height: "100%", background: f.grad, borderRadius: 4, transition: "width .5s" }} />}
                          </div>
                          <div style={{ fontSize: 12, fontWeight: 700, color: f.count > 0 ? f.color : "var(--text3)", width: 32, textAlign: "right" }}>{f.count}</div>
                        </div>
                      </React.Fragment>
                    );
                  });
                })()}
                {emAF.length === 0 && <div style={{ color: "var(--text3)", fontSize: 12, textAlign: "center", padding: 20 }}>Nenhum processo em andamento.</div>}
              </div>

              {/* Estagnados */}
              <div className="gaq-card" style={{ padding: 22, borderRadius: 24, borderTop: "3px solid #e67e22" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
                  <div>
                    <div style={{ fontWeight: 700, fontSize: 15, color: "var(--gaq-text)" }}>Estagnados</div>
                    <div style={{ fontSize: 11, color: "var(--gaq-text-3)", marginTop: 2 }}>{emAF.filter(r => r.diasParado > 15).length} processos sem movimentação ha mais de 15 d.u.</div>
                  </div>
                  {emAF.filter(r => r.diasParado > 15).length > 8 && (
                    <button className="gaq-btn" style={{ fontSize: 11 }}
                      onClick={() => { setDrillDown({ title: "Processos estagnados (>15 d.u. sem movimentação)", data: emAF.filter(r => r.diasParado > 15).sort((a,b) => b.diasParado - a.diasParado), color: "#e67e22", sourceAba: "operacional" }); setAba("overview"); }}>
                      Ver todos
                    </button>
                  )}
                </div>
                {(() => {
                  const parados = [...emAF].filter(r => r.diasParado > 15).sort((a,b) => b.diasParado - a.diasParado).slice(0, 8);
                  if (parados.length === 0) return <div style={{ color: "var(--text3)", fontSize: 12, textAlign: "center", padding: 12 }}>Nenhum processo estagnado.</div>;
                  return parados.map(r => (
                    <div key={r.ProcessKey} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 0", borderBottom: "1px solid var(--gaq-line)", cursor: "pointer", fontSize: 11 }}
                      onClick={() => setSelProc(r)}>
                      <div style={{ width: 6, height: 6, borderRadius: "50%", background: r.diasParado > 30 ? "#c0392b" : "#e67e22", flexShrink: 0 }} />
                      <div style={{ flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", color: "var(--gaq-text)" }}>{r.NumRC || r.TicketSD || "—"} — {(r.Objeto || "").slice(0, 40)}</div>
                      <div style={{ flexShrink: 0, fontWeight: 700, color: r.diasParado > 30 ? "#c0392b" : "#e67e22" }}>{r.diasParado} d.u.</div>
                      <div style={{ flexShrink: 0, fontSize: 9, color: "var(--gaq-text-3)" }}>{r.faseAtual || "—"}</div>
                    </div>
                  ));
                })()}
              </div>
            </div>

            {/* Status Detalhado + Modalidades */}
            <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 16 }}>
              <div className="gaq-card" style={{ padding: 22, borderRadius: 24 }}>
                <div style={{ fontWeight: 700, fontSize: 15, color: "var(--gaq-text)", marginBottom: 2 }}>Status detalhado</div>
                <div style={{ fontSize: 11, color: "var(--gaq-text-3)", marginBottom: 10 }}>Processos em andamento por status</div>
                {statusDetArr.slice(0, 15).map(([s, cnt]) => (
                  <div key={s} style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", borderBottom: "1px solid var(--gaq-line)", fontSize: 11, gap: 8, alignItems: "center" }}>
                    <span style={{ color: "var(--gaq-text)", flex: 1, wordBreak: "break-word", lineHeight: 1.3 }} title={s}>{s}</span>
                    <span style={{ fontWeight: 700, color: "#2e86c1", flexShrink: 0 }}>{cnt} <span style={{ fontSize: 9, fontWeight: 400, color: "var(--gaq-text-3)" }}>({emAF.length > 0 ? Math.round((cnt/emAF.length)*100) : 0}%)</span></span>
                  </div>
                ))}
              </div>
              <div className="gaq-card" style={{ padding: 22, borderRadius: 24 }}>
                <div style={{ fontWeight: 700, fontSize: 15, color: "var(--gaq-text)", marginBottom: 2 }}>Modalidades activas</div>
                <div style={{ fontSize: 11, color: "var(--gaq-text-3)", marginBottom: 10 }}>Volume e aging por modalidade</div>
                {modArr.slice(0, 10).map(m => (
                  <div key={m.name} style={{ display: "flex", justifyContent: "space-between", padding: "4px 0", borderBottom: "1px solid var(--gaq-line)", fontSize: 11, gap: 6, alignItems: "center", cursor: "pointer" }}
                    onClick={() => {
                      const procs = emAF.filter(r => (r.Modalidade || "N/I") === m.name);
                      setDrillDown({ title: "Modalidade: " + m.name + " — Em Andamento", data: procs, color: "#2e86c1", sourceAba: "operacional" }); setAba("overview");
                    }}>
                    <span style={{ color: "var(--gaq-text)", flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{m.name}</span>
                    <span style={{ display: "flex", gap: 6, alignItems: "center", flexShrink: 0 }}>
                      <span style={{ fontWeight: 700, color: "#2e86c1" }}>{m.total}</span>
                      {m.crit > 0 && <span style={{ fontSize: 9, background: "#c0392b", color: "#fff", borderRadius: 4, padding: "0 5px", fontWeight: 700 }}>{m.crit} crít.</span>}
                      <span style={{ fontSize: 9, color: "var(--gaq-text-3)" }}>~{m.media} d.u.</span>
                    </span>
                  </div>
                ))}
              </div>
            </div>

            {/* Fluxo Mensal — pré-compra e RC */}
            {(() => {
              const MESES = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
              const getYM = (d) => d ? `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}` : null;
              const ymLabel = (ym) => { const [y,m] = ym.split("-"); return MESES[parseInt(m,10)-1] + "/" + y.slice(2); };

              // Coletar todos os meses presentes na base
              const ymSet = new Set();
              fBase.forEach(r => {
                const ym1 = getYM(r.aberturaSD);
                const ym2 = getYM(r.aberturaRC);
                const ym3 = getYM(r.ultimaData);
                if (ym1) ymSet.add(ym1);
                if (ym2) ymSet.add(ym2);
                if (ym3) ymSet.add(ym3);
              });
              const allYM = [...ymSet].sort();
              if (allYM.length < 2) return null;
              // Últimos 12 meses no máximo
              const recentYM = allYM.slice(-12);

              // ── Gráfico SD ──
              // SD: processos que TÊM abertura de SD (todos, mesmo que depois gerem RC)
              const sdBase = fBase.filter(r => r.aberturaSD);
              const sdData = recentYM.map(ym => {
                const abertos = sdBase.filter(r => getYM(r.aberturaSD) === ym).length;
                // Cancelados SD: somenteSD (só tem SD, não virou RC) e está cancelado/fracassado
                const cancelados = sdBase.filter(r => r.somenteSD && (r.isCanceled || r.isFailed) && getYM(r.aberturaSD) === ym).length;
                // Concluídos SD: tem encerramento do SD nesse mês
                const concluidos = sdBase.filter(r => r.encSD && getYM(r.encSD) === ym).length;
                const saldo = abertos - cancelados - concluidos;
                return { mes: ymLabel(ym), Abertos: abertos, Cancelados: cancelados, Concluídos: concluidos, Saldo: saldo };
              });

              // ── Gráfico RC ──
              const rcBase = fBase.filter(r => r.aberturaRC);
              const rcData = recentYM.map(ym => {
                const abertos = rcBase.filter(r => getYM(r.aberturaRC) === ym).length;
                // Cancelados RC: tem RC e está cancelado/fracassado, pela data da RC
                const cancelados = rcBase.filter(r => (r.isCanceled || r.isFailed) && getYM(r.aberturaRC) === ym).length;
                // Concluídos RC: concluído e a última data preenchida cai nesse mês
                const concluidos = rcBase.filter(r => r.isConcluded && r.ultimaData && getYM(r.ultimaData) === ym).length;
                const saldo = abertos - cancelados - concluidos;
                return { mes: ymLabel(ym), Abertos: abertos, Cancelados: cancelados, Concluídos: concluidos, Saldo: saldo };
              });

              return (
                <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 16 }}>
                  <div className="gaq-card" style={{ padding: 22, borderRadius: 24 }}>
                    <div style={{ fontWeight: 700, fontSize: 15, color: "var(--gaq-text)", marginBottom: 2 }}>Fluxo mensal — pré-compra</div>
                    <div style={{ fontSize: 11, color: "var(--gaq-text-3)", marginBottom: 10 }}>Aberturas, cancelamentos e conclusões de pré-compra por mês</div>
                    <Recharts.ResponsiveContainer width="100%" height={220}>
                      <Recharts.ComposedChart data={sdData} margin={{ top: 5, right: 10, left: 0, bottom: 5 }}>
                        <Recharts.CartesianGrid strokeDasharray="3 3" stroke="var(--gaq-line)" />
                        <Recharts.XAxis dataKey="mes" fontSize={9} tick={{ fill: "var(--gaq-text-3)" }} />
                        <Recharts.YAxis fontSize={10} tick={{ fill: "var(--gaq-text-3)" }} />
                        <Recharts.Tooltip formatter={(v, n) => [n === "Saldo" ? (v >= 0 ? "+" + v : v) : v, n]} />
                        <Recharts.Legend wrapperStyle={{ fontSize: 10 }} />
                        <Recharts.Bar dataKey="Abertos" fill="#2e86c1" radius={[3,3,0,0]} />
                        <Recharts.Bar dataKey="Cancelados" fill="#e74c3c" radius={[3,3,0,0]} />
                        <Recharts.Bar dataKey="Concluídos" fill="#27ae60" radius={[3,3,0,0]} />
                        <Recharts.Bar dataKey="Saldo" fill="#d4a017" radius={[3,3,0,0]} />
                        <Recharts.ReferenceLine y={0} stroke="var(--gaq-text-3)" strokeDasharray="3 3" />
                      </Recharts.ComposedChart>
                    </Recharts.ResponsiveContainer>
                  </div>
                  <div className="gaq-card" style={{ padding: 22, borderRadius: 24 }}>
                    <div style={{ fontWeight: 700, fontSize: 15, color: "var(--gaq-text)", marginBottom: 2 }}>Fluxo mensal — RC</div>
                    <div style={{ fontSize: 11, color: "var(--gaq-text-3)", marginBottom: 10 }}>Aberturas, cancelamentos e conclusões de RC por mês</div>
                    <Recharts.ResponsiveContainer width="100%" height={220}>
                      <Recharts.ComposedChart data={rcData} margin={{ top: 5, right: 10, left: 0, bottom: 5 }}>
                        <Recharts.CartesianGrid strokeDasharray="3 3" stroke="var(--gaq-line)" />
                        <Recharts.XAxis dataKey="mes" fontSize={9} tick={{ fill: "var(--gaq-text-3)" }} />
                        <Recharts.YAxis fontSize={10} tick={{ fill: "var(--gaq-text-3)" }} />
                        <Recharts.Tooltip formatter={(v, n) => [n === "Saldo" ? (v >= 0 ? "+" + v : v) : v, n]} />
                        <Recharts.Legend wrapperStyle={{ fontSize: 10 }} />
                        <Recharts.Bar dataKey="Abertos" fill="#2e86c1" radius={[3,3,0,0]} />
                        <Recharts.Bar dataKey="Cancelados" fill="#e74c3c" radius={[3,3,0,0]} />
                        <Recharts.Bar dataKey="Concluídos" fill="#27ae60" radius={[3,3,0,0]} />
                        <Recharts.Bar dataKey="Saldo" fill="#d4a017" radius={[3,3,0,0]} />
                        <Recharts.ReferenceLine y={0} stroke="var(--gaq-text-3)" strokeDasharray="3 3" />
                      </Recharts.ComposedChart>
                    </Recharts.ResponsiveContainer>
                  </div>
                </div>
              );
            })()}

            {/* ── Distribuição por comprador e modalidade / ano ── */}
            {(() => {
              const shortName = name => { const p = (name||"").trim().split(/\s+/); return p.length > 1 ? `${p[0]} ${p[p.length-1]}` : p[0]; };
              const truncate  = (s, n) => s.length > n ? s.slice(0, n-1) + "…" : s;

              // Acumular dados por ano para compradores e modalidades
              const anoCompMap = {};
              const anoModMap  = {};
              fBase.forEach(r => {
                if (r.isCanceled || r.isFailed) return;
                const dataRef = r.aberturaRC || r.aberturaSD;
                if (!dataRef || !(dataRef instanceof Date)) return;
                const ano  = String(dataRef.getFullYear());
                const _validResp = v => { if (!v) return false; const n = nrm(v.trim()); return n && n !== "n/a" && n !== "n/i" && !n.includes("nao aplic") && !n.includes("não aplic"); };
                const comp = (_validResp(r.Comprador) ? r.Comprador : _validResp(r.Avaliador) ? r.Avaliador : "Sem responsável").trim();
                const mod  = (r.Modalidade || "N/I").trim();
                if (!anoCompMap[ano]) anoCompMap[ano] = {};
                anoCompMap[ano][comp] = (anoCompMap[ano][comp] || 0) + 1;
                if (!anoModMap[ano]) anoModMap[ano] = {};
                anoModMap[ano][mod]  = (anoModMap[ano][mod]  || 0) + 1;
              });

              const allAnos    = [...new Set([...Object.keys(anoCompMap), ...Object.keys(anoModMap)])].sort();
              if (allAnos.length === 0) return null;
              const recentAnos = allAnos.slice(-5); // últimos 5 anos para navegação
              const anoMaisRecente = recentAnos[recentAnos.length - 1];

              // displayAnos: null → apenas o mais recente; array → seleção explícita
              const displayAnos = distAnosFilter === null ? [anoMaisRecente] : distAnosFilter;

              // Paleta fixa por posição dentro de recentAnos
              const YEAR_PAL = ["#94a3b8","#64748b","#2e86c1","#059669","#d35400"];
              const anoColor   = ano => YEAR_PAL[recentAnos.indexOf(ano)] ?? "#94a3b8";
              const anoAtivo   = ano => displayAnos.includes(ano);

              const toggleAno = ano => {
                const atual = distAnosFilter === null ? [anoMaisRecente] : distAnosFilter;
                const novo  = anoAtivo(ano) ? atual.filter(a => a !== ano) : [...atual, ano].sort();
                setDistAnosFilter(novo.length === 0 ? [ano] : novo);
              };

              // ── Compradores ──
              // totais dos anos selecionados → determinam quem entra no top e a ordenação
              const buyerTotalsDisplay = {};
              displayAnos.forEach(ano => Object.entries(anoCompMap[ano]||{}).forEach(([c,n]) => { buyerTotalsDisplay[c]=(buyerTotalsDisplay[c]||0)+n; }));
              const buyerTotal = Object.values(buyerTotalsDisplay).reduce((a,b)=>a+b,0);
              const topBuyers = Object.entries(buyerTotalsDisplay).sort((a,b)=>b[1]-a[1]).slice(0,10).map(([n])=>n);

              const buyerData  = topBuyers.map(b => {
                const row = { name: shortName(b) };
                recentAnos.forEach(ano => { row[ano] = (anoCompMap[ano]||{})[b] || 0; });
                row._tot = displayAnos.reduce((s,a)=>s+(row[a]||0),0); // ordena pelo período exibido
                return row;
              }).sort((a,b)=>b._tot-a._tot);

              // ── Modalidades ──
              const modTotalsDisplay = {};
              displayAnos.forEach(ano => Object.entries(anoModMap[ano]||{}).forEach(([m,n]) => { modTotalsDisplay[m]=(modTotalsDisplay[m]||0)+n; }));
              const modTotal = Object.values(modTotalsDisplay).reduce((a,b)=>a+b,0);
              const topMods = Object.entries(modTotalsDisplay).sort((a,b)=>b[1]-a[1]).slice(0,8).map(([n])=>n);

              const modData  = topMods.map(m => {
                const row = { name: truncate(m, 30) };
                recentAnos.forEach(ano => { row[ano] = (anoModMap[ano]||{})[m] || 0; });
                row._tot = displayAnos.reduce((s,a)=>s+(row[a]||0),0);
                return row;
              }).sort((a,b)=>b._tot-a._tot);

              const BAR_H  = 11;
              const ROW_H  = displayAnos.length * (BAR_H + 3) + 18;
              const bChartH = Math.max(160, buyerData.length * ROW_H + 20);
              const mChartH = Math.max(160, modData.length  * ROW_H + 20);
              const YW_B = 88, YW_M = 110;

              // Seletor de anos compartilhado
              const YearLegend = () => (
                <div style={{ display:"flex", gap:5, marginBottom:12, flexWrap:"wrap", alignItems:"center" }}>
                  {recentAnos.map(ano => {
                    const ativo = anoAtivo(ano);
                    const cor   = anoColor(ano);
                    const isLatest = ano === anoMaisRecente;
                    return (
                      <button key={ano} onClick={() => toggleAno(ano)}
                        style={{ display:"flex", alignItems:"center", gap:5, padding:"4px 10px", borderRadius:20, cursor:"pointer", fontSize:11, fontWeight:700, border:`1.5px solid ${ativo ? cor : "var(--gaq-line)"}`, background: ativo ? cor : "var(--gaq-bg-2)", color: ativo ? "#fff" : "var(--gaq-text-3)", transition:"all .15s", position:"relative" }}>
                        {ano}
                        {isLatest && !ativo && <span style={{ fontSize:8, color:"var(--gaq-text-3)", fontWeight:400, marginLeft:2 }}>↩</span>}
                      </button>
                    );
                  })}
                  <button onClick={() => setDistAnosFilter(recentAnos.slice())}
                    style={{ padding:"4px 10px", borderRadius:20, cursor:"pointer", fontSize:11, fontWeight:600, border:"1.5px solid var(--gaq-line)", background:"var(--gaq-bg-2)", color:"var(--gaq-text-3)", transition:"all .15s" }}>
                    Todos
                  </button>
                  {distAnosFilter !== null && displayAnos.length < recentAnos.length && (
                    <button onClick={() => setDistAnosFilter(null)}
                      style={{ padding:"4px 10px", borderRadius:20, cursor:"pointer", fontSize:10, border:"none", background:"transparent", color:"var(--gaq-text-3)", textDecoration:"underline" }}>
                      resetar
                    </button>
                  )}
                </div>
              );

              const Pills = ({ items, totals, total, labelFn }) => (
                <div style={{ marginTop:12, display:"flex", flexWrap:"wrap", gap:5 }}>
                  {items.map(item => {
                    const tot = totals[item]||0;
                    const pct = total>0 ? Math.round((tot/total)*100) : 0;
                    return (
                      <div key={item} style={{ display:"flex", alignItems:"center", gap:4, background:"var(--gaq-bg-2)", borderRadius:6, padding:"2px 8px", fontSize:10, border:"1px solid var(--gaq-line)" }}>
                        <span style={{ color:"var(--gaq-text-2)" }}>{labelFn ? labelFn(item) : truncate(item,22)}</span>
                        <span style={{ fontWeight:700, color:"var(--gaq-text)" }}>{tot}</span>
                        <span style={{ color:"var(--gaq-text-3)" }}>({pct}%)</span>
                      </div>
                    );
                  })}
                </div>
              );

              const commonTooltip = {
                contentStyle:{ background:"var(--gaq-surface)", border:"1px solid var(--gaq-line)", borderRadius:8, fontSize:11, boxShadow:"0 4px 16px rgba(0,0,0,.08)" },
                cursor:{ fill:"var(--gaq-bg-2)" },
                formatter:(v, name) => [v||null, name],
              };

              return (
                <div className="gaq-responsive-grid" style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14, marginBottom:16 }}>

                  {/* ─ Card: Por Comprador ─ */}
                  <div className="gaq-card" style={{ padding:22, borderRadius:24 }}>
                    <div style={{ fontWeight:700, fontSize:15, color:"var(--gaq-text)", marginBottom:2 }}>Distribuição por comprador</div>
                    <div style={{ fontSize:11, color:"var(--gaq-text-3)", marginBottom:10 }}>
                      Processos encaminhados · data ref.: RC; se não houver, abertura SD · <b style={{ color:"var(--gaq-text-2)" }}>{buyerTotal}</b> total
                    </div>
                    <YearLegend />
                    <Recharts.ResponsiveContainer width="100%" height={bChartH}>
                      <Recharts.BarChart layout="vertical" data={buyerData}
                        margin={{ top:0, right:28, left:0, bottom:0 }}
                        barCategoryGap="22%" barGap={2}>
                        <Recharts.CartesianGrid strokeDasharray="3 3" stroke="var(--gaq-line)" horizontal={false} />
                        <Recharts.XAxis type="number" fontSize={9} tick={{ fill:"var(--gaq-text-3)" }} axisLine={false} tickLine={false} allowDecimals={false} />
                        <Recharts.YAxis type="category" dataKey="name" width={YW_B} fontSize={10} tick={{ fill:"var(--gaq-text)", fontWeight:500 }} axisLine={false} tickLine={false} />
                        <Recharts.Tooltip {...commonTooltip} />
                        {displayAnos.map(ano => (
                          <Recharts.Bar key={ano} dataKey={ano} fill={anoColor(ano)} barSize={BAR_H} radius={[0,3,3,0]}>
                            <Recharts.LabelList dataKey={ano} position="right" style={{ fontSize:8, fill:"var(--gaq-text-3)" }} formatter={v => v > 0 ? v : ""} />
                          </Recharts.Bar>
                        ))}
                      </Recharts.BarChart>
                    </Recharts.ResponsiveContainer>
                    <Pills items={topBuyers} totals={buyerTotalsDisplay} total={buyerTotal} labelFn={shortName} />
                  </div>

                  {/* ─ Card: Por Modalidade ─ */}
                  <div className="gaq-card" style={{ padding:22, borderRadius:24 }}>
                    <div style={{ fontWeight:700, fontSize:15, color:"var(--gaq-text)", marginBottom:2 }}>Distribuição por modalidade</div>
                    <div style={{ fontSize:11, color:"var(--gaq-text-3)", marginBottom:10 }}>
                      Volume de processos por modalidade e ano · <b style={{ color:"var(--gaq-text-2)" }}>{modTotal}</b> total
                    </div>
                    <YearLegend />
                    <Recharts.ResponsiveContainer width="100%" height={mChartH}>
                      <Recharts.BarChart layout="vertical" data={modData}
                        margin={{ top:0, right:28, left:0, bottom:0 }}
                        barCategoryGap="22%" barGap={2}>
                        <Recharts.CartesianGrid strokeDasharray="3 3" stroke="var(--gaq-line)" horizontal={false} />
                        <Recharts.XAxis type="number" fontSize={9} tick={{ fill:"var(--gaq-text-3)" }} axisLine={false} tickLine={false} allowDecimals={false} />
                        <Recharts.YAxis type="category" dataKey="name" width={YW_M} fontSize={10} tick={{ fill:"var(--gaq-text)", fontWeight:500 }} axisLine={false} tickLine={false} />
                        <Recharts.Tooltip {...commonTooltip} />
                        {displayAnos.map(ano => (
                          <Recharts.Bar key={ano} dataKey={ano} fill={anoColor(ano)} barSize={BAR_H} radius={[0,3,3,0]}>
                            <Recharts.LabelList dataKey={ano} position="right" style={{ fontSize:8, fill:"var(--gaq-text-3)" }} formatter={v => v > 0 ? v : ""} />
                          </Recharts.Bar>
                        ))}
                      </Recharts.BarChart>
                    </Recharts.ResponsiveContainer>
                    <Pills items={topMods} totals={modTotalsDisplay} total={modTotal} />
                  </div>

                </div>
              );
            })()}

            {/* ── Top 10 áreas: SD concluído sem RC ── */}
            {(() => {
              const semRC = fBase.filter(r => r.encSD && !r.aberturaRC && !r.isCanceled && !r.isFailed);
              if (semRC.length === 0) return null;

              const areaMap = {};
              semRC.forEach(r => {
                const area = (r["Área Requisitante"] || "N/I").trim();
                if (!areaMap[area]) areaMap[area] = { total: 0, urgente: 0 };
                areaMap[area].total++;
                if ((r.diasDesdeEncSD || 0) > 7) areaMap[area].urgente++;
              });

              const chartData = Object.entries(areaMap)
                .sort((a, b) => b[1].total - a[1].total)
                .slice(0, 10)
                .map(([area, v]) => ({
                  name   : area.length > 30 ? area.slice(0, 29) + "…" : area,
                  urgente: v.urgente,
                  atencao: v.total - v.urgente,
                  total  : v.total,
                }));

              const chartH = Math.max(180, chartData.length * 40 + 20);

              return (
                <div className="gaq-card" style={{ padding: 22, borderRadius: 24, marginBottom: 16 }}>
                  <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:14, flexWrap:"wrap", gap:8 }}>
                    <div>
                      <div style={{ fontWeight:700, fontSize:15, color:"var(--gaq-text)" }}>Top 10 áreas — SD encerrado sem RC aberta</div>
                      <div style={{ fontSize:11, color:"var(--gaq-text-3)", marginTop:2 }}>
                        Pré-compra encerrada pelo GAQ aguardando abertura da RC · <b style={{ color:"var(--gaq-text-2)" }}>{semRC.length}</b> processo{semRC.length !== 1 ? "s" : ""} no total
                      </div>
                    </div>
                    <div style={{ display:"flex", gap:12, alignItems:"center", flexWrap:"wrap" }}>
                      <div style={{ display:"flex", alignItems:"center", gap:5, fontSize:10 }}>
                        <div style={{ width:10, height:10, borderRadius:2, background:"#dc2626", flexShrink:0 }}/>
                        <span style={{ color:"var(--gaq-text-3)" }}>Urgente (&gt;7 d.u. s/ RC)</span>
                      </div>
                      <div style={{ display:"flex", alignItems:"center", gap:5, fontSize:10 }}>
                        <div style={{ width:10, height:10, borderRadius:2, background:"#f59e0b", flexShrink:0 }}/>
                        <span style={{ color:"var(--gaq-text-3)" }}>Atenção (≤7 d.u.)</span>
                      </div>
                    </div>
                  </div>

                  <Recharts.ResponsiveContainer width="100%" height={chartH}>
                    <Recharts.BarChart layout="vertical" data={chartData}
                      margin={{ top:0, right:36, left:0, bottom:0 }}
                      barCategoryGap="28%" barGap={0}>
                      <Recharts.CartesianGrid strokeDasharray="3 3" stroke="var(--gaq-line)" horizontal={false} />
                      <Recharts.XAxis type="number" fontSize={9} tick={{ fill:"var(--gaq-text-3)" }} axisLine={false} tickLine={false} allowDecimals={false} />
                      <Recharts.YAxis type="category" dataKey="name" width={138} fontSize={10} tick={{ fill:"var(--gaq-text)", fontWeight:500 }} axisLine={false} tickLine={false} />
                      <Recharts.Tooltip
                        contentStyle={{ background:"var(--gaq-surface)", border:"1px solid var(--gaq-line)", borderRadius:8, fontSize:11, boxShadow:"0 4px 16px rgba(0,0,0,.08)" }}
                        cursor={{ fill:"var(--gaq-bg-2)" }}
                        formatter={(v, name) => [v, name === "urgente" ? "Urgente (>7 d.u. s/ RC)" : "Atenção (≤7 d.u. s/ RC)"]}
                      />
                      <Recharts.Bar dataKey="urgente" stackId="s" fill="#dc2626" barSize={22} radius={[0,0,0,0]} />
                      <Recharts.Bar dataKey="atencao" stackId="s" fill="#f59e0b" barSize={22} radius={[0,3,3,0]}>
                        <Recharts.LabelList
                          content={({ x, y, width, height, value, index }) => {
                            const total = chartData[index]?.total;
                            if (!total) return null;
                            return <text x={x + width + 6} y={y + height / 2 + 1} textAnchor="start" fontSize={11} fontWeight={700} fill="var(--gaq-text)">{total}</text>;
                          }}
                        />
                      </Recharts.Bar>
                    </Recharts.BarChart>
                  </Recharts.ResponsiveContainer>
                </div>
              );
            })()}

            {/* Insights Operacionais */}
            <div className="gaq-card" style={{ padding: 22, borderRadius: 24 }}>
              <div style={{ fontWeight: 700, fontSize: 15, color: "var(--gaq-text)", marginBottom: 2 }}>Insights operacionais</div>
              <div style={{ fontSize: 11, color: "var(--gaq-text-3)", marginBottom: 10 }}>Alertas automáticos baseados nos filtros actuais</div>
              {(() => {
                const opsInsights = [];
                const bigFase = [...fasesExpandidas].filter(f => f.count > 0).sort((a,b) => b.count - a.count)[0];
                if (bigFase && bigFase.count > 5) opsInsights.push({ icon: "🔄", text: `Maior concentração na fase "${bigFase.label}" (${bigFase.area}) com ${bigFase.count} processos. Avaliar ações para destravar esta etapa.`, color: "#d35400" });
                const topComp = compArr[0];
                if (!loginArea && topComp && topComp.crit > 3) opsInsights.push({ icon: "👤", text: `${topComp.name} concentra ${topComp.crit} processos críticos. Priorizar redistribuição.`, color: "#c0392b" });
                if (loginArea && opsCriticos.length > 0) opsInsights.push({ icon: "⚠️", text: `${opsCriticos.length} processo(s) da área ${loginAreaDisp} estão críticos pela combinação SLA + entrega. Verificar priorização junto ao GAQ.`, color: "#c0392b" });
                const paradosTotal = emAF.filter(r => r.diasParado > 15).length;
                if (paradosTotal > 5) opsInsights.push({ icon: "⏸", text: `${paradosTotal} processos estagnados (>15 d.u. sem movimentação). Verificar pendências e destravar.`, color: "#e67e22" });
                if (mediaAging > 50) opsInsights.push({ icon: "⏱", text: `Aging médio de ${mediaAging} d.u. está elevado. Meta recomendada: abaixo de 50 d.u.`, color: "#c0392b" });
                if (semPedido > emAF.length * 0.3 && semPedido > 5) opsInsights.push({ icon: "📦", text: `${semPedido} processos (${Math.round((semPedido/emAF.length)*100)}%) ainda sem envio de pedido. Verificar gargalos no fluxo de aprovação.`, color: "#e67e22" });
                if (opsInsights.length === 0) opsInsights.push({ icon: "✅", text: "Indicadores operacionais dentro dos parâmetros para os filtros selecionados.", color: "#27ae60" });
                return opsInsights.map((ins, i) => (
                  <div key={i} style={{ display: "flex", gap: 10, alignItems: "flex-start", padding: "7px 0", borderBottom: i < opsInsights.length - 1 ? "1px solid var(--gaq-line)" : "none" }}>
                    <span style={{ fontSize: 16 }}>{ins.icon}</span>
                    <span style={{ fontSize: 12, color: ins.color, fontWeight: 500 }}>{ins.text}</span>
                  </div>
                ));
              })()}
            </div>
            </>;
          })()}
        </div> : <div style={{ textAlign: "center", padding: 60, color: "var(--text3)" }}>
          <div style={{ fontSize: 38 }}>📭</div><div style={{ fontWeight: 600, marginTop: 8 }}>Carregue a base para acessar o painel operacional.</div>
          <button onClick={() => setAba("upload")} style={{ marginTop: 12, background: "#1a5276", color: "#fff", border: "none", borderRadius: 8, padding: "9px 22px", cursor: "pointer", fontSize: 13 }}>Upload</button>
        </div>);
}
