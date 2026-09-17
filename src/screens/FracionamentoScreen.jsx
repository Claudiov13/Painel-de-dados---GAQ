import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function FracionamentoScreen({ loginArea }) {
  return ((() => {
          const F = window.__FRACIONAMENTO__;
          if (!F || !Array.isArray(F.itens) || F.itens.length === 0) {
            return <div style={{ maxWidth: 620, margin: "60px auto", textAlign: "center", background: "var(--gaq-surface)", border: "1px solid var(--gaq-line)", borderRadius: "var(--gaq-r-xl)", padding: "32px 28px" }}>
              <div style={{ fontSize: 38 }}>📊</div>
              <div style={{ fontWeight: 700, fontSize: 16, color: "var(--gaq-text)", marginTop: 8 }}>Base de Fracionamento não carregada</div>
              <div style={{ fontSize: 12.5, color: "var(--gaq-text-3)", marginTop: 8, lineHeight: 1.6 }}>
                O arquivo <b>base_fracionamento.js</b> não foi encontrado nesta pasta.<br/>
                Gere-o rodando <b>gerar_base_fracionamento.ps1</b> (lê a planilha "Base de fracionamento MXM.xlsx") e recarregue a página.
              </div>
            </div>;
          }
          const brl = (v) => (v == null ? "—" : Number(v).toLocaleString("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 2 }));
          const brlShort = (v) => {
            const n = Number(v) || 0;
            if (n >= 1e6) return "R$ " + (n / 1e6).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " mi";
            if (n >= 1e3) return "R$ " + (n / 1e3).toLocaleString("pt-BR", { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + " mil";
            return brl(n);
          };
          const itens = F.itens;
          const byDL = (dl) => itens.filter(i => i.dl === dl).slice().sort((a, b) => b.vlrTotal - a.vlrTotal);
          const dl1 = byDL("DL 1");
          const dl2 = byDL("DL 2");
          const T = F.totais || { dl1: { grupos: dl1.length }, dl2: { grupos: dl2.length } };
          const nExceder = itens.filter(i => i.status === "A EXCEDER").length;
          const nSemSaldo = itens.filter(i => i.status === "SEM SALDO").length;
          const statusRows = (dl, st) => itens.filter(i => i.dl === dl && i.status === st).slice().sort((a, b) => b.vlrTotal - a.vlrTotal);

          const DL_COLOR = { "DL 1": "#2e86c1", "DL 2": "#e67e22" };
          // Login de AREA nao ve valores (R$): so ranking de consumo + status.
          // Admin, diretoria e usuario individual continuam com valores visiveis.
          const verValores = !loginArea;
          const tituloMetrica = verValores ? "Soma de Vlr Total por grupo de cotação" : "Ranking de consumo por grupo de cotação";

          // Tooltip custom: oculta o R$ quando o perfil nao pode ver valores.
          const tipContent = (props) => {
            const { active, payload } = props || {};
            if (!active || !payload || !payload.length) return null;
            const p = payload[0].payload;
            return <div style={{ background: "var(--gaq-surface)", border: "1px solid var(--gaq-line)", borderRadius: 6, padding: "6px 9px", fontSize: 11, color: "var(--gaq-text)", boxShadow: "var(--gaq-shadow-2)" }}>
              <div style={{ fontWeight: 700 }}>{p.cod}</div>
              <div style={{ color: "var(--gaq-text-3)" }}>{p.desc}</div>
              {verValores && <div style={{ marginTop: 3, fontWeight: 600 }}>{brl(p.vlr)}</div>}
            </div>;
          };

          // Funcao (nao componente): chamada inline evita criar uma nova
          // identidade de componente a cada render, o que remontava o grafico
          // e fazia o Recharts reanimar do zero (efeito de "piscar").
          const fracChart = (titulo, arr, dl) => {
            const color = DL_COLOR[dl];
            const data = arr.map(i => ({ name: i.cod + " · " + i.desc, cod: i.cod, desc: i.desc, vlr: i.vlrTotal }));
            const h = Math.max(260, data.length * 22 + 40);
            return <div style={{ flex: 1, minWidth: 320, background: "var(--gaq-surface)", border: "1px solid var(--gaq-line)", borderRadius: "var(--gaq-r-xl)", padding: "16px 18px", boxShadow: "var(--gaq-shadow-2)" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 10 }}>
                <div style={{ fontSize: 14, fontWeight: 700, color: "var(--gaq-text)" }}>
                  <span style={{ display: "inline-block", width: 10, height: 10, borderRadius: 3, background: color, marginRight: 7 }}></span>
                  {titulo}
                </div>
                <div style={{ fontSize: 11, color: "var(--gaq-text-3)" }}>{arr.length} grupos</div>
              </div>
              <div style={{ maxHeight: 540, overflowY: "auto", overflowX: "hidden" }}>
                <Recharts.ResponsiveContainer width="100%" height={h}>
                  <Recharts.BarChart data={data} layout="vertical" margin={{ top: 4, right: 24, left: 8, bottom: 4 }}>
                    <Recharts.CartesianGrid strokeDasharray="3 3" stroke="var(--gaq-line)" horizontal={false} />
                    <Recharts.YAxis type="category" dataKey="name" fontSize={9.5} width={168} tick={{ fill: "var(--gaq-text-3)" }} tickFormatter={v => v.length > 26 ? v.slice(0, 24) + "…" : v} interval={0} />
                    <Recharts.XAxis type="number" hide={!verValores} fontSize={9} tick={{ fill: "var(--gaq-text-3)" }} tickFormatter={brlShort} />
                    <Recharts.Tooltip content={tipContent} cursor={{ fill: color + "11" }} />
                    <Recharts.Bar dataKey="vlr" name="Vlr Total" fill={color} radius={[0, 3, 3, 0]} barSize={13} isAnimationActive={false} />
                  </Recharts.BarChart>
                </Recharts.ResponsiveContainer>
              </div>
            </div>;
          };

          const statusTable = (dl, st, color) => {
            const rows = statusRows(dl, st);
            return <div style={{ marginBottom: 14 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "var(--gaq-text)", marginBottom: 6, display: "flex", alignItems: "center", gap: 7 }}>
                <span style={{ display: "inline-block", padding: "2px 9px", borderRadius: 999, fontSize: 10, fontWeight: 700, background: color + "22", color: color }}>{st}</span>
                <span style={{ fontSize: 11, color: "var(--gaq-text-3)", fontWeight: 600 }}>{rows.length} grupo(s)</span>
              </div>
              {rows.length === 0
                ? <div style={{ fontSize: 11.5, color: "var(--gaq-text-3)", fontStyle: "italic", padding: "4px 2px" }}>Nenhum grupo neste status.</div>
                : <div className="gaq-table-scroll"><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11.5 }}>
                    <thead><tr style={{ textAlign: "left", color: "var(--gaq-text-3)", fontSize: 10.5 }}>
                      <th style={{ padding: "4px 6px", fontWeight: 600 }}>Cód.</th>
                      <th style={{ padding: "4px 6px", fontWeight: 600 }}>Grupo de cotação</th>
                      {verValores && <th style={{ padding: "4px 6px", fontWeight: 600, textAlign: "right" }}>Vlr Total</th>}
                      {verValores && <th style={{ padding: "4px 6px", fontWeight: 600, textAlign: "right" }}>Vlr Atend.</th>}
                    </tr></thead>
                    <tbody>
                      {rows.map((r, idx) => <tr key={r.cod + idx} style={{ borderTop: "1px solid var(--gaq-line)" }}>
                        <td style={{ padding: "5px 6px", fontWeight: 700, color: color, whiteSpace: "nowrap" }}>{r.cod}</td>
                        <td style={{ padding: "5px 6px", color: "var(--gaq-text)" }}>{r.desc}</td>
                        {verValores && <td style={{ padding: "5px 6px", textAlign: "right", fontWeight: 600, color: "var(--gaq-text)", whiteSpace: "nowrap" }}>{brl(r.vlrTotal)}</td>}
                        {verValores && <td style={{ padding: "5px 6px", textAlign: "right", color: "var(--gaq-text-3)", whiteSpace: "nowrap" }}>{brl(r.vlrAtend)}</td>}
                      </tr>)}
                    </tbody>
                  </table></div>}
            </div>;
          };

          const statusBlock = (dl) => {
            const color = DL_COLOR[dl];
            return <div style={{ flex: 1, minWidth: 320, background: "var(--gaq-surface)", border: "1px solid var(--gaq-line)", borderRadius: "var(--gaq-r-xl)", padding: "16px 18px", boxShadow: "var(--gaq-shadow-2)" }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: "var(--gaq-text)", marginBottom: 12 }}>
                <span style={{ display: "inline-block", width: 10, height: 10, borderRadius: 3, background: color, marginRight: 7 }}></span>
                {dl === "DL 1" ? "DL I" : "DL II"} — Atenção
              </div>
              {statusTable(dl, "A EXCEDER", "#d4a017")}
              {statusTable(dl, "SEM SALDO", "#e74c3c")}
            </div>;
          };

          const kpi = (v, l, sub, c) => <div style={{ flex: 1, minWidth: 140, background: "var(--gaq-surface)", border: "1px solid var(--gaq-line)", borderTop: "3px solid " + c, borderRadius: "var(--gaq-r-lg)", padding: "14px 16px", boxShadow: "var(--gaq-shadow-2)" }}>
            <div style={{ fontSize: 24, fontWeight: 800, color: c, lineHeight: 1.1 }}>{v}</div>
            <div style={{ fontSize: 10.5, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase", marginTop: 3 }}>{l}</div>
            {sub && <div style={{ fontSize: 11, color: "var(--gaq-text-3)", marginTop: 3 }}>{sub}</div>}
          </div>;

          return <div className="anim-fade">
            <div style={{ background: "var(--gaq-surface)", borderRadius: "var(--gaq-r-xl)", padding: "24px 28px", marginBottom: 18, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, border: "1px solid var(--gaq-line)", boxShadow: "var(--gaq-shadow-2)" }}>
              <div>
                <div style={{ fontSize: 22, fontWeight: 700, color: "var(--gaq-text)", letterSpacing: "-0.02em" }}>Base de Fracionamento (MXM)</div>
                <div style={{ fontSize: 12, color: "var(--gaq-text-3)", marginTop: 2 }}>{tituloMetrica} · DL I e DL II apartados · Fonte: {F.fonte}{F.gerado ? " · atualizado em " + F.gerado : ""}</div>
              </div>
            </div>

            <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 18 }}>
              {kpi(T.grupos != null ? T.grupos : itens.length, "Grupos de cotação", null, "#1a5276")}
              {kpi(T.dl1 ? T.dl1.grupos : dl1.length, "DL I", verValores && T.dl1 ? brlShort(T.dl1.vlrTotal) : null, "#2e86c1")}
              {kpi(T.dl2 ? T.dl2.grupos : dl2.length, "DL II", verValores && T.dl2 ? brlShort(T.dl2.vlrTotal) : null, "#e67e22")}
              {kpi(nExceder, "A Exceder", null, "#d4a017")}
              {kpi(nSemSaldo, "Sem Saldo", null, "#e74c3c")}
            </div>

            <div style={{ fontSize: 13, fontWeight: 700, color: "var(--gaq-text)", margin: "0 0 10px 2px" }}>{tituloMetrica}</div>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 22 }}>
              {fracChart("DL I", dl1, "DL 1")}
              {fracChart("DL II", dl2, "DL 2")}
            </div>

            <div style={{ fontSize: 13, fontWeight: 700, color: "var(--gaq-text)", margin: "0 0 10px 2px" }}>Status — A Exceder e Sem Saldo</div>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              {statusBlock("DL 1")}
              {statusBlock("DL 2")}
            </div>
          </div>;
        })());
}
