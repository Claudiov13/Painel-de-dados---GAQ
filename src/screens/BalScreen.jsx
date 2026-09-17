import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function BalScreen({ cargaCPLPreg, cargaCPLResp, cargaNCL, crossFilter, renderBuscaPanel, setAba, setSelComp }) {
  return (<div>
          {renderBuscaPanel()}
          {[
            { title: "NCL — Carga por Comprador", color: "#1a5276", data: cargaNCL, tipo: "ncl", cols: C_NCL,
              extra: (d) => { const tot = cargaNCL.reduce((a,b) => a + b.total, 0) || 1; const pct = Math.round((d.total/tot)*100); return <span style={{ display: "flex", gap: 6, alignItems: "center" }}>{d.criticos > 0 && <span style={{ background: "#c0392b", color: "#fff", borderRadius: 8, padding: "0 7px", fontSize: 10, fontWeight: 700 }}>{d.criticos} crít.</span>}<b>{d.total}</b>({pct}%)</span>; } },
            { title: "CPL — Carga por Responsável", color: "#6c3483", data: cargaCPLResp, tipo: "cpl", cols: C_CPL, extra: (d) => <span><b>{d.total}</b> ativos</span> },
            { title: "CPL — Divisão por Pregoeiro", color: "#8e44ad", data: cargaCPLPreg, tipo: "cpl", cols: C_RSP,
              extra: (d) => { const tot = cargaCPLPreg.reduce((a,b) => a + b.total, 0) || 1; const pct = Math.round((d.total/tot)*100); return <span><b>{d.total}</b>({pct}%)</span>; } },
          ].map(({ title, color, data, tipo, cols, extra }) => (
            <div key={title} style={{ background: "var(--card)", borderRadius: 10, padding: 20, boxShadow: "var(--shadow)", marginBottom: 14, borderTop: `3px solid ${color}` }}>
              <div style={{ fontWeight: 700, fontSize: 14, color, marginBottom: 12 }}>{title}</div>
              {data.length === 0 ? <div style={{ color: "var(--text3)", fontSize: 13 }}>Sem dados.</div> : <>
                <ResponsiveContainer width="100%" height={Math.max(120, data.length * 30)}>
                  <BarChart data={data} layout="vertical" margin={{ left: 100, right: 20 }}>
                    <XAxis type="number" fontSize={11} /><YAxis type="category" dataKey="name" fontSize={10} width={96} />
                    <Tooltip formatter={v => [v + " proc", "Total"]} />
                    <Bar dataKey="total" radius={[0,5,5,0]} onClick={(entry) => { crossFilter("resp", entry.full); setAba("processos"); }}>
                      {data.map((_, i) => <Cell key={i} fill={cols[i % cols.length]} style={{ cursor: "pointer" }} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
                <div style={{ marginTop: 10 }}>
                  {data.map(d => (
                    <div key={d.full} onClick={() => setSelComp({ nome: d.full, tipo })}
                      style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "5px 0", borderBottom: "1px solid var(--border2)", fontSize: 12, cursor: "pointer" }}>
                      <span style={{ color, textDecoration: "underline" }}>{d.full}</span>{extra(d)}
                    </div>
                  ))}
                </div>
              </>}
            </div>
          ))}
        </div>);
}
