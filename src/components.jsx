import { ModalOverlay } from './ModalOverlay.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from './domain/service-desk.js';
const { useState, useMemo, useRef, useEffect, useCallback } = React;

const XLSX_LIB = window.XLSX;

const Papa = window.Papa;

const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;

const Icon = ({ name, size = 18, stroke = 1.6, color = "currentColor", style = {} }) => {
  const s = size, sw = stroke;
  const common = { width: s, height: s, viewBox: "0 0 24 24", fill: "none", stroke: color, strokeWidth: sw, strokeLinecap: "round", strokeLinejoin: "round", style };
  switch (name) {
    case "home":      return <svg {...common}><path d="M3 11l9-7 9 7v9a2 2 0 01-2 2h-4v-7H9v7H5a2 2 0 01-2-2v-9z"/></svg>;
    case "grid":      return <svg {...common}><rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/></svg>;
    case "chart":     return <svg {...common}><path d="M3 21h18"/><path d="M7 17V11"/><path d="M12 17V7"/><path d="M17 17V13"/></svg>;
    case "list":      return <svg {...common}><path d="M8 6h13"/><path d="M8 12h13"/><path d="M8 18h13"/><circle cx="4" cy="6" r="1"/><circle cx="4" cy="12" r="1"/><circle cx="4" cy="18" r="1"/></svg>;
    case "calendar":  return <svg {...common}><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M8 3v4M16 3v4M3 10h18"/></svg>;
    case "bell":      return <svg {...common}><path d="M6 8a6 6 0 1112 0c0 7 3 7 3 9H3c0-2 3-2 3-9z"/><path d="M10 21a2 2 0 004 0"/></svg>;
    case "alert":     return <svg {...common}><path d="M12 9v4M12 17h.01"/><path d="M10.3 3.7L2 18a2 2 0 001.7 3h16.6A2 2 0 0022 18L13.7 3.7a2 2 0 00-3.4 0z"/></svg>;
    case "shield":    return <svg {...common}><path d="M12 3l8 4v5c0 5-3.5 8-8 9-4.5-1-8-4-8-9V7l8-4z"/></svg>;
    case "users":     return <svg {...common}><circle cx="9" cy="8" r="3.5"/><path d="M3 21c0-3.3 2.7-6 6-6s6 2.7 6 6"/><circle cx="17" cy="9" r="2.5"/><path d="M21 19c0-2.5-1.8-4.5-4-5"/></svg>;
    case "subway":    return <svg {...common}><rect x="5" y="3" width="14" height="14" rx="3"/><circle cx="9" cy="13" r="1"/><circle cx="15" cy="13" r="1"/><path d="M8 17l-2 4M16 17l2 4"/><path d="M5 9h14"/></svg>;
    case "map":       return <svg {...common}><path d="M9 3l-6 3v15l6-3 6 3 6-3V3l-6 3-6-3z"/><path d="M9 3v15M15 6v15"/></svg>;
    case "compass":   return <svg {...common}><circle cx="12" cy="12" r="9"/><path d="M16 8l-2 6-6 2 2-6 6-2z"/></svg>;
    case "trend":     return <svg {...common}><path d="M3 17l6-6 4 4 8-8"/><path d="M14 7h7v7"/></svg>;
    case "search":    return <svg {...common}><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>;
    case "filter":    return <svg {...common}><path d="M3 5h18l-7 9v6l-4-2v-4L3 5z"/></svg>;
    case "more":      return <svg {...common}><circle cx="12" cy="5" r="1.5" fill={color}/><circle cx="12" cy="12" r="1.5" fill={color}/><circle cx="12" cy="19" r="1.5" fill={color}/></svg>;
    case "chevR":     return <svg {...common}><path d="M9 6l6 6-6 6"/></svg>;
    case "chevL":     return <svg {...common}><path d="M15 6l-6 6 6 6"/></svg>;
    case "chevD":     return <svg {...common}><path d="M6 9l6 6 6-6"/></svg>;
    case "chevU":     return <svg {...common}><path d="M6 15l6-6 6 6"/></svg>;
    case "x":         return <svg {...common}><path d="M6 6l12 12M18 6L6 18"/></svg>;
    case "check":     return <svg {...common}><path d="M5 12l5 5 9-12"/></svg>;
    case "plus":      return <svg {...common}><path d="M12 5v14M5 12h14"/></svg>;
    case "download":  return <svg {...common}><path d="M12 3v13M6 11l6 6 6-6"/><path d="M5 21h14"/></svg>;
    case "upload":    return <svg {...common}><path d="M12 21V8M6 13l6-6 6 6"/><path d="M5 3h14"/></svg>;
    case "refresh":   return <svg {...common}><path d="M3 12a9 9 0 0115-6.7L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 01-15 6.7L3 16"/><path d="M3 21v-5h5"/></svg>;
    case "moon":      return <svg {...common}><path d="M21 14A9 9 0 1110 3a7 7 0 0011 11z"/></svg>;
    case "sun":       return <svg {...common}><circle cx="12" cy="12" r="5"/><path d="M12 1v2M12 21v2M4.2 4.2l1.4 1.4M18.4 18.4l1.4 1.4M1 12h2M21 12h2M4.2 19.8l1.4-1.4M18.4 5.6l1.4-1.4"/></svg>;
    case "settings":  return <svg {...common}><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.6 1.6 0 00.3 1.8l.1.1a2 2 0 01-2.8 2.8l-.1-.1a1.6 1.6 0 00-1.8-.3 1.6 1.6 0 00-1 1.5V21a2 2 0 11-4 0v-.1a1.6 1.6 0 00-1-1.5 1.6 1.6 0 00-1.8.3l-.1.1a2 2 0 11-2.8-2.8l.1-.1a1.6 1.6 0 00.3-1.8 1.6 1.6 0 00-1.5-1H3a2 2 0 110-4h.1a1.6 1.6 0 001.5-1 1.6 1.6 0 00-.3-1.8l-.1-.1a2 2 0 112.8-2.8l.1.1a1.6 1.6 0 001.8.3h0a1.6 1.6 0 001-1.5V3a2 2 0 114 0v.1a1.6 1.6 0 001 1.5 1.6 1.6 0 001.8-.3l.1-.1a2 2 0 112.8 2.8l-.1.1a1.6 1.6 0 00-.3 1.8h0a1.6 1.6 0 001.5 1H21a2 2 0 110 4h-.1a1.6 1.6 0 00-1.5 1z"/></svg>;
    case "user":      return <svg {...common}><circle cx="12" cy="8" r="4"/><path d="M4 21c0-4 4-7 8-7s8 3 8 7"/></svg>;
    case "tag":       return <svg {...common}><path d="M21 12l-9 9-9-9V3h9l9 9z"/><circle cx="7.5" cy="7.5" r="1.2" fill={color}/></svg>;
    case "eye":       return <svg {...common}><path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/></svg>;
    case "clock":     return <svg {...common}><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>;
    case "doc":       return <svg {...common}><path d="M14 3H6a2 2 0 00-2 2v14a2 2 0 002 2h12a2 2 0 002-2V9l-6-6z"/><path d="M14 3v6h6"/></svg>;
    case "spark":     return <svg {...common}><path d="M12 3v4M12 17v4M3 12h4M17 12h4M5.6 5.6l2.8 2.8M15.6 15.6l2.8 2.8M5.6 18.4l2.8-2.8M15.6 8.4l2.8-2.8"/></svg>;
    case "logout":    return <svg {...common}><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><path d="M16 17l5-5-5-5"/><path d="M21 12H9"/></svg>;
    case "external":  return <svg {...common}><path d="M14 4h6v6"/><path d="M20 4l-9 9"/><path d="M19 14v5a2 2 0 01-2 2H5a2 2 0 01-2-2V7a2 2 0 012-2h5"/></svg>;
    case "menu":      return <svg {...common}><path d="M3 6h18M3 12h18M3 18h18"/></svg>;
    case "sidebar":   return <svg {...common}><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M9 3v18"/></svg>;
    case "key":       return <svg {...common}><circle cx="8" cy="14" r="4"/><path d="M11 11l9-9"/><path d="M16 6l3 3"/><path d="M14 8l3 3"/></svg>;
    default: return null;
  }
};

const Spark = ({ data = [], color = "var(--gaq-blue)", w = 80, h = 24 }) => {
  if (!data.length) return null;
  const max = Math.max(...data), min = Math.min(...data);
  const span = max - min || 1;
  const pts = data.map((v, i) => `${(i/(data.length-1))*w},${h - ((v-min)/span)*h}`).join(" ");
  return (
    <svg width={w} height={h} style={{ overflow: "visible" }}>
      <polyline points={pts} fill="none" stroke={color} strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/>
      <circle cx={w} cy={h - ((data[data.length-1]-min)/span)*h} r="2.4" fill={color}/>
    </svg>
  );
};

const MiniCard = ({ title, value, sub }) => (
  <div style={{ background: "rgba(255,255,255,0.12)", borderRadius: "var(--gaq-r-sm)", padding: "10px 12px", flex: 1, minWidth: 80 }}>
    <div style={{ fontSize: 9, color: "rgba(255,255,255,0.65)", textTransform: "uppercase", letterSpacing: .5, marginBottom: 2, fontWeight: 600 }}>{title}</div>
    <div style={{ fontSize: 22, fontWeight: 700, color: "#fff", fontVariantNumeric: "tabular-nums" }}>{value}</div>
    {sub && <div style={{ fontSize: 9, color: "rgba(255,255,255,0.55)", marginTop: 1 }}>{sub}</div>}
  </div>
);

const GBox = ({ title, color, children, style = {} }) => (
  <div className="w100-sm" style={{ background: color, borderRadius: "var(--gaq-r-lg)", padding: "14px 16px", boxShadow: "var(--gaq-shadow-2)", ...style }}>
    <div style={{ fontSize: 9, fontWeight: 600, color: "rgba(255,255,255,0.7)", textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: 8 }}>{title}</div>
    <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>{children}</div>
  </div>
);

const WCard = ({ title, border = "var(--gaq-blue)", children }) => (
  <div className="gaq-card" style={{ padding: 18, borderTop: `3px solid ${border}` }}>
    <div style={{ fontWeight: 600, marginBottom: 10, color: "var(--gaq-text)", fontSize: 14, letterSpacing: "-0.01em" }}>{title}</div>
    {children}
  </div>
);

const KpiBox = ({ title, subtitle, color, children }) => (
  <div style={{ flex: 1, border: `1px solid var(--gaq-line)`, borderRadius: "var(--gaq-r-lg)", padding: 16, background: "var(--gaq-surface)", borderTop: `3px solid ${color}` }}>
    <div style={{ fontSize: 11, fontWeight: 600, color, textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: subtitle ? 2 : 10 }}>{title}</div>
    {subtitle && <div style={{ fontSize: 9, color, opacity: .7, marginBottom: 8, fontStyle: "italic" }}>{subtitle}</div>}
    <div className="grid-2 gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>{children}</div>
  </div>
);

const BadgeScore = ({ score }) => {
  const c = score >= 70 ? "var(--gaq-green)" : score >= 50 ? "var(--gaq-orange)" : "var(--gaq-red)";
  const bg = score >= 70 ? "rgba(52,199,89,0.16)" : score >= 50 ? "rgba(255,149,0,0.18)" : "rgba(255,59,48,0.14)";
  const tc = score >= 70 ? "#1e8e3e" : score >= 50 ? "#c66600" : "var(--gaq-red)";
  return <span style={{ background: bg, color: tc, borderRadius: 999, padding: "2px 10px", fontSize: 12, fontWeight: 700, fontVariantNumeric: "tabular-nums" }}>{score}</span>;
};

const Sel = ({ label, value, onChange, opts }) => (
  <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
    <label style={{ fontSize: 11, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.04em" }}>{label}</label>
    <select value={value} onChange={e => onChange(e.target.value)}
      className="gaq-select"
      style={{ padding: "5px 8px", borderRadius: 8, border: value ? "1.5px solid var(--gaq-blue)" : "1px solid var(--gaq-line-2)", fontSize: 12, background: value ? "var(--gaq-blue-tint)" : "var(--gaq-surface)", color: value ? "var(--gaq-blue)" : "var(--gaq-text)", width: "100%", fontWeight: value ? 600 : 400, transition: "border .12s, background .12s", height: 32 }}>
      <option value="">Todos</option>
      {opts.map(o => <option key={o} value={o}>{o}</option>)}
    </select>
  </div>
);

const RankRow = ({ rank, name, badge, badgeColor, extra, onClick }) => (
  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", borderBottom: "1px solid var(--gaq-line)", fontSize: 13 }}>
    <span style={onClick ? { cursor: "pointer", color: "var(--gaq-blue)", fontWeight: 500 } : { fontWeight: 500 }} onClick={onClick}>
      <b style={{ color: "var(--gaq-text-4)", marginRight: 6, fontSize: 11, fontWeight: 700 }}>#{rank}</b>{name}
    </span>
    <span style={{ display: "flex", gap: 6, alignItems: "center" }}>
      <span style={{ background: badgeColor + "20", color: badgeColor, borderRadius: 999, padding: "2px 10px", fontSize: 11, fontWeight: 700, fontVariantNumeric: "tabular-nums" }}>{badge}</span>
      {extra && <span style={{ color: "var(--gaq-text-3)", fontSize: 11 }}>{extra}</span>}
    </span>
  </div>
);

const TagBadge = ({ tag, onRemove, small, onDark }) => {
  const c = TagsManager.getTagColor(tag);
  const bg = onDark ? c : c + "22";
  const fg = onDark ? "#fff" : c;
  const bd = onDark ? "rgba(255,255,255,0.28)" : `${c}44`;
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 3, background: bg, color: fg, borderRadius: 12, padding: small ? "0 6px" : "1px 8px", fontSize: small ? 9 : 10, fontWeight: 700, border: `1px solid ${bd}`, whiteSpace: "nowrap" }}>
      {tag}
      {onRemove && (
        <span
          role="button"
          onClick={(e) => { e.preventDefault(); e.stopPropagation(); onRemove(); }}
          style={{ cursor: "pointer", marginLeft: 3, fontWeight: 900, fontSize: small ? 12 : 14, lineHeight: 1, color: fg, display: "inline-flex", alignItems: "center", justifyContent: "center", width: small ? 14 : 16, height: small ? 14 : 16, borderRadius: "50%", background: onDark ? "rgba(255,255,255,0.18)" : c + "22" }}
        >×</span>
      )}
    </span>
  );
};

const Pagination = ({ total, page, pageSize, onPage, onSize }) => {
  const tp = Math.ceil(total / pageSize);
  if (tp <= 1) return null;
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 0", flexWrap: "wrap", gap: 8 }}>
      <span className="gaq-meta gaq-num">{((page-1)*pageSize)+1}–{Math.min(page*pageSize, total)} de {total}</span>
      <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
        <select value={pageSize} onChange={e => { onSize(+e.target.value); onPage(1); }}
          style={{ padding: "3px 6px", borderRadius: 8, border: "1px solid var(--gaq-line-2)", fontSize: 11, background: "var(--gaq-surface)", color: "var(--gaq-text)", height: 28 }}>
          {PAGE_SIZES.map(n => <option key={n} value={n}>{n}/pág</option>)}
        </select>
        <button className="gaq-icon-btn" disabled={page <= 1} onClick={() => onPage(page-1)} style={{ opacity: page <= 1 ? .3 : 1 }}><Icon name="chevL" size={14}/></button>
        <span style={{ fontSize: 12, fontWeight: 600, color: "var(--gaq-text-2)", minWidth: 60, textAlign: "center" }}>{page}/{tp}</span>
        <button className="gaq-icon-btn" disabled={page >= tp} onClick={() => onPage(page+1)} style={{ opacity: page >= tp ? .3 : 1 }}><Icon name="chevR" size={14}/></button>
      </div>
    </div>
  );
};

const DrillDownPanel = ({
  drillDown,
  defaultSourceAba = "overview",
  setDrillDown,
  setAba,
  setSelProc,
  pgDrillDown,
  pgDrillDownSz,
  setPgDrillDown,
  setPgDrillDownSz,
  phaseIntervals,
}) => {
  if (!drillDown) return null;

  const sourceAba = drillDown.sourceAba || defaultSourceAba;
  const accent = drillDown.color || "var(--gaq-blue)";
  const sortedData = [...drillDown.data].sort((a, b) => {
    const at = a.diasTotais || a.diasSDAberto || a.diasDesdeEncSD || 0;
    const bt = b.diasTotais || b.diasSDAberto || b.diasDesdeEncSD || 0;
    return bt - at;
  });
  const total = sortedData.length;
  const start = (pgDrillDown - 1) * pgDrillDownSz;
  const rows = sortedData.slice(start, start + pgDrillDownSz);
  const backLabel = sourceAba === "operacional"
    ? "ao Operacional"
    : sourceAba === "gestao"
      ? "ao Painel de Gestao"
      : sourceAba === "qualidade"
        ? "a Qualidade de Dados"
      : sourceAba === "executivo"
        ? "ao Executivo"
        : "a Visao Geral";

  const fmtDate = (dt) => dt ? dt.toLocaleDateString("pt-BR") : "Sem data";
  const fmtStamp = (dt) => {
    if (!dt) return "Sem registo";
    const now = new Date();
    return dt.toDateString() === now.toDateString()
      ? dt.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })
      : dt.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" });
  };
  const fmtSignedDU = (n, empty = "Sem data") => {
    if (n == null || Number.isNaN(n)) return empty;
    if (n < 0) return `${Math.abs(n)} d.u. atraso`;
    if (n === 0) return "Hoje";
    return `${n} d.u.`;
  };
  const pillStyle = (fg, bg) => ({
    display: "inline-flex",
    alignItems: "center",
    gap: 6,
    padding: "6px 10px",
    borderRadius: 999,
    background: bg,
    color: fg,
    fontSize: 11,
    fontWeight: 600,
    lineHeight: 1,
    whiteSpace: "nowrap",
  });

  return (
    <div>
      <button
        onClick={() => {
          setDrillDown(null);
          if (sourceAba !== defaultSourceAba) setAba(sourceAba);
        }}
        className="gaq-btn"
        style={{ marginBottom: 14 }}
      >
        <Icon name="chevL" size={14}/> Voltar {backLabel}
      </button>

      <div className="gaq-card" style={{ padding: 22, marginBottom: 16, borderTop: `3px solid ${accent}` }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 14, flexWrap: "wrap" }}>
          <div>
            <div style={{ fontWeight: 700, fontSize: 20, color: accent, letterSpacing: "-0.02em" }}>{drillDown.title}</div>
            <div className="gaq-meta" style={{ marginTop: 4 }}>{total.toLocaleString("pt-BR")} processo(s) no recorte</div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
            <div style={{ ...pillStyle(accent, `${accent}12`), fontSize: 12, fontWeight: 700 }}>
              Lista detalhada com prazo de entrega e SLA
            </div>
            {total > 0 && <button className="gaq-btn" style={{ fontSize: 12 }}
              onClick={() => {
                const slug = (drillDown.title || "lista")
                  .normalize("NFD").replace(new RegExp("[\\u0300-\\u036f]", "g"), "")
                  .replace(/[^a-zA-Z0-9]+/g, "_").replace(/^_+|_+$/g, "").toLowerCase();
                exportToExcel(sortedData, `${slug || "lista"}.xlsx`);
              }}>
              <Icon name="download" size={13}/> Exportar Excel
            </button>}
          </div>
        </div>
      </div>

      {total === 0
        ? <div style={{ textAlign: "center", padding: 30, color: "var(--gaq-text-3)" }}>Nenhum processo nesta categoria.</div>
        : rows.map((r, i) => {
            const sla = r._sla || calcSLAClassification(r, phaseIntervals);
            const bucket = sla ? (SLA_BUCKETS[sla.bucket] || SLA_BUCKETS.no_prazo) : null;
            const aging = sla ? sla.diasConsumidos : (r.diasTotais || r.diasSDAberto || r.diasSD || 0);
            const owner = r.Comprador || r.Pregoeiro || r.AnalistaContrato || r.respNCL || r.respFase || r.Avaliador || "Sem responsavel";
            const watched = !!(r.ProcessKey && WatchlistManager.isWatched(r.ProcessKey));
            const entregaColor = sla?.entregaVencida
              ? "#ff3b30"
              : (sla?.projecaoForaCronograma || sla?.janelaComprometida || sla?.atencaoEntrega || sla?.entregaProxima)
                ? "#ff9500"
                : r.dataEntrega
                  ? "#2e86c1"
                  : "var(--gaq-text-3)";
            const entregaBg = sla?.entregaVencida
              ? "rgba(255,59,48,0.12)"
              : (sla?.projecaoForaCronograma || sla?.janelaComprometida || sla?.atencaoEntrega || sla?.entregaProxima)
                ? "rgba(255,149,0,0.14)"
                : r.dataEntrega
                  ? "rgba(46,134,193,0.10)"
                  : "rgba(107,114,128,0.12)";

            return (
              <div
                key={`${r.ProcessKey || r.NumRC || r.TicketSD || "proc"}-${start + i}`}
                className="gaq-card"
                style={{ padding: "18px 20px", marginBottom: 12, border: `1px solid ${accent}18`, borderRadius: 22, cursor: "pointer" }}
                onClick={() => setSelProc(r)}
              >
                <div style={{ display: "flex", justifyContent: "space-between", gap: 18, alignItems: "flex-start", flexWrap: "wrap" }}>
                  <div style={{ flex: "1 1 520px", minWidth: 0 }}>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center", marginBottom: 10 }}>
                      <span style={pillStyle("var(--gaq-text-3)", "var(--gaq-bg-2)")}>{r.ProcessKey || "Sem chave"}</span>
                      {r.NumRC && <span style={pillStyle("var(--gaq-blue)", "rgba(46,134,193,0.12)")}>{watched ? "Olho " : ""}{r.NumRC}</span>}
                      {r.TicketSD && <span style={pillStyle("var(--gaq-green)", "rgba(52,199,89,0.12)")}>Pré-compra {r.TicketSD}</span>}
                      <span style={pillStyle("var(--gaq-text-2)", "rgba(15,23,42,0.06)")}>{r.Modalidade || "Sem modalidade"}</span>
                      <span style={pillStyle("var(--gaq-text-2)", "rgba(15,23,42,0.06)")}>{owner}</span>
                    </div>

                    <div style={{ fontSize: 13, fontWeight: 400, color: "var(--gaq-text-2)", lineHeight: 1.45, marginBottom: 10 }}>
                      {r.Objeto || "Sem objeto informado"}
                    </div>

                    <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                      <span style={pillStyle("var(--gaq-text-2)", "rgba(15,23,42,0.06)")}>Area: {r["Área Requisitante"] || "Nao informada"}</span>
                      <span style={pillStyle("var(--gaq-text-2)", "rgba(15,23,42,0.06)")}>Status: {r.statusDet || r.status || "Sem status"}</span>
                      <span style={pillStyle("var(--gaq-text-2)", "rgba(15,23,42,0.06)")}>Abertura: {fmtDate(r.dataAbertura)}</span>
                      <span style={pillStyle(entregaColor, entregaBg)}>Prazo de entrega: {fmtDate(r.dataEntrega)} · {fmtSignedDU(sla?.diasEntregaUteis)}</span>
                      <span style={pillStyle("var(--gaq-text-2)", "rgba(15,23,42,0.06)")}>Ultimo movimento: {fmtStamp(r.ultimaData)}</span>
                    </div>
                  </div>

                  <div style={{ minWidth: 126, display: "grid", gap: 10, justifyItems: "end" }}>
                    {bucket && <span style={{ ...pillStyle(bucket.color, bucket.bg), fontWeight: 700 }}>{bucket.label}</span>}
                    <div style={{ textAlign: "right" }}>
                      <div className="gaq-num" style={{ fontSize: 30, fontWeight: 800, lineHeight: 1, color: bucket ? bucket.color : accent }}>
                        {Number(aging || 0).toLocaleString("pt-BR")}
                      </div>
                      <div className="gaq-meta" style={{ marginTop: 4 }}>dias uteis</div>
                    </div>
                    <div className="gaq-meta" style={{ color: "var(--gaq-text-3)" }}>Abrir timeline <Icon name="chevR" size={12}/></div>
                  </div>
                </div>
              </div>
            );
          })}
      {total > 0 && <Pagination total={total} page={pgDrillDown} pageSize={pgDrillDownSz} onPage={setPgDrillDown} onSize={setPgDrillDownSz} />}
    </div>
  );
};

const CopyChip = ({ label, value, color, solid }) => {
  const [copied, setCopied] = React.useState(false);
  const doCopy = async (e) => {
    e.stopPropagation();
    let ok = false;
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(value);
        ok = true;
      }
    } catch (_) {}
    if (!ok) {
      try {
        const ta = document.createElement("textarea");
        ta.value = value;
        ta.style.cssText = "position:fixed;top:-9999px;left:-9999px;width:1px;height:1px;opacity:0;";
        document.body.appendChild(ta);
        ta.focus(); ta.select();
        ok = document.execCommand("copy");
        document.body.removeChild(ta);
      } catch (_) {}
    }
    if (ok !== false) { setCopied(true); setTimeout(() => setCopied(false), 1500); }
  };
  return (
    <button onClick={doCopy} title={`Copiar ${label}: ${value}`}
      style={{ display: "inline-flex", alignItems: "center", gap: 4, background: copied ? "#27ae60" : (solid ? (color || "#1a5276") : "rgba(255,255,255,0.15)"), border: solid ? `1px solid ${color || "#1a5276"}` : "1px solid rgba(255,255,255,0.3)", borderRadius: 5, padding: "2px 8px", fontSize: 10, fontWeight: 700, color: "#fff", cursor: "pointer", transition: "background .2s" }}>
      <span style={{ fontSize: 9, opacity: .75, fontWeight: 400 }}>{label}:</span>
      <span>{value}</span>
      <span style={{ fontSize: 10 }}>{copied ? "✓" : "⎘"}</span>
    </button>
  );
};

const HistoricoModal = ({ proc, onClose }) => (
  <ModalOverlay onDismiss={onClose} style={{ position: "fixed", inset: 0, background: "#0009", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 2000 }} onClick={onClose}>
    <div style={{ background: "var(--card)", borderRadius: 12, width: "min(720px,95vw)", maxHeight: "82vh", display: "flex", flexDirection: "column", boxShadow: "0 8px 40px #0006" }} onClick={e => e.stopPropagation()}>
      <div style={{ padding: "14px 20px", background: "var(--header2)", color: "#fff", borderRadius: "12px 12px 0 0", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 14 }}>Histórico do Processo</div>
          <div style={{ fontSize: 10, opacity: .7 }}>{proc.NumRC || "—"} · {proc.Comprador || proc.Pregoeiro || "—"}</div>
        </div>
        <button onClick={onClose} style={{ background: "none", border: "none", color: "#fff", fontSize: 22, cursor: "pointer" }}>×</button>
      </div>
      <div style={{ overflow: "auto", padding: "18px 22px", flex: 1 }}>
        {/* Histórico do sistema (GAQ/NCL) */}
        {proc.Historico && <div style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 10, fontWeight: 700, color: "#2e86c1", textTransform: "uppercase", letterSpacing: .5, marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}>
            <span style={{ width: 10, height: 10, borderRadius: "50%", background: "#2e86c1", display: "inline-block" }} />
            Histórico NCL
          </div>
          <div style={{ fontSize: 12, color: "var(--text)", whiteSpace: "pre-wrap", lineHeight: 1.7, background: "rgba(46,134,193,0.06)", border: "1px solid rgba(46,134,193,0.2)", borderRadius: 8, padding: 14 }}>{proc.Historico}</div>
        </div>}
        {proc.HistoricoCPL && <div style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 10, fontWeight: 700, color: "#8e44ad", textTransform: "uppercase", letterSpacing: .5, marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}>
            <span style={{ width: 10, height: 10, borderRadius: "50%", background: "#8e44ad", display: "inline-block" }} />
            Histórico CPL
          </div>
          <div style={{ fontSize: 12, color: "var(--text)", whiteSpace: "pre-wrap", lineHeight: 1.7, background: "rgba(142,68,173,0.06)", border: "1px solid rgba(142,68,173,0.25)", borderRadius: 8, padding: 14 }}>{proc.HistoricoCPL}</div>
        </div>}
        {/* Histórico SCONT */}
        {proc.HistoricoScont && <div style={{ marginBottom: 8 }}>
          <div style={{ fontSize: 10, fontWeight: 700, color: "#16a085", textTransform: "uppercase", letterSpacing: .5, marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}>
            <span style={{ width: 10, height: 10, borderRadius: "50%", background: "#16a085", display: "inline-block" }} />
            Histórico Scont
          </div>
          <div style={{ fontSize: 12, color: "var(--text)", whiteSpace: "pre-wrap", lineHeight: 1.7, background: "rgba(22,160,133,0.06)", border: "1px solid rgba(22,160,133,0.25)", borderRadius: 8, padding: 14 }}>{proc.HistoricoScont}</div>
        </div>}
        {!proc.Historico && !proc.HistoricoCPL && !proc.HistoricoScont && (
          <div style={{ textAlign: "center", padding: 30, color: "var(--text3)", fontSize: 13 }}>Sem histórico registrado.</div>
        )}
      </div>
    </div>
  </ModalOverlay>
);

const TagModal = ({ proc, onClose }) => {
  const tags = TagsManager.getForProcess(proc.ProcessKey, proc.TicketSD);

  return (
    <ModalOverlay onDismiss={onClose} style={{ position: "fixed", inset: 0, background: "#0009", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 2100 }} onClick={onClose}>
      <div style={{ background: "var(--card)", borderRadius: 12, width: "min(480px,95vw)", maxHeight: "70vh", display: "flex", flexDirection: "column", boxShadow: "0 8px 40px #0006" }} onClick={e => e.stopPropagation()}>
        <div style={{ padding: "14px 20px", background: "#34495e", color: "#fff", borderRadius: "12px 12px 0 0" }}>
          <div style={{ fontWeight: 700, fontSize: 14 }}>Tag do processo</div>
          <div style={{ fontSize: 10, opacity: .7 }}>{proc.NumRC || "—"} · {(proc.Objeto || "").slice(0,60) || "—"}</div>
        </div>
        <div style={{ padding: "16px 20px", flex: 1, overflow: "auto" }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text2)", marginBottom: 6, textTransform: "uppercase" }}>Tag definida pelo dado</div>
          <div style={{ display: "flex", gap: 5, flexWrap: "wrap", marginBottom: 14, minHeight: 28 }}>
            {tags.length === 0
              ? <span style={{ color: "var(--text3)", fontSize: 12 }}>Sem valor no campo "PROJETO / EVENTO / AÇÃO"</span>
              : tags.map(t => <TagBadge key={t} tag={t} />)}
          </div>
          <div style={{ fontSize: 11, color: "var(--text3)", lineHeight: 1.5, background: "var(--card2)", border: "1px solid var(--border)", borderRadius: 8, padding: "8px 12px" }}>
            A tag é derivada do campo <b>PROJETO / EVENTO / AÇÃO</b> do JSON principal.
            Para alterá-la, edite o valor desse campo na base e recarregue os dados.
          </div>
        </div>
        <div style={{ padding: "10px 20px", borderTop: "1px solid var(--border)", textAlign: "right" }}>
          <button onClick={onClose} style={{ padding: "7px 20px", borderRadius: 6, background: "var(--card2)", border: "1px solid var(--border)", color: "var(--text)", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>Fechar</button>
        </div>
      </div>
    </ModalOverlay>
  );
};

const EmailDiarioModal = ({ alertSDAberto, fBase, serviceDeskData, onClose }) => {
  const [cpStatus, setCpStatus] = React.useState("");
  const MARCIO_EMAIL = "mmarcenal@sesc.com.br";
  const hoje = new Date();
  const ontem = new Date(hoje); ontem.setDate(ontem.getDate() - 1);
  const hojeStr   = hoje.toLocaleDateString("pt-BR", { weekday:"long", year:"numeric", month:"long", day:"numeric" });
  const ontemStr  = ontem.toLocaleDateString("pt-BR", { weekday:"long", year:"numeric", month:"long", day:"numeric" });
  const ontemSimples = ontem.toLocaleDateString("pt-BR");
  const isD1 = d => d instanceof Date && d.toDateString() === ontem.toDateString();
  const fmtD = d => d instanceof Date ? d.toLocaleDateString("pt-BR") : "—";
  const esc  = s => String(s||"—").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");

  const rcHoje  = (fBase||[]).filter(r => r.aberturaRC && isD1(r.aberturaRC) && !r.isCanceled && !r.isFailed);
  const sdHoje  = (alertSDAberto||[]).filter(r => r.aberturaSD && isD1(r.aberturaSD));

  const sdRows = sdHoje.map(r => {
    const sdRec  = getServiceDeskRecord(r.TicketSD, serviceDeskData);
    const sdCli  = sdRec ? cleanServiceDeskObject(sdRec.informacoes_do_cliente || {}) : {};
    const sdInfo = sdRec ? cleanServiceDeskObject(sdRec.informacoes_do_chamado  || {}) : {};
    const sdSum  = getServiceDeskAlertSummary(r, serviceDeskData);
    return {
      ticket   : r.TicketSD || "—",
      abertura : fmtD(r.aberturaSD),
      objeto   : (r.Objeto || "—").slice(0, 100),
      status   : sdSum ? sdSum.estado : (sdInfo.estado || "—"),
      area     : r["Área Requisitante"] || "—",
      avaliador: r.Avaliador || "—",
      cliente  : [sdCli.nome, sdCli.sobrenome].filter(Boolean).join(" ") || "—",
    };
  });

  const rcRows = rcHoje.map(r => ({
    numRC   : r.NumRC || "—",
    abertura: fmtD(r.aberturaRC),
    objeto  : (r.Objeto || "—").slice(0, 100),
    status  : r.statusDet || r.status || "—",
    area    : r["Área Requisitante"] || "—",
    comprador: r.Comprador || r.respNCL || "—",
  }));

  const TH = "padding:8px 10px;background:#0f3d5e;color:#fff;font-size:11px;font-weight:700;text-align:left;border:1px solid #0a2d44";
  const TD = "padding:6px 10px;border:1px solid #dbe3ee;font-size:11px;vertical-align:top;color:#1d1d1f";
  const TBL = "width:100%;border-collapse:collapse;margin-bottom:24px;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif";

  const buildHtmlEmail = () => {
    const sdHtml = sdRows.length
      ? `<table style="${TBL}"><thead><tr>
           <th style="${TH}">Ticket SD</th><th style="${TH}">Abertura</th><th style="${TH}">Objeto</th>
           <th style="${TH}">Status</th><th style="${TH}">Área Requisitante</th><th style="${TH}">Avaliador</th><th style="${TH}">Cliente</th>
         </tr></thead><tbody>
         ${sdRows.map((r,i) => `<tr style="background:${i%2===0?'#fff':'#f9fafb'}">
           <td style="${TD}">${esc(r.ticket)}</td><td style="${TD}" nowrap>${esc(r.abertura)}</td>
           <td style="${TD}">${esc(r.objeto)}</td><td style="${TD}">${esc(r.status)}</td>
           <td style="${TD}">${esc(r.area)}</td><td style="${TD}">${esc(r.avaliador)}</td><td style="${TD}">${esc(r.cliente)}</td>
         </tr>`).join("")}
         </tbody></table>`
      : `<p style="color:#6b7280;font-size:12px">Nenhum ticket aberto nessa data em andamento.</p>`;

    const rcHtml = rcRows.length
      ? `<table style="${TBL}"><thead><tr>
           <th style="${TH}">Nº RC</th><th style="${TH}">Recebimento RC</th><th style="${TH}">Objeto</th>
           <th style="${TH}">Status</th><th style="${TH}">Área Requisitante</th><th style="${TH}">Comprador</th>
         </tr></thead><tbody>
         ${rcRows.map((r,i) => `<tr style="background:${i%2===0?'#fff':'#f9fafb'}">
           <td style="${TD}">${esc(r.numRC)}</td><td style="${TD}" nowrap>${esc(r.abertura)}</td>
           <td style="${TD}">${esc(r.objeto)}</td><td style="${TD}">${esc(r.status)}</td>
           <td style="${TD}">${esc(r.area)}</td><td style="${TD}">${esc(r.comprador)}</td>
         </tr>`).join("")}
         </tbody></table>`
      : `<p style="color:#6b7280;font-size:12px">Nenhuma RC recebida nessa data.</p>`;

    return `<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;color:#1d1d1f;max-width:1020px">
      <p>Prezado Márcio,</p>
      <p>Segue o acompanhamento diário da GAQ referente a <strong>${ontemStr}</strong> (D-1).</p>
      <h3 style="font-size:13px;color:#0f3d5e;border-left:4px solid #2e86c1;padding-left:8px;margin:20px 0 10px">
        Pré-compra SD — abertas em ${ontemSimples} (${sdRows.length})
      </h3>
      ${sdHtml}
      <h3 style="font-size:13px;color:#0f3d5e;border-left:4px solid #27ae60;padding-left:8px;margin:20px 0 10px">
        RCs recebidas em ${ontemSimples} (${rcRows.length})
      </h3>
      ${rcHtml}
      <p style="font-size:11px;color:#6b7280;border-top:1px solid #e5e7eb;padding-top:10px;margin-top:20px">
        Gerado pelo Painel GAQ · ${hojeStr}
      </p>
    </div>`;
  };

  const copiarEEnviar = async () => {
    const html = buildHtmlEmail();
    try {
      const blob = new Blob([html], { type: "text/html" });
      await navigator.clipboard.write([new ClipboardItem({ "text/html": blob })]);
      setCpStatus("copiado");
    } catch {
      setCpStatus("erro");
    }
    const subject = `Acompanhamento GAQ — D-1: ${ontemSimples} · SD: ${sdRows.length} · RC: ${rcRows.length}`;
    const body    = `Prezado Márcio,\n\nSegue o acompanhamento diário da GAQ referente a ${ontemStr} (D-1).\n\n• Pré-compra SD abertas em ${ontemSimples}: ${sdRows.length}\n• RCs recebidas em ${ontemSimples}: ${rcRows.length}\n\n(Cole o conteúdo HTML copiado neste corpo de e-mail para exibir as tabelas formatadas.)\n\nAtenciosamente,\nGAQ — Gerência de Aquisições · DN`;
    setTimeout(() => {
      const a = document.createElement("a");
      a.href = `mailto:${MARCIO_EMAIL}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
      a.click();
      setTimeout(() => setCpStatus(""), 3000);
    }, 200);
  };

  const thStyle = { padding:"7px 10px", background:"#0f3d5e", color:"#fff", fontSize:10, fontWeight:700, textAlign:"left", borderRight:"1px solid #1a5276", whiteSpace:"nowrap" };
  const tdStyle = { padding:"5px 9px", fontSize:11, borderBottom:"1px solid var(--gaq-line)", verticalAlign:"top", color:"var(--gaq-text)", lineHeight:1.35 };
  const tblStyle = { width:"100%", borderCollapse:"collapse", marginBottom:20 };

  return (
    <ModalOverlay onDismiss={onClose} style={{ position:"fixed", inset:0, background:"rgba(15,23,42,.6)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:2200, padding:16 }} onClick={onClose}>
      <div style={{ background:"var(--gaq-surface)", borderRadius:14, width:"min(1020px,98vw)", maxHeight:"92vh", display:"flex", flexDirection:"column", boxShadow:"0 12px 48px rgba(0,0,0,.25)", border:"1px solid var(--gaq-line)" }} onClick={e => e.stopPropagation()}>

        {/* Header */}
        <div style={{ padding:"18px 24px", background:"linear-gradient(135deg,#0f3d5e,#1a5276)", color:"#fff", borderRadius:"14px 14px 0 0", display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexShrink:0 }}>
          <div>
            <div style={{ fontSize:10, opacity:.72, textTransform:"uppercase", letterSpacing:".05em", marginBottom:4 }}>Email Diário · GAQ — Gerência de Aquisições · DN</div>
            <div style={{ fontSize:18, fontWeight:700, letterSpacing:"-.01em" }}>D-1 · {ontemStr}</div>
            <div style={{ fontSize:11, opacity:.65, marginTop:2 }}>Gerado em {hojeStr}</div>
            <div style={{ fontSize:12, opacity:.82, marginTop:6, display:"flex", gap:16, flexWrap:"wrap" }}>
              <span>Pré-compra SD abertas em {ontemSimples}: <b>{sdRows.length}</b></span>
              <span>RCs abertas em {ontemSimples}: <b>{rcRows.length}</b></span>
            </div>
          </div>
          <button onClick={onClose} style={{ background:"rgba(255,255,255,.12)", border:"1px solid rgba(255,255,255,.25)", color:"#fff", fontSize:18, cursor:"pointer", borderRadius:8, width:34, height:34, display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0, marginLeft:14 }}>×</button>
        </div>

        {/* Body — tabelas de preview */}
        <div style={{ overflow:"auto", padding:"20px 24px", flex:1 }}>

          {/* Tabela SD */}
          <div style={{ fontSize:12, fontWeight:700, color:"#0f3d5e", borderLeft:"4px solid #2e86c1", paddingLeft:8, marginBottom:10 }}>
            Pré-compra SD — abertas em {ontemSimples} ({sdRows.length})
          </div>
          {sdRows.length > 0 ? (
            <div style={{ overflowX:"auto", marginBottom:20 }}>
              <div className="gaq-table-scroll"><table style={tblStyle}>
                <thead>
                  <tr>{["Ticket SD","Abertura","Objeto","Status","Área Requisitante","Avaliador","Cliente"].map(h => <th key={h} style={thStyle}>{h}</th>)}</tr>
                </thead>
                <tbody>
                  {sdRows.map((r,i) => (
                    <tr key={r.ticket} style={{ background: i%2===0 ? "var(--gaq-bg-2)" : "var(--gaq-surface)" }}>
                      <td style={{ ...tdStyle, fontWeight:700, whiteSpace:"nowrap" }}>{r.ticket}</td>
                      <td style={{ ...tdStyle, whiteSpace:"nowrap" }}>{r.abertura}</td>
                      <td style={{ ...tdStyle, maxWidth:220 }}>{r.objeto}</td>
                      <td style={{ ...tdStyle, whiteSpace:"nowrap" }}>{r.status}</td>
                      <td style={tdStyle}>{r.area}</td>
                      <td style={{ ...tdStyle, whiteSpace:"nowrap" }}>{r.avaliador}</td>
                      <td style={{ ...tdStyle, whiteSpace:"nowrap" }}>{r.cliente}</td>
                    </tr>
                  ))}
                </tbody>
              </table></div>
            </div>
          ) : <div style={{ fontSize:11, color:"var(--gaq-text-3)", marginBottom:20 }}>Nenhum ticket aberto nessa data em andamento.</div>}

          {/* Tabela RC */}
          <div style={{ fontSize:12, fontWeight:700, color:"#0f3d5e", borderLeft:"4px solid #27ae60", paddingLeft:8, marginBottom:10 }}>
            RCs recebidas em {ontemSimples} ({rcRows.length})
          </div>
          {rcRows.length > 0 ? (
            <div style={{ overflowX:"auto" }}>
              <div className="gaq-table-scroll"><table style={tblStyle}>
                <thead>
                  <tr>{["Nº RC","Recebimento RC","Objeto","Status","Área Requisitante","Comprador"].map(h => <th key={h} style={thStyle}>{h}</th>)}</tr>
                </thead>
                <tbody>
                  {rcRows.map((r,i) => (
                    <tr key={r.numRC + i} style={{ background: i%2===0 ? "var(--gaq-bg-2)" : "var(--gaq-surface)" }}>
                      <td style={{ ...tdStyle, fontWeight:700, whiteSpace:"nowrap" }}>{r.numRC}</td>
                      <td style={{ ...tdStyle, whiteSpace:"nowrap" }}>{r.abertura}</td>
                      <td style={{ ...tdStyle, maxWidth:220 }}>{r.objeto}</td>
                      <td style={tdStyle}>{r.status}</td>
                      <td style={tdStyle}>{r.area}</td>
                      <td style={{ ...tdStyle, whiteSpace:"nowrap" }}>{r.comprador}</td>
                    </tr>
                  ))}
                </tbody>
              </table></div>
            </div>
          ) : <div style={{ fontSize:11, color:"var(--gaq-text-3)" }}>Nenhuma RC recebida nessa data.</div>}
        </div>

        {/* Footer — ações */}
        <div style={{ padding:"14px 24px", borderTop:"1px solid var(--gaq-line)", display:"flex", gap:10, alignItems:"center", flexShrink:0, background:"var(--gaq-bg-2)", borderRadius:"0 0 14px 14px" }}>
          <button onClick={copiarEEnviar}
            style={{ background: cpStatus === "copiado" ? "#059669" : "#0f3d5e", color:"#fff", border:"none", borderRadius:8, padding:"10px 20px", cursor:"pointer", fontSize:12, fontWeight:700, display:"flex", alignItems:"center", gap:7, transition:"all .2s" }}>
            {cpStatus === "copiado" ? "✓ Copiado! Outlook aberto — cole no corpo do e-mail" : cpStatus === "erro" ? "⚠ Copie manualmente e abra o Outlook" : "✉ Copiar tabelas e abrir Outlook"}
          </button>
          <div style={{ fontSize:10, color:"var(--gaq-text-3)", flex:1 }}>
            Copia o conteúdo HTML formatado para a área de transferência e abre o Outlook com o destinatário e assunto preenchidos. Cole (Ctrl+V) no corpo do e-mail.
          </div>
          <button onClick={onClose} style={{ background:"transparent", border:"1px solid var(--gaq-line)", borderRadius:8, padding:"9px 16px", cursor:"pointer", fontSize:11, color:"var(--gaq-text-3)" }}>Fechar</button>
        </div>
      </div>
    </ModalOverlay>
  );
};

const GAQ_SEXTA_TD = { padding:"5px 8px", fontSize:11, borderBottom:"1px solid var(--gaq-line)", verticalAlign:"top", color:"var(--gaq-text)", lineHeight:1.35 };

const GAQ_SEXTA_INPUT = { width:"100%", boxSizing:"border-box", fontSize:11, padding:"4px 6px", border:"1px solid var(--gaq-line)", borderRadius:6, background:"var(--gaq-bg-2)", color:"var(--gaq-text)" };

const GAQ_SEXTA_EDIT_VAZIO = {};

const GaqSextaLinhaEditor = React.memo(({ r, idx, edit, removido, onEdit, onToggleRemove }) => {
  const pk = r.ProcessKey;
  const tag = (r["PROJETO / EVENTO / AÇÃO"] || "").toString().trim();
  const corDu = r.diasTotais > 100 ? "#922b21" : r.diasTotais > 60 ? "#c0392b" : r.diasTotais > 30 ? "#e67e22" : "#27ae60";
  const corPrev = r.diasParaEntrega == null ? "var(--gaq-text-3)" : r.diasParaEntrega < 0 ? "#c0392b" : r.diasParaEntrega <= 15 ? "#e67e22" : "#27ae60";
  const valStatus    = edit.status    != null ? edit.status    : (r.status || "");
  const valStatusDet = edit.statusDet != null ? edit.statusDet : (r.statusDet || "");
  const valObs       = edit.obs       != null ? edit.obs       : "";
  const editado = valStatus !== (r.status || "") || valStatusDet !== (r.statusDet || "") || valObs.trim() !== "";
  return (
    <tr style={{ background: removido ? "rgba(192,57,43,.07)" : editado ? "rgba(243,156,18,.10)" : "transparent", opacity: removido ? .5 : 1 }}>
      <td style={{ ...GAQ_SEXTA_TD, whiteSpace:"nowrap" }}>
        <button onClick={() => onToggleRemove(pk)} title={removido ? "Reincluir este processo no PDF" : "Excluir este processo do PDF (só desta exportação)"}
          style={{ background: removido ? "#27ae6018" : "#c0392b12", color: removido ? "#27ae60" : "#c0392b", border:"1px solid " + (removido ? "#27ae6066" : "#c0392b44"), borderRadius:6, cursor:"pointer", fontSize:10, fontWeight:700, padding:"3px 8px", whiteSpace:"nowrap" }}>
          {removido ? "↩ Reincluir" : "✖"}
        </button>
      </td>
      <td style={{ ...GAQ_SEXTA_TD, color:"var(--gaq-text-3)", whiteSpace:"nowrap" }}>{idx}</td>
      <td style={{ ...GAQ_SEXTA_TD, whiteSpace:"nowrap", textDecoration: removido ? "line-through" : "none" }}>
        <div style={{ fontWeight:700 }}>{r.NumRC || "—"}</div>
        <div style={{ fontSize:10, color:"var(--gaq-text-3)" }}>{r.TicketSD || "—"}</div>
      </td>
      <td style={{ ...GAQ_SEXTA_TD, maxWidth:130 }}>{tag ? <span style={{ fontSize:10, fontWeight:700, color:"#8e44ad", background:"#8e44ad15", borderRadius:4, padding:"1px 6px" }}>{tag}</span> : <span style={{ color:"var(--gaq-text-3)" }}>—</span>}</td>
      <td style={{ ...GAQ_SEXTA_TD, maxWidth:300, textDecoration: removido ? "line-through" : "none" }}>{(r.Objeto || "—").slice(0, 120)}{(r.Objeto || "").length > 120 ? "…" : ""}</td>
      <td style={{ ...GAQ_SEXTA_TD, whiteSpace:"nowrap", fontWeight:700, color:corPrev }}>{r.dataEntrega ? r.dataEntrega.toLocaleDateString("pt-BR") : "—"}</td>
      <td style={{ ...GAQ_SEXTA_TD, whiteSpace:"nowrap", fontWeight:800, color:corDu }}>{r.diasTotais}</td>
      <td style={{ ...GAQ_SEXTA_TD, minWidth:140 }}>
        <input style={{ ...GAQ_SEXTA_INPUT, opacity: removido ? .5 : 1 }} value={valStatus} placeholder={r.status || "—"} disabled={removido}
          onChange={e => onEdit(pk, "status", e.target.value)} />
      </td>
      <td style={{ ...GAQ_SEXTA_TD, minWidth:180 }}>
        <input style={{ ...GAQ_SEXTA_INPUT, opacity: removido ? .5 : 1 }} value={valStatusDet} placeholder={r.statusDet || "—"} disabled={removido}
          onChange={e => onEdit(pk, "statusDet", e.target.value)} />
      </td>
      <td style={{ ...GAQ_SEXTA_TD, minWidth:200 }}>
        <input style={{ ...GAQ_SEXTA_INPUT, opacity: removido ? .5 : 1 }} value={valObs} placeholder="Observação pontual…" disabled={removido}
          onChange={e => onEdit(pk, "obs", e.target.value)} />
      </td>
    </tr>
  );
});

const GaqSextaEditorModal = ({ processos, area, meta, fBase, onClose }) => {
  const [busca, setBusca] = React.useState("");
  const [edits, setEdits] = React.useState({}); // ProcessKey -> {status?, statusDet?, obs?}
  const [removidos, setRemovidos] = React.useState({}); // ProcessKey -> true (fora deste PDF)

  const emA = React.useMemo(() =>
    (processos || []).filter(r => r.emA).slice().sort((a, b) => b.diasTotais - a.diasTotais),
  [processos]);

  const onEdit = React.useCallback((pk, campo, valor) => {
    setEdits(prev => ({ ...prev, [pk]: { ...(prev[pk] || {}), [campo]: valor } }));
  }, []);
  const onToggleRemove = React.useCallback(pk => {
    setRemovidos(prev => { const n = { ...prev }; if (n[pk]) delete n[pk]; else n[pk] = true; return n; });
  }, []);

  const q = busca.trim().toLowerCase();
  const visiveis = q
    ? emA.filter(r => [r.NumRC, r.TicketSD, r.NumProcesso, r.Objeto, r["PROJETO / EVENTO / AÇÃO"], r.Comprador, r["Área Requisitante"], r.statusDet, r.status]
        .map(v => (v || "").toString().toLowerCase()).join(" ").includes(q))
    : emA;

  // Overrides efetivos: só o que de fato difere do valor original da base
  const buildOverrides = () => {
    const overrides = {};
    emA.forEach(r => {
      const e = edits[r.ProcessKey];
      if (!e) return;
      const o = {};
      if (e.status != null && e.status.trim() && e.status.trim() !== (r.status || "")) o.status = e.status.trim();
      if (e.statusDet != null && e.statusDet.trim() && e.statusDet.trim() !== (r.statusDet || "")) o.statusDet = e.statusDet.trim();
      if (e.obs != null && e.obs.trim()) o.obs = e.obs.trim();
      if (Object.keys(o).length) overrides[r.ProcessKey] = o;
    });
    return overrides;
  };
  const editCount = Object.keys(buildOverrides()).length;
  const removeCount = Object.keys(removidos).length;
  // Trava de saída: qualquer digitação ou exclusão pendente exige confirmação
  const temAlteracoes = Object.keys(edits).length > 0 || removeCount > 0;
  const fechar = () => {
    if (temAlteracoes && !window.confirm("Você tem edições/exclusões feitas para este PDF que ainda não foi gerado.\n\nSair e descartar tudo?")) return;
    onClose();
  };

  const gerar = () => {
    // Excluídos saem do PDF inteiro (lista, KPIs e tabelas) — só desta exportação
    const procFiltrados = removeCount ? (processos || []).filter(r => !removidos[r.ProcessKey]) : processos;
    const html = gerarRelatorioGAQSexta({ processos: procFiltrados, area, meta, fBase, overrides: buildOverrides() });
    const win = window.open("", "_blank");
    win.document.write(html);
    win.document.close();
    onClose();
  };

  const thStyle = { padding:"7px 9px", background:"#0f3d5e", color:"#fff", fontSize:10, fontWeight:700, textAlign:"left", borderRight:"1px solid #1a5276", whiteSpace:"nowrap", position:"sticky", top:0, zIndex:1 };

  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(15,23,42,.6)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:2200, padding:16 }}>
      <div style={{ background:"var(--gaq-surface)", borderRadius:14, width:"min(1420px,98vw)", maxHeight:"94vh", display:"flex", flexDirection:"column", boxShadow:"0 12px 48px rgba(0,0,0,.25)", border:"1px solid var(--gaq-line)" }} onClick={e => e.stopPropagation()}>

        {/* Header */}
        <div style={{ padding:"16px 24px", background:"linear-gradient(135deg,#0f3d5e,#1a5276)", color:"#fff", borderRadius:"14px 14px 0 0", display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexShrink:0 }}>
          <div>
            <div style={{ fontSize:10, opacity:.72, textTransform:"uppercase", letterSpacing:".05em", marginBottom:4 }}>GAQ Sexta · Revisão antes do PDF</div>
            <div style={{ fontSize:17, fontWeight:700, letterSpacing:"-.01em" }}>📋 {area} — {emA.length} processo(s) em andamento</div>
            <div style={{ fontSize:11, opacity:.7, marginTop:4 }}>Ajuste Status / Status Det., inclua uma Observação ou exclua (✖) processos deste PDF. Vale só para esta exportação — a base não é alterada.</div>
          </div>
          <button onClick={fechar} style={{ background:"rgba(255,255,255,.12)", border:"1px solid rgba(255,255,255,.25)", color:"#fff", fontSize:18, cursor:"pointer", borderRadius:8, width:34, height:34, display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0, marginLeft:14 }}>×</button>
        </div>

        {/* Busca */}
        <div style={{ padding:"12px 24px", borderBottom:"1px solid var(--gaq-line)", display:"flex", gap:10, alignItems:"center", flexShrink:0 }}>
          <input value={busca} onChange={e => setBusca(e.target.value)} placeholder="🔍 Buscar por RC, pré-compra, processo, TAG, objeto, comprador, área ou status…"
            style={{ ...GAQ_SEXTA_INPUT, maxWidth:480, padding:"7px 10px", fontSize:12 }} />
          <div style={{ fontSize:11, color:"var(--gaq-text-3)" }}>{visiveis.length} de {emA.length} exibidos · ordem do PDF (maior aging primeiro)</div>
        </div>

        {/* Lista */}
        <div style={{ overflow:"auto", flex:1 }}>
          <div className="gaq-table-scroll"><table style={{ width:"100%", borderCollapse:"collapse" }}>
            <thead>
              <tr>
                {["PDF","#","RC / Pré-compra","TAG","Objeto","Prev. Entrega","d.u.","Status","Status Det.","Observação (só p/ este PDF)"].map(h => <th key={h} style={thStyle}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {visiveis.map((r, i) => (
                <GaqSextaLinhaEditor key={(r.ProcessKey || "") + "_" + i} r={r} idx={emA.indexOf(r) + 1}
                  edit={edits[r.ProcessKey] || GAQ_SEXTA_EDIT_VAZIO} removido={!!removidos[r.ProcessKey]}
                  onEdit={onEdit} onToggleRemove={onToggleRemove} />
              ))}
            </tbody>
          </table></div>
          {visiveis.length === 0 && <div style={{ textAlign:"center", padding:40, color:"var(--gaq-text-3)", fontSize:12 }}>Nenhum processo encontrado para "{busca}".</div>}
        </div>

        {/* Footer */}
        <div style={{ padding:"14px 24px", borderTop:"1px solid var(--gaq-line)", display:"flex", gap:10, alignItems:"center", flexShrink:0, background:"var(--gaq-bg-2)", borderRadius:"0 0 14px 14px" }}>
          <button onClick={gerar}
            style={{ background:"#0f3d5e", color:"#fff", border:"none", borderRadius:8, padding:"10px 20px", cursor:"pointer", fontSize:12, fontWeight:700, display:"flex", alignItems:"center", gap:7 }}>
            📄 Gerar PDF{(editCount > 0 || removeCount > 0) ? ` (${[editCount > 0 ? `${editCount} editada${editCount > 1 ? "s" : ""}` : "", removeCount > 0 ? `${removeCount} excluída${removeCount > 1 ? "s" : ""}` : ""].filter(Boolean).join(" · ")})` : ""}
          </button>
          <div style={{ fontSize:10, color:"var(--gaq-text-3)", flex:1 }}>
            Edições, observações e exclusões valem somente para o PDF gerado agora; a base permanece intacta. Para sair sem gerar, use Cancelar ou ✕ — será pedida confirmação se houver alterações.
          </div>
          <button onClick={fechar} style={{ background:"transparent", border:"1px solid var(--gaq-line)", borderRadius:8, padding:"9px 16px", cursor:"pointer", fontSize:11, color:"var(--gaq-text-3)" }}>Cancelar</button>
        </div>
      </div>
    </div>
  );
};

const PantanalModal = ({ proc, onClose }) => {
  const [cpTeams, setCpTeams] = React.useState(false);
  const [cpEmail, setCpEmail] = React.useState(false);
  const hoje = new Date(); hoje.setHours(0,0,0,0);
  // Aging oficial: desde Recebimento RC (ou pré-compra se não há RC)
  const diasAberto = proc.aberturaRC ? du(proc.aberturaRC, hoje) : (proc.aberturaSD ? du(proc.aberturaSD, hoje) : 0);
  const ageColor = diasAberto < 20 ? "#1e8449" : diasAberto < 50 ? "#e67e22" : "#c0392b";
  const ageBg   = diasAberto < 20 ? "#eafaf1" : diasAberto < 50 ? "#fef9e7" : "#fdf0f0";

  const respList = [
    proc.Comprador        && { label: "Comprador NCL",    value: proc.Comprador },
    proc.Avaliador        && { label: "Avaliador",         value: proc.Avaliador },
    proc.cplResp          && { label: "Resp. CPL",         value: proc.cplResp },
    proc.Pregoeiro        && { label: "Pregoeiro",         value: proc.Pregoeiro },
    proc.AnalistaContrato && { label: "Analista SCONT",    value: proc.AnalistaContrato },
    proc.AdvogadoResp     && { label: "Advogado",          value: proc.AdvogadoResp },
    proc.EmpresaContratada && { label: "Empresa Contratada", value: proc.EmpresaContratada },
  ].filter(Boolean);

  const mainResp = proc.Comprador || proc.Avaliador || proc.cplResp || "Responsável";
  const numProc  = proc.NumProcesso ? ` | Nº Processo: ${proc.NumProcesso}` : "";
  const dtAb     = proc.aberturaSD ? proc.aberturaSD.toLocaleDateString("pt-BR") : "—";
  const areaReq  = proc["Área Requisitante"] || "";

  const teamsMsg =
`Olá, ${mainResp}! 👋

Aqui é da equipe de Compras do SESC Pantanal. Identificamos o processo abaixo como uma oportunidade de adesão para nós:

📋 Pré-compra: ${proc.TicketSD || proc.NumRC || "—"}${numProc}
🏷️ Objeto: ${(proc.Objeto || "").slice(0, 180)}
🏛️ Área: ${areaReq || "—"}
📅 Abertura pré-compra: ${dtAb}

Gostaríamos de solicitar a inclusão do SESC Pantanal como aderente nesta contratação, junto ao Departamento Nacional – GAQ.

Poderia nos orientar sobre os próximos passos?

Grato(a) desde já! 🙏
SESC Pantanal – Equipe de Compras`;

  const emailSubject = `Solicitação de Adesão – Pré-compra ${proc.TicketSD || proc.NumRC || "—"} | ${(proc.Objeto || "").slice(0, 55)}`;
  const emailBody =
`Prezado(a) ${mainResp},

Meu nome é __________, da equipe de Compras do SESC Pantanal.

Identificamos o processo a seguir como uma oportunidade de adesão de nosso interesse:

  • Nº Pré-compra: ${proc.TicketSD || proc.NumRC || "—"}${numProc}
  • Objeto: ${proc.Objeto || "—"}
  • Área Requisitante: ${areaReq || "—"}
  • Abertura pré-compra: ${dtAb}
  • Fase atual: ${proc.faseAtual || "—"}

Solicitamos, gentilmente, a avaliação da viabilidade de inclusão do SESC Pantanal como aderente nesta contratação, a ser formalizada junto ao Departamento Nacional – GAQ.

Colocamo-nos à disposição para fornecer informações adicionais ou agendar uma conversa para tratar dos próximos passos.

Atenciosamente,

__________________________
SESC Pantanal – Equipe de Compras`;

  const copyText = async (text, setFn) => {
    try { await navigator.clipboard.writeText(text); } catch (_) {
      try { const ta = document.createElement("textarea"); ta.value = text;
        ta.style.cssText = "position:fixed;top:-9999px;left:-9999px;opacity:0;width:1px;height:1px;";
        document.body.appendChild(ta); ta.focus(); ta.select(); document.execCommand("copy"); document.body.removeChild(ta); } catch (_) {}
    }
    setFn(true); setTimeout(() => setFn(false), 2200);
  };

  const btnBase = { borderRadius: 8, padding: "10px 0", cursor: "pointer", fontSize: 12, fontWeight: 700, transition: "all .2s", display: "flex", alignItems: "center", justifyContent: "center", gap: 6, textDecoration: "none" };

  return (
    <ModalOverlay onDismiss={onClose} style={{ position: "fixed", inset: 0, background: "#0009", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 2200 }} onClick={onClose}>
      <div style={{ background: "var(--card)", borderRadius: 14, width: "min(780px,96vw)", maxHeight: "92vh", display: "flex", flexDirection: "column", boxShadow: "0 10px 50px #0007" }} onClick={e => e.stopPropagation()}>

        {/* Header */}
        <div style={{ padding: "18px 24px", background: "linear-gradient(135deg, #0b3d20 0%, #145a32 50%, #1e8449 100%)", color: "#fff", borderRadius: "14px 14px 0 0", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 11, opacity: .7, marginBottom: 4, letterSpacing: .5 }}>🌿 JANELA DE OPORTUNIDADE PANTANAL</div>
            <div style={{ fontWeight: 800, fontSize: 17, marginBottom: 6 }}>{proc.TicketSD || proc.NumRC || "—"} · {areaReq || "—"}</div>
            <div style={{ fontSize: 11, opacity: .85, lineHeight: 1.55, maxWidth: 540 }}>{(proc.Objeto || "Sem descrição").slice(0, 200)}{(proc.Objeto||"").length > 200 ? "…" : ""}</div>
            <div style={{ display: "flex", gap: 6, marginTop: 10, flexWrap: "wrap" }}>
              {proc.Modalidade && <span style={{ background: "rgba(255,255,255,0.2)", borderRadius: 5, padding: "2px 10px", fontSize: 10, fontWeight: 700 }}>{proc.Modalidade}</span>}
              <span style={{ background: ageBg, color: ageColor, borderRadius: 5, padding: "2px 10px", fontSize: 10, fontWeight: 800 }}>⏱ {diasAberto} d.u. {proc.aberturaRC ? "(desde RC)" : "(desde pré-compra)"}</span>
              {proc.faseAtual && proc.faseAtual !== "—" && <span style={{ background: "rgba(255,255,255,0.18)", borderRadius: 5, padding: "2px 10px", fontSize: 10 }}>📍 {proc.faseAtual}</span>}
            </div>
          </div>
          <button onClick={onClose} style={{ background: "rgba(255,255,255,0.15)", border: "1px solid rgba(255,255,255,0.3)", color: "#fff", fontSize: 18, cursor: "pointer", borderRadius: 8, width: 34, height: 34, display: "flex", alignItems: "center", justifyContent: "center", marginLeft: 14, flexShrink: 0 }}>×</button>
        </div>

        {/* Body */}
        <div style={{ overflow: "auto", padding: "20px 24px", flex: 1 }}>

          {/* Responsáveis */}
          <div style={{ marginBottom: 18 }}>
            <div style={{ fontSize: 10, fontWeight: 700, color: "#1e8449", textTransform: "uppercase", letterSpacing: .6, marginBottom: 10 }}>👥 Responsáveis pelo Processo</div>
            {respList.length > 0 ? (
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(190px,1fr))", gap: 8 }}>
                {respList.map((r, i) => (
                  <div key={i} style={{ background: "var(--card2)", borderRadius: 8, padding: "10px 14px", border: "1px solid var(--border2)" }}>
                    <div style={{ fontSize: 9, fontWeight: 700, color: "#1e8449", textTransform: "uppercase", letterSpacing: .5, marginBottom: 3 }}>{r.label}</div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: "var(--text)" }}>{r.value}</div>
                  </div>
                ))}
              </div>
            ) : <div style={{ color: "var(--text3)", fontSize: 12 }}>Nenhum responsável identificado neste processo.</div>}
          </div>

          {/* Info do processo */}
          <div style={{ marginBottom: 18 }}>
            <div style={{ fontSize: 10, fontWeight: 700, color: "#1e8449", textTransform: "uppercase", letterSpacing: .6, marginBottom: 10 }}>📋 Informações do Processo</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(160px,1fr))", gap: 8, fontSize: 12 }}>
              {[
                ["Nº Pré-compra", proc.TicketSD || proc.NumRC], ["Nº RC", proc.NumRC || "—"],
                ["Modalidade", proc.Modalidade], ["Área", proc["Área Requisitante"]],
                ["Abertura pré-compra", dtAb], [proc.aberturaRC ? "Em aberto (desde RC)" : "Em aberto (desde pré-compra)", `${diasAberto} d.u.`],
                ["Fase atual", proc.faseAtual || "—"], ["Nº Processo", proc.NumProcesso || "—"],
              ].map(([lbl, val]) => (
                <div key={lbl} style={{ background: "var(--card2)", borderRadius: 7, padding: "8px 12px", border: "1px solid var(--border2)" }}>
                  <div style={{ fontSize: 9, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", letterSpacing: .4, marginBottom: 3 }}>{lbl}</div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "var(--text)" }}>{val || "—"}</div>
                </div>
              ))}
            </div>
          </div>

          {/* Objeto completo */}
          {proc.Objeto && (
            <div style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 10, fontWeight: 700, color: "#1e8449", textTransform: "uppercase", letterSpacing: .6, marginBottom: 8 }}>🏷️ Objeto Completo</div>
              <div style={{ background: "var(--card2)", borderRadius: 8, padding: "12px 16px", fontSize: 12, color: "var(--text)", lineHeight: 1.65, border: "1px solid var(--border2)", whiteSpace: "pre-wrap", wordBreak: "break-word" }}>{proc.Objeto}</div>
            </div>
          )}

          {/* Templates */}
          <div style={{ background: "linear-gradient(135deg, #eafaf1 0%, #d5f5e3 100%)", borderRadius: 12, padding: "18px 20px", border: "2px solid #a9dfbf" }}>
            <div style={{ fontWeight: 800, fontSize: 14, color: "#145a32", marginBottom: 3 }}>📨 Templates de Solicitação de Adesão</div>
            <div style={{ fontSize: 11, color: "#1e8449", marginBottom: 18 }}>Mensagens prontas com os dados do processo. Revise antes de enviar se necessário.</div>

            {/* Teams */}
            <div style={{ marginBottom: 16 }}>
              <div style={{ fontSize: 10, fontWeight: 700, color: "#5558af", textTransform: "uppercase", letterSpacing: .5, marginBottom: 6 }}>💬 Microsoft Teams</div>
              <div style={{ background: "#fff", borderRadius: 8, padding: "10px 14px", fontSize: 11, color: "#333", lineHeight: 1.65, whiteSpace: "pre-wrap", border: "1px solid #c5cae9", marginBottom: 8, maxHeight: 130, overflow: "auto" }}>{teamsMsg}</div>
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={() => copyText(teamsMsg, setCpTeams)}
                  style={{ ...btnBase, flex: 1, background: cpTeams ? "#1e8449" : "#fff", color: cpTeams ? "#fff" : "#5558af", border: `2px solid ${cpTeams ? "#1e8449" : "#5558af"}` }}>
                  {cpTeams ? "✓ Copiado!" : "📋 Copiar mensagem"}
                </button>
                <a href={`https://teams.microsoft.com/l/chat/0/0?message=${encodeURIComponent(teamsMsg)}`} target="_blank" rel="noreferrer"
                  style={{ ...btnBase, flex: 1, background: "#5558af", color: "#fff", border: "none" }}>
                  💬 Abrir no Teams
                </a>
              </div>
            </div>

            {/* Email */}
            <div>
              <div style={{ fontSize: 10, fontWeight: 700, color: "#c0392b", textTransform: "uppercase", letterSpacing: .5, marginBottom: 6 }}>📧 E-mail</div>
              <div style={{ background: "#fff", borderRadius: 8, padding: "10px 14px", fontSize: 11, color: "#333", lineHeight: 1.65, whiteSpace: "pre-wrap", border: "1px solid #f1948a", marginBottom: 8, maxHeight: 130, overflow: "auto" }}>{emailBody}</div>
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={() => copyText(emailBody, setCpEmail)}
                  style={{ ...btnBase, flex: 1, background: cpEmail ? "#1e8449" : "#fff", color: cpEmail ? "#fff" : "#c0392b", border: `2px solid ${cpEmail ? "#1e8449" : "#c0392b"}` }}>
                  {cpEmail ? "✓ Copiado!" : "📋 Copiar e-mail"}
                </button>
                <a href={`mailto:?subject=${encodeURIComponent(emailSubject)}&body=${encodeURIComponent(emailBody)}`}
                  style={{ ...btnBase, flex: 1, background: "#c0392b", color: "#fff", border: "none" }}>
                  📧 Abrir no E-mail
                </a>
              </div>
            </div>
          </div>
        </div>
      </div>
    </ModalOverlay>
  );
};

const ServiceDeskAlertStrip = ({ summary, showLastInteraction = true }) => {
  if (!summary) return null;
  const { comQuem, acao, estado, diasDesdeUltimaInteracao, ultimaInteracaoDate } = summary;
  const hasLastInteraction = showLastInteraction && diasDesdeUltimaInteracao != null;
  const diasColor = diasDesdeUltimaInteracao == null ? "var(--gaq-text-3)"
    : diasDesdeUltimaInteracao > 10 ? "#dc2626"
    : diasDesdeUltimaInteracao > 5 ? "#c2410c"
    : "var(--gaq-text-2)";
  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: 0, alignItems: "stretch", margin: "8px 0 7px", border: "1px solid var(--gaq-line)", borderRadius: 8, background: "var(--gaq-bg-2)", overflow: "hidden" }}>
      <div style={{ flex: "0 1 210px", minWidth: 150, padding: "7px 11px", borderRight: "1px solid var(--gaq-line)" }}>
        <div style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".05em", textTransform: "uppercase", color: "var(--gaq-text-3)", marginBottom: 2 }}>Com quem está</div>
        <div style={{ fontSize: 12, fontWeight: 700, color: "var(--gaq-text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{comQuem}</div>
      </div>
      <div style={{ flex: "1 1 260px", minWidth: 200, padding: "7px 11px", borderRight: hasLastInteraction ? "1px solid var(--gaq-line)" : "none" }}>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center", marginBottom: 2 }}>
          <div style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".05em", textTransform: "uppercase", color: "var(--gaq-text-3)" }}>Ação necessária</div>
          <span style={{ fontSize: 9, fontWeight: 600, color: "var(--gaq-text-3)", background: "var(--gaq-surface)", borderRadius: 999, padding: "1px 7px", border: "1px solid var(--gaq-line)", whiteSpace: "nowrap" }}>{estado}</span>
        </div>
        <div style={{ fontSize: 12, fontWeight: 600, color: "var(--gaq-text-2)", lineHeight: 1.35 }}>{acao}</div>
      </div>
      {hasLastInteraction && (
        <div style={{ flex: "0 0 auto", padding: "7px 13px", minWidth: 110, textAlign: "center", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
          <div style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".05em", textTransform: "uppercase", color: "var(--gaq-text-3)", marginBottom: 2 }}>Últ. inter. SD</div>
          <div style={{ fontSize: 17, fontWeight: 800, color: diasColor, lineHeight: 1 }}>{diasDesdeUltimaInteracao}</div>
          <div style={{ fontSize: 9, color: "var(--gaq-text-3)", marginTop: 1 }}>d.u. atrás</div>
          {ultimaInteracaoDate && <div style={{ fontSize: 9, color: "var(--gaq-text-3)", marginTop: 2 }}>{ultimaInteracaoDate.toLocaleDateString("pt-BR")}</div>}
        </div>
      )}
    </div>
  );
};

const ServiceDeskHistoryModal = ({ ticket, record, onClose }) => {
  const summary = summarizeServiceDesk(ticket, record);
  if (!summary) return null;
  const { info, cliente, interacoes, last, estado, comQuem, acao, ultimaPendencia } = summary;
  return (
    <ModalOverlay onDismiss={onClose} style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.58)", zIndex: 2200, display: "flex", alignItems: "center", justifyContent: "center", padding: 18, fontFamily: "var(--gaq-font)", WebkitFontSmoothing: "antialiased", textRendering: "optimizeLegibility" }} onClick={onClose}>
      <div style={{ width: "min(820px,96vw)", maxHeight: "90vh", background: "var(--gaq-surface)", borderRadius: 14, boxShadow: "var(--gaq-shadow-2)", border: "1px solid var(--gaq-line)", overflow: "hidden", display: "flex", flexDirection: "column" }} onClick={e => e.stopPropagation()}>
        <div style={{ padding: "18px 22px", background: "linear-gradient(135deg,#0f3d5e,#1a5276)", color: "#fff", display: "flex", justifyContent: "space-between", gap: 14 }}>
          <div style={{ minWidth: 0 }}>
            <div style={{ fontSize: 10, fontWeight: 600, textTransform: "uppercase", opacity: .75, letterSpacing: ".04em", marginBottom: 4 }}>Histórico Service Desk</div>
            <div style={{ fontSize: 18, fontWeight: 700 }}>{ticket}</div>
            <div style={{ fontSize: 12, opacity: .86, marginTop: 4, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 620 }}>{info.titulo || last.assunto || "Chamado de pré-compra"}</div>
          </div>
          <button type="button" onClick={onClose} style={{ width: 32, height: 32, borderRadius: 8, border: "1px solid rgba(255,255,255,.25)", background: "rgba(255,255,255,.12)", color: "#fff", cursor: "pointer", fontSize: 18 }}>×</button>
        </div>
        <div style={{ padding: 18, overflow: "auto" }}>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 10, marginBottom: 14 }}>
            {[
              ["Status", estado],
              ["Com quem está", comQuem],
              ["Técnico", info.tecnico || "—"],
              ["Fila", info.fila || "—"],
              ["Cliente", [cliente.nome, cliente.sobrenome].filter(Boolean).join(" ") || "—"],
              ["Criado", info.criado || "—"],
            ].map(([label, value]) => (
              <div key={label} style={{ background: "var(--gaq-bg-2)", border: "1px solid var(--gaq-line)", borderRadius: 8, padding: "10px 12px" }}>
                <div className="gaq-eyebrow" style={{ fontSize: 9, marginBottom: 3 }}>{label}</div>
                <div style={{ fontSize: 12, fontWeight: 700, color: "var(--gaq-text)" }}>{value}</div>
              </div>
            ))}
          </div>
          <div style={{ background: "#fff7ed", border: "1px solid #fed7aa", color: "#9a3412", borderRadius: 10, padding: "11px 14px", marginBottom: 16 }}>
            <div style={{ fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: ".04em", marginBottom: 4 }}>Ação pendente</div>
            <div style={{ fontSize: 13, fontWeight: 600 }}>{acao}</div>
            {ultimaPendencia && <div style={{ fontSize: 11, marginTop: 6, lineHeight: 1.5 }}><b>Última interação:</b> {ultimaPendencia.slice(0, 320)}{ultimaPendencia.length > 320 ? "..." : ""}</div>}
          </div>
          <div className="gaq-eyebrow" style={{ marginBottom: 10 }}>Interações ({interacoes.length})</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {interacoes.map((it, idx) => (
              <div key={idx} style={{ border: "1px solid var(--gaq-line)", borderRadius: 10, padding: "11px 13px", background: idx === interacoes.length - 1 ? "rgba(46,134,193,.08)" : "var(--gaq-surface)" }}>
                <div style={{ display: "flex", justifyContent: "space-between", gap: 10, flexWrap: "wrap", marginBottom: 5 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "var(--gaq-text)" }}>{it.numero ? `#${it.numero} · ` : ""}{formatServiceDeskType(it.tipo)}</div>
                  <div style={{ fontSize: 10, color: "var(--gaq-text-3)" }}>{it.criado || "—"}</div>
                </div>
                <div style={{ fontSize: 10, color: "var(--gaq-text-3)", marginBottom: 7 }}>
                  {it.de || "—"}{it.para ? ` → ${it.para}` : ""}{it.assunto ? ` · ${formatServiceDeskSubject(it.assunto)}` : ""}
                </div>
                <div style={{ fontSize: 12, color: "var(--gaq-text-2)", lineHeight: 1.55, whiteSpace: "pre-wrap" }}>{it.mensagem || "—"}</div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </ModalOverlay>
  );
};

const TimelineModal = ({ proc: procRaw, onClose, onOpenTags, onUpdate, phaseIntervals, onToggleWatch, watchVersion, hidePredictions, serviceDeskData }) => {
  const [showHist, setShowHist] = useState(false);
  const [showServiceDeskHist, setShowServiceDeskHist] = useState(false);
  const [showPredictions, setShowPredictions] = useState(false);
  const [showObjeto, setShowObjeto] = useState(false);
  const [showCronMenu, setShowCronMenu] = useState(false);
  const [showResponsaveis, setShowResponsaveis] = useState(false);
  // Enriquece sempre na abertura (datas do log + status efetivo), independente da
  // lista que abriu o processo — evita duas versões do mesmo processo no portal.
  const proc = React.useMemo(() => enrichSD(procRaw, serviceDeskData), [procRaw, serviceDeskData]);
  const tags = TagsManager.getForProcess(proc.ProcessKey, proc.TicketSD);
  const isWatched = WatchlistManager.isWatched(proc.ProcessKey);
  // Ordena fases: a Indicação Analista é volátil e fica em posição cronológica por proc.
  // As demais (incl. Recebimento DJ) seguem a ordem fixa de TL_COLS.
  const ents = getOrderedTimelineEntries(proc);
  const _tlSla = calcSLAClassification(proc, phaseIntervals);
  const _tlBk = _tlSla ? SLA_BUCKETS[_tlSla.bucket] : null;
  const entregaDate = pd(proc.dataEntrega);
  const entregaDU = entregaDate ? businessDaysFromTodaySigned(entregaDate) : null;
  const entregaColor = entregaDU == null ? "#2e86c1" : entregaDU < 0 ? "#c0392b" : entregaDU <= 15 ? "#e67e22" : "#27ae60";
  const entregaLabel = entregaDU == null ? "" : entregaDU < 0 ? `${Math.abs(entregaDU)} d.u. vencida` : entregaDU === 0 ? "vence hoje" : `${entregaDU} d.u. restantes`;
  const serviceDeskRecord = getServiceDeskRecord(proc.TicketSD, serviceDeskData);
  const serviceDeskSummary = summarizeServiceDesk(proc.TicketSD, serviceDeskRecord);
  useEffect(() => { setShowPredictions(false); }, [proc?.ProcessKey]);

  // ── Calcular previsões para datas não preenchidas ──
  const predictions = useMemo(() => {
    if (!phaseIntervals || !proc.emA) return {};
    const mod = proc.Modalidade || "N/I";
    const modData = phaseIntervals.byMod[mod] || {};
    const globalData = phaseIntervals.global || {};
    const preds = {};
    // Modalidades sem CPL param no Envio Pedido/Suite (idx 9)
    const maxPredIdx = proc.isNonCPL ? IDX_ENVIO_PEDIDO : TL_COLS.length - 1;
    // Encontrar última data preenchida até o limite da modalidade
    let lastFilledIdx = -1, lastFilledDate = null;
    for (let i = 0; i < TL_COLS.length && i <= maxPredIdx; i++) {
      const d = pd(proc[TL_COLS[i][1]]);
      if (d) { lastFilledIdx = i; lastFilledDate = d; }
    }
    if (lastFilledIdx < 0 || !lastFilledDate) return {};
    // Projetar para frente a partir da última data preenchida
    let cursor = lastFilledDate;
    for (let i = lastFilledIdx; i < maxPredIdx; i++) {
      const [, keyA] = TL_COLS[i];
      const [, keyB] = TL_COLS[i + 1];
      const pair = keyA + "→" + keyB;
      const data = modData[pair] || globalData[pair];
      if (!data || !data.avg) continue;
      const nextDate = pd(proc[keyB]);
      if (nextDate) { cursor = nextDate; continue; } // já preenchida
      // Projetar: cursor + avg dias úteis (aproximar como dias corridos * 1.4)
      const projDate = new Date(cursor);
      projDate.setDate(projDate.getDate() + Math.round(data.avg * 1.4));
      preds[keyB] = { date: projDate, avg: data.avg, n: data.n, source: modData[pair] ? "modalidade" : "global" };
      cursor = projDate;
    }
    return preds;
  }, [proc, phaseIntervals]);

  // ── Calcular previsão total de conclusão ──
  const totalPrediction = useMemo(() => {
    if (!proc.emA || Object.keys(predictions).length === 0) return null;
    const lastPredKey = Object.keys(predictions).pop();
    const lastPred = predictions[lastPredKey];
    if (!lastPred) return null;
    const duRestantes = businessDaysFromToday(lastPred.date);
    return { date: lastPred.date, daysFromNow: duRestantes };
  }, [predictions, proc]);

  return (<>
    <ModalOverlay onDismiss={onClose} style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.55)", backdropFilter: "blur(2px)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000 }} onClick={onClose}>
      <div style={{ background: "var(--gaq-surface)", borderRadius: "var(--gaq-r-xl)", width: "min(720px,95vw)", maxHeight: "92vh", display: "flex", flexDirection: "column", boxShadow: "var(--gaq-shadow-2)", border: "1px solid var(--gaq-line)", overflow: "hidden" }} onClick={e => e.stopPropagation()}>
        {/* ── Header com gradiente ─────────────────────────────── */}
        <div style={{ padding: "20px 24px", background: "linear-gradient(135deg, #0a1f3d 0%, #1a3a6e 55%, #1a5276 100%)", color: "#fff", display: "flex", justifyContent: "space-between", gap: 12, alignItems: "flex-start" }}>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 10, fontWeight: 600, opacity: .72, letterSpacing: ".06em", textTransform: "uppercase", marginBottom: 4 }}>Timeline · Processo</div>
            <div style={{ fontWeight: 700, fontSize: 18, letterSpacing: "-0.02em", marginBottom: 10 }}>{proc.NumRC || proc.TicketSD || "Processo"}</div>
            {/* Chips de responsáveis organizados por função */}
            {(() => {
              const InfoChip = ({ label, value, color = "rgba(255,255,255,0.10)" }) => value ? (
                <div style={{ display: "flex", flexDirection: "column", background: color, borderRadius: 8, padding: "4px 10px", minWidth: 0, border: "1px solid rgba(255,255,255,0.12)" }}>
                  <span style={{ fontSize: 8, opacity: .72, fontWeight: 700, textTransform: "uppercase", letterSpacing: .5, whiteSpace: "nowrap" }}>{label}</span>
                  <span style={{ fontSize: 11, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 150 }}>{value}</span>
                </div>
              ) : null;
              const respFields = [
                proc.Comprador && { label: "Comprador", value: proc.Comprador },
                proc.Avaliador && { label: "Avaliador", value: proc.Avaliador },
                proc.cplResp && { label: "Resp. CPL", value: proc.cplResp },
                proc.Pregoeiro && { label: "Pregoeiro", value: proc.Pregoeiro },
                proc.AnalistaContrato && { label: "Analista (SCONT)", value: proc.AnalistaContrato },
                proc.AdvogadoResp && { label: "Advogado Responsável", value: proc.AdvogadoResp },
                proc.EmpresaContratada && { label: "Empresa Contratada", value: proc.EmpresaContratada },
              ].filter(Boolean);
              const hasExtra = respFields.length > 0;
              return (<>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 6 }}>
                  <InfoChip label="Modalidade" value={proc.Modalidade} />
                  <InfoChip label="Área" value={proc["Área Requisitante"]} />
                  {!showResponsaveis && respFields.slice(0, 2).map((f, i) => <InfoChip key={i} label={f.label} value={f.value} />)}
                  {hasExtra && (
                    <button onClick={() => setShowResponsaveis(s => !s)}
                      style={{ background: showResponsaveis ? "rgba(255,255,255,0.20)" : "rgba(255,255,255,0.08)", border: "1px solid rgba(255,255,255,0.20)", borderRadius: 8, padding: "4px 10px", cursor: "pointer", fontSize: 10, fontWeight: 700, color: "#fff", whiteSpace: "nowrap" }}>
                      {showResponsaveis ? "▲ Ocultar" : "▼ Responsáveis (" + respFields.length + ")"}
                    </button>
                  )}
                </div>
                {showResponsaveis && (
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 6 }}>
                    {respFields.map((f, i) => <InfoChip key={i} label={f.label} value={f.value} />)}
                  </div>
                )}
              </>);
            })()}
            <div onClick={() => proc.Objeto && proc.Objeto.length > 80 && setShowObjeto(s => !s)}
              style={{ fontSize: 11, opacity: .78, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 460, marginBottom: 6, marginTop: 4, lineHeight: 1.5, cursor: proc.Objeto && proc.Objeto.length > 80 ? "pointer" : "default" }}>
              {proc.Objeto ? proc.Objeto.slice(0, 100) : ""}
              {proc.Objeto && proc.Objeto.length > 100 && <span style={{ opacity: .65, marginLeft: 4, fontSize: 9 }}>(...clique p/ ver completo)</span>}
            </div>
            {/* Chips copiáveis: RC, pré-compra, Nº Processo */}
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 6 }}>
              {proc.NumRC && <CopyChip label="RC" value={proc.NumRC} color="#2e86c1" />}
              {proc.TicketSD && <CopyChip label="Pré-compra" value={proc.TicketSD} color="#27ae60" />}
              {proc.NumProcesso && <CopyChip label="Nº Processo" value={proc.NumProcesso} color="#8e44ad" />}
            </div>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {proc.status && <span style={{ background: "rgba(255,255,255,0.12)", borderRadius: 999, padding: "3px 10px", fontSize: 10, fontWeight: 600, border: "1px solid rgba(255,255,255,0.18)" }}>{proc.status}</span>}
              {proc.statusDet && <span style={{ background: "rgba(255,255,255,0.12)", borderRadius: 999, padding: "3px 10px", fontSize: 10, fontWeight: 600, border: "1px solid rgba(255,255,255,0.18)" }}>{proc.statusDet}</span>}
              {proc.NumPedidoSuite && <span style={{ background: "rgba(255,255,255,0.18)", borderRadius: 999, padding: "3px 10px", fontSize: 10, fontWeight: 700, border: "1px solid rgba(255,255,255,0.20)" }}>{proc.NumPedidoSuite}</span>}
              {_tlSla && _tlBk && <span style={{ background: _tlBk.color, borderRadius: 999, padding: "3px 10px", fontSize: 10, fontWeight: 700, color: "#fff" }}>{_tlBk.label} · {_tlSla.pctSLALabel}% SLA</span>}
              {entregaDate && <span style={{ background: "rgba(255,255,255,0.10)", border: `1px solid ${entregaColor}`, borderRadius: 999, padding: "3px 10px", fontSize: 10, fontWeight: 700 }}>
                Entrega prevista: {fmt(entregaDate)}{entregaLabel ? ` · ${entregaLabel}` : ""}
              </span>}
              {proc.diasParado > 15 && <span style={{ background: "#c0392b", borderRadius: 999, padding: "3px 10px", fontSize: 10, fontWeight: 700, color: "#fff" }}>{String.fromCharCode(9208)} Últ. mov. {proc.diasParado} d.u.</span>}
              {proc.isNonCPL && <span style={{ background: "rgba(255,255,255,0.10)", borderRadius: 999, padding: "3px 10px", fontSize: 10, fontStyle: "italic", border: "1px solid rgba(255,255,255,0.16)" }}>Sem etapas CPL</span>}
            </div>
            {tags.length > 0 && <div style={{ display: "flex", gap: 4, flexWrap: "wrap", marginTop: 8 }}>{tags.map(t => <TagBadge key={t} tag={t} small onDark />)}</div>}
          </div>
          <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 8, flexShrink: 0 }}>
            <button type="button" onClick={(e) => { e.preventDefault(); e.stopPropagation(); onClose && onClose(); }} title="Fechar"
              style={{ background: "rgba(255,255,255,0.10)", border: "1px solid rgba(255,255,255,0.20)", color: "#fff", fontSize: 18, lineHeight: 1, cursor: "pointer", width: 30, height: 30, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center" }}>×</button>
            <button type="button" onClick={(e) => { e.preventDefault(); e.stopPropagation(); onToggleWatch && onToggleWatch(proc); }} title={isWatched ? "Remover do acompanhamento" : "Acompanhar processo"}
              style={{ background: isWatched ? "rgba(243,156,18,0.95)" : "rgba(255,255,255,0.10)", border: isWatched ? "1px solid #f39c12" : "1px solid rgba(255,255,255,0.22)", borderRadius: 999, padding: "5px 12px", cursor: "pointer", fontSize: 11, color: "#fff", fontWeight: 600, transition: "all .2s", whiteSpace: "nowrap" }}>
              {isWatched ? "👁 Acompanhando" : "👁 Acompanhar"}
            </button>
          </div>
        </div>

        {/* ── Corpo ─────────────────────────────────────────────── */}
        <div style={{ overflow: "auto", padding: "18px 24px", flex: 1, background: "var(--gaq-bg)" }}>
          {/* Barra de ações */}
          <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
            <button key="tl-hist" type="button" onClick={(e) => { e.preventDefault(); e.stopPropagation(); setShowHist(true); }} className="gaq-btn"
              style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 600 }}>
              <Icon name="clock" size={13}/> Histórico
              {proc.Historico && <span style={{ background: "#2e86c1", color: "#fff", borderRadius: 999, padding: "1px 7px", fontSize: 9, fontWeight: 700 }}>NCL</span>}
              {proc.HistoricoCPL && <span style={{ background: "#8e44ad", color: "#fff", borderRadius: 999, padding: "1px 7px", fontSize: 9, fontWeight: 700 }}>CPL</span>}
              {proc.HistoricoScont && <span style={{ background: "#16a085", color: "#fff", borderRadius: 999, padding: "1px 7px", fontSize: 9, fontWeight: 700 }}>Scont</span>}
              {!proc.Historico && !proc.HistoricoCPL && !proc.HistoricoScont && <span style={{ opacity: .55, fontWeight: 400, fontSize: 11 }}>(sem)</span>}
            </button>
            <button key="tl-tags" type="button" onClick={(e) => { e.preventDefault(); e.stopPropagation(); onOpenTags(proc); }} className="gaq-btn"
              style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 600 }}>
              <Icon name="tag" size={13}/> Tags ({tags.length})
            </button>
            <button key="tl-obj" type="button" onClick={(e) => { e.preventDefault(); e.stopPropagation(); setShowObjeto(s => !s); }}
              className={showObjeto ? "gaq-btn is-primary" : "gaq-btn"}
              style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 600 }}>
              <Icon name="doc" size={13}/> Objeto {showObjeto ? "▲" : "▼"}
            </button>
            {!hidePredictions && proc.emA && Object.keys(predictions).length > 0 && (
              <button key="tl-pred" type="button" onClick={(e) => { e.preventDefault(); e.stopPropagation(); setShowPredictions(p => !p); }}
                className="gaq-btn"
                style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 600, background: showPredictions ? "#e67e22" : "var(--gaq-surface)", color: showPredictions ? "#fff" : "var(--gaq-text)", borderColor: showPredictions ? "#e67e22" : "var(--gaq-line)" }}>
                <Icon name="spark" size={13}/> Previsões {showPredictions ? "ON" : "OFF"}
              </button>
            )}
            <div key="tl-cron-wrap" style={{ position: "relative" }}>
              <button type="button" onClick={(e) => { e.preventDefault(); e.stopPropagation(); setShowCronMenu(m => !m); }} className="gaq-btn is-primary"
                style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 600 }}>
                <Icon name="calendar" size={13}/> Cronograma {showCronMenu ? "▲" : "▼"}
              </button>
              {showCronMenu && (
                <div style={{ position: "absolute", top: "calc(100% + 6px)", right: 0, background: "var(--gaq-surface)", borderRadius: 10, boxShadow: "var(--gaq-shadow-2)", zIndex: 50, minWidth: 200, overflow: "hidden", border: "1px solid var(--gaq-line)" }}>
                  <button type="button" onClick={(e) => { e.preventDefault(); e.stopPropagation(); exportCronogramaExcel(proc, phaseIntervals); setShowCronMenu(false); }}
                    style={{ width: "100%", display: "flex", alignItems: "center", gap: 10, padding: "11px 14px", background: "none", border: "none", borderBottom: "1px solid var(--gaq-line)", color: "var(--gaq-text)", cursor: "pointer", fontSize: 13, textAlign: "left" }}
                    onMouseEnter={e => e.currentTarget.style.background = "var(--gaq-bg-2)"}
                    onMouseLeave={e => e.currentTarget.style.background = "none"}>
                    <Icon name="download" size={14}/>
                    <div><div style={{ fontWeight: 700 }}>Exportar Excel</div><div style={{ fontSize: 10, color: "var(--gaq-text-3)" }}>Cronograma + dados do processo</div></div>
                  </button>
                  <button type="button" onClick={(e) => { e.preventDefault(); e.stopPropagation(); exportCronogramaPDF(proc, phaseIntervals); setShowCronMenu(false); }}
                    style={{ width: "100%", display: "flex", alignItems: "center", gap: 10, padding: "11px 14px", background: "none", border: "none", color: "var(--gaq-text)", cursor: "pointer", fontSize: 13, textAlign: "left" }}
                    onMouseEnter={e => e.currentTarget.style.background = "var(--gaq-bg-2)"}
                    onMouseLeave={e => e.currentTarget.style.background = "none"}>
                    <Icon name="doc" size={14}/>
                    <div><div style={{ fontWeight: 700 }}>Gerar PDF</div><div style={{ fontSize: 10, color: "var(--gaq-text-3)" }}>Cronograma para impressão</div></div>
                  </button>
                </div>
              )}
            </div>
          </div>

          {/* Objeto completo (expandível) */}
          {showObjeto && proc.Objeto && (
            <div className="gaq-card" style={{ padding: "14px 16px", marginBottom: 14, animation: "fadeSlideIn .2s ease-out" }}>
              <div className="gaq-eyebrow" style={{ marginBottom: 6 }}>Objeto da Contratação</div>
              <div style={{ fontSize: 13, color: "var(--gaq-text)", lineHeight: 1.6, whiteSpace: "pre-wrap", wordBreak: "break-word" }}>{proc.Objeto}</div>
            </div>
          )}

          {/* Previsão de conclusão — estilo clean */}
          {proc.emA && totalPrediction && showPredictions && (
            <div className="gaq-card" style={{ padding: "14px 18px", marginBottom: 14, background: "linear-gradient(180deg, #e67e221A 0%, var(--gaq-surface) 100%)", borderTop: "3px solid #e67e22" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                <div style={{ minWidth: 0 }}>
                  <div className="gaq-eyebrow" style={{ color: "#d35400", marginBottom: 4 }}>Previsão de Conclusão</div>
                  <div style={{ fontSize: 12, color: "var(--gaq-text-2)", lineHeight: 1.55 }}>
                    Baseada na média histórica de processos na modalidade <b style={{ color: "var(--gaq-text)" }}>{proc.Modalidade || "N/I"}</b> (todos com ambas datas preenchidas).
                  </div>
                </div>
                <div style={{ textAlign: "right", flexShrink: 0 }}>
                  <div className="gaq-num" style={{ fontSize: 22, fontWeight: 800, color: "#d35400", letterSpacing: "-0.02em" }}>~{totalPrediction.date.toLocaleDateString("pt-BR")}</div>
                  <div style={{ fontSize: 11, color: "var(--gaq-text-3)", marginTop: 2 }}>≈ {totalPrediction.daysFromNow} d.u. a partir de hoje</div>
                </div>
              </div>
            </div>
          )}

          {/* Suite Sesc — colunas novas da planilha (lookup por chave normalizada
              p/ tolerar variação de caixa/acentos no cabeçalho do Excel) */}
          {(() => {
            const findSuiteKey = (id) => Object.keys(proc).find(k => nrm(k).replace(/[\s_\-]/g, "") === id);
            const kId = findSuiteKey("identificadorsuitesesc");
            const kSt = findSuiteKey("statussuitesesc");
            const kHab = findSuiteKey("habilitadaemsuitesesc");
            const kResp = findSuiteKey("responsabilidadesuitesesc");
            if (!kId && !kSt && !kHab && !kResp) return null; // colunas ainda não existem no dados.js
            const txt = (v) => (v == null ? "" : v.toString().trim());
            const vId = txt(kId && proc[kId]);
            const vSt = txt(kSt && proc[kSt]);
            const dHab = kHab ? pd(proc[kHab]) : null;
            const vHab = dHab ? fmt(dHab) : txt(kHab && proc[kHab]);
            const vResp = txt(kResp && proc[kResp]);
            const Campo = ({ label, value }) => (
              <div style={{ background: "var(--gaq-bg-2)", borderRadius: 8, padding: "6px 12px", minWidth: 0 }}>
                <div style={{ fontSize: 9, fontWeight: 700, color: "var(--gaq-text-3)", textTransform: "uppercase", letterSpacing: ".04em", marginBottom: 2 }}>{label}</div>
                <div style={{ fontSize: 12, fontWeight: 600, color: "var(--gaq-text)", wordBreak: "break-word" }}>{value || "—"}</div>
              </div>
            );
            return (
              <div className="gaq-card" style={{ padding: "14px 18px", marginBottom: 14, borderTop: "3px solid #16a085" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", marginBottom: 8 }}>
                  <div className="gaq-eyebrow" style={{ color: "#16a085" }}>Suite Sesc</div>
                  {vId ? <CopyChip label="Identificador" value={vId} color="#16a085" solid />
                       : <span style={{ fontSize: 10, color: "var(--gaq-text-3)" }}>Identificador: —</span>}
                </div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  <Campo label="Status" value={vSt} />
                  <Campo label="Habilitada em" value={vHab} />
                  <Campo label="Responsabilidade" value={vResp} />
                </div>
              </div>
            );
          })()}

          {/* Linha do tempo das fases */}
          <div className="gaq-card" style={{ padding: "16px 18px 8px" }}>
            <div className="gaq-eyebrow" style={{ marginBottom: 12 }}>Etapas do processo</div>
            {serviceDeskSummary && (
              <div onClick={() => setShowServiceDeskHist(true)}
                style={{ display: "flex", minHeight: 68, cursor: "pointer" }}
                title="Abrir histórico do Service Desk">
                <div style={{ display: "flex", flexDirection: "column", alignItems: "center", width: 28, flexShrink: 0 }}>
                  <div style={{ width: 14, height: 14, borderRadius: "50%", marginTop: 4, flexShrink: 0, background: "#f97316", border: "2px solid #f97316", boxShadow: "0 0 0 3px #f9731622" }} />
                  <div style={{ width: 2, flex: 1, minHeight: 22, background: "var(--gaq-line)" }} />
                </div>
                <div style={{ flex: 1, paddingBottom: 14, paddingLeft: 12 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "flex-start" }}>
                    <div style={{ minWidth: 0, flex: 1 }}>
                      <div style={{ fontSize: 10, fontWeight: 700, color: "#ea580c", textTransform: "uppercase", letterSpacing: ".03em", marginBottom: 3, display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                        <span>Service Desk · Pré-compra {proc.TicketSD}</span>
                        <span style={{ fontSize: 9, fontWeight: 600, color: "#fff", background: "#f97316", borderRadius: 999, padding: "1px 7px", textTransform: "none", letterSpacing: 0 }}>{serviceDeskSummary.estado}</span>
                      </div>
                      <div style={{ fontSize: 13, fontWeight: 600, color: "var(--gaq-text)", marginBottom: 3 }}>{serviceDeskSummary.acao}</div>
                      <div style={{ fontSize: 11, color: "var(--gaq-text-2)", lineHeight: 1.45 }}>
                        Com quem está: <b>{serviceDeskSummary.comQuem}</b>
                        {serviceDeskSummary.ultimaPendencia ? <> · Última interação: {serviceDeskSummary.ultimaPendencia.slice(0, 150)}{serviceDeskSummary.ultimaPendencia.length > 150 ? "..." : ""}</> : null}
                      </div>
                    </div>
                    <div style={{ textAlign: "right", flexShrink: 0 }}>
                      <div style={{ fontSize: 8, color: "var(--gaq-text-3)", textTransform: "uppercase", letterSpacing: ".04em" }}>histórico</div>
                      <div style={{ fontSize: 13, fontWeight: 700, color: "#ea580c" }}>{serviceDeskSummary.interacoes.length} inter.</div>
                    </div>
                  </div>
                </div>
              </div>
            )}
            {ents.map((en, idx) => {
              let prev = null;
              for (let i = idx - 1; i >= 0; i--) { if (ents[i].date) { prev = ents[i].date; break; } }
              const intv = (en.date && prev) ? du(prev, en.date) : null;
              const pred = (!en.date && showPredictions) ? predictions[en.key] : null;
              const isPred = !!pred && !en.date;
              const isLast = idx === ents.length - 1;
              const phaseColor = en.isCPL ? "#8e44ad" : "#1a5276";
              return (
                <div key={en.label} style={{ display: "flex", minHeight: 54 }}>
                  <div style={{ display: "flex", flexDirection: "column", alignItems: "center", width: 28, flexShrink: 0 }}>
                    <div style={{ width: 14, height: 14, borderRadius: "50%", marginTop: 4, flexShrink: 0,
                      background: en.date ? phaseColor : isPred ? "#e67e2226" : "var(--gaq-bg-2)",
                      border: `2px ${isPred ? "dashed" : "solid"} ${en.date ? phaseColor : isPred ? "#e67e22" : "var(--gaq-line)"}`,
                      boxShadow: en.date ? `0 0 0 3px ${phaseColor}1A` : "none" }} />
                    {!isLast && <div style={{ width: 2, flex: 1, minHeight: 22,
                      background: isPred ? "transparent" : "var(--gaq-line)",
                      borderLeft: isPred ? "2px dashed #e67e2266" : "none" }} />}
                  </div>
                  <div style={{ flex: 1, paddingBottom: 14, paddingLeft: 12 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 10, alignItems: "flex-start" }}>
                      <div style={{ minWidth: 0, flex: 1 }}>
                        <div style={{ fontSize: 10, fontWeight: 700, color: en.date ? (en.isCPL ? "#8e44ad" : "var(--gaq-text-3)") : isPred ? "#e67e22" : "var(--gaq-text-3)", textTransform: "uppercase", letterSpacing: ".04em", marginBottom: 3, display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                          <span>{en.label}</span>
                          {PHASE_RESP[en.key] && <span style={{ fontSize: 9, fontWeight: 700, color: en.date ? "#fff" : "var(--gaq-text-3)", background: en.date ? phaseColor : "var(--gaq-bg-2)", border: en.date ? "none" : "1px solid var(--gaq-line)", borderRadius: 999, padding: "1px 7px", textTransform: "none", letterSpacing: 0 }}>{PHASE_RESP[en.key]}</span>}
                        </div>
                        {en.date ? (
                          <div>
                            <div className="gaq-num" style={{ fontSize: 14, fontWeight: 700, color: "var(--gaq-text)", letterSpacing: "-0.01em" }}>{fmt(en.date)}</div>
                            {en.fromHistory && <span style={{ fontSize: 9, color: "var(--gaq-text-3)", fontStyle: "italic", letterSpacing: ".02em" }}>⊙ hist. SD</span>}
                          </div>
                        ) : isPred ? (
                          <div style={{ fontSize: 13, fontWeight: 600, color: "#e67e22", fontStyle: "italic", display: "flex", alignItems: "baseline", gap: 6, flexWrap: "wrap" }}>
                            <span>~{pred.date.toLocaleDateString("pt-BR")}</span>
                            <span style={{ fontSize: 10, fontWeight: 500, color: "var(--gaq-text-3)", fontStyle: "normal" }}>
                              previsão ({pred.source === "modalidade" ? proc.Modalidade : "geral"} · ~{pred.avg} d.u. · {pred.n} amostras)
                            </span>
                          </div>
                        ) : (
                          <div style={{ fontSize: 13, fontWeight: 400, color: "var(--gaq-text-3)" }}>—</div>
                        )}
                      </div>
                      {intv !== null && (
                        <div style={{ textAlign: "right", flexShrink: 0 }}>
                          <div style={{ fontSize: 8, color: "var(--gaq-text-3)", textTransform: "uppercase", letterSpacing: ".04em" }}>intervalo</div>
                          <div className="gaq-num" style={{ fontSize: 13, fontWeight: 700, color: intv > 50 ? "#c0392b" : intv > 20 ? "#e67e22" : "#27ae60" }}>+{intv} d.u.</div>
                        </div>
                      )}
                      {isPred && (
                        <div style={{ textAlign: "right", flexShrink: 0 }}>
                          <div style={{ fontSize: 8, color: "#e67e22", textTransform: "uppercase", letterSpacing: ".04em" }}>estimado</div>
                          <div className="gaq-num" style={{ fontSize: 13, fontWeight: 600, color: "#e67e22" }}>~{pred.avg} d.u.</div>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>

          {/* Legenda */}
          <div style={{ marginTop: 12, padding: "10px 14px", background: "var(--gaq-bg-2)", border: "1px solid var(--gaq-line)", borderRadius: 10, fontSize: 11, color: "var(--gaq-text-3)", lineHeight: 1.55 }}>
            <b style={{ color: "var(--gaq-text-2)" }}>Legenda:</b>
            <span style={{ marginLeft: 6 }}>d.u. = dias úteis</span>
            <span style={{ margin: "0 6px" }}>·</span>
            <span style={{ color: "#8e44ad", fontWeight: 600 }}>roxo</span> = etapas CPL
            <span style={{ margin: "0 6px" }}>·</span>
            <span>cinza = não preenchida</span>
            <span style={{ margin: "0 6px" }}>·</span>
            <span style={{ background: "#1a5276", color: "#fff", borderRadius: 999, padding: "1px 7px", fontSize: 9, fontWeight: 700 }}>badge</span> = área responsável
            {showPredictions && Object.keys(predictions).length > 0 && (<>
              <span style={{ margin: "0 6px" }}>·</span>
              <span style={{ color: "#e67e22", fontWeight: 600 }}>laranja tracejado</span> = previsão baseada em histórico
            </>)}
          </div>
        </div>
      </div>
    </ModalOverlay>
    {showHist && <HistoricoModal proc={proc} onClose={() => setShowHist(false)} />}
    {showServiceDeskHist && serviceDeskRecord && <ServiceDeskHistoryModal ticket={proc.TicketSD} record={serviceDeskRecord} onClose={() => setShowServiceDeskHist(false)} />}
  </>);
};

const ProcCard = ({ r, tipo, onClickComp, onClickProc, onOpenTags, tagVersion, onToggleWatch, watchVersion }) => {
  const isNCL = tipo === "ncl", dias = isNCL ? r.diasTotaisGestao : r.diasCpl;
  const cor = isNCL ? (dias > 100 ? "#922b21" : dias > 70 ? "#c0392b" : "#e67e22") : (dias > 100 ? "#5b2c6f" : dias > 70 ? "#8e44ad" : "#a569bd");
  const tags = TagsManager.getForProcess(r.ProcessKey, r.TicketSD);
  const watched = WatchlistManager.isWatched(r.ProcessKey);
  const rAtivoColor = r.faseSubarea === "CPL" ? "#8e44ad" : r.faseSubarea === "Scont" ? "#16a085" : "#2e86c1";
  return (
    <div className="alert-card-modern" style={{ padding: "12px 14px", marginBottom: 8, borderLeft: `5px solid ${cor}`, outline: watched ? "2px solid #f39c1288" : "none" }}>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 10 }}>
        <div style={{ flex: 1, minWidth: 0 }}>
          {/* Linha 1: chips de identificação */}
          <div style={{ display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center", marginBottom: 4 }}>
            {watched && <span style={{ fontSize: 11, color: "#f39c12" }} title="Na watchlist">👁</span>}
            <span className="list-chip-cta" style={{ fontSize: 9, color: "var(--text3)" }}>{r.ProcessKey || "—"}</span>
            <span className="list-chip-cta" style={{ cursor: "pointer" }} onClick={() => onClickProc(r)}>{r.NumRC || "—"}</span>
            {r.TicketSD && <span className="list-chip-cta" style={{ color: "#27ae60" }}>Pré-compra: {r.TicketSD}</span>}
            {r["Área Requisitante"] && <span className="list-chip-cta" style={{ color: "#2e86c1" }}>{r["Área Requisitante"]}</span>}
            <span className="list-chip-cta" style={{ color: "#2e86c1" }}>{r.Modalidade || "—"}</span>
            {r.isLegado && <span style={{ fontSize: 9, background: "#ecf0f1", color: "#7f8c8d", borderRadius: 4, padding: "1px 6px", fontWeight: 700 }}>📦 Legado</span>}
            <span onClick={(e) => { e.stopPropagation(); onOpenTags(r); }} className="mini-link-btn" title="Gerenciar tags">#tag</span>
            <span onClick={(e) => { e.stopPropagation(); onToggleWatch && onToggleWatch(r); }} className="mini-link-btn" style={{ color: watched ? "#f39c12" : "var(--text3)", borderColor: watched ? "#f39c1266" : "var(--border2)", background: watched ? "#fff7e6" : "var(--card2)" }} title={watched ? "Remover da watchlist" : "Acompanhar"}>{watched ? "★" : "☆"}</span>
          </div>
          {/* Linha 2: responsável ativo (destaque) + fase + statusDet */}
          <div style={{ display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center", marginBottom: 3 }}>
            <span style={{ fontWeight: 800, fontSize: 12, color: rAtivoColor }}>◉ {r.respAtivo || "—"}</span>
            {r.faseSubarea && <span style={{ fontSize: 9, background: `${rAtivoColor}18`, color: rAtivoColor, borderRadius: 4, padding: "1px 6px", fontWeight: 600 }}>{r.faseSubarea}</span>}
            {r.statusDet && <span className="soft-chip" style={{ fontSize: 9, color: "#8e44ad" }}>{r.statusDet}</span>}
            {r.faseAtual && r.faseAtual !== "—" && <span className="soft-chip" style={{ fontSize: 9 }}>{r.faseAtual}</span>}
          </div>
          {/* Linha 3: objeto */}
          <div style={{ fontSize: 11, color: "var(--text)", marginBottom: 3, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{(r.Objeto || "").slice(0, 100) || "—"}</div>
          {/* Linha 4: todos os responsáveis */}
          <div style={{ display: "flex", gap: 4, flexWrap: "wrap", alignItems: "center", fontSize: 9, color: "var(--text3)" }}>
            {r.Comprador && <span>Comp: <b style={{ color: "var(--text2)" }}>{r.Comprador}</b></span>}
            {r.Avaliador && <span>· Aval: <b style={{ color: "var(--text2)" }}>{r.Avaliador}</b></span>}
            {r.cplResp && <span>· CPL: <b style={{ color: "#8e44ad" }}>{r.cplResp}</b></span>}
            {r.Pregoeiro && r.Pregoeiro !== r.cplResp && <span>· Preg: <b style={{ color: "#8e44ad" }}>{r.Pregoeiro}</b></span>}
            {r.AnalistaContrato && <span>· Scont: <b style={{ color: "#16a085" }}>{r.AnalistaContrato}</b></span>}
            {r.AdvogadoResp && <span>· Adv: <b style={{ color: "#16a085" }}>{r.AdvogadoResp}</b></span>}
          </div>
          <div style={{ marginTop: 3 }}>
            {tags.map(t => <TagBadge key={t} tag={t} small />)}
          </div>
        </div>
        <div style={{ textAlign: "center", minWidth: 58, flexShrink: 0 }}>
          <div style={{ fontSize: 22, fontWeight: 800, color: cor }}>{dias}</div>
          <div style={{ fontSize: 9, color: "var(--text3)" }}>d.u. {tipo.toUpperCase()}</div>
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
};

const MiniTimeline = ({ r }) => {
  if (!r || r.faseAtualIdx < 0) return null;
  const total = r.isNonCPL ? (IDX_ENVIO_PEDIDO + 1) : TL_COLS.length;
  const cur = Math.min(r.faseAtualIdx, total - 1);
  const pct = Math.round((cur / Math.max(total - 1, 1)) * 100);
  const color = r.faseSubarea === "CPL" ? "#8e44ad" : r.faseSubarea === "Scont" ? "#16a085" : "#2e86c1";
  return (
    <div style={{ marginTop: 6 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 2 }}>
        <span style={{ fontSize: 8, color: "var(--text3)" }}>Início</span>
        <span style={{ fontSize: 8, color, fontWeight: 700 }}>{r.faseAtual && r.faseAtual !== "—" ? r.faseAtual : "—"}</span>
        <span style={{ fontSize: 8, color: "var(--text3)" }}>Conclusão</span>
      </div>
      <div style={{ position: "relative", height: 5, borderRadius: 3, background: "var(--border2)", overflow: "hidden" }}>
        <div style={{ position: "absolute", left: 0, top: 0, height: "100%", width: pct + "%", background: color, borderRadius: 3 }} />
      </div>
      <div style={{ textAlign: "right", fontSize: 8, color: "var(--text3)", marginTop: 1 }}>{pct}% concluído</div>
    </div>
  );
};

const PatternCard = ({ p, accent = "#e67e22" }) => (
  <div className="alert-card-modern" style={{ padding: "14px 16px", marginBottom: 10, borderLeft: `5px solid ${p.count >= 5 ? "#c0392b" : p.count >= 3 ? accent : "#f0c030"}` }}>
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
      <div style={{ flex: 1 }}>
        <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap", marginBottom: 5 }}>
          <span style={{ background: p.count >= 5 ? "#c0392b" : p.count >= 3 ? accent : "#f39c12", color: "#fff", borderRadius: 999, padding: "2px 12px", fontSize: 12, fontWeight: 800 }}>"{p.word}"</span>
          {p.cancelados > 0 && <span className="soft-chip" style={{ color: "#c0392b" }}>{p.cancelados} cancelados</span>}
          {p.fracassados > 0 && <span className="soft-chip" style={{ color: "#8e44ad" }}>{p.fracassados} fracassados</span>}
          {p.avgLT && <span className="soft-chip" style={{ color: "#2e86c1" }}>⏱ {p.avgLT} d.u.</span>}
        </div>
        {p.areas && p.areas.length > 0 && <div style={{ fontSize: 10, color: "var(--text3)", marginBottom: 3 }}>Áreas: {p.areas.join(" · ")}</div>}
        {p.compradores && p.compradores.length > 0 && <div style={{ fontSize: 10, color: "var(--text3)", marginBottom: 3 }}>Responsáveis: {p.compradores.join(" · ")}</div>}
        {p.exemplos && p.exemplos.length > 0 && <div style={{ marginBottom: 4 }}>
          {p.exemplos.slice(0, 2).map((ex, ei) => <div key={ei} style={{ fontSize: 11, color: "var(--text2)", background: "var(--card2)", borderRadius: 4, padding: "2px 7px", marginBottom: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>"{ex.slice(0, 130)}"</div>)}
        </div>}
        <div style={{ fontSize: 11, color: "var(--text2)" }}><b>RCs:</b> {p.nums.slice(0, 6).join(" · ")}{p.nums.length > 6 && <span style={{ color: "var(--text3)" }}> +{p.nums.length - 6}</span>}</div>
      </div>
      <div style={{ textAlign: "center", minWidth: 50, flexShrink: 0 }}>
        <div style={{ fontSize: 28, fontWeight: 800, color: p.count >= 5 ? "#c0392b" : p.count >= 3 ? "#e67e22" : "#f39c12" }}>{p.count}</div>
        <div style={{ fontSize: 9, color: "var(--text3)" }}>casos</div>
      </div>
    </div>
  </div>
);

const RecCard = ({ r, i, pool, dem, onTempAssign }) => {
  const poolColor = pool === "cpl" ? "#8e44ad" : "#1a5276";
  const accents = pool === "cpl" ? ["#8e44ad","#a569bd","#c39bd3"] : ["#1e8449","#2e86c1","#85c1e9"];
  const areaLabel = dem && dem.area ? dem.area : "—";
  // Justificativa textual: por que esta pessoa é uma boa pista
  const motivos = [];
  if (r.simArea >= 3) motivos.push(`${r.simArea} processos já tocados na mesma área`);
  if (r.simObj >= 1) motivos.push(`${r.simObj} casos com objeto semelhante`);
  if (!r.ignMod && r.simMod >= 3) motivos.push(`${r.simMod} processos na mesma modalidade`);
  if (r.totSD >= 15) motivos.push(`alto volume de SDs absorvidos (${r.totSD})`);
  if (r.throughput >= 70) motivos.push(`alta vazão de entregas (${r.throughput}% throughput)`);
  if (r.cargaArea <= 2 && dem && dem.area) motivos.push(`baixa carga atual em "${areaLabel}" (${r.cargaArea})`);
  if (r.cargaAnaliseSD <= 1) motivos.push(`pouco backlog em análise SD (${r.cargaAnaliseSD})`);
  if (r.med != null && r.med <= 40) motivos.push(`lead médio rápido (${r.med} d.u.)`);
  if (r.mediaTocadosPool > 0 && (r.tocadosAno || 0) < r.mediaTocadosPool * 0.9) motivos.push(`abaixo da média do pool no ano (${r.tocadosAno} vs ~${Math.round(r.mediaTocadosPool)})`);
  if ((r.conclAno || r.concl2026 || 0) >= 5) motivos.push(`${r.conclAno || r.concl2026} concluído(s) no ano`);
  if (pool !== "cpl" && (r.sdAtual || 0) <= 2) motivos.push(`baixo volume de SD aberto agora (${r.sdAtual || 0})`);
  const subScores = [
    ["Expertise", r.sA, 20, false],
    ["Performance", r.sB, 8, false],
    ["Carteira", r.sC, 24, false],
    ["Complex.", r.sD, 4, r.ignMod],
    ["Distrib. Justa", r.sE, 18, false],
    ["Vol. SD", r.sF, 6, false],
    ["Ano atual", r.sG, 14, false],
  ];
  return (
    <div className="alert-card-modern" style={{ padding: 18, marginBottom: 10, borderLeft: `5px solid ${accents[i] || accents[0]}` }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
        {i < 3
          ? <span style={{ fontSize: 18 }}>{"🥇🥈🥉"[i]}</span>
          : <span style={{ fontSize: 12, fontWeight: 700, color: "var(--text3)", background: "var(--card2)", borderRadius: 6, padding: "2px 8px" }}>#{i + 1}</span>}
        <span style={{ fontWeight: 700, fontSize: 15, color: "var(--text)" }}>{r.comp}</span>
        <BadgeScore score={r.total} />
        <span style={{ background: poolColor + "22", color: poolColor, borderRadius: 6, padding: "1px 8px", fontSize: 10, fontWeight: 700 }}>{pool.toUpperCase()}</span>
        {r.ignMod && <span style={{ fontSize: 10, background: "#ecf0f1", border: "1px solid #bdc3c7", color: "#7f8c8d", borderRadius: 4, padding: "1px 7px" }}>Modalidade ignorada</span>}
        {!r.ignMod && r.taxaCF > 25 && <span style={{ fontSize: 11, background: "#fef9e7", border: "1px solid #e67e22", color: "#d35400", borderRadius: 4, padding: "1px 7px" }}>⚠️ {r.taxaCF}% canc./frac.</span>}
        {onTempAssign && (
          <button type="button" onClick={() => onTempAssign(r.comp, pool)}
            style={{ marginLeft: "auto", background: poolColor, color: "#fff", border: "none", borderRadius: 6, padding: "4px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
            + temp.
          </button>
        )}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)", gap: 6, marginBottom: 8 }}>
        {subScores.map(([l, v, mx, dimmed]) => (
          <div key={l} style={{ background: "var(--card2)", borderRadius: 5, padding: "7px 6px", textAlign: "center", opacity: dimmed ? 0.4 : 1 }}>
            <div style={{ fontSize: 9, color: "var(--text3)" }}>{l} (/{mx})</div>
            <div style={{ fontWeight: 700, color: !dimmed && v < 0 ? "#c0392b" : poolColor, fontSize: 15 }} title={v < 0 ? "Sobrecarga: carga atual acima do saudável" : undefined}>{dimmed ? "—" : v}</div>
          </div>
        ))}
      </div>
      <div style={{ fontSize: 11, color: "var(--text2)", display: "grid", gap: 4 }}>
        <div>• <b>Histórico:</b> {r.simArea} na área · {r.ignMod ? "modalidade ignorada" : `${r.simMod} na modalidade`} · {r.simObj} objeto similar · lead {r.med != null ? r.med + " d.u." : "—"}</div>
        <div>• <b>Ano atual:</b> {r.totSD} SD(s) no histórico · {r.tocadosAno ?? r.receb2026 ?? 0} processo(s) tocado(s){r.mediaTocadosPool > 0 && <span style={{ color: "var(--text3)" }}> (média do pool ~{Math.round(r.mediaTocadosPool)})</span>} · {r.conclAno ?? r.concl2026 ?? 0} concluído(s) · <b style={{ color: r.throughput >= 70 ? "#1e8449" : r.throughput >= 40 ? "#f39c12" : "#c0392b" }}>throughput {r.throughput || 0}%</b></div>
        <div>• <b>Carga {pool.toUpperCase()} atual:</b> {r.carga} distribuída(s){r.tempAtual > 0 && <> · <b style={{ color: "#d35400" }}>{r.tempAtual} temp.</b></>} · {r.cargaPond ?? r.carga} ponderado{pool !== "cpl" && <> · {r.sdAtual ?? 0} SD(s) aberto(s)</>} · {r.crit} críticos</div>
        <div style={{ background: "var(--card2)", borderRadius: 6, padding: "6px 10px", marginTop: 2, borderLeft: `3px solid ${poolColor}` }}>
          <div style={{ fontSize: 10, color: "var(--text3)", textTransform: "uppercase", letterSpacing: ".05em", fontWeight: 700, marginBottom: 3 }}>Distribuição ponderada (Comprador 1.0 · Avaliador 0.5)</div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", fontSize: 11 }}>
            <span><b style={{ color: poolColor }}>{r.cargaArea}</b> em "{areaLabel}"</span>
            <span>·</span>
            <span><b style={{ color: r.cargaAnaliseSD > 3 ? "#c0392b" : poolColor }}>{r.cargaAnaliseSD}</b> em análise pré-compra</span>
            <span>·</span>
            <span><b style={{ color: poolColor }}>{r.cargaTotalSub}</b> total {pool.toUpperCase()}</span>
            {r.cargaTotalSubPond != null && r.cargaTotalSubPond !== r.cargaTotalSub && (
              <>
                <span>·</span>
                <span><b style={{ color: poolColor }}>{r.cargaTotalSubPond}</b> ponderado</span>
              </>
            )}
          </div>
        </div>
        {motivos.length > 0 && (
          <div style={{ fontSize: 10, color: "var(--text3)", fontStyle: "italic", marginTop: 2 }}>
            ↳ {motivos.slice(0, 3).join(" · ")}
          </div>
        )}
      </div>
    </div>
  );
};
export { XLSX_LIB, Papa, Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard };
