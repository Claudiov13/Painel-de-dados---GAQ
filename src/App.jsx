import { DirHomeScreen } from './screens/DirHomeScreen.jsx';
import { DirDiagScreen } from './screens/DirDiagScreen.jsx';
import { DirTimelineScreen } from './screens/DirTimelineScreen.jsx';
import { UploadScreen } from './screens/UploadScreen.jsx';
import { OverviewScreen } from './screens/OverviewScreen.jsx';
import { QualidadeScreen } from './screens/QualidadeScreen.jsx';
import { ExecutivoScreen } from './screens/ExecutivoScreen.jsx';
import { CronogramaScreen } from './screens/CronogramaScreen.jsx';
import { FracionamentoScreen } from './screens/FracionamentoScreen.jsx';
import { OperacionalScreen } from './screens/OperacionalScreen.jsx';
import { ProcessosScreen } from './screens/ProcessosScreen.jsx';
import { NovaScreen } from './screens/NovaScreen.jsx';
import { BalScreen } from './screens/BalScreen.jsx';
import { ProjetosScreen } from './screens/ProjetosScreen.jsx';
import { AlertasScreen } from './screens/AlertasScreen.jsx';
import { AlertaRCScreen } from './screens/AlertaRCScreen.jsx';
import { GeradorSenhaScreen } from './screens/GeradorSenhaScreen.jsx';
import { RelCustomScreen } from './screens/RelCustomScreen.jsx';
import { AlertaSDScreen } from './screens/AlertaSDScreen.jsx';
import { MapeamentoScreen } from './screens/MapeamentoScreen.jsx';
import { GestaoScreen } from './screens/GestaoScreen.jsx';
import { MetroScreen } from './screens/MetroScreen.jsx';
import { PantanalScreen } from './screens/PantanalScreen.jsx';
import { PlanejamentoScreen } from './screens/PlanejamentoScreen.jsx';
import { ModalOverlay } from './ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from './AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from './components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from './domain/service-desk.js';
function App() {
  const [rawRows, setRawRows] = useState([]);
  const [base, setBase] = useState([]);
  const [serviceDeskData, setServiceDeskData] = useState({});
  const [meta, setMeta] = useState(null);
  const [aba, setAba] = useState("upload");
  const [loading, setLoading] = useState(false);
  const [erro, setErro] = useState("");
  const [senhaInput, setSenhaInput] = useState("");
  const [senhaOk, setSenhaOk] = useState(false);
  const [senhaErro, setSenhaErro] = useState(false);
  const [dark, setDark] = useState(false);
  const [presMode, setPresMode] = useState(false);
  const [slideshow, setSlideshow] = useState(false);
  const [slideshowTick, setSlideshowTick] = useState(60);
  const [vencidosOpen, setVencidosOpen] = useState(false);
  const [futurosTodosOpen, setFuturosTodosOpen] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(() => !window.matchMedia('(max-width: 900px)').matches);
  const [accessOpen, setAccessOpen] = useState(false);
  const accessSession = useRef(null);
  useEffect(() => {
    const query = window.matchMedia('(max-width: 900px)');
    const resize = () => setSidebarOpen(!query.matches);
    query.addEventListener('change', resize);
    return () => query.removeEventListener('change', resize);
  }, []);
  useEffect(() => { if (window.matchMedia('(max-width: 900px)').matches) setSidebarOpen(false); }, [aba]);
  useEffect(() => {
    if (!sidebarOpen || !window.matchMedia('(max-width: 900px)').matches) return;
    const close = event => { if (event.key === 'Escape') { setSidebarOpen(false); document.querySelector('[aria-controls="gaq-navigation"]')?.focus(); } };
    document.addEventListener('keydown', close);
    return () => document.removeEventListener('keydown', close);
  }, [sidebarOpen]);
  const slideshowRef = useRef(null);
  const slideshowTickRef = useRef(null);
  const SLIDESHOW_ABAS = ["overview", "executivo", "operacional", "dirTimeline", "cronograma", "alertaSD", "alertaRC", "mapeamento", "planejamento"];
  const SLIDESHOW_LABELS = { overview: "Visão Geral", executivo: "Executivo", operacional: "Operacional", dirTimeline: "Timeline", cronograma: "Cronograma", alertaRC: "Alerta RC", alertaSD: "Alerta Pré-compra", mapeamento: "Mapeamento", planejamento: "Planejamento" };
  const [spotlight, setSpotlight] = useState(false);
  const spotlightRef = useRef(null);

  // ── Perfis de acesso — definidos em auth.js como var PERFIS_ACESSO (global) ──

  // ── Estado de login (persiste na sessão do browser via sessionStorage) ────
  const [loginOk, setLoginOk] = useState(() => {
    try { return !!sessionStorage.getItem("painel_login_ok"); } catch { return false; }
  });
  const [loginUser, setLoginUser] = useState(() => {
    try { const u = sessionStorage.getItem("painel_login_user"); return (u && u !== "") ? u : null; } catch { return null; }
  });
  const [loginMaster, setLoginMaster] = useState(() => {
    try { const m = sessionStorage.getItem("painel_login_master"); return m ? JSON.parse(m) : null; } catch { return null; }
  });
  const [loginArea, setLoginArea] = useState(() => {
    try { const a = sessionStorage.getItem("painel_login_area"); return (a && a !== "") ? a : null; } catch { return null; }
  });
  // loginArea pode vir como "GIN,ALMOXARIFADO" (perfil de área multi-área) — versão pra exibir na UI
  const loginAreaDisp = loginArea ? loginArea.split(",").join(" + ") : null;
  const [loginInput, setLoginInput] = useState("");
  const [loginErroMsg, setLoginErroMsg] = useState(false);
  const [loginTab, setLoginTab] = useState("gaq");
  const [areaPassInput, setAreaPassInput] = useState("");
  const [areaPassErr, setAreaPassErr] = useState(false);
  // ── Estado do Diretor ──
  const [loginDiretor, setLoginDiretor] = useState(() => {
    try { const d = sessionStorage.getItem("painel_login_diretor"); return d ? JSON.parse(d) : null; } catch { return null; }
  });
  const [dirPassInput, setDirPassInput] = useState("");
  const [dirPassErr, setDirPassErr] = useState(false);
  const [dirAreaSel, setDirAreaSel] = useState(null); // área selecionada na home do diretor
  const [dirExpandNCL, setDirExpandNCL] = useState(false);
  const [dirExpandNCLScont, setDirExpandNCLScont] = useState(false);
  const [dirShowSDList, setDirShowSDList] = useState(false);
  const [dirTLDrill, setDirTLDrill] = useState(null); // { title, procs } — drill-down da timeline
  const isDiretor = !!loginDiretor;
  const isAdminMaster = !!loginMaster;

  async function handleLogin() {
    const inputHash = await gaqHashSenha(loginInput);
    if (typeof ADMIN_MASTER_PERFIS !== "undefined" && Object.prototype.hasOwnProperty.call(ADMIN_MASTER_PERFIS, inputHash)) {
      const master = ADMIN_MASTER_PERFIS[inputHash];
      try {
        sessionStorage.setItem("painel_login_ok", "1");
        sessionStorage.setItem("painel_login_master", JSON.stringify(master));
        sessionStorage.setItem("painel_login_user", "");
        sessionStorage.setItem("painel_login_area", "");
        sessionStorage.removeItem("painel_login_diretor");
      } catch {}
      setLoginMaster(master);
      setLoginUser(null);
      setLoginArea(null);
      setLoginDiretor(null);
      setLoginOk(true);
      setLoginErroMsg(false);
      setLoginInput("");
      if (base.length > 0) setAba("qualidade");
    } else if (Object.prototype.hasOwnProperty.call(PERFIS_ACESSO, inputHash)) {
      const user = PERFIS_ACESSO[inputHash];
      try {
        sessionStorage.setItem("painel_login_ok", "1");
        sessionStorage.setItem("painel_login_user", user || "");
        sessionStorage.removeItem("painel_login_master");
        sessionStorage.removeItem("painel_login_diretor");
      } catch {}
      setLoginMaster(null);
      setLoginUser(user);
      setLoginDiretor(null);
      setLoginOk(true);
      setLoginErroMsg(false);
      setLoginInput("");
      if (base.length > 0) setAba("overview");
    } else {
      setLoginErroMsg(true);
      setLoginInput("");
    }
  }

  async function handleAreaLogin() {
    const inputHash = await gaqHashSenha(areaPassInput);
    if (typeof AREA_PERFIS !== "undefined" && Object.prototype.hasOwnProperty.call(AREA_PERFIS, inputHash)) {
      const area = AREA_PERFIS[inputHash];
      try {
        sessionStorage.setItem("painel_login_ok", "1");
        sessionStorage.setItem("painel_login_area", area || "");
        sessionStorage.setItem("painel_login_user", "");
        sessionStorage.removeItem("painel_login_master");
        sessionStorage.removeItem("painel_login_diretor");
      } catch {}
      setLoginArea(area);
      setLoginMaster(null);
      setLoginUser(null);
      setLoginDiretor(null);
      setLoginOk(true);
      setAreaPassErr(false);
      setAreaPassInput("");
      if (base.length > 0) setAba("overview");
    } else {
      setAreaPassErr(true);
      setAreaPassInput("");
    }
  }

  async function handleDiretorLogin() {
    const inputHash = await gaqHashSenha(dirPassInput);
    if (typeof DIRETOR_PERFIS !== "undefined" && Object.prototype.hasOwnProperty.call(DIRETOR_PERFIS, inputHash)) {
      const cfg = DIRETOR_PERFIS[inputHash];
      try {
        sessionStorage.setItem("painel_login_ok", "1");
        sessionStorage.setItem("painel_login_diretor", JSON.stringify(cfg));
        sessionStorage.setItem("painel_login_user", "");
        sessionStorage.setItem("painel_login_area", "");
        sessionStorage.removeItem("painel_login_master");
        // Persistir hash no localStorage para auto-login futuro (só nesta máquina/navegador)
        localStorage.setItem("painel_diretor_remember", inputHash);
      } catch {}
      setLoginDiretor(cfg);
      setLoginMaster(null);
      setLoginUser(null);
      setLoginArea(null);
      setLoginOk(true);
      setDirPassErr(false);
      setDirPassInput("");
      if (base.length > 0) setAba("overview");
    } else {
      setDirPassErr(true);
      setDirPassInput("");
    }
  }

  function handleLogout() {
    if (loginOk) window.PainelAccess.record(loginDiretor?.nome || loginMaster?.nome || loginUser || loginArea || 'Administrador', loginDiretor ? 'Diretoria' : loginArea ? 'Área' : loginUser ? 'Comprador' : 'Administrador', 'Saída do painel');
    if (!window.confirm("Deseja realmente sair do sistema?")) return;
    try { sessionStorage.removeItem("painel_login_ok"); sessionStorage.removeItem("painel_login_user"); sessionStorage.removeItem("painel_login_area"); sessionStorage.removeItem("painel_login_master"); sessionStorage.removeItem("painel_login_diretor"); localStorage.removeItem("painel_diretor_remember"); } catch {}
    setLoginOk(false);
    setLoginMaster(null);
    setLoginUser(null);
    setLoginArea(null);
    setLoginDiretor(null);
    setDirAreaSel(null);
    setLoginInput("");
    setLoginErroMsg(false);
    setAreaPassInput("");
    setAreaPassErr(false);
    setDirPassInput("");
    setDirPassErr(false);
  }

  // ── Sync dados.js/tags.js (sem quebrar file://) ─────────────────────────

const [autoSync, setAutoSync] = useState(false);
const [syncInfo, setSyncInfo] = useState({ state: "idle", msg: "", at: null });
const [tagVersion, setTagVersion] = useState(0);
const refreshTags = useCallback(() => setTagVersion(v => v + 1), []);
const syncBusy = useRef(false);
const syncApplied = useRef(false);

const syncNow = useCallback(async (force = false) => {
  if (syncBusy.current) return;
  syncBusy.current = true;
  try {
    setSyncInfo({ state: "loading", msg: "Verificando as quatro bases…", at: new Date() });
    const snapshot = await window.PainelRepository.read({ force });
    if (snapshot.changed || !syncApplied.current) {
      const parsed = parseBase(snapshot.dados);
      window.PainelRepository.commit(snapshot);
      TagsManager.loadFromBase(parsed);
      setRawRows(snapshot.dados);
      setBase(parsed);
      setServiceDeskData(snapshot.servicedesk);
      refreshTags();
      setMeta(window.PainelRepository.metadata(snapshot, parsed.length));
      setSenhaOk(true);
      syncApplied.current = true;
    } else {
      setMeta(previous => previous ? { ...previous, verificadoEm: new Date().toISOString() } : previous);
    }
    setSyncInfo({ state: "ok", msg: snapshot.changed ? "Bases carregadas e validadas" : "Nenhuma alteração na publicação", at: new Date() });
  } catch (e) {
    setSyncInfo({ state: "error", msg: (e && e.message) ? e.message : "Falha ao atualizar", at: new Date() });
  } finally { syncBusy.current = false; }
}, [refreshTags]);

useEffect(() => {
  // expõe para debug/uso externo
  window.__PAINEL_SYNC__ = syncNow;
  return () => { try { delete window.__PAINEL_SYNC__; } catch {} };
}, [syncNow]);

useEffect(() => {
  if (!autoSync) return;
  const check = () => { if (!document.hidden) syncNow(); };
  const id = setInterval(check, 60000);
  document.addEventListener('visibilitychange', check);
  return () => { clearInterval(id); document.removeEventListener('visibilitychange', check); };
}, [autoSync, syncNow]);


  const [tagModal, setTagModal] = useState(null);

  // FIX: Nova demanda — estados separados para NCL e CPL
  const [dem, setDem] = useState({ area: "", modalidade: "", objeto: "", ignorarMod: false });
  const [recsNCL, setRecsNCL] = useState([]);
  const [recsCPL, setRecsCPL] = useState([]);
  const [tempDemandas, setTempDemandas] = useState([]);
  // Distribuição rápida temporária: só a pessoa é obrigatória, o resto é opcional
  const [quickTemp, setQuickTemp] = useState({ responsavel: "", pool: "", area: "", modalidade: "", objeto: "", quantidade: 1 });
  // Redistribuição de carteira de quem está de férias/indisponível: origem|pool -> { receiver, qty }
  const [redistForm, setRedistForm] = useState({});
  // Compradores/Avaliadores marcados como indisponíveis (férias, atestado etc.)
  const [comprIgnorados, setComprIgnorados] = useState(() => new Set());
  // Mostrar ranking completo (todos os candidatos) ou só top 3
  const [verRankingCompleto, setVerRankingCompleto] = useState(false);

  const [fAnoSD, setFAnoSD] = useState(""); const [fAnoRC, setFAnoRC] = useState("");
  const [fMod, setFMod] = useState(""); const [fArea, setFArea] = useState("");
  const [fSt, setFSt] = useState(""); const [fStDet, setFStDet] = useState(""); const [fComp, setFComp] = useState(""); const [fAval, setFAval] = useState(""); const [fAnalista, setFAnalista] = useState(""); const [fSubarea, setFSubarea] = useState("");
  const [fTag, setFTag] = useState("");
  const [fWatch, setFWatch] = useState(false);
  const [gestaoUnlocked, setGestaoUnlocked] = useState(false);
  const [gestaoPassInput, setGestaoPassInput] = useState("");
  const [gestaoPassErr, setGestaoPassErr] = useState(false);
  // ── Estado do Gerador de Senha (Admin Master) ──
  const [genTipoPerfil, setGenTipoPerfil] = useState("perfil");
  const [genSenha, setGenSenha] = useState("");
  const [genNome, setGenNome] = useState("");
  const [genArea, setGenArea] = useState("");
  const [genAreasDestaque, setGenAreasDestaque] = useState("");
  const [genIsAdmin, setGenIsAdmin] = useState(false);
  const [genResultado, setGenResultado] = useState(null); // { linha, varName, tipo }
  const [genCopiado, setGenCopiado] = useState(false);
  // ── Relatório Customizado (aba admin) ──
  const [rcuAreas, setRcuAreas] = useState([]);           // áreas selecionadas (vazio = todas)
  const [rcuDias, setRcuDias] = useState("");             // nº de d.u. a contar da RC ("" = sem filtro)
  const [rcuDiasModo, setRcuDiasModo] = useState("ate");  // "ate" = RC nos últimos N d.u. | "acima" = RC há mais de N d.u.
  const [rcuStatus, setRcuStatus] = useState(["andamento"]); // andamento|concluidos|cancelfrac|suspensos
  const [rcuInds, setRcuInds] = useState(["total","andamento","concluidos","slaok","criticos","lead","aging","parados"]);
  const [rcuExcl, setRcuExcl] = useState([]);             // processos excluídos manualmente do relatório (ids)
  const [metroComp, setMetroComp] = useState("");
  const [metroBusca, setMetroBusca] = useState("");
  const [metroAgrupado, setMetroAgrupado] = useState(false);
  const [pgMetro, setPgMetro] = useState(1);
  const [pgMetroSz, setPgMetroSz] = useState(50);
  const [pgPantanal, setPgPantanal] = useState(1);
  const [pgPantanalSz, setPgPantanalSz] = useState(50);
  const [pantanalSel, setPantanalSel] = useState(null);
  const [showEmailDiario, setShowEmailDiario] = useState(false);
  const [showGaqSextaEditor, setShowGaqSextaEditor] = useState(false);
  const [pantanalBusca, setPantanalBusca] = useState("");
  const [topMenuOpen, setTopMenuOpen] = useState(false);
  const [busca, setBusca] = useState("");

  const [selComp, setSelComp] = useState(null);
  const [selProc, setSelProc] = useState(null);
  const [projetoTag, setProjetoTag] = useState("");
  const [drillDown, setDrillDown] = useState(null);
  const [pgDrillDown, setPgDrillDown] = useState(1);
  const [pgDrillDownSz, setPgDrillDownSz] = useState(PAGE_SIZES[0] || 20);
  const [overviewInfoCard, setOverviewInfoCard] = useState("");
  const [ocultarProblematicosOverview, setOcultarProblematicosOverview] = useState(false);
  const [showFunilInfo, setShowFunilInfo] = useState(false);
  const [showFilters, setShowFilters] = useState(false);
  const [watchVersion, setWatchVersion] = useState(0);
  const [watchChanges, setWatchChanges] = useState([]);
  const [baseDiffFeed, setBaseDiffFeed] = useState([]);
  const [baseDiffInfo, setBaseDiffInfo] = useState({ hasPrevious: false, previousAt: null, currentAt: null });
  const [showWatchPanel, setShowWatchPanel] = useState(false);
  const [notifs, setNotifs] = useState([]);
  const [showNotifPanel, setShowNotifPanel] = useState(false);
  const [showHS, setShowHS] = useState(true);
  const [showStag, setShowStag] = useState(true);
  const [showSLA, setShowSLA] = useState(true);
  // ── Cronograma de entregas ─────────────────────────────────────────
  const [calMes, setCalMes] = useState(() => { const d = new Date(); return { y: d.getFullYear(), m: d.getMonth() }; });
  const [calDiaSel, setCalDiaSel] = useState(null);
  // ── Planejamento: mês expandido no cronograma anual de recorrência ────────
  const [recMesAberto, setRecMesAberto] = useState(null);
  const [pgHS, setPgHS] = useState(1); const [pgHSSz, setPgHSSz] = useState(20);
  const [pgStag, setPgStag] = useState(1); const [pgStagSz, setPgStagSz] = useState(20);
  const toggleWatch = useCallback((proc) => {
    WatchlistManager.toggle(proc.ProcessKey, proc);
    setWatchVersion(v => v + 1);
  }, []);

  const [pgAlert, setPgAlert] = useState(1); const [pgAlertSz, setPgAlertSz] = useState(20);
  const [pgCpl, setPgCpl] = useState(1); const [pgCplSz, setPgCplSz] = useState(20);
  const [pgAlertRC, setPgAlertRC] = useState(1); const [pgAlertRCSz, setPgAlertRCSz] = useState(20);
  const [pgAlertSD, setPgAlertSD] = useState(1); const [pgAlertSDSz, setPgAlertSDSz] = useState(20);
  const [pgAlertSDEnc, setPgAlertSDEnc] = useState(1); const [pgAlertSDEncSz, setPgAlertSDEncSz] = useState(20);
  const [sdCopyStatus, setSdCopyStatus] = useState("");
  const [distAnosFilter, setDistAnosFilter] = useState(null); // null = auto (ano mais recente)
  const [pgAll, setPgAll] = useState(1); const [pgAllSz, setPgAllSz] = useState(50);
  const [sortAll, setSortAll] = useState({ col: "diasTotais", dir: "desc" });
  const [hideSDConcluido, setHideSDConcluido] = useState(false);


  useEffect(() => { document.documentElement.setAttribute("data-theme", dark ? "dark" : "light"); }, [dark]);

  // ── Watchlist: detectar mudanças ao carregar base ────────────────
  const phaseIntervals = useMemo(() => base.length > 0 ? calcPhaseIntervals(base) : { byMod: {}, global: {} }, [base]);
  useEffect(() => {
    if (base.length > 0 && WatchlistManager.count() > 0) {
      const ch = WatchlistManager.detectChanges(base);
      setWatchChanges(ch);
      if (ch.length > 0) setShowWatchPanel(true);
    }
  }, [base]);
  useEffect(() => { setPgDrillDown(1); }, [drillDown]);

  useEffect(() => {
    if (!base.length) {
      setBaseDiffFeed([]);
      setBaseDiffInfo({ hasPrevious: false, previousAt: null, currentAt: null });
      return;
    }

    const SNAP_KEY = "painel_gaq_base_snapshot_v3";
    const nowLabel = (meta && meta.atualizadoEm) ? meta.atualizadoEm : new Date().toLocaleString("pt-BR");
    const snapKeyFor = (r) => {
      if (r.ProcessKey) return `pk:${r.ProcessKey}`;
      const rc = String(r.NumRC || "").trim();
      const sd = String(r.TicketSD || "").trim();
      if (rc || sd) return `rcsd:${rc}|${sd}`;
      return `obj:${nrm(r.Objeto || "").slice(0, 80)}|${r.dataAbertura ? r.dataAbertura.getTime() : ""}`;
    };
    const shortLabel = (r) => r.NumRC || r.TicketSD || r.ProcessKey || "Sem identificador";
    const buildSnap = (r) => ({
      key: snapKeyFor(r),
      label: shortLabel(r),
      numRC: r.NumRC || "",
      ticketSD: r.TicketSD || "",
      status: r.status || "",
      statusDet: r.statusDet || "",
      statusLabel: r.statusDet || r.status || "Sem status",
      faseSubarea: r.faseSubarea || "",
      emA: !!r.emA,
      isEncerrado: !!r.isEncerrado,
      dataEntregaMs: r.dataEntrega ? r.dataEntrega.getTime() : 0,
      ultimaDataMs: r.ultimaData ? r.ultimaData.getTime() : 0,
      area: r["Área Requisitante"] || "",
      modalidade: r.Modalidade || "",
    });
    const stampLabel = (ms) => {
      if (!ms) return "sem data";
      const d = new Date(ms);
      const hojeRef = new Date();
      return d.toDateString() === hojeRef.toDateString()
        ? d.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })
        : d.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" });
    };

    const currentProcMap = new Map();
    const currentSnap = base.map(r => {
      const snap = buildSnap(r);
      currentProcMap.set(snap.key, r);
      return snap;
    });

    let prev = null;
    try { prev = JSON.parse(localStorage.getItem(SNAP_KEY) || "null"); } catch {}

    if (!prev || !Array.isArray(prev.rows)) {
      setBaseDiffFeed([]);
      setBaseDiffInfo({ hasPrevious: false, previousAt: null, currentAt: nowLabel });
      try { localStorage.setItem(SNAP_KEY, JSON.stringify({ at: nowLabel, rows: currentSnap })); } catch {}
      return;
    }

    const prevMap = new Map(prev.rows.map(r => [r.key, r]));
    const currentMap = new Map(currentSnap.map(r => [r.key, r]));
    const changes = [];

    currentSnap.forEach(snap => {
      const old = prevMap.get(snap.key);
      const proc = currentProcMap.get(snap.key) || null;
      if (!old) {
        changes.push({
          key: snap.key,
          icon: "+",
          color: "var(--gaq-blue)",
          title: "Entrou na base",
          detail: `${snap.label} · ${snap.modalidade || "Sem modalidade"}`,
          stamp: stampLabel(snap.ultimaDataMs),
          sortAt: snap.ultimaDataMs || Date.now(),
          weight: 90,
          proc,
        });
        return;
      }
      if (!old.numRC && snap.numRC) {
        changes.push({
          key: snap.key,
          icon: "RC",
          color: "var(--gaq-green)",
          title: "Recebeu RC",
          detail: `${snap.label} saiu da triagem pré-compra`,
          stamp: stampLabel(snap.ultimaDataMs),
          sortAt: snap.ultimaDataMs || Date.now(),
          weight: 80,
          proc,
        });
        return;
      }
      if (!old.isEncerrado && snap.isEncerrado) {
        changes.push({
          key: snap.key,
          icon: "OK",
          color: "var(--gaq-green)",
          title: "Foi encerrado",
          detail: `${snap.label} · ${snap.statusLabel}`,
          stamp: stampLabel(snap.ultimaDataMs),
          sortAt: snap.ultimaDataMs || Date.now(),
          weight: 75,
          proc,
        });
        return;
      }
      if (old.faseSubarea && snap.faseSubarea && old.faseSubarea !== snap.faseSubarea) {
        changes.push({
          key: snap.key,
          icon: "SA",
          color: "var(--gaq-cpl)",
          title: "Mudou de subárea",
          detail: `${snap.label} · ${old.faseSubarea} -> ${snap.faseSubarea}`,
          stamp: stampLabel(snap.ultimaDataMs),
          sortAt: snap.ultimaDataMs || Date.now(),
          weight: 70,
          proc,
        });
        return;
      }
      if (old.statusLabel !== snap.statusLabel) {
        changes.push({
          key: snap.key,
          icon: "ST",
          color: "var(--gaq-orange)",
          title: "Mudou de status",
          detail: `${snap.label} · ${old.statusLabel} -> ${snap.statusLabel}`,
          stamp: stampLabel(snap.ultimaDataMs),
          sortAt: snap.ultimaDataMs || Date.now(),
          weight: 65,
          proc,
        });
        return;
      }
      if (old.dataEntregaMs !== snap.dataEntregaMs && snap.dataEntregaMs) {
        changes.push({
          key: snap.key,
          icon: "DT",
          color: "var(--gaq-red)",
          title: "Data de entrega atualizada",
          detail: `${snap.label} · ${new Date(snap.dataEntregaMs).toLocaleDateString("pt-BR")}`,
          stamp: stampLabel(snap.ultimaDataMs),
          sortAt: snap.ultimaDataMs || Date.now(),
          weight: 55,
          proc,
        });
      }
    });

    prev.rows
      .filter(r => !currentMap.has(r.key))
      .slice(0, 3)
      .forEach(r => {
        changes.push({
          key: `out:${r.key}`,
          icon: "-",
          color: "var(--gaq-text-3)",
          title: "Saiu da base atual",
          detail: `${r.label} · ${r.statusLabel || "Sem status"}`,
          stamp: prev.at || "base anterior",
          sortAt: r.ultimaDataMs || 0,
          weight: 40,
          proc: null,
        });
      });

    changes.sort((a, b) => b.sortAt - a.sortAt || b.weight - a.weight);
    setBaseDiffFeed(changes.slice(0, 8));
    setBaseDiffInfo({ hasPrevious: true, previousAt: prev.at || null, currentAt: nowLabel });
    try { localStorage.setItem(SNAP_KEY, JSON.stringify({ at: nowLabel, rows: currentSnap })); } catch {}
  }, [base, meta && meta.atualizadoEm]);

  // ── Ctrl+K Spotlight Search ──────────────────────────────────────
  useEffect(() => {
    const handler = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "k") { e.preventDefault(); setSpotlight(s => !s); }
      if (e.key === "Escape" && spotlight) setSpotlight(false);
    };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, [spotlight]);
  useEffect(() => { if (spotlight && spotlightRef.current) spotlightRef.current.focus(); }, [spotlight]);

  // ── Presentation mode body class ────────────────────────────────
  useEffect(() => {
    if (presMode) document.body.classList.add("pres-mode");
    else { document.body.classList.remove("pres-mode"); setSlideshow(false); }
  }, [presMode]);

  // ── Slideshow automático (modo apresentação) ─────────────────────
  useEffect(() => {
    if (slideshow && presMode) {
      setSlideshowTick(60);
      slideshowRef.current = setInterval(() => {
        setAba(prev => {
          const idx = SLIDESHOW_ABAS.indexOf(prev);
          return SLIDESHOW_ABAS[(idx + 1) % SLIDESHOW_ABAS.length];
        });
        setSlideshowTick(60);
      }, 60000);
      slideshowTickRef.current = setInterval(() => {
        setSlideshowTick(t => t > 0 ? t - 1 : 60);
      }, 1000);
    } else {
      clearInterval(slideshowRef.current);
      clearInterval(slideshowTickRef.current);
    }
    return () => { clearInterval(slideshowRef.current); clearInterval(slideshowTickRef.current); };
  }, [slideshow, presMode]);

  // ── Persistent filters (localStorage) ───────────────────────────
  useEffect(() => {
    try {
      const saved = JSON.parse(localStorage.getItem("painel_gaq_filters") || "{}");
      if (saved.fAnoSD) setFAnoSD(saved.fAnoSD);
      if (saved.fAnoRC) setFAnoRC(saved.fAnoRC);
      if (saved.fMod) setFMod(saved.fMod);
      if (saved.fArea) setFArea(saved.fArea);
      if (saved.fSt) setFSt(saved.fSt);
      if (saved.fStDet) setFStDet(saved.fStDet);
      // Filtros de comprador/avaliador só são restaurados para o perfil admin
      if (!loginUser) {
        if (saved.fComp) setFComp(saved.fComp);
        if (saved.fAval) setFAval(saved.fAval);
        // Legacy: migrate old fResp to fComp
        if (saved.fResp && !saved.fComp) setFComp(saved.fResp);
      }
      if (saved.fAnalista) setFAnalista(saved.fAnalista);
      if (saved.fSubarea) setFSubarea(saved.fSubarea);
      if (saved.fTag) setFTag(saved.fTag);
    } catch {}
  }, []);
  useEffect(() => {
    try { localStorage.setItem("painel_gaq_filters", JSON.stringify({ fAnoSD, fAnoRC, fMod, fArea, fSt, fStDet, fComp, fAval, fAnalista, fSubarea, fTag })); } catch {}
  }, [fAnoSD, fAnoRC, fMod, fArea, fSt, fStDet, fComp, fAval, fAnalista, fSubarea, fTag]);

  // ── Notificações: detectar mudanças desde o último acesso ──
  useEffect(() => {
    if (!loginUser || !base.length) return;
    const key = "painel_lastAccess_" + loginUser;
    const lastTs = parseInt(localStorage.getItem(key) || "0", 10);
    try { localStorage.setItem(key, Date.now().toString()); } catch {}
    if (!lastTs) return;
    const lastDate = new Date(lastTs);
    const changed = base.filter(r => r.emA && r.ultimaData && r.ultimaData > lastDate && (
      r.respNCL === loginUser || r.Comprador === loginUser || r.cplResp === loginUser || r.Pregoeiro === loginUser || r.AnalistaContrato === loginUser
    ));
    setNotifs(changed);
  }, [loginUser, base]);

  // ── Auto-login diretor via localStorage (lembra a senha entre sessões) ──
  useEffect(() => {
    if (loginOk) return; // já logado
    try {
      const saved = localStorage.getItem("painel_diretor_remember");
      if (saved && typeof DIRETOR_PERFIS !== "undefined" && Object.prototype.hasOwnProperty.call(DIRETOR_PERFIS, saved)) {
        const cfg = DIRETOR_PERFIS[saved];
        sessionStorage.setItem("painel_login_ok", "1");
        sessionStorage.setItem("painel_login_diretor", JSON.stringify(cfg));
        sessionStorage.setItem("painel_login_user", "");
        sessionStorage.setItem("painel_login_area", "");
        setLoginDiretor(cfg);
        setLoginUser(null);
        setLoginArea(null);
        setLoginOk(true);
      }
    } catch {}
  }, []);

  // Render the login immediately, then read a complete validated publication.
  useEffect(() => {
    syncNow();
    setAutoSync(true);
    try { if (sessionStorage.getItem('painel_login_ok')) setAba('overview'); } catch {}
  }, [syncNow]);

  useEffect(() => {
    if (!loginOk) { accessSession.current = null; return; }
    const profile = loginDiretor?.nome || loginMaster?.nome || loginUser || loginArea || 'Administrador';
    const type = loginDiretor ? 'Diretoria' : loginArea ? 'Área' : loginUser ? 'Comprador' : 'Administrador';
    const id = type + ':' + profile;
    if (accessSession.current !== id) {
      window.PainelAccess.record(profile, type, 'Entrada no painel');
      accessSession.current = id;
    }
  }, [loginOk, loginUser, loginArea, loginMaster, loginDiretor]);

  function handleUpload(e) {
    const f = e.target.files[0]; if (!f) return;
    setLoading(true); setErro("");
    const rdr = new FileReader();
    rdr.onload = ev => {
      try {
        let rows = [];
        if (f.name.toLowerCase().endsWith(".csv")) {
          const res = Papa.parse(ev.target.result, { header: true, skipEmptyLines: true, dynamicTyping: false, delimitersToGuess: [",", ";", "\t"] });
          rows = res.data;
        } else {
          const wb = XLSX_LIB.read(ev.target.result, { type: "binary" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          rows = XLSX_LIB.utils.sheet_to_json(ws, { defval: "" });
        }
        const parsed = parseBase(rows);
        setRawRows(rows); setBase(parsed);
        TagsManager.loadFromBase(parsed); refreshTags();
        setAutoSync(false);
        setSyncInfo({ state: 'warn', msg: 'Arquivo importado neste navegador. Sincronização automática pausada.', at: new Date() });
        setMeta({ nomeArq: f.name, atualizadoEm: 'Geração não informada (importação manual)', verificadoEm: new Date().toISOString(), total: parsed.length });
        setAba("overview");
      } catch (err) { setErro("Erro ao processar: " + err.message); }
      setLoading(false);
    };
    rdr.onerror = () => { setErro("Falha ao ler o arquivo."); setLoading(false); };
    f.name.toLowerCase().endsWith(".csv") ? rdr.readAsText(f, "UTF-8") : rdr.readAsBinaryString(f);
  }

  // ── Vista filtrada pelo perfil do usuário logado ──────────────────────────
  const baseVis = useMemo(() => {
    let arr = base;
    if (loginUser) {
      const u = nrm(loginUser);
      arr = arr.filter(r =>
        nrm(r.respNCL  || "").includes(u) ||
        nrm(r.Pregoeiro || "").includes(u) ||
        nrm(r.cplResp  || "").includes(u)
      );
    }
    if (loginArea) {
      arr = arr.filter(r => areaMatch(r["Área Requisitante"] || r.Area || r["ÁREA"] || "", loginArea));
    }
    if (isDiretor && dirAreaSel) {
      const a = nrm(dirAreaSel);
      arr = arr.filter(r => nrm(r["Área Requisitante"] || r.Area || r["ÁREA"] || "").includes(a));
    }
    return arr;
  }, [base, loginUser, loginArea, isDiretor, dirAreaSel]);

  // Base enriquecida com o log do SD + status efetivo (statusEfetivoSD). É a
  // fonte canônica do detalhe do processo — definida cedo para que a busca e
  // todos os painéis enxerguem exatamente os mesmos dados (datas e status).
  const baseEnriched = useMemo(() => {
    const sd = (serviceDeskData && Object.keys(serviceDeskData).length) ? serviceDeskData : null;
    return baseVis.map(proc => enrichSD(proc, sd));
  }, [baseVis, serviceDeskData]);

  // ── Busca (debounced) ──────────────────────────────────────────────
  const [buscaDefer, setBuscaDefer] = useState("");
  const buscaTimerRef = useRef(null);
  const handleBuscaChange = useCallback((val) => {
    setBusca(val);
    if (buscaTimerRef.current) clearTimeout(buscaTimerRef.current);
    buscaTimerRef.current = setTimeout(() => setBuscaDefer(val), 200);
  }, []);
  const buscaIndex = useMemo(() => baseEnriched.map(row => ({ row, fields: [row.NumRC, row.TicketSD, row.ProcessKey, row.Objeto, row.NumPedidoSuite, row.NumProcesso].map(nrm) })), [baseEnriched]);
  useEffect(() => () => { clearTimeout(buscaTimerRef.current); }, []);
  const buscaResults = useMemo(() => {
    if (!buscaDefer.trim() || buscaDefer.length < 3) return [];
    const q = nrm(buscaDefer.trim());
    const results = [];
    for (const entry of buscaIndex) { if (entry.fields.some(field => field.includes(q))) results.push(entry.row); if (results.length === 30) break; }
    return results;
  }, [buscaDefer, buscaIndex]);

  const anosSD = useMemo(() => [...new Set(baseVis.map(r => r.anoSD).filter(Boolean))].sort((a,b) => b-a).map(String), [baseVis]);
  const anosRC = useMemo(() => [...new Set(baseVis.map(r => r.anoRC).filter(Boolean))].sort((a,b) => b-a).map(String), [baseVis]);
  const mods = useMemo(() => {
    const m = new Map();
    baseVis.forEach(r => { const k = nrm(r.Modalidade); if (k && !m.has(k)) m.set(k, r.Modalidade); });
    return [...m.values()].sort((a,b) => nrm(a).localeCompare(nrm(b)));
  }, [baseVis]);
  const areas = useMemo(() => [...new Set(baseVis.map(r => r["Área Requisitante"]).filter(Boolean))].sort(), [baseVis]);
  const sts = useMemo(() => [...new Set(baseVis.map(r => r.status).filter(Boolean))].sort(), [baseVis]);
  const stsDet = useMemo(() => [...new Set(baseVis.map(r => r.statusDet).filter(Boolean))].sort(), [baseVis]);
  const todosComp = useMemo(() => {
    const seen = new Map();
    baseVis.forEach(r => {
      const c = (r.Comprador || "").trim();
      if (!c) return;
      const key = c.toLowerCase();
      if (!seen.has(key)) seen.set(key, c.charAt(0).toUpperCase() + c.slice(1));
    });
    return [...seen.values()].sort();
  }, [baseVis]);
  const todosAval = useMemo(() => [...new Set(baseVis.map(r => r.Avaliador).filter(Boolean))].sort(), [baseVis]);
  const todosAnalistas = useMemo(() => [...new Set(baseVis.map(r => r.AnalistaContrato).filter(Boolean))].sort(), [baseVis]);
  const allTagsList = useMemo(() => { tagVersion; return TagsManager.getTagList(); }, [tagVersion]);

  // Pools separados NCL vs CPL
  // NCL: somente compradores com processo em andamento HOJE na fase NCL
  //      (mesmo critério do "Ranking de compradores" do painel executivo).
  //      Descarta históricos inativos (ex.: Marcus, Ana Carolina) e placeholders ("N/I", "N/A").
  const NOME_INVALIDO = ["", "n/i", "n/a", "nao informado", "nao informada", "sem comprador", "sem responsavel"];
  const compsNCLPool = useMemo(() => {
    const ativos = new Set();
    base.filter(r => r.emA && r.faseSubarea === "NCL" && (r.Comprador || r.Avaliador))
      .forEach(r => ativos.add(r.Comprador || r.Avaliador));
    return [...ativos]
      .filter(c => c && !IGNORAR.includes(nrm(c)) && !NOME_INVALIDO.includes(nrm(c)))
      .sort();
  }, [base]);
  const compsCPLPool = useMemo(() => {
    const s = new Set();
    base.filter(r => r.ehCPL && r.emA).forEach(r => { if (r.Pregoeiro) s.add(r.Pregoeiro); if (r.cplResp) s.add(r.cplResp); });
    return [...s].filter(c => c && !IGNORAR.includes(nrm(c)) && !NOME_INVALIDO.includes(nrm(c))).sort();
  }, [base]);

  const fBase = useMemo(() => baseEnriched.filter(r => {
    if (fAnoSD && String(r.anoSD) !== fAnoSD) return false;
    if (fAnoRC && String(r.anoRC) !== fAnoRC) return false;
    if (fMod && nrm(r.Modalidade) !== nrm(fMod)) return false;
    if (fArea && nrm(r["Área Requisitante"]) !== nrm(fArea)) return false;
    if (fSt && nrm(r.status) !== nrm(fSt)) return false;
    if (fStDet && nrm(r.statusDet) !== nrm(fStDet)) return false;
    if (fComp && nrm(r.Comprador) !== nrm(fComp) && nrm(r.respNCL) !== nrm(fComp)) return false;
    if (fAval && nrm(r.Avaliador) !== nrm(fAval)) return false;
    if (fAnalista && nrm(r.AnalistaContrato) !== nrm(fAnalista)) return false;
    if (fSubarea && r.faseSubarea !== fSubarea) return false;
    if (fTag && !TagsManager.getForProcess(r.ProcessKey, r.TicketSD).includes(fTag)) return false;
    if (fWatch && !WatchlistManager.isWatched(r.ProcessKey)) return false;
    return true;
  }), [baseEnriched, fAnoSD, fAnoRC, fMod, fArea, fSt, fStDet, fComp, fAval, fAnalista, fSubarea, fTag, tagVersion, fWatch, watchVersion]);

  const fBack = useMemo(() => fBase.filter(r => r.emA), [fBase]);

  const counts = useMemo(() => ({
    emA: fBase.filter(r => r.emA).length,
    conc: fBase.filter(r => r.isConcluded).length,
    canc: fBase.filter(r => r.isCanceled).length,
    frac: fBase.filter(r => r.isFailed).length,
  }), [fBase]);

  const nclBack = useMemo(() => fBack.filter(r => !r.ehCPL), [fBack]);
  // alertNCL = críticos que ainda estão em fase NCL (baseado em aging de gestão: a partir da RC)
  const alertNCL = useMemo(() => fBack.filter(r => r.criticoNclGestao && r.faseSubarea === "NCL").sort((a,b) => b.diasTotaisGestao - a.diasTotaisGestao), [fBack]);
  // NCL: usa respNCL (Comprador || Avaliador) para todos os KPIs
  const compsNCL = useMemo(() => [...new Set(fBack.map(r => r.respNCL).filter(c => c && !IGNORAR.includes(nrm(c))))].sort(), [fBack]);
  // mediaNCL: média de aging gestão (desde RC) para processos ainda em fase NCL
  const mediaNCL = useMemo(() => { const p = fBack.filter(r => r.faseSubarea === "NCL" && r.diasTotaisGestao > 0); return p.length ? Math.round(p.reduce((s,r) => s + r.diasTotaisGestao, 0) / p.length) : 0; }, [fBack]);

  // rankNCL: críticos por comprador — aging gestão (desde RC, >50 d.u.), fase NCL
  const rankNCL = useMemo(() => [...new Set(fBack.map(r => r.respNCL).filter(Boolean))].map(c => ({
    name: c,
    criticos: fBack.filter(r => r.respNCL === c && r.criticoNclGestao && r.faseSubarea === "NCL").length,
    total:    fBack.filter(r => r.respNCL === c && r.faseSubarea === "NCL").length,
  })).sort((a,b) => b.criticos - a.criticos).filter(x => x.criticos > 0), [fBack]);

  // rankGeralNCL: média d.u. gestão (desde RC) — processos em fase NCL em andamento
  const rankGeralNCL = useMemo(() => [...new Set(fBack.map(r => r.respNCL).filter(Boolean))].map(c => {
    const p = fBack.filter(r => r.respNCL === c && r.faseSubarea === "NCL" && r.diasTotaisGestao > 0);
    return { name: c, media: p.length ? Math.round(p.reduce((s,r) => s + r.diasTotaisGestao, 0) / p.length) : 0, total: p.length };
  }).filter(x => x.total > 0).sort((a,b) => b.media - a.media), [fBack]);

  // SCONT: processos em fase Scont
  const scontBack = useMemo(() => fBack.filter(r => r.faseSubarea === "Scont"), [fBack]);
  const alertScont = useMemo(() => scontBack.filter(r => r.diasTotais > 50).sort((a,b) => b.diasTotais - a.diasTotais), [scontBack]);
  const rankScont = useMemo(() => [...new Set(scontBack.map(r => r.AnalistaContrato).filter(Boolean))].map(c => ({
    name: c,
    criticos: scontBack.filter(r => r.AnalistaContrato === c && r.diasTotais > 50).length,
    total:    scontBack.filter(r => r.AnalistaContrato === c).length,
  })).sort((a,b) => b.criticos - a.criticos).filter(x => x.criticos > 0), [scontBack]);
  const rankGeralScont = useMemo(() => [...new Set(scontBack.map(r => r.AnalistaContrato).filter(Boolean))].map(c => {
    const p = scontBack.filter(r => r.AnalistaContrato === c && r.diasTotais > 0);
    return { name: c, media: p.length ? Math.round(p.reduce((s,r) => s + r.diasTotais, 0) / p.length) : 0, total: p.length };
  }).filter(x => x.total > 0).sort((a,b) => b.media - a.media), [scontBack]);

  // CPL performance: TODOS os processos emA com CPL_ENCONTRADO=VERDADEIRO, agrupados por CPL_RESPONSAVEL_FINAL
  // Média usa diasCplTotal = du(CPL_DATA_RECEBIMENTO_FINAL → CPL_DATA_HOMOLOGACAO_FINAL ou hoje)
  const cplPerfBack = useMemo(() => fBack.filter(r => r.ehCPL), [fBack]);
  const rankCPLPerf = useMemo(() => [...new Set(cplPerfBack.map(r => r.cplResp).filter(Boolean))].map(resp => {
    const ativos = cplPerfBack.filter(r => r.cplResp === resp);
    const comDias = ativos.filter(r => r.diasCplTotal > 0);
    return { name: resp, total: ativos.length, criticos: ativos.filter(r => r.criticoCpl).length, media: comDias.length ? Math.round(comDias.reduce((a,b) => a + b.diasCplTotal, 0) / comDias.length) : 0 };
  }).filter(x => x.total > 0).sort((a,b) => b.total - a.total), [cplPerfBack]);

  // Scont performance: TODOS os processos emA com AnalistaContrato preenchido
  // Média usa diasScontFase = du(Envio Pedido/Suite → Recebimento DJ ou hoje)
  const scontPerfBack = useMemo(() => fBack.filter(r => r.AnalistaContrato), [fBack]);
  const rankScontPerf = useMemo(() => [...new Set(scontPerfBack.map(r => r.AnalistaContrato).filter(Boolean))].map(c => {
    const ativos = scontPerfBack.filter(r => r.AnalistaContrato === c);
    const comDias = ativos.filter(r => r.diasScontFase > 0);
    return { name: c, total: ativos.length, criticos: ativos.filter(r => r.diasTotais > 50).length, media: comDias.length ? Math.round(comDias.reduce((a,b) => a + b.diasScontFase, 0) / comDias.length) : 0 };
  }).filter(x => x.total > 0).sort((a,b) => b.total - a.total), [scontPerfBack]);

  // Concluídos — ranking histórico de performance por pessoa e modalidade
  const baseConc = useMemo(() => baseVis.filter(r => r.isConcluded), [baseVis]);
  // rankNCLConc: média d.u. gestão (desde RC) — concluídos
  const rankNCLConc = useMemo(() => [...new Set(baseConc.map(r => r.respNCL).filter(Boolean))].map(c => {
    const p = baseConc.filter(r => r.respNCL === c && r.diasTotaisGestao > 0);
    return { name: c, total: p.length, media: p.length ? Math.round(p.reduce((a,b) => a + b.diasTotaisGestao, 0) / p.length) : 0 };
  }).filter(x => x.total > 0).sort((a,b) => b.media - a.media), [baseConc]);
  const rankCPLConc = useMemo(() => [...new Set(baseConc.filter(r => r.ehCPL).map(r => r.cplResp).filter(Boolean))].map(c => {
    const p = baseConc.filter(r => r.ehCPL && r.cplResp === c && r.diasCplTotal > 0);
    return { name: c, total: p.length, media: p.length ? Math.round(p.reduce((a,b) => a + b.diasCplTotal, 0) / p.length) : 0 };
  }).filter(x => x.total > 0).sort((a,b) => b.media - a.media), [baseConc]);
  const rankScontConc = useMemo(() => [...new Set(baseConc.filter(r => r.AnalistaContrato).map(r => r.AnalistaContrato).filter(Boolean))].map(c => {
    const p = baseConc.filter(r => r.AnalistaContrato === c && r.diasScontFase > 0);
    return { name: c, total: p.length, media: p.length ? Math.round(p.reduce((a,b) => a + b.diasScontFase, 0) / p.length) : 0 };
  }).filter(x => x.total > 0).sort((a,b) => b.media - a.media), [baseConc]);
  // nclByMod: média d.u. gestão (desde RC) por modalidade — concluídos
  const nclByMod = useMemo(() => [...new Set(baseConc.map(r => r.Modalidade).filter(Boolean))].map(m => {
    const p = baseConc.filter(r => r.Modalidade === m && r.diasTotaisGestao > 0);
    return { name: m, total: p.length, media: p.length ? Math.round(p.reduce((a,b) => a + b.diasTotaisGestao, 0) / p.length) : 0 };
  }).filter(x => x.total >= 2).sort((a,b) => b.media - a.media), [baseConc]);

  // Aging SD por comprador NCL: média de (Abertura SD → Encerramento SD) — concluídos/encerrados SD
  const rankNCLAgingSD = useMemo(() => [...new Set(baseConc.filter(r => r.respNCL && r.diasAgingSD > 0).map(r => r.respNCL))].map(c => {
    const p = baseConc.filter(r => r.respNCL === c && r.diasAgingSD > 0);
    return { name: c, total: p.length, media: p.length ? Math.round(p.reduce((a,b) => a + b.diasAgingSD, 0) / p.length) : 0 };
  }).filter(x => x.total > 0).sort((a,b) => b.media - a.media), [baseConc]);

  // Aging SD em andamento — processos SD ainda abertos (sem encSD)
  const rankNCLAgingSDAtivo = useMemo(() => [...new Set(fBack.filter(r => r.respNCL && r.aberturaSD && r.faseSubarea === "NCL").map(r => r.respNCL))].map(c => {
    const p = fBack.filter(r => r.respNCL === c && r.aberturaSD && r.faseSubarea === "NCL" && r.diasSD > 0);
    return { name: c, total: p.length, media: p.length ? Math.round(p.reduce((a,b) => a + b.diasSD, 0) / p.length) : 0 };
  }).filter(x => x.total > 0).sort((a,b) => b.media - a.media), [fBack]);
  const cplByMod = useMemo(() => [...new Set(baseConc.filter(r => r.ehCPL).map(r => r.Modalidade).filter(Boolean))].map(m => {
    const p = baseConc.filter(r => r.ehCPL && r.Modalidade === m && r.diasCplTotal > 0);
    return { name: m, total: p.length, media: p.length ? Math.round(p.reduce((a,b) => a + b.diasCplTotal, 0) / p.length) : 0 };
  }).filter(x => x.total >= 2).sort((a,b) => b.media - a.media), [baseConc]);
  const scontByMod = useMemo(() => [...new Set(baseConc.filter(r => r.AnalistaContrato).map(r => r.Modalidade).filter(Boolean))].map(m => {
    const p = baseConc.filter(r => r.AnalistaContrato && r.Modalidade === m && r.diasScontFase > 0);
    return { name: m, total: p.length, media: p.length ? Math.round(p.reduce((a,b) => a + b.diasScontFase, 0) / p.length) : 0 };
  }).filter(x => x.total >= 2).sort((a,b) => b.media - a.media), [baseConc]);

  // Per-person × modality breakdown with auto-insights
  const nclModPerf = useMemo(() => {
    const comps = [...new Set(baseConc.filter(r => r.respNCL && r.Modalidade).map(r => r.respNCL))];
    return comps.map(comp => {
      const concProcs = baseConc.filter(r => r.respNCL === comp);
      const andProcs = fBack.filter(r => r.respNCL === comp && r.faseSubarea === "NCL");
      const mods = [...new Set([...concProcs, ...andProcs].map(r => r.Modalidade).filter(Boolean))];
      const modStats = mods.map(m => {
        const c = concProcs.filter(r => r.Modalidade === m && r.diasTotaisGestao > 0);
        const a = andProcs.filter(r => r.Modalidade === m);
        const avg = c.length ? Math.round(c.reduce((s,r) => s + r.diasTotaisGestao,0) / c.length) : null;
        return { mod: m, nConc: c.length, avg, nAnd: a.length };
      }).filter(x => x.nConc > 0 || x.nAnd > 0).sort((a,b) => (a.avg||999) - (b.avg||999));
      const withAvg = modStats.filter(x => x.avg !== null && x.nConc >= 3);
      let insight = null;
      if (withAvg.length >= 2) {
        const best = withAvg[0], worst = withAvg[withAvg.length-1];
        if (best.avg < worst.avg * 0.7) insight = `Melhor em ${best.mod} (${best.avg} d.u.) · Mais lento em ${worst.mod} (${worst.avg} d.u.)`;
      }
      return { name: comp, mods: modStats, insight, totalConc: concProcs.length };
    }).filter(x => x.mods.length > 0 && x.totalConc >= 3).sort((a,b) => b.totalConc - a.totalConc);
  }, [baseConc, fBack]);

  const cplModPerf = useMemo(() => {
    const comps = [...new Set(baseConc.filter(r => r.ehCPL && r.cplResp && r.Modalidade).map(r => r.cplResp))];
    return comps.map(comp => {
      const concProcs = baseConc.filter(r => r.ehCPL && r.cplResp === comp);
      const andProcs = cplPerfBack.filter(r => r.cplResp === comp);
      const mods = [...new Set([...concProcs, ...andProcs].map(r => r.Modalidade).filter(Boolean))];
      const modStats = mods.map(m => {
        const c = concProcs.filter(r => r.Modalidade === m && r.diasCplTotal > 0);
        const a = andProcs.filter(r => r.Modalidade === m);
        const avg = c.length ? Math.round(c.reduce((s,r) => s + r.diasCplTotal,0) / c.length) : null;
        return { mod: m, nConc: c.length, avg, nAnd: a.length };
      }).filter(x => x.nConc > 0 || x.nAnd > 0).sort((a,b) => (a.avg||999) - (b.avg||999));
      const withAvg = modStats.filter(x => x.avg !== null && x.nConc >= 2);
      let insight = null;
      if (withAvg.length >= 2) {
        const best = withAvg[0], worst = withAvg[withAvg.length-1];
        if (best.avg < worst.avg * 0.7) insight = `Melhor em ${best.mod} (${best.avg} d.u.) · Mais lento em ${worst.mod} (${worst.avg} d.u.)`;
      }
      return { name: comp, mods: modStats, insight, totalConc: concProcs.length };
    }).filter(x => x.mods.length > 0 && x.totalConc >= 2).sort((a,b) => b.totalConc - a.totalConc);
  }, [baseConc, cplPerfBack]);

  const scontModPerf = useMemo(() => {
    const comps = [...new Set(baseConc.filter(r => r.AnalistaContrato && r.Modalidade).map(r => r.AnalistaContrato))];
    return comps.map(comp => {
      const concProcs = baseConc.filter(r => r.AnalistaContrato === comp);
      const andProcs = scontPerfBack.filter(r => r.AnalistaContrato === comp);
      const mods = [...new Set([...concProcs, ...andProcs].map(r => r.Modalidade).filter(Boolean))];
      const modStats = mods.map(m => {
        const c = concProcs.filter(r => r.Modalidade === m && r.diasScontFase > 0);
        const a = andProcs.filter(r => r.Modalidade === m);
        const avg = c.length ? Math.round(c.reduce((s,r) => s + r.diasScontFase,0) / c.length) : null;
        return { mod: m, nConc: c.length, avg, nAnd: a.length };
      }).filter(x => x.nConc > 0 || x.nAnd > 0).sort((a,b) => (a.avg||999) - (b.avg||999));
      const withAvg = modStats.filter(x => x.avg !== null && x.nConc >= 2);
      let insight = null;
      if (withAvg.length >= 2) {
        const best = withAvg[0], worst = withAvg[withAvg.length-1];
        if (best.avg < worst.avg * 0.7) insight = `Melhor em ${best.mod} (${best.avg} d.u.) · Mais lento em ${worst.mod} (${worst.avg} d.u.)`;
      }
      return { name: comp, mods: modStats, insight, totalConc: concProcs.length };
    }).filter(x => x.mods.length > 0 && x.totalConc >= 2).sort((a,b) => b.totalConc - a.totalConc);
  }, [baseConc, scontPerfBack]);

  // ── Dados históricos por ano para o perfil de Área ──────────────────────
  const areaAnoData = useMemo(() => {
    if (!loginArea) return [];
    const procs = base.filter(r => areaMatch(r["Área Requisitante"] || r.Area || r["ÁREA"] || "", loginArea));
    const byAno = {};
    procs.forEach(r => {
      const dt = r.aberturaSD || r.aberturaRC;
      if (!dt) return;
      let ano; try { ano = new Date(dt).getFullYear(); } catch { return; }
      if (isNaN(ano) || ano < 2019 || ano > 2030) return;
      if (!byAno[ano]) byAno[ano] = { total: 0, conc: 0, dias: [], mods: {} };
      byAno[ano].total++;
      if (!r.emA) {
        byAno[ano].conc++;
        if (r.diasTotais > 0) byAno[ano].dias.push(r.diasTotais);
      }
      if (r.Modalidade) byAno[ano].mods[r.Modalidade] = (byAno[ano].mods[r.Modalidade] || 0) + 1;
    });
    return Object.keys(byAno).sort().map(ano => {
      const d = byAno[ano];
      const media = d.dias.length ? Math.round(d.dias.reduce((s,v) => s+v,0) / d.dias.length) : null;
      const topMod = Object.entries(d.mods).sort((a,b) => b[1]-a[1])[0];
      return { ano, total: d.total, conc: d.conc, andamento: d.total - d.conc, media, topMod: topMod ? topMod[0] : "—" };
    });
  }, [base, loginArea]);

  const cargaNCL = useMemo(() => [...new Set(fBack.map(r => r.respNCL).filter(Boolean))].map(c => ({
    name: (c.split(" ")[0] || c), full: c, total: fBack.filter(r => r.respNCL === c).length, criticos: fBack.filter(r => r.respNCL === c && r.criticoNclGestao).length,
  })).filter(d => d.total > 0).sort((a,b) => b.total - a.total), [fBack]);

  const cplBack = useMemo(() => {
    const rows = fBase.filter(r => r.emA && r.ehCPL);
    const seen = new Set();
    return rows.filter(r => { const k = r.ProcessKey; if (!k) return true; if (seen.has(k)) return false; seen.add(k); return true; });
  }, [fBase]);

  // alertCPL = críticos que ainda estão em fase CPL (última data preenchida é CPL)
  const cmpText = (a, b) => String(a || "").localeCompare(String(b || ""), "pt-BR", { numeric: true, sensitivity: "base" });
  const sortAlertRCDesc = (a, b) =>
    (b.diasTotaisGestao || 0) - (a.diasTotaisGestao || 0) ||
    (b.diasRC || 0) - (a.diasRC || 0) ||
    cmpText(a.NumRC || a.ProcessKey, b.NumRC || b.ProcessKey);
  const sortAlertSDDesc = (a, b) =>
    (b.diasTotais || 0) - (a.diasTotais || 0) ||
    (b.diasSD || 0) - (a.diasSD || 0) ||
    cmpText(a.TicketSD || a.ProcessKey, b.TicketSD || b.ProcessKey);

  const alertCPL = useMemo(() => cplBack.filter(r => r.criticoCpl && r.faseSubarea === "CPL").sort((a,b) => b.diasTotais - a.diasTotais), [cplBack]);

  // ── Alerta RC: em andamento COM RC preenchida, ordenado por aging da RC ──
  const alertRC = useMemo(() => fBack.filter(r => r.NumRC && r.NumRC.trim() !== "").sort(sortAlertRCDesc), [fBack]);
  // ── Alerta SD: em andamento COM SD mas SEM RC, ordenado por aging do SD ──
  const alertSD = useMemo(() => fBack.filter(r => r.TicketSD && r.TicketSD.trim() !== "" && (!r.NumRC || r.NumRC.trim() === "")).sort(sortAlertSDDesc), [fBack]);
  // ── SD dividido: em andamento (sem encerramento) vs encerrado aguardando RC pela área ──
  const alertSDAberto = useMemo(() => alertSD.filter(r => !r.encSD).sort((a, b) => (b.diasSDAberto || 0) - (a.diasSDAberto || 0)), [alertSD]);
  const alertSDEncerrado = useMemo(() => alertSD.filter(r => !!r.encSD).sort((a, b) => (b.diasDesdeEncSD || 0) - (a.diasDesdeEncSD || 0)), [alertSD]);
  const dirScopedBase = useMemo(() => {
    if (!isDiretor) return [];
    if (!dirAreaSel) return base;
    const area = nrm(dirAreaSel);
    return base.filter(r => nrm(r["Área Requisitante"] || "").includes(area));
  }, [isDiretor, base, dirAreaSel]);
  const diretorDiag = useMemo(() => {
    if (!isDiretor || !dirScopedBase.length) return null;

    const emA = dirScopedBase.filter(r => r.emA);
    const ativosComRC = emA.filter(r => r.NumRC && r.NumRC.trim() !== "");
    const sdTriagem = emA.filter(r => r.TicketSD && r.TicketSD.trim() !== "" && (!r.NumRC || r.NumRC.trim() === ""));
    const classificados = ativosComRC
      .map(r => {
        const sla = calcSLAClassification(r, phaseIntervals);
        const diag = calcDirectorPriority(r);
        return sla ? { ...r, _sla: sla, _diag: diag } : null;
      })
      .filter(Boolean)
      .sort((a, b) => (b._sla.pctSLA || 0) - (a._sla.pctSLA || 0) || (b.diasTotais || 0) - (a.diasTotais || 0));

    const criticos = classificados.filter(r => r._sla.bucket === "critico");
    const atencao = classificados.filter(r => r._sla.bucket === "atencao");
    const slaVencidos = classificados.filter(r => r._sla.bucket === "sla_vencido");
    const controlados = classificados.filter(r => r._sla.bucket === "no_prazo");
    const atrasoPrazo = slaVencidos;
    const entregaVencida = classificados.filter(r => r._sla.entregaVencida);
    const janelaComprometida = classificados.filter(r => r._sla.janelaComprometida);
    const entregaCurta = classificados.filter(r => r._sla.entregaProxima);
    const semDataEntrega = classificados.filter(r => !r._sla.temEntrega);
    const consumoElevado = classificados.filter(r => (r._sla.pctSLA || 0) >= 0.7);
    const atencaoForte = classificados.filter(r => r._sla.atencaoForte);
    const atencaoCombinada = classificados.filter(r => r._sla.atencaoCombinada);
    const areaScoreDir = calcAreaSlaScore(classificados, phaseIntervals);

    const summarizeByMod = (rows) => Object.entries(rows.reduce((acc, r) => {
      const mod = r.Modalidade || "N/I";
      acc[mod] = (acc[mod] || 0) + 1;
      return acc;
    }, {})).sort((a, b) => b[1] - a[1]).slice(0, 3).map(([name, count]) => ({ name, count }));

    const flowYears = new Set();
    dirScopedBase.forEach(r => {
      if (r.dataAbertura) flowYears.add(r.dataAbertura.getFullYear());
      if (r.isEncerrado && r.ultimaData) flowYears.add(r.ultimaData.getFullYear());
    });

    const flowByYear = [...flowYears].sort((a, b) => b - a).map(ano => {
      const entradasList = dirScopedBase.filter(r => r.dataAbertura && r.dataAbertura.getFullYear() === ano);
      const concluidosList = dirScopedBase.filter(r => r.isConcluded && r.ultimaData && r.ultimaData.getFullYear() === ano);
      const encerradosList = dirScopedBase.filter(r => r.isEncerrado && r.ultimaData && r.ultimaData.getFullYear() === ano);
      const entradas = entradasList.length;
      const concluidos = concluidosList.length;
      const encerrados = encerradosList.length;
      return {
        ano,
        entradas,
        concluidos,
        encerrados,
        saldo: entradas - encerrados,
        taxaConclusao: entradas ? Math.round((concluidos / entradas) * 100) : 0,
        taxaEncerramento: entradas ? Math.round((encerrados / entradas) * 100) : 0,
        entradasList,
        concluidosList,
        encerradosList,
      };
    });

    const anoAtual = new Date().getFullYear();
    const flowAtual = flowByYear.find(r => r.ano === anoAtual) || {
      ano: anoAtual,
      entradas: 0,
      concluidos: 0,
      encerrados: 0,
      saldo: 0,
      taxaConclusao: 0,
      taxaEncerramento: 0,
      entradasList: [],
      concluidosList: [],
      encerradosList: [],
    };

    return {
      scoped: dirScopedBase,
      emA,
      ativosComRC,
      sdTriagem,
      classificados,
      criticos,
      atencao,
      controlados,
      atrasoPrazo,
      entregaVencida,
      janelaComprometida,
      entregaCurta,
      semDataEntrega,
      consumoElevado,
      atencaoForte,
      atencaoCombinada,
      slaVencidos,
      mediaScore: areaScoreDir.score,
      mediaConsumoPrazo: classificados.length ? Math.round(classificados.reduce((s, r) => s + (r._sla.pctSLALabel || 0), 0) / classificados.length) : 0,
      criticosTopMods: summarizeByMod(criticos),
      atencaoTopMods: summarizeByMod(atencao),
      controladosTopMods: summarizeByMod(controlados),
      topCriticos: criticos.slice(0, 8),
      flowByYear,
      flowAtual,
    };
  }, [isDiretor, dirScopedBase, phaseIntervals]);
  // Reset paginação quando filtros mudam (evita página fora do range)
  useEffect(() => { setPgAlertRC(1); }, [alertRC]);
  useEffect(() => { setPgAlertSD(1); }, [alertSDAberto]);
  useEffect(() => { setPgAlertSDEnc(1); }, [alertSDEncerrado]);
  const pregoeirosCPL = useMemo(() => [...new Set(cplBack.map(r => r.Pregoeiro).filter(Boolean))], [cplBack]);
  const mediaCPL = useMemo(() => { const p = cplBack.filter(r => r.diasCpl > 0); return p.length ? Math.round(p.reduce((s,r) => s + r.diasCpl, 0) / p.length) : 0; }, [cplBack]);

  const rankCPLResp = useMemo(() => {
    // FIX: usar APENAS processos em andamento (cplBack) para todos os KPIs CPL
    const emACPL = cplBack.filter(r => r.cplResp);
    return [...new Set(emACPL.map(r => r.cplResp).filter(Boolean))].map(resp => {
      const ativos = emACPL.filter(r => r.cplResp === resp);
      const comDias = ativos.filter(r => r.diasCplTotal > 0);
      return { name: resp, total: ativos.length, media: comDias.length ? Math.round(comDias.reduce((a,b) => a + b.diasCplTotal, 0) / comDias.length) : 0, totalHist: ativos.length };
    }).filter(x => x.total > 0).sort((a,b) => b.total - a.total);
  }, [cplBack]);

  const cargaCPLResp = useMemo(() => rankCPLResp.map(r => ({ name: (r.name.split(" ")[0] || r.name), full: r.name, total: r.total || 0 })).filter(d => d.total > 0), [rankCPLResp]);
  const cargaCPLPreg = useMemo(() => [...new Set(cplBack.map(r => r.Pregoeiro).filter(Boolean))].map(p => ({
    name: (p.split(" ")[0] || p), full: p, total: cplBack.filter(r => r.Pregoeiro === p).length,
  })).filter(d => d.total > 0).sort((a,b) => b.total - a.total), [cplBack]);

  const topStatus = useMemo(() => {
    const cnt = {};
    fBack.forEach(r => { const s = (r.statusDet || r.status || "Sem status").toString().trim(); if (s) cnt[s] = (cnt[s] || 0) + 1; });
    return Object.entries(cnt).sort((a,b) => b[1] - a[1]).slice(0, 10);
  }, [fBack]);

  const mapInsights = useMemo(() => calcMapInsights(baseVis), [baseVis]);
  const alertMapDem = useMemo(() => {
    if (!dem.objeto || !mapInsights.padroes.length) return [];
    const words = new Set(kw(dem.objeto, 4));
    return mapInsights.padroes.filter(p => words.has(p.word));
  }, [dem.objeto, mapInsights]);

  // ── Executive dashboard data ──────────────────────────────────────
  const execData = useMemo(() => {
    if (!baseVis.length) return null;
    const tot = baseVis.length;
    const emAL = baseVis.filter(r => r.emA);
    const conc = baseVis.filter(r => r.isConcluded);
    const canc = baseVis.filter(r => r.isCanceled);
    const frac = baseVis.filter(r => r.isFailed);
    const prob = [...canc, ...frac];
    const taxaProb = tot ? Math.round((prob.length / tot) * 100) : 0;
    const taxaConc = tot ? Math.round((conc.length / tot) * 100) : 0;
    const concLTs = conc.filter(r => r.LeadTime > 0).map(r => r.LeadTime);
    const mediaLTConc = concLTs.length ? Math.round(concLTs.reduce((a,b) => a+b,0)/concLTs.length) : 0;
    const aging = [
      { faixa: "0-30 d.u.", count: emAL.filter(r => r.diasTotaisGestao <= 30).length, color: "#27ae60" },
      { faixa: "31-50 d.u.", count: emAL.filter(r => r.diasTotaisGestao > 30 && r.diasTotaisGestao <= 50).length, color: "#f39c12" },
      { faixa: "51-100 d.u.", count: emAL.filter(r => r.diasTotaisGestao > 50 && r.diasTotaisGestao <= 100).length, color: "#e67e22" },
      { faixa: ">100 d.u.", count: emAL.filter(r => r.diasTotaisGestao > 100).length, color: "#c0392b" },
    ];
    const areaProb = {};
    prob.forEach(r => { const a = r["Área Requisitante"] || "N/I"; areaProb[a] = (areaProb[a]||0)+1; });
    const topAreasProb = Object.entries(areaProb).sort((a,b) => b[1]-a[1]).slice(0,5);
    // ── NOVO: Tempo médio (d.u.) por Modalidade (comparativo anual)
// Regra: início = DATA DO RECEBIMENTO DA RC (aging oficial a partir da RC).
//        fim   = CPL_DATA_HOMOLOGACAO_FINAL; senão, DATA DE ENVIO DO PEDIDO (fornecedor) OU SUITE SESC (SCONT ou CPL);
//        se fim vazio, usa hoje (para processos ainda abertos).
const durByModYear = {};
const durCountByModYear = {};
const yearsSet = new Set();
const today0 = new Date(); today0.setHours(0,0,0,0);

function getInicioFim(r){
  const ini = pd(r["DATA DO RECEBIMENTO DA RC"] || r["Data Recebimento RC"] || r["Recebimento RC"]);
  if (!ini) return { ini: null, fim: null };
  const fimRaw = pd(r["CPL_DATA_HOMOLOGACAO_FINAL"]) || pd(r["DATA DE ENVIO DO PEDIDO (fornecedor) OU SUITE SESC (SCONT ou CPL)"]);
  const fim = fimRaw || today0; // hoje já está zerado no parseBase, mas aqui garantimos zero hora
  return { ini, fim };
}

baseVis.forEach(r => {
  if (!r.Modalidade) return;
  const { ini, fim } = getInicioFim(r);
  if (!ini || !fim) return;
  const y = ini.getFullYear();
  yearsSet.add(y);
  const d = Math.max(0, du(ini, fim));
  if (!durByModYear[r.Modalidade]) durByModYear[r.Modalidade] = {};
  if (!durCountByModYear[r.Modalidade]) durCountByModYear[r.Modalidade] = {};
  durByModYear[r.Modalidade][y] = (durByModYear[r.Modalidade][y] || 0) + d;
  durCountByModYear[r.Modalidade][y] = (durCountByModYear[r.Modalidade][y] || 0) + 1;
});

const yearsAll = [...yearsSet].sort((a,b)=>a-b);
const latestYear = yearsAll.length ? yearsAll[yearsAll.length-1] : null;
const prevYear = yearsAll.length > 1 ? yearsAll[yearsAll.length-2] : null;

const tempoPorModalidadeAno = (() => {
  if (!latestYear) return { years: [], data: [] };
  const y1 = prevYear, y2 = latestYear;
  const rows = Object.keys(durByModYear).map(mod => {
    const sum2 = (durByModYear[mod][y2] || 0), cnt2 = (durCountByModYear[mod][y2] || 0);
    const sum1 = y1 ? (durByModYear[mod][y1] || 0) : 0, cnt1 = y1 ? (durCountByModYear[mod][y1] || 0) : 0;
    const avg2 = cnt2 ? Math.round(sum2/cnt2) : null;
    const avg1 = cnt1 ? Math.round(sum1/cnt1) : null;
    const totalCnt = cnt1 + cnt2;
    return { name: mod, avg1, avg2, cnt1, cnt2, totalCnt, [String(y1)]: avg1, [String(y2)]: avg2 };
  })
  // manter modalidades com volume mínimo para comparativo fazer sentido
  .filter(r => (r.cnt2 >= 3) || (r.cnt1 >= 3))
  // prioriza volume do ano mais recente
  .sort((a,b) => (b.cnt2 - a.cnt2) || (b.totalCnt - a.totalCnt))
  .slice(0, 12);

  return { years: y1 ? [y1, y2] : [y2], data: rows };
})();

// Modalidades Mais Lentas: APENAS processos concluídos
// Início: DATA DO RECEBIMENTO DA RC (aging oficial)
// Fim: CPL_DATA_HOMOLOGACAO_FINAL, senão DATA DE ENVIO DO PEDIDO
const topModLentas = (() => {
  const modMap = {};
  conc.forEach(r => {
    if (!r.Modalidade) return;
    const ini = pd(r["DATA DO RECEBIMENTO DA RC"] || r["Data Recebimento RC"] || r["Recebimento RC"]);
    const fim = pd(r["CPL_DATA_HOMOLOGACAO_FINAL"]) || pd(r["DATA DE ENVIO DO PEDIDO (fornecedor) OU SUITE SESC (SCONT ou CPL)"]);
    if (!ini || !fim) return;
    const d = Math.max(0, du(ini, fim));
    if (!modMap[r.Modalidade]) modMap[r.Modalidade] = { sum: 0, cnt: 0 };
    modMap[r.Modalidade].sum += d;
    modMap[r.Modalidade].cnt++;
  });
  return Object.entries(modMap)
    .filter(([_, v]) => v.cnt >= 3)
    .map(([name, v]) => ({ name, media: Math.round(v.sum / v.cnt), count: v.cnt }))
    .sort((a, b) => b.media - a.media)
    .slice(0, 5);
})();

    const now = new Date();
    const meses = [];
    for (let i = 11; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const label = d.toLocaleDateString("pt-BR", { month: "short", year: "2-digit" });
      const y = d.getFullYear(), mo = d.getMonth();
      const abertos = baseVis.filter(r => r._ref && r._ref.getFullYear() === y && r._ref.getMonth() === mo).length;
      const conclM = conc.filter(r => {
        const env = pd(r["DATA DE ENVIO DO PEDIDO (fornecedor) OU SUITE SESC (SCONT ou CPL)"]);
        return env && env.getFullYear() === y && env.getMonth() === mo;
      }).length;
      meses.push({ label, abertos, concluidos: conclM });
    }
    const compCarga = {};
    emAL.forEach(r => {
      const c = r.ehCPL ? (r.Pregoeiro || r.cplResp) : r.respNCL; if (!c) return;
      if (!compCarga[c]) compCarga[c] = { total: 0, criticos: 0 };
      compCarga[c].total++;
      if (r.criticoNclGestao || r.criticoCpl) compCarga[c].criticos++;
    });
    const topSobrecarregados = Object.entries(compCarga)
      .map(([name, v]) => ({ name, ...v })).sort((a,b) => b.criticos - a.criticos || b.total - a.total).slice(0,5);
    const concDentroSLA = conc.filter(r => r.LeadTime > 0 && r.LeadTime <= 50).length;
    const taxaSLA = concLTs.length ? Math.round((concDentroSLA / concLTs.length) * 100) : 0;
    const risco = emAL.filter(r => r.diasTotais > 100);

    // ── Trends: comparativo mês atual vs mês anterior ──
    const nowT = new Date();
    const curM = nowT.getMonth(), curY = nowT.getFullYear();
    const prevM = curM === 0 ? 11 : curM - 1;
    const prevYT = curM === 0 ? curY - 1 : curY;
    const curMonthProcs = baseVis.filter(r => r._ref && r._ref.getMonth() === curM && r._ref.getFullYear() === curY);
    const prevMonthProcs = baseVis.filter(r => r._ref && r._ref.getMonth() === prevM && r._ref.getFullYear() === prevYT);
    // Conclusões/problemas pelo mês em que foram encerrados (ultimaData), não pelo mês de abertura
    const curConc = baseVis.filter(r => r.isConcluded && r.ultimaData && r.ultimaData.getMonth() === curM && r.ultimaData.getFullYear() === curY).length;
    const prevConc = baseVis.filter(r => r.isConcluded && r.ultimaData && r.ultimaData.getMonth() === prevM && r.ultimaData.getFullYear() === prevYT).length;
    const curProb = baseVis.filter(r => (r.isCanceled || r.isFailed) && r.ultimaData && r.ultimaData.getMonth() === curM && r.ultimaData.getFullYear() === curY).length;
    const prevProb = baseVis.filter(r => (r.isCanceled || r.isFailed) && r.ultimaData && r.ultimaData.getMonth() === prevM && r.ultimaData.getFullYear() === prevYT).length;
    const trendConc = curConc > prevConc ? "up" : curConc < prevConc ? "down" : "flat";
    const trendProb = curProb < prevProb ? "up" : curProb > prevProb ? "down" : "flat";

    // ── Comparativo Período Mês vs Mês Anterior (KPIs múltiplos) ──
    const curMA = curMonthProcs.length, prevMA = prevMonthProcs.length;
    const curConcN = curMonthProcs.filter(r => r.isConcluded).length;
    const prevConcN = prevMonthProcs.filter(r => r.isConcluded).length;
    const curAgingArr = curMonthProcs.filter(r => r.emA && r.diasTotaisGestao > 0).map(r => r.diasTotaisGestao);
    const prevAgingArr = prevMonthProcs.filter(r => r.emA && r.diasTotaisGestao > 0).map(r => r.diasTotaisGestao);
    const curMediaAging = curAgingArr.length ? Math.round(curAgingArr.reduce((a,b)=>a+b,0)/curAgingArr.length) : 0;
    const prevMediaAging = prevAgingArr.length ? Math.round(prevAgingArr.reduce((a,b)=>a+b,0)/prevAgingArr.length) : 0;
    const curCrit = curMonthProcs.filter(r => r.emA && r.criticoNclGestao).length;
    const prevCrit = prevMonthProcs.filter(r => r.emA && r.criticoNclGestao).length;
    const pctDelta = (cur, prev) => prev === 0 ? (cur > 0 ? "+100" : "0") : ((cur - prev) / prev * 100).toFixed(0);
    const comparativo = {
      aberturas: { cur: curMA, prev: prevMA, delta: pctDelta(curMA, prevMA) },
      conclusoes: { cur: curConcN, prev: prevConcN, delta: pctDelta(curConcN, prevConcN) },
      problematicos: { cur: curProb, prev: prevProb, delta: pctDelta(curProb, prevProb) },
      agingMedio: { cur: curMediaAging, prev: prevMediaAging, delta: pctDelta(curMediaAging, prevMediaAging) },
      criticos: { cur: curCrit, prev: prevCrit, delta: pctDelta(curCrit, prevCrit) },
    };

    return {
      tot, emA: emAL.length, conc: conc.length, canc: canc.length, frac: frac.length,
      taxaProb, taxaConc, mediaLTConc, taxaSLA, aging, topAreasProb, topModLentas,
      tempoPorModalidadeAno, comparativo,
      criticalTotal: (alertNCL.length + alertCPL.length),
      meses, topSobrecarregados,
      risco: risco.length,
      alertNCLCount: alertNCL.length, alertCPLCount: alertCPL.length,
      nclBackCount: nclBack.length, cplBackCount: cplBack.length,
      trendConc, trendProb, curConc, prevConc, curProb, prevProb,
    };
  }, [baseVis, alertNCL, alertCPL, nclBack, cplBack]);

  const overviewAdmin = useMemo(() => {
    if (loginUser || loginArea || !base.length) return null;

    const rowKey = (r) => r.ProcessKey || `${r.NumRC || ""}|${r.TicketSD || ""}|${r.Objeto || ""}`;
    const uniqRows = (rows) => {
      const seen = new Set();
      return rows.filter(r => {
        const key = rowKey(r);
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    };
    // ── Classificação SLA global (calcSLAClassification) ──
    const ativosComRC = fBack.filter(r => r.NumRC && String(r.NumRC).trim() !== "");
    const classificadosOverview = fBack.map(r => {
      const sla = calcSLAClassification(r, phaseIntervals);
      if (!sla) return null;
      const diag = calcDirectorPriority(r);
      return { ...r, _sla: sla, _diag: diag, _overview: { criticoEntrega: sla.bucket === "critico", criticoReason: sla.bucket, overviewScore: sla.processScore } };
    }).filter(Boolean);
    const areaScore = calcAreaSlaScore(classificadosOverview, phaseIntervals);
    const sdTriagem = fBack.filter(r => r.TicketSD && String(r.TicketSD).trim() !== "" && (!r.NumRC || String(r.NumRC).trim() === ""));
    const sdAtrasado = sdTriagem.filter(r => r.atrasoSD);
    const criticos = classificadosOverview.filter(r => r._sla.bucket === "critico");
    const slaVencidosScore = classificadosOverview.filter(r => r._sla.bucket === "sla_vencido");
    const atencao = classificadosOverview.filter(r => r._sla.bucket === "atencao");
    const slaAtendidosScore = classificadosOverview.filter(r => r._sla.bucket === "no_prazo");
    const atrasadosRC = ativosComRC.filter(r => r.atrasoGeral);
    const atencaoKeys = new Set(atencao.map(rowKey));
    const slaAtendidosRCBruto = ativosComRC.filter(r => !r.atrasoGeral);
    const atencaoDentroSlaRC = slaAtendidosRCBruto.filter(r => atencaoKeys.has(rowKey(r)));
    const slaAtendidosRC = slaAtendidosRCBruto.filter(r => !atencaoKeys.has(rowKey(r)));
    const slaAtendidosSD = sdTriagem.filter(r => !r.atrasoSD);
    const slaVencidos = uniqRows([...atrasadosRC, ...sdAtrasado]);
    const slaAtendidos = uniqRows([...slaAtendidosRC, ...slaAtendidosSD]);
    const criticosEntregaVencida = criticos.filter(r => r._sla.entregaVencida);
    const atencaoJanela = atencao.filter(r => r._sla.janelaComprometida);
    const atencaoPrazo = atencao.filter(r => r._sla.atencaoPrazo || r._sla.projecaoEstouraSla);
    const criticosSlaEntrega = criticos.filter(r => r._sla.slaVencido && (r._sla.entregaVencida || r._sla.projecaoForaCronograma));
    const conclValidos = fBase.filter(r => {
      if (!r.isConcluded || (r.diasTotais || 0) <= 0) return false;
      const prazo = r.prazoGeral || getPrazoGeral(nrm(r.Modalidade || ""));
      return prazo > 0;
    });
    const conclNoPrazo = conclValidos.filter(r => {
      const prazo = r.prazoGeral || getPrazoGeral(nrm(r.Modalidade || ""));
      return (r.diasTotais || 0) <= prazo;
    });
    const slaCumprido = conclValidos.length ? Math.round((conclNoPrazo.length / conclValidos.length) * 100) : 0;
    const mediaScore = areaScore.score;
    const scoreColor = mediaScore >= 75 ? "var(--gaq-green)" : mediaScore >= 50 ? "var(--gaq-orange)" : "var(--gaq-red)";
    const activeAttention = criticos.length + atencao.length;
    const firstValid = (...vals) => vals.find(v => v && v !== "—" && v !== "N/A") || "Sistema";
    const processLabel = (r) => r.NumRC || r.TicketSD || r.ProcessKey || "Sem identificador";
    const stampLabel = (dt) => {
      if (!dt) return "sem data";
      const now = new Date();
      return dt.toDateString() === now.toDateString()
        ? dt.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })
        : dt.toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" });
    };
    const monthBuckets = Array.from({ length: 8 }, (_, idx) => {
      const d = new Date();
      d.setDate(1);
      d.setMonth(d.getMonth() - (7 - idx));
      return { y: d.getFullYear(), m: d.getMonth() };
    });
    const byMonth = (rows, pickDate, calcValue) => monthBuckets.map(({ y, m }) => {
      const scoped = rows.filter(r => {
        const dt = pickDate(r);
        return dt && dt.getFullYear() === y && dt.getMonth() === m;
      });
      return calcValue(scoped);
    });
    const sparkEmAndamento = byMonth(fBack, r => r.dataAbertura, rows => rows.length);
    const sparkAtrasados = byMonth(slaVencidos, r => r.dataAbertura, rows => rows.length);
    const sparkCriticos = byMonth(criticos, r => r.dataAbertura, rows => rows.length);
    const sparkSla = byMonth(slaAtendidos, r => r.dataAbertura, rows => rows.length);
    const sparkScore = byMonth(classificadosOverview, r => r.dataAbertura, rows => rows.length
      ? calcAreaSlaScore(rows, phaseIntervals).score
      : 0);

    const subareaStatusDefs = [
      { key: "critico", label: "Crítico", color: "#d35400" },
      { key: "atencao", label: "Em atenção", color: "#ff9500" },
      { key: "sla_vencido", label: "SLA vencido", color: "#ff3b30" },
      { key: "no_prazo", label: "No prazo", color: "#34c759" },
    ];
    const subareas = [
      { key: "NCL", label: "NCL · Núcleo de Compras", color: "var(--gaq-ncl)", rows: fBack.filter(r => r.faseSubarea === "NCL") },
      { key: "CPL", label: "CPL · Comissão de Licitação", color: "var(--gaq-cpl)", rows: fBack.filter(r => r.faseSubarea === "CPL") },
      { key: "Scont", label: "Scont · Contratos", color: "var(--gaq-rsp)", rows: fBack.filter(r => r.faseSubarea === "Scont") },
    ].map(item => {
      const classifiedRows = classificadosOverview.filter(r => r.faseSubarea === item.key);
      const statusSegments = subareaStatusDefs.map(def => {
        const rows = classifiedRows.filter(r => r._sla.bucket === def.key);
        return {
          ...def,
          rows,
          count: rows.length,
          pct: item.rows.length ? Math.round((rows.length / item.rows.length) * 100) : 0,
        };
      });
      return {
        ...item,
        pct: fBack.length ? Math.round((item.rows.length / fBack.length) * 100) : 0,
        statusSegments,
      };
    });

    const recentActivity = [...fBase]
      .filter(r => r.ultimaData)
      .sort((a, b) => b.ultimaData - a.ultimaData)
      .slice(0, 6)
      .map(r => {
        const who = firstValid(r.respFase, r.respAtivo, r.Comprador, r.Pregoeiro, r.AnalistaContrato, r.respNCL);
        const label = processLabel(r);
        const detail = r.statusDet || r.status || r.faseAtual || "Sem status";
        const color = r.isConcluded ? "var(--gaq-green)"
          : r.atrasoGeral ? "var(--gaq-red)"
          : r.faseSubarea === "CPL" ? "var(--gaq-cpl)"
          : r.faseSubarea === "Scont" ? "var(--gaq-rsp)"
          : "var(--gaq-blue)";
        const verb = r.isConcluded ? "encerrou"
          : r.isCanceled ? "cancelou"
          : r.isFailed ? "marcou fracasso em"
          : r.NumRC ? "movimentou"
          : "atualizou";
        return { proc: r, who, label, detail, color, verb, stamp: stampLabel(r.ultimaData) };
      });

    const rankRows = (rows, field) => Object.entries(rows.reduce((acc, r) => {
      const key = r[field] || "Não informado";
      if (!acc[key]) acc[key] = [];
      acc[key].push(r);
      return acc;
    }, {}))
      .map(([label, procRows]) => ({ label, count: procRows.length, rows: procRows }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 10);

    const modalidades = rankRows(fBack, "Modalidade");
    const areasAtivas = rankRows(fBack, "Área Requisitante");
    const statusAtivos = topStatus.map(([label, count]) => ({
      label,
      count,
      rows: fBack.filter(r => (r.statusDet || r.status || "Sem status").toString().trim() === label),
    }));
    // ── Total de processos por situação (base completa, não só "em andamento") ──
    const situacaoConcluidos = fBase.filter(r => r.isConcluded);
    const situacaoCancelados = fBase.filter(r => r.isCanceled);
    const situacaoFracassados = fBase.filter(r => r.isFailed);
    const situacaoSuspensos = fBase.filter(r => r.isSuspended);
    const situacaoOutrosRows = fBase.filter(r => !r.emA && !r.isConcluded && !r.isCanceled && !r.isFailed && !r.isSuspended);
    const totalPorSituacao = [
      { key: "emA", label: "Em andamento", count: fBack.length, color: "var(--gaq-blue)", rows: fBack, problematico: false },
      { key: "concluido", label: "Concluído", count: situacaoConcluidos.length, color: "var(--gaq-green)", rows: situacaoConcluidos, problematico: false },
      { key: "suspenso", label: "Suspenso", count: situacaoSuspensos.length, color: "var(--gaq-purple)", rows: situacaoSuspensos, problematico: false },
      { key: "cancelado", label: "Cancelado", count: situacaoCancelados.length, color: "var(--gaq-orange)", rows: situacaoCancelados, problematico: true },
      { key: "fracassado", label: "Fracassado", count: situacaoFracassados.length, color: "var(--gaq-red)", rows: situacaoFracassados, problematico: true },
      { key: "outros", label: "Outros/Sem status", count: situacaoOutrosRows.length, color: "var(--gaq-text-3)", rows: situacaoOutrosRows, problematico: false },
    ].filter(item => item.count > 0);

    // ── Total de processos por modalidade (base completa) ──
    const totalPorModalidade = Object.entries(fBase.reduce((acc, r) => {
      const key = r.Modalidade || "Não informado";
      if (!acc[key]) acc[key] = [];
      acc[key].push(r);
      return acc;
    }, {}))
      .map(([label, rows]) => ({
        label,
        count: rows.length,
        rows,
        countAtivo: rows.filter(r => !r.isCanceled && !r.isFailed).length,
        rowsAtivo: rows.filter(r => !r.isCanceled && !r.isFailed),
      }))
      .sort((a, b) => b.count - a.count);

    const rowScopeKey = (r) => r ? (r.ProcessKey || `${r.NumRC || ""}|${r.TicketSD || ""}|${r.Objeto || ""}`) : "";
    const scopeKeys = new Set(fBase.map(rowScopeKey));
    const scopedBaseDiffFeed = isDiretor
      ? baseDiffFeed.filter(change => change.proc && scopeKeys.has(rowScopeKey(change.proc)))
      : baseDiffFeed;

    const hour = new Date().getHours();
    const saudacao = hour < 12 ? "Bom dia" : hour < 19 ? "Boa tarde" : "Boa noite";

    return {
      saudacao,
      emAndamento: fBack.length,
      ativosComRC,
      sdTriagem,
      sdAtrasado,
      criticos,
      atencao,
      slaVencidos,
      slaAtendidos,
      slaAtendidosRC,
      slaAtendidosSD,
      atencaoDentroSlaRC,
      slaVencidosScore,
      slaAtendidosScore,
      activeAttention,
      atrasadosSla: slaVencidos,
      atrasadosRC,
      conclNoPrazo,
      conclValidos,
      slaCumprido,
      mediaScore,
      scoreColor,
      subareas,
      recentActivity,
      modalidades,
      areasAtivas,
      statusAtivos,
      totalGeral: fBase.length,
      totalPorSituacao,
      totalPorModalidade,
      baseDiffFeed: scopedBaseDiffFeed,
      baseDiffInfo,
      mediaNCL,
      mediaCPL,
      above60: fBack.filter(r => (r.diasTotais || 0) > 60),
      criticosEntregaVencida,
      atencaoJanela,
      atencaoPrazo,
      criticosSlaEntrega,
      updatedAt: meta && meta.atualizadoEm ? meta.atualizadoEm : null,
      sparkEmAndamento,
      sparkAtrasados,
      sparkCriticos,
      sparkSla,
      sparkScore,
    };
  }, [loginUser, loginArea, base, fBack, fBase, topStatus, baseDiffFeed, baseDiffInfo, mediaNCL, mediaCPL, meta, phaseIntervals]);

  const overviewInfoDefsLegacy = {
    em_andamento: {
      title: "Em andamento",
      color: "#2e86c1",
      bg: "#eef6ff",
      lines: [
        "Mostra todos os processos em andamento no recorte atual.",
        "Inclui processos ainda só em pré-compra e processos que já receberam RC.",
      ],
    },
    atrasados_sla: {
      title: "Atrasados por SLA",
      color: "#c0392b",
      bg: "#fdf0ee",
      lines: [
        "Considera apenas SLA não cumprido.",
        "Pré-compra sem RC: entra quando fica aberta por mais de 10 dias úteis.",
        "Processos com RC: entra quando os dias do processo passam do prazo geral da modalidade.",
        "A data de entrega não entra sozinha nesta caixa.",
      ],
    },
    criticos: {
      title: "Críticos",
      color: "#d35400",
      bg: "#fff4ea",
      lines: [
        "Mostra pressão real de entrega.",
        "Entra se a entrega já venceu.",
        "Entra se a janela estiver comprometida: o prazo restante do processo é maior do que a folga até a entrega.",
        "Entra também quando o SLA já estourou e a entrega está em até 30 dias úteis.",
      ],
    },
    sla_cumprido: {
      title: "SLA cumprido",
      color: "#1e8449",
      bg: "#eefaf2",
      lines: [
        "Percentual de processos concluídos dentro do prazo geral da modalidade.",
        "Considera apenas processos concluídos com prazo válido para comparação.",
      ],
    },
    score_medio: {
      title: "Score médio",
      color: "#8e44ad",
      bg: "#f6f0fb",
      lines: [
        "O score médio usa a régua nova da Visão Geral.",
        "Atraso só de SLA pesa menos do que criticidade ligada à entrega.",
        "Casos críticos por entrega ou janela comprometida puxam o score para cima.",
      ],
    },
  };

  const overviewInfoDefs = {
    em_andamento: {
      title: "Em andamento",
      color: "#2e86c1",
      bg: "#eef6ff",
      lines: [
        "Mostra todos os processos em andamento no recorte atual.",
        "Inclui processos ainda só em pré-compra e processos que já receberam RC.",
      ],
    },
    atrasados_sla: {
      title: "SLA Vencido",
      color: "#c0392b",
      bg: "#fdf0ee",
      lines: [
        "Considera apenas prazo de SLA não cumprido.",
        "Pré-compra sem RC: entra quando fica aberta por mais de 10 dias úteis.",
        "Processos com RC: entra quando o aging da RC passa de 30 ou 90 dias, conforme a modalidade.",
        "A data de entrega não altera esta caixa.",
      ],
    },
    criticos: {
      title: "Crítico",
      color: "#d35400",
      bg: "#fff4ea",
      lines: [
        "Crítico exige duas condições simultâneas: SLA vencido E entrega vencida (ou projeção fora do cronograma).",
        "Para RC, a entrega fica crítica quando a projeção de conclusão passa da data prevista ou a entrega já venceu.",
        "Para pré-compra sem RC, o bucket máximo possível é SLA Vencido — não atinge Crítico (sem entrega definida).",
        "Atenção (próximo de Crítico) dispara quando qualquer margem individual fica ≤ 3 d.u. OU SLA e entrega ambos ≤ 8 d.u. — não vira Crítico até o SLA realmente estourar.",
      ],
    },
    atencao: {
      title: "Atenção (regra Opção C)",
      color: "#e67e22",
      bg: "#fff7e6",
      lines: [
        "Atenção é o aviso de quase-crítico: dispara somente quando o processo está perto de virar Crítico, não como early-warning genérico.",
        "Sinal forte (qualquer um dispara sozinho): margem de SLA ≤ 3 d.u., OU margem de entrega ≤ 3 d.u., OU a projeção atual já estoura o SLA.",
        "Sinal combinado (precisa dos dois juntos): margem de SLA ≤ 8 d.u. E margem de entrega ≤ 8 d.u. — pressão dupla mesmo que cada lado isolado não seja crítico.",
        "Margem de SLA = prazo da modalidade − (consumido + projetado restante). Margem de entrega = d.u. até a data prevista − d.u. até a conclusão projetada.",
        "Um processo com 15 d.u. de folga numa frente e 30 d.u. na outra não entra em Atenção — está confortável. A régua agora corta só o realmente apertado.",
      ],
    },
    sla_cumprido: {
      title: "SLA Atendido",
      color: "#1e8449",
      bg: "#eefaf2",
      lines: [
        "Mostra os processos em andamento que ainda estão dentro do prazo simples de SLA.",
        "Para pré-compra, considera 10 dias úteis. Para RC, considera 30 ou 90 dias conforme a modalidade.",
        "Esta caixa pode incluir processos em Atenção, desde que o SLA ainda não tenha vencido.",
        `O histórico concluído dentro do prazo continua a ser acompanhado e hoje está em ${overviewAdmin ? overviewAdmin.slaCumprido : 0}%.`,
      ],
    },
    score_medio: {
      title: "Score da área",
      color: "#8e44ad",
      bg: "#f6f0fb",
      lines: [
        "Base de 1000 pontos, normalizada para 0 a 100.",
        "O score usa buckets exclusivos: Crítico, Atenção, SLA Vencido e SLA Atendido.",
        "Cada caso Crítico tira 30 pontos, Atenção tira 15 e SLA Vencido tira 10.",
        "Cada processo classificado como SLA Atendido adiciona 30 pontos.",
        "Quanto maior o score, melhor a saúde atual da carteira.",
      ],
    },
  };

  const executiveStrategic = useMemo(() => {
    if (!base.length) return null;

    const rowKey = (r) => r ? (r.ProcessKey || `${r.NumRC || ""}|${r.TicketSD || ""}|${r.Objeto || ""}`) : "";
    const modColorsExec = {
      "pregao": "#1a5276",
      "concorrencia": "#6c3483",
      "dispensa": "#e67e22",
      "inexigibilidade": "#c0392b",
      "adesao": "#16a085",
      "credenciamento": "#2e86c1",
      "dialogo": "#8e44ad",
      "n/i": "#8e8e93",
    };
    const uniqRows = (rows) => {
      const seen = new Set();
      return rows.filter(r => {
        const key = rowKey(r);
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    };

    // ── Classificação SLA global (calcSLAClassification) ──
    const classificados = fBack.map(r => {
      const sla = calcSLAClassification(r, phaseIntervals);
      if (!sla) return null;
      const diag = calcDirectorPriority(r);
      return { ...r, _sla: sla, _diag: diag };
    }).filter(Boolean);

    const criticoRows = classificados.filter(r => r._sla.bucket === "critico");
    const slaVencidoRows = classificados.filter(r => r._sla.bucket === "sla_vencido");
    const atrasoRows = slaVencidoRows;
    const atencaoRows = classificados.filter(r => r._sla.bucket === "atencao");
    const regularRows = classificados.filter(r => r._sla.bucket === "no_prazo");
    const areaScore = calcAreaSlaScore(classificados, phaseIntervals);

    const problematicosRows = uniqRows([...criticoRows, ...atencaoRows, ...slaVencidoRows]);
    const taxaProblematica = fBack.length ? Math.round((problematicosRows.length / fBack.length) * 100) : 0;
    const mediaConsumoPrazo = classificados.length ? Math.round(classificados.reduce((s, r) => s + (r._sla.pctSLALabel || 0), 0) / classificados.length) : 0;
    const scoreAreaExec = areaScore.score;
    const scoreColorExec = scoreAreaExec >= 75 ? "#27ae60" : scoreAreaExec >= 50 ? "#ff9500" : "#c0392b";
    const scoreLabelExec = scoreAreaExec >= 75 ? "saúde estável" : scoreAreaExec >= 50 ? "atenção" : "pressão alta";
    const scoreColor = mediaConsumoPrazo >= 100 ? "#c0392b" : mediaConsumoPrazo >= 70 ? "#e67e22" : "#27ae60";
    const scoreLabel = mediaConsumoPrazo >= 100 ? "pressão alta" : mediaConsumoPrazo >= 70 ? "atenção" : "estável";

    const modMap = {};
    fBack.forEach(r => {
      const mod = r.Modalidade || "N/I";
      if (!modMap[mod]) modMap[mod] = [];
      modMap[mod].push(r);
    });
    const modalidadesAtivas = Object.entries(modMap)
      .map(([label, rows]) => ({ label, count: rows.length, rows }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 6);

    const modMapAll = {};
    fBase.forEach(r => {
      const mod = r.Modalidade || "N/I";
      if (!modMapAll[mod]) modMapAll[mod] = [];
      modMapAll[mod].push(r);
    });
    const modalidades = Object.entries(modMapAll)
      .map(([label, rows]) => ({ label, count: rows.length, rows }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 6);

    const riskSegments = [
      { label: "Crítico", value: criticoRows.length, color: "#d35400", rows: criticoRows },
      { label: "SLA Vencido", value: slaVencidoRows.length, color: "#ff3b30", rows: slaVencidoRows },
      { label: "Atenção", value: atencaoRows.length, color: "#ff9500", rows: atencaoRows },
      { label: "SLA Atendido", value: regularRows.length, color: "#34c759", rows: regularRows },
    ];

    riskSegments.splice(0, riskSegments.length, riskSegments[0], riskSegments[2], riskSegments[1], riskSegments[3]);

    const semaforoLegacy = [
      { label: "CrÃ­tico", rows: criticoRows, color: "#d35400", detail: "SLA vencido e entrega ou cronograma rompidos." },
      { label: "Atenção", rows: atencaoRows, color: "#ff9500", detail: "Próximo do limite de SLA (≥80%) ou entrega em ≤15 dias" },
      { label: "SLA Vencido", rows: atrasoRows, color: "#ff3b30", detail: "Fora do SLA da modalidade, mas ainda dentro do cronograma de entrega" },
      { label: "Crítico", rows: criticoRows, color: "#d35400", detail: "Fora do SLA e fora do prazo de entrega/cronograma" },
    ];

    const semaforo = [
      { label: "Crítico", rows: criticoRows, color: "#d35400", detail: "SLA vencido E entrega vencida (ou projeção fora do cronograma)." },
      { label: "Atenção", rows: atencaoRows, color: "#ff9500", detail: "Margem ≤ 3 d.u. em SLA ou entrega, OU SLA e entrega ambos ≤ 8 d.u., OU projeção estoura SLA." },
      { label: "SLA Vencido", rows: slaVencidoRows, color: "#ff3b30", detail: "Fora do SLA da modalidade, mas ainda com cronograma de entrega preservado." },
      { label: "SLA Atendido", rows: regularRows, color: "#34c759", detail: "Dentro do SLA da modalidade e sem pressão relevante na entrega." },
    ];

    const critBase = criticoRows.length > 0 ? criticoRows : (atencaoRows.length > 0 ? atencaoRows : slaVencidoRows);
    const topCriticos = [...critBase]
      .sort((a, b) =>
        (b._sla?.pctSLA || 0) - (a._sla?.pctSLA || 0) ||
        Number(!!b._sla?.entregaVencida) - Number(!!a._sla?.entregaVencida) ||
        (b.diasTotais || 0) - (a.diasTotais || 0)
      )
      .slice(0, 5);

    const problemByMod = modalidadesAtivas
      .map(m => ({ ...m, sobPressao: problematicosRows.filter(r => nrm(r.Modalidade) === nrm(m.label)).length }))
      .sort((a, b) => b.sobPressao - a.sobPressao || b.count - a.count);

    // Ranking de compradores por volume de demandas + média d.u. NCL.
    // Exclui pré-compra com SD concluído e ainda sem RC: não há esforço do
    // comprador enquanto a RC não é aberta, então não conta na carteira dele.
    const buyerMap = {};
    fBack.filter(r => r.faseSubarea === "NCL" && (r.Comprador || r.Avaliador) && !isSDConcluidoSemRC(r)).forEach(r => {
      const nome = r.Comprador || r.Avaliador;
      if (!buyerMap[nome]) buyerMap[nome] = { nome, demandas: 0, totalDias: 0, count: 0, rows: [] };
      buyerMap[nome].demandas++;
      buyerMap[nome].rows.push(r);
      if (r.diasTotais > 0) { buyerMap[nome].totalDias += r.diasTotais; buyerMap[nome].count++; }
    });
    const buyerRanking = Object.values(buyerMap)
      .map(b => ({ ...b, mediaDU: b.count > 0 ? Math.round(b.totalDias / b.count) : 0 }))
      .sort((a, b) => b.demandas - a.demandas)
      .slice(0, 10);

    // Pré-compra (Service Desk com número de SD mas SEM RC vinculada) por Status.
    // Usa statusEfetivoSD (core.js): aproveita o log do ticket — "EM ANÁLISE
    // SERVICE DESK" com data de encerramento (encSD) vira "SERVICE DESK
    // CONCLUÍDO"; demais status do JSON são mantidos. Este gráfico MOSTRA inclusive
    // os concluídos (é o panorama da pré-compra), mas eles não entram na carga do
    // comprador no ranking acima nem no balanceamento de nova demanda.
    const preCompraRows = fBack.filter(r => r.somenteSD && (r.Comprador || r.Avaliador));
    const preCompraStatusMap = {};
    preCompraRows.forEach(r => {
      const st = statusEfetivoSD(r);
      if (!preCompraStatusMap[st]) preCompraStatusMap[st] = { status: st, count: 0, rows: [] };
      preCompraStatusMap[st].count++;
      preCompraStatusMap[st].rows.push(r);
    });
    const preCompraStatus = Object.values(preCompraStatusMap).sort((a, b) => b.count - a.count);
    const preCompraTotal = preCompraRows.length;

    // Mesma carteira (SD sem RC) quebrada por responsável NCL (respNCL = Comprador
    // || Avaliador). Por pessoa: quantos "em análise" e quantos "concluído sem RC".
    // A soma de todos continua = preCompraTotal (157).
    const preCompraByRespMap = {};
    preCompraRows.forEach(r => {
      const nome = r.respNCL || r.Comprador || r.Avaliador || "—";
      const n = nrm(statusEfetivoSD(r));
      if (!preCompraByRespMap[nome]) preCompraByRespMap[nome] = { nome, emAnalise: 0, concluido: 0, outros: 0, total: 0, emAnaliseRows: [], concluidoRows: [], outrosRows: [], rows: [] };
      const g = preCompraByRespMap[nome];
      g.total++; g.rows.push(r);
      if (n.includes("concluido")) { g.concluido++; g.concluidoRows.push(r); }
      else if (n.includes("em analise service desk")) { g.emAnalise++; g.emAnaliseRows.push(r); }
      else { g.outros++; g.outrosRows.push(r); }
    });
    const preCompraByResp = Object.values(preCompraByRespMap)
      .sort((a, b) => b.total - a.total || b.emAnalise - a.emAnalise || a.nome.localeCompare(b.nome));

    const buyerHistoryMap = {};
    const concludedModTotals = {};
    fBase.filter(r => r.isConcluded && (r.Comprador || r.Avaliador || r.Pregoeiro || r.respNCL)).forEach(r => {
      const nome = r.Comprador || r.Avaliador || r.Pregoeiro || r.respNCL;
      const mod = r.Modalidade || "N/I";
      concludedModTotals[mod] = (concludedModTotals[mod] || 0) + 1;
      if (!buyerHistoryMap[nome]) buyerHistoryMap[nome] = { nome, concluidos: 0, totalDias: 0, count: 0, rows: [], mods: {} };
      buyerHistoryMap[nome].concluidos++;
      buyerHistoryMap[nome].rows.push(r);
      buyerHistoryMap[nome].mods[mod] = (buyerHistoryMap[nome].mods[mod] || 0) + 1;
      if (r.diasTotais > 0) {
        buyerHistoryMap[nome].totalDias += r.diasTotais;
        buyerHistoryMap[nome].count++;
      }
    });
    const buyerRankingConcluded = Object.values(buyerHistoryMap)
      .map(b => {
        const modParts = Object.entries(b.mods || {})
          .map(([label, count]) => {
            const pctBuyer = b.concluidos ? Math.round((count / b.concluidos) * 100) : 0;
            const pctTotal = concludedModTotals[label] ? Math.round((count / concludedModTotals[label]) * 100) : 0;
            const color = modColorsExec[nrm(label)] || "#8e8e93";
            return { label, count, pctBuyer, pctTotal, color };
          })
          .sort((a, bPart) => bPart.count - a.count || a.label.localeCompare(bPart.label));
        return { ...b, mediaDU: b.count > 0 ? Math.round(b.totalDias / b.count) : 0, modParts };
      })
      .sort((a, b) => b.concluidos - a.concluidos || a.mediaDU - b.mediaDU || a.nome.localeCompare(b.nome))
      .slice(0, 10);

    const insights = [];
    const criticosEntregaVencida = criticoRows.filter(r => r._sla?.entregaVencida);
    const atencaoJanelaExec = atencaoRows.filter(r => r._sla?.janelaComprometida);
    if (criticosEntregaVencida.length > 0) insights.push({
      icon: "ENT",
      color: "#c0392b",
      title: "Entregas já vencidas",
      text: `${criticosEntregaVencida.length} processo(s) com SLA vencido e entrega já ultrapassada.`,
    });
    if (atencaoJanelaExec.length > 0) insights.push({
      icon: "JAN",
      color: "#ff9500",
      title: "Casos em atenÃ§Ã£o",
      text: `${atencaoJanelaExec.length} processo(s) com janela de entrega curta ou comprometida.`,
    });
    if (slaVencidoRows.length > 0) insights.push({
      icon: "SLA",
      color: "#ff3b30",
      title: "SLA vencido",
      text: `${slaVencidoRows.length} processo(s) fora do SLA da modalidade mas ainda dentro do cronograma de entrega.`,
    });
    if (problemByMod.length > 0 && problemByMod[0].sobPressao > 0) insights.push({
      icon: "MOD",
      color: "#1a5276",
      title: "Modalidade sob pressão",
      text: `${problemByMod[0].label} concentra ${problemByMod[0].sobPressao} caso(s) sob pressão no recorte actual.`,
    });
    if (insights.length === 0) insights.push({
      icon: "OK",
      color: "#27ae60",
      title: "Sem pressão relevante",
      text: "Não há sinais relevantes de ruptura entre SLA e calendário de entrega no recorte actual.",
    });

    return {
      total: fBase.length,
      ativos: fBack.length,
      abertos: fBase.length,
      concluidos: fBase.filter(r => r.isConcluded).length,
      emAndamento: fBack.length,
      cancelados: fBase.filter(r => r.isCanceled).length,
      fracassados: fBase.filter(r => r.isFailed).length,
      classificados,
      criticos: criticoRows,
      atencao: atencaoRows,
      slaVencidos: slaVencidoRows,
      slaAtendidos: regularRows,
      problematicosRows,
      taxaProblematica,
      mediaConsumoPrazo,
      scoreArea: scoreAreaExec,
      scoreColor: scoreColorExec,
      scoreLabel: scoreLabelExec,
      modalidades,
      riskSegments,
      semaforo,
      topCriticos,
      insights,
      buyerRanking,
      buyerRankingConcluded,
      preCompraStatus,
      preCompraTotal,
      preCompraByResp,
      atrasadosSla: slaVencidoRows,
    };
  }, [base, fBase, fBack, phaseIntervals]);

  const qualityData = useMemo(() => {
    if (!base.length) return null;

    // Avalia toda a base acessível (baseVis), ignorando filtros de UI — qualidade é métrica global, não de recorte
    const rows = baseVis;
    const rowKey = (r) => r ? (r.ProcessKey || `${r.NumRC || ""}|${r.TicketSD || ""}|${r.Objeto || ""}`) : "";
    const addDecor = (list, issueKey, detailFn) => list.map(r => ({
      ...r,
      _dqIssue: issueKey,
      _dqDetail: typeof detailFn === "function" ? detailFn(r) : detailFn,
    }));
    const uniqRows = (list) => {
      const seen = new Set();
      return (list || []).filter(r => {
        const key = rowKey(r);
        if (!key || seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    };

    const duplicateRows = (field) => {
      const groups = {};
      rows.forEach(r => {
        const raw = (r[field] || "").toString().trim();
        if (!raw) return;
        if (!groups[raw]) groups[raw] = [];
        groups[raw].push(r);
      });
      return Object.entries(groups)
        .filter(([, list]) => list.length > 1)
        .flatMap(([value, list]) => addDecor(list, `dup_${field}`, () => `${field}: ${value}`));
    };

    const dateOrderIssues = rows.flatMap(r => {
      // Cancelados/fracassados podem ter ordem de datas atípica por conta do encerramento — ignorar
      if (r.isCanceled || r.isFailed) return [];
      const modNrm = nrm(r.Modalidade || "");
      const idxs = r.NumRC ? getTLColsForMod(modNrm) : [0, 1, 2];
      let prevDate = null;
      let prevLabel = "";
      for (const idx of idxs) {
        const [label, key] = TL_COLS[idx] || [];
        if (!key) continue;
        // Indicação Analista é fase volátil — pode acontecer fora da sequência cronológica padrão.
        if (VOLATILE_TIMELINE_KEYS.has(key)) continue;
        const current = pd(r[key]);
        if (!current) continue;
        if (prevDate && current < prevDate) {
          return [{
            ...r,
            _dqIssue: "datas_incoerentes",
            _dqDetail: `${label}: ${current.toLocaleDateString("pt-BR")} antes de ${prevLabel}: ${prevDate.toLocaleDateString("pt-BR")}`,
          }];
        }
        prevDate = current;
        prevLabel = label;
      }
      return [];
    });

    const issueDefs = [
      {
        key: "datas_incoerentes",
        label: "Datas incoerentes",
        severity: "critical",
        color: "#d35400",
        desc: "Etapa posterior com data anterior a uma etapa já registada no processo.",
        rows: dateOrderIssues,
      },
      {
        key: "rc_sem_recebimento",
        label: "RC sem recebimento",
        severity: "critical",
        color: "#c0392b",
        desc: "Processos com RC aberta, mas sem data de recebimento da RC. Excluídos: cancelados, fracassados e RCs reprovadas para ajustes (aguardando reenvio).",
        rows: addDecor(rows.filter(r => {
          if (!r.NumRC || r.aberturaRC || r.isCanceled || r.isFailed) return false;
          const sd = (r.statusDet || "").toString().toLowerCase();
          if (sd.includes("reprovad")) return false; // RC reprovada para ajustes — aguarda reenvio com data nova
          return true;
        }), "rc_sem_recebimento", "RC sem data de recebimento."),
      },
      {
        key: "duplicidade_rc",
        label: "RC duplicada",
        severity: "critical",
        color: "#e74c3c",
        desc: "Mesmo número de RC encontrado em mais de um registo do recorte.",
        rows: duplicateRows("NumRC"),
      },
      {
        key: "ativo_sem_responsavel",
        label: "Ativo sem responsável",
        severity: "critical",
        color: "#ff3b30",
        desc: "Processos em andamento sem responsável activo definido.",
        rows: addDecor(rows.filter(r => r.emA && (!r.respAtivo || r.respAtivo === "—" || r.respAtivo === "N/A")), "ativo_sem_responsavel", "Processo activo sem responsável da etapa."),
      },
      {
        key: "rc_sem_modalidade",
        label: "RC sem modalidade",
        severity: "attention",
        color: "#ff9500",
        desc: "RCs sem modalidade impedem SLA e cronograma correctos. Cancelados/fracassados são excluídos.",
        rows: addDecor(rows.filter(r => r.NumRC && !(r.Modalidade || "").toString().trim() && !r.isCanceled && !r.isFailed), "rc_sem_modalidade", "RC sem modalidade informada."),
      },
      {
        key: "ativo_sem_entrega",
        label: "Ativo com RC sem entrega",
        severity: "attention",
        color: "#f39c12",
        desc: "Processos em andamento com RC, mas sem data prevista de entrega/início.",
        rows: addDecor(rows.filter(r => r.emA && r.NumRC && !r.dataEntrega), "ativo_sem_entrega", "Sem data de entrega prevista."),
      },
      {
        key: "cpl_sem_responsavel",
        label: "CPL sem responsável",
        severity: "attention",
        color: "#9b59b6",
        desc: "Processos com fluxo CPL sem responsável CPL ou pregoeiro preenchido.",
        rows: addDecor(rows.filter(r => r.emA && r.ehCPL && !(r.cplResp || r.Pregoeiro)), "cpl_sem_responsavel", "Fluxo CPL sem responsável identificado."),
      },
      {
        key: "duplicidade_sd",
        label: "Ticket pré-compra duplicado",
        severity: "attention",
        color: "#5dade2",
        desc: "Mesmo ticket de pré-compra encontrado em mais de um registo do recorte.",
        rows: duplicateRows("TicketSD"),
      },
    ].map(issue => ({ ...issue, rows: uniqRows(issue.rows), count: uniqRows(issue.rows).length }));

    const allIssueRows = uniqRows(issueDefs.flatMap(issue => issue.rows));
    const criticalRows = uniqRows(issueDefs.filter(issue => issue.severity === "critical").flatMap(issue => issue.rows));
    const attentionRows = uniqRows(issueDefs.filter(issue => issue.severity === "attention").flatMap(issue => issue.rows));

    const bySubarea = ["NCL", "CPL", "Scont", "Sem subárea"].map(label => {
      const match = label === "Sem subárea"
        ? allIssueRows.filter(r => !r.faseSubarea)
        : allIssueRows.filter(r => r.faseSubarea === label);
      const totalSub = label === "Sem subárea"
        ? rows.filter(r => !r.faseSubarea).length
        : rows.filter(r => r.faseSubarea === label).length;
      return {
        label,
        count: match.length,
        total: totalSub,
        pct: totalSub ? Math.round((match.length / totalSub) * 100) : 0,
        rows: match,
      };
    }).filter(item => item.total > 0);

    const maxPenalty = Math.max(rows.length * 6, 1);
    const appliedPenalty = (criticalRows.length * 6) + (attentionRows.length * 3);
    const score = Math.max(0, Math.min(100, Math.round(100 - ((appliedPenalty / maxPenalty) * 100))));

    return {
      totalRows: rows.length,
      activeRows: fBack.length,
      issueDefs,
      allIssueRows,
      criticalRows,
      attentionRows,
      totalIssues: allIssueRows.length,
      criticalCount: criticalRows.length,
      attentionCount: attentionRows.length,
      score,
      bySubarea,
      topExamples: allIssueRows.slice(0, 12),
    };
  }, [base, baseVis, fBack]);

  const allProcesses = useMemo(() => {
    let arr = [...fBase];
    if (hideSDConcluido) arr = arr.filter(r => !nrm(r.statusDet || r.status || "").includes("service desk concluido"));
    const { col, dir } = sortAll;
    arr.sort((a, b) => {
      let va = a[col], vb = b[col];
      if (typeof va === "string") va = nrm(va);
      if (typeof vb === "string") vb = nrm(vb);
      if (va == null) va = dir === "desc" ? -Infinity : Infinity;
      if (vb == null) vb = dir === "desc" ? -Infinity : Infinity;
      return dir === "desc" ? (vb > va ? 1 : vb < va ? -1 : 0) : (va > vb ? 1 : va < vb ? -1 : 0);
    });
    return arr;
  }, [fBase, sortAll, hideSDConcluido]);

  const allPage = useMemo(() => allProcesses.slice((pgAll-1)*pgAllSz, pgAll*pgAllSz), [allProcesses, pgAll, pgAllSz]);

  const procComp = useMemo(() => {
    if (!selComp) return [];
    if (selComp.tipo === "cpl") return cplBack.filter(r => r.Pregoeiro === selComp.nome || r.cplResp === selComp.nome);
    return fBack.filter(r => r.respNCL === selComp.nome);
  }, [selComp, cplBack, fBack]);

  const onClickComp = (nome, tipo) => setSelComp({ nome, tipo });
  const onClickProc = (r) => setSelProc(r);
  const onOpenTags = (r) => setTagModal(r);

  function buildTempBacklogRows(tempList) {
    const hoje = new Date();
    return (tempList || []).map((t, idx) => {
      const isCpl = t.pool === "cpl";
      const responsavel = t.responsavel || "";
      return {
        _tempDemand: true,
        ProcessKey: t.id || `TEMP_${idx}`,
        NumRC: "",
        TicketSD: t.ticket || "",
        Objeto: t.objeto || "",
        "Área Requisitante": t.area || "",
        Modalidade: t.modalidade || "",
        emA: true,
        isCanceled: false,
        isFailed: false,
        isConcluded: false,
        faseSubarea: isCpl ? "CPL" : "NCL",
        Comprador: isCpl ? "" : responsavel,
        Avaliador: "",
        respNCL: isCpl ? "" : responsavel,
        Pregoeiro: isCpl ? responsavel : "",
        cplResp: isCpl ? responsavel : "",
        aberturaSD: isCpl ? null : hoje,
        encSD: null,
        aberturaRC: hoje,
        statusDet: isCpl ? "DEMANDA TEMPORARIA CPL" : "EM ANALISE SERVICE DESK - TEMPORARIA",
        criticoNcl: false,
        criticoNclGestao: false,
        criticoCpl: false,
        LeadTime: 0,
      };
    });
  }

  // FIX: gerarRecs com pools separados + limpar
  function gerarRecs(tempList = tempDemandas) {
    const histBase = base;
    // Backlog a partir da base canônica enriquecida (datas do log + status
    // efetivo), para que a regra de "SD concluído sem RC" — que tira o processo
    // da carga do comprador — valha igual no balanceamento e no painel executivo.
    const allBack = [
      ...baseEnriched.filter(r => r.emA && !r.isCanceled && !r.isFailed),
      ...buildTempBacklogRows(tempList),
    ]; // backlog global sem filtros de UI + temporarias da sessao
    const opts = { ignorarModalidade: !!dem.ignorarMod || !dem.modalidade, anoAtual: new Date().getFullYear() };
    // Filtra do pool quem foi marcado como indisponível
    const ignSet = comprIgnorados;
    const poolNCL = compsNCLPool.filter(c => !ignSet.has(c));
    const poolCPL = compsCPLPool.filter(c => !ignSet.has(c));
    // Mantém TODOS os candidatos rankeados; UI decide se mostra top 3 ou tudo
    // Ordena por rankScore (total bruto, pode ser negativo p/ sobrecarregados)
    const ordenarRank = (a, b) => (b.rankScore ?? b.total) - (a.rankScore ?? a.total);
    // Duas passadas: a 1ª só extrai tocadosAno de cada candidato para obter a
    // média do pool; a 2ª pontua de verdade com essa média (componente de
    // equilíbrio anual do sG — quem já tocou mais que a média do time perde).
    const mediaTocados = (pool, tipo) => {
      const vals = pool.map(c => calcularScore(c, histBase, allBack, dem, tipo, opts).tocadosAno || 0);
      return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
    };
    const optsNCL = { ...opts, mediaTocadosPool: mediaTocados(poolNCL, "ncl") };
    const optsCPL = { ...opts, mediaTocadosPool: mediaTocados(poolCPL, "cpl") };
    const scNCL = poolNCL.map(c => ({ comp: c, ...calcularScore(c, histBase, allBack, dem, "ncl", optsNCL) })).sort(ordenarRank);
    const scCPL = poolCPL.map(c => ({ comp: c, ...calcularScore(c, histBase, allBack, dem, "cpl", optsCPL) })).sort(ordenarRank);
    setRecsNCL(scNCL);
    setRecsCPL(scCPL);
  }
  // Núcleo: adiciona N demandas temporárias para uma pessoa. Só responsavel+pool
  // são obrigatórios; área/modalidade/objeto são opcionais. origem registra de
  // quem a carteira foi redistribuída (férias), apenas para rótulo — não muda o JSON.
  function addTempDemandas(payload) {
    const responsavel = (payload.responsavel || "").trim();
    const pool = payload.pool || "ncl";
    if (!responsavel) return false;
    const qty = Math.max(1, Math.min(99, parseInt(payload.quantidade, 10) || 1));
    const novas = [];
    for (let k = 0; k < qty; k++) {
      novas.push({
        id: `TEMP_${Date.now()}_${Math.random().toString(36).slice(2, 7)}_${k}`,
        pool,
        responsavel,
        area: payload.area || "",
        modalidade: payload.modalidade || "",
        objeto: payload.objeto || (payload.origem ? `Redistribuição da carteira de ${payload.origem}` : "Demanda temporária"),
        origem: payload.origem || "",
        criadaEm: new Date().toISOString(),
      });
    }
    const next = [...tempDemandas, ...novas];
    setTempDemandas(next);
    gerarRecs(next);
    return true;
  }
  // Usada pelos cards de recomendação: herda os campos da demanda em edição.
  function registrarTempDemanda(responsavel, pool) {
    addTempDemandas({
      responsavel, pool,
      area: dem.area || "",
      modalidade: dem.ignorarMod ? "" : (dem.modalidade || ""),
      objeto: dem.objeto || "Demanda temporária",
    });
  }
  // Distribuição rápida (painel próprio): só a pessoa é obrigatória.
  function adicionarQuickTemp() {
    if (!quickTemp.responsavel) return;
    const ok = addTempDemandas({
      responsavel: quickTemp.responsavel,
      pool: quickTemp.pool || (compsCPLPool.includes(quickTemp.responsavel) && !compsNCLPool.includes(quickTemp.responsavel) ? "cpl" : "ncl"),
      area: quickTemp.area || "",
      modalidade: quickTemp.modalidade || "",
      objeto: quickTemp.objeto || "",
      quantidade: quickTemp.quantidade,
    });
    if (ok) setQuickTemp({ responsavel: "", pool: "", area: "", modalidade: "", objeto: "", quantidade: 1 });
  }
  // Conta a carteira ATIVA real de uma pessoa (base global, sem filtros de UI),
  // separada por pool — base para a redistribuição de férias.
  function contarCarteira(nome) {
    if (!nome) return { ncl: 0, cpl: 0 };
    const ativos = base.filter(r => r.emA && !r.isCanceled && !r.isFailed);
    const ncl = ativos.filter(r => r.respNCL === nome || r.Comprador === nome || r.Avaliador === nome).length;
    const cpl = ativos.filter(r => r.cplResp === nome || r.Pregoeiro === nome).length;
    return { ncl, cpl };
  }
  // Redistribui a carteira de quem está indisponível (férias) para outra pessoa.
  function redistribuirCarteira(origem, pool) {
    const key = `${origem}|${pool}`;
    const f = redistForm[key] || {};
    if (!f.receiver) return;
    // Sem qtd informada → transfere a carteira ativa inteira daquele pool
    const carteira = contarCarteira(origem);
    const qtyDefault = pool === "cpl" ? carteira.cpl : carteira.ncl;
    const ok = addTempDemandas({
      responsavel: f.receiver,
      pool,
      objeto: `Redistribuição da carteira de ${origem} (férias/indisponível)`,
      origem,
      quantidade: f.qty != null && f.qty !== "" ? f.qty : qtyDefault,
    });
    if (ok) setRedistForm(prev => { const n = { ...prev }; delete n[key]; return n; });
  }
  function removerTempDemanda(id) {
    const next = tempDemandas.filter(t => t.id !== id);
    setTempDemandas(next);
    if (recsNCL.length > 0 || recsCPL.length > 0) gerarRecs(next);
  }
  function limparTempDemandas() {
    setTempDemandas([]);
    if (recsNCL.length > 0 || recsCPL.length > 0) gerarRecs([]);
  }
  function limparDemanda() {
    setDem({ area: "", modalidade: "", objeto: "", ignorarMod: false });
    setRecsNCL([]);
    setRecsCPL([]);
    setVerRankingCompleto(false);
  }
  function toggleIgnorado(nome) {
    setComprIgnorados(prev => {
      const next = new Set(prev);
      if (next.has(nome)) next.delete(nome); else next.add(nome);
      return next;
    });
  }

  function crossFilter(key, val) { if (key === "resp") setFComp(val); if (key === "area") setFArea(val); if (key === "mod") setFMod(val); }
  function toggleSort(col) { setSortAll(prev => ({ col, dir: prev.col === col && prev.dir === "desc" ? "asc" : "desc" })); setPgAll(1); }

  const hasF = fAnoSD || fAnoRC || fMod || fArea || fSt || fStDet || fComp || fAval || fAnalista || fSubarea || fTag || fWatch;
  const activeFilters = useMemo(() => {
    const arr = [];
    if (fAnoSD) arr.push({ key: "fAnoSD", label: "Ano pré-compra", value: fAnoSD, clear: () => setFAnoSD("") });
    if (fAnoRC) arr.push({ key: "fAnoRC", label: "Ano RC", value: fAnoRC, clear: () => setFAnoRC("") });
    if (fMod) arr.push({ key: "fMod", label: "Modalidade", value: fMod.length > 18 ? fMod.slice(0,16)+"…" : fMod, clear: () => setFMod("") });
    if (fArea) arr.push({ key: "fArea", label: "Área", value: fArea.length > 18 ? fArea.slice(0,16)+"…" : fArea, clear: () => setFArea("") });
    if (fSt) arr.push({ key: "fSt", label: "Status", value: fSt, clear: () => setFSt("") });
    if (fStDet) arr.push({ key: "fStDet", label: "Status Det.", value: fStDet.length > 20 ? fStDet.slice(0,18)+"…" : fStDet, clear: () => setFStDet("") });
    if (fComp && !loginUser) arr.push({ key: "fComp", label: "Comprador", value: fComp, clear: () => setFComp("") });
    if (fAval && !loginUser) arr.push({ key: "fAval", label: "Avaliador", value: fAval, clear: () => setFAval("") });
    if (fAnalista) arr.push({ key: "fAnalista", label: "Analista (SCONT)", value: fAnalista, clear: () => setFAnalista("") });
    if (fSubarea) arr.push({ key: "fSubarea", label: "Subárea", value: fSubarea, clear: () => setFSubarea("") });
    if (fTag) arr.push({ key: "fTag", label: "Tag", value: fTag, clear: () => setFTag("") });
    if (fWatch) arr.push({ key: "fWatch", label: "Watchlist", value: WatchlistManager.count(), clear: () => setFWatch(false) });
    return arr;
  }, [fAnoSD, fAnoRC, fMod, fArea, fSt, fStDet, fComp, fAval, fAnalista, fSubarea, fTag, fWatch, watchVersion]);
  const renderBuscaPanel = () => (
    <div className="gaq-card" style={{ marginBottom: 16, border: "1px solid var(--gaq-blue)" }}>
      <div style={{ fontWeight: 600, color: "var(--gaq-blue)", fontSize: 13, marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}>
        <Icon name="search" size={14} color="var(--gaq-blue)" /> Busca por N RC, ticket pré-compra ou objeto
      </div>
      <input value={busca} onChange={e => handleBuscaChange(e.target.value)} placeholder="Digite o codigo RC, numero da pré-compra ou parte do objeto..."
        className="gaq-input" style={{ width: "100%", padding: "10px 14px", fontSize: 13 }} />
      {buscaDefer.length >= 3 && (buscaResults.length === 0
        ? <div style={{ fontSize: 12, color: "var(--gaq-text-3)", marginTop: 6 }}>Nenhum resultado para "{buscaDefer}".</div>
        : <div style={{ marginTop: 8 }}>
            <div style={{ fontSize: 11, color: "var(--gaq-text-3)", marginBottom: 6 }}>{buscaResults.length} resultado(s)</div>
            {buscaResults.map((r, i) => (
              <div key={i} style={{ background: "var(--gaq-bg-2)", borderRadius: "var(--gaq-r-sm)", padding: "10px 14px", marginBottom: 6, cursor: "pointer", border: "1px solid var(--gaq-line)", transition: "background .15s" }}
                onClick={() => setSelProc(r)}
                onMouseEnter={e => e.currentTarget.style.background = "var(--gaq-surface)"}
                onMouseLeave={e => e.currentTarget.style.background = "var(--gaq-bg-2)"}>
                <div style={{ display: "flex", justifyContent: "space-between", gap: 8 }}>
                  <div style={{ minWidth: 0, flex: 1 }}>
                    <div style={{ display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center", marginBottom: 3 }}>
                      <span style={{ fontSize: 9, color: "var(--gaq-text-3)", background: "var(--gaq-bg-2)", borderRadius: 3, padding: "0 4px" }}>{r.ProcessKey || "—"}</span>
                      <span style={{ fontWeight: 600, color: "var(--gaq-blue)", fontSize: 12 }}>{r.NumRC || "—"}</span>
                      {r.TicketSD && <span style={{ fontSize: 10, color: "var(--gaq-green)" }}>Pré-compra: {r.TicketSD}</span>}
                      <span style={{ fontSize: 10, color: "var(--gaq-ncl)" }}>{r.Modalidade || "—"}</span>
                      <span style={{ fontSize: 10, color: r.emA ? "var(--gaq-green)" : r.isConcluded ? "var(--gaq-blue)" : "var(--gaq-red)" }}>{r.status}</span>
                    </div>
                    <div style={{ fontSize: 11.5, color: "var(--gaq-text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{(r.Objeto || "").slice(0, 100) || "—"}</div>
                  </div>
                  <div style={{ flexShrink: 0, fontSize: 10, color: "var(--gaq-text-3)", display: "flex", alignItems: "center", gap: 3 }}>timeline <Icon name="chevR" size={10}/></div>
                </div>
              </div>
            ))}
          </div>
      )}
    </div>
  );

  // Sidebar nav groups config
  const sidebarGroups = isDiretor ? [
    { label: "Visão", items: [
      ["overview","grid","Visão Geral",null],
      ["dirHome","grid","Painel Diretoria",null],
      ["dirDiag","compass","Diagnóstico",diretorDiag ? diretorDiag.criticos.length : 0],
    ]},
    { label: "Processos", items: [
      ["processos","list","Processos",fBase.length],
      ["metro","subway","Linha do Metrô",null],
      ["dirTimeline","clock","Timeline",null],
      ["projetos","calendar","Acompanhamento de Projetos",null],
    ]},
    { label: "Riscos & Alertas", items: [
      ["alertaSD","doc","Alerta Pré-compra",alertSD.length],
      ["alertaRC","shield","Alerta RC",alertRC.length],
      ["fracionamento","trend","Fracionamento (MXM)",null],
    ]},
  ] : [
    { label: "Visão", items: [
      ...(isAdminMaster ? [["qualidade","settings","Qualidade de Dados",qualityData ? qualityData.totalIssues : 0]] : []),
      ["overview","grid","Visão Geral",null],
      ["executivo","spark","Executivo",null],
      ["operacional","chart","Operacional",null],
      ["fracionamento","trend","Fracionamento (MXM)",null],
    ]},
    { label: "Processos", items: [
      ["processos","list","Processos",fBase.length],
      ["metro","subway","Linha do Metrô",null],
      ["dirTimeline","clock","Timeline",null],
      ["cronograma","calendar","Cronograma",null],
      ["projetos","calendar","Acompanhamento de Projetos",null],
    ]},
    { label: "Riscos & Alertas", items: [
      ["alertas","alert","Alertas",(alertNCL.length+alertCPL.length)],
      ["alertaSD","doc","Alerta Pré-compra",alertSD.length],
      ["alertaRC","shield","Alerta RC",alertRC.length],
      ["mapeamento","map","Mapeamento",null],
    ]},
    ...(!loginUser && !loginArea ? [{ label: "Gestão", items: [
      ["gestao","users","Performance",null],
      ["planejamento","compass","Planejamento",null],
      ["pantanal","trend","Pantanal",null],
    ]}] : []),
    ...(isAdminMaster ? [{ label: "Administração", items: [
      ["geradorSenha","key","Gerador de Senha",null],
      ["relCustom","doc","Relatório Customizado",null],
    ]}] : []),
  ];

  // Tab title map
  const tabTitles = {
    upload: "Upload", overview: "Visão Geral", executivo: "Executivo", operacional: "Operacional", qualidade: "Qualidade de Dados",
    processos: "Processos", metro: "Linha do Metrô", dirTimeline: "Timeline", cronograma: "Cronograma",
    alertas: "Alertas", alertaSD: "Alerta Pré-compra", alertaRC: "Alerta RC", mapeamento: "Mapeamento",
    gestao: "Performance", planejamento: "Planejamento", pantanal: "Pantanal", nova: "Nova Demanda",
    bal: "Balanceamento", dirHome: "Painel Diretoria", dirDiag: "Diagnóstico",
    geradorSenha: "Gerador de Senha", relCustom: "Relatório Customizado",
    projetos: "Acompanhamento de Projetos",
    fracionamento: "Fracionamento (MXM)",
  };

  return (
    <div style={{ fontFamily: "var(--gaq-font)", minHeight: "100vh", background: "var(--gaq-bg)", color: "var(--gaq-text)", display: "flex", transition: "background .3s" }}>

      {/* ── OVERLAY DE LOGIN ─────────────────────────────────────────────── */}
      {!loginOk && (
        <div style={{ position: "fixed", inset: 0, background: "var(--gaq-bg)", zIndex: 9999, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
          <div style={{ background: "var(--gaq-surface)", borderRadius: "var(--gaq-r-xl)", padding: "40px 44px", width: "min(420px, 94vw)", boxShadow: "var(--gaq-shadow-3)", border: "1px solid var(--gaq-line)" }}>
            <div style={{ textAlign: "center", marginBottom: 28 }}>
              <div style={{
                width: 48, height: 48, borderRadius: 14, margin: "0 auto 16px",
                background: "linear-gradient(135deg, var(--gaq-blue), var(--gaq-cpl))",
                display: "flex", alignItems: "center", justifyContent: "center",
                color: "#fff", fontWeight: 700, fontSize: 20, letterSpacing: "-0.02em",
                boxShadow: "0 4px 12px rgba(0,113,227,0.3)"
              }}>G</div>
              <div style={{ fontSize: 24, fontWeight: 700, color: "var(--gaq-text)", lineHeight: 1.2, marginBottom: 4, letterSpacing: "-0.025em" }}>Painel GAQ</div>
              <div style={{ fontSize: 13, color: "var(--gaq-text-3)" }}>Sistema de Compras · Uso Interno</div>
            </div>
            {/* Tabs de tipo de acesso */}
            <div style={{ display: "flex", gap: 0, marginBottom: 22, borderRadius: 10, overflow: "hidden", border: "1px solid var(--gaq-line-2)" }}>
              {[["gaq","GAQ"],["area","Área"],["diretor","Diretoria"]].map(([tab, label], idx, arr) => (
                <button key={tab} onClick={() => { setLoginTab(tab); setLoginErroMsg(false); setAreaPassErr(false); setDirPassErr(false); }}
                  style={{ flex: 1, padding: "10px 0", fontSize: 12, fontWeight: 600, cursor: "pointer", border: "none", borderRight: idx < arr.length - 1 ? "1px solid var(--gaq-line-2)" : "none",
                    background: loginTab === tab ? "var(--gaq-blue)" : "var(--gaq-surface)",
                    color: loginTab === tab ? "#fff" : "var(--gaq-text-3)", transition: "all .15s" }}>
                  {label}
                </button>
              ))}
            </div>
            {loginTab === "gaq" ? (
              <>
                <div style={{ marginBottom: 18 }}>
                  <label style={{ fontSize: 12, fontWeight: 600, color: "var(--gaq-text-2)", display: "block", marginBottom: 6 }}>Senha de acesso</label>
                  <input type="password" value={loginInput} autoFocus
                    onChange={e => { setLoginInput(e.target.value); setLoginErroMsg(false); }}
                    onKeyDown={e => { if (e.key === "Enter") handleLogin(); }}
                    placeholder="Digite sua senha..."
                    className="gaq-input"
                    style={{ width: "100%", padding: "12px 14px", height: 44, fontSize: 15, borderColor: loginErroMsg ? "var(--gaq-red)" : undefined }}
                  />
                  <div style={{ fontSize: 11, color: "var(--gaq-text-3)", marginTop: 6 }}>Admin, compradores e Admin Master usam este acesso.</div>
                  {loginErroMsg && <div style={{ fontSize: 12, color: "var(--gaq-red)", marginTop: 6, fontWeight: 600 }}>Senha incorreta. Tente novamente.</div>}
                </div>
                <button onClick={handleLogin} className="gaq-btn is-primary"
                  style={{ width: "100%", padding: "13px 0", fontSize: 14, justifyContent: "center" }}>
                  Entrar
                </button>
              </>
            ) : loginTab === "area" ? (
              <>
                <div style={{ background: "var(--gaq-purple-t)", border: "1px solid rgba(175,82,222,0.2)", borderRadius: 10, padding: "10px 14px", marginBottom: 14, fontSize: 12, color: "var(--gaq-purple)" }}>
                  Acesso restrito aos processos da sua área.
                </div>
                <div style={{ marginBottom: 18 }}>
                  <label style={{ fontSize: 12, fontWeight: 600, color: "var(--gaq-text-2)", display: "block", marginBottom: 6 }}>Senha da Área</label>
                  <input type="password" value={areaPassInput} autoFocus
                    onChange={e => { setAreaPassInput(e.target.value); setAreaPassErr(false); }}
                    onKeyDown={e => { if (e.key === "Enter") handleAreaLogin(); }}
                    placeholder="Senha fornecida pelo GAQ..."
                    className="gaq-input"
                    style={{ width: "100%", padding: "12px 14px", height: 44, fontSize: 15, borderColor: areaPassErr ? "var(--gaq-red)" : undefined }}
                  />
                  {areaPassErr && <div style={{ fontSize: 12, color: "var(--gaq-red)", marginTop: 6, fontWeight: 600 }}>Senha incorreta. Contate o GAQ.</div>}
                </div>
                <button onClick={handleAreaLogin} className="gaq-btn is-primary"
                  style={{ width: "100%", padding: "13px 0", fontSize: 14, justifyContent: "center", background: "var(--gaq-purple)" }}>
                  Acessar Área
                </button>
              </>
            ) : (
              <>
                <div style={{ background: "var(--gaq-green-t)", border: "1px solid rgba(52,199,89,0.2)", borderRadius: 10, padding: "10px 14px", marginBottom: 14, fontSize: 12, color: "#1e8e3e" }}>
                  Visão simplificada para Diretoria.
                </div>
                <div style={{ marginBottom: 18 }}>
                  <label style={{ fontSize: 12, fontWeight: 600, color: "var(--gaq-text-2)", display: "block", marginBottom: 6 }}>Senha da Diretoria</label>
                  <input type="password" value={dirPassInput} autoFocus
                    onChange={e => { setDirPassInput(e.target.value); setDirPassErr(false); }}
                    onKeyDown={e => { if (e.key === "Enter") handleDiretorLogin(); }}
                    placeholder="Senha fornecida pelo GAQ..."
                    className="gaq-input"
                    style={{ width: "100%", padding: "12px 14px", height: 44, fontSize: 15, borderColor: dirPassErr ? "var(--gaq-red)" : undefined }}
                  />
                  {dirPassErr && <div style={{ fontSize: 12, color: "var(--gaq-red)", marginTop: 6, fontWeight: 600 }}>Senha incorreta. Contate o GAQ.</div>}
                </div>
                <button onClick={handleDiretorLogin} className="gaq-btn is-primary"
                  style={{ width: "100%", padding: "13px 0", fontSize: 14, justifyContent: "center", background: "var(--gaq-green)" }}>
                  Acessar Diretoria
                </button>
              </>
            )}
            <p style={{ fontSize: 11, color: 'var(--gaq-text-3)', lineHeight: 1.5, margin: '16px 0 0' }}>Ao entrar, o perfil e o horário são registrados apenas neste navegador.</p>
          </div>
        </div>
      )}

      {/* ── SIDEBAR ─────────────────────────────────────────────── */}
      {sidebarOpen && <button className="gaq-sidebar-backdrop" aria-label="Fechar menu" onClick={() => setSidebarOpen(false)} />}
      <aside id="gaq-navigation" className={`pres-hide gaq-sidebar${sidebarOpen ? ' is-open' : ' is-collapsed'}`} inert={sidebarOpen && loginOk ? undefined : ''} onClick={event => { if (event.target.closest('.gaq-side-item') && window.matchMedia('(max-width: 900px)').matches) setSidebarOpen(false); }} style={{
        width: sidebarOpen ? 232 : 0, flex: "none",
        background: "var(--gaq-bg-2)",
        borderRight: sidebarOpen ? "1px solid var(--gaq-line)" : "none",
        display: "flex", flexDirection: "column",
        height: "100vh", position: "sticky", top: 0,
        overflow: "hidden", transition: "width .25s cubic-bezier(.4,0,.2,1)",
        zIndex: 20,
      }}>
        {/* Search shortcut */}
        <div style={{ padding: "16px 14px 12px", minWidth: 232 }}>
          <div className="gaq-search" style={{ height: 30 }} onClick={() => setSpotlight(true)}>
            <Icon name="search" size={14} color="var(--gaq-text-3)" />
            <input placeholder="Buscar processo, RC..." readOnly style={{ cursor: "pointer" }} />
            <span style={{
              fontSize: 10, color: "var(--gaq-text-3)", fontWeight: 600,
              border: "1px solid var(--gaq-line)", padding: "1px 5px",
              borderRadius: 4, background: "var(--gaq-surface)"
            }}>Ctrl+K</span>
          </div>
        </div>

        {/* Nav groups */}
        <nav style={{ flex: 1, overflowY: "auto", padding: "0 10px 10px", minWidth: 232 }}>
          {sidebarGroups.map(g => (
            <div key={g.label} style={{ marginBottom: 14 }}>
              <div style={{
                fontSize: 10.5, fontWeight: 600, color: "var(--gaq-text-3)",
                padding: "6px 10px 4px", letterSpacing: "0.06em", textTransform: "uppercase"
              }}>{g.label}</div>
              {g.items.map(([id, ic, lbl, badge]) => (
                <button key={id}
                  className={`gaq-side-item${aba === id ? " is-active" : ""}`}
                  onClick={() => { setAba(id); setDrillDown(null); }}>
                  <Icon name={ic} size={16} stroke={aba === id ? 1.9 : 1.6} />
                  <span>{lbl}</span>
                  {badge > 0 && <span className={`gaq-badge ${id.startsWith("alerta") || id === "alertas" ? "" : "is-soft"}`}>{badge > 99 ? "99+" : badge}</span>}
                </button>
              ))}
            </div>
          ))}
          {/* Admin-only: Upload & extras */}
          {!loginUser && !loginArea && !isDiretor && (
            <div style={{ marginBottom: 14 }}>
              <div style={{
                fontSize: 10.5, fontWeight: 600, color: "var(--gaq-text-3)",
                padding: "6px 10px 4px", letterSpacing: "0.06em", textTransform: "uppercase"
              }}>Admin</div>
              <button className={`gaq-side-item${aba === "upload" ? " is-active" : ""}`}
                onClick={() => { setAba("upload"); setDrillDown(null); }}>
                <Icon name="upload" size={16} stroke={aba === "upload" ? 1.9 : 1.6} />
                <span>Upload</span>
              </button>
              <button className={`gaq-side-item${aba === "nova" ? " is-active" : ""}`}
                onClick={() => { setAba("nova"); setDrillDown(null); }}>
                <Icon name="plus" size={16} stroke={aba === "nova" ? 1.9 : 1.6} />
                <span>Nova Demanda</span>
              </button>
              <button className={`gaq-side-item${aba === "bal" ? " is-active" : ""}`}
                onClick={() => { setAba("bal"); setDrillDown(null); }}>
                <Icon name="chart" size={16} stroke={aba === "bal" ? 1.9 : 1.6} />
                <span>Balanceamento</span>
              </button>
            </div>
          )}
        </nav>

        {/* Footer (user) */}
        <div style={{
          padding: "10px 12px", borderTop: "1px solid var(--gaq-line)",
          display: "flex", alignItems: "center", gap: 10, minWidth: 232,
        }}>
          <div style={{
            width: 30, height: 30, borderRadius: "50%",
            background: isDiretor ? "linear-gradient(135deg,#34c759,#148f77)" : loginArea ? "linear-gradient(135deg,#af52de,#8e44ad)" : isAdminMaster ? "linear-gradient(135deg,#f59e0b,#d97706)" : "linear-gradient(135deg,#0071e3,#5e5ce6)",
            color: "#fff", display: "flex", alignItems: "center", justifyContent: "center",
            fontWeight: 600, fontSize: 12, flexShrink: 0,
          }}>{isDiretor ? "D" : loginArea ? "A" : isAdminMaster ? "M" : loginUser ? loginUser.charAt(0).toUpperCase() : "G"}</div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 12.5, fontWeight: 600, color: "var(--gaq-text)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
              {isDiretor ? loginDiretor.nome : loginArea ? `Área: ${loginAreaDisp}` : loginUser ? loginUser : "GAQ Admin"}
            </div>
            <div style={{ fontSize: 11, color: "var(--gaq-text-3)" }}>
              {isDiretor ? "Diretoria" : loginArea ? "Perfil Área" : loginUser ? "Comprador" : "Administrador"}
            </div>
            {isAdminMaster && <div style={{ fontSize: 10, color: "#d97706", fontWeight: 700 }}>Qualidade de Dados liberada</div>}
          </div>
          <button className="gaq-icon-btn" title="Sair" onClick={handleLogout}><Icon name="logout" size={15}/></button>
        </div>
      </aside>

      {/* ── MAIN CONTENT AREA ────────────────────────────────────────── */}
      <div className="gaq-main" inert={loginOk ? undefined : ''} style={{ flex: 1, display: "flex", flexDirection: "column", minWidth: 0, minHeight: "100vh" }}>

        {/* ── TOP BAR ── */}
        <header className="pres-hide" style={{
          height: 56, padding: "0 24px",
          borderBottom: "1px solid var(--gaq-line)",
          background: dark ? "rgba(13,25,41,0.82)" : "rgba(255,255,255,0.72)",
          backdropFilter: "saturate(180%) blur(12px)",
          WebkitBackdropFilter: "saturate(180%) blur(12px)",
          display: "flex", alignItems: "center", gap: 16,
          position: "sticky", top: 0, zIndex: 5, flexShrink: 0,
        }}>
          {/* Toggle sidebar */}
          <button className="gaq-icon-btn" aria-controls="gaq-navigation" aria-expanded={sidebarOpen} onClick={() => setSidebarOpen(o => !o)} title={sidebarOpen ? "Recolher menu" : "Expandir menu"}>
            <Icon name={sidebarOpen ? "sidebar" : "menu"} size={17}/>
          </button>

          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
              <h1 style={{ margin: 0, fontSize: 17, fontWeight: 600, letterSpacing: "-0.012em" }}>{tabTitles[aba] || aba}</h1>
              {base.length > 0 && <span className="gaq-meta">{fBase.length.toLocaleString("pt-BR")} registros</span>}
            </div>
          </div>

          <div className="gaq-header-actions" style={{ display: "flex", alignItems: "center", gap: 4 }}>
            {/* Sync */}
            {syncInfo && syncInfo.state !== "idle" && (
              <span className="gaq-meta gaq-sync-label" title={syncInfo.msg} style={{ marginRight: 4, fontSize: 10 }}>
                {syncInfo.state === "ok" ? "Verificado" : syncInfo.state === "loading" ? "Verificando..." : "Ver aviso"}
              </span>
            )}
            <button className="gaq-icon-btn" title="Sincronizar as quatro bases" disabled={syncInfo.state === 'loading'} onClick={() => { setAutoSync(true); syncNow(true); }}><Icon name="refresh" size={16}/></button>
            {!loginUser && !loginArea && !isDiretor && <button className="gaq-btn" onClick={() => setAccessOpen(true)}>Acessos locais</button>}
            {/* Notifications */}
            {loginUser && notifs.length > 0 && (
              <button className="gaq-icon-btn" title="Notificações" style={{ position: "relative" }} onClick={() => setShowNotifPanel(v => !v)}>
                <Icon name="bell" size={16}/>
                <span style={{ position: "absolute", top: 6, right: 6, width: 7, height: 7, borderRadius: 999, background: "var(--gaq-red)", border: "2px solid var(--gaq-bg)" }}/>
              </button>
            )}
            {/* Watchlist */}
            {base.length > 0 && (
              <button className="gaq-icon-btn" title={`Acompanhamento${watchChanges.length > 0 ? ` (${watchChanges.length})` : ""}`}
                onClick={() => setShowWatchPanel(p => !p)}>
                <Icon name="eye" size={16}/>
              </button>
            )}
            {/* Theme */}
            <button className="gaq-icon-btn" title={dark ? "Tema Claro" : "Tema Escuro"} onClick={() => setDark(d => !d)}>
              <Icon name={dark ? "sun" : "moon"} size={16}/>
            </button>
            {/* Presentation mode */}
            {base.length > 0 && (
              <button className="gaq-icon-btn" title={presMode ? "Sair Apresentação" : "Apresentação"} onClick={() => setPresMode(p => !p)}>
                <Icon name="grid" size={16}/>
              </button>
            )}
            <div style={{ width: 1, height: 22, background: "var(--gaq-line)", margin: "0 6px" }}/>
            {/* Export */}
            <button className="gaq-btn" onClick={() => setSpotlight(true)}>
              <Icon name="search" size={14}/> Buscar
            </button>
          </div>
        </header>

        {syncInfo.state === 'error' && <div role="alert" className="gaq-data-notice">{syncInfo.msg} {base.length > 0 ? 'Mantivemos os dados já carregados.' : 'Os indicadores estarão disponíveis após carregar as bases.'}<button className="gaq-btn" onClick={() => syncNow(true)}>Tentar novamente</button></div>}
        {syncInfo.state === 'loading' && !base.length && <div role="status" className="gaq-data-notice">Carregando os processos e históricos completos…</div>}
        {syncInfo.state === 'warn' && <div role="status" className="gaq-data-notice">{syncInfo.msg}</div>}
        {/* ── FILTER BAR ── */}
        {base.length > 0 && (
          <div className="pres-hide" style={{
            borderBottom: "1px solid var(--gaq-line)",
            background: "var(--gaq-bg-2)",
            padding: "10px 24px",
            display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", flexShrink: 0,
          }}>
            <button className="gaq-btn" style={{ padding: "6px 12px" }} onClick={() => setShowFilters(f => !f)}>
              <Icon name="filter" size={14}/> Filtros
              {activeFilters.length > 0 && <span className="gaq-badge is-blue" style={{ marginLeft: 4 }}>{activeFilters.length}</span>}
            </button>
            <div style={{ width: 1, height: 18, background: "var(--gaq-line)", margin: "0 4px" }}/>

            {/* Quick filter chips */}
            {[
              { label: "Modalidade", value: fMod, onChange: setFMod, opts: mods },
              { label: "Área", value: fArea, onChange: setFArea, opts: areas },
              { label: "Status", value: fSt, onChange: setFSt, opts: sts },
              ...(!loginUser ? [{ label: "Comprador", value: fComp, onChange: setFComp, opts: todosComp }] : []),
            ].map(({ label, value, onChange, opts }) => (
              <div key={label} style={{ position: "relative" }}>
                <select value={value} onChange={e => onChange(e.target.value)}
                  className={`gaq-chip${value ? " is-active" : ""}`}
                  style={{ appearance: "none", cursor: "pointer", paddingRight: 22, border: value ? "1px solid var(--gaq-blue)" : undefined }}>
                  <option value="">{label}</option>
                  {opts.map(o => <option key={o} value={o}>{o.length > 22 ? o.slice(0,20)+"..." : o}</option>)}
                </select>
              </div>
            ))}

            {/* Active filter tags */}
            {activeFilters.length > 0 && activeFilters.map(f => (
              <span key={f.key} className="gaq-chip-filter">
                <span style={{ opacity: 0.7, fontWeight: 500, fontSize: 11 }}>{f.label}: {f.value}</span>
                <span className="x" onClick={(e) => { e.stopPropagation(); f.clear(); }}><Icon name="x" size={10}/></span>
              </span>
            ))}

            {activeFilters.length > 0 && (
              <button onClick={() => { setFAnoSD(""); setFAnoRC(""); setFMod(""); setFArea(""); setFSt(""); setFStDet(""); setFComp(""); setFAval(""); setFAnalista(""); setFSubarea(""); setFTag(""); setFWatch(false); try{localStorage.removeItem("painel_gaq_filters");}catch{} }}
                style={{ padding: "4px 10px", borderRadius: 6, border: "none", color: "var(--gaq-red)", background: "rgba(255,59,48,0.08)", cursor: "pointer", fontSize: 11, fontWeight: 600 }}>
                Limpar tudo
              </button>
            )}

            <div style={{ marginLeft: "auto", display: "flex", gap: 12, alignItems: "center" }}>
              <span className="gaq-meta gaq-num"><b style={{ color: "var(--gaq-text)" }}>{fBase.length.toLocaleString("pt-BR")}</b> registros</span>
              <span className="gaq-meta gaq-num" style={{ color: "var(--gaq-blue)" }}><b>{fBack.length}</b> em and.</span>
            </div>
          </div>
        )}

        {/* Expandable filter panel */}
        {base.length > 0 && (
          <div className="pres-hide" style={{ maxHeight: showFilters ? 200 : 0, overflow: "hidden", transition: "max-height .25s ease-in-out, opacity .2s", opacity: showFilters ? 1 : 0, borderBottom: showFilters ? "1px solid var(--gaq-line)" : "none", flexShrink: 0 }}>
            <div style={{ padding: "12px 24px 14px", display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: 8, background: "var(--gaq-bg-2)" }}>
              <Sel label="Ano pré-compra" value={fAnoSD} onChange={setFAnoSD} opts={anosSD} />
              <Sel label="Ano RC" value={fAnoRC} onChange={setFAnoRC} opts={anosRC} />
              <Sel label="Modalidade" value={fMod} onChange={setFMod} opts={mods} />
              <Sel label="Área" value={fArea} onChange={setFArea} opts={areas} />
              <Sel label="Status" value={fSt} onChange={setFSt} opts={sts} />
              <Sel label="Status Det." value={fStDet} onChange={setFStDet} opts={stsDet} />
              {!loginUser && <Sel label="Comprador" value={fComp} onChange={setFComp} opts={todosComp} />}
              {!loginUser && <Sel label="Avaliador" value={fAval} onChange={setFAval} opts={todosAval} />}
              {loginUser && (
                <div style={{ display: "flex", flexDirection: "column", gap: 2, gridColumn: "span 2" }}>
                  <label style={{ fontSize: 10, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase", letterSpacing: .4 }}>Perfil ativo</label>
                  <div style={{ padding: "5px 10px", borderRadius: 6, background: "rgba(52,199,89,0.08)", border: "1px solid rgba(52,199,89,0.2)", fontSize: 12, color: "var(--gaq-green)", fontWeight: 700, height: 28, display: "flex", alignItems: "center", gap: 6 }}>
                    {loginUser} — visualizacao restrita ao seu perfil
                  </div>
                </div>
              )}
              {loginArea && (
                <div style={{ display: "flex", flexDirection: "column", gap: 2, gridColumn: "span 2" }}>
                  <label style={{ fontSize: 10, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase", letterSpacing: .4 }}>Area ativa</label>
                  <div style={{ padding: "5px 10px", borderRadius: 6, background: "rgba(175,82,222,0.08)", border: "1px solid rgba(175,82,222,0.2)", fontSize: 12, color: "var(--gaq-purple)", fontWeight: 700, height: 28, display: "flex", alignItems: "center", gap: 6 }}>
                    Area: {loginAreaDisp} — filtro bloqueado nesta area
                  </div>
                </div>
              )}
              <Sel label="Analista (SCONT)" value={fAnalista} onChange={setFAnalista} opts={todosAnalistas} />
              <Sel label="Subarea" value={fSubarea} onChange={setFSubarea} opts={["NCL","CPL","Scont"]} />
              <Sel label="Tag" value={fTag} onChange={setFTag} opts={allTagsList} />
              <div style={{ display: "flex", flexDirection: "column", gap: 2, justifyContent: "flex-end" }}>
                <label style={{ fontSize: 10, color: "var(--gaq-text-3)", fontWeight: 600, textTransform: "uppercase", letterSpacing: .4 }}>Watchlist</label>
                <button onClick={() => setFWatch(w => !w)} title="Filtrar apenas processos na Watchlist"
                  className={`gaq-btn${fWatch ? " is-primary" : ""}`}
                  style={{ height: 28, fontSize: 12, justifyContent: "center", background: fWatch ? "var(--gaq-orange)" : undefined, borderColor: fWatch ? "var(--gaq-orange)" : undefined }}>
                  <Icon name="eye" size={12}/> {fWatch ? `Ativa (${WatchlistManager.count()})` : "Filtrar"}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ── SCROLLABLE CONTENT ── */}
        <div style={{ flex: 1, minHeight: 0, padding: "24px 24px 36px" }}>

        {/* ════════════ ABA: DIRETORIA HOME ════════════ */}
        {aba === "dirHome" && isDiretor && base.length > 0 && <DirHomeScreen base={base} dirAreaSel={dirAreaSel} dirExpandNCL={dirExpandNCL} dirExpandNCLScont={dirExpandNCLScont} dirScopedBase={dirScopedBase} dirShowSDList={dirShowSDList} loginDiretor={loginDiretor} setDirAreaSel={setDirAreaSel} setDirExpandNCL={setDirExpandNCL} setDirExpandNCLScont={setDirExpandNCLScont} setDirShowSDList={setDirShowSDList} setSelProc={setSelProc} />}

        {/* ════════════ ABA: DIAGNÓSTICO (DIRETORIA) ════════════ */}
        {aba === "dirDiag" && isDiretor && diretorDiag && <DirDiagScreen dirAreaSel={dirAreaSel} diretorDiag={diretorDiag} setDirTLDrill={setDirTLDrill} setSelProc={setSelProc} />}

        {/* ════════════ ABA: TIMELINE (DIRETORIA) ════════════ */}
        {aba === "dirTimeline" && base.length > 0 && <DirTimelineScreen base={base} dirTLDrill={dirTLDrill} phaseIntervals={phaseIntervals} setDirTLDrill={setDirTLDrill} setSelProc={setSelProc} />}

        {/* ════════════ ABA: UPLOAD ════════════ */}
        {isDiretor && dirTLDrill && aba !== "dirTimeline" && <>
          <div onClick={() => setDirTLDrill(null)} style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.5)", zIndex: 9998 }} />
          <div style={{ position: "fixed", top: "5vh", left: "5vw", right: "5vw", bottom: "5vh", background: "var(--bg, #fff)", borderRadius: 16, boxShadow: "0 12px 48px rgba(0,0,0,0.3)", zIndex: 9999, display: "flex", flexDirection: "column", overflow: "hidden" }}>
            <div style={{ padding: "16px 24px", borderBottom: "2px solid var(--border2)", display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0, background: "var(--card)" }}>
              <div>
                <div style={{ fontWeight: 800, fontSize: 16, color: "#1a5276" }}>{dirTLDrill.title}</div>
                <div style={{ fontSize: 12, color: "var(--text3)", marginTop: 2 }}>{dirTLDrill.procs.length} processo(s) · Clique em uma linha para ver detalhes</div>
              </div>
              <button onClick={() => setDirTLDrill(null)} style={{ background: "#e74c3c15", border: "1px solid #e74c3c33", color: "#e74c3c", cursor: "pointer", fontSize: 14, fontWeight: 700, borderRadius: 8, padding: "6px 16px", transition: "background .15s" }}
                onMouseEnter={e => e.currentTarget.style.background = "#e74c3c25"}
                onMouseLeave={e => e.currentTarget.style.background = "#e74c3c15"}>Fechar</button>
            </div>
            <div style={{ flex: 1, overflowY: "auto", padding: "0 4px" }}>
              {dirTLDrill.procs.length === 0
                ? <div style={{ textAlign: "center", padding: 40, color: "var(--text3)", fontSize: 14 }}>Nenhum processo nesta seleção.</div>
                : (() => {
                  const hasPhaseCtx = !!dirTLDrill.phaseCtx;
                  const delayMap = new Map();
                  if (hasPhaseCtx && dirTLDrill.phaseCtx.procsComDelay) {
                    dirTLDrill.phaseCtx.procsComDelay.forEach(x => delayMap.set(x.proc, x));
                  }
                  const sortedProcs = hasPhaseCtx
                    ? [...dirTLDrill.procs].sort((a, b) => {
                        const da = delayMap.get(a), db = delayMap.get(b);
                        if (da && db) { if (da.atrasoFase !== db.atrasoFase) return da.atrasoFase ? -1 : 1; return (db.diasNaFase || 0) - (da.diasNaFase || 0); }
                        return (b.diasTotais || 0) - (a.diasTotais || 0);
                      })
                    : [...dirTLDrill.procs].sort((a, b) => (b.diasTotais || 0) - (a.diasTotais || 0));
                  const slaFaseLabel = hasPhaseCtx && dirTLDrill.phaseCtx.nextSLA ? `${dirTLDrill.phaseCtx.nextSLA} d.u.` : "—";
                  const totalAtrasoFase = hasPhaseCtx ? [...delayMap.values()].filter(x => x.atrasoFase).length : 0;
                  const totalOkFase = hasPhaseCtx ? dirTLDrill.procs.length - totalAtrasoFase : 0;
                  return <>
                    {hasPhaseCtx && (
                      <div style={{ display: "flex", gap: 12, padding: "12px 16px", borderBottom: "1px solid var(--border2)", background: "var(--card2)", flexWrap: "wrap", alignItems: "center" }}>
                        <div style={{ fontSize: 11, fontWeight: 700, color: "var(--text2)" }}>SLA da etapa: <span style={{ color: "#1a5276" }}>{slaFaseLabel}</span></div>
                        <div style={{ fontSize: 11 }}><span style={{ background: "#27ae6022", color: "#27ae60", borderRadius: 4, padding: "2px 8px", fontWeight: 700 }}>{totalOkFase} no prazo da etapa</span></div>
                        <div style={{ fontSize: 11 }}><span style={{ background: totalAtrasoFase > 0 ? "#e74c3c" : "#27ae6022", color: totalAtrasoFase > 0 ? "#fff" : "#27ae60", borderRadius: 4, padding: "2px 8px", fontWeight: 700 }}>{totalAtrasoFase} em atraso na etapa</span></div>
                      </div>
                    )}
                    <div className="gaq-table-scroll"><table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                      <thead>
                        <tr style={{ background: "var(--card2)", borderBottom: "2px solid var(--border2)", position: "sticky", top: 0, zIndex: 1 }}>
                          <th style={{ padding: "10px 8px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase" }}>Nº RC</th>
                          <th style={{ padding: "10px 8px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase" }}>Ticket pré-compra</th>
                          <th style={{ padding: "10px 8px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase", maxWidth: 220 }}>Objeto</th>
                          <th style={{ padding: "10px 8px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase" }}>Área</th>
                          <th style={{ padding: "10px 8px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase" }}>Comprador</th>
                          {hasPhaseCtx && <>
                            <th style={{ padding: "10px 8px", textAlign: "center", fontSize: 10, fontWeight: 700, color: "#8e44ad", textTransform: "uppercase" }}>Dias na Fase</th>
                            <th style={{ padding: "10px 8px", textAlign: "center", fontSize: 10, fontWeight: 700, color: "#8e44ad", textTransform: "uppercase" }}>SLA Fase</th>
                            <th style={{ padding: "10px 8px", textAlign: "center", fontSize: 10, fontWeight: 700, color: "#8e44ad", textTransform: "uppercase" }}>Status Fase</th>
                          </>}
                          <th style={{ padding: "10px 8px", textAlign: "center", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase" }}>Dias (RC)</th>
                          <th style={{ padding: "10px 8px", textAlign: "center", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase" }}>Prazo Geral</th>
                          <th style={{ padding: "10px 8px", textAlign: "center", fontSize: 10, fontWeight: 700, color: "var(--text3)", textTransform: "uppercase" }}>Status Geral</th>
                        </tr>
                      </thead>
                      <tbody>
                        {sortedProcs.map((r, i) => {
                          const late = r.atrasoGeral;
                          const cor = late ? "#c0392b" : (r.diasTotais || 0) > r.prazoGeral * 0.7 ? "#e67e22" : "#27ae60";
                          const phDelay = delayMap.get(r);
                          return (
                            <tr key={r.ProcessKey || i} onClick={() => { setDirTLDrill(null); setSelProc(r); }} style={{ cursor: "pointer", borderBottom: "1px solid var(--border2)", transition: "background .15s", background: phDelay && phDelay.atrasoFase ? "#fdecea44" : "transparent" }}
                              onMouseEnter={e => e.currentTarget.style.background = "var(--hover)"}
                              onMouseLeave={e => e.currentTarget.style.background = phDelay && phDelay.atrasoFase ? "#fdecea44" : "transparent"}>
                              <td style={{ padding: "9px 8px", fontWeight: 700, color: "#2e86c1" }}>{r.NumRC || "—"}</td>
                              <td style={{ padding: "9px 8px", color: "#27ae60", fontSize: 11 }}>{r.TicketSD || "—"}</td>
                              <td style={{ padding: "9px 8px", color: "var(--text)", maxWidth: 220, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={r.Objeto}>{(r.Objeto || "").slice(0, 60) || "—"}</td>
                              <td style={{ padding: "9px 8px", color: "var(--text2)" }}>{r["Área Requisitante"] || "—"}</td>
                              <td style={{ padding: "9px 8px", color: "var(--text2)" }}>{r.Comprador || r.respNCL || "—"}</td>
                              {hasPhaseCtx && phDelay && <>
                                <td style={{ padding: "9px 8px", textAlign: "center", fontWeight: 800, color: phDelay.atrasoFase ? "#c0392b" : "#1a5276" }}>{phDelay.diasNaFase}</td>
                                <td style={{ padding: "9px 8px", textAlign: "center", fontSize: 11, color: "#8e44ad" }}>{phDelay.slaFase != null ? `${phDelay.slaFase} d.u.` : "—"}</td>
                                <td style={{ padding: "9px 8px", textAlign: "center" }}>
                                  {phDelay.atrasoFase
                                    ? <span style={{ background: "#e74c3c", color: "#fff", borderRadius: 4, padding: "2px 8px", fontSize: 10, fontWeight: 700 }}>ATRASO</span>
                                    : <span style={{ background: "#27ae6022", color: "#27ae60", borderRadius: 4, padding: "2px 8px", fontSize: 10, fontWeight: 600 }}>Ok</span>}
                                </td>
                              </>}
                              {hasPhaseCtx && !phDelay && <>
                                <td style={{ padding: "9px 8px", textAlign: "center", color: "var(--text3)" }}>?</td>
                                <td style={{ padding: "9px 8px", textAlign: "center", color: "var(--text3)" }}>?</td>
                                <td style={{ padding: "9px 8px", textAlign: "center", color: "var(--text3)" }}>?</td>
                              </>}
                              <td style={{ padding: "9px 8px", textAlign: "center", fontWeight: 800, color: cor }}>{r.diasTotais || 0}</td>
                              <td style={{ padding: "9px 8px", textAlign: "center", fontSize: 11, color: "var(--text3)" }}>{r.prazoGeral} d.u.</td>
                              <td style={{ padding: "9px 8px", textAlign: "center" }}>
                                {late
                                  ? <span style={{ background: "#e74c3c", color: "#fff", borderRadius: 4, padding: "2px 8px", fontSize: 10, fontWeight: 700 }}>ATRASO</span>
                                  : <span style={{ background: "#27ae6022", color: "#27ae60", borderRadius: 4, padding: "2px 8px", fontSize: 10, fontWeight: 600 }}>No prazo</span>}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table></div>
                  </>;
                })()}
            </div>
          </div>
        </>}

        {aba === "upload" && <UploadScreen base={base} erro={erro} handleUpload={handleUpload} loading={loading} rawRows={rawRows} senhaErro={senhaErro} senhaInput={senhaInput} senhaOk={senhaOk} setAba={setAba} setSenhaErro={setSenhaErro} setSenhaInput={setSenhaInput} setSenhaOk={setSenhaOk} />}

        {/* ════════════ ABA: OVERVIEW ════════════ */}
        {aba === "overview" && (<OverviewScreen areaAnoData={areaAnoData} base={base} drillDown={drillDown} fBack={fBack} fBase={fBase} loginArea={loginArea} loginAreaDisp={loginAreaDisp} loginUser={loginUser} ocultarProblematicosOverview={ocultarProblematicosOverview} overviewAdmin={overviewAdmin} overviewInfoCard={overviewInfoCard} overviewInfoDefs={overviewInfoDefs} pgDrillDown={pgDrillDown} pgDrillDownSz={pgDrillDownSz} phaseIntervals={phaseIntervals} renderBuscaPanel={renderBuscaPanel} setAba={setAba} setDrillDown={setDrillDown} setOcultarProblematicosOverview={setOcultarProblematicosOverview} setOverviewInfoCard={setOverviewInfoCard} setPgDrillDown={setPgDrillDown} setPgDrillDownSz={setPgDrillDownSz} setSelProc={setSelProc} />
        )}

        {/* ════════════ ABA: QUALIDADE DE DADOS ════════════ */}
        {aba === "qualidade" && isAdminMaster && (<QualidadeScreen baseVis={baseVis} drillDown={drillDown} pgDrillDown={pgDrillDown} pgDrillDownSz={pgDrillDownSz} phaseIntervals={phaseIntervals} qualityData={qualityData} setAba={setAba} setDrillDown={setDrillDown} setPgDrillDown={setPgDrillDown} setPgDrillDownSz={setPgDrillDownSz} setSelProc={setSelProc} />)}

        {/* ════════════ ABA: EXECUTIVO ════════════ */}
        {aba === "executivo" && (<ExecutivoScreen alertCPL={alertCPL} alertNCL={alertNCL} base={base} baseVis={baseVis} drillDown={drillDown} execData={execData} executiveStrategic={executiveStrategic} fBack={fBack} fBase={fBase} loginArea={loginArea} loginAreaDisp={loginAreaDisp} loginUser={loginUser} meta={meta} notifs={notifs} pgDrillDown={pgDrillDown} pgDrillDownSz={pgDrillDownSz} phaseIntervals={phaseIntervals} rankCPLResp={rankCPLResp} rankGeralNCL={rankGeralNCL} rankNCL={rankNCL} setAba={setAba} setDrillDown={setDrillDown} setPgDrillDown={setPgDrillDown} setPgDrillDownSz={setPgDrillDownSz} setSelProc={setSelProc} showNotifPanel={showNotifPanel} />)}

        {/* ════════════ ABA: OPERACIONAL ════════════ */}
        {/* ════════════ ABA: CRONOGRAMA DE ENTREGAS ════════════ */}
        {aba === "cronograma" && <CronogramaScreen calDiaSel={calDiaSel} calMes={calMes} fArea={fArea} fBack={fBack} futurosTodosOpen={futurosTodosOpen} loginUser={loginUser} meta={meta} setCalDiaSel={setCalDiaSel} setCalMes={setCalMes} setFuturosTodosOpen={setFuturosTodosOpen} setSelProc={setSelProc} setVencidosOpen={setVencidosOpen} vencidosOpen={vencidosOpen} />}

        {/* ════════════ ABA: FRACIONAMENTO (MXM) ════════════ */}
        {aba === "fracionamento" && <FracionamentoScreen loginArea={loginArea} />}

        {/* ════════════ ABA: OPERACIONAL ════════════ */}
        {aba === "operacional" && (<OperacionalScreen base={base} distAnosFilter={distAnosFilter} fArea={fArea} fBack={fBack} fBase={fBase} loginArea={loginArea} loginAreaDisp={loginAreaDisp} meta={meta} phaseIntervals={phaseIntervals} setAba={setAba} setDistAnosFilter={setDistAnosFilter} setDrillDown={setDrillDown} setSelProc={setSelProc} setShowEmailDiario={setShowEmailDiario} setShowFunilInfo={setShowFunilInfo} setShowGaqSextaEditor={setShowGaqSextaEditor} showFunilInfo={showFunilInfo} />)}

        {/* ════════════ ABA: TODOS OS PROCESSOS ════════════ */}
        {aba === "processos" && <ProcessosScreen allPage={allPage} allProcesses={allProcesses} fBase={fBase} hideSDConcluido={hideSDConcluido} pgAll={pgAll} pgAllSz={pgAllSz} phaseIntervals={phaseIntervals} renderBuscaPanel={renderBuscaPanel} setHideSDConcluido={setHideSDConcluido} setPgAll={setPgAll} setPgAllSz={setPgAllSz} setSelProc={setSelProc} sortAll={sortAll} toggleSort={toggleSort} />}

        {/* ════════════ ABA: NOVA DEMANDA (NCL + CPL separados) ════════════ */}
        {aba === "nova" && <NovaScreen adicionarQuickTemp={adicionarQuickTemp} alertMapDem={alertMapDem} areas={areas} base={base} comprIgnorados={comprIgnorados} compsCPLPool={compsCPLPool} compsNCLPool={compsNCLPool} contarCarteira={contarCarteira} dem={dem} gerarRecs={gerarRecs} limparDemanda={limparDemanda} limparTempDemandas={limparTempDemandas} mods={mods} quickTemp={quickTemp} recsCPL={recsCPL} recsNCL={recsNCL} redistForm={redistForm} redistribuirCarteira={redistribuirCarteira} registrarTempDemanda={registrarTempDemanda} removerTempDemanda={removerTempDemanda} renderBuscaPanel={renderBuscaPanel} setComprIgnorados={setComprIgnorados} setDem={setDem} setQuickTemp={setQuickTemp} setRedistForm={setRedistForm} setVerRankingCompleto={setVerRankingCompleto} tempDemandas={tempDemandas} toggleIgnorado={toggleIgnorado} verRankingCompleto={verRankingCompleto} />}

        {/* ════════════ ABA: BALANCEAMENTO ════════════ */}
        {aba === "bal" && <BalScreen cargaCPLPreg={cargaCPLPreg} cargaCPLResp={cargaCPLResp} cargaNCL={cargaNCL} crossFilter={crossFilter} renderBuscaPanel={renderBuscaPanel} setAba={setAba} setSelComp={setSelComp} />}

        {/* ════════════ ABA: ACOMPANHAMENTO DE PROJETOS ════════════ */}
        {aba === "projetos" && <ProjetosScreen base={base} dirAreaSel={dirAreaSel} isDiretor={isDiretor} loginArea={loginArea} meta={meta} projetoTag={projetoTag} renderBuscaPanel={renderBuscaPanel} setProjetoTag={setProjetoTag} setSelProc={setSelProc} />}

        {/* ════════════ ABA: ALERTAS ════════════ */}
        {aba === "alertas" && <AlertasScreen alertCPL={alertCPL} alertNCL={alertNCL} fBack={fBack} onClickComp={onClickComp} onClickProc={onClickProc} onOpenTags={onOpenTags} pgAlert={pgAlert} pgAlertSz={pgAlertSz} pgCpl={pgCpl} pgCplSz={pgCplSz} pgHS={pgHS} pgHSSz={pgHSSz} pgStag={pgStag} pgStagSz={pgStagSz} phaseIntervals={phaseIntervals} renderBuscaPanel={renderBuscaPanel} setPgAlert={setPgAlert} setPgAlertSz={setPgAlertSz} setPgCpl={setPgCpl} setPgCplSz={setPgCplSz} setPgHS={setPgHS} setPgHSSz={setPgHSSz} setPgStag={setPgStag} setPgStagSz={setPgStagSz} setSelProc={setSelProc} setShowHS={setShowHS} setShowSLA={setShowSLA} setShowStag={setShowStag} showHS={showHS} showSLA={showSLA} showStag={showStag} tagVersion={tagVersion} toggleWatch={toggleWatch} watchVersion={watchVersion} />}

        {/* ════════════ ABA: ALERTA RC ════════════ */}
        {aba === "alertaRC" && <AlertaRCScreen alertRC={alertRC} fArea={fArea} fBase={fBase} meta={meta} onOpenTags={onOpenTags} pgAlertRC={pgAlertRC} pgAlertRCSz={pgAlertRCSz} renderBuscaPanel={renderBuscaPanel} setPgAlertRC={setPgAlertRC} setPgAlertRCSz={setPgAlertRCSz} setSelProc={setSelProc} toggleWatch={toggleWatch} />}

        {/* ════════════ ABA: GERADOR DE SENHA (Admin Master) ════════════ */}
        {aba === "geradorSenha" && isAdminMaster && <GeradorSenhaScreen genArea={genArea} genAreasDestaque={genAreasDestaque} genCopiado={genCopiado} genNome={genNome} genResultado={genResultado} genSenha={genSenha} genTipoPerfil={genTipoPerfil} setGenArea={setGenArea} setGenAreasDestaque={setGenAreasDestaque} setGenCopiado={setGenCopiado} setGenIsAdmin={setGenIsAdmin} setGenNome={setGenNome} setGenResultado={setGenResultado} setGenSenha={setGenSenha} setGenTipoPerfil={setGenTipoPerfil} />}

        {/* ════════════ ABA: RELATÓRIO CUSTOMIZADO (Admin Master) ════════════ */}
        {aba === "relCustom" && isAdminMaster && <RelCustomScreen areas={areas} baseEnriched={baseEnriched} meta={meta} phaseIntervals={phaseIntervals} rcuAreas={rcuAreas} rcuDias={rcuDias} rcuDiasModo={rcuDiasModo} rcuExcl={rcuExcl} rcuInds={rcuInds} rcuStatus={rcuStatus} setRcuAreas={setRcuAreas} setRcuDias={setRcuDias} setRcuDiasModo={setRcuDiasModo} setRcuExcl={setRcuExcl} setRcuInds={setRcuInds} setRcuStatus={setRcuStatus} setSelProc={setSelProc} />}

        {/* ════════════ ABA: ALERTA PRÉ-COMPRA ════════════ */}
        {aba === "alertaSD" && <AlertaSDScreen alertSD={alertSD} alertSDAberto={alertSDAberto} alertSDEncerrado={alertSDEncerrado} fArea={fArea} fBase={fBase} meta={meta} onOpenTags={onOpenTags} pgAlertSD={pgAlertSD} pgAlertSDEnc={pgAlertSDEnc} pgAlertSDEncSz={pgAlertSDEncSz} pgAlertSDSz={pgAlertSDSz} renderBuscaPanel={renderBuscaPanel} sdCopyStatus={sdCopyStatus} serviceDeskData={serviceDeskData} setPgAlertSD={setPgAlertSD} setPgAlertSDEnc={setPgAlertSDEnc} setPgAlertSDEncSz={setPgAlertSDEncSz} setPgAlertSDSz={setPgAlertSDSz} setSdCopyStatus={setSdCopyStatus} setSelProc={setSelProc} toggleWatch={toggleWatch} />}

        {/* ════════════ ABA: MAPEAMENTO (ENRIQUECIDO) ════════════ */}
        {aba === "mapeamento" && <MapeamentoScreen mapInsights={mapInsights} renderBuscaPanel={renderBuscaPanel} setAba={setAba} setFArea={setFArea} />}

      </div>

      {/* MODAL: Comprador/Responsável */}
      {/* ════════════ WATCHLIST NOTIFICATION PANEL ════════════ */}
      {showWatchPanel && (
        <div style={{ position: "fixed", top: 80, right: 20, width: 420, maxHeight: "70vh", zIndex: 850, background: "var(--card)", borderRadius: 14, boxShadow: "0 8px 40px rgba(0,0,0,0.25)", border: `2px solid ${watchChanges.length > 0 ? "#f39c12" : "#2e86c1"}`, display: "flex", flexDirection: "column", overflow: "hidden" }}>
          <div style={{ padding: "14px 18px", background: watchChanges.length > 0 ? "linear-gradient(135deg, #f39c12 0%, #e67e22 100%)" : "linear-gradient(135deg, #1a5276 0%, #2c3e50 100%)", color: "#fff", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div>
              <div style={{ fontWeight: 700, fontSize: 14 }}>👁 Watchlist {watchChanges.length > 0 ? "— Mudanças Detectadas" : ""}</div>
              <div style={{ fontSize: 10, opacity: .85 }}>
                {watchChanges.length > 0
                  ? `${watchChanges.length} processo(s) com alterações desde a última verificação`
                  : `${WatchlistManager.count()} processo(s) acompanhados`}
              </div>
            </div>
            <button onClick={() => setShowWatchPanel(false)} style={{ background: "none", border: "none", color: "#fff", fontSize: 20, cursor: "pointer" }}>×</button>
          </div>
          <div style={{ overflow: "auto", padding: "12px 16px", flex: 1 }}>
            {watchChanges.length > 0 ? watchChanges.map((ch, i) => (
              <div key={ch.pk} style={{ background: "var(--card2)", borderRadius: 8, padding: "10px 14px", marginBottom: 8, borderLeft: "4px solid #f39c12", cursor: "pointer" }}
                onClick={() => { setSelProc(ch.proc); setShowWatchPanel(false); }}>
                <div style={{ fontWeight: 700, fontSize: 12, color: "#2e86c1", marginBottom: 4 }}>
                  {ch.proc.NumRC || "—"} {ch.proc.TicketSD ? `· Pré-compra: ${ch.proc.TicketSD}` : ""}
                </div>
                <div style={{ fontSize: 11, color: "var(--text2)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", marginBottom: 6 }}>
                  {(ch.proc.Objeto || "").slice(0, 80) || "—"}
                </div>
                {ch.items.map((it, j) => (
                  <div key={j} style={{ fontSize: 11, color: "#d35400", padding: "2px 0", display: "flex", gap: 6, alignItems: "flex-start" }}>
                    <span>{it.icon}</span><span>{it.desc}</span>
                  </div>
                ))}
              </div>
            )) : (() => {
              const wList = WatchlistManager.getList();
              const wProcs = wList.map(pk => base.find(r => r.ProcessKey === pk)).filter(Boolean);
              return wProcs.length === 0
                ? <div style={{ textAlign: "center", padding: 30, color: "var(--text3)" }}>
                    <div style={{ fontSize: 32, marginBottom: 8 }}>👁</div>
                    <div style={{ fontSize: 13, fontWeight: 600 }}>Nenhum processo na watchlist</div>
                    <div style={{ fontSize: 11, marginTop: 4 }}>Abra a timeline de um processo e clique em "Acompanhar" ou use a ☆ nos cards de alerta.</div>
                  </div>
                : wProcs.map(r => (
                  <div key={r.ProcessKey} style={{ background: "var(--card2)", borderRadius: 8, padding: "8px 12px", marginBottom: 6, cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", borderLeft: `3px solid ${r.emA ? "#2e86c1" : r.isConcluded ? "#27ae60" : "#e67e22"}` }}
                    onClick={() => { setSelProc(r); setShowWatchPanel(false); }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 700, fontSize: 12, color: "#2e86c1" }}>{r.NumRC || "—"}</div>
                      <div style={{ fontSize: 10, color: "var(--text2)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{(r.Objeto || "").slice(0, 60)}</div>
                      <div style={{ fontSize: 10, color: "var(--text3)" }}>{r.status} · {r.diasTotais} d.u. · {r.respNCL || r.Pregoeiro || "—"}</div>
                    </div>
                    <button onClick={(e) => { e.stopPropagation(); toggleWatch(r); }} title="Remover da watchlist"
                      style={{ background: "none", border: "none", color: "#e74c3c", fontSize: 14, cursor: "pointer", flexShrink: 0 }}>✕</button>
                  </div>
                ));
            })()}
          </div>
          <div style={{ padding: "10px 16px", borderTop: "1px solid var(--border)", display: "flex", gap: 8, justifyContent: "space-between" }}>
            {watchChanges.length > 0 && <button onClick={() => { WatchlistManager.acknowledgeAll(base); setWatchChanges([]); }}
              style={{ background: "#27ae60", color: "#fff", border: "none", borderRadius: 7, padding: "7px 16px", cursor: "pointer", fontSize: 11, fontWeight: 700, flex: 1 }}>
              ✓ Marcar tudo como visto
            </button>}
            {WatchlistManager.count() > 0 && <button onClick={() => { setFWatch(true); setAba("processos"); setShowWatchPanel(false); }}
              style={{ background: "#1a5276", color: "#fff", border: "none", borderRadius: 7, padding: "7px 16px", cursor: "pointer", fontSize: 11, fontWeight: 600, flex: 1 }}>
              📋 Ver na tabela
            </button>}
            <button onClick={() => setShowWatchPanel(false)}
              style={{ background: "var(--card2)", color: "var(--text2)", border: "1px solid var(--border)", borderRadius: 7, padding: "7px 16px", cursor: "pointer", fontSize: 11, fontWeight: 600 }}>
              Fechar
            </button>
          </div>
        </div>
      )}

      {selComp && <ModalOverlay onDismiss={() => setSelComp(null)} style={{ position: "fixed", inset: 0, background: "#0007", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 900 }} onClick={() => setSelComp(null)}>
        <div style={{ background: "var(--card)", borderRadius: 12, width: "min(860px,95vw)", maxHeight: "85vh", display: "flex", flexDirection: "column", boxShadow: "0 8px 40px #0004" }} onClick={e => e.stopPropagation()}>
          <div style={{ padding: "16px 22px", background: selComp.tipo === "cpl" ? "#6c3483" : "#1a5276", color: "#fff", borderRadius: "12px 12px 0 0", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div><div style={{ fontWeight: 700, fontSize: 15 }}>{selComp.nome}</div>
              <div style={{ fontSize: 10, opacity: .8 }}>{procComp.length} processo(s) · {selComp.tipo === "cpl" ? "CPL" : "NCL"}</div></div>
            <button onClick={() => setSelComp(null)} style={{ background: "none", border: "none", color: "#fff", fontSize: 22, cursor: "pointer" }}>×</button>
          </div>
          <div style={{ overflow: "auto", padding: "14px 22px", flex: 1 }}>
            {procComp.length === 0 ? <div style={{ color: "var(--text3)", textAlign: "center", padding: 20 }}>Nenhum processo.</div>
              : procComp.map((r, i) => {
                const tags = TagsManager.getForProcess(r.ProcessKey, r.TicketSD);
                return (
                  <div key={i} style={{ borderBottom: "1px solid var(--border2)", padding: "9px 0", cursor: "pointer" }} onClick={() => setSelProc(r)}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 6 }}>
                      <div style={{ flex: 1 }}>
                        <div style={{ display: "flex", gap: 5, flexWrap: "wrap", alignItems: "center", marginBottom: 2 }}>
                          <span style={{ fontSize: 9, color: "var(--text3)", background: "var(--card2)", borderRadius: 3, padding: "0 4px" }}>{r.ProcessKey || "—"}</span>
                          <span style={{ fontWeight: 700, fontSize: 12, color: "#2e86c1" }}>{r.NumRC || "—"}</span>
                          {r.TicketSD && <span style={{ fontSize: 10, color: "#27ae60" }}>Pré-compra: {r.TicketSD}</span>}
                          <span style={{ fontSize: 10, color: "#2e86c1" }}>{r.Modalidade || "—"}</span>
                          {r.criticoNclGestao && <span style={{ fontSize: 10, color: "#c0392b", fontWeight: 700 }}>NCL crit.</span>}
                          {r.criticoCpl && <span style={{ fontSize: 10, color: "#8e44ad", fontWeight: 700 }}>CPL crit.</span>}
                          {tags.map(t => <TagBadge key={t} tag={t} small />)}
                        </div>
                        <div style={{ fontSize: 11, color: "var(--text)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 480 }}>{(r.Objeto || "").slice(0, 120) || "—"}</div>
                        <div style={{ fontSize: 10, color: "var(--text3)", marginTop: 2 }}>Status: <b>{r.status}</b>{r.statusDet ? <span style={{ fontSize: 10, background: "#f0ebfa33", borderRadius: 4, padding: "0 5px", color: "#8e44ad", fontWeight: 600, marginLeft: 5 }}>{r.statusDet}</span> : ""} | Abertura: {r.dataAbertura ? r.dataAbertura.toLocaleDateString("pt-BR") : "—"}{r.aberturaSD ? "" : " (RC)"}</div>
                      </div>
                      <div style={{ textAlign: "right", flexShrink: 0, minWidth: 75 }}>
                        {r.diasTotaisGestao > 0 && <div style={{ fontSize: 11, fontWeight: 700, color: r.criticoNclGestao ? "#c0392b" : "#e67e22" }}>{r.diasTotaisGestao} d.u. RC</div>}
                        {r.diasCpl > 0 && <div style={{ fontSize: 11, fontWeight: 700, color: r.criticoCpl ? "#8e44ad" : "#a569bd" }}>{r.diasCpl} d.u. CPL</div>}
                        <div style={{ fontSize: 9, color: "var(--text3)", marginTop: 3 }}>timeline →</div>
                      </div>
                    </div>
                  </div>
                );
              })}
          </div>
        </div>
      </ModalOverlay>}

        {/* ------------ ABA: GESTAO / PERFORMANCE ------------ */}
        {aba === "gestao" && !loginUser && <GestaoScreen alertCPL={alertCPL} alertNCL={alertNCL} alertScont={alertScont} base={base} cplByMod={cplByMod} fBack={fBack} gestaoPassErr={gestaoPassErr} gestaoPassInput={gestaoPassInput} gestaoUnlocked={gestaoUnlocked} nclByMod={nclByMod} onClickComp={onClickComp} phaseIntervals={phaseIntervals} rankCPLConc={rankCPLConc} rankCPLPerf={rankCPLPerf} rankGeralNCL={rankGeralNCL} rankNCL={rankNCL} rankNCLAgingSD={rankNCLAgingSD} rankNCLAgingSDAtivo={rankNCLAgingSDAtivo} rankNCLConc={rankNCLConc} rankScontConc={rankScontConc} rankScontPerf={rankScontPerf} scontBack={scontBack} scontByMod={scontByMod} scontPerfBack={scontPerfBack} setAba={setAba} setDrillDown={setDrillDown} setGestaoPassErr={setGestaoPassErr} setGestaoPassInput={setGestaoPassInput} setGestaoUnlocked={setGestaoUnlocked} setSelProc={setSelProc} />}

        {/* ════════════ ABA: LINHA DE METRO ════════════ */}
        {aba === "metro" && <MetroScreen base={base} fBase={fBase} metroAgrupado={metroAgrupado} metroBusca={metroBusca} metroComp={metroComp} pgMetro={pgMetro} pgMetroSz={pgMetroSz} phaseIntervals={phaseIntervals} setAba={setAba} setMetroAgrupado={setMetroAgrupado} setMetroBusca={setMetroBusca} setMetroComp={setMetroComp} setPgMetro={setPgMetro} setPgMetroSz={setPgMetroSz} setSelProc={setSelProc} />}

        {/* ════════════ ABA: PANTANAL ════════════ */}
        {aba === "pantanal" && <PantanalScreen base={base} pantanalBusca={pantanalBusca} pgPantanal={pgPantanal} pgPantanalSz={pgPantanalSz} setAba={setAba} setPantanalBusca={setPantanalBusca} setPantanalSel={setPantanalSel} setPgPantanal={setPgPantanal} setPgPantanalSz={setPgPantanalSz} />}

        {/* ════════════ ABA: PLANEJAMENTO DE COMPRA ════════════ */}
        {aba === "planejamento" && !loginUser && !loginArea && <PlanejamentoScreen base={base} fBack={fBack} recMesAberto={recMesAberto} setAba={setAba} setRecMesAberto={setRecMesAberto} setSelProc={setSelProc} />}

      {selProc && <TimelineModal proc={selProc} onClose={() => setSelProc(null)} onOpenTags={onOpenTags} onUpdate={refreshTags} phaseIntervals={phaseIntervals} onToggleWatch={toggleWatch} watchVersion={watchVersion} hidePredictions={isDiretor} serviceDeskData={serviceDeskData} />}
      {tagModal && <TagModal proc={tagModal} onClose={() => setTagModal(null)} onUpdate={refreshTags} />}
      {pantanalSel && <PantanalModal proc={pantanalSel} onClose={() => setPantanalSel(null)} />}
      {showEmailDiario && <EmailDiarioModal alertSDAberto={alertSDAberto} fBase={fBase} serviceDeskData={serviceDeskData} onClose={() => setShowEmailDiario(false)} />}
      {showGaqSextaEditor && <GaqSextaEditorModal processos={fBase} area={fArea || "Todas as Áreas"} meta={meta} fBase={fBase} onClose={() => setShowGaqSextaEditor(false)} />}

      {/* Floating controls for Presentation Mode */}
      {presMode && (
        <div style={{ position: "fixed", top: 16, right: 16, zIndex: 4000, display: "flex", flexDirection: "column", gap: 8, alignItems: "flex-end" }}>
          {/* Slideshow controls */}
          <div style={{ background: "rgba(15,25,45,0.92)", borderRadius: 12, padding: "10px 14px", boxShadow: "0 4px 24px rgba(0,0,0,.4)", display: "flex", alignItems: "center", gap: 10, backdropFilter: "blur(8px)" }}>
            {/* Tab pills */}
            <div style={{ display: "flex", gap: 4 }}>
              {SLIDESHOW_ABAS.map(a => (
                <button key={a} onClick={() => { setAba(a); if(slideshow){setSlideshowTick(60); clearInterval(slideshowRef.current); slideshowRef.current = setInterval(() => { setAba(prev => { const idx = SLIDESHOW_ABAS.indexOf(prev); return SLIDESHOW_ABAS[(idx+1)%SLIDESHOW_ABAS.length]; }); setSlideshowTick(60); }, 60000);} }}
                  style={{ background: aba === a ? "#2e86c1" : "rgba(255,255,255,0.12)", border: "none", borderRadius: 6, padding: "4px 9px", cursor: "pointer", fontSize: 10, fontWeight: 700, color: "#fff", transition: "background .2s" }}>
                  {SLIDESHOW_LABELS[a]}
                </button>
              ))}
            </div>
            {/* Slideshow toggle + countdown */}
            <div style={{ width: 1, height: 24, background: "rgba(255,255,255,0.2)" }} />
            <button onClick={() => setSlideshow(s => !s)}
              style={{ background: slideshow ? "#27ae60" : "rgba(255,255,255,0.15)", border: "none", borderRadius: 8, padding: "5px 12px", cursor: "pointer", fontSize: 11, fontWeight: 700, color: "#fff", display: "flex", alignItems: "center", gap: 5, transition: "background .2s" }}>
              {slideshow ? "⏸" : "▶"} {slideshow ? `${slideshowTick}s` : "Auto"}
            </button>
          </div>
          {/* Exit button */}
          <button onClick={() => setPresMode(false)}
            style={{ background: "#e74c3c", color: "#fff", border: "none", borderRadius: 10, padding: "8px 16px", cursor: "pointer", fontSize: 12, fontWeight: 700, boxShadow: "0 4px 16px rgba(0,0,0,.3)" }}>
            ✕ Sair
          </button>
        </div>
      )}

      {/* Spotlight Search (Ctrl+K) */}
      {spotlight && <div className="spotlight-overlay" onClick={() => setSpotlight(false)}>
        <div className="spotlight-box" onClick={e => e.stopPropagation()}>
          <input ref={spotlightRef} placeholder="Buscar RC, pré-compra ou objeto... (Esc para fechar)"
            onChange={e => handleBuscaChange(e.target.value)} autoFocus />
          <div className="spotlight-results">
            {buscaResults.length === 0 && busca.length >= 3 && <div style={{ padding: 20, textAlign: "center", color: "var(--text3)", fontSize: 13 }}>Nenhum resultado.</div>}
            {buscaResults.map((r, i) => (
              <div key={i} className="spotlight-item" onClick={() => { setSelProc(r); setSpotlight(false); }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <span style={{ fontSize: 9, color: "var(--text3)", marginRight: 5 }}>{r.ProcessKey || ""}</span>
                    <span style={{ fontWeight: 700, color: "#2e86c1", fontSize: 13, marginRight: 8 }}>{r.NumRC || "—"}</span>
                    {r.TicketSD && <span style={{ fontSize: 11, color: "#27ae60", marginRight: 8 }}>Pré-compra: {r.TicketSD}</span>}
                    <span style={{ fontSize: 11, color: "var(--text2)" }}>{r.respNCL || r.Pregoeiro || "—"}</span>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <span style={{ fontSize: 11, fontWeight: 700, color: r.diasTotais > 100 ? "#c0392b" : r.diasTotais > 50 ? "#e67e22" : "#27ae60" }}>{r.diasTotais} d.u.</span>
                  </div>
                </div>
                <div style={{ fontSize: 11, color: "var(--text3)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", marginTop: 2 }}>{(r.Objeto || "").slice(0, 100)}</div>
              </div>
            ))}
            {busca.length < 3 && <div style={{ padding: 20, textAlign: "center", color: "var(--text3)", fontSize: 12 }}>Digite pelo menos 3 caracteres...</div>}
          </div>
        </div>
      </div>}

      {/* Footer */}
      <div className="pres-hide no-print" style={{ borderTop: "1px solid var(--gaq-line)", padding: "14px 24px", marginTop: 30, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
        <div style={{ fontSize: 11, color: "var(--gaq-text-3)" }}>
          <b style={{ color: "var(--gaq-text-2)" }}>Sesc · GAQ</b> — Painel de Dados de Compras
        </div>
        <div style={{ fontSize: 10, color: "var(--gaq-text-3)", display: "flex", gap: 16 }}>
          <span>v6.0</span>
          {meta && <span>{meta.total} registros</span>}
          {meta && <span>Base gerada: {meta.atualizadoEm}</span>}
          {meta?.verificadoEm && <span>Verificada: {new Date(meta.verificadoEm).toLocaleString('pt-BR')}</span>}
        </div>
      </div>

      </div>
      {accessOpen && loginOk && !loginUser && !loginArea && !isDiretor && <AccessPanel onClose={() => setAccessOpen(false)} />}
    </div>
  );
}

ReactDOM.createRoot(document.getElementById("root")).render(<App />);
