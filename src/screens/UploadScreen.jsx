import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function UploadScreen({ base, erro, handleUpload, loading, rawRows, senhaErro, senhaInput, senhaOk, setAba, setSenhaErro, setSenhaInput, setSenhaOk }) {
  return (<div style={{ maxWidth: 540, margin: "40px auto" }}>
          <div style={{ background: "var(--card)", borderRadius: 14, padding: 36, boxShadow: "var(--shadow2)", border: "2px dashed #85c1e9", textAlign: "center" }}>
            <div style={{ fontSize: 44, marginBottom: 10 }}>📂</div>
            <div style={{ fontWeight: 700, fontSize: 17, marginBottom: 6, color: "#1a5276" }}>Carregar Base Consolidada</div>
            <div style={{ fontSize: 12, color: "var(--text2)", marginBottom: 20 }}>Aceita <b>.xlsx</b> ou <b>.csv</b> — ou coloque <b>dados.js</b> na mesma pasta para carga automática.</div>

            {!senhaOk
              ? <div style={{ background: "var(--card2)", borderRadius: 10, padding: "20px 24px", marginBottom: 16, textAlign: "left" }}>
                  <div style={{ fontSize: 13, fontWeight: 700, color: "#1a5276", marginBottom: 6 }}>Acesso restrito</div>
                  <div style={{ fontSize: 12, color: "var(--text2)", marginBottom: 12 }}>Digite a senha para liberar o upload.</div>
                  <input type="password" value={senhaInput} onChange={e => { setSenhaInput(e.target.value); setSenhaErro(false); }}
                    onKeyDown={async e => { if (e.key === "Enter") { const h = await gaqHashSenha(senhaInput); if (typeof UPLOAD_PASS !== "undefined" && h === UPLOAD_PASS) { setSenhaOk(true); setSenhaErro(false); } else { setSenhaErro(true); setSenhaInput(""); } } }}
                    placeholder="Senha de acesso..."
                    style={{ width: "100%", padding: "9px 12px", borderRadius: 7, border: `1.5px solid ${senhaErro ? "#e74c3c" : "var(--input-bd)"}`, fontSize: 13, background: "var(--input-bg)", color: "var(--text)", marginBottom: 8 }} />
                  {senhaErro && <div style={{ fontSize: 12, color: "#c0392b", marginBottom: 8 }}>Senha incorreta.</div>}
                  <button onClick={async () => { const h = await gaqHashSenha(senhaInput); if (typeof UPLOAD_PASS !== "undefined" && h === UPLOAD_PASS) { setSenhaOk(true); setSenhaErro(false); } else { setSenhaErro(true); setSenhaInput(""); } }}
                    style={{ background: "#1a5276", color: "#fff", border: "none", borderRadius: 8, padding: "9px 24px", cursor: "pointer", fontWeight: 700, fontSize: 13, width: "100%" }}>Confirmar</button>
                </div>
              : <>
                  <div style={{ background: "#eafaf1", borderRadius: 8, padding: "8px 14px", fontSize: 12, color: "#1e8449", fontWeight: 600, marginBottom: 16 }}>Acesso autorizado</div>
                  <div style={{ fontSize: 11, color: "var(--text3)", marginBottom: 20, background: "var(--card2)", borderRadius: 8, padding: "10px 14px", textAlign: "left", lineHeight: 1.6 }}>
                    <b>Como compartilhar (Opção C):</b><br/>
                    1. Faça upload da planilha abaixo<br/>
                    2. Clique "Exportar dados.js" e "Exportar tags.js"<br/>
                    3. Substitua os arquivos exportados na pasta compartilhada completa do painel.<br/>
                    4. Execute <b>geradores/publicar_dados.ps1</b> para validar e publicar as quatro bases.<br/>
                    5. Colegas abrem o index.html — dados carregam automaticamente
                  </div>
                  {loading
                    ? <div style={{ background: "#eaf4fb", borderRadius: 8, padding: 16, color: "#1a5276", fontWeight: 600 }}>Processando...</div>
                    : <label style={{ background: "#1a5276", color: "#fff", padding: "12px 30px", borderRadius: 8, cursor: "pointer", fontSize: 14, fontWeight: 600, display: "inline-block" }}>
                        Selecionar arquivo
                        <input type="file" accept=".xlsx,.csv" onChange={handleUpload} style={{ display: "none" }} />
                      </label>}
                  {erro && <div style={{ marginTop: 14, background: "#fdecea", borderRadius: 8, padding: 12, fontSize: 12, color: "#c0392b" }}>{erro}</div>}
                  {base.length > 0 && !loading && <div style={{ marginTop: 16 }}>
                    <div style={{ fontSize: 13, color: "#1e8449", fontWeight: 600, marginBottom: 10 }}>
                      {base.length} registros carregados · <span style={{ cursor: "pointer", textDecoration: "underline" }} onClick={() => setAba("overview")}>Ver painel →</span>
                    </div>
                    <div style={{ display: "flex", gap: 8, justifyContent: "center", flexWrap: "wrap" }}>
                      <button onClick={() => exportAsJS(rawRows)}
                        style={{ background: "#2e86c1", color: "#fff", border: "none", borderRadius: 7, padding: "8px 18px", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
                        💾 Exportar dados.js
                      </button>
                      <button onClick={() => exportTagsAsJS()}
                        style={{ background: "#34495e", color: "#fff", border: "none", borderRadius: 7, padding: "8px 18px", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
                        🏷️ Exportar tags.js
                      </button>
                    </div>
                  </div>}
                </>}
          </div>
        </div>);
}
