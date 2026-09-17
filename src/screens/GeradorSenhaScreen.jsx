import { ModalOverlay } from '../ModalOverlay.jsx';
const { useState, useMemo, useRef, useEffect, useCallback } = React;
import { AccessPanel } from '../AccessPanel.jsx';
const XLSX_LIB = window.XLSX;
const Papa = window.Papa;
const { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } = Recharts;
import { Icon, Spark, MiniCard, GBox, WCard, KpiBox, BadgeScore, Sel, RankRow, TagBadge, Pagination, DrillDownPanel, CopyChip, HistoricoModal, TagModal, EmailDiarioModal, GAQ_SEXTA_TD, GAQ_SEXTA_INPUT, GAQ_SEXTA_EDIT_VAZIO, GaqSextaLinhaEditor, GaqSextaEditorModal, PantanalModal, ServiceDeskAlertStrip, ServiceDeskHistoryModal, TimelineModal, ProcCard, MiniTimeline, PatternCard, RecCard } from '../components.jsx';
import { getServiceDeskRecord, fixServiceDeskText, cleanServiceDeskObject, cleanServiceDeskParty, formatServiceDeskType, formatServiceDeskSubject, summarizeServiceDesk, getServiceDeskAlertSummary, isServiceDeskConcluido, enrichWithSDHistory, enrichSD, gerarScriptExtratorOTRS } from '../domain/service-desk.js';

export function GeradorSenhaScreen({ genArea, genAreasDestaque, genCopiado, genNome, genResultado, genSenha, genTipoPerfil, setGenArea, setGenAreasDestaque, setGenCopiado, setGenIsAdmin, setGenNome, setGenResultado, setGenSenha, setGenTipoPerfil }) {
  return (<div className="anim-fade">
          <div style={{ background: "linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #334155 100%)", color: "#fff", borderRadius: "var(--gaq-r-xl)", padding: "26px 30px", marginBottom: 18, boxShadow: "var(--gaq-shadow-2)" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
              <div>
                <div style={{ fontSize: 11, fontWeight: 600, opacity: .72, letterSpacing: ".05em", textTransform: "uppercase", marginBottom: 4 }}>Admin Master · Cadastro de Senhas</div>
                <div style={{ fontSize: 22, fontWeight: 700, letterSpacing: "-.02em" }}>Gerador de Senha</div>
                <div style={{ fontSize: 12, opacity: .82, marginTop: 6, maxWidth: 720, lineHeight: 1.5 }}>
                  Use esta tela para criar a linha pronta de uma nova senha, ou substituir uma existente. A senha em si nunca é guardada — só o hash. Após gerar, copie a linha e cole no <code style={{ background: "rgba(255,255,255,0.12)", padding: "1px 6px", borderRadius: 4 }}>auth.js</code> dentro do bloco indicado.
                </div>
              </div>
            </div>
          </div>

          <div className="gaq-responsive-grid" style={{ display: "grid", gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1fr)", gap: 18 }}>
            {/* Coluna esquerda — formulário */}
            <div className="gaq-card" style={{ padding: 22 }}>
              <div className="gaq-h3" style={{ marginBottom: 4 }}>Dados da nova senha</div>
              <div className="sub" style={{ marginBottom: 16 }}>Preencha conforme o tipo de perfil escolhido.</div>

              <label style={{ display: "block", fontSize: 11, fontWeight: 700, color: "var(--gaq-text-3)", textTransform: "uppercase", letterSpacing: ".04em", marginBottom: 6 }}>Tipo de perfil</label>
              <select value={genTipoPerfil} onChange={e => { setGenTipoPerfil(e.target.value); setGenResultado(null); setGenIsAdmin(false); setGenNome(""); setGenArea(""); setGenAreasDestaque(""); }}
                style={{ width: "100%", padding: "10px 12px", borderRadius: 8, border: "1px solid var(--gaq-line)", background: "var(--input-bg)", color: "var(--gaq-text)", fontSize: 13, marginBottom: 16, boxSizing: "border-box" }}>
                <option value="perfil">Comprador / Avaliador (perfil individual)</option>
                <option value="admin">Admin (visão completa)</option>
                <option value="master">Admin Master</option>
                <option value="area">Área (filtra por área)</option>
                <option value="diretor">Diretor</option>
                <option value="gestao">Painel de Gestão</option>
                <option value="upload">Upload de base</option>
              </select>

              <label style={{ display: "block", fontSize: 11, fontWeight: 700, color: "var(--gaq-text-3)", textTransform: "uppercase", letterSpacing: ".04em", marginBottom: 6 }}>Senha que o usuário vai digitar</label>
              <input type="text" value={genSenha} onChange={e => { setGenSenha(e.target.value); setGenResultado(null); }}
                placeholder="Ex: Joao#Compr2026!"
                style={{ width: "100%", padding: "10px 12px", borderRadius: 8, border: "1px solid var(--gaq-line)", background: "var(--input-bg)", color: "var(--gaq-text)", fontSize: 13, fontFamily: "var(--gaq-mono)", marginBottom: 16, boxSizing: "border-box" }} />

              {genTipoPerfil === "perfil" && <>
                <label style={{ display: "block", fontSize: 11, fontWeight: 700, color: "var(--gaq-text-3)", textTransform: "uppercase", letterSpacing: ".04em", marginBottom: 6 }}>Nome exato do responsável (filtra processos por este nome)</label>
                <input type="text" value={genNome} onChange={e => { setGenNome(e.target.value); setGenResultado(null); }}
                  placeholder="Ex: João Silva"
                  style={{ width: "100%", padding: "10px 12px", borderRadius: 8, border: "1px solid var(--gaq-line)", background: "var(--input-bg)", color: "var(--gaq-text)", fontSize: 13, marginBottom: 16, boxSizing: "border-box" }} />
              </>}

              {genTipoPerfil === "master" && <>
                <label style={{ display: "block", fontSize: 11, fontWeight: 700, color: "var(--gaq-text-3)", textTransform: "uppercase", letterSpacing: ".04em", marginBottom: 6 }}>Nome de exibição do Admin Master</label>
                <input type="text" value={genNome} onChange={e => { setGenNome(e.target.value); setGenResultado(null); }}
                  placeholder="Ex: Maria Silva"
                  style={{ width: "100%", padding: "10px 12px", borderRadius: 8, border: "1px solid var(--gaq-line)", background: "var(--input-bg)", color: "var(--gaq-text)", fontSize: 13, marginBottom: 16, boxSizing: "border-box" }} />
              </>}

              {genTipoPerfil === "area" && <>
                <label style={{ display: "block", fontSize: 11, fontWeight: 700, color: "var(--gaq-text-3)", textTransform: "uppercase", letterSpacing: ".04em", marginBottom: 6 }}>Código da área (igual ao da coluna ÁREA no dados.js) — separe por vírgula para dar acesso a mais de uma área</label>
                <input type="text" value={genArea} onChange={e => { setGenArea(e.target.value.toUpperCase()); setGenResultado(null); }}
                  placeholder="Ex: GTI  ou  GIN,ALMOXARIFADO"
                  style={{ width: "100%", padding: "10px 12px", borderRadius: 8, border: "1px solid var(--gaq-line)", background: "var(--input-bg)", color: "var(--gaq-text)", fontSize: 13, fontFamily: "var(--gaq-mono)", marginBottom: 16, boxSizing: "border-box" }} />
              </>}

              {genTipoPerfil === "diretor" && <>
                <label style={{ display: "block", fontSize: 11, fontWeight: 700, color: "var(--gaq-text-3)", textTransform: "uppercase", letterSpacing: ".04em", marginBottom: 6 }}>Nome do diretor</label>
                <input type="text" value={genNome} onChange={e => { setGenNome(e.target.value); setGenResultado(null); }}
                  placeholder="Ex: Carlos Souza"
                  style={{ width: "100%", padding: "10px 12px", borderRadius: 8, border: "1px solid var(--gaq-line)", background: "var(--input-bg)", color: "var(--gaq-text)", fontSize: 13, marginBottom: 12, boxSizing: "border-box" }} />
                <label style={{ display: "block", fontSize: 11, fontWeight: 700, color: "var(--gaq-text-3)", textTransform: "uppercase", letterSpacing: ".04em", marginBottom: 6 }}>Áreas de destaque (separadas por vírgula)</label>
                <input type="text" value={genAreasDestaque} onChange={e => { setGenAreasDestaque(e.target.value.toUpperCase()); setGenResultado(null); }}
                  placeholder="Ex: ASCOM, TI, GIN"
                  style={{ width: "100%", padding: "10px 12px", borderRadius: 8, border: "1px solid var(--gaq-line)", background: "var(--input-bg)", color: "var(--gaq-text)", fontSize: 13, fontFamily: "var(--gaq-mono)", marginBottom: 16, boxSizing: "border-box" }} />
              </>}

              {(genTipoPerfil === "admin" || genTipoPerfil === "gestao" || genTipoPerfil === "upload") && (
                <div style={{ background: "var(--gaq-bg-2)", borderRadius: 8, padding: "10px 12px", marginBottom: 16, fontSize: 12, color: "var(--gaq-text-3)", border: "1px solid var(--gaq-line)" }}>
                  {genTipoPerfil === "admin" && <>Administrador comum: visão completa do painel. Não precisa de campos extras.</>}
                  {genTipoPerfil === "gestao" && <>Senha única que destrava o painel de Gestão. Esta operação <b>substitui</b> a senha atual de gestão.</>}
                  {genTipoPerfil === "upload" && <>Senha única que libera o upload de base. Esta operação <b>substitui</b> a senha atual de upload.</>}
                </div>
              )}

              <button onClick={async () => {
                if (!genSenha) { alert("Digite a senha que o usuário vai usar."); return; }
                const hash = await gaqHashSenha(genSenha);
                let linha = "", varName = "", instrucoes = "";
                if (genTipoPerfil === "perfil") {
                  if (!genNome.trim()) { alert("Digite o nome do responsável."); return; }
                  linha = `  "${hash}": "${genNome.trim()}",`;
                  varName = "PERFIS_ACESSO";
                  instrucoes = "Adicione esta linha dentro do bloco PERFIS_ACESSO no auth.js.";
                } else if (genTipoPerfil === "admin") {
                  linha = `  "${hash}": null,`;
                  varName = "PERFIS_ACESSO";
                  instrucoes = "Adicione esta linha dentro do bloco PERFIS_ACESSO no auth.js.";
                } else if (genTipoPerfil === "master") {
                  if (!genNome.trim()) { alert("Digite o nome de exibição."); return; }
                  linha = `  "${hash}": { nome: "${genNome.trim()}" },`;
                  varName = "ADMIN_MASTER_PERFIS";
                  instrucoes = "Adicione esta linha dentro do bloco ADMIN_MASTER_PERFIS no auth.js.";
                } else if (genTipoPerfil === "area") {
                  if (!genArea.trim()) { alert("Digite o código da área."); return; }
                  linha = `  "${hash}": "${genArea.trim()}",`;
                  varName = "AREA_PERFIS";
                  instrucoes = "Adicione esta linha dentro do bloco AREA_PERFIS no auth.js.";
                } else if (genTipoPerfil === "diretor") {
                  if (!genNome.trim()) { alert("Digite o nome do diretor."); return; }
                  const areas = genAreasDestaque.split(",").map(a => a.trim()).filter(Boolean).map(a => `"${a}"`).join(", ");
                  linha = `  "${hash}": { nome: "${genNome.trim()}", areasDestaque: [${areas}] },`;
                  varName = "DIRETOR_PERFIS";
                  instrucoes = "Adicione esta linha dentro do bloco DIRETOR_PERFIS no auth.js.";
                } else if (genTipoPerfil === "gestao") {
                  linha = `var GESTAO_PASS = "${hash}";`;
                  varName = "GESTAO_PASS";
                  instrucoes = "Substitua a linha atual de var GESTAO_PASS = ... no auth.js por esta.";
                } else if (genTipoPerfil === "upload") {
                  linha = `var UPLOAD_PASS = "${hash}";`;
                  varName = "UPLOAD_PASS";
                  instrucoes = "Substitua a linha atual de var UPLOAD_PASS = ... no auth.js por esta.";
                }
                setGenResultado({ linha, varName, instrucoes, hash });
                setGenCopiado(false);
              }}
                style={{ background: "var(--gaq-blue)", color: "#fff", border: "none", borderRadius: 8, padding: "11px 28px", fontWeight: 700, fontSize: 13, cursor: "pointer", width: "100%" }}>
                Gerar linha para o auth.js
              </button>
            </div>

            {/* Coluna direita — resultado */}
            <div className="gaq-card" style={{ padding: 22 }}>
              <div className="gaq-h3" style={{ marginBottom: 4 }}>Linha gerada</div>
              <div className="sub" style={{ marginBottom: 16 }}>Copie a linha abaixo e cole no <code style={{ background: "var(--gaq-bg-2)", padding: "1px 6px", borderRadius: 4, fontSize: 11 }}>auth.js</code>.</div>

              {!genResultado ? (
                <div style={{ background: "var(--gaq-bg-2)", borderRadius: 8, padding: "30px 16px", textAlign: "center", fontSize: 13, color: "var(--gaq-text-3)", border: "1px dashed var(--gaq-line)" }}>
                  Preencha os campos ao lado e clique em <b>Gerar</b> para ver a linha aqui.
                </div>
              ) : (
                <>
                  <div style={{ background: "linear-gradient(180deg, var(--gaq-bg-2) 0%, var(--gaq-surface) 100%)", borderLeft: "3px solid var(--gaq-blue)", borderRadius: 8, padding: "12px 14px", marginBottom: 12, fontSize: 12, color: "var(--gaq-text-2)", lineHeight: 1.5 }}>
                    <b style={{ color: "var(--gaq-blue)" }}>Onde colar:</b> {genResultado.instrucoes}
                  </div>
                  <pre style={{ background: "var(--gaq-bg-2)", borderRadius: 8, padding: "12px 14px", fontSize: 12.5, fontFamily: "var(--gaq-mono)", color: "var(--gaq-text)", border: "1px solid var(--gaq-line)", overflow: "auto", marginBottom: 12, whiteSpace: "pre-wrap", wordBreak: "break-all" }}>{genResultado.linha}</pre>
                  <div style={{ display: "flex", gap: 8 }}>
                    <button onClick={async () => {
                      try { await navigator.clipboard.writeText(genResultado.linha); setGenCopiado(true); setTimeout(() => setGenCopiado(false), 2000); } catch { alert("Não foi possível copiar automaticamente. Selecione o texto e copie manualmente."); }
                    }}
                      style={{ background: genCopiado ? "var(--gaq-green)" : "var(--gaq-blue)", color: "#fff", border: "none", borderRadius: 8, padding: "9px 22px", fontWeight: 700, fontSize: 12.5, cursor: "pointer", flex: 1 }}>
                      {genCopiado ? "Copiado!" : "Copiar linha"}
                    </button>
                    <button onClick={() => { setGenResultado(null); setGenSenha(""); setGenNome(""); setGenArea(""); setGenAreasDestaque(""); }}
                      style={{ background: "var(--gaq-surface)", color: "var(--gaq-text)", border: "1px solid var(--gaq-line)", borderRadius: 8, padding: "9px 18px", fontWeight: 600, fontSize: 12.5, cursor: "pointer" }}>
                      Limpar
                    </button>
                  </div>
                  <div style={{ marginTop: 14, fontSize: 11, color: "var(--gaq-text-3)", lineHeight: 1.55, background: "var(--gaq-bg-2)", borderRadius: 8, padding: "10px 12px", border: "1px solid var(--gaq-line)" }}>
                    <b>Próximos passos:</b><br/>
                    1. Abra o arquivo <code style={{ fontFamily: "var(--gaq-mono)" }}>auth.js</code> no Bloco de Notas.<br/>
                    2. Localize o bloco <code style={{ fontFamily: "var(--gaq-mono)" }}>{genResultado.varName}</code>.<br/>
                    3. Cole a linha (ou substitua a existente, se for senha única).<br/>
                    4. Salve o arquivo. A senha já fica ativa no próximo login.
                  </div>
                </>
              )}
            </div>
          </div>

          <div className="gaq-card" style={{ padding: 18, marginTop: 18 }}>
            <div className="gaq-h3" style={{ marginBottom: 6 }}>Como funciona</div>
            <div style={{ fontSize: 12.5, color: "var(--gaq-text-2)", lineHeight: 1.6 }}>
              A senha que você digita acima é convertida em um <b>hash SHA-256</b> com salt ({GAQ_SALT}). Apenas o hash é gravado no <code style={{ fontFamily: "var(--gaq-mono)" }}>auth.js</code>. Quando o usuário digita a senha original, o painel calcula o hash dela e compara — se baterem, libera o acesso. Não há como recuperar a senha original a partir do hash, então guarde a senha original em local seguro antes de fechar esta tela.
            </div>
          </div>
        </div>);
}
