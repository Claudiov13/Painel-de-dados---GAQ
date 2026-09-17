import { spawnSync } from 'node:child_process';
import assert from 'node:assert/strict';
const exe=process.env.GAQ_BROWSER_BIN;
if(!exe) throw Error('Set GAQ_BROWSER_BIN');
function cli(args,input){const r=spawnSync(exe,['--session','gaq-review',...args],{input,encoding:'utf8',timeout:45000});if(r.status!==0)throw Error(r.stdout+r.stderr);return r.stdout.trim();}
function evaluate(code){const output=cli(['eval','--stdin'],code);try{return JSON.parse(output);}catch{return output;}}
function ref(kind,label){const snapshot=cli(['snapshot','-i']);const line=snapshot.split('\n').find(x=>x.includes(`${kind} "${label}"`));const found=line?.match(/ref=(e\d+)/);if(!found)throw Error('Missing '+label);return '@'+found[1];}
const profiles=[
 {name:'Área',tab:'Área',button:'Acessar Área',inject:'AREA_PERFIS[h]="ALMOXARIFADO"'},
 {name:'Comprador',tab:'GAQ',button:'Entrar',inject:'PERFIS_ACESSO[h]=Object.values(PERFIS_ACESSO).find(x=>typeof x==="string"&&x)'},
 {name:'Diretoria',tab:'Diretoria',button:'Acessar Diretoria',inject:'DIRETOR_PERFIS[h]={...Object.values(DIRETOR_PERFIS)[0],nome:"Teste de verificação"}'},
 {name:'Admin Master',tab:'GAQ',button:'Entrar',inject:'ADMIN_MASTER_PERFIS[h]={...Object.values(ADMIN_MASTER_PERFIS)[0],nome:"Teste de verificação"}'}
];
for(const profile of profiles){
 evaluate('sessionStorage.clear();localStorage.removeItem("painel_diretor_remember")');cli(['reload']);
 evaluate(`(async()=>{await window.__PAINEL_SYNC__();const h=await gaqHashSenha('GAQ-perfil-teste-descartavel');${profile.inject};return true})()`);
 cli(['click',ref('button',profile.tab)]);
 const snapshot=cli(['snapshot','-i']);const input=snapshot.split('\n').find(x=>x.includes('textbox ')).match(/ref=(e\d+)/)[1];
 cli(['fill','@'+input,'GAQ-perfil-teste-descartavel']);cli(['click',ref('button',profile.button)]);
 const state=evaluate('({logged:sessionStorage.getItem("painel_login_ok"),error:document.getElementById("boot-status").textContent,adminButton:[...document.querySelectorAll(".gaq-header-actions button")].some(b=>b.textContent==="Acessos locais"),title:document.querySelector("h1")?.textContent})');
 assert.equal(state.logged,'1');assert.equal(state.error,'');assert.equal(state.adminButton,profile.name==='Admin Master');
 const pages=evaluate('[...document.querySelectorAll(".gaq-side-item")].map(x=>x.textContent)');
 for(let i=0;i<pages.length;i++){
   evaluate(`document.querySelectorAll('.gaq-side-item')[${i}].click()`);
   const error=evaluate('document.getElementById("boot-status").textContent');assert.equal(error,'',profile.name+' / '+pages[i]+': '+error);
 }
 console.log(profile.name+': login and '+pages.length+' menu pages OK');
}
