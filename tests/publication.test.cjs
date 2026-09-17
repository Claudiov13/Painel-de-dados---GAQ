const test=require('node:test');const assert=require('node:assert/strict');const fs=require('node:fs');const os=require('node:os');const path=require('node:path');const vm=require('node:vm');const {spawnSync}=require('node:child_process');
test('publisher validates before writing, preserves content, commits manifest and leaves no backups',()=>{
  const root=fs.mkdtempSync(path.join(os.tmpdir(),'gaq-publication-test-'));
  const globals={'dados.js':['__PAINEL_DADOS__',[{id:1,Objeto:'Teste; ação'}]],'tags.js':['__PAINEL_TAGS__',{tags:{},tagList:[]}],'chamados_servicedesk.js':['__SERVICE_DESK__',{'1':{interacoes:[]}}],'base_fracionamento.js':['__FRACIONAMENTO__',{itens:[]}]};
  try {
    for(const [file,[name,data]] of Object.entries(globals)) fs.writeFileSync(path.join(root,file),'window.'+name+' = '+JSON.stringify(data)+';\n');
    const run=()=>spawnSync('powershell.exe',['-NoProfile','-ExecutionPolicy','Bypass','-File',path.resolve('geradores/publicar_dados.ps1'),'-Root',root],{encoding:'utf8'});
    let result=run(); assert.equal(result.status,0,result.stdout+result.stderr);
    const context={window:{}};
    for(const [file,[name,data]] of Object.entries(globals)){vm.runInNewContext(fs.readFileSync(path.join(root,file),'utf8'),context);assert.equal(JSON.stringify(context.window[name]),JSON.stringify(data));}
    vm.runInNewContext(fs.readFileSync(path.join(root,'publicacao.js'),'utf8'),context); assert.equal(context.window.__PAINEL_MANIFEST__.sources.dados.count,1);
    const original=fs.readFileSync(path.join(root,'publicacao.js'),'utf8');
    fs.writeFileSync(path.join(root,'chamados_servicedesk.js'),'invalid json');
    const before=fs.readFileSync(path.join(root,'dados.js'),'utf8');result=run();assert.notEqual(result.status,0);
    assert.equal(fs.readFileSync(path.join(root,'publicacao.js'),'utf8'),original);assert.equal(fs.readFileSync(path.join(root,'dados.js'),'utf8'),before);
    assert.equal(fs.readdirSync(root).length,5);
  } finally {
    // Only files explicitly created in this test's unique directory.
    for(const file of [...Object.keys(globals),'publicacao.js']) {const p=path.join(root,file);if(fs.existsSync(p))fs.unlinkSync(p);}
    fs.rmdirSync(root);
  }
});
