const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('src/data-repository.js', 'utf8');
function setup() {
  const entries = {dados:['dados.js','__PAINEL_DADOS__',[{id:1}]], tags:['tags.js','__PAINEL_TAGS__',{}], servicedesk:['chamados_servicedesk.js','__SERVICE_DESK__',{'1':{}}], fracionamento:['base_fracionamento.js','__FRACIONAMENTO__',{itens:[]}]};
  const manifest = {schemaVersion:1, revision:'release-1',sources:Object.fromEntries(Object.entries(entries).map(([key,[file,,value]])=>[key,{file,revision:key+'-v1',count:key==='dados'?value.length:key==='fracionamento'?value.itens.length:Object.keys(value).length,generatedAt:null}]))};
  const reads = [], failures = new Set(), metadata = {}; let frames = 0;
  const context = {Date, URL, setTimeout, clearTimeout, console}; context.window = context;
  context.document = {baseURI:'file:///panel/index.html',body:{appendChild(){}}, createElement() {
    frames++;
    const win = {};
    return {contentWindow:win, remove(){frames--;}, contentDocument:{createElement(){return {};},head:{appendChild(script){
      const file = new URL(script.src).pathname.split('/').at(-1); reads.push(file);
      setTimeout(()=>{
        if(failures.has(file)) return script.onerror?.();
        if(file==='publicacao.js') win.__PAINEL_MANIFEST__=manifest;
        else for(const [key,[path,name,value]] of Object.entries(entries)) if(path===file){win[name]=value;win.__PAINEL_SOURCE_META__={[key]:{revision:metadata[key]||key+'-v1'}};}
        script.onload?.();
      },1);
    }}}};
  }};
  vm.runInNewContext(source,context);
  return {repo:context.PainelRepository, context, manifest, reads, failures, metadata, frames:()=>frames};
}
test('load all four sources, then check only manifest when unchanged',async()=>{
  const x=setup(); const snapshot=await x.repo.read();
  assert.equal(x.context.__PAINEL_DADOS__,undefined);
  x.repo.commit(snapshot); assert.equal(x.context.__PAINEL_DADOS__.length,1);
  assert.equal(x.reads.length,5); assert.equal(x.frames(),0);
  assert.equal((await x.repo.read()).changed,false); assert.equal(x.reads.length,6);
});
test('reject incomplete sync and preserve last committed globals; retry succeeds',async()=>{
  const x=setup(); x.repo.commit(await x.repo.read()); const before=x.context.__PAINEL_DADOS__;
  x.manifest.revision='release-2'; x.metadata.servicedesk='not-synced';
  await assert.rejects(x.repo.read(),/sincronizou/);
  assert.equal(x.context.__PAINEL_DADOS__,before); assert.equal(x.frames(),0);
  delete x.metadata.servicedesk; assert.equal((await x.repo.read()).changed,true);
});
test('file failure cleans frame and can be retried',async()=>{
  const x=setup();x.failures.add('dados.js');await assert.rejects(x.repo.read(),/indisponível/);assert.equal(x.frames(),0);
  x.failures.clear();assert.equal((await x.repo.read()).dados.length,1);
});
test('concurrent requests share a single read; uncommitted candidate is never cached',async()=>{
  const x=setup();const [a,b]=await Promise.all([x.repo.read(),x.repo.read()]);assert.equal(a,b);assert.equal(x.reads.length,5);
  await x.repo.read();assert.equal(x.reads.length,10);
});
test('invalid row count rejected and unknown generation time stays unknown',async()=>{
  const x=setup();x.manifest.sources.dados.count=2;await assert.rejects(x.repo.read(),/inválido/);
  x.manifest.sources.dados.count=1;const snap=await x.repo.read();assert.equal(x.repo.metadata(snap,1).atualizadoEm,'Geração não informada');
});
