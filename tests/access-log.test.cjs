const test=require('node:test');const assert=require('node:assert/strict');const fs=require('node:fs');const vm=require('node:vm');
test('local history excludes secrets, is bounded and handles unavailable storage',()=>{
  const data=new Map();const ctx={window:{},Date,localStorage:{getItem:k=>data.get(k),setItem:(k,v)=>data.set(k,v)}};
  vm.runInNewContext(fs.readFileSync('src/access-log.js','utf8'),ctx);
  for(let i=0;i<520;i++)ctx.window.PainelAccess.record('Perfil teste','Área','Entrada no painel');
  assert.equal(ctx.window.PainelAccess.list().length,500);
  assert.deepEqual(Object.keys(ctx.window.PainelAccess.list()[0]),['at','profile','type','event']);
  ctx.localStorage.setItem=()=>{throw Error('blocked');};assert.equal(ctx.window.PainelAccess.record('teste','Área','Entrada'),false);
});
