// Run against a disposable browser session already authenticated with test-only profiles.
import { spawnSync } from 'node:child_process';
import assert from 'node:assert/strict';
const exe = process.env.GAQ_BROWSER_BIN;
if (!exe) throw new Error('Set GAQ_BROWSER_BIN to the agent-browser executable.');
const session = process.env.GAQ_BROWSER_SESSION || 'gaq-review';
function cli(args, input) {
  const result = spawnSync(exe, ['--session', session, ...args], { input, encoding:'utf8', timeout:45000 });
  if (result.status !== 0) throw Error(result.stdout + result.stderr);
  return result.stdout.trim();
}
function evaluate(code) {
  let result = cli(['eval','--stdin'], code);
  try { result = JSON.parse(result); } catch {}
  return result;
}
const pages = evaluate('[...document.querySelectorAll(".gaq-side-item")].map(x=>x.textContent)');
const report = [];
for (const width of [1366, 390]) {
  cli(['set','viewport',String(width),'900']);
  for (let i=0;i<pages.length;i++) {
    if (width < 900) evaluate('document.querySelector("[aria-controls=gaq-navigation]").getAttribute("aria-expanded")==="false" && document.querySelector("[aria-controls=gaq-navigation]").click()');
    evaluate(`document.querySelectorAll('.gaq-side-item')[${i}].click()`);
    const result=evaluate(`({title:document.querySelector('h1')?.textContent,error:document.querySelector('#boot-status')?.textContent,overflow:document.documentElement.scrollWidth-innerWidth,sidebar:document.querySelector('[aria-controls=gaq-navigation]')?.getAttribute('aria-expanded')})`);
    assert.equal(result.error, '', `${pages[i]}: ${result.error}`);
    assert.ok(result.title, `Missing page after ${pages[i]}`);
    report.push({page:result.title,width,overflow:result.overflow});
  }
}
console.log(JSON.stringify(report,null,2));
