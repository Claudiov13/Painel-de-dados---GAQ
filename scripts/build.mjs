import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
const require = createRequire(import.meta.url);
const moduleRoot = process.env.GAQ_BUILD_MODULES || path.join(process.env.LOCALAPPDATA || '.', 'PainelGAQ', 'build', 'node_modules');
const { transform, build } = require(path.join(moduleRoot, 'esbuild'));
fs.mkdirSync('assets', { recursive: true });
const libraries = [
  ['react.production.min.js', 'react@18.2.0/umd/react.production.min.js'],
  ['react-dom.production.min.js', 'react-dom@18.2.0/umd/react-dom.production.min.js'],
  ['xlsx.full.min.js', 'xlsx@0.18.5/dist/xlsx.full.min.js'],
  ['papaparse.min.js', 'papaparse@5.4.1/papaparse.min.js'],
  ['react-is.production.min.js', 'react-is@18.2.0/umd/react-is.production.min.js'],
  ['prop-types.min.js', 'prop-types@15.8.1/prop-types.min.js'],
  ['Recharts.min.js', 'recharts@2.12.7/umd/Recharts.min.js']
];
// Pinned original library versions: no framework upgrade during this migration.
const vendorParts = await Promise.all(libraries.map(async ([name, remotePath]) => {
  const local = path.join(moduleRoot, '.gaq-vendor', name);
  if (fs.existsSync(local)) return fs.readFileSync(local, 'utf8');
  const response = await fetch('https://cdn.jsdelivr.net/npm/' + remotePath);
  if (!response.ok) throw new Error('Library download failed: ' + remotePath);
  const source = await response.text();
  fs.mkdirSync(path.dirname(local), { recursive: true });
  fs.writeFileSync(local, source);
  return source;
}));
const vendors = vendorParts.join('\n;\n');
const app = await build({ entryPoints: ['src/App.jsx'], bundle: true, write: false, format: 'iife', target: ['chrome100', 'edge100'], minify: true, charset: 'utf8', legalComments: 'inline', define: { 'process.env.NODE_ENV': '"production"' } });
// Domain globals stay compatible with existing reports and auth.js.
const runtime = (await transform(['config.js','core.js','reports.js','src/data-repository.js','src/access-log.js'].map(path => fs.readFileSync(path,'utf8')).join('\n;\n'), { minify: true, target: 'es2020', charset: 'utf8', legalComments: 'inline' })).code;
const files = { 'assets/vendor.js': vendors, 'assets/runtime.js': runtime, 'assets/app.js': app.outputFiles[0].text };
const revision = createHash('sha256').update(Object.values(files).join('') + fs.readFileSync('styles.css', 'utf8') + fs.readFileSync('auth.js', 'utf8')).digest('hex').slice(0,12);
for (const [file, contents] of Object.entries(files)) { fs.writeFileSync(file + '.tmp', contents); fs.renameSync(file + '.tmp', file); }
const template = fs.readFileSync('src/index.template.html', 'utf8');
const html = template.replaceAll('{{REVISION}}', revision).replace('{{OTRS_TEMPLATES}}', fs.readFileSync('src/otrs-templates.html', 'utf8'));
fs.writeFileSync('index.html.tmp', html); fs.renameSync('index.html.tmp', 'index.html');
console.log(`Build ${revision}: app ${Math.round(files['assets/app.js'].length/1024)} KB; browser compilation removed; all libraries local.`);
