// prova-mobil.mjs - Prova la web del MOBIL (docs/): la pantalla d'inici, el
// planol des del Drive i el "Fer informe" que obre el formulari amb l'ID GIA.
//
//   cd suport/tests/navegador && npm install && node prova-mobil.mjs
//
// Serveix docs/ amb un servidor local (el formulari fa fetch de dades/*.json,
// que des de file:// no va) i canvia drive.js per un DOBLE: el Drive de debo
// demana el compte de Google. El doble torna el planol de prova
// (genera-planol-prova.ps1) i una activitat.
import { chromium } from 'playwright';
import { spawnSync } from 'node:child_process';
import http from 'node:http';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { AQUI, check, eq, seccio, serveixLeaflet, resultat } from './comu.mjs';

const DOCS = path.resolve(AQUI, '..', '..', '..', 'docs');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'mobil-nav-'));
const gen = spawnSync('pwsh', ['-NoProfile', '-File', path.join(AQUI, 'genera-planol-prova.ps1'), '-Dir', TMP], { stdio: 'inherit' });
if (gen.status !== 0) { console.error('No s\'ha pogut generar el planol de prova.'); process.exit(2); }
const PLANOL = fs.readFileSync(path.join(TMP, 'Planol.html'), 'utf8');

const TIPUS = { '.html': 'text/html; charset=utf-8', '.js': 'application/javascript', '.css': 'text/css', '.json': 'application/json',
                '.svg': 'image/svg+xml', '.png': 'image/png', '.webmanifest': 'application/manifest+json' };
const servidor = http.createServer((req, res) => {
  const ruta = decodeURIComponent(new URL(req.url, 'http://x').pathname);
  if (ruta === '/__planol.html') { res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' }); res.end(PLANOL); return; }
  const f = path.join(DOCS, ruta === '/' ? 'index.html' : ruta);
  if (!f.startsWith(DOCS) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { 'Content-Type': TIPUS[path.extname(f)] || 'application/octet-stream' });
  fs.createReadStream(f).pipe(res);
});
await new Promise((ok) => servidor.listen(0, '127.0.0.1', ok));
const BASE = 'http://127.0.0.1:' + servidor.address().port + '/';

// El DOBLE del Drive: connectat, el planol de prova i una activitat.
const DRIVE_DOBLE = `
window.Drive = {
  disponible: function () { return true; },
  connectar: function () { return Promise.resolve(); },
  reconnectarSilenci: function () { return Promise.resolve(true); },
  connectat: function () { return true; },
  llegirPlanol: function () { return fetch('/__planol.html').then(function (r) { return r.text(); }); },
  llegirActivitats: function () { return Promise.resolve({ Source: 'prova', ById: { '1447': { TITULAR: 'EL RACO SL', MOBIL: '600000000', EMAIL: 'raco@exemple.cat' } } }); },
  pujarPaquet: function () { return Promise.resolve({}); }
};`;

const nav = await chromium.launch({ headless: true });
const errors = [];
try {
  const ctx = await nav.newContext({ viewport: { width: 390, height: 800 }, isMobile: true, hasTouch: true });
  await serveixLeaflet(ctx);
  await ctx.route('**/drive.js', (r) => r.fulfill({ status: 200, contentType: 'application/javascript', body: DRIVE_DOBLE }));
  await ctx.route(/accounts\.google\.com|emailjs|googleapis/, (r) => r.abort());
  const p = await ctx.newPage();
  p.on('pageerror', (e) => errors.push(String(e)));

  seccio('La pantalla d\'inici');
  await p.goto(BASE + 'index.html');
  check(await p.isVisible('#inici'), 'surt la pantalla d\'inici');
  eq(await p.locator('.opcio-inici b').allTextContents(), ['1 · Consultar plànol', '2 · Generar informe'], 'amb les dues opcions');
  check(!(await p.isVisible('#navegacio')), 'i encara cap pas de l\'informe');
  await p.click('#btn-generar');
  await p.waitForSelector('#navegacio:not(.ocult)', { timeout: 5000 });
  check(!(await p.isVisible('#inici')), '«Generar informe»: comença l\'informe com sempre');

  seccio('Consultar el plànol (del Drive)');
  await p.goto(BASE + 'index.html');
  await p.click('#btn-planol');
  await p.waitForURL(/planol\.html/);
  await p.waitForSelector('#marc', { state: 'visible', timeout: 10000 });
  const marc = p.frameLocator('#marc');
  const frame = p.frames().find((f) => f !== p.mainFrame());
  await frame.waitForFunction(() => typeof capes !== 'undefined' && capes.length === PARCELES.length, null, { timeout: 15000 });
  check(true, 'el plànol del Drive es mostra dins de la pàgina');
  check(await marc.locator('#bt-side').isVisible(), 'a la pantalla del mòbil, el lateral va amb un botó «Filtres»');
  check(!(await marc.locator('#side').isVisible()), '...tancat');
  await marc.locator('#bt-side').click();
  check(await marc.locator('#side').isVisible(), '...que l\'obre');
  await marc.locator('#bt-side').click();

  seccio('Fer informe des d\'una activitat');
  await frame.evaluate(() => { const i = PARCELES.findIndex((x) => x.k === '2295827DF2729E'); obreFitxaActivitat('1447', PARCELES[i].c); });
  const boto = marc.locator('.leaflet-popup-content a.informe');
  eq(await boto.count(), 1, 'la fitxa de l\'activitat porta «Fer informe» (només dins de l\'app)');
  await boto.click();
  await p.waitForURL(/index\.html/, { timeout: 5000 });
  await p.waitForFunction(() => document.getElementById('in-gia') && document.getElementById('in-gia').value === '1447', null, { timeout: 5000 });
  check(await p.isVisible('#pas-capcalera'), 'obre l\'informe al Pas 2...');
  await p.waitForFunction(() => document.getElementById('tit-rao').textContent === 'EL RACO SL', null, { timeout: 5000 });
  check(true, '...amb l\'ID 1447 ja cercat (les dades del titular, posades)');
  check(!(await p.isChecked('#chk-titular')), 'i les has de confirmar tu abans de prémer Següent');
  eq(new URL(p.url()).search, '', 'i l\'adreça ja no porta el ?gia (tornar a carregar no torna a cercar)');

  seccio('Al PC (fora de l\'app) no hi ha «Fer informe»');
  const pc = await ctx.newPage();
  await pc.goto(BASE + '__planol.html');
  await pc.waitForFunction(() => typeof capes !== 'undefined' && capes.length === PARCELES.length, null, { timeout: 15000 });
  await pc.evaluate(() => { const i = PARCELES.findIndex((x) => x.k === '2295827DF2729E'); obreFitxaActivitat('1447', PARCELES[i].c); });
  eq(await pc.locator('.leaflet-popup-content a.informe').count(), 0, 'el plànol obert sol no ensenya el botó');

  seccio('Errors de JavaScript');
  eq(errors, [], 'cap error a les pàgines');
} catch (e) {
  check(false, 'la prova ha petat: ' + e);
} finally {
  await nav.close();
  servidor.close();
  fs.rmSync(TMP, { recursive: true, force: true });
}
process.exit(resultat('MOBIL'));
