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
  pujarPaquet: function () { return Promise.resolve({}); },
  // La CCO del Drive privat (correu-mobil.json); window.__ccoFalla la fa petar.
  llegirCorreuMobil: function () {
    if (window.__ccoFalla) { return Promise.reject(new Error('sense fitxer')); }
    return Promise.resolve({ bcc: [{ addr: 'sempre@x.cat', def: true }, { addr: 'tria1@x.cat', def: false }, { addr: 'tria2@x.cat', def: false }] });
  }
};`;

const nav = await chromium.launch({ headless: true });
const errors = [];
try {
  // Una posicio de prova a Cornella (el permis d'ubicacio, ja donat).
  const GEO = { latitude: 41.3560, longitude: 2.0740, accuracy: 15 };
  const ctx = await nav.newContext({ viewport: { width: 390, height: 800 }, isMobile: true, hasTouch: true, geolocation: GEO, permissions: ['geolocation'] });
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

  seccio('Tornar a l\'inici des de «Generar informe»');
  check(await p.isVisible('#btn-tornar-inici'), 'hi ha la fletxa a dalt a l\'esquerra (com al plànol)');
  await p.click('#btn-tornar-inici');
  await p.waitForSelector('#inici', { state: 'visible', timeout: 5000 });
  check(!(await p.isVisible('#navegacio')) && !(await p.isVisible('#btn-tornar-inici')), 'sense res posat, torna a l\'inici sense preguntar');
  await p.click('#btn-generar');
  await p.waitForSelector('#navegacio:not(.ocult)', { timeout: 5000 });
  await p.evaluate(() => { const b = document.getElementById('btn-seguent'); if (document.getElementById('pas-cataleg').classList.contains('ocult') === false) { b.click(); } });
  await p.waitForSelector('#in-gia', { state: 'visible', timeout: 5000 });
  await p.fill('#in-gia', '1447');
  let preguntes = 0;
  p.once('dialog', (d) => { preguntes++; d.dismiss(); });
  await p.click('#btn-tornar-inici');
  await p.waitForTimeout(300);
  check(preguntes === 1 && await p.isVisible('#navegacio') && (await p.inputValue('#in-gia')) === '1447', 'amb un informe començat, pregunta; si dius que no, continua on era');
  p.once('dialog', (d) => { preguntes++; d.accept(); });
  await p.goBack();
  await p.waitForSelector('#inici', { state: 'visible', timeout: 5000 });
  check(preguntes === 2, 'l\'enrere del telèfon fa el mateix (pregunta i, si dius que sí, torna a l\'inici)');
  await p.click('#btn-generar');
  await p.waitForSelector('#navegacio:not(.ocult)', { timeout: 5000 });
  check((await p.inputValue('#in-gia')) === '', 'i l\'informe que hi havia s\'ha descartat');

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

  seccio('Següent sense confirmar el titular: avís sobre el botó');
  check(!(await p.isVisible('#msg-nav')), 'sense res a dir, no hi ha cap avís');
  await p.click('#btn-seguent');
  check(await p.isVisible('#pas-capcalera'), 'no avança...');
  check(await p.isVisible('#msg-nav') && (await p.textContent('#msg-nav')).includes('Confirmo que les dades del titular'), '...i diu per què, a sobre del botó Següent');
  const posAvis = await p.evaluate(() => [document.getElementById('msg-nav').getBoundingClientRect().bottom, document.getElementById('btn-seguent').getBoundingClientRect().top]);
  check(posAvis[0] <= posAvis[1] + 1, 'l\'avís queda damunt del botó');
  eq(await p.evaluate(() => getComputedStyle(document.getElementById('msg-nav')).color), 'rgb(192, 57, 43)', 'en vermell (el color d\'error de l\'app)');
  await p.check('#chk-titular');
  check(!(await p.isVisible('#msg-nav')), 'en marcar la casella, l\'avís se\'n va');

  seccio('La CCO del mòbil: la de sempre marcada i les de triar, desmarcades');
  const cco = await p.evaluate(() => new Promise((ok) => {
    const c = document.createElement('div'); document.body.appendChild(c);
    Cco.muntar(c, []);
    setTimeout(() => ok({ grups: Array.from(c.querySelectorAll('.bcc-grup')).map((g) => g.textContent),
      caselles: Array.from(c.querySelectorAll('input')).map((i) => i.value + ':' + i.checked), sel: Cco.seleccionat(c) }), 200);
  }));
  eq(cco.grups, ['CCO sempre (desmarca-la si no la vols):', 'CCO a triar:'], 'dos grups: CCO sempre i CCO a triar');
  eq(cco.caselles, ['sempre@x.cat:true', 'tria1@x.cat:false', 'tria2@x.cat:false'], 'la de sempre marcada, les a triar desmarcades');
  eq(cco.sel, 'sempre@x.cat', 'i se n\'envia la marcada');
  const ccoKo = await p.evaluate(() => new Promise((ok) => {
    window.__ccoFalla = true;
    const c = document.createElement('div'); document.body.appendChild(c);
    Cco.muntar(c, [], true);
    setTimeout(() => { const t = c.textContent; window.__ccoFalla = false; c.querySelector('a.bcc-reintent').click();
      setTimeout(() => ok({ t: t, despres: c.querySelectorAll('input').length }), 200); }, 200);
  }));
  check(ccoKo.t.includes("No s'ha pogut llegir la CCO") && ccoKo.t.includes('Tornar a carregar'), 'si el Drive no la dona, ho diu (no surt buit) i es pot tornar a provar');
  eq(ccoKo.despres, 3, '...i en tornar-ho a provar, surten les caselles');
  const formes = await p.evaluate(() => [
    Cco.llistaDe({ bcc: { value: [{ addr: 'a@x.cat', def: true }, { addr: 'b@x.cat', def: false }], Count: 2 } }),
    Cco.llistaDe({ bcc: { addr: 'sol@x.cat', def: true } }),
    Cco.llistaDe({ bcc: [{ Addr: 'maj@x.cat', Default: true }, null, { addr: '' }] }),
    Cco.llistaDe({})].map((l) => l.map((e) => e.addr + ':' + e.def).join(',')));
  eq(formes, ['a@x.cat:true,b@x.cat:false', 'sol@x.cat:true', 'maj@x.cat:true', ''], 'la CCO es llegeix encara que el fitxer vingui en una altra forma (value/Count, un sol objecte, majúscules)');
  const buida = await p.evaluate(() => new Promise((ok) => {
    const orig = Drive.llegirCorreuMobil;
    Drive.llegirCorreuMobil = () => Promise.resolve({ bcc: [], __modificat: '2026-10-20T09:15:00Z' });
    const c = document.createElement('div'); document.body.appendChild(c);
    Cco.muntar(c, [], true);
    setTimeout(() => { Drive.llegirCorreuMobil = orig; ok(c.textContent); }, 200);
  }));
  check(buida.includes('no porta cap adreça') && buida.includes('20/10/2026') && buida.includes('Desar'), 'si el fitxer del Drive és buit, diu de quan és i com tornar-lo a pujar');
  eq(new URL(p.url()).search, '', 'i l\'adreça ja no porta el ?gia (tornar a carregar no torna a cercar)');

  seccio('Al PC (fora de l\'app) no hi ha «Fer informe»');
  const pc = await ctx.newPage();
  await pc.goto(BASE + '__planol.html');
  await pc.waitForFunction(() => typeof capes !== 'undefined' && capes.length === PARCELES.length, null, { timeout: 15000 });
  await pc.evaluate(() => { const i = PARCELES.findIndex((x) => x.k === '2295827DF2729E'); obreFitxaActivitat('1447', PARCELES[i].c); });
  eq(await pc.locator('.leaflet-popup-content a.informe').count(), 0, 'el plànol obert sol no ensenya el botó');

  seccio('On sóc (la ubicació del mòbil al plànol)');
  const pg = await ctx.newPage();
  pg.on('pageerror', (e) => errors.push(String(e)));
  await pg.goto(BASE + 'planol.html');
  await pg.waitForSelector('#marc', { state: 'visible', timeout: 10000 });
  const fg = pg.frames().find((f) => f !== pg.mainFrame());
  await fg.waitForFunction(() => typeof capes !== 'undefined' && capes.length === PARCELES.length, null, { timeout: 15000 });
  check(await pg.isVisible('#btn-ubica'), 'hi ha el botó «On sóc» a sobre del plànol');
  await pg.tap('#btn-ubica');
  await fg.waitForFunction(() => { let n = 0; map.eachLayer((l) => { if (l.options && l.options.fillColor === '#1a73e8' && !(l instanceof L.Circle)) { n++; } }); return n > 0; }, null, { timeout: 8000 }).catch(() => {});
  const ubi = await fg.evaluate(() => { let ll = null; map.eachLayer((l) => { if (l instanceof L.CircleMarker && !(l instanceof L.Circle) && !l.feature && l.options.fillColor === '#1a73e8') { ll = l.getLatLng(); } }); const c = map.getCenter(); return { ll: ll && [ll.lat, ll.lng], c: [c.lat, c.lng], z: map.getZoom() }; });
  eq(ubi.ll, [41.356, 2.074], 'el punt blau és on és el mòbil');
  check(Math.abs(ubi.c[0] - 41.356) < 0.0005 && Math.abs(ubi.c[1] - 2.074) < 0.0005 && ubi.z >= 17, 'i el mapa s\'hi centra, de prop');
  check(await pg.evaluate(() => document.getElementById('btn-ubica').classList.contains('actiu')), 'el botó queda encès');
  await pg.tap('#btn-ubica');
  const queden = await fg.evaluate(() => { let n = 0; map.eachLayer((l) => { if (l.options && l.options.fillColor === '#1a73e8') { n++; } }); return n; });
  eq(queden, 0, 'un altre toc l\'apaga i treu el punt');

  seccio('Google: no tornar a iniciar sessió a cada pàgina (drive.js de debò)');
  // Un Google simulat: cada petició de token s'apunta a window.__peticions.
  const ctxG = await nav.newContext({ viewport: { width: 390, height: 800 }, isMobile: true, hasTouch: true });
  await ctxG.route('https://accounts.google.com/gsi/client', (r) => r.fulfill({ status: 200, contentType: 'application/javascript', body: `
    window.__peticions = window.__peticions || [];
    window.google = { accounts: { oauth2: { initTokenClient: function (cfg) {
      var c = { callback: null, error_callback: null, requestAccessToken: function (o) {
        window.__peticions.push(o || {});
        var self = this;
        setTimeout(function () {
          if (o && o.prompt === 'none' && window.__senseSessio) { self.error_callback({ type: 'popup_failed_to_open', message: 'bloquejada' }); return; }
          self.callback({ access_token: 'tok' + window.__peticions.length, expires_in: 3599 });
        }, 10);
      } };
      return c;
    } } } };` }));
  await ctxG.route(/googleapis\.com\/drive\/v3\/about/, (r) => r.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify({ user: { emailAddress: 'jo@exemple.cat' } }) }));
  await ctxG.route(/googleapis\.com\/drive\/v3\/files/, (r) => r.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify({ files: [] }) }));
  await ctxG.route(/emailjs/, (r) => r.abort());
  const g = await ctxG.newPage();
  g.on('pageerror', (e) => errors.push(String(e)));
  await g.goto(BASE + 'index.html');
  await g.waitForFunction(() => window.Drive && typeof Drive.connectar === 'function');
  const primera = await g.evaluate(() => Drive.connectar().then(() => ({ ok: Drive.connectat(), p: window.__peticions })));
  check(primera.ok, 'el primer toc connecta');
  eq(primera.p.map((x) => x.prompt), [''], 'una sola petició, i sense forçar la pantalla de permisos (prompt "", no "consent")');
  await g.waitForFunction(() => localStorage.getItem('driveCompte') === 'jo@exemple.cat', null, { timeout: 3000 }).catch(() => {});
  eq(await g.evaluate(() => localStorage.getItem('driveCompte')), 'jo@exemple.cat', 'es recorda el compte');
  await g.goto(BASE + 'index.html?gia=1447');
  await g.waitForFunction(() => window.Drive && typeof Drive.connectat === 'function');
  await g.waitForTimeout(300);
  check(await g.evaluate(() => Drive.connectat()), 'en canviar de pàgina (com «Fer informe» des del plànol), segueix connectat');
  eq(await g.evaluate(() => (window.__peticions || []).length), 0, '...sense tornar a demanar res a Google');
  check(await g.evaluate(() => Drive.reconnectarSilenci()), 'i la reconnexió de l\'arrencada ho troba fet');
  const caducat = await g.evaluate(() => {
    sessionStorage.setItem('driveToken', JSON.stringify({ t: 'vell', exp: Date.now() + 30000 }));
    return true;
  });
  // Sense sessio de Google, la reconnexio de l'arrencada (sense toc) no en pot treure cap.
  await g.addInitScript(() => { window.__senseSessio = true; });
  await g.goto(BASE + 'index.html');
  await g.waitForFunction(() => window.Drive && typeof Drive.connectat === 'function');
  await g.waitForTimeout(300);
  check(caducat && !(await g.evaluate(() => Drive.connectat())), 'un token a punt de caducar no es fa servir');
  eq(await g.evaluate(() => (window.__peticions || []).map((x) => x.prompt)), ['none'], 'i la reconnexió de l\'arrencada no ensenya res (prompt "none")');
  const nova = await g.evaluate(() => { window.__peticions = []; return Drive.connectar().then(() => window.__peticions); });
  eq(nova.map((x) => x.prompt + '|' + x.hint), ['|jo@exemple.cat'], 'en tornar a demanar-ne, es diu el compte (no cal triar-lo)');
  await ctxG.close();

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
