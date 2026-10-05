// prova-mapa-coordenades.mjs - Prova el mapa de l'eina Coordenades en un
// Chromium DE DEBO (Playwright), amb el Leaflet de veritat.
//
// Per que existeix: tot el que l'usuari fa al mapa (validar, arrossegar,
// recuperar el repas l'endema, baixar l'Excel) es JavaScript, i la suite de
// PowerShell nomes pot mirar que l'HTML porti el que toca. El setembre de 2026
// es va provar a ma un sol cop; aixo ho fa repetible.
//
// Com s'executa (Linux de desenvolupament, amb Node i pwsh; al PC de la feina
// no cal):
//   cd suport/tests/navegador && npm install && node prova-mapa-coordenades.mjs
//
// El Leaflet: el mapa el demana a unpkg amb SRI. Aqui es serveix des de
// node_modules/leaflet (el MATEIX paquet 1.9.4 de npm, byte a byte: el hash
// del SRI hi coincideix), aixi que NO cal treure l'integrity a cap copia. Si
// algun dia no coincidis, el navegador no carregaria el Leaflet i la prova ho
// diria a la primera comprovacio. Les tessel.les d'OpenStreetMap s'avorten: el
// mapa funciona igual sense el fons.
//
// El navegador es PERSISTENT (un perfil en una carpeta temporal): es l'unica
// manera de provar que el repas sobreviu a tancar el Chrome i a obrir un mapa
// nou, amb un altre nom de fitxer, l'endema.

import { chromium } from 'playwright';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const AQUI = path.dirname(fileURLToPath(import.meta.url));
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'coord-nav-'));
const PERFIL = path.join(TMP, 'perfil');
const LEAFLET = path.join(AQUI, 'node_modules', 'leaflet', 'dist');
const CLAU = 'coordenades:2026-08-18 ACTIVITATS.xls';

let ok = 0, fail = 0;
function check(cond, msg) {
  if (cond) { ok++; console.log('  OK   ' + msg); } else { fail++; console.log('  FAIL ' + msg); }
}
function eq(obtingut, esperat, msg) {
  const igual = JSON.stringify(obtingut) === JSON.stringify(esperat);
  check(igual, msg + (igual ? '' : ` (esperat ${JSON.stringify(esperat)}, obtingut ${JSON.stringify(obtingut)})`));
}
function seccio(t) { console.log('\n--- ' + t + ' ---'); }

// 1. Els mapes de prova, amb les funcions de debo de Coordenades.ps1.
const gen = spawnSync('pwsh', ['-NoProfile', '-File', path.join(AQUI, 'genera-mapa-prova.ps1'), '-Dir', TMP],
                      { stdio: 'inherit' });
if (gen.status !== 0) { console.error('No s\'han pogut generar els mapes de prova.'); process.exit(2); }
if (!fs.existsSync(path.join(LEAFLET, 'leaflet.js'))) {
  console.error('Falta node_modules/leaflet: executa "npm install" a ' + AQUI); process.exit(2);
}

async function obreContext() {
  const ctx = await chromium.launchPersistentContext(PERFIL, { headless: true, acceptDownloads: true });
  await ctx.route('https://unpkg.com/leaflet@1.9.4/dist/**', (route) => {
    const nom = path.basename(new URL(route.request().url()).pathname);
    route.fulfill({
      path: path.join(LEAFLET, nom),
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Content-Type': nom.endsWith('.css') ? 'text/css' : 'application/javascript',
      },
    });
  });
  await ctx.route(/tile\.openstreetmap\.org/, (route) => route.abort());
  return ctx;
}

const errorsPagina = [];
async function obre(ctx, rel) {
  const page = await ctx.newPage();
  page.on('pageerror', (e) => errorsPagina.push(rel + ': ' + e));
  page.dialegs = [];
  page.on('dialog', (d) => { page.dialegs.push(d.message()); d.accept(); });
  await page.goto(pathToFileURL(path.join(TMP, rel)).href);
  await page.waitForFunction(() => typeof capes !== 'undefined' && capes.length === ITEMS.length,
                             null, { timeout: 10000 });
  return page;
}

// Centra el mapa en una activitat al zoom maxim, perque les apilades quedin
// separades i el clic vagi al punt que toca.
async function centra(page, i) {
  await page.evaluate((i) => { map.setView([estat[i].lat, estat[i].lon], 19, { animate: false }); }, i);
}
async function capsaVerd(page, i) {
  const h = await page.evaluateHandle((i) => capes[i].verd.getElement(), i);
  return h.asElement().boundingBox();
}
const desat = (page) => page.evaluate((c) => JSON.parse(localStorage.getItem(c) || '{}'), CLAU);

let ctx = await obreContext();
try {
  seccio('Arrencada');
  const a = await obre(ctx, 'a/Coordenades_A.html');
  eq(await a.evaluate(() => typeof L), 'object', 'el Leaflet carrega (amb el SRI)');
  eq(await a.locator('.marker-verd').count(), 6, 'un punt verd per activitat');
  eq(await a.locator('#tbody tr').count(), 6, 'una fila per activitat');
  check((await a.textContent('#bar')).includes('Esborrar el meu repàs'), 'els accents de la plantilla surten bé');
  eq(await a.evaluate(() => ITEMS[3].adreca), 'C/ Falsa 1 </script><b>',
     'un </script> a l\'adreça no trenca la pàgina');

  seccio('Validar amb un clic');
  await centra(a, 0);
  let b0 = await capsaVerd(a, 0);
  await a.mouse.click(b0.x + b0.width / 2, b0.y + b0.height / 2);
  eq(await a.evaluate(() => estat[0].revisada), true, 'el clic valida el punt');
  eq(Object.keys(await desat(a)), ['101'], 'i queda desat al navegador per ID');
  check(await a.evaluate(() => capes[0].fila.classList.contains('revisada')), 'la fila surt com a validada');

  seccio('Arrossegar');
  await centra(a, 1);
  const b1 = await capsaVerd(a, 1);
  const cx = b1.x + b1.width / 2, cy = b1.y + b1.height / 2;
  await a.mouse.move(cx, cy);
  await a.mouse.down();
  await a.mouse.move(cx + 30, cy, { steps: 5 });
  await a.mouse.move(cx + 60, cy, { steps: 5 });
  await a.mouse.up();
  const e1 = await a.evaluate(() => estat[1]);
  eq(e1.origen, 'manual', 'arrossegat: queda com a mogut a mà');
  eq(e1.revisada, true, 'i validat (el clic que Leaflet llança després del drag no el desvalida)');
  const d1 = (await desat(a))['102'];
  const xf1 = await a.evaluate(() => ITEMS[1].xf);
  check(d1 && d1.x - xf1 > 5, 'la X nova desada s\'ha mogut cap a l\'est (' + (d1 ? (d1.x - xf1).toFixed(1) : '?') + ' m)');

  seccio('Filtre');
  await a.fill('#cerca', 'huelva');
  eq(await a.locator('#tbody tr:visible').count(), 1, 'el cercador filtra per adreça');
  await a.fill('#cerca', '');
  eq(await a.locator('#tbody tr:visible').count(), 6, 'i buidar-lo ho torna a mostrar tot');

  seccio('Baixar l\'Excel');
  const [baixada] = await Promise.all([a.waitForEvent('download'), a.click('text=Baixar Excel (.xlsx)')]);
  const xlsx = path.join(TMP, 'baixat.xlsx');
  await baixada.saveAs(xlsx);
  const bytes = fs.readFileSync(xlsx);
  eq(bytes.subarray(0, 2).toString(), 'PK', 'el fitxer és un ZIP');
  const py = spawnSync('python3', ['-c', [
    'import sys, json, warnings, openpyxl',
    'warnings.simplefilter("error")',
    'ws = openpyxl.load_workbook(sys.argv[1]).active',
    'print(json.dumps([[c.value for c in r] for r in ws.iter_rows()]))',
  ].join('\n'), xlsx], { encoding: 'utf8' });
  if (py.status === 0) {
    const files = JSON.parse(py.stdout);
    eq(files[0][0], 'ID GIA', 'openpyxl l\'obre sense avisos i hi ha la capçalera');
    eq(files.slice(1).map((r) => r[0]).sort(), ['101', '102'], 'hi ha les dues activitats repassades');
    check(typeof files[1][6] === 'number', 'les coordenades són números, no text');
  } else {
    console.log('  (sense python3 + openpyxl: no es rellegeix l\'Excel) ' + (py.stderr || '').split('\n').slice(-2).join(' '));
  }

  seccio('Tornar-hi l\'endemà: tancar el navegador i obrir un mapa NOU');
  await ctx.close();
  ctx = await obreContext();
  const a2 = await obre(ctx, 'a/Coordenades_A.html');
  eq(await a2.evaluate(() => [estat[0].revisada, estat[1].origen]), [true, 'manual'],
     'el mateix mapa recupera el repàs');
  const b = await obre(ctx, 'b/Coordenades_B.html');
  eq(await b.evaluate(() => estat[0].revisada), true,
     'un mapa NOU (un altre nom i carpeta) de la mateixa base hi veu el repàs');
  check((await b.textContent('#estat')).includes('2 en total'), 'i compta les 2 validades de la base');
  const c = await obre(ctx, 'c/Coordenades_C.html');
  eq(await c.evaluate(() => estat.filter((e) => e.revisada).length), 0,
     'el mapa d\'una ALTRA base comença net');

  seccio('Esborrar el repàs');
  await b.click('text=Esborrar el meu repàs');
  check(b.dialegs.some((m) => m.includes('Segur')), 'demana confirmació');
  eq(await desat(b), {}, 'i, acceptat, el repàs d\'aquesta base desapareix');

  seccio('Errors de JavaScript');
  eq(errorsPagina, [], 'cap error de JavaScript a cap pàgina');
} catch (e) {
  fail++;
  console.log('  FAIL la prova ha petat: ' + (e && e.stack || e));
} finally {
  await ctx.close().catch(() => {});
  fs.rmSync(TMP, { recursive: true, force: true });
}

console.log('\n========================================');
console.log(`RESULTAT NAVEGADOR: ${ok} OK, ${fail} FAIL`);
console.log('========================================');
process.exit(fail === 0 ? 0 : 1);
