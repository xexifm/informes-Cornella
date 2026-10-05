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

// Quins CDN responen: per defecte, unpkg (el de la pagina). Amb { unpkg: false }
// es prova el segon intent (jsDelivr) i amb tots dos a false, el missatge.
async function serveixLeaflet(ctx, cdn = { unpkg: true, jsdelivr: true }) {
  const serveix = (actiu) => (route) => {
    if (!actiu) { route.abort(); return; }
    const nom = path.basename(new URL(route.request().url()).pathname);
    route.fulfill({
      path: path.join(LEAFLET, nom),
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Content-Type': nom.endsWith('.css') ? 'text/css' : 'application/javascript',
      },
    });
  };
  await ctx.route('https://unpkg.com/leaflet@1.9.4/dist/**', serveix(cdn.unpkg));
  await ctx.route('https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/**', serveix(cdn.jsdelivr));
  await ctx.route(/tile\.openstreetmap\.org/, (route) => route.abort());
}

async function obreContext() {
  const ctx = await chromium.launchPersistentContext(PERFIL, { headless: true, acceptDownloads: true });
  await serveixLeaflet(ctx);
  return ctx;
}

const errorsPagina = [];
async function obre(ctx, rel, { espera = true } = {}) {
  const page = await ctx.newPage();
  page.errors = [];
  page.on('pageerror', (e) => { page.errors.push(String(e)); if (espera) errorsPagina.push(rel + ': ' + e); });
  page.dialegs = [];
  page.on('dialog', (d) => { page.dialegs.push(d.message()); d.accept(); });
  await page.goto(pathToFileURL(path.join(TMP, rel)).href);
  if (espera) {
    await page.waitForFunction(() => typeof capes !== 'undefined' && capes.length === ITEMS.length,
                               null, { timeout: 10000 });
  }
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
  // 2 i no 1: la que acabes d'arrossegar (102) es queda a la vista mentre la
  // tinguis seleccionada, encara que no passi el filtre.
  eq(await a.locator('#tbody tr:visible td.id').allTextContents(), ['102', '103'],
     'el cercador filtra per adreça (i la seleccionada no desapareix)');
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

  seccio('Filtre per estat i recomptes (mapa d\'una altra base, comença net)');
  const visibles = (p) => p.locator('#tbody tr:visible td.id').allTextContents();
  await c.selectOption('#filtreEstat', 'cadastre');
  eq(await visibles(c), ['104'], 'només les que no tenen portal');
  eq(await c.locator('.marker-verd').count(), 1, 'i al mapa també només aquella');
  check((await c.textContent('#filtreEstat')).includes('Sense portal (blanc) (1)'), 'el desplegable diu quantes n\'hi ha');
  check((await c.textContent('#llegenda')).includes('portal dubtós'), 'la llegenda porta tots els colors, també el groc');
  await c.selectOption('#filtreEstat', 'tots');
  eq((await visibles(c)).length, 6, 'i «Totes» les torna a mostrar');

  seccio('Ressaltar la parella');
  await c.evaluate(() => vesA(1));
  eq(await c.evaluate(() => [capes[1].fila.classList.contains('sel'),
                             capes[1].verd.getElement().firstChild.classList.contains('sel'),
                             capes[1].linia.options.weight, capes[1].vermell.getRadius()]),
     [true, true, 3, 8], 'fila, punt verd, línia i vermell ressaltats');
  await c.evaluate(() => vesA(2));
  eq(await c.evaluate(() => [capes[1].fila.classList.contains('sel'), capes[1].linia.options.weight,
                             capes[1].vermell.getRadius(), capes[2].linia.options.weight]),
     [false, 1, 5, 3], 'en triar-ne una altra, l\'anterior torna a la normalitat');

  seccio('Següent pendent (tecla N)');
  await c.evaluate(() => vesA(0));
  await c.click('#map', { position: { x: 5, y: 5 } });   // el focus fora del cercador
  const ordre = [];
  for (let k = 0; k < 5; k++) { await c.keyboard.press('n'); ordre.push(await c.evaluate(() => sel)); }
  eq(ordre.slice(0, 3).sort(), [1, 2, 3], 'primer les apilades al mateix edifici (les més properes)');
  eq(ordre.slice(3).sort(), [4, 5], 'després salta al següent edifici');
  eq(new Set(ordre).size, 5, 'cap repetida: no torna a la d\'on has sortit');
  await c.keyboard.press('n');
  eq(await c.evaluate(() => sel), 4, 'quan les ha ensenyat totes, torna a començar (per la més propera)');
  await c.selectOption('#filtreEstat', 'pendents');
  await c.evaluate(() => vesA(3));
  await c.evaluate(() => commutaValidada(3));
  check((await visibles(c)).includes('104'), 'amb «Pendents», la que acabes de validar no et desapareix de sota');
  await c.click('#map', { position: { x: 5, y: 5 } });
  await c.keyboard.press('n');
  check(!(await visibles(c)).includes('104'), 'i marxa quan passes a la següent');
  await c.selectOption('#filtreEstat', 'tots');

  seccio('Desfer (Ctrl+Z)');
  await c.evaluate(() => vesA(0));
  const b5 = await capsaVerd(c, 0);
  await c.mouse.move(b5.x + b5.width / 2, b5.y + b5.height / 2);
  await c.mouse.down();
  await c.mouse.move(b5.x + 60, b5.y + 40, { steps: 8 });
  await c.mouse.up();
  eq(await c.evaluate(() => estat[0].origen), 'manual', 'arrossegat');
  await c.keyboard.press('Control+z');
  eq(await c.evaluate(() => [estat[0].origen, estat[0].revisada, estat[0].lat === ITEMS[0].latf]),
     ['facana', false, true], 'Ctrl+Z el torna al portal, sense validar, com era');
  eq(await c.evaluate(() => Object.keys(JSON.parse(localStorage.getItem('coordenades:2026-10-01 ACTIVITATS.xls') || '{}'))),
     ['104'], 'i al navegador només hi queda el que sí que has validat');
  // Desvalidar un punt MOGUT el torna al Cadastre: era la manera de perdre
  // sense voler la posició que li havies donat.
  await c.evaluate(() => vesA(1));
  const b6 = await capsaVerd(c, 1);
  await c.mouse.move(b6.x + b6.width / 2, b6.y + b6.height / 2);
  await c.mouse.down();
  await c.mouse.move(b6.x - 50, b6.y + 20, { steps: 8 });
  await c.mouse.up();
  const mogut = await c.evaluate(() => [estat[1].lat, estat[1].lon]);
  const b7 = await capsaVerd(c, 1);
  await c.mouse.click(b7.x + b7.width / 2, b7.y + b7.height / 2);
  eq(await c.evaluate(() => estat[1].revisada), false, 'un clic al punt mogut el desvalida (i el torna al portal)');
  await c.click('#btnDesfer');
  eq(await c.evaluate(() => [estat[1].lat, estat[1].lon, estat[1].origen, estat[1].revisada]),
     [...mogut, 'manual', true], '«Desfer» recupera la posició que li havies donat');
  await c.focus('#cerca');
  const pila = await c.evaluate(() => pilaDesfer.length);
  await c.keyboard.press('Control+z');
  eq(await c.evaluate(() => pilaDesfer.length), pila, 'dins del cercador, Ctrl+Z és del text i no desfà cap punt');
  while (await c.evaluate(() => pilaDesfer.length) > 0) { await c.click('#btnDesfer'); }
  check(await c.locator('#btnDesfer').isDisabled(), 'sense res per desfer, el botó queda apagat');
  eq(await c.evaluate(() => estat.filter((e) => e.revisada).length), 0, 'i desfent-ho tot, es torna a l\'inici');

  seccio('Si el CDN del mapa no respon');
  {
    const nav = await chromium.launch({ headless: true });
    try {
      const c1 = await nav.newContext();
      await serveixLeaflet(c1, { unpkg: false, jsdelivr: true });
      const p1 = await obre(c1, 'a/Coordenades_A.html');
      eq(await p1.locator('.marker-verd').count(), 6, 'unpkg caigut: el mapa arrenca amb jsDelivr (mateix SRI)');
      const c2 = await nav.newContext();
      await serveixLeaflet(c2, { unpkg: false, jsdelivr: false });
      const p2 = await obre(c2, 'a/Coordenades_A.html', { espera: false });
      check((await p2.textContent('#map')).includes("No s'ha pogut carregar el mapa"),
            'tots dos caiguts: la pàgina ho diu en clar, no es queda en blanc');
      check(p2.errors.length === 1 && p2.errors[0].includes('Leaflet no carregat'),
            'i el codi del mapa s\'atura amb un sol error, que ho explica');
    } finally {
      await nav.close();
    }
  }

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
