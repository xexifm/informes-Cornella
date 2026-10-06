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
import { pathToFileURL } from 'node:url';
import { AQUI, LEAFLET, check, eq, seccio, serveixLeaflet, resultat, comptes } from './comu.mjs';

const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'coord-nav-'));
const PERFIL = path.join(TMP, 'perfil');
const CLAU = 'coordenades:2026-08-18 ACTIVITATS.xls';

// 1. Els mapes de prova, amb les funcions de debo de Coordenades.ps1.
const gen = spawnSync('pwsh', ['-NoProfile', '-File', path.join(AQUI, 'genera-mapa-prova.ps1'), '-Dir', TMP],
                      { stdio: 'inherit' });
if (gen.status !== 0) { console.error('No s\'han pogut generar els mapes de prova.'); process.exit(2); }
if (!fs.existsSync(path.join(LEAFLET, 'leaflet.js'))) {
  console.error('Falta node_modules/leaflet: executa "npm install" a ' + AQUI); process.exit(2);
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
  // Que el mapa hagi acabat de moure's (vesA fa zoom i obre la fitxa, i el
  // popup el desplaça): un marcador a mig moviment no s'arrossega (el Leaflet
  // ignora el clic durant l'animacio) i la prova fallava a vegades.
  await page.evaluate(() => new Promise((ok) => {
    let t = setTimeout(ok, 500);
    map.on('movestart zoomstart', () => { clearTimeout(t); });
    map.on('moveend zoomend', () => { clearTimeout(t); t = setTimeout(ok, 300); });
  }));
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

  seccio('El titular');
  check((await a.locator('#tbody tr').first().textContent()).includes('EL RACO DE CADIS SL'), 'a la llista, sota l\'adreça');
  await a.evaluate(() => capes[0].verd.openPopup());
  check((await a.textContent('.leaflet-popup-content')).includes('Titular: EL RACO DE CADIS SL'), 'i a la fitxa del punt');
  await a.evaluate(() => map.closePopup());
  await a.fill('#cerca', 'raco de cadis');
  eq(await a.evaluate(() => capes.filter((c) => c.fila.style.display !== 'none').length), 1, 'el cercador també busca pel titular');
  await a.fill('#cerca', '');
  await a.dispatchEvent('#cerca', 'input');

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
    eq(files[0][10], 'Base de dades', 'l\'última columna diu de quina base és el repàs');
    eq(files[1][10], '2026-08-18 ACTIVITATS.xls', 'i cada fila porta el nom de la base');
  } else {
    console.log('  (sense python3 + openpyxl: no es rellegeix l\'Excel) ' + (py.stderr || '').split('\n').slice(-2).join(' '));
  }

  seccio('Projecció d\'anada i tornada (per carregar un repàs, que només porta metres)');
  const errMax = await a.evaluate(() => {
    let pitjor = 0;
    for (let x = 418000; x <= 426000; x += 500) {
      for (let y = 4574000; y <= 4582000; y += 500) {
        const ll = utm31ToLatLon(x, y), u = latLonToUtm31(ll[0], ll[1]);
        pitjor = Math.max(pitjor, Math.abs(u[0] - x), Math.abs(u[1] - y));
      }
    }
    return pitjor;
  });
  check(errMax < 0.001, 'UTM -> graus -> UTM sobre tot el terme: error màxim ' + (errMax * 1000).toFixed(3) + ' mm (< 1 mm)');

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

  seccio('Carregar el repàs des de l\'Excel baixat (la còpia de seguretat)');
  const carrega = async (p, fitxer) => {
    const dlg = p.waitForEvent('dialog');
    await p.setInputFiles('#fitxerRepas', fitxer);
    return (await dlg).message();
  };
  const posAbans = await a2.evaluate(() => [estat[1].lat, estat[1].lon]);
  // b ha esborrat el repàs de la base A: el navegador l'ha perdut.
  await a2.reload();
  await a2.waitForFunction(() => typeof capes !== 'undefined' && capes.length === ITEMS.length);
  eq(await a2.evaluate(() => estat.filter((e) => e.revisada).length), 0, 'sense repàs al navegador (esborrat)');
  let msg = await carrega(a2, xlsx);
  check(msg.includes('Recuperades 2'), 'recarregant l\'Excel baixat es recuperen les 2: «' + msg.split('\n')[0] + '»');
  const rec = await a2.evaluate(() => [estat[0].revisada, estat[1].origen, estat[1].lat, estat[1].lon]);
  eq(rec.slice(0, 2), [true, 'manual'], 'amb el seu estat (validada, moguda a mà)');
  check(Math.abs(rec[2] - posAbans[0]) < 1e-7 && Math.abs(rec[3] - posAbans[1]) < 1e-7,
        'i el punt mogut torna exactament on era (< 1 cm)');
  msg = await carrega(a2, xlsx);
  check(msg.includes('ja eren al navegador'), 'carregar-lo dos cops no duplica res: «' + msg.split('\n')[1] + '»');
  // El mateix fitxer obert i desat amb un altre programa: comprimit i amb els
  // textos a sharedStrings.xml, com quan l'obres i el deses amb l'Excel.
  const desatAmbAltre = path.join(TMP, 'desat-amb-excel.xlsx');
  // (simula-excel.py: openpyxl no serveix, desa els textos inline com nosaltres)
  const pyDesa = spawnSync('python3', [path.join(AQUI, 'simula-excel.py'), xlsx, desatAmbAltre], { encoding: 'utf8' });
  if (pyDesa.status === 0) {
    const zipCru = fs.readFileSync(desatAmbAltre).toString('latin1');
    check(zipCru.includes('sharedStrings.xml') && !zipCru.includes('Coordenades<'),
          '(la còpia desada de nou porta sharedStrings i va comprimida, com la de l\'Excel)');
    // Desa'n una còpia per a la suite de PowerShell, que ha de llegir el mateix.
    if (process.env.COORD_DESA_FIXTURES) {
      fs.copyFileSync(xlsx, path.join(process.env.COORD_DESA_FIXTURES, 'repas-navegador.xlsx'));
      fs.copyFileSync(desatAmbAltre, path.join(process.env.COORD_DESA_FIXTURES, 'repas-desat-excel.xlsx'));
    }
    await a2.evaluate(() => { localStorage.clear(); });
    await a2.reload();
    await a2.waitForFunction(() => typeof capes !== 'undefined' && capes.length === ITEMS.length);
    msg = await carrega(a2, desatAmbAltre);
    check(msg.includes('Recuperades 2'), 'un Excel desat de nou (comprimit, sharedStrings) també es carrega');
    const r2 = await a2.evaluate(() => [estat[1].lat, estat[1].lon]);
    check(Math.abs(r2[0] - posAbans[0]) < 1e-7 && Math.abs(r2[1] - posAbans[1]) < 1e-7, 'i amb les mateixes coordenades');
  } else {
    console.log('  (sense python3 + openpyxl: no es prova l\'Excel desat de nou)');
  }
  msg = await carrega(c, xlsx);
  check(msg.includes('altra base de dades') && msg.includes('2026-08-18 ACTIVITATS.xls'),
        'en un mapa d\'una ALTRA base no es carrega, i diu de quina és');
  eq(await c.evaluate(() => estat.filter((e) => e.revisada).length), 0, 'i no hi toca res');
  const noRepas = path.join(TMP, 'no-repas.xlsx');
  if (spawnSync('python3', ['-c', 'import sys, openpyxl; wb = openpyxl.Workbook(); wb.active.append(["Nom", "Cognom"]); wb.save(sys.argv[1])', noRepas]).status === 0) {
    msg = await carrega(c, noRepas);
    check(msg.includes('No s\'ha pogut carregar') && msg.includes('ID GIA'), 'un Excel que no és un repàs: ho diu i no peta');
  }

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

  seccio('Zoom de prop (l\'usuari: «deixa\'m fer més zoom»)');
  eq(await c.evaluate(() => map.getMaxZoom()), 22, 'el mapa arriba fins al zoom 22 (abans, 19)');
  await capsaVerd(c, 2);   // que s'hagi acabat el moviment de vesA(2)
  eq(await c.evaluate(() => { map.setZoom(21.5, { animate: false }); return map.getZoom(); }), 21.5, 'es pot apropar més enllà del 19');
  eq(await c.evaluate(() => { vesA(1); return map.getZoom(); }), 21.5, 'anar a una activitat no t\'allunya si ja eres més a prop');
  await c.evaluate(() => { map.setZoom(15, { animate: false }); vesA(2); });
  await capsaVerd(c, 2);   // el zoom de vesA va animat
  eq(await c.evaluate(() => map.getZoom()), 19, 'i si eres lluny, t\'hi apropa (19)');

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

  seccio('El plànol del Cadastre (a sobre del fons)');
  eq(await c.evaluate(() => [document.getElementById('chkCadastre').checked, map.hasLayer(FONS.cadastre)]), [false, false],
     'apagat per defecte (carrega)');
  const [wms] = await Promise.all([c.waitForRequest(/ServidorWMS\.aspx/), c.check('#chkCadastre')]);
  check(/layers=Catastro/i.test(wms.url()) && /transparent=true/i.test(wms.url()), 'la casella l\'encén: demana el WMS oficial del Cadastre, transparent');
  eq(await c.evaluate(() => map.hasLayer(FONS.cadastre)), true, 'i la capa és al mapa');
  eq(await c.evaluate(() => document.querySelector('.leaflet-control-layers-overlays input').checked), true,
     'el selector de dalt diu el mateix (és la mateixa capa)');
  const c2 = await obre(ctx, 'a/Coordenades_A.html');
  eq(await c2.evaluate(() => [document.getElementById('chkCadastre').checked, map.hasLayer(FONS.cadastre)]), [true, true],
     'es recorda: un mapa nou ja l\'obre ences');
  await c2.evaluate(() => document.querySelector('.leaflet-control-layers-overlays input').click());
  eq(await c2.evaluate(() => [document.getElementById('chkCadastre').checked, map.hasLayer(FONS.cadastre)]), [false, false],
     'apagat des del selector, la casella també s\'apaga');
  await c2.close();
  await c.uncheck('#chkCadastre');

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
  comptes.fail++;
  console.log('  FAIL la prova ha petat: ' + (e && e.stack || e));
} finally {
  await ctx.close().catch(() => {});
  fs.rmSync(TMP, { recursive: true, force: true });
}

process.exit(resultat('NAVEGADOR'));
