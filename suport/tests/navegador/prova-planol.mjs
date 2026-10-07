// prova-planol.mjs - Prova el "Planol activitats" (PlanolMapa.html) en un
// Chromium de debo, amb el Leaflet de veritat.
//
//   cd suport/tests/navegador && npm install && node prova-planol.mjs
//
// El mapa el genera genera-planol-prova.ps1 amb les funcions de debo de
// PlanolDades.ps1: Cadis 19 (una precintada, una amb requeriment, un local buit,
// amb la geometria de la fixture), Hospitalet 147 (favorable, en un local
// marcat com a buit, sense geometria: un punt), un local buit sol i una activitat
// sense informes.
import { chromium } from 'playwright';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { AQUI, LEAFLET, check, eq, seccio, serveixLeaflet, resultat, comptes, PNG_1x1 } from './comu.mjs';

const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'planol-nav-'));
const gen = spawnSync('pwsh', ['-NoProfile', '-File', path.join(AQUI, 'genera-planol-prova.ps1'), '-Dir', TMP], { stdio: 'inherit' });
if (gen.status !== 0) { console.error('No s\'ha pogut generar el planol de prova.'); process.exit(2); }
if (!fs.existsSync(path.join(LEAFLET, 'leaflet.js'))) { console.error('Falta node_modules: npm install'); process.exit(2); }

const nav = await chromium.launch({ headless: true });
const errors = [];
try {
  const ctx = await nav.newContext({ acceptDownloads: true, viewport: { width: 1280, height: 800 } });
  await serveixLeaflet(ctx);
  const p = await ctx.newPage();
  p.on('pageerror', (e) => errors.push(String(e)));
  const peticions = [];
  p.on('request', (r) => peticions.push(r.url()));
  await p.goto(pathToFileURL(path.join(TMP, 'Planol.html')).href);
  await p.waitForFunction(() => typeof capes !== 'undefined' && capes.length === PARCELES.length, null, { timeout: 10000 });
  const idx = (k) => p.evaluate((k) => PARCELES.findIndex((x) => x.k === k), k);
  const iCadis = await idx('2295827DF2729E'), iHosp = await idx('4091106DF2749A'), iBuit = await idx('1111111DF1111A');
  const vista = (i) => p.evaluate((i) => { const v = capes[i].vista; return { color: v.color, acts: [...new Set(v.acts.map((e) => e.g))], revisar: v.revisar, alMapa: capes[i].alMapa }; }, i);

  seccio('Arrencada');
  eq(await p.evaluate(() => typeof L), 'object', 'el Leaflet carrega (amb el SRI)');
  eq(await p.evaluate(() => PARCELES.length), 4, 'quatre parcel·les');
  check((await p.textContent('#avisos')).includes('</script> amb <b>'), 'els avisos surten escapats (cap etiqueta s\'interpreta)');
  eq(await p.evaluate((i) => capes[i].forma instanceof L.Polygon, iCadis), true, 'Cadis 19: un polígon (té geometria)');
  eq(await p.evaluate((i) => capes[i].forma instanceof L.CircleMarker, iHosp), true, 'Hospitalet 147: un punt (sense geometria)');
  eq(await p.evaluate(() => [document.getElementById('f-turisme').value, CLASSIFS]), ['no', ['II', 'III', 'L18 Cert', '']],
     'per defecte, sense allotjaments turístics; una casella per cada classificació (i la buida, al final)');

  seccio('Colors: mana el pitjor');
  eq(await vista(iCadis), { color: 'vermell', acts: ['1447', '1403'], revisar: false, alMapa: true }, 'Cadis 19: vermell (1447 precintada) amb 1403 en groc');
  eq((await vista(iHosp)).color, 'verd', 'Hospitalet 147: verd (favorable)');
  eq((await vista(iHosp)).revisar, true, 'i marcada per revisar: activitat en un local marcat com a buit');
  eq(await p.evaluate((i) => capes[i].forma.options.dashArray, iHosp), '5,4', 'amb la vora discontínua');
  eq((await vista(iBuit)).alMapa, false, 'el local buit sol: no es pinta per defecte');
  eq(await p.evaluate(() => ['vermell', 'groc', 'blau', 'verd', 'buits'].map((k) => document.getElementById('n-' + k).textContent)),
     ['1', '1', '1', '1', '2'], 'recomptes: una de cada color i dos locals buits');

  seccio('Filtres');
  await p.uncheck('#f-vermell');
  eq((await vista(iCadis)).color, 'groc', 'sense les precintades, Cadis 19 passa a groc (1403)');
  await p.uncheck('#f-groc');
  eq((await vista(iCadis)).alMapa, false, 'i sense requeriments, ja no es pinta');
  await p.check('#f-vermell'); await p.check('#f-groc');
  await p.check('#f-buits');
  eq(await vista(iBuit), { color: 'gris', acts: [], revisar: false, alMapa: true }, 'locals buits: el sol es pinta en gris');
  eq(await p.evaluate((i) => capes[i].forma.options.fillColor, iBuit), '#b0b0b0', 'gris de debò');
  await p.uncheck('#f-buits');
  await p.selectOption('#f-revisar', 'mb');
  eq([(await vista(iCadis)).alMapa, (await vista(iHosp)).alMapa], [false, true], 'per revisar «local marcat com a buit»: només Hospitalet');
  await p.selectOption('#f-revisar', '');

  seccio('Allotjaments turístics (CCAE 552/5520)');
  const actsHosp = async () => (await vista(iHosp)).acts;
  eq(await actsHosp(), ['9'], 'per defecte, l\'hotel (30) no surt');
  check((await p.textContent('#f-turisme')).includes('Només allotjaments turístics (1)'), 'el desplegable diu quants n\'hi ha');
  await p.selectOption('#f-turisme', 'nomes');
  eq(await actsHosp(), ['30'], '«Només allotjaments turístics»: només l\'hotel');
  eq((await vista(iCadis)).alMapa, false, 'i la resta de parcel·les, fora');
  await p.selectOption('#f-turisme', 'tot');
  eq(await actsHosp(), ['9', '30'], '«Totes»: les dues');
  await p.selectOption('#f-turisme', 'no');

  seccio('Classificació (annex)');
  const iII = await p.evaluate(() => CLASSIFS.indexOf('II'));
  await p.uncheck('#fc-' + iII);
  eq((await vista(iCadis)).acts, ['1447'], 'sense l\'annex II, el 1403 (II) desapareix');
  eq(await p.textContent('#nc-' + iII), '1', 'i la casella diu quantes n\'hi ha');
  await p.check('#fc-' + iII);

  seccio('Per revisar: coordenada fora de la parcel·la');
  await p.selectOption('#f-revisar', 'ne');
  eq((await vista(iCadis)).acts, ['1403'], 'coordenada fora: el 1403 (la seva UTM cau lluny de la parcel·la)');
  check((await p.textContent('#ajuda-revisar')).includes('vermell'), 'amb l\'explicació a sota');
  check((await p.textContent('#f-revisar')).includes('Coordenada UTM fora de la parcel·la o sense (ID en vermell) (1)'), 'i el recompte al desplegable');
  await p.selectOption('#f-revisar', '');

  seccio('Etiquetes: a la coordenada UTM i segons el zoom');
  const etiqueta = (i) => p.evaluate((i) => {
    const t = capes[i].etiqs.filter((e) => e.alMapa).map((e) => (e.def.v ? 'V:' : '') + e.tt.getElement().textContent.trim());
    return t.length ? t.join(' | ') : null;
  }, i);
  await p.evaluate((i) => map.setView(PARCELES[i].c, 15, { animate: false }), iCadis);
  eq(await etiqueta(iCadis), null, 'zoom 15: cap etiqueta');
  await p.evaluate((i) => map.setView(PARCELES[i].c, 16.5, { animate: false }), iCadis);
  eq(await etiqueta(iCadis), '1447 | V:1403', 'zoom 16,5: el 1447 a la seva coordenada i el 1403 (fora de la parcel·la) en vermell');
  // El 1447, a la seva coordenada UTM (a prop de la façana de baix), DINS de la
  // parcel·la, emmarcat amb el color del seu estat (vermell: precintada).
  const ent = await p.evaluate((i) => {
    const e = capes[i].etiqs.filter((x) => !x.def.v)[0];
    const ll = e.tt.getLatLng();
    const dins = capes[i].forma.getBounds().contains(ll) && ll.lat > PARCELES[i].c[0] - 0.001;
    return { dir: e.def.d, xip: e.tt.getElement().querySelector('.xip').className, dins: dins, sotaCentre: ll.lat < PARCELES[i].c[0] };
  }, iCadis);
  eq(ent, { dir: 't', xip: 'xip xip-vermell', dins: true, sotaCentre: true }, 'a l\'entrada (a baix), dins la parcel·la, creixent cap a dins, amb el marc del seu color');
  eq(await p.evaluate((i) => capes[i].etiqs.filter((x) => x.def.v)[0].tt.getElement().querySelector('.xip').className, iCadis),
     'xip xip-err', 'la no trobada: el text en vermell');
  // La LÍNIA DE PUNTS de l'etiqueta (UTM de l'Excel d'activitats) al punt de la
  // parcel·la al Cadastre.
  const lin = await p.evaluate((i) => {
    const e = capes[i].etiqs.filter((x) => !x.def.v)[0];
    return e.linia ? { alMapa: map.hasLayer(e.linia), dash: e.linia.options.dashArray, color: e.linia.options.color,
                       punts: e.linia.getLatLngs().map((q) => [q.lat, q.lng]), c: e.def.c, r: e.def.r } : null;
  }, iCadis);
  check(lin !== null && lin.alMapa && lin.dash === '1,4', 'una línia de punts surt de l\'etiqueta del 1447');
  eq(lin && lin.punts, lin && [lin.c, lin.r], '...i va fins al punt de la parcel·la al Cadastre');
  eq(await p.evaluate((i) => capes[i].etiqs.filter((x) => x.def.v)[0].linia, iCadis), null, 'la vermella del centre no en porta');
  await p.evaluate((i) => map.setView(PARCELES[i].c, 18.5, { animate: false }), iCadis);
  eq(await etiqueta(iCadis), '1447 | V:1403', 'zoom 18,5: només els ID (el local/planta/porta tapava les naus: és a la fitxa)');
  eq(await p.evaluate(() => document.querySelectorAll('.leaflet-tooltip.ent .sub').length), 0, 'cap text de local damunt del plànol');
  // Dotze ID a la mateixa porta (Sant Ferran): de quatre en quatre, no una columna.
  eq(await p.evaluate(() => {
    const vis = {}; const g = [];
    for (let i = 1; i <= 12; i++) { g.push(String(i)); vis[String(i)] = { e: 'blau' }; }
    const d = document.createElement('div'); d.innerHTML = textEntrada({ g: g, v: 0 }, vis, 18.5, null);
    return [d.querySelectorAll('.xip').length, d.querySelectorAll('br').length, textEntrada({ g: g, v: 0 }, vis, 16.5, null).includes('+8')];
  }), [12, 2, true], '12 ID a la mateixa porta: tres files de quatre de prop, i 4 + «+8» a mig zoom');
  await p.uncheck('#f-etiq');
  eq(await etiqueta(iCadis), null, 'i es poden amagar');
  eq(await p.evaluate((i) => map.hasLayer(capes[i].etiqs.filter((x) => !x.def.v)[0].linia), iCadis), false, 'i la línia s\'amaga amb l\'etiqueta');
  await p.check('#f-etiq');

  seccio('El plànol del Cadastre (MapaFons.js, el mateix que Coordenades)');
  eq(await p.evaluate(() => [document.getElementById('f-cadastre').checked, map.hasLayer(FONS.cadastre)]), [false, false], 'apagat per defecte');
  await p.check('#f-cadastre');
  eq(await p.evaluate(() => [map.hasLayer(FONS.cadastre), document.querySelector('.leaflet-control-layers-overlays input').checked]), [true, true],
     'la casella del lateral encén la capa comuna, i el selector ho diu');
  await p.uncheck('#f-cadastre');
  eq(await p.evaluate(() => map.hasLayer(FONS.cadastre)), false, 'i l\'apaga');

  seccio('La fitxa');
  await p.evaluate((i) => obrePopup(i), iCadis);
  const fitxa = await p.textContent('.leaflet-popup-content');
  check(fitxa.includes('1447') && fitxa.includes('EL RACO') && fitxa.includes('PRECINTADA'), 'hi ha la precintada, amb el nom');
  check(fitxa.includes('(Cadastre)'), 'i diu que la planta/porta surt del Cadastre');
  check(fitxa.includes('Requeriment') && fitxa.includes('2 informes'), 'la de requeriment, amb el nombre d\'informes');
  check(fitxa.includes('1 local buit'), 'i el local buit (el duplicat del GIA no hi compta)');
  eq(await p.evaluate(() => [...document.querySelectorAll('.leaflet-popup-content td.gia')].map((t) => t.textContent)), ['1447', '1403'],
     'una fila per ACTIVITAT: el 1447, amb dos establiments, surt un sol cop');
  check(fitxa.includes('2 establiments:') && fitxa.includes('Local 7'), '...amb els seus dos establiments');
  check(fitxa.includes('també com a BUIT (establiment 8)'), 'i l\'avís del local buit duplicat al GIA');
  check((await p.innerHTML('.leaflet-popup-content')).includes('rc1=2295827&amp;rc2=DF2729E'), 'amb l\'enllaç a la fitxa del Cadastre');
  check(fitxa.includes('Esc. 1 - Pl. 2 - Pt. 16') && fitxa.includes('Local 5'), 'el local/planta/porta, a la fitxa (ja no al plànol)');
  check(fitxa.includes('Adreça (base d\'activitats): C CADIS 21'), 'l\'adreça de la base d\'activitats, amb el seu nom');
  check(fitxa.includes('Adreça (Cadastre): CL CADIS 19'), 'i la del Cadastre (el portal amb el seu número), diferenciada');
  check(fitxa.includes('la coordenada UTM cau fora de la parcel·la'), 'la que va en vermell diu per què');
  check(!fitxa.includes('situa'), 'res de situar a mà: mana la coordenada de l\'Excel');
  check(fitxa.includes('Adreces al Cadastre: CL CADIS 19'), 'i les de la parcel·la al Cadastre, a dalt');

  seccio('Clic a un ID: la fitxa d\'aquella activitat');
  await p.evaluate(() => map.closePopup());
  await p.evaluate((i) => map.setView(PARCELES[i].c, 18.5, { animate: false }), iCadis);
  eq(await p.evaluate((i) => capes[i].etiqs.filter((x) => !x.def.v)[0].tt.options.direction, iCadis), 'center',
     'l\'etiqueta, centrada al seu punt (on acaba la línia de punts)');
  await p.locator('.leaflet-tooltip.ent .xip[data-g="1447"]').click();
  const fAct = await p.locator('.leaflet-popup-content').last().textContent();
  check(fAct.includes('ID 1447') && fAct.includes('2 establiments:') && !fAct.includes('1403'), 'només el 1447, amb els seus establiments (cap altra activitat)');
  check(fAct.includes('Totes les activitats de 2295827DF2729E'), 'i un enllaç a tota la parcel·la');
  await p.locator('.leaflet-popup-content a', { hasText: 'Totes les activitats' }).last().click();
  await p.waitForFunction(() => { const c = [...document.querySelectorAll('.leaflet-popup-content')].pop(); return c && c.textContent.includes('1403'); }, null, { timeout: 3000 });
  check(true, 'que obre la fitxa de la parcel·la, amb totes');
  await p.evaluate(() => map.closePopup());
  await p.selectOption('#f-revisar', 'bd');
  eq((await vista(iCadis)).acts, ['1447'], 'per revisar «local buit duplicat al GIA»: el 1447');
  await p.selectOption('#f-revisar', '');

  seccio('Cercador');
  await p.fill('#cerca', '144');
  eq(await p.locator('#resultats div').count(), 1, 'per ID GIA (començament): 1447');
  await p.locator('#resultats div').first().click();
  check((await p.textContent('.leaflet-popup-content')).includes('EL RACO'), 'i en clicar-hi, s\'obre la fitxa');
  await p.fill('#cerca', 'hospitalet');
  eq(await p.locator('#resultats div').count(), 2, 'per adreça (també l\'hotel, que el filtre amaga)');
  await p.uncheck('#f-verd');
  await p.fill('#cerca', 'acme');
  check((await p.textContent('#resultats')).includes('amagada pel filtre'), 'troba també el que el filtre amaga, i ho diu');
  await p.check('#f-verd');

  seccio('Baixar el CSV');
  const [baixada] = await Promise.all([p.waitForEvent('download'), p.click('text=Baixar la llista')]);
  const csv = fs.readFileSync(await baixada.path(), 'utf8');
  check(csv.charCodeAt(0) === 0xfeff, 'amb BOM (l\'Excel l\'obre amb els accents bé)');
  const linies = csv.replace(/^﻿/, '').trim().split('\r\n');
  eq(linies.length, 5, 'capçalera + les quatre activitats que es veuen');
  check(linies[0].startsWith('ID GIA;Nom;Activitat'), 'separat per ;');

  seccio('El fons del mapa (MapaFons.js)');
  // OpenStreetMap rebutja les pagines obertes des del disc ("Access blocked"):
  // cap peticio hi ha d'anar.
  eq(peticions.filter((u) => /openstreetmap\.org/.test(u)).length, 0, 'cap rajola es demana a OpenStreetMap');
  check(peticions.some((u) => /geoserveis\.icgc\.cat/.test(u)), 'el primer fons es el de l\'ICGC');
  // Aqui no respon cap fons (totes les rajoles s'avorten): es diu, i el mapa
  // segueix funcionant.
  await p.waitForFunction(() => document.querySelector('.fons-avis') && /Cap fons/.test(document.querySelector('.fons-avis').textContent), null, { timeout: 15000 });
  check(true, 'si cap fons no respon, ho diu (i les dades hi són igual)');

  // Si l'ICGC no respon i Esri si, passa SOL a Esri.
  const ctx2 = await nav.newContext({ viewport: { width: 1280, height: 800 } });
  await serveixLeaflet(ctx2);
  const serveix = (route) => route.fulfill({ body: PNG_1x1, contentType: 'image/png' });
  await ctx2.route(/arcgisonline\.com/, serveix);
  const p2 = await ctx2.newPage();
  p2.on('pageerror', (e) => errors.push(String(e)));
  await p2.goto(pathToFileURL(path.join(TMP, 'Planol.html')).href);
  await p2.waitForFunction(() => typeof FONS !== 'undefined' && FONS.actiu() === 'Mapa (Esri)', null, { timeout: 15000 });
  check((await p2.textContent('.fons-avis')).includes('ara es veu «Mapa (Esri)»'), 'l\'ICGC no respon: passa sol a Esri, i ho diu');
  eq(await p2.evaluate(() => FONS_MAPA.map((f) => f.nom)), ['Mapa (ICGC)', 'Ortofoto (ICGC)', 'Mapa (Esri)'],
     'CARTO ja no hi és (demana clau: pinta «API KEY REQUIRED» i el recanvi no ho veu)');
  // Ara l'ICGC respon: l'usuari tria l'ortofoto amb el selector.
  await ctx2.route(/geoserveis\.icgc\.cat/, serveix);
  const iCadis2 = await p2.evaluate(() => PARCELES.findIndex((x) => x.k === '2295827DF2729E'));
  const vora = () => p2.evaluate((i) => [capes[i].forma.options.color, capes[i].forma.options.weight], iCadis2);
  const voraMapa = await vora();
  await p2.hover('.leaflet-control-layers');
  await p2.locator('.leaflet-control-layers label', { hasText: 'Ortofoto (ICGC)' }).click();
  eq(await p2.evaluate(() => FONS.actiu()), 'Ortofoto (ICGC)', 'el selector canvia el fons');
  eq(await vora(), ['#ffffff', 2.5], 'sobre l\'ortofoto, la vora de les parcel·les és blanca i més gruixuda (es distingeix de la foto)');
  check(voraMapa[0] !== '#ffffff', 'i sobre el mapa, la de sempre');
  await p2.reload();
  await p2.waitForFunction(() => typeof FONS !== 'undefined', null, { timeout: 10000 });
  eq(await p2.evaluate(() => FONS.actiu()), 'Ortofoto (ICGC)', 'en tornar a obrir el mapa, es recorda');
  eq(await vora(), ['#ffffff', 2.5], 'i ja surt amb la vora blanca');
  // Si el que ha triat l'usuari deixa de respondre, tampoc es queda sense fons.
  await ctx2.unroute(/geoserveis\.icgc\.cat/);
  await p2.reload();
  await p2.waitForFunction(() => typeof FONS !== 'undefined' && FONS.actiu() === 'Mapa (Esri)', null, { timeout: 15000 });
  check(true, 'el fons triat no respon: en torna a posar un que sí');
  eq((await vora())[0] !== '#ffffff', true, 'i la vora torna a ser la de sempre');
  await ctx2.close();

  seccio('Errors de JavaScript');
  eq(errors, [], 'cap error de JavaScript');
} catch (e) {
  comptes.fail++;
  console.log('  FAIL la prova ha petat: ' + (e && e.stack || e));
} finally {
  await nav.close();
  fs.rmSync(TMP, { recursive: true, force: true });
}
process.exit(resultat('PLANOL'));
