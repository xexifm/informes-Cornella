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
  const vista = (i) => p.evaluate((i) => { const v = capes[i].vista; return { color: v.color, acts: v.acts.map((e) => e.g), revisar: v.revisar, alMapa: capes[i].alMapa }; }, i);

  seccio('Arrencada');
  eq(await p.evaluate(() => typeof L), 'object', 'el Leaflet carrega (amb el SRI)');
  eq(await p.evaluate(() => PARCELES.length), 4, 'quatre parcel·les');
  check((await p.textContent('#avisos')).includes('</script> amb <b>'), 'els avisos surten escapats (cap etiqueta s\'interpreta)');
  eq(await p.evaluate((i) => capes[i].forma instanceof L.Polygon, iCadis), true, 'Cadis 19: un polígon (té geometria)');
  eq(await p.evaluate((i) => capes[i].forma instanceof L.CircleMarker, iHosp), true, 'Hospitalet 147: un punt (sense geometria)');

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

  seccio('Etiquetes segons el zoom');
  const etiqueta = (i) => p.evaluate((i) => capes[i].etiqAlMapa ? capes[i].etiq.getContent() : null, i);
  await p.evaluate((i) => map.setView(PARCELES[i].c, 15, { animate: false }), iCadis);
  eq(await etiqueta(iCadis), null, 'zoom 15: cap etiqueta');
  await p.evaluate((i) => map.setView(PARCELES[i].c, 16.5, { animate: false }), iCadis);
  eq(await etiqueta(iCadis), '1447, 1403', 'zoom 16,5: els ID GIA');
  await p.evaluate((i) => map.setView(PARCELES[i].c, 18.5, { animate: false }), iCadis);
  eq(await etiqueta(iCadis), '<b>1447</b> Esc. 1 - Pl. 2 - Pt. 16<br><b>1403</b> Local 5', 'zoom 18,5: cada ID amb el seu local/planta/porta');
  await p.uncheck('#f-etiq');
  eq(await etiqueta(iCadis), null, 'i es poden amagar');
  await p.check('#f-etiq');

  seccio('La fitxa');
  await p.evaluate((i) => obrePopup(i), iCadis);
  const fitxa = await p.textContent('.leaflet-popup-content');
  check(fitxa.includes('1447') && fitxa.includes('EL RACO') && fitxa.includes('PRECINTADA'), 'hi ha la precintada, amb el nom');
  check(fitxa.includes('(Cadastre)'), 'i diu que la planta/porta surt del Cadastre');
  check(fitxa.includes('Requeriment') && fitxa.includes('2 informes'), 'la de requeriment, amb el nombre d\'informes');
  check(fitxa.includes('1 local buit'), 'i el local buit');
  check((await p.innerHTML('.leaflet-popup-content')).includes('rc1=2295827&amp;rc2=DF2729E'), 'amb l\'enllaç a la fitxa del Cadastre');

  seccio('Cercador');
  await p.fill('#cerca', '144');
  eq(await p.locator('#resultats div').count(), 1, 'per ID GIA (començament): 1447');
  await p.locator('#resultats div').first().click();
  check((await p.textContent('.leaflet-popup-content')).includes('EL RACO'), 'i en clicar-hi, s\'obre la fitxa');
  await p.fill('#cerca', 'hospitalet');
  eq(await p.locator('#resultats div').count(), 1, 'per adreça');
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

  // Si l'ICGC no respon i CARTO si, passa SOL a CARTO.
  const ctx2 = await nav.newContext({ viewport: { width: 1280, height: 800 } });
  await serveixLeaflet(ctx2);
  // CARTO i Esri responen; l'ICGC no.
  await ctx2.route(/basemaps\.cartocdn\.com|arcgisonline\.com/, (route) => route.fulfill({ body: PNG_1x1, contentType: 'image/png' }));
  const p2 = await ctx2.newPage();
  p2.on('pageerror', (e) => errors.push(String(e)));
  await p2.goto(pathToFileURL(path.join(TMP, 'Planol.html')).href);
  await p2.waitForFunction(() => typeof FONS !== 'undefined' && FONS.actiu() === 'Mapa (CARTO)', null, { timeout: 15000 });
  check((await p2.textContent('.fons-avis')).includes('ara es veu «Mapa (CARTO)»'), 'l\'ICGC no respon: passa sol a CARTO, i ho diu');
  // La tria de l'usuari es recorda (el selector de dalt a la dreta).
  await p2.hover('.leaflet-control-layers');
  await p2.locator('.leaflet-control-layers label', { hasText: 'Mapa (Esri)' }).click();
  eq(await p2.evaluate(() => FONS.actiu()), 'Mapa (Esri)', 'el selector canvia el fons');
  await p2.reload();
  await p2.waitForFunction(() => typeof FONS !== 'undefined', null, { timeout: 10000 });
  eq(await p2.evaluate(() => FONS.actiu()), 'Mapa (Esri)', 'i en tornar a obrir el mapa, es recorda');
  // Si el que ha triat l'usuari deixa de respondre, tampoc es queda sense fons.
  await ctx2.unroute(/basemaps\.cartocdn\.com|arcgisonline\.com/);
  await ctx2.route(/basemaps\.cartocdn\.com/, (route) => route.fulfill({ body: PNG_1x1, contentType: 'image/png' }));
  await p2.reload();
  await p2.waitForFunction(() => typeof FONS !== 'undefined' && FONS.actiu() === 'Mapa (CARTO)', null, { timeout: 15000 });
  check(true, 'el fons triat no respon: en torna a posar un que sí');
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
