// comu.mjs - El que comparteixen les proves dels mapes al navegador
// (prova-mapa-coordenades.mjs, prova-planol.mjs): els comptadors d'OK/FAIL i
// servir el Leaflet des de node_modules en lloc del CDN.
//
// El Leaflet de node_modules es el MATEIX paquet 1.9.4 que serveixen unpkg i
// jsDelivr, byte a byte: el SRI dels <script> hi coincideix i no cal treure'l.
// Les tessel.les (OpenStreetMap, el WMS del Cadastre) s'avorten: el mapa
// funciona igual sense fons.
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export const AQUI = path.dirname(fileURLToPath(import.meta.url));
export const LEAFLET = path.join(AQUI, 'node_modules', 'leaflet', 'dist');

export const comptes = { ok: 0, fail: 0 };
export function check(cond, msg) {
  if (cond) { comptes.ok++; console.log('  OK   ' + msg); } else { comptes.fail++; console.log('  FAIL ' + msg); }
}
export function eq(obtingut, esperat, msg) {
  const igual = JSON.stringify(obtingut) === JSON.stringify(esperat);
  check(igual, msg + (igual ? '' : ` (esperat ${JSON.stringify(esperat)}, obtingut ${JSON.stringify(obtingut)})`));
}
export function seccio(t) { console.log('\n--- ' + t + ' ---'); }

// Quins CDN responen: per defecte, unpkg (el de la pagina). Amb { unpkg: false }
// es prova el segon intent (jsDelivr) i amb tots dos a false, el missatge.
export async function serveixLeaflet(ctx, cdn = { unpkg: true, jsdelivr: true }) {
  const serveix = (actiu) => (route) => {
    if (!actiu) { route.abort(); return; }
    // El cami DINS de dist/ (les icones son a dist/images/: la del selector de
    // capes del fons, per exemple).
    const rel = new URL(route.request().url()).pathname.split('/dist/')[1] || '';
    const tipus = rel.endsWith('.css') ? 'text/css' : rel.endsWith('.png') ? 'image/png' : 'application/javascript';
    route.fulfill({
      path: path.join(LEAFLET, ...rel.split('/')),
      headers: { 'Access-Control-Allow-Origin': '*', 'Content-Type': tipus },
    });
  };
  await ctx.route('https://unpkg.com/leaflet@1.9.4/dist/**', serveix(cdn.unpkg));
  await ctx.route('https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/**', serveix(cdn.jsdelivr));
  // Les rajoles de TOTS els fons (MapaFons.js) i del Cadastre. Cap prova surt a
  // Internet; la del fons serveix ella mateixa les que vol veure carregar.
  await ctx.route(TESSELES, (route) => route.abort());
}

export const TESSELES = /tile\.openstreetmap\.org|ovc\.catastro\.meh\.es|geoserveis\.icgc\.cat|arcgisonline\.com/;

// Un PNG de 1x1, per servir una rajola que "carrega".
export const PNG_1x1 = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');

export function resultat(nom) {
  console.log('\n========================================');
  console.log(`RESULTAT ${nom}: ${comptes.ok} OK, ${comptes.fail} FAIL`);
  console.log('========================================');
  return comptes.fail === 0 ? 0 : 1;
}
