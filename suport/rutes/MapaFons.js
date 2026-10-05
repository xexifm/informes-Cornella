// MapaFons.js - EL FONS dels mapes de 'rutes/' (Planol activitats, Coordenades
// i Ruta). No es carrega com a fitxer: MapaHtml.ps1 (Get-MapaFonsJs) el posa
// TAL QUAL dins de l'HTML, perque els mapes s'obren des del disc i han de
// funcionar sols.
//
// PER QUE NO OPENSTREETMAP (octubre 2026). Els servidors d'OpenStreetMap
// demanen que cada peticio digui de quina web ve (la capcalera Referer). Un
// HTML obert des del disc (file://) no n'envia cap, i el que tornaven era una
// rajola "Access blocked" / "403" a cada quadre del mapa.
//
// L'ICGC (la cartografia oficial de Catalunya, sense clau ni Referer) va
// primer. I com que un servei pot canviar d'adreca o caure, SI UN FONS NO
// CARREGA CAP RAJOLA i ja n'han fallat unes quantes, es passa SOL al seguent
// de la llista i es diu a la pantalla. La tria de l'usuari (el selector de dalt
// a la dreta) es recorda per a tots els mapes.
//
// ASCII pur (els accents, amb \u).

var FONS_MAPA = [
  { nom: 'Mapa (ICGC)',
    url: 'https://geoserveis.icgc.cat/servei/catalunya/mapa-base/wmts/topografic/MON3857NW/{z}/{x}/{y}.png',
    opt: { maxNativeZoom: 18, attribution: '&copy; <a href="https://www.icgc.cat">Institut Cartogr\u00e0fic i Geol\u00f2gic de Catalunya</a>' } },
  { nom: 'Ortofoto (ICGC)',
    url: 'https://geoserveis.icgc.cat/servei/catalunya/mapa-base/wmts/orto/MON3857NW/{z}/{x}/{y}.png',
    opt: { maxNativeZoom: 18, attribution: '&copy; <a href="https://www.icgc.cat">Institut Cartogr\u00e0fic i Geol\u00f2gic de Catalunya</a>' } },
  { nom: 'Mapa (CARTO)',
    url: 'https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png',
    opt: { subdomains: 'abcd', maxNativeZoom: 19, attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> &copy; <a href="https://carto.com/attributions">CARTO</a>' } },
  { nom: 'Mapa (Esri)',
    url: 'https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}',
    opt: { maxNativeZoom: 18, attribution: 'Tiles &copy; Esri' } }
];

// Quantes rajoles han de fallar, SENSE CAP QUE CARREGUI, per donar un fons per
// mort. Una de sola pot fallar per atzar; aixo ja no.
var FONS_ERRORS_MAX = 4;
var FONS_CLAU = 'informesCornella.fonsMapa';

// Posa el fons al mapa (el recordat o el primer) i el selector. Torna l'estat,
// per a les proves: { capes: {nom: capa}, actiu: function () -> nom }.
function afegeixFonsMapa(map) {
  var capes = {}, ordre = [], fallats = {}, actiu = null;
  FONS_MAPA.forEach(function (f) {
    // maxZoom alt: si el mapa s'apropa mes que el que dona el servei, la
    // rajola s'amplia (maxNativeZoom) en lloc de desapareixer.
    var c = L.tileLayer(f.url, L.extend({ maxZoom: 22 }, f.opt));
    var vist = { ok: 0, ko: 0 };
    c.on('tileload', function () { vist.ok++; });
    c.on('tileerror', function () {
      vist.ko++;
      if (vist.ok === 0 && vist.ko >= FONS_ERRORS_MAX && actiu === f.nom) { passaAlSeguent(f.nom); }
    });
    capes[f.nom] = c; ordre.push(f.nom);
  });

  function posa(nom) {
    if (actiu && capes[actiu]) { map.removeLayer(capes[actiu]); }
    actiu = nom;
    capes[nom].addTo(map);
    capes[nom].bringToBack();
  }

  function passaAlSeguent(nom) {
    fallats[nom] = true;
    var seg = null;
    for (var i = 0; i < ordre.length; i++) { if (!fallats[ordre[i]]) { seg = ordre[i]; break; } }
    if (!seg) { avisa('Cap fons del mapa no respon (sense connexi\u00f3?). Les dades s\u00ed que hi s\u00f3n.'); return; }
    posa(seg);
    avisa('El fons \u00ab' + nom + '\u00bb no respon: ara es veu \u00ab' + seg + '\u00bb.');
  }

  var caixaAvis = null;
  function avisa(text) {
    if (!caixaAvis) {
      var Avis = L.Control.extend({ options: { position: 'bottomleft' },
        onAdd: function () { var d = L.DomUtil.create('div', 'fons-avis'); d.style.cssText = 'background:#fff8e1;border:1px solid #e0c060;padding:4px 8px;font:12px sans-serif;max-width:340px'; return d; } });
      caixaAvis = new Avis(); caixaAvis.addTo(map);
    }
    caixaAvis.getContainer().textContent = text;
  }

  var triat = null;
  try { triat = window.localStorage.getItem(FONS_CLAU); } catch (e) { triat = null; }
  posa(capes[triat] ? triat : ordre[0]);

  // El selector. Nomes la tria de l'USUARI es recorda (baselayerchange no salta
  // quan el canvi el fa passaAlSeguent).
  L.control.layers(capes, null, { position: 'topright', collapsed: true }).addTo(map);
  map.on('baselayerchange', function (e) {
    actiu = e.name;
    try { window.localStorage.setItem(FONS_CLAU, e.name); } catch (er) { }
  });
  return { capes: capes, actiu: function () { return actiu; } };
}
