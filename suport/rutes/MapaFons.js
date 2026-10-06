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
// CARTO NO HI ES (octubre 2026): les seves rajoles ara demanen una clau i
// carreguen igualment, amb "API KEY REQUIRED" pintat a sobre. Com que carreguen,
// el recanvi automatic no ho veu: un fons que respon amb una imatge d'error es
// pitjor que un que no respon. Abans d'afegir un fons nou, mira que no en demani.
//
// Quan canvia el fons (el tria l'usuari o el recanvi), el mapa rep l'event
// 'fonscanviat' amb { nom }: el Planol hi canvia el contorn de les parcel.les
// (sobre l'ortofoto, la linia fosca no es veu).
//
// EL PLANOL DEL CADASTRE (WMS oficial: totes les parcel.les i els edificis, tambe
// els que no tenen cap activitat) es una capa A SOBRE del fons, que s'encen al
// mateix selector. Era nomes del Planol activitats; l'usuari la va voler tambe a
// Coordenades ("posa'm el planol del cadastre a Coordenades tambe") i per aixo
// viu aqui, un sol cop. Apagada per defecte (carrega); si l'encens, es recorda.
//
// ASCII pur (els accents, amb \u).

var FONS_MAPA = [
  { nom: 'Mapa (ICGC)',
    url: 'https://geoserveis.icgc.cat/servei/catalunya/mapa-base/wmts/topografic/MON3857NW/{z}/{x}/{y}.png',
    opt: { maxNativeZoom: 18, attribution: '&copy; <a href="https://www.icgc.cat">Institut Cartogr\u00e0fic i Geol\u00f2gic de Catalunya</a>' } },
  { nom: 'Ortofoto (ICGC)',
    url: 'https://geoserveis.icgc.cat/servei/catalunya/mapa-base/wmts/orto/MON3857NW/{z}/{x}/{y}.png',
    opt: { maxNativeZoom: 18, attribution: '&copy; <a href="https://www.icgc.cat">Institut Cartogr\u00e0fic i Geol\u00f2gic de Catalunya</a>' } },
  { nom: 'Mapa (Esri)',
    url: 'https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}',
    opt: { maxNativeZoom: 18, attribution: 'Tiles &copy; Esri' } }
];

// Quantes rajoles han de fallar, SENSE CAP QUE CARREGUI, per donar un fons per
// mort. Una de sola pot fallar per atzar; aixo ja no.
var FONS_ERRORS_MAX = 4;
var FONS_CLAU = 'informesCornella.fonsMapa';

var CADASTRE_WMS = {
  nom: 'Pl\u00e0nol del Cadastre',
  url: 'https://ovc.catastro.meh.es/Cartografia/WMS/ServidorWMS.aspx',
  // Un WMS es dibuixa a qualsevol escala: arriba al zoom maxim sense ampliar.
  opt: { layers: 'Catastro', format: 'image/png', transparent: true, maxZoom: 22, opacity: 0.6,
         attribution: 'Direcci\u00f3n General del Catastro' }
};
var CADASTRE_CLAU = 'informesCornella.planolCadastre';

// Posa el fons al mapa (el recordat o el primer) i el selector, amb el planol del
// Cadastre a sobre si s'havia deixat ences. Torna { capes: {nom: capa}, actiu:
// function () -> nom, cadastre: la capa del Cadastre } (el Planol hi lliga la
// seva casella; les proves ho miren).
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
    map.fire('fonscanviat', { nom: nom });
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

  // El planol del Cadastre: es recorda TANT si l'encens al selector com des
  // d'una casella de la pagina (per aixo escolta la capa i no el selector).
  var cadastre = L.tileLayer.wms(CADASTRE_WMS.url, CADASTRE_WMS.opt);
  var sobre = {}; sobre[CADASTRE_WMS.nom] = cadastre;
  cadastre.on('add', function () { try { window.localStorage.setItem(CADASTRE_CLAU, '1'); } catch (e) { } });
  cadastre.on('remove', function () { try { window.localStorage.setItem(CADASTRE_CLAU, '0'); } catch (e) { } });
  var cadastreEnces = false;
  try { cadastreEnces = window.localStorage.getItem(CADASTRE_CLAU) === '1'; } catch (e) { cadastreEnces = false; }
  if (cadastreEnces) { cadastre.addTo(map); }

  // El selector. Nomes la tria de l'USUARI es recorda (baselayerchange no salta
  // quan el canvi el fa passaAlSeguent).
  L.control.layers(capes, sobre, { position: 'topright', collapsed: true }).addTo(map);
  map.on('baselayerchange', function (e) {
    actiu = e.name;
    map.fire('fonscanviat', { nom: e.name });
    try { window.localStorage.setItem(FONS_CLAU, e.name); } catch (er) { }
  });
  return { capes: capes, actiu: function () { return actiu; }, cadastre: cadastre };
}

// Una casella de la pagina que encen i apaga el planol del Cadastre, sempre
// d'acord amb el selector de dalt (la capa es la mateixa). La fan servir el
// Planol activitats i Coordenades.
function lligaCasellaCadastre(map, fons, idCasella) {
  var c = document.getElementById(idCasella);
  if (!c) { return; }
  c.checked = map.hasLayer(fons.cadastre);
  c.addEventListener('change', function () {
    if (c.checked) { fons.cadastre.addTo(map); } else { map.removeLayer(fons.cadastre); }
  });
  fons.cadastre.on('add remove', function () { c.checked = map.hasLayer(fons.cadastre); });
}
