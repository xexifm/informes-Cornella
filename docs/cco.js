// La Còpia Oculta (CCO) del correu de l'informe, al pas final del mòbil.
// -----------------------------------------------------------------------------
// Les adreces NO són al GitHub (que és públic): les hi deixa el PC al Drive
// privat (Dades/correu-mobil.json; Configuració → Configuració correu
// electrònic → «Informe de visita · des del mòbil»). Cada una porta 'def':
//   - CCO SEMPRE (def = true): surten MARCADES; es poden desmarcar.
//   - CCO A TRIAR (def = false): surten DESMARCADES.
//
// Abans vivia a app.js i, quan el Drive no responia o encara no hi havia el
// fitxer, el quadre sortia BUIT sense dir res (l'usuari, octubre 2026: «no em
// deixa escollir el CCO»). Ara sempre diu què passa i es pot tornar a provar.
//
// Exposa window.Cco amb: muntar(cont, delJson), seleccionat(cont).

(function () {
  var estat = { carregant: false, fet: false };

  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }

  // Les adreces d'un grup: un títol i una casella per adreça.
  function grup(cont, titol, llista, marcades) {
    if (!llista.length) return;
    cont.appendChild(el("div", "bcc-grup", titol));
    llista.forEach(function (o) {
      var lab = el("label", "bcc-item");
      var cb = document.createElement("input");
      cb.type = "checkbox";
      cb.value = o.addr;
      cb.checked = marcades;
      lab.appendChild(cb);
      lab.appendChild(document.createTextNode(" " + o.addr));
      cont.appendChild(lab);
    });
  }

  // La llista del fitxer, tal com vingui: {bcc:[{addr,def}]} és el normal, però
  // es tolera el que podria escriure un PowerShell 5.1 ({value:[...]}, un sol
  // objecte, Addr/Default amb majúscula). Torna [{addr, def}].
  function llistaDe(o) {
    var l = o && (o.bcc || o.Bcc || o.BCC);
    if (l && !Array.isArray(l) && Array.isArray(l.value)) l = l.value;
    if (l && !Array.isArray(l)) l = [l];
    return (l || []).map(function (e) {
      if (typeof e === "string") return { addr: e, def: false };
      return e ? { addr: String(e.addr || e.Addr || "").trim(), def: !!(e.def != null ? e.def : e.Default) } : null;
    }).filter(function (e) { return e && e.addr; });
  }

  // "2026-10-20T09:15:00.000Z" -> "20/10/2026 11:15". PURA.
  function quan(iso) {
    var d = new Date(iso);
    if (!iso || isNaN(d.getTime())) return "";
    var z = function (n) { return (n < 10 ? "0" : "") + n; };
    return z(d.getDate()) + "/" + z(d.getMonth() + 1) + "/" + d.getFullYear() + " " + z(d.getHours()) + ":" + z(d.getMinutes());
  }

  // Pinta la llista ({addr, def}) o, si no n'hi ha, el missatge.
  function pintar(cont, llista, avis) {
    cont.innerHTML = "";
    var bones = (Array.isArray(llista) ? llista : []).filter(function (o) { return o && o.addr; });
    grup(cont, "CCO sempre (desmarca-la si no la vols):", bones.filter(function (o) { return !!o.def; }), true);
    grup(cont, "CCO a triar:", bones.filter(function (o) { return !o.def; }), false);
    if (!bones.length) {
      cont.appendChild(el("p", "ajuda", avis || "No hi ha cap adreça de CCO per al mòbil. Al PC: Configuració → Configuració correu electrònic → «Informe de visita · des del mòbil»."));
    }
    estat.fet = bones.length > 0;
  }

  // Un missatge d'error amb un enllaç per tornar-ho a provar.
  function error(cont, text, delJson) {
    pintar(cont, delJson, text);
    var a = el("a", "bcc-reintent", "Tornar a carregar la CCO");
    a.href = "#";
    a.addEventListener("click", function (ev) { ev.preventDefault(); estat.fet = false; muntar(cont, delJson, true); });
    cont.appendChild(a);
  }

  // Carrega la CCO del Drive i la pinta. Un sol cop per sessió si ha anat bé;
  // si no (sense Drive o sense fitxer), es torna a provar la propera vegada.
  function muntar(cont, delJson, forca) {
    if (!cont || estat.carregant || (estat.fet && !forca)) return;
    var dj = Array.isArray(delJson) ? delJson : [];
    var D = window.Drive;
    if (!D || !D.llegirCorreuMobil) { pintar(cont, dj); return; }
    estat.carregant = true;
    cont.innerHTML = "";
    cont.appendChild(el("p", "ajuda", "Carregant la CCO del Drive…"));
    var connexio = (D.connectat && D.connectat()) ? Promise.resolve(true)
      : (D.reconnectarSilenci ? D.reconnectarSilenci() : Promise.resolve(false));
    connexio.then(function (ok) {
      if (!ok) throw new Error("Connecta el Drive (a la pantalla d'inici) per veure la CCO.");
      return D.llegirCorreuMobil();
    }).then(function (o) {
      var l = llistaDe(o);
      var q = quan(o && o.__modificat);
      pintar(cont, l.length ? l : dj, "El fitxer de la CCO del Drive (correu-mobil.json" + (q ? ", actualitzat el " + q : "") +
        ") no porta cap adreça. Al PC: Configuració → Configuració correu electrònic → Desar (la torna a pujar al Drive).");
    }).catch(function (e) {
      error(cont, "No s'ha pogut llegir la CCO: " + (e && e.message ? e.message : e), dj);
    }).then(function () { estat.carregant = false; });
  }

  // Les adreces marcades, unides per coma (buit si cap).
  function seleccionat(cont) {
    if (!cont) return "";
    var out = [];
    var cbs = cont.querySelectorAll("input[type=checkbox]");
    for (var i = 0; i < cbs.length; i++) { if (cbs[i].checked) out.push(cbs[i].value); }
    return out.join(",");
  }

  window.Cco = { muntar: muntar, seleccionat: seleccionat, llistaDe: llistaDe };
})();
