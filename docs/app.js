// Formulari del mòbil per preparar informes Cornellà.
// -----------------------------------------------------------------------------
// Llegeix les plantilles (dades/*.json servides per GitHub Pages), guia l'usuari
// pels passos, i al final: (a) ofereix enviar els requeriments per correu, i
// (b) prepara un "paquet" JSON que el PC convertirà en el .docx complet.
//
// La lògica de camps/claus/selecció replica EXACTAMENT la del PC
// (GenerarInforme.ps1) perquè el paquet sigui interpretable sense sorpreses.

(function () {
  "use strict";

  // ------- Utilitats que repliquen el PC -------------------------------------

  // PowerShell ConvertTo-Json (5.1) pot convertir un array d'1 element en
  // escalar. Normalitzem sempre a array.
  function asArray(x) {
    if (x === null || x === undefined) return [];
    return Array.isArray(x) ? x : [x];
  }

  // _ItemKey: "Secció::Item[::Fill]"
  function itemKey(sec, item, child) {
    return child ? (sec + "::" + item + "::" + child) : (sec + "::" + item);
  }

  var RE_CAMP = /\[CAMP:\s*([^\]]+?)\s*\]/g;
  var RE_OPCIO = /\[OPCIO:\s*([^\]]+?)\s*\]/g;

  // _ParseOpcio: "nom | A | B" -> {name, options:[A,B]}
  function parseOpcio(raw) {
    var segs = raw.split("|");
    var name = segs[0].trim();
    var opts = [];
    for (var i = 1; i < segs.length; i++) {
      var o = segs[i].trim();
      if (o !== "") opts.push(o);
    }
    return { name: name, options: opts };
  }

  // _AddFieldsFromText: afegeix camps detectats (sense duplicar) a fields/order.
  function addFieldsFromText(fields, order, text) {
    if (!text) return;
    var m;
    RE_CAMP.lastIndex = 0;
    while ((m = RE_CAMP.exec(text)) !== null) {
      var raw = m[1].trim();
      var name = raw, hint = "";
      var p = raw.indexOf("(");
      if (p >= 0) { name = raw.substring(0, p).trim(); hint = raw.substring(p).trim().replace(/^\(/, "").replace(/\)$/, ""); }
      if (!fields[name]) { fields[name] = { name: name, type: "text", hint: hint, options: [], value: "" }; order.push(name); }
    }
    RE_OPCIO.lastIndex = 0;
    while ((m = RE_OPCIO.exec(text)) !== null) {
      var parsed = parseOpcio(m[1]);
      if (!parsed.name) continue;
      if (!fields[parsed.name]) {
        fields[parsed.name] = {
          name: parsed.name, type: "choice", hint: "",
          options: parsed.options, value: parsed.options.length ? parsed.options[0] : ""
        };
        order.push(parsed.name);
      }
    }
  }

  // Apply-Fields: substitueix [OPCIO:] i [CAMP:] pels valors.
  function applyFields(text, values) {
    if (!text) return "";
    var out = text.replace(RE_OPCIO, function (_, g) {
      var name = parseOpcio(g).name;
      return (values && values[name] !== undefined) ? String(values[name]) : "";
    });
    out = out.replace(RE_CAMP, function (_, g) {
      var raw = g.trim(), name = raw;
      var p = raw.indexOf("(");
      if (p >= 0) name = raw.substring(0, p).trim();
      return (values && values[name] !== undefined) ? String(values[name]) : "";
    });
    return out;
  }

  // _SplitTextAndUrls: separa el text dels enllaços d'una línia.
  function splitTextAndUrls(line) {
    if (!line) return { text: "", urls: [] };
    if (line.indexOf("[[URL]] ") === 0) return { text: "", urls: [line.substring(8).trim()] };
    var tokens = line.split(/\s+/);
    var urls = tokens.filter(function (t) { return /^https?:\/\//i.test(t); });
    if (!urls.length) return { text: line, urls: [] };
    var idx = line.search(/https?:\/\//i);
    var text = (idx >= 0 ? line.substring(0, idx) : line).trim();
    return { text: text, urls: urls };
  }

  // Build-SelectionFromKeys: reconstrueix la selecció (seccions amb items).
  function reconstructSelection(catalog, keysSet) {
    var result = [];
    asArray(catalog.Sections).forEach(function (sec) {
      var chosen = [];
      asArray(sec.Items).forEach(function (el) {
        if (el.Kind === "subsection" || el.Kind === "intro") {
          chosen.push({ Kind: el.Kind, Short: el.Short, BodyLines: asArray(el.BodyLines), Children: [], Selected: false });
          return;
        }
        var isSel = keysSet.has(itemKey(sec.Title, el.Short));
        var chosenChildren = [];
        asArray(el.Children).forEach(function (ch) {
          if (keysSet.has(itemKey(sec.Title, el.Short, ch.Short))) {
            chosenChildren.push({ Kind: ch.Kind, Short: ch.Short, BodyLines: asArray(ch.BodyLines) });
          }
        });
        if (isSel || chosenChildren.length > 0) {
          chosen.push({ Kind: "item", Short: el.Short, BodyLines: asArray(el.BodyLines), Children: chosenChildren, Selected: isSel });
        }
      });
      if (chosen.some(function (x) { return x.Kind === "item"; })) {
        result.push({ Title: sec.Title, Items: chosen });
      }
    });
    return result;
  }

  // Get-FieldsFromSelection + Add-FieldsFromConclusions.
  function collectFields(selSections, conclusions, always) {
    var fields = {}, order = [];
    selSections.forEach(function (sec) {
      sec.Items.forEach(function (it) {
        var allText = it.BodyLines.join(" ");
        (it.Children || []).forEach(function (ch) { allText += " " + ch.BodyLines.join(" "); });
        addFieldsFromText(fields, order, allText);
      });
    });
    conclusions.forEach(function (c) { addFieldsFromText(fields, order, c.Body); });
    always.forEach(function (a) { addFieldsFromText(fields, order, a); });
    return { fields: fields, order: order };
  }

  // Text dels requeriments per al correu (numerats globalment, com el document).
  function buildRequirementsText(selSections, values) {
    var lines = [];
    var n = 0;
    selSections.forEach(function (sec) {
      lines.push("");
      lines.push(sec.Title.toUpperCase());
      sec.Items.forEach(function (it) {
        if (it.Kind !== "item") {
          if (it.Kind === "subsection") lines.push("— " + it.Short + " —");
          return;
        }
        n++;
        var parts = [];
        var urls = [];
        it.BodyLines.forEach(function (ln) {
          var s = splitTextAndUrls(applyFields(ln, values));
          if (s.text) parts.push(s.text);
          urls = urls.concat(s.urls);
        });
        lines.push(n + ". " + parts.join(" "));
        urls.forEach(function (u) { lines.push("   " + u); });
        (it.Children || []).forEach(function (ch) {
          var cp = [];
          ch.BodyLines.forEach(function (ln) {
            var s = splitTextAndUrls(applyFields(ln, values));
            if (s.text) cp.push(s.text);
          });
          if (cp.length) lines.push("   - " + cp.join(" "));
        });
      });
    });
    return lines.join("\n").trim();
  }

  // ------- Estat de l'aplicació ----------------------------------------------

  var manifest = null, conclusions = null, capcalera = null;
  var catalegCache = {};       // baseName -> JSON del catàleg
  var estat = {
    cataleg: null,             // baseName
    header: {},                // ID_GIA, EXP_NUM, ...
    keys: new Set(),           // claus de deficiències seleccionades
    conclTitles: new Set(),    // títols de conclusions seleccionades
    fieldOrder: [],            // ordre dels camps del pas 5
    fieldDefs: {},             // name -> {type, options, ...}
    fieldValues: {}            // name -> valor
  };

  var PASSOS = ["cataleg", "capcalera", "deficiencies", "conclusions", "camps", "final"];
  var passActual = 0;

  var HEADER_KEYS = ["ID_GIA", "EXP_NUM", "TITULAR", "ADRECA", "ACTIVITAT", "NUM_ANOTACIO", "DATA_ANOTACIO"];

  // ------- DOM ----------------------------------------------------------------
  function $(id) { return document.getElementById(id); }
  function mostrar(el, si) { el.classList.toggle("ocult", !si); }

  // ------- Càrrega de dades ---------------------------------------------------
  function carregarJson(ruta) {
    return fetch(ruta, { cache: "no-cache" }).then(function (r) {
      if (!r.ok) throw new Error("No s'ha pogut carregar " + ruta + " (" + r.status + ")");
      return r.json();
    });
  }

  function inici() {
    Promise.all([
      carregarJson("dades/manifest.json"),
      carregarJson("dades/conclusions.json").catch(function () { return { HeaderText: "", Selectable: [], Always: [] }; }),
      carregarJson("dades/capcalera.json").catch(function () { return { Placeholders: HEADER_KEYS }; })
    ]).then(function (res) {
      manifest = res[0];
      conclusions = res[1];
      capcalera = res[2];
      $("carregant").classList.add("ocult");
      mostrar($("navegacio"), true);
      muntarCataleg();
      muntarCapcalera();
      muntarDrive();
      anarA(0);
    }).catch(function (e) {
      $("carregant").innerHTML = '<span class="error">Error carregant les dades: ' + e.message +
        "<br>Comprova que el PC hagi pujat les dades (Actualitzar.bat) i que GitHub Pages estigui actiu.</span>";
    });
  }

  // ------- Pas 1: catàleg -----------------------------------------------------
  function muntarCataleg() {
    var sel = $("sel-cataleg");
    sel.innerHTML = "";
    asArray(manifest.Catalegs).forEach(function (b) {
      var o = document.createElement("option");
      o.value = b; o.textContent = b;
      sel.appendChild(o);
    });
    estat.cataleg = asArray(manifest.Catalegs)[0] || null;
    sel.value = estat.cataleg;
    sel.addEventListener("change", function () { estat.cataleg = sel.value; });
  }

  // ------- Pas 2: capçalera ---------------------------------------------------
  function muntarCapcalera() {
    var cont = $("camps-capcalera");
    cont.innerHTML = "";
    var placeholders = asArray(capcalera.Placeholders).length ? asArray(capcalera.Placeholders) : HEADER_KEYS;
    placeholders.forEach(function (k) {
      if (k === "ID_GIA") return; // ja té el seu camp principal
      var lbl = document.createElement("label");
      lbl.textContent = k;
      var inp = document.createElement("input");
      inp.type = "text";
      inp.id = "hdr-" + k;
      inp.addEventListener("input", function () { estat.header[k] = inp.value; });
      cont.appendChild(lbl);
      cont.appendChild(inp);
    });

    $("in-gia").addEventListener("input", function () { estat.header.ID_GIA = $("in-gia").value.trim(); });
    $("btn-cercar").addEventListener("click", cercarActivitat);
  }

  function cercarActivitat() {
    var gia = $("in-gia").value.trim();
    var msg = $("msg-cerca");
    if (!gia) { msg.textContent = "Escriu un ID GIA."; return; }
    if (!window.Drive || !Drive.disponible()) {
      msg.textContent = "Drive no configurat: omple les dades a mà (o el PC les omplirà en generar).";
      return;
    }
    msg.textContent = "Connectant amb Drive…";
    var pas = Drive.connectat() ? Promise.resolve() : Drive.connectar();
    pas.then(function () {
      estatDrive(true);
      msg.textContent = "Cercant…";
      return Drive.llegirActivitats();
    }).then(function (data) {
      var byId = (data && data.ById) || {};
      var act = byId[gia];
      if (!act) { msg.textContent = "L'ID GIA " + gia + " no s'ha trobat a la base de dades."; return; }
      HEADER_KEYS.forEach(function (k) {
        if (k === "ID_GIA") return;
        if (act[k] !== undefined) {
          estat.header[k] = act[k];
          var inp = $("hdr-" + k);
          if (inp) inp.value = act[k];
        }
      });
      $("det-capcalera").open = true;
      msg.textContent = "Dades trobades (" + (data.Source || "") + ").";
    }).catch(function (e) {
      msg.innerHTML = '<span class="error">' + e.message + "</span>";
    });
  }

  // ------- Pas 3: deficiències ------------------------------------------------
  function carregarCataleg() {
    if (catalegCache[estat.cataleg]) return Promise.resolve(catalegCache[estat.cataleg]);
    return carregarJson("dades/cataleg-" + estat.cataleg + ".json").then(function (c) {
      catalegCache[estat.cataleg] = c;
      return c;
    });
  }

  function muntarDeficiencies(catalog) {
    var cont = $("arbre-def");
    cont.innerHTML = "";
    asArray(catalog.Sections).forEach(function (sec) {
      var hSec = document.createElement("div");
      hSec.className = "seccio";
      hSec.textContent = sec.Title;
      cont.appendChild(hSec);

      asArray(sec.Items).forEach(function (el) {
        if (el.Kind === "intro") return;
        if (el.Kind === "subsection") {
          var sub = document.createElement("div");
          sub.className = "subseccio";
          sub.textContent = el.Short;
          cont.appendChild(sub);
          return;
        }
        // item
        var key = itemKey(sec.Title, el.Short);
        var fila = filaCheckbox(key, el.Short, el, "item");
        cont.appendChild(fila.wrap);

        // fills
        asArray(el.Children).forEach(function (ch) {
          var ckey = itemKey(sec.Title, el.Short, ch.Short);
          var cf = filaCheckbox(ckey, ch.Short, ch, "fill");
          cont.appendChild(cf.wrap);
          // propagació (com el TreeView del PC): marcar/desmarcar l'item
          // arrossega els fills.
          fila.cb.addEventListener("change", function () {
            cf.cb.checked = fila.cb.checked;
            if (fila.cb.checked) estat.keys.add(ckey); else estat.keys.delete(ckey);
          });
        });
      });
    });
    $("filtre-def").value = "";
    $("filtre-def").oninput = function () { filtrarArbre(this.value); };
  }

  function filaCheckbox(key, etiqueta, el, classe) {
    var wrap = document.createElement("div");
    wrap.className = classe;
    wrap.dataset.text = (etiqueta || "").toLowerCase();
    var cb = document.createElement("input");
    cb.type = "checkbox";
    cb.id = "cb-" + key;
    cb.checked = estat.keys.has(key);
    cb.addEventListener("change", function () {
      if (cb.checked) estat.keys.add(key); else estat.keys.delete(key);
    });
    var lbl = document.createElement("label");
    lbl.setAttribute("for", cb.id);
    lbl.textContent = etiqueta;
    // previsualització breu (primera línia de text, sense aplicar camps)
    var prev = primeraLiniaText(el);
    if (prev) {
      var sp = document.createElement("span");
      sp.className = "preview";
      sp.textContent = prev;
      lbl.appendChild(sp);
    }
    wrap.appendChild(cb);
    wrap.appendChild(lbl);
    return { wrap: wrap, cb: cb };
  }

  function primeraLiniaText(el) {
    var lines = asArray(el.BodyLines);
    for (var i = 0; i < lines.length; i++) {
      var s = splitTextAndUrls(lines[i]);
      if (s.text) { return s.text.length > 90 ? s.text.substring(0, 90) + "…" : s.text; }
    }
    return "";
  }

  function filtrarArbre(needle) {
    needle = (needle || "").toLowerCase();
    var cont = $("arbre-def");
    [].forEach.call(cont.querySelectorAll(".item, .fill"), function (w) {
      var match = !needle || w.dataset.text.indexOf(needle) >= 0;
      mostrar(w, match);
    });
  }

  // ------- Pas 4: conclusions -------------------------------------------------
  function muntarConclusions() {
    var cont = $("llista-concl");
    cont.innerHTML = "";
    var sel = asArray(conclusions.Selectable);
    if (!sel.length) { cont.innerHTML = '<p class="ajuda">No hi ha conclusions triables.</p>'; return; }
    sel.forEach(function (c) {
      var wrap = document.createElement("div");
      wrap.className = "concl";
      var cb = document.createElement("input");
      cb.type = "checkbox";
      cb.id = "cc-" + c.Title;
      cb.checked = estat.conclTitles.has(c.Title);
      cb.addEventListener("change", function () {
        if (cb.checked) estat.conclTitles.add(c.Title); else estat.conclTitles.delete(c.Title);
      });
      var div = document.createElement("div");
      var lbl = document.createElement("label");
      lbl.setAttribute("for", cb.id);
      lbl.textContent = c.Title;
      var cos = document.createElement("div");
      cos.className = "cos";
      cos.textContent = c.Body;
      div.appendChild(lbl);
      div.appendChild(cos);
      wrap.appendChild(cb);
      wrap.appendChild(div);
      cont.appendChild(wrap);
    });
  }

  // ------- Pas 5: camps -------------------------------------------------------
  function muntarCamps() {
    return carregarCataleg().then(function (catalog) {
      var selSections = reconstructSelection(catalog, estat.keys);
      var conclSel = asArray(conclusions.Selectable).filter(function (c) { return estat.conclTitles.has(c.Title); });
      var always = asArray(conclusions.Always);
      var cf = collectFields(selSections, conclSel, always);

      // conserva valors ja introduïts
      cf.order.forEach(function (name) {
        if (estat.fieldValues[name] !== undefined) cf.fields[name].value = estat.fieldValues[name];
      });
      estat.fieldOrder = cf.order;
      estat.fieldDefs = cf.fields;

      var cont = $("llista-camps");
      cont.innerHTML = "";
      mostrar($("sense-camps"), cf.order.length === 0);
      cf.order.forEach(function (name) {
        var f = cf.fields[name];
        var wrap = document.createElement("div");
        wrap.className = "camp";
        var lbl = document.createElement("label");
        lbl.textContent = name;
        wrap.appendChild(lbl);
        if (f.hint) {
          var h = document.createElement("div");
          h.className = "ajuda";
          h.textContent = f.hint;
          wrap.appendChild(h);
        }
        var inp;
        if (f.type === "choice") {
          inp = document.createElement("select");
          f.options.forEach(function (o) {
            var op = document.createElement("option");
            op.value = o; op.textContent = o;
            inp.appendChild(op);
          });
          inp.value = f.value || (f.options[0] || "");
        } else {
          inp = document.createElement("input");
          inp.type = "text";
          inp.value = f.value || "";
        }
        inp.addEventListener("input", function () { estat.fieldValues[name] = inp.value; });
        inp.addEventListener("change", function () { estat.fieldValues[name] = inp.value; });
        estat.fieldValues[name] = inp.value; // valor inicial (default dels desplegables)
        wrap.appendChild(inp);
        cont.appendChild(wrap);
      });
    });
  }

  // ------- Pas 6: final -------------------------------------------------------
  function muntarFinal() {
    return carregarCataleg().then(function (catalog) {
      var selSections = reconstructSelection(catalog, estat.keys);
      var text = buildRequirementsText(selSections, estat.fieldValues);
      $("prev-requeriments").value = text;
      if (!$("in-destinatari").value && window.CONFIG) $("in-destinatari").value = CONFIG.EMAIL_DESTINATARI || "";
    });
  }

  function assumpte() {
    var prefix = (window.CONFIG && CONFIG.EMAIL_ASSUMPTE_PREFIX) || "Requeriments activitat GIA";
    return prefix + " " + (estat.header.ID_GIA || "");
  }

  function enviarEmail() {
    var dest = $("in-destinatari").value.trim();
    var cos = $("prev-requeriments").value;
    var url = "mailto:" + encodeURIComponent(dest) +
      "?subject=" + encodeURIComponent(assumpte()) +
      "&body=" + encodeURIComponent(cos);
    window.location.href = url;
  }

  function construirPaquet() {
    var header = {};
    HEADER_KEYS.forEach(function (k) { header[k] = estat.header[k] || ""; });
    return {
      Version: 2,
      Origen: "mobil",
      Timestamp: new Date().toISOString(),
      CatalegBaseName: estat.cataleg,
      Header: header,
      SelectedKeys: Array.from(estat.keys),
      FieldValues: estat.fieldValues,
      ConclusionTexts: Array.from(estat.conclTitles),
      Email: { Destinatari: $("in-destinatari").value.trim(), Assumpte: assumpte() }
    };
  }

  function nomPaquet() {
    var gia = (estat.header.ID_GIA || "sense-gia").replace(/[^\w-]/g, "");
    var ts = new Date().toISOString().replace(/[:.]/g, "-");
    return "paquet_GIA" + gia + "_" + ts + ".json";
  }

  function enviarAlPc() {
    var msg = $("msg-final");
    if (!estat.keys.size) { msg.innerHTML = '<span class="error">No has seleccionat cap deficiència.</span>'; return; }
    var paquet = construirPaquet();
    var nom = nomPaquet();
    if (window.Drive && Drive.disponible()) {
      msg.textContent = "Connectant amb Drive…";
      var pas = Drive.connectat() ? Promise.resolve() : Drive.connectar();
      pas.then(function () { estatDrive(true); msg.textContent = "Pujant el paquet…"; return Drive.pujarPaquet(nom, paquet); })
        .then(function () { msg.innerHTML = "✅ Paquet enviat. L'informe es generarà al PC."; })
        .catch(function (e) { msg.innerHTML = '<span class="error">' + e.message + "</span>"; });
    } else {
      baixarPaquet(paquet, nom);
      msg.textContent = "Drive no configurat: s'ha baixat el paquet. Deixa'l a la carpeta Entrada del Drive.";
    }
  }

  function baixarPaquet(paquet, nom) {
    paquet = paquet || construirPaquet();
    nom = nom || nomPaquet();
    var blob = new Blob([JSON.stringify(paquet, null, 2)], { type: "application/json" });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = nom;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  // ------- Drive (estat visual) ----------------------------------------------
  function estatDrive(on) {
    var el = $("estat-drive");
    el.textContent = on ? "Drive: connectat" : "Drive: no connectat";
    el.classList.toggle("estat-on", on);
    el.classList.toggle("estat-off", !on);
  }
  function muntarDrive() {
    if (!window.Drive || !Drive.disponible()) {
      $("estat-drive").textContent = "Drive: no configurat";
    }
    $("btn-email").addEventListener("click", enviarEmail);
    $("btn-enviar-pc").addEventListener("click", enviarAlPc);
    $("btn-baixar").addEventListener("click", function () { baixarPaquet(); });
  }

  // ------- Navegació ----------------------------------------------------------
  function passId(i) { return "pas-" + PASSOS[i]; }

  function anarA(i) {
    // salta el pas de catàleg si només n'hi ha un
    if (PASSOS[i] === "cataleg" && asArray(manifest.Catalegs).length <= 1) {
      return anarA(i + (i >= passActual ? 1 : -1));
    }
    passActual = i;
    PASSOS.forEach(function (p, idx) { mostrar($("pas-" + p), idx === i); });
    $("btn-enrere").disabled = (i === 0) || (i === 1 && asArray(manifest.Catalegs).length <= 1);
    $("btn-seguent").textContent = (i === PASSOS.length - 1) ? "Fet" : "Següent";
    $("pas-indicador").textContent = "Pas " + (i + 1) + " / " + PASSOS.length;
    mostrar($("navegacio"), PASSOS[i] !== "final" ? true : true);
  }

  function validarPas() {
    var p = PASSOS[passActual];
    if (p === "capcalera" && !$("in-gia").value.trim()) {
      $("msg-cerca").innerHTML = '<span class="error">Cal un ID GIA.</span>';
      return false;
    }
    if (p === "deficiencies" && estat.keys.size === 0) {
      alert("Selecciona almenys una deficiència.");
      return false;
    }
    return true;
  }

  function seguent() {
    if (!validarPas()) return;
    var seg = passActual + 1;
    if (seg >= PASSOS.length) { return; } // ja al final
    // accions de preparació en entrar a certs passos
    var entrar = PASSOS[seg];
    var prep = Promise.resolve();
    if (entrar === "deficiencies") prep = carregarCataleg().then(muntarDeficiencies);
    else if (entrar === "conclusions") { muntarConclusions(); }
    else if (entrar === "camps") prep = muntarCamps();
    else if (entrar === "final") prep = muntarFinal();
    prep.then(function () { anarA(seg); }).catch(function (e) {
      alert("Error: " + e.message);
    });
  }

  function enrere() {
    var ant = passActual - 1;
    if (ant < 0) return;
    var entrar = PASSOS[ant];
    var prep = Promise.resolve();
    if (entrar === "deficiencies") prep = carregarCataleg().then(muntarDeficiencies);
    else if (entrar === "conclusions") muntarConclusions();
    else if (entrar === "camps") prep = muntarCamps();
    prep.then(function () { anarA(ant); });
  }

  // ------- Arrencada ----------------------------------------------------------
  document.addEventListener("DOMContentLoaded", function () {
    $("btn-seguent").addEventListener("click", seguent);
    $("btn-enrere").addEventListener("click", enrere);
    inici();
  });
})();
