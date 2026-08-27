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

  // Treu els marcadors de negreta/cursiva (en text pla del correu no es poden
  // representar, així que es mostren sense els símbols **...** i //...//).
  function stripMarkers(t) {
    if (!t) return "";
    return String(t).replace(/\*\*(.+?)\*\*/g, "$1").replace(/\/\/(.+?)\/\//g, "$1");
  }

  // ------- Camps inline (opcions/text dins del propi text) --------------------
  // Treu TOTS els marcadors ** i // (encara que un parell quedi partit per un
  // [OPCIO:]/[CAMP:]); només per a la previsualització en pantalla.
  function stripMarkersLoose(t) {
    if (!t) return "";
    return String(t).replace(/\*\*/g, "").replace(/\/\//g, "");
  }

  // Segmenta un text amb [OPCIO:]/[CAMP:] en trossos ordenats per renderitzar-lo
  // amb controls inline. Equival a _SegmentRichText del PC.
  function segmentRichText(text) {
    var segs = [];
    if (!text) return segs;
    var rx = /\[OPCIO:\s*([^\]]+?)\s*\]|\[CAMP:\s*([^\]]+?)\s*\]/g;
    var pos = 0, m;
    while ((m = rx.exec(text)) !== null) {
      if (m.index > pos) {
        var plain = stripMarkersLoose(text.substring(pos, m.index));
        if (plain) segs.push({ kind: "text", text: plain });
      }
      if (m[1] !== undefined) {
        var p = parseOpcio(m[1]);
        segs.push({ kind: "opcio", name: p.name, options: p.options });
      } else {
        var raw = m[2].trim(), name = raw, hint = "";
        var pi = raw.indexOf("(");
        if (pi >= 0) { name = raw.substring(0, pi).trim(); hint = raw.substring(pi).trim().replace(/^\(/, "").replace(/\)$/, ""); }
        segs.push({ kind: "camp", name: name, hint: hint });
      }
      pos = m.index + m[0].length;
    }
    if (pos < text.length) {
      var plain2 = stripMarkersLoose(text.substring(pos));
      if (plain2) segs.push({ kind: "text", text: plain2 });
    }
    return segs;
  }

  // Assegura que un camp tingui un valor per defecte a estat.fieldValues.
  function ensureFieldValue(name, type, options) {
    if (estat.fieldValues[name] === undefined) {
      estat.fieldValues[name] = (type === "choice" && options.length) ? options[0] : "";
    }
  }

  // Desa un valor de camp i sincronitza tots els controls amb el mateix nom
  // (un mateix camp pot sortir a diversos llocs alhora).
  function setFieldValue(name, value) {
    estat.fieldValues[name] = value;
    [].forEach.call(document.querySelectorAll("[data-fieldname]"), function (el) {
      if (el.dataset.fieldname === name && el.value !== value) el.value = value;
    });
  }

  // Text "ric" d'un element (item/fill): només la part de TEXT de cada BodyLine
  // (els URLs no s'editen aquí), unit amb espais.
  function richTextOf(el) {
    var parts = [];
    asArray(el.BodyLines).forEach(function (ln) {
      var s = splitTextAndUrls(ln);
      if (s.text) parts.push(s.text);
    });
    return parts.join(" ");
  }

  // Renderitza un text dins d'un contenidor barrejant text i controls inline
  // per als [OPCIO:]/[CAMP:]. Equival a _RenderRichInto del PC.
  function renderRich(container, text) {
    segmentRichText(text).forEach(function (seg) {
      if (seg.kind === "text") {
        container.appendChild(document.createTextNode(seg.text + " "));
        return;
      }
      if (!seg.name) return;
      if (seg.kind === "opcio") {
        ensureFieldValue(seg.name, "choice", seg.options);
        var sel = document.createElement("select");
        sel.className = "inline-camp";
        sel.dataset.fieldname = seg.name;
        seg.options.forEach(function (o) {
          var op = document.createElement("option");
          op.value = o; op.textContent = o;
          sel.appendChild(op);
        });
        sel.value = (estat.fieldValues[seg.name] !== undefined) ? estat.fieldValues[seg.name] : (seg.options[0] || "");
        estat.fieldValues[seg.name] = sel.value;
        sel.addEventListener("change", function () { setFieldValue(seg.name, sel.value); });
        container.appendChild(sel);
      } else {
        ensureFieldValue(seg.name, "text", []);
        var inp = document.createElement("input");
        inp.type = "text";
        inp.className = "inline-camp";
        inp.dataset.fieldname = seg.name;
        if (seg.hint) inp.placeholder = seg.hint;
        inp.value = estat.fieldValues[seg.name] || "";
        inp.addEventListener("input", function () { setFieldValue(seg.name, inp.value); });
        container.appendChild(inp);
      }
    });
  }

  // Llista de requeriments numerada, PORTADA FIDELMENT de _WriteCatalegBody del
  // PC (mateixa numeració global, subseccions emeses només quan les segueix un
  // ítem real, fills sagnats sense numerar, separació text/URL). Així el correu
  // coincideix amb l'informe generat.
  function buildRequirementsList(selSections, values) {
    var lines = [];
    var n = 0;
    var lastSection = null;
    selSections.forEach(function (sec) {
      // Secció / subsecció derivada del títol amb " - " (com fa el PC).
      var idx = sec.Title.indexOf(" - ");
      var secName, subFromTitle = null;
      if (idx >= 0) { secName = sec.Title.substring(0, idx).trim(); subFromTitle = sec.Title.substring(idx + 3).trim(); }
      else secName = sec.Title.trim();
      if (secName !== lastSection) {
        lines.push("");
        lines.push(secName.toUpperCase());
        lastSection = secName;
      }
      if (subFromTitle) lines.push(subFromTitle);

      var pendingSub = null, pendingIntro = null;
      sec.Items.forEach(function (el) {
        if (el.Kind === "subsection") { pendingSub = el; pendingIntro = null; return; }
        if (el.Kind === "intro") { pendingIntro = el; return; }

        var itemLines = asArray(el.BodyLines).map(function (l) { return applyFields(l, values); });
        var hasChildren = !!(el.Children && el.Children.length > 0);
        if (!(el.Selected || hasChildren)) return;   // res a emetre per a aquest ítem

        // Subsecció/intro pendents: només quan ve un ítem real.
        if (pendingSub) { lines.push("· " + stripMarkers(pendingSub.Short)); pendingSub = null; }
        if (pendingIntro) {
          asArray(pendingIntro.BodyLines).forEach(function (l) {
            var s = splitTextAndUrls(applyFields(l, values));
            if (s.text) lines.push(stripMarkers(s.text));
            s.urls.forEach(function (u) { lines.push("   " + u); });
          });
          pendingIntro = null;
        }

        var itemWritten = false;
        if (itemLines.length > 0) {
          n++;
          var p0 = splitTextAndUrls(itemLines[0]);
          lines.push(n + ". " + stripMarkers(p0.text));
          p0.urls.forEach(function (u) { lines.push("   " + u); });
          for (var i = 1; i < itemLines.length; i++) {
            var s = splitTextAndUrls(itemLines[i]);
            if (s.text) lines.push("   " + stripMarkers(s.text));
            s.urls.forEach(function (u) { lines.push("   " + u); });
          }
          itemWritten = true;
        }
        if (hasChildren) {
          el.Children.forEach(function (ch) {
            var childLines = asArray(ch.BodyLines).map(function (l) { return applyFields(l, values); });
            if (!childLines.length) return;
            if (!itemWritten) { n++; itemWritten = true; }
            for (var j = 0; j < childLines.length; j++) {
              var s = splitTextAndUrls(childLines[j]);
              if (s.text) lines.push("   - " + stripMarkers(s.text));
              s.urls.forEach(function (u) { lines.push("     " + u); });
            }
          });
        }
      });
    });
    return lines.join("\n").trim();
  }

  function avuiDDMMYYYY() {
    var d = new Date();
    function p(x) { return (x < 10 ? "0" : "") + x; }
    return p(d.getDate()) + "/" + p(d.getMonth() + 1) + "/" + d.getFullYear();
  }

  // Textos EDITABLES del correu (des de dades/email-textos.json; l'app del PC els
  // pot modificar). NOMES hi ha ASSUMPTE i COS: al cos hi surt TOT, i els
  // requeriments seleccionats s'insereixen alla on posis la variable
  // {REQUERIMENTS}. Variables: {REQUERIMENTS} {ID_GIA} {ADRECA} {ACTIVITAT}
  // {TITULAR} {DATA}. El cos accepta **negreta** i els enllacos http(s) es
  // tornen clicables sols. Aquests son els valors PER DEFECTE (fallback).
  var EMAIL_TEXTOS_DEFAULT = {
    assumpte: "GIA {ID_GIA} Requeriments",
    cos: [
      "ID GIA: {ID_GIA}",
      "Adreça: {ADRECA}",
      "Activitat: {ACTIVITAT}",
      "Titular: {TITULAR}",
      "",
      "Aquestes són les deficiències que s'han detectat a la visita de l'activitat per part de l'Ajuntament el dia {DATA} i que s'han d'esmenar:",
      "",
      "{REQUERIMENTS}",
      "",
      "**Com presentar la documentació / Cómo presentar la documentación**",
      "Heu de presentar **tota la documentació alhora** (important: no la presenteu per parts), mitjançant una **instància genèrica** de la seu electrònica de l'Ajuntament de Cornellà de Llobregat:",
      "https://seuelectronica.cornella.cat/portal/entidades.do?ent_id=1&idioma=2",
      "Debe presentar **toda la documentación a la vez** (importante: no la presente por partes), mediante una **instancia genérica** de la sede electrónica del Ayuntamiento de Cornellà de Llobregat.",
      "Indiqueu que la instància va **a l'atenció del Departament d'Activitats** / Indique que la instancia va dirigida **a la atención del Departamento de Actividades**, i feu-hi constar: ID GIA {ID_GIA}, Adreça {ADRECA}, Titular {TITULAR}.",
      "",
      "________________________________________",
      "",
      "IMPORTANT: aquest és un correu automàtic i no s'admeten respostes. Aquest llistat NO és definitiu ni oficial i pot variar respecte del requeriment oficial que rebreu properament. Per a qualsevol consulta podeu adreçar-vos al Departament d'Activitats de l'Ajuntament de Cornellà de Llobregat (Carrer de l'Energia, 97) o trucar al 93 377 02 12, extensió 1227.",
      "",
      "IMPORTANTE: este es un correo automático y no se admiten respuestas. Este listado NO es definitivo ni oficial y puede variar respecto del requerimiento oficial que recibirá próximamente. Para cualquier consulta puede dirigirse al Departamento de Actividades del Ayuntamiento de Cornellà de Llobregat (Calle de l'Energia, 97) o llamar al 93 377 02 12, extensión 1227."
    ].join("\n")
  };
  var emailTextos = EMAIL_TEXTOS_DEFAULT;   // se sobreescriu a inici() amb el JSON

  // Aplica el fitxer carregat sobre els valors per defecte (només claus amb text).
  function aplicarEmailTextos(obj) {
    if (!obj) return;
    var merged = {};
    Object.keys(EMAIL_TEXTOS_DEFAULT).forEach(function (k) {
      merged[k] = (obj[k] != null && String(obj[k]) !== "") ? obj[k] : EMAIL_TEXTOS_DEFAULT[k];
    });
    emailTextos = merged;
  }

  // Substitueix els placeholders {ID_GIA}, {ADRECA}, {ACTIVITAT}, {TITULAR}, {DATA}.
  function fillPh(s, h) {
    h = h || {};
    return String(s == null ? "" : s)
      .replace(/\{ID_GIA\}/g, h.ID_GIA || "")
      .replace(/\{ADRECA\}/g, h.ADRECA || "")
      .replace(/\{ACTIVITAT\}/g, h.ACTIVITAT || "")
      .replace(/\{TITULAR\}/g, h.TITULAR || "")
      .replace(/\{DATA\}/g, avuiDDMMYYYY());
  }

  // Converteix una línia de cos a HTML: escapa, **negreta** -> <b> i els URLs
  // http(s) es tornen enllaços clicables.
  function autolinkHtml(s) {
    var re = /(https?:\/\/[^\s]+)/g, out = "", last = 0, m;
    while ((m = re.exec(s)) !== null) {
      out += mdHtml(s.substring(last, m.index));
      out += '<a href="' + esc(m[1]) + '">' + esc(m[1]) + '</a>';
      last = re.lastIndex;
    }
    out += mdHtml(s.substring(last));
    return out;
  }

  // Cos del correu en TEXT pla (fallback mailto). El cos editable amb la variable
  // {REQUERIMENTS} substituïda per la llista de requeriments; **marques** tretes.
  function buildEmailBody(selSections, values) {
    var h = estat.header || {};
    var parts = fillPh(emailTextos.cos, h).split("{REQUERIMENTS}");
    var out = [];
    for (var i = 0; i < parts.length; i++) {
      out.push(stripMarkers(parts[i]));
      if (i < parts.length - 1) out.push(buildRequirementsList(selSections, values));
    }
    return out.join("");
  }

  // ------- Versió HTML del correu (per a EmailJS i la previsualització) --------
  function esc(s) {
    return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
  // **negreta** -> <b>, //cursiva// -> <i> (després d'escapar l'HTML).
  function mdHtml(s) {
    return esc(s).replace(/\*\*(.+?)\*\*/g, "<b>$1</b>").replace(/\/\/(.+?)\/\//g, "<i>$1</i>");
  }
  function urlHtml(u, ml) {
    var m = ml ? (' style="margin-left:' + ml + 'px"') : '';
    return '<div' + m + '><a href="' + esc(u) + '">' + esc(u) + '</a></div>';
  }

  // Requeriments en HTML (mateixa estructura que buildRequirementsList, però amb
  // format: seccions en negreta, SUBSECCIONS SUBRATLLADES, negreta/cursiva i
  // enllaços, per assemblar-se a l'informe de Word).
  function buildRequirementsHTML(selSections, values) {
    var H = [];
    var n = 0, lastSection = null;
    selSections.forEach(function (sec) {
      var idx = sec.Title.indexOf(" - ");
      var secName = idx >= 0 ? sec.Title.substring(0, idx).trim() : sec.Title.trim();
      var subT = idx >= 0 ? sec.Title.substring(idx + 3).trim() : null;
      if (secName !== lastSection) {
        H.push('<div style="font-weight:bold;margin-top:12px">' + esc(secName.toUpperCase()) + '</div>');
        lastSection = secName;
      }
      if (subT) H.push('<div style="text-decoration:underline;margin-top:4px">' + esc(subT) + '</div>');

      var pendingSub = null, pendingIntro = null;
      sec.Items.forEach(function (el) {
        if (el.Kind === "subsection") { pendingSub = el; pendingIntro = null; return; }
        if (el.Kind === "intro") { pendingIntro = el; return; }
        var itemLines = asArray(el.BodyLines).map(function (l) { return applyFields(l, values); });
        var hasChildren = !!(el.Children && el.Children.length > 0);
        if (!(el.Selected || hasChildren)) return;
        if (pendingSub) { H.push('<div style="text-decoration:underline;margin-top:4px">' + esc(stripMarkers(pendingSub.Short)) + '</div>'); pendingSub = null; }
        if (pendingIntro) {
          asArray(pendingIntro.BodyLines).forEach(function (l) {
            var s = splitTextAndUrls(applyFields(l, values));
            if (s.text) H.push('<div>' + mdHtml(s.text) + '</div>');
            s.urls.forEach(function (u) { H.push(urlHtml(u)); });
          });
          pendingIntro = null;
        }
        var itemWritten = false;
        if (itemLines.length > 0) {
          n++;
          var p0 = splitTextAndUrls(itemLines[0]);
          H.push('<div style="margin-top:6px">' + n + '. ' + mdHtml(p0.text) + '</div>');
          p0.urls.forEach(function (u) { H.push(urlHtml(u, 18)); });
          for (var i = 1; i < itemLines.length; i++) {
            var s = splitTextAndUrls(itemLines[i]);
            if (s.text) H.push('<div style="margin-left:18px">' + mdHtml(s.text) + '</div>');
            s.urls.forEach(function (u) { H.push(urlHtml(u, 18)); });
          }
          itemWritten = true;
        }
        if (hasChildren) {
          el.Children.forEach(function (ch) {
            var cl = asArray(ch.BodyLines).map(function (l) { return applyFields(l, values); });
            if (!cl.length) return;
            if (!itemWritten) { n++; itemWritten = true; }
            cl.forEach(function (cx) {
              var s = splitTextAndUrls(cx);
              if (s.text) H.push('<div style="margin-left:24px">- ' + mdHtml(s.text) + '</div>');
              s.urls.forEach(function (u) { H.push(urlHtml(u, 24)); });
            });
          });
        }
      });
    });
    return H.join("\n");
  }

  function buildEmailHTML(selSections, values) {
    var h = estat.header || {};
    var parts = fillPh(emailTextos.cos, h).split("{REQUERIMENTS}");
    var H = [];
    for (var i = 0; i < parts.length; i++) {
      // Cada línia del cos -> un <div> (línia buida = petit espai).
      parts[i].split("\n").forEach(function (l) {
        if (l.trim() === "") { H.push('<div style="height:8px"></div>'); }
        else { H.push('<div>' + autolinkHtml(l) + '</div>'); }
      });
      if (i < parts.length - 1) H.push(buildRequirementsHTML(selSections, values));
    }
    return '<div style="font-family:Calibri,Arial,sans-serif;font-size:14px;line-height:1.4">' + H.join("\n") + '</div>';
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

  // Els camps (opcions/text lliure) ja no tenen un pas propi: s'omplen inline
  // dins del text, al pas de deficiències i al de conclusions.
  var PASSOS = ["cataleg", "capcalera", "deficiencies", "conclusions", "final"];
  var passActual = 0;

  var HEADER_KEYS = ["ID_GIA", "EXP_NUM", "TITULAR", "ADRECA", "ACTIVITAT", "ORIGEN_TIPUS", "NUM_ANOTACIO", "DATA_ANOTACIO", "DATA_INSPECCIO"];

  // Claus de capçalera que NO es mostren al bloc genèric "opcional" (les gestiona
  // el bloc "Origen de l'informe" o són el camp principal / d'estat intern).
  var HEADER_SKIP_GENERIC = { ID_GIA: 1, ORIGEN_TIPUS: 1, NUM_ANOTACIO: 1, DATA_ANOTACIO: 1, DATA_INSPECCIO: 1 };

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
      carregarJson("dades/capcalera.json").catch(function () { return { Placeholders: HEADER_KEYS }; }),
      carregarJson("dades/email-textos.json").catch(function () { return null; })
    ]).then(function (res) {
      manifest = res[0];
      conclusions = res[1];
      capcalera = res[2];
      aplicarEmailTextos(res[3]);
      $("carregant").classList.add("ocult");
      mostrar($("navegacio"), true);
      muntarCataleg();
      muntarCapcalera();
      muntarDrive();
      // Reconnexió silenciosa amb Drive si ja s'havia autoritzat abans (sense
      // tornar a mostrar la finestra de permisos).
      if (window.Drive && Drive.reconnectarSilenci) {
        Drive.reconnectarSilenci().then(function (ok) { if (ok) estatDrive(true); });
      }
      // Precàrrega des del PC (eina "Enviar correu"): si la URL porta un informe
      // (#c=...), el carreguem i anem directament a la pantalla d'enviar.
      var pre = llegirInformePrecarregat();
      if (pre) { precarregarInforme(pre).catch(function () { anarA(0); }); }
      else { anarA(0); }
    }).catch(function (e) {
      $("carregant").innerHTML = '<span class="error">Error carregant les dades: ' + e.message +
        "<br>Comprova que el PC hagi pujat les dades (Actualitzar.bat) i que GitHub Pages estigui actiu.</span>";
    });
  }

  // ------- Precàrrega d'un informe des del PC (eina "Enviar correu") ----------
  // El PC obre aquesta web amb l'últim informe al fragment: #c=<json url-encoded>.
  // El fragment queda al navegador (no s'envia al servidor).
  function llegirInformePrecarregat() {
    try {
      var h = location.hash || "";
      var m = h.match(/[#&]c=([^&]+)/);
      if (!m) return null;
      return JSON.parse(decodeURIComponent(m[1]));
    } catch (e) { return null; }
  }

  function precarregarInforme(d) {
    estat.cataleg = d.CatalegBaseName || (asArray(manifest.Catalegs)[0] || null);
    estat.header = d.Header || {};
    estat.keys = new Set(asArray(d.SelectedKeys));
    estat.conclTitles = new Set(asArray(d.ConclusionTexts));
    estat.fieldValues = d.FieldValues || {};
    if (d.Header && d.Header.EMAIL) { $("in-destinatari").value = d.Header.EMAIL; }
    // Netegem el fragment perquè no es repeteixi en recarregar.
    try { history.replaceState(null, "", location.pathname + location.search); } catch (e) {}
    return muntarFinal().then(function () { anarA(PASSOS.indexOf("final")); });
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
      if (HEADER_SKIP_GENERIC[k]) return; // ID_GIA té camp propi; l'anotació/inspecció van al bloc Origen
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
    muntarOrigen();
  }

  // Bloc "Origen de l'informe": documentació aportada (Núm./Data d'anotació) o
  // visita d'inspecció (Data d'inspecció). Al mòbil, per defecte INSPECCIÓ.
  // Es lliga UN sol cop (a muntarCapcalera); l'estat es reinicia a reinici().
  function muntarOrigen() {
    if (!estat.header.ORIGEN_TIPUS) estat.header.ORIGEN_TIPUS = "insp";
    [].forEach.call(document.getElementsByName("origen"), function (r) {
      r.checked = (r.value === estat.header.ORIGEN_TIPUS);
      r.addEventListener("change", function () {
        if (r.checked) { estat.header.ORIGEN_TIPUS = r.value; renderCampsOrigen(); }
      });
    });
    renderCampsOrigen();
  }

  // Mostra els camps que corresponen a l'origen triat (prefills des de l'estat).
  function renderCampsOrigen() {
    var cont = $("camps-origen");
    cont.innerHTML = "";
    var camps = (estat.header.ORIGEN_TIPUS === "doc")
      ? [["NUM_ANOTACIO", "Núm. d'anotació"], ["DATA_ANOTACIO", "Data d'anotació (dd/mm/aaaa)"]]
      : [["DATA_INSPECCIO", "Data d'inspecció (dd/mm/aaaa)"]];
    camps.forEach(function (c) {
      var k = c[0];
      var lbl = document.createElement("label");
      lbl.textContent = c[1];
      var inp = document.createElement("input");
      inp.type = "text";
      inp.id = "hdr-" + k;
      inp.value = estat.header[k] || "";
      inp.addEventListener("input", function () { estat.header[k] = inp.value; });
      cont.appendChild(lbl);
      cont.appendChild(inp);
    });
  }

  function omplirCapcalera(act) {
    HEADER_KEYS.forEach(function (k) {
      if (k === "ID_GIA") return;
      if (act[k] !== undefined) {
        estat.header[k] = act[k];
        var inp = $("hdr-" + k);
        if (inp) inp.value = act[k];
      }
    });
    // Dades del titular (només a l'app, per confirmar-les). Es desmarca la
    // casella perquè s'hagin de tornar a confirmar amb cada cerca nova.
    $("tit-rao").textContent = act.TITULAR || "—";
    $("tit-mobil").textContent = act.MOBIL || "—";
    $("tit-email").textContent = act.EMAIL || "—";
    estat.header.MOBIL = act.MOBIL || "";
    estat.header.EMAIL = act.EMAIL || "";
    $("chk-titular").checked = false;
    $("det-capcalera").open = true;
  }

  function cercarActivitat() {
    var gia = $("in-gia").value.trim();
    var msg = $("msg-cerca");
    if (!gia) { msg.textContent = "Escriu un ID GIA."; return; }

    if (!window.Drive || !Drive.disponible()) {
      // Sense Drive: provem el fitxer local de DEMO (activitats falses).
      carregarJson("dades/activitats.json").then(function (data) {
        var act = ((data && data.ById) || {})[gia];
        if (!act) { msg.textContent = "(Demo) Prova amb 1001, 1002 o 1003 — o omple les dades a mà."; return; }
        omplirCapcalera(act);
        msg.textContent = "(Demo) Dades de prova carregades.";
      }).catch(function () {
        msg.textContent = "Drive no configurat: omple les dades a mà (o el PC les omplirà en generar).";
      });
      return;
    }

    msg.textContent = "Connectant amb Drive…";
    var pas = Drive.connectat() ? Promise.resolve() : Drive.connectar();
    pas.then(function () {
      estatDrive(true);
      msg.textContent = "Cercant…";
      return Drive.llegirActivitats();
    }).then(function (data) {
      var act = ((data && data.ById) || {})[gia];
      if (!act) { msg.textContent = "L'ID GIA " + gia + " no s'ha trobat a la base de dades."; return; }
      omplirCapcalera(act);
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
        // item: bloc amb la casella + el text (i camps inline) que apareix en
        // marcar-la.
        var block = document.createElement("div");
        block.className = "item-block";
        block.dataset.text = (el.Short || "").toLowerCase();

        var key = itemKey(sec.Title, el.Short);
        var fila = filaCheckbox(key, el.Short, el, "item");
        block.appendChild(fila.wrap);

        var det = document.createElement("div");
        det.className = "detall ocult";
        renderRich(det, richTextOf(el));
        block.appendChild(det);
        var mostraDet = function () { mostrar(det, fila.cb.checked && det.childNodes.length > 0); };
        fila.cb.addEventListener("change", mostraDet);
        mostraDet();

        // fills
        asArray(el.Children).forEach(function (ch) {
          var ckey = itemKey(sec.Title, el.Short, ch.Short);
          var cf = filaCheckbox(ckey, ch.Short, ch, "fill");
          block.appendChild(cf.wrap);

          var cdet = document.createElement("div");
          cdet.className = "detall detall-fill ocult";
          renderRich(cdet, richTextOf(ch));
          block.appendChild(cdet);
          var mostraCdet = function () { mostrar(cdet, cf.cb.checked && cdet.childNodes.length > 0); };
          cf.cb.addEventListener("change", mostraCdet);
          mostraCdet();

          // propagació (com el TreeView del PC): marcar/desmarcar l'item
          // arrossega els fills.
          fila.cb.addEventListener("change", function () {
            cf.cb.checked = fila.cb.checked;
            if (fila.cb.checked) estat.keys.add(ckey); else estat.keys.delete(ckey);
            mostraCdet();
          });
        });

        cont.appendChild(block);
      });
    });
    $("filtre-def").value = "";
    $("filtre-def").oninput = function () { filtrarArbre(this.value); };
  }

  function filaCheckbox(key, etiqueta, el, classe) {
    // La fila SENCERA és un <label> amb l'<input> a dins: així tocar qualsevol
    // punt de la fila marca/desmarca de manera fiable, sense dependre de cap
    // id/for (les claus tenen espais i símbols que feien fallar l'associació).
    var wrap = document.createElement("label");
    wrap.className = classe;
    wrap.dataset.text = (etiqueta || "").toLowerCase();
    var cb = document.createElement("input");
    cb.type = "checkbox";
    cb.dataset.key = key;
    cb.checked = estat.keys.has(key);
    cb.addEventListener("change", function () {
      if (cb.checked) estat.keys.add(key); else estat.keys.delete(key);
    });
    var txt = document.createElement("span");
    txt.className = "etq";
    txt.textContent = etiqueta;
    var prev = primeraLiniaText(el);
    if (prev) {
      var sp = document.createElement("span");
      sp.className = "preview";
      sp.textContent = prev;
      txt.appendChild(sp);
    }
    wrap.appendChild(cb);
    wrap.appendChild(txt);
    return { wrap: wrap, cb: cb };
  }

  // Reconstrueix estat.keys a partir de les caselles realment marcades al DOM.
  // És la font de veritat abans de generar camps, correu o paquet, de manera
  // que SelectedKeys, el correu i l'informe sempre coincideixen.
  function recollirKeysDelDOM() {
    var cont = $("arbre-def");
    if (!cont) return;
    var boxes = cont.querySelectorAll("input[type=checkbox]");
    if (!boxes.length) return; // l'arbre encara no s'ha construït: no toquem res
    var s = new Set();
    [].forEach.call(boxes, function (cb) {
      if (cb.checked && cb.dataset.key) s.add(cb.dataset.key);
    });
    estat.keys = s;
  }

  function recollirConclTitlesDelDOM() {
    var cont = $("llista-concl");
    if (!cont) return;
    var boxes = cont.querySelectorAll("input[type=checkbox]");
    if (!boxes.length) return;
    var s = new Set();
    [].forEach.call(boxes, function (cb) {
      if (cb.checked && cb.dataset.title) s.add(cb.dataset.title);
    });
    estat.conclTitles = s;
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
    [].forEach.call(cont.querySelectorAll(".item-block"), function (w) {
      var match = !needle || (w.dataset.text || "").indexOf(needle) >= 0;
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
      // El títol (amb la casella) ha de ser clicable, però el cos NO pot ser-ho
      // (conté selects/inputs). Per això la casella i el títol van dins d'un
      // <label> i el cos va a fora.
      var wrap = document.createElement("div");
      wrap.className = "concl";

      var cap = document.createElement("label");
      cap.className = "concl-cap";
      var cb = document.createElement("input");
      cb.type = "checkbox";
      cb.dataset.title = c.Title;
      cb.checked = estat.conclTitles.has(c.Title);
      cb.addEventListener("change", function () {
        if (cb.checked) estat.conclTitles.add(c.Title); else estat.conclTitles.delete(c.Title);
      });
      var tit = document.createElement("span");
      tit.className = "etq";
      tit.textContent = c.Title;
      cap.appendChild(cb);
      cap.appendChild(tit);

      var cos = document.createElement("div");
      cos.className = "cos";
      renderRich(cos, c.Body);

      wrap.appendChild(cap);
      wrap.appendChild(cos);
      cont.appendChild(wrap);
    });

    // Frases fixes (::SEMPRE::): si tenen camps, també s'han de poder omplir.
    var always = asArray(conclusions.Always);
    var withFields = always.filter(function (a) { return /\[OPCIO:|\[CAMP:/.test(a); });
    if (withFields.length) {
      var fx = document.createElement("div");
      fx.className = "concl-fixes";
      var t = document.createElement("div");
      t.className = "ajuda";
      t.textContent = "Es posa sempre al final:";
      fx.appendChild(t);
      withFields.forEach(function (a) {
        var cos = document.createElement("div");
        cos.className = "cos";
        renderRich(cos, a);
        fx.appendChild(cos);
      });
      cont.appendChild(fx);
    }
  }

  // ------- Pas 5: final -------------------------------------------------------
  function muntarFinal() {
    return carregarCataleg().then(function (catalog) {
      recollirKeysDelDOM();
      var selSections = reconstructSelection(catalog, estat.keys);
      estat.emailHTML = buildEmailHTML(selSections, estat.fieldValues);
      estat.emailText = buildEmailBody(selSections, estat.fieldValues);
      $("prev-requeriments").innerHTML = estat.emailHTML;
      // Destinatari per defecte: l'e-mail del titular (Raó soc. E-mail), editable.
      if (!$("in-destinatari").value) {
        $("in-destinatari").value = estat.header.EMAIL || ((window.CONFIG && CONFIG.EMAIL_DESTINATARI) || "");
      }
    });
  }

  function assumpte() {
    return fillPh(emailTextos.assumpte, estat.header || {});
  }

  function enviarEmail() {
    var dest = $("in-destinatari").value.trim();
    var subj = assumpte();
    var msg = $("msg-email");
    if (!dest) { msg.innerHTML = '<span class="error">Indica un destinatari.</span>'; return; }
    if (window.Mail && Mail.configurat()) {
      // Enviament automàtic (un sol clic): enviem la versió HTML amb format.
      msg.textContent = "Enviant…";
      $("btn-email").disabled = true;
      Mail.enviar(dest, subj, estat.emailHTML || "").then(function () {
        msg.innerHTML = "✅ Correu enviat a " + dest + ".";
      }).catch(function (e) {
        msg.innerHTML = '<span class="error">No s\'ha pogut enviar: ' + e.message + "</span>";
      }).then(function () { $("btn-email").disabled = false; });
    } else {
      // Sense EmailJS: obrim l'app de correu amb la versió en text (mailto no
      // admet HTML). Per tenir el format (subratllat/negreta) cal EmailJS.
      window.location.href = "mailto:" + encodeURIComponent(dest) +
        "?subject=" + encodeURIComponent(subj) + "&body=" + encodeURIComponent(estat.emailText || "");
    }
  }

  function construirPaquet() {
    recollirKeysDelDOM();
    recollirConclTitlesDelDOM();
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
  }

  // Reinicia el formulari per fer un informe nou i torna a l'inici.
  function reinici() {
    estat.keys = new Set();
    estat.conclTitles = new Set();
    estat.fieldValues = {};
    estat.fieldOrder = [];
    estat.fieldDefs = {};
    estat.header = {};
    $("in-gia").value = "";
    $("msg-cerca").textContent = "";
    $("msg-final").textContent = "";
    $("msg-email").textContent = "";
    $("prev-requeriments").value = "";
    $("chk-titular").checked = false;
    $("tit-rao").textContent = "—";
    $("tit-mobil").textContent = "—";
    $("tit-email").textContent = "—";
    [].forEach.call($("camps-capcalera").querySelectorAll("input"), function (i) { i.value = ""; });
    // Torna l'origen al valor per defecte del mòbil (visita d'inspecció).
    estat.header.ORIGEN_TIPUS = "insp";
    [].forEach.call(document.getElementsByName("origen"), function (r) { r.checked = (r.value === "insp"); });
    renderCampsOrigen();
    $("arbre-def").innerHTML = "";
    $("llista-concl").innerHTML = "";
    anarA(0);
    window.scrollTo(0, 0);
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
    if (p === "capcalera") {
      if (!$("in-gia").value.trim()) {
        $("msg-cerca").innerHTML = '<span class="error">Cal un ID GIA.</span>';
        return false;
      }
      if (!$("chk-titular").checked) {
        $("msg-cerca").innerHTML = '<span class="error">Has de confirmar les dades del titular.</span>';
        return false;
      }
    }
    if (p === "deficiencies") {
      recollirKeysDelDOM();
      if (estat.keys.size === 0) {
        alert("Selecciona almenys una deficiència.");
        return false;
      }
    }
    return true;
  }

  function seguent() {
    if (!validarPas()) return;
    var seg = passActual + 1;
    if (seg >= PASSOS.length) { reinici(); return; } // "Fet": torna a l'inici
    // accions de preparació en entrar a certs passos
    var entrar = PASSOS[seg];
    var prep = Promise.resolve();
    if (entrar === "deficiencies") prep = carregarCataleg().then(muntarDeficiencies);
    else if (entrar === "conclusions") { muntarConclusions(); }
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
    prep.then(function () { anarA(ant); });
  }

  // ------- Arrencada ----------------------------------------------------------
  document.addEventListener("DOMContentLoaded", function () {
    $("btn-seguent").addEventListener("click", seguent);
    $("btn-enrere").addEventListener("click", enrere);
    inici();
  });
})();
