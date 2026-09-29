// El FORMAT del correu de requeriments: el mateix que l'informe de REQ1.
// -----------------------------------------------------------------------------
// Es la copia en JavaScript de suport/CorreuFormat.ps1 (i de Build-CatalegBlocs
// / _BlocsDItem de MotorInforme.ps1), i HA DE DONAR EL MATEIX HTML: hi ha una
// prova (suport/tests/proves/04-correu.ps1) que passa la mateixa seleccio de
// REQ1 pels dos i compara el resultat caracter a caracter.
//
// Les MIDES no s'escriuen aqui: arriben de dades/correu-format.json, que genera
// el PC a partir de $ReportFormatConfig (Format.ps1), que es d'on les treu el
// Word. Aixi el correu del mobil i l'informe no poden divergir.
//
// Funcions PURES: no toquen el DOM. Exposa window.Correu (al navegador) i
// module.exports (a Node, per a la prova).

(function (arrel) {
  function esc(s) {
    return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  // _TextToHtml (EnviarCorreu.ps1): escapa, aparta els URLs ABANS de mirar la
  // cursiva (un "https://" porta un "//"), **negreta**, //cursiva// i torna a
  // posar els URLs com a enllac.
  function textHtml(s) {
    if (s == null || s === "") return "";
    var h = esc(s);
    var urls = [];
    h = h.replace(/(https?:\/\/[^\s<]+)/g, function (u) { urls.push(u); return "\u0001" + (urls.length - 1) + "\u0001"; });
    h = h.replace(/\*\*(.+?)\*\*/g, "<b>$1</b>").replace(/\/\/(.+?)\/\//g, "<i>$1</i>");
    for (var i = 0; i < urls.length; i++) {
      h = h.split("\u0001" + i + "\u0001").join('<a href="' + urls[i] + '">' + urls[i] + "</a>");
    }
    return h.replace(/\r\n/g, "\n").replace(/\n/g, "<br>");
  }

  // _SplitTextAndUrls (Camps.ps1).
  function splitTextAndUrls(line) {
    if (line == null || String(line).trim() === "") return { text: "", urls: [] };
    line = String(line);
    if (line.indexOf("[[URL]] ") === 0) return { text: "", urls: [line.substring(8).trim()] };
    var m = /https?:\/\//.exec(line);
    if (!m) return { text: line.trim(), urls: [] };
    var text = line.substring(0, m.index).trim();
    var urls = line.substring(m.index).split(/\s+/).filter(function (t) { return /^https?:\/\//i.test(t); });
    return { text: text, urls: urls };
  }

  // _EsFraseTancament (Format.ps1).
  function esFraseTancament(t) {
    return String(t == null ? "" : t).split("**").join("").split("//").join("").trim().toLowerCase()
      .indexOf("ho poso al seu coneixement") === 0;
  }

  // --- El contracte d'un paragraf (_CorreuParagrafHtml) ----------------------
  function num(x) { return String(Number(x)); }
  function base(fmt) {
    return "font-family:" + fmt.Lletra + ";font-size:" + num(fmt.MidaPt) + "pt;line-height:" + fmt.InterlineatPct + "%;color:#000000";
  }
  function buit(fmt) { return '<p style="' + base(fmt) + ';margin:0">&nbsp;</p>'; }
  function paragraf(contingut, m, fmt, abansTab) {
    var b = base(fmt), esq = Number(m.Esq) | 0, pen = Number(m.Penjat) | 0;
    var marge = num(m.Abans) + "pt 0 " + num(m.Despres) + "pt ";
    if (abansTab != null && pen > 0) {
      return '<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin:' +
        marge + (esq - pen) + 'px"><tr><td style="' + b + ";width:" + pen +
        'px;padding:0;vertical-align:top">' + abansTab + '</td><td style="' + b +
        ";padding:0;vertical-align:top;text-align:" + m.Alinea + '">' + contingut + "</td></tr></table>";
    }
    if (abansTab != null) contingut = abansTab + " " + contingut;
    var st = b + ";margin:" + marge + esq + "px";
    if (pen !== 0) st += ";text-indent:" + (-pen) + "px";
    st += ";text-align:" + m.Alinea;
    return '<p style="' + st + '">' + contingut + "</p>";
  }
  function enllac(url, midaPt) {
    var u = esc(url);
    return '<a href="' + u + '" style="font-size:' + num(midaPt) + 'pt;word-break:break-all">' + u + "</a>";
  }

  // --- Blocs -> HTML (_CorreuBlocsAHtml) --------------------------------------
  function blocsAHtml(blocs, fmt) {
    var estat = { PrimerFill: false, Escrits: 0, UltimBuit: true, out: [] };
    blocsRec(blocs, fmt, estat);
    return estat.out.join("");
  }
  function blocsRec(blocs, fmt, estat) {
    function posaBuit() { estat.out.push(buit(fmt)); estat.UltimBuit = true; }
    (blocs || []).forEach(function (b) {
      if (!b) return;
      var t = String(b.T), fill = !!b.Fill, B = fmt.Blocs, h;
      if (t === "unitat") {
        var abans = estat.Escrits;
        estat.PrimerFill = true;
        blocsRec(b.Blocs, fmt, estat);
        if (estat.Escrits > abans && fmt.Aire.item) posaBuit();
        return;
      }
      if (t === "aire") { if (fmt.Aire[String(b.Clau).toLowerCase()]) posaBuit(); return; }
      if (t === "espai") { posaBuit(); return; }
      if (t === "separa") { if (!estat.UltimBuit) posaBuit(); return; }
      switch (t) {
        case "seccio": h = paragraf(esc(String(b.Text).toUpperCase()), B.seccio, fmt); break;
        case "subseccio": h = paragraf(esc(b.Text), B.subseccio, fmt); break;
        case "item": h = paragraf("<b>" + esc(b.Num) + "</b> " + textHtml(b.Text), B.item, fmt); break;
        case "cos":
          var x = textHtml(b.Text);
          if (b.Negreta) x = "<b>" + x + "</b>";
          h = paragraf(x, fill ? B.cosFill : B.cos, fmt); break;
        case "enllac":
          var me = fill ? B.enllacFill : B.enllac;
          h = paragraf(enllac(b.Url, me.MidaPt), me, fmt); break;
        case "pic":
          var primer = ("First" in b) ? !!b.First : estat.PrimerFill;
          var mp = fill ? (primer ? B.picFillPrimer : B.picFill) : (primer ? B.picPrimer : B.pic);
          estat.PrimerFill = false;
          h = paragraf(textHtml(b.Text), mp, fmt, "•"); break;
        case "conclusiocap": h = paragraf("<b>" + esc(b.Text) + "</b>", B.conclusiocap, fmt); break;
        case "conclusio": h = paragraf(textHtml(b.Text), B.conclusio, fmt); break;
        default: throw new Error("Correu.blocsAHtml: tipus de bloc desconegut '" + t + "'");
      }
      estat.out.push(h);
      estat.Escrits++;
      estat.UltimBuit = false;
    });
  }

  // --- La seleccio del cataleg -> blocs (Build-CatalegBlocs + _BlocsDItem) ----
  // linies(node): les linies del node amb els camps resolts (applyFieldsToLines).
  function blocsDeLinia(linia, fill) {
    var out = [];
    if (linia == null || String(linia).trim() === "") return out;
    var p = splitTextAndUrls(linia);
    if (p.text.trim() !== "") out.push({ T: "cos", Text: p.text, Fill: fill });
    p.urls.forEach(function (u) { out.push({ T: "enllac", Url: u, Fill: fill }); });
    return out;
  }
  function blocsDItem(el, linies, comptador) {
    var dins = [], ls = linies(el), fills = el.Children || [], escrit = false;
    if ((el.Selected || fills.length > 0) && ls.length > 0) {
      comptador.n++;
      var p0 = splitTextAndUrls(ls[0]);
      dins.push({ T: "item", Num: comptador.n + ".", Text: p0.text });
      p0.urls.forEach(function (u) { dins.push({ T: "enllac", Url: u }); });
      for (var i = 1; i < ls.length; i++) dins = dins.concat(blocsDeLinia(ls[i], false));
      escrit = true;
    }
    fills.forEach(function (ch) {
      var cl = linies(ch);
      if (!cl.length) return;
      if (!escrit) { comptador.n++; escrit = true; }
      var pc = splitTextAndUrls(cl[0]);
      if (pc.text.trim() !== "") dins.push({ T: "pic", Text: pc.text, Fill: true });
      pc.urls.forEach(function (u) { dins.push({ T: "enllac", Url: u, Fill: true }); });
      for (var j = 1; j < cl.length; j++) dins = dins.concat(blocsDeLinia(cl[j], true));
    });
    if (!escrit) return [];
    return [{ T: "unitat", Blocs: dins }];
  }
  function blocsDeSeleccio(seccions, linies) {
    var b = [], comptador = { n: 0 }, darrera = null;
    (seccions || []).forEach(function (sec) {
      var titol = String(sec.Title || ""), k = titol.indexOf(" - ");
      if (k >= 0) {
        var nomSec = titol.substring(0, k).trim();
        if (nomSec !== darrera) { b.push({ T: "seccio", Text: nomSec }); b.push({ T: "aire", Clau: "seccio" }); darrera = nomSec; }
        b.push({ T: "subseccio", Text: titol.substring(k + 3).trim() }); b.push({ T: "aire", Clau: "subseccio" });
      } else {
        b.push({ T: "seccio", Text: titol }); b.push({ T: "aire", Clau: "seccio" }); darrera = titol;
      }
      // Un text fix es de la SECCIO (abans de la primera subseccio) o de la
      // SUBSECCIO on es: la mateixa regla que Build-CatalegBlocs.
      var subPendent = null, introPendent = null, introSeccio = null, dinsSub = false;
      (sec.Items || []).forEach(function (el) {
        if (el.Kind === "subsection") { subPendent = el; dinsSub = true; introPendent = null; return; }
        if (el.Kind === "intro") { if (dinsSub) introPendent = el; else introSeccio = el; return; }
        var bi = blocsDItem(el, linies, comptador);
        if (!bi.length) return;
        if (introSeccio) {
          linies(introSeccio).forEach(function (ln) { b = b.concat(blocsDeLinia(ln, false)); });
          b.push({ T: "aire", Clau: "intro" }); introSeccio = null;
        }
        if (subPendent) { b.push({ T: "subseccio", Text: String(subPendent.Short) }); b.push({ T: "aire", Clau: "subseccio" }); subPendent = null; }
        if (introPendent) {
          linies(introPendent).forEach(function (ln) { b = b.concat(blocsDeLinia(ln, false)); });
          b.push({ T: "aire", Clau: "intro" }); introPendent = null;
        }
        b = b.concat(bi);
      });
    });
    return b;
  }

  // Les conclusions triades (_BlocsConclusions sense el tancament: "Ho poso al
  // seu coneixement" i "Cornella de Llobregat," no van al correu). Sense cap
  // conclusio no hi ha bloc: un "CONCLUSIONS" sol no diu res.
  function blocsConclusions(titol, textos) {
    textos = (textos || []).filter(function (t) { return String(t || "").trim() !== ""; });
    if (!textos.length) return [];
    var b = [{ T: "aire", Clau: "conclusions" }];
    if (String(titol || "").trim() !== "") b.push({ T: "conclusiocap", Text: titol });
    textos.forEach(function (t) {
      if (esFraseTancament(t)) b.push({ T: "separa" });
      b.push({ T: "conclusio", Text: t });
    });
    return b;
  }

  // --- Capcalera, origen i frase d'introduccio --------------------------------
  function omplePlantilla(pl, valors) {
    return String(pl == null ? "" : pl).replace(/<<\s*([A-Za-z0-9_]+)\s*>>/g, function (_, k) {
      return (valors && valors[k] != null) ? String(valors[k]) : "";
    });
  }
  // _BuildOrigenText: el text de la linia "Objecte:".
  function origenText(plantilles, h) {
    var tipus = (h && h.ORIGEN_TIPUS) || "doc";
    if (tipus === "cap") return "";
    return omplePlantilla(tipus === "insp" ? plantilles.insp : plantilles.doc, h);
  }
  // _CorreuCapcaleraHtml: les linies SENSE valor no surten.
  function capcaleraHtml(linies, fmt) {
    var px = Number(fmt.EtiquetaPx) | 0, m = { Esq: px, Penjat: px, Abans: 0, Despres: 0, Alinea: "left" };
    return (linies || []).filter(function (l) { return String(l.Valor || "").trim() !== ""; }).map(function (l) {
      return paragraf(esc(String(l.Valor).trim()), m, fmt, "<b>" + esc(l.Etiqueta) + "</b>");
    }).join("");
  }
  // _CorreuIntro.
  function intro(textos, o, avui) {
    var clau = (o.ORIGEN_TIPUS === "insp") ? "introInsp"
      : ((String(o.NUM_ANOTACIO || "").trim() && String(o.DATA_ANOTACIO || "").trim()) ? "introDoc" : "introDocSenseAnotacio");
    var dataInsp = String(o.DATA_INSPECCIO || "").trim() ? o.DATA_INSPECCIO : avui;
    return String((textos && textos[clau]) || "")
      .split("{NUM_ANOTACIO}").join(o.NUM_ANOTACIO || "")
      .split("{DATA_ANOTACIO}").join(o.DATA_ANOTACIO || "")
      .split("{DATA_INSPECCIO}").join(dataInsp);
  }
  // _CorreuCosAHtml: una linia = un paragraf; {CAPCALERA} i {REQUERIMENTS},
  // sols a la seva linia, son blocs.
  function cosAHtml(cos, omple, capHtml, reqHtml, fmt) {
    var out = [];
    String(cos == null ? "" : cos).replace(/\r\n/g, "\n").split("\n").forEach(function (ln) {
      var t = ln.trim();
      if (t === "{CAPCALERA}") { out.push(capHtml); return; }
      if (t === "{REQUERIMENTS}") { out.push(reqHtml); return; }
      if (t === "") { out.push(buit(fmt)); return; }
      out.push(paragraf(textHtml(omple(ln)), fmt.Blocs.cos, fmt));
    });
    return '<div style="' + base(fmt) + '">' + out.join("") + "</div>";
  }

  var Correu = {
    esc: esc, textHtml: textHtml, splitTextAndUrls: splitTextAndUrls,
    blocsAHtml: blocsAHtml, blocsDeSeleccio: blocsDeSeleccio, blocsConclusions: blocsConclusions,
    omplePlantilla: omplePlantilla, origenText: origenText, capcaleraHtml: capcaleraHtml,
    intro: intro, cosAHtml: cosAHtml
  };
  if (typeof module !== "undefined" && module.exports) module.exports = Correu;
  else arrel.Correu = Correu;
})(this);
