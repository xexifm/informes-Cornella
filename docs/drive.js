// Integració amb Google Drive (opcional).
// -----------------------------------------------------------------------------
// Fa servir Google Identity Services (GIS) per obtenir un token d'accés i la
// API REST de Drive (v3) per llegir activitats.json i pujar el paquet.
//
// Si CONFIG.GOOGLE_CLIENT_ID està buit, tot queda desactivat i el formulari
// funciona en mode manual (descarregar el paquet, capçalera a mà).
//
// Exposa window.Drive amb: disponible(), connectar(), llegirActivitats(),
// pujarPaquet(nom, objecte).

(function () {
  var SCOPE = "https://www.googleapis.com/auth/drive";
  var tokenClient = null;
  var accessToken = null;
  var gisCarregat = false;

  function disponible() {
    return !!(window.CONFIG && CONFIG.GOOGLE_CLIENT_ID);
  }

  // Carrega la biblioteca GIS sota demanda (només si cal connectar).
  function carregarGis() {
    return new Promise(function (resolve, reject) {
      if (gisCarregat) return resolve();
      var s = document.createElement("script");
      s.src = "https://accounts.google.com/gsi/client";
      s.async = true;
      s.defer = true;
      s.onload = function () { gisCarregat = true; resolve(); };
      s.onerror = function () { reject(new Error("No s'ha pogut carregar Google Identity Services.")); };
      document.head.appendChild(s);
    });
  }

  // Recorda (en aquest navegador) que l'usuari ja ha autoritzat alguna vegada.
  // Google NO dóna "refresh tokens" a les apps de navegador (els tokens caduquen
  // ~1h), però si ja s'ha consentit es pot renovar el token EN SILENCI (sense
  // tornar a mostrar la finestra de permisos) mentre hi hagi sessió de Google.
  var CLAU_AUTORITZAT = "driveAutoritzat";
  function jaAutoritzat() {
    try { return localStorage.getItem(CLAU_AUTORITZAT) === "1"; } catch (e) { return false; }
  }
  function recordarAutoritzat() {
    try { localStorage.setItem(CLAU_AUTORITZAT, "1"); } catch (e) {}
  }

  // MENYS INICIS DE SESSIÓ (octubre 2026, l'usuari: «em fa iniciar sessió a
  // Google moltes vegades»). Tres coses ho feien:
  //  1. El token només vivia a la memòria de la PÀGINA: anar del plànol a
  //     l'informe («Fer informe»), tornar enrere o recarregar el perdia, i la
  //     renovació "silenciosa" de l'arrencada no surt d'un toc de l'usuari, o
  //     sigui que el mòbil li bloqueja la finestra. Ara es desa a la
  //     sessionStorage fins que caduca: dins de la mateixa sessió de l'app, les
  //     pàgines se'l passen. (La sessionStorage i no la localStorage: no queda
  //     al telèfon en tancar l'app.)
  //  2. Si la renovació silenciosa fallava, es demanava amb prompt "consent":
  //     la pantalla de permisos SENCERA cada vegada. Ara amb prompt "" Google
  //     només la ensenya si de debò cal (la primera vegada).
  //  3. Sense dir quin compte, amb més d'un compte al telèfon calia triar-lo
  //     cada vegada. Ara es recorda el compte (localStorage) i es passa com a
  //     'hint'.
  var CLAU_TOKEN = "driveToken";
  var CLAU_COMPTE = "driveCompte";
  var caducaA = 0;
  function desarToken(t, segons) {
    accessToken = t;
    caducaA = Date.now() + (Number(segons) > 0 ? Number(segons) : 3600) * 1000;
    try { sessionStorage.setItem(CLAU_TOKEN, JSON.stringify({ t: t, exp: caducaA })); } catch (e) {}
  }
  function oblidarToken() {
    accessToken = null; caducaA = 0;
    try { sessionStorage.removeItem(CLAU_TOKEN); } catch (e) {}
  }
  // El de la sessió, si encara li queden més de 2 minuts.
  (function recuperarToken() {
    try {
      var o = JSON.parse(sessionStorage.getItem(CLAU_TOKEN) || "null");
      if (o && o.t && Number(o.exp) - Date.now() > 120000) { accessToken = o.t; caducaA = Number(o.exp); }
    } catch (e) {}
  })();
  function compte() {
    try { return localStorage.getItem(CLAU_COMPTE) || ""; } catch (e) { return ""; }
  }
  // L'adreça del compte, per al 'hint' de la propera vegada (no bloqueja res).
  function recordarCompte() {
    fetch("https://www.googleapis.com/drive/v3/about?fields=user(emailAddress)", { headers: { Authorization: "Bearer " + accessToken } })
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (d) { if (d && d.user && d.user.emailAddress) { try { localStorage.setItem(CLAU_COMPTE, d.user.emailAddress); } catch (e) {} } })
      .catch(function () {});
  }

  function _initClient() {
    if (tokenClient) return;
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CONFIG.GOOGLE_CLIENT_ID,
      scope: SCOPE,
      callback: function () {},        // es reassigna a cada petició
      error_callback: function () {}   // es reassigna a cada petició
    });
  }

  // Demana un token a Google. prompt "" = només ensenya el que calgui (res, si
  // ja s'ha consentit i hi ha sessió); "none" = no ensenya res mai (falla si
  // calgués).
  function _demanarToken(prompt) {
    return carregarGis().then(function () {
      return new Promise(function (resolve, reject) {
        try {
          _initClient();
          tokenClient.callback = function (resp) {
            if (resp && resp.access_token) {
              desarToken(resp.access_token, resp.expires_in);
              recordarAutoritzat();
              if (!compte()) recordarCompte();
              resolve(accessToken);
            } else {
              reject(new Error("No s'ha obtingut el token de Google."));
            }
          };
          tokenClient.error_callback = function (err) {
            reject(new Error((err && err.message) || "Autorització de Google cancel·lada."));
          };
          var opcions = { prompt: prompt };
          if (compte()) opcions.hint = compte();
          tokenClient.requestAccessToken(opcions);
        } catch (e) { reject(e); }
      });
    });
  }

  // Connecta amb Drive (des d'un TOC de l'usuari). Amb un token encara bo, res;
  // si no, UNA sola petició: Google només ensenya la finestra si cal.
  function connectar() {
    if (!disponible()) return Promise.reject(new Error("Google Drive no configurat."));
    if (connectat()) return Promise.resolve(accessToken);
    return _demanarToken("");
  }

  // Reconnexió silenciosa en segon pla (a l'arrencada de cada pàgina): mai
  // mostra res ni llança. Amb el token de la sessió, de seguida i sense
  // Google. Resol a true si n'hi ha, false si no (llavors caldrà un toc).
  function reconnectarSilenci() {
    if (connectat()) return Promise.resolve(true);
    if (!disponible() || !jaAutoritzat()) return Promise.resolve(false);
    return _demanarToken("none").then(function () { return true; }).catch(function () { return false; });
  }

  function connectat() { return !!accessToken && Date.now() < caducaA - 60000; }

  function _check(resp) {
    // Un token caducat o retirat: s'oblida, i el proper toc en demana un altre.
    if (resp.status === 401) oblidarToken();
    if (!resp.ok) {
      return resp.text().then(function (t) {
        throw new Error("Drive " + resp.status + ": " + t);
      });
    }
    return resp;
  }

  // Cerca un fitxer per nom dins d'una carpeta. Retorna l'id o null.
  function cercarFitxer(nom, folderId) {
    var q = "name='" + nom.replace(/'/g, "\\'") + "' and trashed=false";
    if (folderId) q += " and '" + folderId + "' in parents";
    var url = "https://www.googleapis.com/drive/v3/files?orderBy=modifiedTime desc&pageSize=1&fields=files(id,name)&q=" +
      encodeURIComponent(q);
    return fetch(url, { headers: { Authorization: "Bearer " + accessToken } })
      .then(_check).then(function (r) { return r.json(); })
      .then(function (d) { return (d.files && d.files.length) ? d.files[0].id : null; });
  }

  // Llegeix un fitxer de la carpeta Dades (privada) com a TEXT. siNoHiEs: el
  // missatge si no hi es (diu que s'ha de fer al PC).
  function llegirTextDades(nom, siNoHiEs) {
    if (!connectat()) return Promise.reject(new Error("Cal connectar Drive primer."));
    return cercarFitxer(nom, CONFIG.DRIVE_DADES_FOLDER_ID).then(function (id) {
      if (!id) throw new Error(siNoHiEs || ("No s'ha trobat " + nom + " a Drive."));
      var url = "https://www.googleapis.com/drive/v3/files/" + id + "?alt=media";
      return fetch(url, { headers: { Authorization: "Bearer " + accessToken } })
        .then(_check).then(function (r) { return r.text(); });
    });
  }

  // Llegeix i parseja activitats.json de la carpeta Dades.
  function llegirActivitats() {
    return llegirTextDades("activitats.json", "No s'ha trobat activitats.json a Drive. Genera un informe al PC primer.")
      .then(function (t) { return JSON.parse(t); });
  }

  // El Planol activitats que el PC hi deixa (Dades/planol.html, PujaPlanol.ps1).
  function llegirPlanol() {
    return llegirTextDades("planol.html", "Encara no hi ha cap plànol al Drive: genera'l al PC (Plànol activitats) i torna-ho a provar.");
  }

  // La CCO del correu (Dades/correu-mobil.json): les adreces NO son al GitHub,
  // que es public; les hi deixa el PC (Configuracio correu electronic, fila
  // de l'app del mobil, i Actualitzar).
  // Torna { bcc: [{addr, def}] }.
  function llegirCorreuMobil() {
    return llegirTextDades("correu-mobil.json", "Encara no hi ha la CCO del correu al Drive: desa-la al PC (Configuració → Configuració correu electrònic → «Informe de visita · des del mòbil») o fes Actualitzar.")
      .then(function (t) { return JSON.parse(t); });
  }

  // Puja un objecte JSON com a fitxer nou a la carpeta Entrada.
  function pujarPaquet(nom, obj) {
    if (!connectat()) return Promise.reject(new Error("Cal connectar Drive primer."));
    var metadata = { name: nom, mimeType: "application/json" };
    if (CONFIG.DRIVE_ENTRADA_FOLDER_ID) metadata.parents = [CONFIG.DRIVE_ENTRADA_FOLDER_ID];
    var boundary = "-------informescornella" + Date.now();
    var body =
      "--" + boundary + "\r\n" +
      "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
      JSON.stringify(metadata) + "\r\n" +
      "--" + boundary + "\r\n" +
      "Content-Type: application/json\r\n\r\n" +
      JSON.stringify(obj, null, 2) + "\r\n" +
      "--" + boundary + "--";
    return fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", {
      method: "POST",
      headers: {
        Authorization: "Bearer " + accessToken,
        "Content-Type": "multipart/related; boundary=" + boundary
      },
      body: body
    }).then(_check).then(function (r) { return r.json(); });
  }

  window.Drive = {
    disponible: disponible,
    connectar: connectar,
    reconnectarSilenci: reconnectarSilenci,
    connectat: connectat,
    llegirCorreuMobil: llegirCorreuMobil,
    llegirActivitats: llegirActivitats,
    llegirPlanol: llegirPlanol,
    pujarPaquet: pujarPaquet
  };
})();
