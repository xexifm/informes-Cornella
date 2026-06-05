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

  // Demana (interactivament) un token d'accés. Cal una acció de l'usuari.
  function connectar() {
    if (!disponible()) return Promise.reject(new Error("Google Drive no configurat."));
    return carregarGis().then(function () {
      return new Promise(function (resolve, reject) {
        try {
          if (!tokenClient) {
            tokenClient = google.accounts.oauth2.initTokenClient({
              client_id: CONFIG.GOOGLE_CLIENT_ID,
              scope: SCOPE,
              callback: function (resp) {
                if (resp && resp.access_token) {
                  accessToken = resp.access_token;
                  resolve(accessToken);
                } else {
                  reject(new Error("No s'ha obtingut el token de Google."));
                }
              }
            });
          } else {
            tokenClient.callback = function (resp) {
              if (resp && resp.access_token) { accessToken = resp.access_token; resolve(accessToken); }
              else { reject(new Error("No s'ha obtingut el token de Google.")); }
            };
          }
          tokenClient.requestAccessToken({ prompt: accessToken ? "" : "consent" });
        } catch (e) { reject(e); }
      });
    });
  }

  function connectat() { return !!accessToken; }

  function _check(resp) {
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

  // Llegeix i parseja activitats.json de la carpeta Dades.
  function llegirActivitats() {
    if (!connectat()) return Promise.reject(new Error("Cal connectar Drive primer."));
    return cercarFitxer("activitats.json", CONFIG.DRIVE_DADES_FOLDER_ID).then(function (id) {
      if (!id) throw new Error("No s'ha trobat activitats.json a Drive. Genera un informe al PC primer.");
      var url = "https://www.googleapis.com/drive/v3/files/" + id + "?alt=media";
      return fetch(url, { headers: { Authorization: "Bearer " + accessToken } })
        .then(_check).then(function (r) { return r.json(); });
    });
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
    connectat: connectat,
    llegirActivitats: llegirActivitats,
    pujarPaquet: pujarPaquet
  };
})();
