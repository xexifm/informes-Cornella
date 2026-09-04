// Configuracio del formulari del mobil.
// -----------------------------------------------------------------------------
// Aquest fitxer NO conte dades personals: nomes identificadors de serveis i el
// correu del destinatari per defecte. Es pot pujar al GitHub public sense
// problema. Edita'l seguint la guia DESPLEGAMENT-MOBIL.md.

window.CONFIG = {
  // ----- Google Drive (auto-emplenar capçalera + enviar el paquet al PC) -----
  // ID de client OAuth de Google (tipus "Aplicacio web"). Buit = Drive
  // desactivat (capçalera a mà i botó "Baixar paquet").
  GOOGLE_CLIENT_ID: "464628466232-k2l3frmi7r5aav82psjrlfltv68bqvmt.apps.googleusercontent.com",
  DRIVE_ENTRADA_FOLDER_ID: "1EZteYidosVE8iJeli4tNdRg6WxzOk5fn",   // carpeta on el mobil deixa els paquets
  DRIVE_DADES_FOLDER_ID: "1bRof2NwgLYKpzvUD0aGximQzZQkv73od",     // carpeta on el PC deixa activitats.json
  // Aquesta nomes la fa servir el PC (el mobil no hi toca mai), pero viu aqui
  // perque els tres IDs siguin al MATEIX lloc: abans aquest era a
  // suport/config.ps1 i els altres dos hi eren TAMBE, o sigui que calia
  // escriure'ls dues vegades, en dos formats i dos blocs del manual.
  DRIVE_PROCESSATS_FOLDER_ID: "131iyeU0B9eWgDA8E6ncxixNFwx5Yz68X", // carpeta on el PC mou els paquets ja generats

  // ----- Correu (enviar els requeriments amb un sol clic, sense obrir res) ----
  // S'usa EmailJS (https://www.emailjs.com), gratuit per a poc volum. Crea un
  // compte, connecta-hi el teu correu (Gmail...) -> aquesta sera l'adreça DES DE
  // la qual s'envia-, crea una plantilla amb les variables {{to_email}},
  // {{subject}} i {{message}}, i posa aqui les tres claus. Si les deixes buides,
  // el boto fa el comportament antic (obre l'app de correu).
  EMAILJS_PUBLIC_KEY: "ySQ5MNJ4QdJNBIpoG",
  EMAILJS_SERVICE_ID: "service_xzjkbnq",
  EMAILJS_TEMPLATE_ID: "template_yxi61y7",

  // Nom del remitent (camp {{name}} de la plantilla d'EmailJS).
  EMAIL_FROM_NAME: "Ajuntament de Cornellà de Llobregat - Activitats",

  // Destinatari per defecte del correu (es pot canviar al formulari). Per
  // defecte ja s'agafa l'e-mail del titular; això és només un fallback.
  EMAIL_DESTINATARI: ""
};
