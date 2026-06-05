// Configuracio del formulari del mobil.
// -----------------------------------------------------------------------------
// Aquest fitxer NO conte dades personals: nomes identificadors de Google i el
// correu del destinatari per defecte. Es pot pujar al GitHub public sense
// problema. Edita'l amb les teves dades seguint la guia DESPLEGAMENT-MOBIL.md.
//
// El formulari FUNCIONA encara que deixis tot aixo buit: en aquest cas no es
// connectara a Google Drive (no auto-emplenara la capcalera ni pujara el
// paquet sol), pero podras omplir la capcalera a ma i DESCARREGAR el paquet
// per deixar-lo tu a la carpeta de Drive. Omple-ho per tenir-ho automatic.

window.CONFIG = {
  // ID de client OAuth de Google (tipus "Aplicacio web"), creat a Google Cloud
  // Console. Acaba en ".apps.googleusercontent.com". Buit = Drive desactivat.
  GOOGLE_CLIENT_ID: "",

  // ID de la carpeta de Drive "Entrada" (on el mobil deixa els paquets i el
  // vigilant del PC els recull). El treus de la URL de la carpeta a Drive:
  // https://drive.google.com/drive/folders/<AIXO_ES_L_ID>
  DRIVE_ENTRADA_FOLDER_ID: "",

  // ID de la carpeta de Drive "Dades" (on el PC exporta activitats.json amb la
  // base de dades per auto-emplenar la capcalera). Mateixa manera de treure'l.
  DRIVE_DADES_FOLDER_ID: "",

  // Destinatari per defecte del correu amb els requeriments (es pot canviar al
  // formulari abans d'enviar). Pots deixar-lo buit.
  EMAIL_DESTINATARI: "",

  // Prefix de l'assumpte del correu. Se li afegeix l'ID GIA.
  EMAIL_ASSUMPTE_PREFIX: "Requeriments activitat GIA"
};
