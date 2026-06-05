// Enviament de correu amb un sol clic (sense obrir l'app de correu).
// -----------------------------------------------------------------------------
// Fa servir EmailJS, que envia el correu des del servidor del servei amb el teu
// compte connectat (l'adreça "des de la qual s'envia" es configura al panell
// d'EmailJS). Si no esta configurat, qui crida pot fer el fallback a mailto.
//
// Exposa window.Mail amb: configurat(), enviar(dest, assumpte, missatge).

(function () {
  var carregat = false;

  function configurat() {
    return !!(window.CONFIG && CONFIG.EMAILJS_PUBLIC_KEY && CONFIG.EMAILJS_SERVICE_ID && CONFIG.EMAILJS_TEMPLATE_ID);
  }

  function carregar() {
    return new Promise(function (res, rej) {
      if (carregat && window.emailjs) return res();
      var s = document.createElement("script");
      s.src = "https://cdn.jsdelivr.net/npm/@emailjs/browser@4/dist/email.min.js";
      s.onload = function () {
        try { emailjs.init({ publicKey: CONFIG.EMAILJS_PUBLIC_KEY }); carregat = true; res(); }
        catch (e) { rej(e); }
      };
      s.onerror = function () { rej(new Error("No s'ha pogut carregar EmailJS (xarxa?).")); };
      document.head.appendChild(s);
    });
  }

  function enviar(dest, assumpte, missatge) {
    if (!configurat()) return Promise.reject(new Error("EmailJS no configurat."));
    return carregar().then(function () {
      return emailjs.send(CONFIG.EMAILJS_SERVICE_ID, CONFIG.EMAILJS_TEMPLATE_ID, {
        to_email: dest,
        subject: assumpte,
        message: missatge
      });
    });
  }

  window.Mail = { configurat: configurat, enviar: enviar };
})();
