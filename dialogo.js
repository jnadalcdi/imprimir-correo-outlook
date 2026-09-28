/*
 * Ventana de impresión: recibe el correo ya preparado, abre la vista previa
 * de impresión (que ocupa toda la ventana) y avisa a la función del botón
 * para que la cierre al terminar.
 */
(function () {
  "use strict";

  var enviado = false;

  Office.onReady(function () {
    var html = Comun.leer();
    if (html) {
      mostrarEImprimir(html);
      return;
    }
    // Plan B: pedir el HTML a la función del botón por mensaje.
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived,
      function (arg) { mostrarEImprimir(arg.message); });
    Office.context.ui.messageParent("pedir");
  });

  function mostrarEImprimir(html) {
    var documento = document.getElementById("documento");
    if (enviado || documento.innerHTML) return;
    documento.innerHTML = html;
    esperarImagenes(documento).then(function () {
      window.addEventListener("afterprint", cerrarCuandoTermine, { once: true });
      window.print();
    });
  }

  // "afterprint" llega al cerrarse la vista previa. Si el usuario eligió el
  // cuadro de diálogo del sistema, este se abre justo después y cerrar la
  // ventana lo cortaría: se espera a que la ventana recupere el foco.
  function cerrarCuandoTermine() {
    var conFoco = 0;
    setTimeout(function () {
      var timer = setInterval(function () {
        conFoco = document.hasFocus() ? conFoco + 1 : 0;
        if (conFoco >= 2) {
          clearInterval(timer);
          avisar("hecho");
        }
      }, 200);
    }, 300);
  }

  function avisar(mensaje) {
    if (enviado) return;
    enviado = true;
    Office.context.ui.messageParent(mensaje);
  }

  // Imprimir antes de que carguen las imágenes las dejaría en blanco.
  function esperarImagenes(contenedor) {
    var pendientes = Array.prototype.filter.call(contenedor.querySelectorAll("img"),
      function (img) { return !img.complete; });
    var todas = Promise.all(pendientes.map(function (img) {
      return new Promise(function (resolve) {
        img.addEventListener("load", resolve, { once: true });
        img.addEventListener("error", resolve, { once: true });
      });
    }));
    var limite = new Promise(function (resolve) { setTimeout(resolve, 5000); });
    return Promise.race([todas, limite]);
  }
})();
