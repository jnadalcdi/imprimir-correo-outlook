/*
 * Función del botón "Imprimir correo" (sin panel):
 *   1. Prepara el correo (Comun.prepararCorreo).
 *   2. Abre una ventana grande de Office con dialogo.html, que muestra
 *      directamente el cuadro de impresión a tamaño normal.
 *   3. Cuando la ventana avisa de que ha terminado (o el usuario la cierra),
 *      la cierra y da la función por completada.
 */
(function () {
  "use strict";

  function imprimirCorreo(event) {
    var item = Office.context.mailbox.item;
    var html = null;
    var terminado = false;

    function terminar() {
      if (terminado) return;
      terminado = true;
      Comun.borrar();
      event.completed();
    }

    Comun.prepararCorreo(item)
      .then(function (resultado) {
        html = resultado;
        // Si no cabe en localStorage, la ventana pedirá el HTML por mensaje.
        Comun.guardar(html);
      })
      .then(function () {
        // URL sin parámetros: es lo que Outlook enseña en el título de la ventana.
        var url = new URL("dialogo.html", window.location.href).href;
        Office.context.ui.displayDialogAsync(url, { height: 90, width: 70, displayInIframe: true },
          function (res) {
            if (res.status !== Office.AsyncResultStatus.Succeeded) {
              avisar(item, "No se ha podido abrir la ventana de impresión: " + res.error.message);
              terminar();
              return;
            }
            var ventana = res.value;
            ventana.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
              if (arg.message === "pedir") {
                try { ventana.messageChild(html); } catch (e) { /* DialogApi 1.2 no disponible */ }
                return;
              }
              // "hecho" o "error:..."
              if (arg.message && arg.message.indexOf("error:") === 0) {
                avisar(item, arg.message.substring(6));
              }
              ventana.close();
              terminar();
            });
            // El usuario ha cerrado la ventana con la X.
            ventana.addEventHandler(Office.EventType.DialogEventReceived, terminar);
          });
      })
      .catch(function (e) {
        avisar(item, e.message);
        terminar();
      });
  }

  // Aviso en la barra amarilla del correo.
  function avisar(item, texto) {
    try {
      item.notificationMessages.replaceAsync("imprimirCorreo", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: String(texto).substring(0, 150)
      });
    } catch (e) { /* sin avisos */ }
  }

  Office.onReady(function () {
    if (Office.actions && Office.actions.associate) {
      Office.actions.associate("imprimirCorreo", imprimirCorreo);
    }
  });
  // Versiones de Outlook que buscan la función por nombre global.
  window.imprimirCorreo = imprimirCorreo;
})();
