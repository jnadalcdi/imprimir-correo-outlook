/*
 * Código común del complemento "Imprimir correo":
 *   - prepararCorreo(item): cabecera + cuerpo HTML limpio (sin anchos fijos,
 *     con las imágenes insertadas incrustadas) listo para imprimir.
 *   - guardar/leer/borrar: pasan ese HTML de la función del botón a la
 *     ventana de impresión (localStorage, mismo origen).
 */
var Comun = (function () {
  "use strict";

  // ── Preparar el correo ──────────────────────────────────────────────────

  function prepararCorreo(item) {
    return new Promise(function (resolve, reject) {
      item.body.getAsync(Office.CoercionType.Html, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("No se ha podido leer el correo: " + res.error.message));
          return;
        }
        var doc = document.implementation.createHTMLDocument("correo");
        var cuerpo = doc.createElement("div");
        cuerpo.className = "correo";
        cuerpo.innerHTML = limpiarHtml(res.value);

        incrustarImagenes(item, cuerpo).then(function () {
          resolve(cabecera(doc, item) + cuerpo.outerHTML);
        });
      });
    });
  }

  function cabecera(doc, item) {
    var h1 = doc.createElement("h1");
    h1.id = "asunto";
    h1.textContent = item.subject || "(sin asunto)";

    var tabla = doc.createElement("table");
    tabla.id = "cabecera";
    [
      ["De:", persona(item.from)],
      ["Enviado:", item.dateTimeCreated ? item.dateTimeCreated.toLocaleString("es-ES") : ""],
      ["Para:", personas(item.to)],
      ["CC:", personas(item.cc)],
      ["Datos adjuntos:", (item.attachments || [])
        .filter(function (a) { return !a.isInline; })
        .map(function (a) { return a.name; })
        .join("; ")]
    ].forEach(function (f) {
      if (!f[1]) return;
      var tr = tabla.insertRow();
      var th = doc.createElement("th");
      th.textContent = f[0];
      tr.appendChild(th);
      tr.insertCell().textContent = f[1];
    });
    return h1.outerHTML + tabla.outerHTML;
  }

  function persona(p) {
    if (!p) return "";
    if (p.displayName && p.emailAddress && p.displayName !== p.emailAddress) {
      return p.displayName + " <" + p.emailAddress + ">";
    }
    return p.displayName || p.emailAddress || "";
  }

  function personas(lista) {
    return (lista || []).map(persona).join("; ");
  }

  // ── Limpieza del HTML ───────────────────────────────────────────────────

  function limpiarHtml(html) {
    var doc = new DOMParser().parseFromString(html, "text/html");

    // Nada ejecutable ni que cambie la página.
    doc.querySelectorAll("script, iframe, object, embed, base, meta, link, title")
      .forEach(function (el) { el.remove(); });

    doc.querySelectorAll("*").forEach(function (el) {
      // Atributos de eventos (onclick, onload...).
      Array.prototype.slice.call(el.attributes).forEach(function (a) {
        if (/^on/i.test(a.name)) el.removeAttribute(a.name);
      });

      var tag = el.tagName;
      if (tag === "TABLE" || tag === "TD" || tag === "TH" || tag === "COL" ||
          tag === "COLGROUP" || tag === "DIV" || tag === "P") {
        // Los anchos fijos en píxeles son los que hacen que la tabla se salga.
        el.removeAttribute("width");
        el.removeAttribute("nowrap");
        el.style.removeProperty("width");
        el.style.removeProperty("min-width");
        el.style.removeProperty("white-space");
      }
      if (tag === "IMG") {
        // Mantener la proporción al reducir.
        el.removeAttribute("height");
        el.style.removeProperty("height");
      }
    });

    // Los <style> del correo se conservan (colores, fuentes); imprimir.css
    // usa !important para imponer el ajuste de anchos.
    var estilos = Array.prototype.map.call(doc.querySelectorAll("style"),
      function (s) { return s.outerHTML; }).join("");
    return estilos + doc.body.innerHTML;
  }

  // ── Imágenes insertadas en el cuerpo (src="cid:...") ────────────────────

  function incrustarImagenes(item, contenedor) {
    var imgs = Array.prototype.filter.call(contenedor.querySelectorAll("img"),
      function (img) { return /^cid:/i.test(img.getAttribute("src") || ""); });
    var adjuntos = (item.attachments || []).filter(function (a) { return a.isInline; });

    if (imgs.length === 0 || adjuntos.length === 0 || !item.getAttachmentContentAsync) {
      return Promise.resolve();
    }

    return Promise.all(adjuntos.map(function (adj) {
      return new Promise(function (resolve) {
        item.getAttachmentContentAsync(adj.id, function (res) {
          if (res.status === Office.AsyncResultStatus.Succeeded &&
              res.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
            resolve({ adjunto: adj, base64: res.value.content });
          } else {
            resolve(null);
          }
        });
      });
    })).then(function (contenidos) {
      contenidos = contenidos.filter(Boolean);
      imgs.forEach(function (img, i) {
        // cid:image001.png@01DB1234.56789ABC  ->  image001.png
        var cid = img.getAttribute("src").substring(4);
        var nombre = cid.split("@")[0].toLowerCase();
        var match = contenidos.filter(function (c) {
          return (c.adjunto.name || "").toLowerCase() === nombre;
        })[0] || (imgs.length === contenidos.length ? contenidos[i] : null);
        if (match) {
          img.setAttribute("src", "data:" + tipoMime(match.adjunto) + ";base64," + match.base64);
        }
      });
    });
  }

  function tipoMime(adj) {
    if (adj.contentType && adj.contentType.indexOf("/") > 0) return adj.contentType;
    var ext = (adj.name || "").split(".").pop().toLowerCase();
    return { png: "image/png", gif: "image/gif", bmp: "image/bmp", jpg: "image/jpeg", jpeg: "image/jpeg" }[ext]
      || "image/png";
  }

  // ── Paso del HTML a la ventana de impresión (localStorage, mismo origen) ─
  // Si no cabe (límite ~5 MB con muchas imágenes), guardar() devuelve false y
  // la ventana lo pide por mensaje (plan B en funciones.js / dialogo.js).

  // Clave fija: así la ventana se abre con una URL limpia (sin parámetros),
  // que es lo que Outlook muestra en su barra de título.
  var CLAVE = "imprimir-correo";

  function guardar(html) {
    try {
      localStorage.setItem(CLAVE, html);
      return true;
    } catch (e) {
      return false;
    }
  }

  function leer() {
    try { return localStorage.getItem(CLAVE); } catch (e) { return null; }
  }

  function borrar() {
    try { localStorage.removeItem(CLAVE); } catch (e) { /* nada */ }
  }

  return { prepararCorreo: prepararCorreo, guardar: guardar, leer: leer, borrar: borrar };
})();
