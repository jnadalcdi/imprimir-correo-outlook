# Imprimir correo (complemento de Outlook)

Añade un botón **Imprimir correo** a Outlook. Sirve para imprimir correos que
con la impresión normal salen cortados: tablas anchas, imágenes grandes, etc.
El botón ajusta todo al ancho de la hoja y abre directamente el cuadro de
impresión.

## Cómo se usa

1. Abre el correo que quieres imprimir.
2. Pulsa el botón **Imprimir correo** en la cinta de opciones (grupo *Imprimir*).
3. Se abre la vista previa. Elige la impresora y pulsa **Imprimir**.

El papel lleva el asunto, el remitente, la fecha, los destinatarios, los
nombres de los adjuntos y el cuerpo del correo.

> El botón solo aparece al **leer** un correo, no al redactarlo.

## Cómo se instala

Solo hace falta un fichero: **`manifest-final.xml`**.

### Solo para ti (para probar)

1. En Outlook, pulsa **Obtener complementos** (o *Aplicaciones*).
2. Ve a **Mis complementos** → **Agregar un complemento personalizado** →
   **Agregar desde archivo**.
3. Elige `manifest-final.xml` y acepta.

### Para toda la empresa (administrador)

1. Entra en el [centro de administración de Microsoft 365](https://admin.microsoft.com).
2. Ve a **Configuración** → **Aplicaciones integradas** → **Cargar aplicaciones personalizadas**.
3. Sube `manifest-final.xml` y elige a qué usuarios se asigna.

El botón puede tardar unas horas en aparecerles a todos.

## Cómo se cambia algo

La web del complemento se publica sola con GitHub Pages en
<https://jnadalcdi.github.io/imprimir-correo-outlook/>.

- **Cambios en el código** (`.js`, `.html`, `.css`): haz commit y push. En
  unos minutos todos usan la versión nueva. No hay que volver a instalar nada.
- **Cambios en el manifiesto** (nombre, botón, iconos…):
  1. Edita `manifest.xml` y sube el número de `<Version>` (por ejemplo, de
     `1.0.2.0` a `1.0.3.0`).
  2. Genera el manifiesto final:

     ```powershell
     .\preparar.ps1 -Url https://jnadalcdi.github.io/imprimir-correo-outlook
     ```

  3. Haz commit y push, y vuelve a subir `manifest-final.xml` a Outlook o al
     centro de administración.
- **Iconos nuevos**: ponlos en una carpeta nueva (`iconos/v3/`) y cambia las
  rutas en `manifest.xml`. Si los guardas con el mismo nombre, Outlook sigue
  enseñando los antiguos, que tiene en caché.

## Qué hay en cada fichero

| Fichero | Para qué sirve |
| --- | --- |
| `manifest-final.xml` | El que se instala en Outlook. Se genera con `preparar.ps1`; no lo edites a mano. |
| `manifest.xml` | Plantilla del manifiesto (lleva `__BASE_URL__` en vez de la dirección). |
| `preparar.ps1` | Crea `manifest-final.xml` a partir de la plantilla. |
| `funciones.html` / `funciones.js` | Lo que se ejecuta al pulsar el botón. |
| `dialogo.html` / `dialogo.js` | La ventana donde sale la vista previa de impresión. |
| `comun.js` | Prepara el correo: cabecera, limpieza del HTML e imágenes. |
| `imprimir.css` | Hace que todo quepa en el ancho de la hoja. |
| `iconos/` | Iconos del botón. |

## Requisitos

Outlook de Microsoft 365, Outlook 2019 o posterior, el nuevo Outlook o Outlook
en la web. No funciona en el móvil.
