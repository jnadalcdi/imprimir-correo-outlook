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

### Solo para ti, desde el navegador (lo más rápido)

1. Abre este enlace e inicia sesión con tu cuenta del trabajo:
   <https://aka.ms/olksideload>
2. Se abre Outlook en la web con la ventana de complementos. Ve a
   **Mis complementos** → **Agregar un complemento personalizado** →
   **Agregar desde archivo**.
3. Elige `manifest-final.xml` y acepta.

Queda instalado en tu cuenta, así que el botón también aparece en Outlook de
escritorio (puede tardar un rato; si no sale, cierra y abre Outlook).

### Solo para ti, desde Outlook

1. En Outlook, pulsa **Obtener complementos** (o *Aplicaciones*).
2. Ve a **Mis complementos** → **Agregar un complemento personalizado** →
   **Agregar desde archivo**.
3. Elige `manifest-final.xml` y acepta.

### Para toda la empresa (administrador)

1. Entra en el [centro de administración de Microsoft 365](https://admin.microsoft.com).
2. Ve a **Configuración** → **Aplicaciones integradas** → **Cargar aplicaciones personalizadas**.
3. Sube `manifest-final.xml` y elige a qué usuarios se asigna.

El botón puede tardar unas horas en aparecerles a todos.

### Para un usuario concreto, con PowerShell (administrador)

Hace falta el módulo de Exchange Online (solo la primera vez):

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

Conéctate con tu cuenta de administrador e instala el complemento en el buzón
del usuario (ejecútalo desde la carpeta del repositorio):

```powershell
Connect-ExchangeOnline -UserPrincipalName admin@cdi-ibense.com

New-App -Mailbox usuario@cdi-ibense.com -FileData ([IO.File]::ReadAllBytes("$PWD\manifest-final.xml"))
```

Otros comandos útiles:

```powershell
# Ver si el usuario lo tiene instalado
Get-App -Mailbox usuario@cdi-ibense.com | Where-Object DisplayName -eq "Imprimir correo"

# Quitárselo
Remove-App -Mailbox usuario@cdi-ibense.com -Identity 365c5906-4b38-44fb-a694-3c6fed11b4bd -Confirm:$false

# Desconectarse al terminar
Disconnect-ExchangeOnline -Confirm:$false
```

Para actualizarlo tras cambiar el manifiesto, quítalo y vuelve a instalarlo.
El botón puede tardar un rato en aparecer; si no sale, cierra y abre Outlook.

## Si sale "No pudimos acceder a Imprimir correo"

Outlook no consigue cargar la página del complemento. Lo normal es que tenga
instalada una versión antigua que apunta a otra dirección (por ejemplo,
`localhost:3000`, la de pruebas).

1. Quita el complemento: **Obtener complementos** → **Mis complementos** →
   *Imprimir correo* → **Quitar**.
2. Cierra Outlook por completo.
3. Borra la caché de complementos: pulsa `Win + R`, escribe
   `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef` y borra todo lo que hay dentro.
4. Abre Outlook e instala otra vez `manifest-final.xml`.

Si sigue fallando, comprueba que desde ese equipo se abre
<https://jnadalcdi.github.io/imprimir-correo-outlook/funciones.html> en el
navegador (debe salir una página en blanco, sin error). Si no se abre, algo en
la red (proxy, cortafuegos) está bloqueando la dirección.

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
