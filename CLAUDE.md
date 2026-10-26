# App de pedidos a proveedores · Buenas y Santas

Restaurante en Madrid. App para tablet con la que se hacen los pedidos a proveedores y se envían por WhatsApp. El usuario no es programador: explícale los pasos de forma sencilla y en español.

## Piezas

- **App web** (carpeta `SUBIR A GITHUB/`): se publica en GitHub Pages y se abre en la tablet como app instalada.
  - `index.html`: toda la app en un solo archivo (HTML, CSS y JS). **No lleva ninguna lista de productos**: todo viene de la hoja y se guarda en la tablet la última lista recibida (`pp.cat3`). El script tampoco lleva lista (se quitaron `CATALOGO_BASE` y `DATOS` el 18/10/2026 porque el usuario quiere que todo salga de la hoja). No volver a meterlas.
  - `sw.js`: guarda la app para que funcione sin wifi. **Sube el número de `CACHE` (`pedidos-vN`) en cada cambio**, o la tablet seguirá mostrando la versión vieja.
  - `manifest.webmanifest`, `icon-192.png`, `icon-512.png`: instalación en la tablet.
- **Google Sheet + Apps Script** (`PEGAR EN SCRIPT.txt` = `Codigo.gs`): la hoja es la base de datos.
  - URL del script publicado (ya está en index.html como `URL_HOJA`):
    `https://script.google.com/macros/s/AKfycbyMwaUKlKn46HqJtZ8vmceOulgwoXaF-qGCzHZp4w32uV2Spvk-4q-Cf2qY6u7B0bur/exec`
  - Sin clave (el usuario no quiere clave).
  - Peticiones POST con JSON `{accion: 'catalogo' | 'pedido' | 'historial'}`.
  - Pestañas: `APP_Proveedores` (Proveedor, Contacto, Teléfono, WhatsApp, Entregas, Nota, Activo), `APP_Productos` (ID, Proveedor, Sección, Producto, Unidad, Mínimo, Ref, Precio, Activo), `APP_Pedidos` (registro de todo lo enviado, nunca se borra). La pestaña `pedidos` es la hoja original del usuario; el script no la toca.
  - **El script lo sube Claude con clasp, no el usuario** (no quiere hacerlo él). Carpeta `SCRIPT GOOGLE/` (`.clasp.json`, `appsscript.json`, `Código.js`). La hoja es de **lfisbein@gmail.com** (la cuenta buenasysantasmadrid solo tiene lectura), así que se usa `--user lfisbein`:
    1. Copiar `Codigo.gs` a `SCRIPT GOOGLE/Código.js` (y a `PEGAR EN SCRIPT.txt` y `SUBIR A GITHUB/Codigo.gs`).
    2. En `SCRIPT GOOGLE/`: `clasp --user lfisbein push --force`
    3. `clasp --user lfisbein deploy -i AKfycbyMwaUKlKn46HqJtZ8vmceOulgwoXaF-qGCzHZp4w32uV2Spvk-4q-Cf2qY6u7B0bur -d "<descripción>"` (misma URL, versión nueva).
    4. Probar con un POST `{"accion":"catalogo"}` a la URL.
  - Script ID: `1iJHMjQFJDYV0GL9H-78uq5PyWHBWqB1bezQOm0lasuvBhSD37lznmBE-`. Hoja (Drive): `11X7L9CcmklQtHbqdCyi29O0QPzpeEUPAK-GH5c8Fr0I`.
- `pedido_proveedores2026.xlsx`: copia de la hoja de Google.
- La carpeta `github/` es una versión antigua; no usarla.

## Qué hace la app

- Inicio: lista de los 15 proveedores con su estado (pendiente, a medias o enviado) y el botón "Pedidos anteriores".
- Proveedor: productos con botones − / + y un teclado con ¼ ½ ¾. Incluye "Poner habituales", "Repetir último", "Solo lo pedido", "Vaciar" y pestañas por día de entrega (Iberolactal M/J, Disbesa M/V…).
- La columna `Habitual` de APP_Productos es en realidad el **mínimo** de stock: la app la muestra como "Mínimo X". El "Habitual" lo calcula la app (función `habituales()`): mediana de los últimos 8 pedidos de ese producto, con al menos 3 pedidos. "Poner habituales" solo usa ese cálculo.
- Orden de proveedores = orden de filas de APP_Proveedores (19/10/2026: Mercadona, Atlanta, Iberolactal, Disbesa, Monbake, Hughes, Envapro, el resto, y al final CBG, Bilandera, Duo Harinero). Berlys se renombró a **MONBAKE** en todas las pestañas APP_ (función `cambioUnaVez_ordenMonbake`, ya ejecutada; propiedad `orden_monbake`; se puede borrar).
- En el inicio solo sale el nombre del proveedor, su estado y la estimación de la compra: "Pedido ≈ X €" si hay algo puesto hoy, si no "Normalmente ≈ X €" (`gastoHabitual()`: mediana de los últimos 8 pedidos, mínimo 2, con cantidades de entonces × precios de hoy). Arriba, la pastilla "Hoy ≈ X €". Sin contacto ni entregas: eso se ve al entrar. Hay un bloque CSS `@media (max-width:560px)` al final del `<style>` con letra y botones más pequeños para móvil.
- La barra de buscar y la barra de "Revisar y enviar" no son fijas: van arriba y al final de la lista (el usuario no quiere que tapen productos).
- "Otros productos": al final de cada proveedor hay líneas libres (una sola al principio y un botón "+ Añadir otro producto" que añade más; `ui.nOtros`) para escribir productos que no están en la hoja. Claves `x1, x2…`: texto en `d.otros`, cantidad en `d.q`. Se envían y registran con ID vacío; "Repetir" las recupera. Funciones `lineas()`, `huecosOtros()`, `ponerOtro()`.
- **Precios** (19/10/2026): salen de las facturas escaneadas en la app de **Cierre de caja** (`Desktop/CIERRE-CAJA`, botón "Escanear"). Tras escanear, esa app manda la foto en segundo plano a `accionPreciosFactura` (`CIERRE-CAJA/Precios.gs`): guarda la factura en PDF en Drive (`Facturas proveedores/<AAAA>/<MM>/<proveedor>`, carpeta de buenasysantasmadrid compartida con lfisbein; los súper —Carrefour, Dia…— van a `Super` con su nombre en el archivo), Claude lee las líneas y las empareja con los ID de pedidos, y se mandan a este script (`accion: 'precios'` → pestaña `APP_Precios`, columna Precio de APP_Productos). Precios sin IVA. La app muestra precio por producto, ↑/↓ % (último mes), importe aproximado del pedido (no va en el WhatsApp), gasto en el historial (cantidades × precio actual) y el aviso "N productos han cambiado de precio" (`cat.cambios`, vistos en `pp.cambiosVistos`). El cierre se publica SOLO con `CIERRE-CAJA/.claude/publicar.ps1` (exige clasp con la cuenta lfisbein: con otra cuenta Contabilidad y Albaranes dejan de guardarse). Otra sesión de Claude trabaja también en CIERRE-CAJA: mirar `git log` y `git status` allí antes de tocar nada.
- Inicio: cada proveedor tiene un botón ✓ (`marcarEnviado`): con productos registra el pedido en la hoja como "Marcar como enviado"; sin productos solo lo marca; otro toque lo desmarca (lo ya apuntado en la hoja se queda). Botón "Imprimir semana" (`imprimirSemana`, `#print` + `@media print`): pedidos de lunes a domingo por proveedor, con lo de hoy sin enviar marcado.
- Barra de abajo del proveedor: "≈ X €" (+ aviso de pedido mínimo) y un único botón: **Enviar por WhatsApp** (abre wa.me directamente y registra) o, sin WhatsApp, "Enviar pedido" (ventana de copiar/marcar). La ventana `revisar()` ya solo se usa para proveedores sin WhatsApp.
- **Leer precios de una factura** (20/10/2026): botón al final de cada proveedor. Llama directamente al script del Cierre de caja (`URL_CIERRE` en index.html; acciones sin sesión `accionFacturasPedidos` y `accionPreciosDesdeDrive`, en `CIERRE-CAJA/Precios.gs`, máx. 40 relecturas al día): lista las facturas de Drive de los últimos 2 meses de ese proveedor (o de todos) y marca las que ya tienen precios (este script: `accion:'facturasConPrecios'`, por ID del PDF o nº de factura). "Volver a leer" manda `reemplazar:true` y `guardarPrecios` quita antes las líneas de esa factura. Este script NO usa UrlFetchApp ni Drive (añadir servicios pediría volver a autorizar y la app dejaría de funcionar hasta hacerlo).
- `APP_Proveedores` tiene columna **Pedido mínimo** (€, sin IVA; DUO HARINERO = 80). `APP_Pedidos` tiene columnas K "Precio aprox." y L "Importe aprox." (la app manda `precio` en cada línea).
- (Antes) "Revisar y enviar": abre WhatsApp con el mensaje ya escrito y apunta el pedido en APP_Pedidos. Si no hay conexión, lo guarda y lo sube más tarde.
- Pedidos anteriores: últimos 6 meses, con filtros por periodo, proveedor y producto. Se puede ver por pedidos, por productos o por proveedores, repetir un pedido y descargar un CSV.
- Proveedores y productos se cambian en la hoja, solo en APP_Proveedores y APP_Productos (el usuario lo eligió así; las pestañas `pedidos` y `cocina` no se leen). Ambas tienen columna Activo (SI/NO). La app los recoge sola al abrirse; el orden de los productos es el de las filas.
- El script (`ponerAlDia()`) añade la columna Activo a APP_Proveedores si falta, renombra "Habitual" a "Mínimo" y pone ID (p###) a productos sin ID o con ID repetido.
- Los proveedores de la pestaña `cocina` (FRUTAPRO, JAVIER DIAZ, LAS 3 B, MERCANADALUCIA, LA JAULA) no están en la app: el usuario dijo que no los añadiera (17/10/2026).

## Estilo elegido: "Olivo · tarjetas"

Barra superior verde oliva `#556b2f`, fondo `#f4f5ef` y tarjetas blancas redondeadas con borde fino. Nombres en Playfair Display cursiva y textos en Figtree. El usuario descartó muchos estilos antes de elegir este: no cambiar la estética si no lo pide.

## Pendiente / ideas

- GitHub ya está conectado: `SUBIR A GITHUB/` es un repositorio Git enlazado a `https://github.com/buenasysantasmadrid/pedidos` (rama `main`). La cuenta `buenasysantasmadrid` está guardada en el Administrador de credenciales de Windows. Para subir cambios: `git add -A`, `git commit`, `git push` dentro de esa carpeta (no hay `gh` instalado).
- El script también está en el repositorio como `SUBIR A GITHUB/Codigo.gs` (copia de `Codigo.gs` / `PEGAR EN SCRIPT.txt`). Si cambias el script, actualiza las tres copias y súbelo.
- Posible pantalla para editar productos desde la app.
- WINEUP tiene el teléfono incompleto en la hoja.

## Email de compras de los viernes (21/10/2026)

Todos los viernes a las 18:00 (Madrid) llega a **info@buenasysantas.es** un email con: 1) cambios de precio (APP_Precios), 2) productos pedidos más que su media de 8 semanas (APP_Pedidos), 3) gasto por proveedor frente a su media (planilla de **Albaranes** del cierre, con IVA), 4) pedidos de hace 1-14 días sin factura escaneada (Drive + APP_Precios) y 5) conclusiones de la IA. Está en el script del **Cierre de caja** (`CIERRE-CAJA/InformeSemanalPedidos.gs`), que ya tenía permiso para mandar emails y usar Claude. El activador es de la cuenta lfisbein. Acciones (POST con `claveClaude` de `ClaveClaude.gs`): `accionInformeSemanal: 'vista' | 'enviar' | 'instalar'`. Publicar siempre con `CIERRE-CAJA/.claude/publicar.ps1`.
