# Despensa — Arquitectura (rev. 3, 2026-07-24)

PWA de gestión del contenido de la despensa, derivada de la app de Tuppers
(rev. 3). Mismo patrón: vanilla JS/HTML, MSAL, Microsoft Graph, CSV en
OneDrive, despliegue en GitHub Pages. Sufijo `_despensa` en todo lo nuevo
para evitar colisiones con Tuppers.

## Despliegue

- Repo GitHub: `atxcclt` (organización `atxcclt-26`), carpeta `Despensa/`
  — igual que Tuppers, no un repo aparte.
- URL publicada: `https://atxcclt-26.github.io/atxcclt/Despensa/`.
  La carpeta contiene un `index.html` reenviador que conserva query y hash
  y salta a `index_despensa.html` (donde vive la app). Así la URL corta
  funciona como página de uso, de instalación de la PWA y como redirect de
  MSAL.
- App registration Azure: la misma que Tuppers / Revisión de cargos
  (`clientId 24e9d6d3-d9ad-437e-b7f6-1a27f48c2696`). El redirect URI SPA
  `https://atxcclt-26.github.io/atxcclt/Despensa/` ya está registrado;
  `REDIRECT_PAGE` en `app_despensa.js` coincide con él.
- Scopes: `User.Read`, `Files.ReadWrite` (delegados, sin cambios).

## Datos

- Carpeta OneDrive: `ZZZ_SG_Despensa` (en la raíz del OneDrive; se admite
  también carpeta compartida vía `remoteItem`, con enrutado por `driveId`
  como en Tuppers).
- Fichero: `despensa.csv`, separador `;`, UTF-8 con BOM, CRLF.
- Columnas:

  | Columna | Valores | Notas |
  |---|---|---|
  | Nombre | texto libre | obligatorio |
  | Cantidad | texto libre (`30 gr`, `2 bricks`…) | ver semántica abajo |
  | Casa | Comarruga, Castefa | |
  | Ubicación | Frigo, Congelador, Armario, Arriba | |
  | Fecha caducidad | `YYYY-MM-DD` (admite `DD/MM/YYYY` en lectura) | opcional |
  | Corto plazo abierto | Sí / No | check "una vez abierto caduca pronto" |
  | Categoría | Cesped, Frutos secos, Harinas, Otros | |

- **Semántica de Cantidad** (clave del modelo, rev. 2):
  - **Vacía** → plantilla del catálogo: NO se muestra en el listado y se
    propone al pulsar "Nuevo artículo".
  - **`0`** → "No queda nada": visible en el listado (recordatorio de
    compra), tarjeta atenuada con borde gris, sin badges de ubicación ni
    caducidad.
  - **Cualquier otro texto** → artículo en stock normal.

- Cabecera tolerante a alias y a acentos (`crearMapaCabecera`), igual que
  en Tuppers. Sin columna Historial y sin fotos.
- Escritura: la app reescribe el CSV completo (`PUT …:/content`) tras cada
  alta, edición o borrado, ordenado por `_indiceOriginal` para preservar el
  orden del fichero. Rollback en memoria si el PUT falla.

## Funcionalidad

- **Alta**: "＋ Nuevo artículo" abre primero el selector de plantillas
  (`#selectorNuevo`) con los registros de Cantidad vacía, más el botón
  "Crear desde cero". Elegir una plantilla abre el editor sobre ese
  registro (rellenar cantidad = reponer). Si no hay plantillas, va directo
  al editor en blanco.
- **Eliminación en dos pasos**: el botón Eliminar del editor abre
  `#dialogoEliminar` con dos opciones:
  - *No queda nada*: conserva la fila con `Cantidad = 0` y vacía Ubicación
    y Fecha caducidad (rollback en memoria si falla el PUT).
  - *Eliminar completamente*: borra la fila del CSV.
- **Sin resumen** superior.
- **Botones rápidos** en dos grupos, OR dentro de grupo y AND entre grupos:
  - Casa: Castefa, Comarruga.
  - Categoría: Cesped, Frutos secos, Harinas, Otros. La categoría **no se
    muestra en las tarjetas**: solo se filtra desde aquí (o desde su combo).
- **Panel de filtros** plegable (Mostrar/Ocultar filtros) con combos de
  multiselección por todos los campos: OR dentro del campo, AND entre
  campos, opción `(vacío)`. Idéntico mecanismo que Tuppers.
- **Tarjetas**: nombre grande + badges de cantidad, casa, ubicación,
  caducidad y aviso `⚠ Corto plazo abierto`. La categoría no aparece.
  - Fondo: matiz por casa (Castefa azul, Comarruga verde), tono por
    ubicación (oscuro congelador, medio frigo, claro armario/arriba).
  - Borde izquierdo por estado: rojo caducado, gris "No queda nada"
    (tarjeta atenuada), verde en el resto. Sin aviso previo de caducidad.
- **Orden del listado**: casa → ubicación → fecha de caducidad ascendente
  (sin fecha al final) → nombre. Los "No queda nada" quedan agrupados por
  casa con ubicación vacía (al final de su casa).

## Ficheros

- `index.html` — reenviador mínimo a `index_despensa.html` conservando
  query y hash (necesario para el redirect de MSAL sobre la URL de carpeta).
- `index_despensa.html` — vista única + editor, CSS embebido. Enlaza a
  `manifest_despensa.json` y carga `app_despensa.js`.
- `app_despensa.js` — toda la lógica (`window.despensaAuth` para depuración).
- `manifest_despensa.json` — PWA. El nombre del manifest es libre: el
  navegador usa el que indique `<link rel="manifest">`. Su `start_url`
  apunta a `./index_despensa.html`.
- `despensa.csv` — fichero de ejemplo para inicializar `ZZZ_SG_Despensa`.
- Iconos pendientes: `icon-180_despensa.png`, `icon-512_despensa.png`,
  `favicon-32_despensa.png`.

## Puesta en marcha

1. Crear la carpeta `ZZZ_SG_Despensa` en la raíz de OneDrive y subir
   `despensa.csv` (el de ejemplo o uno vacío con solo la cabecera).
2. Crear el repo `Despensa` en `atxcclt-26`, subir los ficheros y activar
   GitHub Pages (rama main, raíz).
3. Añadir el redirect URI SPA en Azure (ver arriba).
4. Generar los iconos `icon-180_despensa.png`, `icon-512_despensa.png` y
   `favicon-32_despensa.png`.
