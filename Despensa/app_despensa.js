// app.js — Gestión de despensa sobre OneDrive (rev. 1)
//
// Estructura esperada en OneDrive:
//   ZZZ_SG_Despensa/
//     despensa.csv
//     fotos/                       ← opcional
//       almendras.jpg
//
// Columnas de despensa.csv (separador ";"):
// Nombre;Cantidad;Casa;Ubicación;Fecha caducidad;Corto plazo abierto;
// Categoría;Foto;Fecha último cambio;Historial
//
// Todos los campos son editables desde la web (Historial es interno). La app
// permite crear y eliminar artículos.
//
// Rev. 5 (2026-07-24):
// - La foto se gestiona desde el editor (a diferencia de Tuppers): botón
//   "Subir foto" (galería/cámara, redimensionado a máx. 1600 px JPEG antes
//   de subir a "fotos/" vía Graph) y botón "Quitar foto". La subida se hace
//   al Guardar; Cancelar no deja ficheros huérfanos. Al sustituir o quitar
//   una foto propia (ruta relativa), el fichero antiguo se borra de OneDrive.
// - Los artículos "No queda nada" (Cantidad 0) ya NO se muestran en la
//   pantalla principal: se proponen en el selector de "Nuevo artículo"
//   junto a las plantillas sin cantidad.
//
// Rev. 4 (2026-07-24):
// - Foto como en Tuppers: columna Foto (nombre de fichero, ruta relativa o
//   URL), miniatura en tarjeta vía Graph con caché y fallback a "fotos/",
//   clic para abrir la imagen completa.
// - Nuevo campo Fecha último cambio: al abrir el editor se propone hoy;
//   Cancelar conserva la fecha que hubiera.
// - Historial de los últimos 5 estados (columna JSON interna, como en
//   Tuppers pero con 5 en vez de 3): instantánea del estado anterior en
//   cada guardado con cambios y en "No queda nada". Panel plegable en el
//   editor para consultarlos.
//
// Rev. 3 (2026-07-24):
// - Ficheros renombrados con sufijo "_despensa": index_despensa.html,
//   app_despensa.js, manifest_despensa.json.
// - Despliegue como carpeta Despensa/ dentro del repo atxcclt (igual que
//   Tuppers). REDIRECT_PAGE = https://atxcclt-26.github.io/atxcclt/Despensa/
//   (URI ya registrado en Azure). Un index.html reenviador conserva el hash
//   de la respuesta MSAL y salta a index_despensa.html.
//
// Rev. 2 (2026-07-24):
// - Nuevo campo Cantidad (texto libre, p. ej. "30 gr").
//   · Cantidad vacía  → el artículo NO se muestra: es una plantilla del
//     catálogo, y se propone al pulsar "Nuevo artículo".
//   · Cantidad "0"    → "No queda nada": visible en el listado (recordatorio
//     de compra), atenuado y con borde gris.
//   · Resto           → artículo en stock normal.
// - "Eliminar" ofrece dos opciones: "No queda nada" (conserva la fila con
//   Cantidad 0 y vacía Ubicación y Fecha caducidad) o "Eliminar
//   completamente" (borra la fila del CSV).
// - Eliminado el aviso de "caduca pronto" a 30 días: solo se marca en rojo
//   lo ya caducado.
//
// Rev. 1 (2026-07-24):
// - Derivada de la app de Tuppers rev. 3 con sufijo "_despensa" en carpeta,
//   fichero y nombres públicos para evitar colisiones.
// - Sin fotos, sin historial y sin resumen superior.
// - Botones rápidos en dos grupos: casa (Castefa/Comarruga) y categoría
//   (Cesped/Frutos secos/Harinas/Otros). OR dentro de grupo, AND entre grupos.
// - La categoría no se muestra en el listado: solo se filtra con los botones
//   rápidos (y con su combo del panel de filtros).
// - Panel de filtros plegable con combos de multiselección (OR dentro del
//   campo, AND entre campos), con opción "(vacío)".
// - Fondo de tarjeta: matiz por casa, tono por ubicación (oscuro congelador,
//   medio frigo, claro armario/arriba). Borde izquierdo por urgencia:
//   rojo caducado, naranja caduca en ≤ 30 días, verde resto.
// - Orden: casa → ubicación → fecha de caducidad ascendente → nombre.

// ====== Ajustes ======
const CARPETA_RAIZ = "ZZZ_SG_Despensa";
const FILE_DATOS = "despensa.csv";
const GRAPH = "https://graph.microsoft.com/v1.0";
// =====================

// Misma arquitectura de autenticación MSAL que Tuppers: el redirect URI es
// la URL de la carpeta dentro del repo atxcclt, ya registrada en Azure.
// index.html (reenviador) conserva el hash y salta a index_despensa.html.
const REDIRECT_PAGE = "https://atxcclt-26.github.io/atxcclt/Despensa/";
const msalConfig = {
  auth: {
    clientId: "24e9d6d3-d9ad-437e-b7f6-1a27f48c2696",
    redirectUri: REDIRECT_PAGE
  },
  cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
  system: {
    loggerOptions: {
      logLevel: msal.LogLevel.Warning,
      loggerCallback: (nivel, mensaje, contieneDatosPersonales) => {
        if (!contieneDatosPersonales) console.log(`[MSAL][${nivel}]`, mensaje);
      }
    }
  }
};
const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["User.Read", "Files.ReadWrite"] };
const fileScopes = { scopes: ["Files.ReadWrite"] };

const COLUMNAS = [
  { key: "nombre", label: "Nombre", aliases: ["nombre", "nombre articulo", "articulo"] },
  { key: "cantidad", label: "Cantidad", aliases: ["cantidad", "qty", "existencias"] },
  { key: "casa", label: "Casa", aliases: ["casa", "en q casa esta", "en que casa esta"] },
  { key: "ubicacion", label: "Ubicación", aliases: ["ubicacion"] },
  { key: "fechaCaducidad", label: "Fecha caducidad", aliases: ["fecha caducidad", "caducidad", "fecha de caducidad"] },
  { key: "cortoPlazo", label: "Corto plazo abierto", aliases: ["corto plazo abierto", "corto plazo", "caduca abierto"] },
  { key: "categoria", label: "Categoría", aliases: ["categoria"] },
  { key: "foto", label: "Foto", aliases: ["foto", "foto articulo", "imagen"] },
  { key: "fechaUltimoCambio", label: "Fecha último cambio", aliases: ["fecha ultimo cambio", "ultimo cambio", "fecha cambio"] },
  { key: "historial", label: "Historial", interno: true, aliases: ["historial"] }
];

const MAX_HISTORIAL = 5;
const CAMPOS_HISTORIAL = ["cantidad", "casa", "ubicacion", "fechaCaducidad", "cortoPlazo", "categoria", "fechaUltimoCambio"];

const CASAS = ["Comarruga", "Castefa"];
const UBICACIONES = ["Frigo", "Congelador", "Armario", "Arriba"];
const CATEGORIAS = ["Cesped", "Frutos secos", "Harinas", "Otros"];
const ORDEN_CASAS = new Map([["comarruga", 0], ["castefa", 1]]);
const ORDEN_UBICACIONES = new Map([["frigo", 0], ["congelador", 1], ["armario", 2], ["arriba", 3], ["", 4]]);
const collator = new Intl.Collator("es", { numeric: true, sensitivity: "base" });

const G = {
  driveId: null,
  folderId: null,
  registros: [],
  fotoCache: new Map(),
  cargaId: 0
};

let inicializado = false;
let guardando = false;

// ---------- Utilidades ----------
function $(id) { return document.getElementById(id); }
function setEstado(mensaje, error = false) {
  const el = $("estado-carga");
  if (!el) return;
  el.textContent = mensaje;
  el.classList.toggle("error", Boolean(error));
}
function mostrarVista(id) {
  document.querySelectorAll(".vista").forEach(v => v.classList.toggle("activa", v.id === id));
}
function escapeHtml(valor) {
  return String(valor == null ? "" : valor)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
function quitarAcentos(valor) {
  return String(valor == null ? "" : valor).normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}
function norm(valor) { return quitarAcentos(valor).trim().toLowerCase(); }
function normalizarCabecera(valor) { return norm(valor).replace(/[^a-z0-9]+/g, " ").trim(); }
function esSi(valor) { return ["si", "sí", "yes", "1", "true", "x"].includes(norm(valor)); }
function canonSiNo(valor, vacioPermitido = true) {
  const n = norm(valor);
  if (!n) return vacioPermitido ? "" : "No";
  if (["si", "sí", "yes", "1", "true", "x"].includes(n)) return "Sí";
  if (["no", "0", "false"].includes(n)) return "No";
  return String(valor).trim();
}
function valorVisible(valor, alternativa = "—") {
  const v = String(valor == null ? "" : valor).trim();
  return v || alternativa;
}
function encodeGraphPath(ruta) {
  return String(ruta || "").split("/").filter(Boolean).map(encodeURIComponent).join("/");
}
function fechaParaInput(valor) {
  const v = String(valor || "").trim();
  if (!v) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
  const m = v.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  return "";
}
function fechaVisible(valor) {
  const iso = fechaParaInput(valor);
  if (!iso) return valorVisible(valor);
  const [a, m, d] = iso.split("-");
  return `${d}/${m}/${a}`;
}
function hoyIso() {
  const d = new Date();
  const p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}
function normalizarRutaFoto(ruta) {
  return String(ruta || "").trim().replace(/^\/+/, "");
}
function esUrl(ruta) { return /^(https?:|data:|blob:)/i.test(String(ruta || "").trim()); }

// Cantidad: "" = plantilla oculta del catálogo; "0" = agotado ("No queda
// nada"); cualquier otro texto = en stock. En la pantalla principal solo se
// muestran los que están en stock; vacíos y agotados se proponen al crear.
function cantidadDe(registro) { return String(registro.cantidad || "").trim(); }
function tieneCantidad(registro) { return cantidadDe(registro) !== ""; }
function esAgotado(registro) { return cantidadDe(registro) === "0"; }
function esVisibleEnListado(registro) { return tieneCantidad(registro) && !esAgotado(registro); }

// Historial: JSON con los últimos MAX_HISTORIAL estados anteriores.
function parsearHistorial(valor) {
  if (!valor) return [];
  try {
    const lista = JSON.parse(valor);
    return Array.isArray(lista) ? lista.filter(v => v && typeof v === "object").slice(0, MAX_HISTORIAL) : [];
  } catch (error) {
    console.warn("Historial no válido; se ignora.", error);
    return [];
  }
}
function serializarHistorial(lista) {
  return JSON.stringify((Array.isArray(lista) ? lista : []).slice(0, MAX_HISTORIAL));
}
function crearInstantanea(registro) {
  const uso = { guardadoEn: new Date().toISOString() };
  CAMPOS_HISTORIAL.forEach(key => { uso[key] = registro[key] || ""; });
  return uso;
}
// Añade al historial del registro una instantánea de su estado actual.
function apilarHistorial(registro) {
  registro.historial = serializarHistorial([crearInstantanea(registro), ...parsearHistorial(registro.historial)].slice(0, MAX_HISTORIAL));
}

// Días desde hoy hasta la caducidad. null si no hay fecha válida.
// Negativo = caducado, 0 = caduca hoy.
function diasHastaCaducidad(valor) {
  const iso = fechaParaInput(valor);
  if (!iso) return null;
  const [a, m, d] = iso.split("-").map(Number);
  const objetivo = new Date(a, m - 1, d);
  const hoy = new Date();
  const base = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
  return Math.round((objetivo - base) / 86400000);
}

// ---------- CSV ----------
function parseCSV(texto, delimitador = ";") {
  const filas = [];
  let fila = [], campo = "", i = 0, entreComillas = false;
  while (i < texto.length) {
    const c = texto[i];
    if (entreComillas) {
      if (c === '"') {
        if (texto[i + 1] === '"') { campo += '"'; i += 2; continue; }
        entreComillas = false; i++; continue;
      }
      campo += c; i++; continue;
    }
    if (c === '"') { entreComillas = true; i++; continue; }
    if (c === delimitador) { fila.push(campo); campo = ""; i++; continue; }
    if (c === "\r") { i++; continue; }
    if (c === "\n") { fila.push(campo); filas.push(fila); fila = []; campo = ""; i++; continue; }
    campo += c; i++;
  }
  if (campo.length || fila.length) { fila.push(campo); filas.push(fila); }
  return filas.filter(f => f.some(v => String(v || "").trim() !== ""));
}
function serializarCampo(valor, delimitador) {
  const v = valor == null ? "" : String(valor);
  return /["\n\r]/.test(v) || v.includes(delimitador) ? `"${v.replace(/"/g, '""')}"` : v;
}
function serializeCSV(filas, delimitador = ";") {
  return filas.map(f => f.map(v => serializarCampo(v, delimitador)).join(delimitador)).join("\r\n");
}
function pareceCabecera(fila) {
  if (!fila || !fila.length) return false;
  const normalizadas = fila.map(normalizarCabecera);
  const coincidencias = COLUMNAS.filter(col => col.aliases.some(a => normalizadas.includes(normalizarCabecera(a)))).length;
  return coincidencias >= Math.min(4, COLUMNAS.length);
}
function crearMapaCabecera(cabecera) {
  const mapa = new Map();
  cabecera.forEach((nombre, indice) => mapa.set(normalizarCabecera(nombre), indice));
  return COLUMNAS.map((col, indicePorDefecto) => {
    for (const alias of col.aliases) {
      const encontrado = mapa.get(normalizarCabecera(alias));
      if (encontrado !== undefined) return encontrado;
    }
    return indicePorDefecto;
  });
}
function registroDesdeFila(fila, mapaIndices, indiceOriginal) {
  const registro = { _indiceOriginal: indiceOriginal };
  COLUMNAS.forEach((col, i) => registro[col.key] = String(fila[mapaIndices[i]] ?? "").trim());
  registro.cortoPlazo = canonSiNo(registro.cortoPlazo, false);
  registro.historial = serializarHistorial(parsearHistorial(registro.historial));
  return registro;
}
function filasParaGuardar() {
  const cabecera = COLUMNAS.map(c => c.label);
  const datos = G.registros
    .slice()
    .sort((a, b) => a._indiceOriginal - b._indiceOriginal)
    .map(registro => COLUMNAS.map(col => String(registro[col.key] ?? "").trim()));
  return [cabecera, ...datos];
}

// ---------- Graph / OneDrive ----------
function itemBase() {
  return G.driveId ? `${GRAPH}/drives/${G.driveId}` : `${GRAPH}/me/drive`;
}
async function getToken(request) {
  const active = msalInstance.getActiveAccount();
  const req = Object.assign({}, request, { account: request.account || active });
  try {
    return (await msalInstance.acquireTokenSilent(req)).accessToken;
  } catch (error) {
    console.warn("Adquisición silenciosa falló; se redirige al inicio de sesión.", error && error.errorCode);
    await msalInstance.acquireTokenRedirect(request);
    return null;
  }
}
async function gFetch(url, opciones = {}) {
  const token = await getToken(fileScopes);
  if (!token) throw new Error("No se pudo obtener el permiso de OneDrive.");
  opciones.headers = Object.assign({ Authorization: `Bearer ${token}` }, opciones.headers || {});
  return fetch(url, opciones);
}
async function localizarCarpeta(token) {
  const headers = { Authorization: `Bearer ${token}` };
  let url = `${GRAPH}/me/drive/root/children?$top=200`;
  const hijos = [];
  while (url) {
    const res = await fetch(url, { headers });
    if (!res.ok) {
      if (res.status === 401 || res.status === 403) {
        throw new Error(`Sin permiso para OneDrive (HTTP ${res.status}). Comprueba Files.ReadWrite en Azure.`);
      }
      throw new Error(`No se pudo leer la raíz de OneDrive (HTTP ${res.status}).`);
    }
    const json = await res.json();
    hijos.push(...(json.value || []));
    url = json["@odata.nextLink"] || null;
  }
  const objetivo = CARPETA_RAIZ.toLowerCase();
  const carpeta = hijos.find(item => (item.folder || item.remoteItem) && String(item.name || "").toLowerCase() === objetivo);
  if (!carpeta) return false;
  if (carpeta.remoteItem) {
    G.driveId = carpeta.remoteItem.parentReference && carpeta.remoteItem.parentReference.driveId;
    G.folderId = carpeta.remoteItem.id;
  } else {
    G.driveId = null;
    G.folderId = carpeta.id;
  }
  return true;
}
async function descargarTexto(ruta) {
  const path = encodeGraphPath(ruta);
  const res = await gFetch(`${itemBase()}/items/${G.folderId}:/${path}:/content`);
  if (!res.ok) {
    if (res.status === 404) throw new Error(`No existe "${ruta}" dentro de "${CARPETA_RAIZ}".`);
    throw new Error(`No se pudo descargar "${ruta}" (HTTP ${res.status}).`);
  }
  return new TextDecoder("utf-8").decode(await res.arrayBuffer()).replace(/^\uFEFF/, "");
}
async function subirTexto(ruta, texto) {
  const path = encodeGraphPath(ruta);
  const res = await gFetch(`${itemBase()}/items/${G.folderId}:/${path}:/content`, {
    method: "PUT",
    headers: { "Content-Type": "text/csv;charset=utf-8" },
    body: `\uFEFF${texto}`
  });
  if (!res.ok) throw new Error(`No se pudo guardar "${ruta}" (HTTP ${res.status}: ${await res.text()}).`);
  return res.json();
}
async function resolverFoto(rutaOriginal) {
  const ruta = normalizarRutaFoto(rutaOriginal);
  if (!ruta) return null;
  if (esUrl(ruta)) return { miniatura: ruta, completa: ruta };
  if (G.fotoCache.has(ruta)) return G.fotoCache.get(ruta);

  const promesa = (async () => {
    const intentos = ruta.includes("/") ? [ruta] : [ruta, `fotos/${ruta}`];
    for (const intento of intentos) {
      const path = encodeGraphPath(intento);
      const res = await gFetch(`${itemBase()}/items/${G.folderId}:/${path}?$expand=thumbnails`);
      if (!res.ok) continue;
      const item = await res.json();
      const th = item.thumbnails && item.thumbnails[0];
      const mini = th && (th.medium || th.large || th.small);
      const completa = item["@microsoft.graph.downloadUrl"] || (mini && mini.url) || "";
      if (mini || completa) return { miniatura: (mini && mini.url) || completa, completa };
    }
    return null;
  })();
  G.fotoCache.set(ruta, promesa);
  return promesa;
}
// Subida de fotos (rev. 5). Redimensionado en cliente para no superar el
// límite de 4 MB del PUT simple de Graph y ahorrar espacio.
function sanitizarNombreFichero(valor) {
  const limpio = quitarAcentos(String(valor || "")).toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
  return limpio || "articulo";
}
function generarRutaFoto(nombreArticulo) {
  const d = new Date();
  const p = n => String(n).padStart(2, "0");
  const marca = `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
  return `fotos/${sanitizarNombreFichero(nombreArticulo)}_${marca}.jpg`;
}
function redimensionarFoto(file, maxLado = 1600, calidad = 0.85) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      const escala = Math.min(1, maxLado / Math.max(img.width, img.height));
      const w = Math.max(1, Math.round(img.width * escala));
      const h = Math.max(1, Math.round(img.height * escala));
      const canvas = document.createElement("canvas");
      canvas.width = w;
      canvas.height = h;
      canvas.getContext("2d").drawImage(img, 0, 0, w, h);
      canvas.toBlob(blob => blob ? resolve(blob) : reject(new Error("No se pudo procesar la imagen.")), "image/jpeg", calidad);
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("No se pudo leer la imagen seleccionada.")); };
    img.src = url;
  });
}
async function subirFotoOneDrive(ruta, blob) {
  const path = encodeGraphPath(normalizarRutaFoto(ruta));
  const res = await gFetch(`${itemBase()}/items/${G.folderId}:/${path}:/content`, {
    method: "PUT",
    headers: { "Content-Type": "image/jpeg" },
    body: blob
  });
  if (!res.ok) throw new Error(`No se pudo subir la foto (HTTP ${res.status}).`);
  G.fotoCache.delete(normalizarRutaFoto(ruta));
  return res.json();
}
// Borra el fichero de una foto propia (ruta relativa). Tolerante a fallos:
// un fichero huérfano en fotos/ no rompe nada.
async function borrarFotoOneDrive(ruta) {
  const limpia = normalizarRutaFoto(ruta);
  if (!limpia || esUrl(limpia)) return;
  try {
    const path = encodeGraphPath(limpia);
    const res = await gFetch(`${itemBase()}/items/${G.folderId}:/${path}`, { method: "DELETE" });
    if (!res.ok && res.status !== 404) console.warn("No se pudo borrar la foto", limpia, res.status);
  } catch (error) {
    console.warn("No se pudo borrar la foto", limpia, error);
  }
  G.fotoCache.delete(limpia);
}

// ---------- Datos ----------
async function cargarDatos() {
  const cargaId = ++G.cargaId;
  setEstado(`Cargando "${FILE_DATOS}"…`);
  const texto = await descargarTexto(FILE_DATOS);
  if (cargaId !== G.cargaId) return;
  const filas = parseCSV(texto);
  if (!filas.length) {
    G.registros = [];
  } else {
    const tieneCabecera = pareceCabecera(filas[0]);
    const cabecera = tieneCabecera ? filas[0] : COLUMNAS.map(c => c.label);
    const mapa = crearMapaCabecera(cabecera);
    const datos = tieneCabecera ? filas.slice(1) : filas;
    G.registros = datos.map((fila, indice) => registroDesdeFila(fila, mapa, indice));
  }
  G.fotoCache.clear();
  poblarFiltros();
  render();
  setEstado(`${G.registros.length} artículo(s) cargado(s).`);
}
async function guardarDatos() {
  if (guardando) throw new Error("Ya hay un guardado en curso.");
  guardando = true;
  const boton = $("btnGuardarEditor");
  if (boton) boton.disabled = true;
  try {
    setEstado("Guardando cambios en OneDrive…");
    await subirTexto(FILE_DATOS, serializeCSV(filasParaGuardar()));
    setEstado("Cambios guardados en OneDrive.");
  } finally {
    guardando = false;
    if (boton) boton.disabled = false;
  }
}
function siguienteIndiceOriginal() {
  return G.registros.reduce((max, r) => Math.max(max, r._indiceOriginal), -1) + 1;
}

// ---------- Filtros (combos de multiselección) ----------
// Todos los filtros son combos con los valores distintos existentes en los
// datos, con multiselección: OR dentro del campo, AND entre campos.
// La opción "(vacío)" permite filtrar valores en blanco.
const VACIO = "__vacio__";
const CAMPOS_FILTRO = [
  { key: "nombre", label: "Nombre", tipo: "texto" },
  { key: "cantidad", label: "Cantidad", tipo: "texto" },
  { key: "casa", label: "Casa", tipo: "texto", base: CASAS },
  { key: "ubicacion", label: "Ubicación", tipo: "texto", base: UBICACIONES },
  { key: "fechaCaducidad", label: "Fecha caducidad", tipo: "fecha" },
  { key: "cortoPlazo", label: "Corto plazo abierto", tipo: "texto", base: ["Sí", "No"] },
  { key: "categoria", label: "Categoría", tipo: "texto", base: CATEGORIAS },
  { key: "foto", label: "Foto", tipo: "texto" },
  { key: "fechaUltimoCambio", label: "Fecha último cambio", tipo: "fecha" }
];
// Estado de selección: key → Set de valores normalizados (ISO en fechas).
const filtrosSel = new Map(CAMPOS_FILTRO.map(c => [c.key, new Set()]));

function opcionesUnicas(key, base = []) {
  const vistos = new Map();
  base.forEach(v => vistos.set(norm(v), v));
  G.registros.forEach(r => {
    const v = String(r[key] || "").trim();
    if (v && !vistos.has(norm(v))) vistos.set(norm(v), v);
  });
  return [...vistos.entries()]; // [valorNormalizado, etiqueta]
}
function opcionesFiltro(campo) {
  const hayVacio = G.registros.some(r => {
    const bruto = String(r[campo.key] || "").trim();
    return campo.tipo === "fecha" ? !fechaParaInput(bruto) : !bruto;
  });
  let opciones;
  if (campo.tipo === "fecha") {
    const vistas = new Map();
    G.registros.forEach(r => {
      const iso = fechaParaInput(r[campo.key]);
      if (iso && !vistas.has(iso)) vistas.set(iso, fechaVisible(iso));
    });
    opciones = [...vistas.entries()]
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([valor, etiqueta]) => ({ valor, etiqueta }));
  } else {
    const entradas = opcionesUnicas(campo.key, campo.base || []);
    if (campo.key === "casa") {
      entradas.sort((a, b) => {
        const ra = ORDEN_CASAS.has(a[0]) ? ORDEN_CASAS.get(a[0]) : 100;
        const rb = ORDEN_CASAS.has(b[0]) ? ORDEN_CASAS.get(b[0]) : 100;
        return (ra - rb) || collator.compare(a[1], b[1]);
      });
    } else if (campo.key === "ubicacion") {
      entradas.sort((a, b) => {
        const ra = ORDEN_UBICACIONES.has(a[0]) ? ORDEN_UBICACIONES.get(a[0]) : 40;
        const rb = ORDEN_UBICACIONES.has(b[0]) ? ORDEN_UBICACIONES.get(b[0]) : 40;
        return (ra - rb) || collator.compare(a[1], b[1]);
      });
    } else {
      entradas.sort((a, b) => collator.compare(a[1], b[1]));
    }
    opciones = entradas.map(([valor, etiqueta]) => ({ valor, etiqueta }));
  }
  if (hayVacio) opciones.push({ valor: VACIO, etiqueta: "(vacío)" });
  return opciones;
}
function construirFiltros() {
  const cont = $("rejillaFiltros");
  if (!cont) return;
  cont.innerHTML = CAMPOS_FILTRO.map(c => `
    <div class="campo ms" data-key="${c.key}">
      <span class="ms-etiqueta">${escapeHtml(c.label)}</span>
      <button type="button" class="ms-boton" aria-haspopup="listbox" aria-expanded="false">Todos</button>
      <div class="ms-panel" hidden></div>
    </div>`).join("");
}
function poblarFiltros() {
  CAMPOS_FILTRO.forEach(campo => {
    const opciones = opcionesFiltro(campo);
    const sel = filtrosSel.get(campo.key);
    // Conservar solo las selecciones cuyo valor sigue existiendo.
    [...sel].forEach(v => { if (!opciones.some(o => o.valor === v)) sel.delete(v); });
    const panel = document.querySelector(`.ms[data-key="${campo.key}"] .ms-panel`);
    if (!panel) return;
    panel.innerHTML = opciones.length
      ? opciones.map(o => `<label><input type="checkbox" value="${escapeHtml(o.valor)}"${sel.has(o.valor) ? " checked" : ""} /> ${escapeHtml(o.etiqueta)}</label>`).join("")
      : `<p class="ms-vacia">Sin valores</p>`;
    actualizarBotonMs(campo.key);
  });
}
function actualizarBotonMs(key) {
  const cont = document.querySelector(`.ms[data-key="${key}"]`);
  if (!cont) return;
  const boton = cont.querySelector(".ms-boton");
  const marcadas = [...cont.querySelectorAll("input:checked")].map(i => i.parentElement.textContent.trim());
  boton.textContent = marcadas.length ? marcadas.join(", ") : "Todos";
  boton.title = boton.textContent;
  boton.classList.toggle("con-seleccion", marcadas.length > 0);
}
function cerrarPanelesMs() {
  document.querySelectorAll(".ms-panel").forEach(p => { p.hidden = true; });
  document.querySelectorAll(".ms-boton").forEach(b => b.setAttribute("aria-expanded", "false"));
}
function valorFiltrable(campo, registro) {
  const bruto = String(registro[campo.key] || "").trim();
  if (campo.tipo === "fecha") return fechaParaInput(bruto);
  return norm(bruto);
}

// Botones rápidos en dos grupos: casa y categoría.
// OR dentro de cada grupo, AND entre grupos y con el panel de filtros.
const casasBoton = new Set();       // "castefa" | "comarruga"
const categoriasBoton = new Set();  // "cesped" | "frutos secos" | "harinas" | "otros"
function coincideBotones(registro) {
  if (casasBoton.size && !casasBoton.has(norm(registro.casa))) return false;
  if (categoriasBoton.size && !categoriasBoton.has(norm(registro.categoria))) return false;
  return true;
}
function registroCoincide(registro) {
  for (const campo of CAMPOS_FILTRO) {
    const sel = filtrosSel.get(campo.key);
    if (!sel.size) continue;
    const v = valorFiltrable(campo, registro);
    if (!v) {
      if (!sel.has(VACIO)) return false;
      continue;
    }
    if (!sel.has(v)) return false;
  }
  return coincideBotones(registro);
}
function compararRegistros(a, b) {
  const ca = norm(a.casa), cb = norm(b.casa);
  const ra = ORDEN_CASAS.has(ca) ? ORDEN_CASAS.get(ca) : 100;
  const rb = ORDEN_CASAS.has(cb) ? ORDEN_CASAS.get(cb) : 100;
  if (ra !== rb) return ra - rb;
  const porCasa = collator.compare(a.casa || "", b.casa || "");
  if (porCasa) return porCasa;
  const ua = ORDEN_UBICACIONES.has(norm(a.ubicacion)) ? ORDEN_UBICACIONES.get(norm(a.ubicacion)) : 50;
  const ub = ORDEN_UBICACIONES.has(norm(b.ubicacion)) ? ORDEN_UBICACIONES.get(norm(b.ubicacion)) : 50;
  if (ua !== ub) return ua - ub;
  // Caducidad ascendente; sin fecha, al final del grupo.
  const fa = fechaParaInput(a.fechaCaducidad) || "9999-12-31";
  const fb = fechaParaInput(b.fechaCaducidad) || "9999-12-31";
  const porFecha = fa.localeCompare(fb);
  if (porFecha) return porFecha;
  return collator.compare(a.nombre || "", b.nombre || "");
}
function limpiarFiltros() {
  filtrosSel.forEach(sel => sel.clear());
  document.querySelectorAll(".ms-panel input[type=checkbox]").forEach(i => { i.checked = false; });
  CAMPOS_FILTRO.forEach(c => actualizarBotonMs(c.key));
  cerrarPanelesMs();
  casasBoton.clear();
  categoriasBoton.clear();
  actualizarBotonesRapidos();
  render();
}

// ---------- Render ----------
function claseCasa(casa) {
  const c = norm(casa);
  if (c === "castefa") return "casa-castefa";
  if (c === "comarruga") return "casa-comarruga";
  return "";
}
// Tono por ubicación: oscuro congelador, medio frigo, claro armario/arriba.
function claseTono(registro) {
  const u = norm(registro.ubicacion);
  if (u === "congelador") return "tono-oscuro";
  if (u === "frigo") return "tono-medio";
  if (u === "armario" || u === "arriba") return "tono-claro";
  return "";
}
// Urgencia: rojo caducado, verde resto.
function claseUrgencia(registro, dias) {
  if (dias != null && dias < 0) return "caducado";
  return "sin-urgencia";
}
function badgeCaducidad(registro, dias) {
  if (dias == null) {
    return String(registro.fechaCaducidad || "").trim()
      ? `<span class="badge">Caducidad: ${escapeHtml(valorVisible(registro.fechaCaducidad))}</span>`
      : `<span class="badge">Sin caducidad</span>`;
  }
  const fecha = fechaVisible(registro.fechaCaducidad);
  if (dias < 0) return `<span class="badge caducidad-caducado">Caducado: ${escapeHtml(fecha)}</span>`;
  if (dias === 0) return `<span class="badge caducidad-caducado">Caduca hoy</span>`;
  return `<span class="badge">Caduca: ${escapeHtml(fecha)}</span>`;
}
function tarjetaHtml(registro, indice) {
  const dias = diasHastaCaducidad(registro.fechaCaducidad);
  const clase = `${claseCasa(registro.casa)} ${claseTono(registro)} ${claseUrgencia(registro, dias)}`;
  const cortoPlazo = esSi(registro.cortoPlazo)
    ? `<span class="badge corto-plazo">⚠ Corto plazo abierto</span>`
    : "";
  const foto = normalizarRutaFoto(registro.foto);
  const marcoFoto = foto
    ? `<div class="foto-marco" data-foto-marco="${indice}"><span class="cargando-foto">Cargando foto…</span></div>`
    : `<div class="foto-marco"><span class="sin-foto">Sin foto</span></div>`;
  // La categoría no se muestra en el listado (se filtra con los botones rápidos).
  return `
    <article class="articulo ${clase}" data-indice="${indice}">
      ${marcoFoto}
      <div class="vista-previa-articulo">
        <p class="nombre-principal">${escapeHtml(valorVisible(registro.nombre, "Artículo sin nombre"))}</p>
        <div class="badges">
          <span class="badge cantidad">Cantidad: ${escapeHtml(cantidadDe(registro))}</span>
          <span class="badge">Casa: ${escapeHtml(valorVisible(registro.casa))}</span>
          <span class="badge ubicacion">Ubicación: ${escapeHtml(valorVisible(registro.ubicacion))}</span>
          ${badgeCaducidad(registro, dias)}
          ${cortoPlazo}
        </div>
        <button class="editar" type="button" data-editar="${indice}">Ver / editar</button>
      </div>
    </article>`;
}
function render() {
  // Pantalla principal: solo artículos en stock. Los de cantidad vacía
  // (plantillas) y los "No queda nada" (cantidad 0) se proponen al pulsar
  // "Nuevo artículo".
  const visibles = G.registros
    .map((registro, indice) => ({ registro, indice }))
    .filter(({ registro }) => esVisibleEnListado(registro) && registroCoincide(registro))
    .sort((a, b) => compararRegistros(a.registro, b.registro));

  const lista = $("listaDespensa");
  if (!visibles.length) {
    lista.innerHTML = `<div class="vacio-lista">No hay artículos que coincidan con los filtros.</div>`;
    return;
  }
  lista.innerHTML = visibles.map(({ registro, indice }) => tarjetaHtml(registro, indice)).join("");
  cargarFotosRenderizadas(visibles);
}
async function cargarFotosRenderizadas(visibles) {
  for (const { registro, indice } of visibles) {
    const marco = document.querySelector(`[data-foto-marco="${indice}"]`);
    if (!marco || !registro.foto) continue;
    try {
      const foto = await resolverFoto(registro.foto);
      if (!marco.isConnected) continue;
      if (!foto || !foto.miniatura) {
        marco.innerHTML = `<span class="sin-foto">Foto no encontrada</span>`;
        continue;
      }
      marco.innerHTML = `<img src="${escapeHtml(foto.miniatura)}" alt="Foto de ${escapeHtml(registro.nombre)}" loading="lazy" />`
        + (foto.completa ? `<button type="button" data-abrir-foto="${indice}" aria-label="Abrir foto"></button>` : "");
      marco.dataset.urlCompleta = foto.completa || foto.miniatura;
    } catch (error) {
      if (marco.isConnected) marco.innerHTML = `<span class="sin-foto">No se pudo cargar</span>`;
      console.warn("No se pudo cargar la foto", registro.foto, error);
    }
  }
}

// ---------- Editor (crear / editar / eliminar) ----------
// editorIndice = "" → alta de artículo nuevo; número → edición del registro.
// Foto pendiente de subir (se sube al Guardar; Cancelar la descarta).
let fotoSeleccionada = null;      // { file, ruta } o null
let urlPreviewLocal = null;       // objectURL de la preview local
let previewToken = 0;
function liberarPreviewLocal() {
  if (urlPreviewLocal) { URL.revokeObjectURL(urlPreviewLocal); urlPreviewLocal = null; }
}
async function actualizarPreviewFoto(valor) {
  const marco = $("previewFoto");
  const token = ++previewToken;
  const ruta = normalizarRutaFoto(valor);
  if (!ruta) { marco.innerHTML = `<span class="sin-foto">Sin foto</span>`; return; }
  marco.innerHTML = `<span class="cargando-foto">Cargando…</span>`;
  try {
    const foto = await resolverFoto(ruta);
    if (token !== previewToken) return;
    marco.innerHTML = foto && foto.miniatura
      ? `<img src="${escapeHtml(foto.miniatura)}" alt="Foto del artículo" />`
      : `<span class="sin-foto">No encontrada</span>`;
  } catch (error) {
    if (token === previewToken) marco.innerHTML = `<span class="sin-foto">No se pudo cargar</span>`;
  }
}
function mostrarPreviewLocal(file) {
  liberarPreviewLocal();
  previewToken++;
  urlPreviewLocal = URL.createObjectURL(file);
  $("previewFoto").innerHTML = `<img src="${urlPreviewLocal}" alt="Foto seleccionada" />`;
}
function seleccionarFoto(file) {
  if (!file) return;
  const ruta = generarRutaFoto($("eNombre").value || file.name.replace(/\.[^.]+$/, ""));
  fotoSeleccionada = { file, ruta };
  $("eFoto").value = ruta;
  mostrarPreviewLocal(file);
}
function quitarFoto() {
  fotoSeleccionada = null;
  liberarPreviewLocal();
  $("eFoto").value = "";
  $("inputFotoFichero").value = "";
  previewToken++;
  $("previewFoto").innerHTML = `<span class="sin-foto">Sin foto</span>`;
}
function abrirEditor(indice) {
  const esNuevo = indice == null;
  const registro = esNuevo ? null : G.registros[indice];
  if (!esNuevo && !registro) return;

  $("editorIndice").value = esNuevo ? "" : String(indice);
  $("tituloEditor").textContent = esNuevo
    ? "Nuevo artículo"
    : valorVisible(registro.nombre, "Artículo");

  $("eNombre").value = esNuevo ? "" : (registro.nombre || "");
  $("eCantidad").value = esNuevo ? "" : (registro.cantidad || "");
  $("eCasa").value = esNuevo
    ? CASAS[0]
    : (CASAS.find(v => norm(v) === norm(registro.casa)) || CASAS[0]);
  $("eUbicacion").value = esNuevo
    ? ""
    : (UBICACIONES.find(v => norm(v) === norm(registro.ubicacion)) || "");
  $("eFechaCaducidad").value = esNuevo ? "" : fechaParaInput(registro.fechaCaducidad);
  $("eCortoPlazo").value = !esNuevo && esSi(registro.cortoPlazo) ? "Sí" : "No";
  $("eCategoria").value = esNuevo
    ? ""
    : (CATEGORIAS.find(v => norm(v) === norm(registro.categoria)) || "");
  $("eFoto").value = esNuevo ? "" : (registro.foto || "");
  fotoSeleccionada = null;
  liberarPreviewLocal();
  $("inputFotoFichero").value = "";
  actualizarPreviewFoto($("eFoto").value);
  // Fecha último cambio: se propone SIEMPRE hoy; Cancelar no guarda nada,
  // así que la fecha anterior se conserva si no se confirma.
  $("eFechaUltimoCambio").value = hoyIso();

  $("btnEliminarEditor").hidden = esNuevo;
  prepararHistorial(registro);

  const dialogo = $("editor");
  if (typeof dialogo.showModal === "function") dialogo.showModal();
  else dialogo.setAttribute("open", "");
}
function cerrarEditor() {
  fotoSeleccionada = null;
  liberarPreviewLocal();
  const dialogo = $("editor");
  if (typeof dialogo.close === "function") dialogo.close();
  else dialogo.removeAttribute("open");
}
// ---------- Historial en el editor ----------
function fechaHoraVisible(valor) {
  if (!valor) return "Fecha no registrada";
  const fecha = new Date(valor);
  if (Number.isNaN(fecha.getTime())) return valorVisible(valor);
  return new Intl.DateTimeFormat("es-ES", { dateStyle: "short", timeStyle: "short" }).format(fecha);
}
function datoHtml(label, valor, fecha = false) {
  const texto = fecha ? fechaVisible(valor) : valorVisible(valor);
  return `<div class="dato"><dt>${escapeHtml(label)}</dt><dd>${escapeHtml(texto)}</dd></div>`;
}
function historialHtml(registro) {
  const historial = parsearHistorial(registro && registro.historial);
  if (!historial.length) return `<p class="historial-vacio">Todavía no hay estados anteriores guardados.</p>`;
  return historial.map((uso, indice) => `
    <article class="uso-anterior">
      <div class="uso-anterior-cabecera">
        <strong>Estado anterior ${indice + 1}</strong>
        <span>${escapeHtml(fechaHoraVisible(uso.guardadoEn))}</span>
      </div>
      <dl class="datos-historial">
        ${datoHtml("Cantidad", uso.cantidad)}
        ${datoHtml("Casa", uso.casa)}
        ${datoHtml("Ubicación", uso.ubicacion)}
        ${datoHtml("Fecha caducidad", uso.fechaCaducidad, true)}
        ${datoHtml("Corto plazo abierto", uso.cortoPlazo)}
        ${datoHtml("Categoría", uso.categoria)}
        ${datoHtml("Fecha último cambio", uso.fechaUltimoCambio, true)}
      </dl>
    </article>`).join("");
}
function prepararHistorial(registro) {
  $("historialUsos").innerHTML = historialHtml(registro);
  $("panelHistorial").hidden = true;
  $("btnHistorial").setAttribute("aria-expanded", "false");
  $("btnHistorial").textContent = `Mostrar los ${MAX_HISTORIAL} estados anteriores`;
}
function alternarHistorial() {
  const panel = $("panelHistorial");
  const mostrar = panel.hidden;
  panel.hidden = !mostrar;
  $("btnHistorial").setAttribute("aria-expanded", mostrar ? "true" : "false");
  $("btnHistorial").textContent = mostrar ? "Ocultar estados anteriores" : `Mostrar los ${MAX_HISTORIAL} estados anteriores`;
}
async function guardarEditor(evento) {
  evento.preventDefault();
  const form = $("formEditor");
  if (!form.reportValidity()) return;

  const valorIndice = $("editorIndice").value;
  const esNuevo = valorIndice === "";
  const nuevo = {
    nombre: $("eNombre").value.trim(),
    cantidad: $("eCantidad").value.trim(),
    casa: $("eCasa").value,
    ubicacion: $("eUbicacion").value,
    fechaCaducidad: $("eFechaCaducidad").value,
    cortoPlazo: $("eCortoPlazo").value,
    categoria: $("eCategoria").value,
    foto: $("eFoto").value.trim(),
    fechaUltimoCambio: $("eFechaUltimoCambio").value
  };

  // Foto pendiente: solo se sube si el campo sigue apuntando a su ruta
  // (si se ha editado o vaciado el campo después, se descarta).
  const pendiente = fotoSeleccionada && fotoSeleccionada.ruta === nuevo.foto ? fotoSeleccionada : null;
  if (pendiente) {
    const botonGuardar = $("btnGuardarEditor");
    if (botonGuardar) botonGuardar.disabled = true;
    try {
      setEstado("Subiendo foto…");
      const blob = await redimensionarFoto(pendiente.file);
      await subirFotoOneDrive(pendiente.ruta, blob);
    } catch (error) {
      setEstado(`Error subiendo la foto: ${error.message}`, true);
      if (botonGuardar) botonGuardar.disabled = false;
      return;
    }
    if (botonGuardar) botonGuardar.disabled = false;
    fotoSeleccionada = null;
  }

  if (esNuevo) {
    const registro = Object.assign({ _indiceOriginal: siguienteIndiceOriginal(), historial: serializarHistorial([]) }, nuevo);
    G.registros.push(registro);
    poblarFiltros();
    render();
    try {
      await guardarDatos();
      cerrarEditor();
    } catch (error) {
      G.registros = G.registros.filter(r => r !== registro);
      poblarFiltros();
      render();
      setEstado(`Error guardando: ${error.message}`, true);
    }
    return;
  }

  const indice = Number(valorIndice);
  const registro = G.registros[indice];
  if (!registro) return;
  const copiaAnterior = Object.assign({}, registro);
  // La fecha último cambio se propone sola con hoy, así que no cuenta como
  // cambio por sí misma: si no se ha tocado nada más, no se guarda nada y
  // se conserva la fecha anterior.
  const hayCambios = Object.keys(nuevo)
    .filter(key => key !== "fechaUltimoCambio")
    .some(key => String(registro[key] || "").trim() !== String(nuevo[key] || "").trim());
  if (!hayCambios) {
    setEstado("No había cambios que guardar.");
    cerrarEditor();
    return;
  }
  apilarHistorial(registro);
  Object.assign(registro, nuevo);
  poblarFiltros();
  render();
  try {
    await guardarDatos();
    // Si la foto propia anterior se ha sustituido o quitado, borrar el
    // fichero antiguo de OneDrive (tolerante a fallos).
    const fotoAntigua = normalizarRutaFoto(copiaAnterior.foto);
    if (fotoAntigua && !esUrl(fotoAntigua) && fotoAntigua !== normalizarRutaFoto(nuevo.foto)) {
      borrarFotoOneDrive(fotoAntigua);
    }
    cerrarEditor();
  } catch (error) {
    Object.assign(registro, copiaAnterior);
    poblarFiltros();
    render();
    setEstado(`Error guardando: ${error.message}`, true);
  }
}
// El botón Eliminar abre un diálogo con dos opciones:
// - "No queda nada": conserva la fila con Cantidad "0" y vacía Ubicación y
//   Fecha caducidad (el artículo sigue visible, atenuado, como recordatorio).
// - "Eliminar completamente": borra la fila del CSV.
function abrirDialogoEliminar() {
  const valorIndice = $("editorIndice").value;
  if (valorIndice === "") return;
  const registro = G.registros[Number(valorIndice)];
  if (!registro) return;
  $("nombreEliminar").textContent = valorVisible(registro.nombre, "este artículo");
  const dialogo = $("dialogoEliminar");
  if (typeof dialogo.showModal === "function") dialogo.showModal();
  else dialogo.setAttribute("open", "");
}
function cerrarDialogoEliminar() {
  const dialogo = $("dialogoEliminar");
  if (typeof dialogo.close === "function") dialogo.close();
  else dialogo.removeAttribute("open");
}
async function marcarSinExistencias() {
  const valorIndice = $("editorIndice").value;
  if (valorIndice === "") return;
  const registro = G.registros[Number(valorIndice)];
  if (!registro) return;
  const nombre = valorVisible(registro.nombre, "este artículo");
  const copiaAnterior = Object.assign({}, registro);
  apilarHistorial(registro);
  registro.cantidad = "0";
  registro.ubicacion = "";
  registro.fechaCaducidad = "";
  registro.fechaUltimoCambio = hoyIso();
  poblarFiltros();
  render();
  try {
    await guardarDatos();
    cerrarDialogoEliminar();
    cerrarEditor();
    setEstado(`"${nombre}" marcado como "No queda nada".`);
  } catch (error) {
    Object.assign(registro, copiaAnterior);
    poblarFiltros();
    render();
    setEstado(`Error guardando: ${error.message}`, true);
  }
}
async function eliminarDefinitivo() {
  const valorIndice = $("editorIndice").value;
  if (valorIndice === "") return;
  const indice = Number(valorIndice);
  const registro = G.registros[indice];
  if (!registro) return;
  const nombre = valorVisible(registro.nombre, "este artículo");
  const copia = G.registros.slice();
  G.registros.splice(indice, 1);
  poblarFiltros();
  render();
  try {
    await guardarDatos();
    const fotoPropia = normalizarRutaFoto(registro.foto);
    if (fotoPropia && !esUrl(fotoPropia)) borrarFotoOneDrive(fotoPropia);
    cerrarDialogoEliminar();
    cerrarEditor();
    setEstado(`"${nombre}" eliminado completamente.`);
  } catch (error) {
    G.registros = copia;
    poblarFiltros();
    render();
    setEstado(`Error eliminando: ${error.message}`, true);
  }
}

// ---------- Selector de artículo nuevo ----------
// "Nuevo artículo" propone primero los artículos sin existencias: plantillas
// (Cantidad vacía) y "No queda nada" (Cantidad 0), además de crear desde cero.
function plantillasCatalogo() {
  return G.registros
    .map((registro, indice) => ({ registro, indice }))
    .filter(({ registro }) => !esVisibleEnListado(registro))
    .sort((a, b) => collator.compare(a.registro.nombre || "", b.registro.nombre || ""));
}
function abrirSelectorNuevo() {
  const plantillas = plantillasCatalogo();
  if (!plantillas.length) { abrirEditor(null); return; }
  $("listaPlantillas").innerHTML = plantillas.map(({ registro, indice }) => {
    const detalle = [registro.casa, registro.categoria].filter(Boolean).join(" · ") || "—";
    const marca = esAgotado(registro) ? ` · <strong>No queda nada</strong>` : "";
    return `
    <button type="button" class="plantilla" data-plantilla="${indice}">
      <span class="plantilla-nombre">${escapeHtml(valorVisible(registro.nombre, "Sin nombre"))}</span>
      <span class="plantilla-detalle">${escapeHtml(detalle)}${marca}</span>
    </button>`;
  }).join("");
  const dialogo = $("selectorNuevo");
  if (typeof dialogo.showModal === "function") dialogo.showModal();
  else dialogo.setAttribute("open", "");
}
function cerrarSelectorNuevo() {
  const dialogo = $("selectorNuevo");
  if (typeof dialogo.close === "function") dialogo.close();
  else dialogo.removeAttribute("open");
}

// ---------- Autenticación e inicio ----------
async function iniciar() {
  if (inicializado) return;
  inicializado = true;
  setEstado(`Buscando la carpeta "${CARPETA_RAIZ}"…`);
  try {
    const token = await getToken(fileScopes);
    if (!token) { inicializado = false; return; }
    const encontrada = await localizarCarpeta(token);
    if (!encontrada) {
      throw new Error(`No existe la carpeta "${CARPETA_RAIZ}" en la raíz de OneDrive.`);
    }
    await cargarDatos();
    mostrarVista("inventario");
  } catch (error) {
    setEstado(`Error: ${error.message}`, true);
    inicializado = false;
  }
}
function updateAuthUI() {
  const account = msalInstance.getActiveAccount();
  const loginBtn = $("loginBtn");
  if (!loginBtn) return;
  if (account) {
    loginBtn.textContent = `Cerrar sesión (${account.name || account.username})`;
    document.body.classList.add("logged-in");
    loginBtn.onclick = async event => {
      event.preventDefault();
      loginBtn.disabled = true;
      try {
        await msalInstance.logoutPopup({ account });
        msalInstance.setActiveAccount(null);
      } catch (error) {
        console.error(error);
      } finally {
        loginBtn.disabled = false;
        inicializado = false;
        G.registros = [];
        G.fotoCache.clear();
        $("listaDespensa").innerHTML = "";
        mostrarVista("");
        setEstado("Inicia sesión para empezar.");
        updateAuthUI();
      }
    };
    iniciar();
  } else {
    loginBtn.textContent = "Iniciar sesión";
    document.body.classList.remove("logged-in");
    loginBtn.onclick = async event => {
      event.preventDefault();
      loginBtn.disabled = true;
      try { await msalInstance.loginRedirect(loginRequest); }
      finally { loginBtn.disabled = false; }
    };
  }
}
function resetMsal() {
  Object.keys(localStorage)
    .filter(k => k.startsWith("msal.") || k.includes(msalConfig.auth.clientId))
    .forEach(k => localStorage.removeItem(k));
  if (location.hash) history.replaceState(null, "", location.pathname + location.search);
  console.log("Estado MSAL limpiado.");
}

// ---------- Eventos ----------
construirFiltros();

// Mostrar/ocultar el panel de filtros (encima del listado, plegado por defecto).
$("btnFiltros").addEventListener("click", () => {
  const panel = $("panelFiltros");
  const mostrar = panel.hidden;
  panel.hidden = !mostrar;
  const boton = $("btnFiltros");
  boton.textContent = mostrar ? "Ocultar filtros" : "Mostrar filtros";
  boton.setAttribute("aria-expanded", mostrar ? "true" : "false");
  if (!mostrar) cerrarPanelesMs();
});

// Apertura/cierre de los combos de multiselección.
$("rejillaFiltros").addEventListener("click", event => {
  const boton = event.target.closest(".ms-boton");
  if (!boton) return;
  const panel = boton.parentElement.querySelector(".ms-panel");
  const abrir = panel.hidden;
  cerrarPanelesMs();
  panel.hidden = !abrir;
  boton.setAttribute("aria-expanded", abrir ? "true" : "false");
});
// Cambios en las casillas de los combos.
$("rejillaFiltros").addEventListener("change", event => {
  const input = event.target.closest('.ms-panel input[type="checkbox"]');
  if (!input) return;
  const ms = input.closest(".ms");
  const sel = filtrosSel.get(ms.dataset.key);
  if (input.checked) sel.add(input.value);
  else sel.delete(input.value);
  actualizarBotonMs(ms.dataset.key);
  render();
});
// Cerrar los desplegables al pulsar fuera.
document.addEventListener("click", event => {
  if (!event.target.closest(".ms")) cerrarPanelesMs();
});

$("btnLimpiar").addEventListener("click", limpiarFiltros);

// Botones rápidos: grupo casa (Castefa/Comarruga) y grupo categoría.
const BOTONES_RAPIDOS = [
  ["btnCasaCastefa", "castefa", casasBoton],
  ["btnCasaComarruga", "comarruga", casasBoton],
  ["btnCatCesped", "cesped", categoriasBoton],
  ["btnCatFrutosSecos", "frutos secos", categoriasBoton],
  ["btnCatHarinas", "harinas", categoriasBoton],
  ["btnCatOtros", "otros", categoriasBoton]
];
function actualizarBotonesRapidos() {
  BOTONES_RAPIDOS.forEach(([id, clave, grupo]) => {
    const btn = $(id);
    if (btn) {
      btn.classList.toggle("activo", grupo.has(clave));
      btn.setAttribute("aria-pressed", grupo.has(clave) ? "true" : "false");
    }
  });
}
BOTONES_RAPIDOS.forEach(([id, clave, grupo]) => {
  const btn = $(id);
  if (!btn) return;
  btn.addEventListener("click", () => {
    if (grupo.has(clave)) grupo.delete(clave);
    else grupo.add(clave);
    actualizarBotonesRapidos();
    render();
  });
});

$("btnNuevo").addEventListener("click", abrirSelectorNuevo);
$("btnRecargar").addEventListener("click", () => cargarDatos().catch(error => setEstado(`Error: ${error.message}`, true)));
$("listaDespensa").addEventListener("click", event => {
  const editar = event.target.closest("[data-editar]");
  if (editar) { abrirEditor(Number(editar.dataset.editar)); return; }
  const abrirFoto = event.target.closest("[data-abrir-foto]");
  if (abrirFoto) {
    const marco = abrirFoto.closest(".foto-marco");
    const url = marco && marco.dataset.urlCompleta;
    if (url) window.open(url, "_blank", "noopener");
  }
});
$("btnHistorial").addEventListener("click", alternarHistorial);

// Gestión de foto en el editor.
$("btnSubirFoto").addEventListener("click", () => $("inputFotoFichero").click());
$("inputFotoFichero").addEventListener("change", event => {
  const file = event.target.files && event.target.files[0];
  if (file) seleccionarFoto(file);
});
$("btnQuitarFoto").addEventListener("click", quitarFoto);
// Si se edita el campo a mano, se descarta la foto pendiente y se refresca
// la preview con lo escrito.
$("eFoto").addEventListener("change", () => {
  if (fotoSeleccionada && fotoSeleccionada.ruta !== $("eFoto").value.trim()) {
    fotoSeleccionada = null;
    liberarPreviewLocal();
  }
  if (!fotoSeleccionada) actualizarPreviewFoto($("eFoto").value.trim());
});

// Selector de nuevo: elegir plantilla del catálogo o crear desde cero.
$("listaPlantillas").addEventListener("click", event => {
  const plantilla = event.target.closest("[data-plantilla]");
  if (!plantilla) return;
  cerrarSelectorNuevo();
  abrirEditor(Number(plantilla.dataset.plantilla));
});
$("btnDesdeCero").addEventListener("click", () => { cerrarSelectorNuevo(); abrirEditor(null); });
$("btnCerrarSelector").addEventListener("click", cerrarSelectorNuevo);
$("selectorNuevo").addEventListener("click", event => {
  if (event.target === $("selectorNuevo")) cerrarSelectorNuevo();
});

// Diálogo de eliminación con dos opciones.
$("btnEliminarEditor").addEventListener("click", abrirDialogoEliminar);
$("btnNoQuedaNada").addEventListener("click", marcarSinExistencias);
$("btnEliminarDefinitivo").addEventListener("click", eliminarDefinitivo);
$("btnCancelarEliminar").addEventListener("click", cerrarDialogoEliminar);
$("dialogoEliminar").addEventListener("click", event => {
  if (event.target === $("dialogoEliminar")) cerrarDialogoEliminar();
});
$("btnCerrarEditor").addEventListener("click", cerrarEditor);
$("btnCancelarEditor").addEventListener("click", cerrarEditor);
$("formEditor").addEventListener("submit", guardarEditor);
$("editor").addEventListener("click", event => {
  if (event.target === $("editor")) cerrarEditor();
});

(async () => {
  try {
    const respuesta = await msalInstance.handleRedirectPromise();
    if (respuesta && respuesta.account) msalInstance.setActiveAccount(respuesta.account);
    else {
      const cuentas = msalInstance.getAllAccounts();
      if (cuentas.length === 1) msalInstance.setActiveAccount(cuentas[0]);
    }
  } catch (error) {
    console.error(error);
    setEstado(`Error de autenticación: ${error.message}`, true);
  }
  updateAuthUI();
})();

window.despensaAuth = { msalInstance, getToken, resetMsal, G };
