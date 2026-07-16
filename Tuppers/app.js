// app.js — Gestión de tuppers sobre OneDrive
//
// Estructura esperada en OneDrive:
//   ZZZ_SG_Tuppers/
//     tuppers.csv
//     fotos/                       ← opcional
//       tupper_azul.jpg
//
// Columnas de tuppers.csv (separador ";"):
// Nombre tupper;Comentarios tupper;Material tapa;Material tupper;Fecha compra tupper;
// Foto tupper;Ocupado;En q casa está;Ubicación;Comida;Url cookido;Gluten;Leche;Fecha preparación;Historial
//
// La web solo modifica: Ocupado, En q casa está, Ubicación, Comida, Url cookido, Gluten, Leche y Fecha preparación.
// El resto de campos se modifica directamente en OneDrive.
// Casas predefinidas: Comarruga, Castefa, 3, 4 (antes CM y CS; el CSV usa ya los nombres largos).
// Gluten/Leche admiten: Sí, No, NPC/NPI.

// ====== Ajustes ======
const CARPETA_RAIZ = "ZZZ_SG_Tuppers";
const FILE_DATOS = "tuppers.csv";
const GRAPH = "https://graph.microsoft.com/v1.0";
// =====================

// Misma arquitectura de autenticación MSAL que la app de referencia.
const REDIRECT_PAGE =
  "https://atxcclt-26.github.io/atxcclt/Tuppers/";
const msalConfig = {
  auth: {
    clientId: "7765de16-766b-4051-b73d-d7167f5897dc",
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
  { key: "nombreTupper", label: "Nombre tupper", editable: false, aliases: ["nombre tupper"] },
  { key: "comentariosTupper", label: "Comentarios tupper", editable: false, aliases: ["comentarios tupper", "comentarios tuppers", "comentario tupper"] },
  { key: "materialTapa", label: "Material tapa", editable: false, aliases: ["material tapa"] },
  { key: "materialTupper", label: "Material tupper", editable: false, aliases: ["material tupper"] },
  { key: "fechaCompraTupper", label: "Fecha compra tupper", editable: false, aliases: ["fecha compra tupper"] },
  { key: "fotoTupper", label: "Foto tupper", editable: false, aliases: ["foto tupper", "foto"] },
  { key: "ocupado", label: "Ocupado", editable: true, aliases: ["ocupado"] },
  { key: "enQueCasaEsta", label: "En q casa está", editable: true, aliases: ["en q casa esta", "en que casa esta", "casa"] },
  { key: "ubicacion", label: "Ubicación", editable: true, aliases: ["ubicacion", "estado"] },
  { key: "comida", label: "Comida", editable: true, aliases: ["comida"] },
  { key: "urlCookido", label: "Url cookido", editable: true, aliases: ["url cookido", "url cooked", "cookido"] },
  { key: "gluten", label: "Gluten", editable: true, aliases: ["gluten"] },
  { key: "leche", label: "Leche", editable: true, aliases: ["leche"] },
  { key: "fechaPreparacion", label: "Fecha preparación", editable: true, aliases: ["fecha preparacion"] },
  { key: "historial", label: "Historial", editable: false, interno: true, aliases: ["historial"] }
];

const CAMPOS_FIJOS = COLUMNAS.filter(c => !c.editable && !c.interno);
const CASAS_PREDEFINIDAS = ["Comarruga", "Castefa", "3", "4"];
const OPCIONES_GLUTEN_LECHE = ["Sí", "No", "NPC/NPI"];
const ORDEN_CASAS = new Map([["comarruga", 0], ["castefa", 1], ["3", 2], ["4", 3]]);
const ORDEN_UBICACIONES = new Map([["congelador", 0], ["frigo", 1], ["fuera", 2], ["", 3]]);
const CAMPOS_HISTORIAL = ["ocupado", "enQueCasaEsta", "ubicacion", "comida", "urlCookido", "gluten", "leche", "fechaPreparacion"];
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
function normalizarRutaFoto(ruta) {
  return String(ruta || "").trim().replace(/^\/+/, "");
}
function esUrl(ruta) { return /^(https?:|data:|blob:)/i.test(String(ruta || "").trim()); }
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
  return coincidencias >= Math.min(5, COLUMNAS.length);
}
function crearMapaCabecera(cabecera) {
  const mapa = new Map();
  cabecera.forEach((nombre, indice) => mapa.set(normalizarCabecera(nombre), indice));
  const formatoNuevo = mapa.has(normalizarCabecera("En q casa está")) || mapa.has(normalizarCabecera("En que casa está"));

  return COLUMNAS.map((col, indicePorDefecto) => {
    let aliases = col.aliases;
    // Compatibilidad con la versión anterior: "Ubicación" era la casa y "Estado" era Frigo/Congelador/Fuera.
    if (col.key === "enQueCasaEsta" && !formatoNuevo) aliases = ["ubicacion", ...aliases];
    if (col.key === "ubicacion") aliases = formatoNuevo ? ["ubicacion", "estado"] : ["estado", "ubicacion"];
    for (const alias of aliases) {
      const encontrado = mapa.get(normalizarCabecera(alias));
      if (encontrado !== undefined) return encontrado;
    }
    return indicePorDefecto;
  });
}
function parsearHistorial(valor) {
  if (!valor) return [];
  try {
    const lista = JSON.parse(valor);
    return Array.isArray(lista) ? lista.filter(v => v && typeof v === "object").slice(0, 3) : [];
  } catch (error) {
    console.warn("Historial no válido; se ignora.", error);
    return [];
  }
}
function serializarHistorial(lista) {
  return JSON.stringify((Array.isArray(lista) ? lista : []).slice(0, 3));
}
function registroDesdeFila(fila, mapaIndices, indiceOriginal) {
  const registro = { _indiceOriginal: indiceOriginal };
  COLUMNAS.forEach((col, i) => registro[col.key] = String(fila[mapaIndices[i]] ?? "").trim());
  registro.ocupado = canonSiNo(registro.ocupado, false);
  registro.gluten = canonSiNo(registro.gluten, true);
  registro.leche = canonSiNo(registro.leche, true);
  registro.historial = serializarHistorial(parsearHistorial(registro.historial));
  if (!esSi(registro.ocupado)) {
    registro.ubicacion = "";
    registro.comida = "";
    registro.urlCookido = "";
    registro.gluten = "";
    registro.leche = "";
    registro.fechaPreparacion = "";
  }
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
  setEstado(`${G.registros.length} tupper(s) cargado(s).`);
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

// ---------- Filtros y orden ----------
const FILTROS = {
  nombreTupper: "fNombre",
  comentariosTupper: "fComentarios",
  materialTapa: "fMaterialTapa",
  materialTupper: "fMaterialTupper",
  fechaCompraTupper: "fFechaCompra",
  fotoTupper: "fFoto",
  ocupado: "fOcupado",
  enQueCasaEsta: "fCasa",
  ubicacion: "fUbicacion",
  comida: "fComida",
  urlCookido: "fUrlCookido",
  gluten: "fGluten",
  leche: "fLeche",
  fechaPreparacion: "fFechaPreparacion"
};
function leerFiltros() {
  return Object.fromEntries(Object.entries(FILTROS).map(([key, id]) => [key, $(id).value]));
}
function coincideTexto(valor, filtro) { return !norm(filtro) || norm(valor).includes(norm(filtro)); }
function coincideFecha(valor, filtro) { return !filtro || fechaParaInput(valor) === filtro; }
function coincideExacto(valor, filtro) { return !filtro || norm(valor) === filtro; }
function coincideSiNo(valor, filtro) {
  if (!filtro) return true;
  if (!norm(valor)) return false;
  return (esSi(valor) ? "si" : "no") === filtro;
}
// Botones rápidos de casa (Castefa / Comarruga): pueden estar activos ambos, uno o ninguno.
const casasBoton = new Set();
function coincideCasaBotones(registro) {
  return !casasBoton.size || casasBoton.has(norm(registro.enQueCasaEsta));
}
function registroCoincide(registro, filtros) {
  return coincideTexto(registro.nombreTupper, filtros.nombreTupper)
    && coincideTexto(registro.comentariosTupper, filtros.comentariosTupper)
    && coincideTexto(registro.materialTapa, filtros.materialTapa)
    && coincideTexto(registro.materialTupper, filtros.materialTupper)
    && coincideFecha(registro.fechaCompraTupper, filtros.fechaCompraTupper)
    && coincideTexto(registro.fotoTupper, filtros.fotoTupper)
    && coincideSiNo(registro.ocupado, filtros.ocupado)
    && coincideExacto(registro.enQueCasaEsta, filtros.enQueCasaEsta)
    && coincideCasaBotones(registro)
    && coincideExacto(registro.ubicacion, filtros.ubicacion)
    && coincideTexto(registro.comida, filtros.comida)
    && coincideTexto(registro.urlCookido, filtros.urlCookido)
    && coincideExacto(registro.gluten, filtros.gluten)
    && coincideExacto(registro.leche, filtros.leche)
    && coincideFecha(registro.fechaPreparacion, filtros.fechaPreparacion);
}
function compararCasaYUbicacion(a, b) {
  const ca = norm(a.enQueCasaEsta), cb = norm(b.enQueCasaEsta);
  const ra = ORDEN_CASAS.has(ca) ? ORDEN_CASAS.get(ca) : 100;
  const rb = ORDEN_CASAS.has(cb) ? ORDEN_CASAS.get(cb) : 100;
  if (ra !== rb) return ra - rb;
  const porCasa = collator.compare(a.enQueCasaEsta || "", b.enQueCasaEsta || "");
  if (porCasa) return porCasa;
  const ua = ORDEN_UBICACIONES.has(norm(a.ubicacion)) ? ORDEN_UBICACIONES.get(norm(a.ubicacion)) : 50;
  const ub = ORDEN_UBICACIONES.has(norm(b.ubicacion)) ? ORDEN_UBICACIONES.get(norm(b.ubicacion)) : 50;
  if (ua !== ub) return ua - ub;
  const porUbicacion = collator.compare(a.ubicacion || "", b.ubicacion || "");
  if (porUbicacion) return porUbicacion;
  return collator.compare(a.nombreTupper || "", b.nombreTupper || "");
}
function limpiarFiltros() {
  Object.values(FILTROS).forEach(id => { $(id).value = ""; });
  casasBoton.clear();
  actualizarBotonesCasa();
  render();
}

// ---------- Combos dinámicos ----------
function opcionesUnicas(key, base = []) {
  const vistos = new Map();
  base.forEach(v => vistos.set(norm(v), v));
  G.registros.forEach(r => {
    const v = String(r[key] || "").trim();
    if (v && !vistos.has(norm(v))) vistos.set(norm(v), v);
  });
  return [...vistos.entries()]; // [valorNormalizado, etiqueta]
}
function rellenarSelect(id, entradas, textoTodos) {
  const sel = $(id);
  if (!sel) return;
  const actual = sel.value;
  sel.innerHTML = `<option value="">${textoTodos}</option>`
    + entradas.map(([v, etiqueta]) => `<option value="${escapeHtml(v)}">${escapeHtml(etiqueta)}</option>`).join("");
  if ([...sel.options].some(o => o.value === actual)) sel.value = actual;
}
function poblarFiltros() {
  const casas = opcionesUnicas("enQueCasaEsta").sort((a, b) => {
    const ra = ORDEN_CASAS.has(a[0]) ? ORDEN_CASAS.get(a[0]) : 100;
    const rb = ORDEN_CASAS.has(b[0]) ? ORDEN_CASAS.get(b[0]) : 100;
    return (ra - rb) || collator.compare(a[1], b[1]);
  });
  rellenarSelect("fCasa", casas, "Todas");
  rellenarSelect("fUbicacion", opcionesUnicas("ubicacion", ["Congelador", "Frigo", "Fuera"]), "Todas");
  rellenarSelect("fGluten", opcionesUnicas("gluten", OPCIONES_GLUTEN_LECHE), "Todos");
  rellenarSelect("fLeche", opcionesUnicas("leche", OPCIONES_GLUTEN_LECHE), "Todos");
}

// ---------- Render ----------
function datoHtml(label, valor, fecha = false) {
  const texto = fecha ? fechaVisible(valor) : valorVisible(valor);
  return `<div class="dato"><dt>${escapeHtml(label)}</dt><dd>${escapeHtml(texto)}</dd></div>`;
}
function claseCasa(casa) {
  const c = norm(casa);
  if (c === "castefa") return "casa-castefa";
  if (c === "comarruga") return "casa-comarruga";
  if (c === "3" || c === "4") return "casa-naranja";
  return "";
}
function tarjetaHtml(registro, indice) {
  const ocupado = esSi(registro.ocupado);
  const clase = `${ocupado ? "ocupado" : "libre"} ${claseCasa(registro.enQueCasaEsta)}`;
  const comida = ocupado ? valorVisible(registro.comida, "Contenido sin indicar") : "Tupper libre";
  const foto = normalizarRutaFoto(registro.fotoTupper);
  const marcoFoto = foto
    ? `<div class="foto-marco" data-foto-marco="${indice}"><span class="cargando-foto">Cargando foto…</span></div>`
    : `<div class="foto-marco"><span class="sin-foto">Sin foto</span></div>`;

  return `
    <article class="tupper ${clase}" data-indice="${indice}">
      ${marcoFoto}
      <div class="vista-previa-tupper">
        <p class="comida-principal ${ocupado ? "" : "vacio"}">${escapeHtml(comida)}</p>
        <h3>${escapeHtml(valorVisible(registro.nombreTupper, "Tupper sin nombre"))}</h3>
        <div class="badges">
          <span class="badge">En q casa está: ${escapeHtml(valorVisible(registro.enQueCasaEsta))}</span>
          <span class="badge estado">Ubicación: ${escapeHtml(ocupado ? valorVisible(registro.ubicacion) : "Libre")}</span>
        </div>
        <button class="editar" type="button" data-editar="${indice}">Ver / editar tupper</button>
      </div>
    </article>`;
}
function actualizarResumen(visibles) {
  const ocupados = G.registros.filter(r => esSi(r.ocupado)).length;
  $("totalTuppers").textContent = String(G.registros.length);
  $("totalOcupados").textContent = String(ocupados);
  $("totalLibres").textContent = String(G.registros.length - ocupados);
  $("totalVisibles").textContent = String(visibles);
}
function render() {
  const filtros = leerFiltros();
  const visibles = G.registros
    .map((registro, indice) => ({ registro, indice }))
    .filter(({ registro }) => registroCoincide(registro, filtros))
    .sort((a, b) => compararCasaYUbicacion(a.registro, b.registro));

  actualizarResumen(visibles.length);
  const lista = $("listaTuppers");
  if (!visibles.length) {
    lista.innerHTML = `<div class="vacio-lista">No hay tuppers que coincidan con los filtros.</div>`;
    return;
  }
  lista.innerHTML = visibles.map(({ registro, indice }) => tarjetaHtml(registro, indice)).join("");
  cargarFotosRenderizadas(visibles);
}
async function cargarFotosRenderizadas(visibles) {
  for (const { registro, indice } of visibles) {
    const marco = document.querySelector(`[data-foto-marco="${indice}"]`);
    if (!marco || !registro.fotoTupper) continue;
    try {
      const foto = await resolverFoto(registro.fotoTupper);
      if (!marco.isConnected) continue;
      if (!foto || !foto.miniatura) {
        marco.innerHTML = `<span class="sin-foto">Foto no encontrada</span>`;
        continue;
      }
      marco.innerHTML = `<img src="${escapeHtml(foto.miniatura)}" alt="Foto de ${escapeHtml(registro.nombreTupper)}" loading="lazy" />`
        + (foto.completa ? `<button type="button" data-abrir-foto="${indice}" aria-label="Abrir foto"></button>` : "");
      marco.dataset.urlCompleta = foto.completa || foto.miniatura;
    } catch (error) {
      if (marco.isConnected) marco.innerHTML = `<span class="sin-foto">No se pudo cargar</span>`;
      console.warn("No se pudo cargar la foto", registro.fotoTupper, error);
    }
  }
}

// ---------- Editor ----------
function normalizarUsoHistorico(uso) {
  const formatoNuevo = Object.prototype.hasOwnProperty.call(uso || {}, "enQueCasaEsta");
  return {
    guardadoEn: uso && uso.guardadoEn ? uso.guardadoEn : "",
    ocupado: uso && uso.ocupado ? uso.ocupado : "No",
    enQueCasaEsta: formatoNuevo ? (uso.enQueCasaEsta || "") : ((uso && uso.ubicacion) || ""),
    ubicacion: formatoNuevo ? (uso.ubicacion || "") : ((uso && uso.estado) || ""),
    comida: (uso && uso.comida) || "",
    urlCookido: (uso && uso.urlCookido) || "",
    gluten: (uso && uso.gluten) || "",
    leche: (uso && uso.leche) || "",
    fechaPreparacion: (uso && uso.fechaPreparacion) || ""
  };
}
function fechaHoraVisible(valor) {
  if (!valor) return "Fecha no registrada";
  const fecha = new Date(valor);
  if (Number.isNaN(fecha.getTime())) return valorVisible(valor);
  return new Intl.DateTimeFormat("es-ES", { dateStyle: "short", timeStyle: "short" }).format(fecha);
}
function enlaceHistorial(url) {
  const texto = String(url || "").trim();
  if (!texto) return "—";
  if (!/^https?:\/\//i.test(texto)) return escapeHtml(texto);
  return `<a href="${escapeHtml(texto)}" target="_blank" rel="noopener">Abrir enlace</a>`;
}
function historialHtml(registro) {
  const historial = parsearHistorial(registro.historial).map(normalizarUsoHistorico);
  if (!historial.length) return `<p class="historial-vacio">Todavía no hay usos anteriores guardados.</p>`;
  return historial.map((uso, indice) => `
    <article class="uso-anterior">
      <div class="uso-anterior-cabecera">
        <strong>Uso anterior ${indice + 1}</strong>
        <span>${escapeHtml(fechaHoraVisible(uso.guardadoEn))}</span>
      </div>
      <dl class="datos-historial">
        ${datoHtml("Ocupado", canonSiNo(uso.ocupado, false))}
        ${datoHtml("En q casa está", uso.enQueCasaEsta)}
        ${datoHtml("Ubicación", uso.ubicacion)}
        ${datoHtml("Comida", uso.comida)}
        <div class="dato"><dt>Url cookido</dt><dd>${enlaceHistorial(uso.urlCookido)}</dd></div>
        ${datoHtml("Gluten", uso.gluten)}
        ${datoHtml("Leche", uso.leche)}
        ${datoHtml("Fecha preparación", uso.fechaPreparacion, true)}
      </dl>
    </article>`).join("");
}
function prepararHistorial(registro) {
  $("historialUsos").innerHTML = historialHtml(registro);
  $("panelHistorial").hidden = true;
  $("btnHistorial").setAttribute("aria-expanded", "false");
  $("btnHistorial").textContent = "Mostrar los 3 usos anteriores";
}
function alternarHistorial() {
  const panel = $("panelHistorial");
  const mostrar = panel.hidden;
  panel.hidden = !mostrar;
  $("btnHistorial").setAttribute("aria-expanded", mostrar ? "true" : "false");
  $("btnHistorial").textContent = mostrar ? "Ocultar usos anteriores" : "Mostrar los 3 usos anteriores";
}
function crearInstantanea(registro) {
  return {
    guardadoEn: new Date().toISOString(),
    ocupado: registro.ocupado || "No",
    enQueCasaEsta: registro.enQueCasaEsta || "",
    ubicacion: registro.ubicacion || "",
    comida: registro.comida || "",
    urlCookido: registro.urlCookido || "",
    gluten: registro.gluten || "",
    leche: registro.leche || "",
    fechaPreparacion: registro.fechaPreparacion || ""
  };
}
function abrirEditor(indice) {
  const registro = G.registros[indice];
  if (!registro) return;
  $("editorIndice").value = String(indice);
  $("tituloEditor").textContent = valorVisible(registro.nombreTupper, "Tupper");
  $("datosFijos").innerHTML = CAMPOS_FIJOS.map(col => datoHtml(col.label, registro[col.key], col.key === "fechaCompraTupper")).join("");

  $("eOcupado").value = esSi(registro.ocupado) ? "si" : "no";
  const casa = String(registro.enQueCasaEsta || "").trim();
  const predefinida = CASAS_PREDEFINIDAS.find(v => norm(v) === norm(casa));
  $("eCasaTipo").value = predefinida || "otra";
  $("eCasaLibre").value = predefinida ? "" : casa;
  $("eUbicacion").value = ["Congelador", "Frigo", "Fuera"].find(v => norm(v) === norm(registro.ubicacion)) || "";
  $("eComida").value = registro.comida || "";
  $("eUrlCookido").value = registro.urlCookido || "";
  $("eGluten").value = OPCIONES_GLUTEN_LECHE.find(v => norm(v) === norm(registro.gluten)) || "";
  $("eLeche").value = OPCIONES_GLUTEN_LECHE.find(v => norm(v) === norm(registro.leche)) || "";
  $("eFechaPreparacion").value = fechaParaInput(registro.fechaPreparacion);
  prepararHistorial(registro);
  actualizarEditor();

  const dialogo = $("editor");
  if (typeof dialogo.showModal === "function") dialogo.showModal();
  else dialogo.setAttribute("open", "");
}
function cerrarEditor() {
  const dialogo = $("editor");
  if (typeof dialogo.close === "function") dialogo.close();
  else dialogo.removeAttribute("open");
}
function actualizarEditor() {
  const ocupado = $("eOcupado").value === "si";
  const otraCasa = $("eCasaTipo").value === "otra";
  $("campoCasaLibre").hidden = !otraCasa;
  $("eCasaLibre").required = otraCasa;
  $("contenidoEditor").disabled = !ocupado;
  $("eUbicacion").required = ocupado;
  $("eComida").required = ocupado;
  $("eGluten").required = ocupado;
  $("eLeche").required = ocupado;
  $("eFechaPreparacion").required = ocupado;
  // Url cookido es siempre opcional.
  $("eUrlCookido").required = false;
  if (ocupado && !$("eFechaPreparacion").value) $("eFechaPreparacion").value = hoyIso();
}
async function guardarEditor(evento) {
  evento.preventDefault();
  const form = $("formEditor");
  if (!form.reportValidity()) return;
  const indice = Number($("editorIndice").value);
  const registro = G.registros[indice];
  if (!registro) return;

  const copiaAnterior = Object.assign({}, registro);
  const ocupado = $("eOcupado").value === "si";
  const casa = $("eCasaTipo").value === "otra"
    ? $("eCasaLibre").value.trim()
    : $("eCasaTipo").value;
  const nuevo = {
    ocupado: ocupado ? "Sí" : "No",
    enQueCasaEsta: casa,
    ubicacion: ocupado ? $("eUbicacion").value : "",
    comida: ocupado ? $("eComida").value.trim() : "",
    urlCookido: ocupado ? $("eUrlCookido").value.trim() : "",
    gluten: ocupado ? $("eGluten").value : "",
    leche: ocupado ? $("eLeche").value : "",
    fechaPreparacion: ocupado ? $("eFechaPreparacion").value : ""
  };

  const hayCambios = CAMPOS_HISTORIAL.some(key => String(registro[key] || "").trim() !== String(nuevo[key] || "").trim());
  if (!hayCambios) {
    setEstado("No había cambios que guardar.");
    cerrarEditor();
    return;
  }

  const historial = parsearHistorial(registro.historial);
  registro.historial = serializarHistorial([crearInstantanea(registro), ...historial].slice(0, 3));
  Object.assign(registro, nuevo);

  poblarFiltros();
  render();
  try {
    await guardarDatos();
    cerrarEditor();
  } catch (error) {
    Object.assign(registro, copiaAnterior);
    render();
    setEstado(`Error guardando: ${error.message}`, true);
  }
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
        $("listaTuppers").innerHTML = "";
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
Object.values(FILTROS).forEach(id => {
  $(id).addEventListener("input", render);
  $(id).addEventListener("change", render);
});
$("btnLimpiar").addEventListener("click", limpiarFiltros);
const BOTONES_CASA = [["btnCasaCastefa", "castefa"], ["btnCasaComarruga", "comarruga"]];
function actualizarBotonesCasa() {
  BOTONES_CASA.forEach(([id, casa]) => {
    const btn = $(id);
    if (btn) {
      btn.classList.toggle("activo", casasBoton.has(casa));
      btn.setAttribute("aria-pressed", casasBoton.has(casa) ? "true" : "false");
    }
  });
}
BOTONES_CASA.forEach(([id, casa]) => {
  const btn = $(id);
  if (!btn) return;
  btn.addEventListener("click", () => {
    if (casasBoton.has(casa)) casasBoton.delete(casa);
    else casasBoton.add(casa);
    actualizarBotonesCasa();
    render();
  });
});
$("btnRecargar").addEventListener("click", () => cargarDatos().catch(error => setEstado(`Error: ${error.message}`, true)));
$("listaTuppers").addEventListener("click", event => {
  const editar = event.target.closest("[data-editar]");
  if (editar) { abrirEditor(Number(editar.dataset.editar)); return; }
  const abrirFoto = event.target.closest("[data-abrir-foto]");
  if (abrirFoto) {
    const marco = abrirFoto.closest(".foto-marco");
    const url = marco && marco.dataset.urlCompleta;
    if (url) window.open(url, "_blank", "noopener");
  }
});
$("eOcupado").addEventListener("change", actualizarEditor);
$("eCasaTipo").addEventListener("change", actualizarEditor);
$("btnHistorial").addEventListener("click", alternarHistorial);
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

window.tuppersAuth = { msalInstance, getToken, resetMsal, G };
