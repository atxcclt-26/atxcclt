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
// Foto tupper;Ocupado;Ubicación;Estado;Comida;Url cookido;Gluten;Leche;Fecha preparación;Historial
//
// La web solo modifica: Ocupado, Ubicación, Estado, Comida, Url cookido, Gluten, Leche y Fecha preparación.
// Historial es una columna interna con las tres instantáneas anteriores del tupper.
// El resto de campos se modifica directamente en OneDrive.

// ====== Ajustes ======
const CARPETA_RAIZ = "ZZZ_SG_Tuppers";
const FILE_DATOS = "tuppers.csv";
const GRAPH = "https://graph.microsoft.com/v1.0";
// =====================

// Misma arquitectura de autenticación MSAL que la app de referencia.
const REDIRECT_PAGE = window.location.origin + window.location.pathname;
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
  { key: "nombreTupper", label: "Nombre tupper", editable: false, aliases: ["nombre tupper"] },
  { key: "comentariosTupper", label: "Comentarios tupper", editable: false, aliases: ["comentarios tupper", "comentarios tuppers", "comentario tupper"] },
  { key: "materialTapa", label: "Material tapa", editable: false, aliases: ["material tapa"] },
  { key: "materialTupper", label: "Material tupper", editable: false, aliases: ["material tupper"] },
  { key: "fechaCompraTupper", label: "Fecha compra tupper", editable: false, aliases: ["fecha compra tupper"] },
  { key: "fotoTupper", label: "Foto tupper", editable: false, aliases: ["foto tupper", "foto"] },
  { key: "ocupado", label: "Ocupado", editable: true, aliases: ["ocupado"] },
  { key: "ubicacion", label: "Ubicación", editable: true, aliases: ["ubicacion"] },
  { key: "estado", label: "Estado", editable: true, aliases: ["estado"] },
  { key: "comida", label: "Comida", editable: true, aliases: ["comida"] },
  { key: "urlCookido", label: "Url cookido", editable: true, aliases: ["url cookido", "url cookidoo", "cookido", "cookidoo"] },
  { key: "gluten", label: "Gluten", editable: true, aliases: ["gluten"] },
  { key: "leche", label: "Leche", editable: true, aliases: ["leche"] },
  { key: "fechaPreparacion", label: "Fecha preparación", editable: true, aliases: ["fecha preparacion"] }
];

const COLUMNA_HISTORIAL = { key: "_historialCsv", label: "Historial", aliases: ["historial"] };
const COLUMNAS_CSV = [...COLUMNAS, COLUMNA_HISTORIAL];
const CAMPOS_FIJOS = COLUMNAS.filter(c => !c.editable);
const CAMPOS_EDITABLES = COLUMNAS.filter(c => c.editable);
const CLAVES_USO = CAMPOS_EDITABLES.map(c => c.key);
const ORDEN_UBICACIONES = new Map([["cm", 0], ["cs", 1], ["3", 2], ["4", 3]]);
const ORDEN_ESTADOS = new Map([["congelador", 0], ["frigo", 1], ["fuera", 2], ["", 3]]);
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
function ahoraIsoLocal() {
  const d = new Date();
  const p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}
function fechaHoraVisible(valor) {
  const v = String(valor || "").trim();
  const m = v.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?/);
  if (!m) return valorVisible(v);
  return `${m[3]}/${m[2]}/${m[1]} ${m[4]}:${m[5]}`;
}
function urlHttpSegura(valor) {
  const v = String(valor || "").trim();
  return /^https?:\/\//i.test(v) ? v : "";
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
function crearMapaCabecera(cabecera, columnas, usarPosicionPorDefecto) {
  const mapa = new Map();
  cabecera.forEach((nombre, indice) => mapa.set(normalizarCabecera(nombre), indice));
  return columnas.map((col, indicePorDefecto) => {
    for (const alias of col.aliases) {
      const encontrado = mapa.get(normalizarCabecera(alias));
      if (encontrado !== undefined) return encontrado;
    }
    return usarPosicionPorDefecto ? indicePorDefecto : -1;
  });
}
function normalizarHistorial(valor) {
  if (!valor) return [];
  try {
    const parsed = typeof valor === "string" ? JSON.parse(valor) : valor;
    if (!Array.isArray(parsed)) return [];
    return parsed.slice(0, 3).map(item => ({
      guardado: String(item.guardado || ""),
      ocupado: canonSiNo(item.ocupado, false),
      ubicacion: String(item.ubicacion || ""),
      estado: String(item.estado || ""),
      comida: String(item.comida || ""),
      urlCookido: String(item.urlCookido || item.urlCookidoo || ""),
      gluten: canonSiNo(item.gluten, true),
      leche: canonSiNo(item.leche, true),
      fechaPreparacion: String(item.fechaPreparacion || "")
    }));
  } catch (error) {
    console.warn("Historial no válido; se ignora.", error);
    return [];
  }
}
function instantaneaUso(registro) {
  return {
    guardado: ahoraIsoLocal(),
    ocupado: canonSiNo(registro.ocupado, false),
    ubicacion: String(registro.ubicacion || "").trim(),
    estado: String(registro.estado || "").trim(),
    comida: String(registro.comida || "").trim(),
    urlCookido: String(registro.urlCookido || "").trim(),
    gluten: canonSiNo(registro.gluten, true),
    leche: canonSiNo(registro.leche, true),
    fechaPreparacion: String(registro.fechaPreparacion || "").trim()
  };
}
function anadirAlHistorial(registro, usoAnterior) {
  const historial = normalizarHistorial(registro._historial);
  registro._historial = [instantaneaUso(usoAnterior), ...historial].slice(0, 3);
}
function registroDesdeFila(fila, mapaIndices, indiceOriginal) {
  const registro = { _indiceOriginal: indiceOriginal };
  COLUMNAS.forEach((col, i) => registro[col.key] = String(mapaIndices[i] >= 0 ? (fila[mapaIndices[i]] ?? "") : "").trim());
  const indiceHistorial = mapaIndices[COLUMNAS.length];
  registro._historial = normalizarHistorial(indiceHistorial >= 0 ? fila[indiceHistorial] : "");
  registro.ocupado = canonSiNo(registro.ocupado, false);
  registro.gluten = canonSiNo(registro.gluten, true);
  registro.leche = canonSiNo(registro.leche, true);
  if (!esSi(registro.ocupado)) {
    registro.estado = "";
    registro.comida = "";
    registro.urlCookido = "";
    registro.gluten = "";
    registro.leche = "";
    registro.fechaPreparacion = "";
  }
  return registro;
}
function filasParaGuardar() {
  const cabecera = COLUMNAS_CSV.map(c => c.label);
  const datos = G.registros
    .slice()
    .sort((a, b) => a._indiceOriginal - b._indiceOriginal)
    .map(registro => [
      ...COLUMNAS.map(col => String(registro[col.key] ?? "").trim()),
      JSON.stringify(normalizarHistorial(registro._historial))
    ]);
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
    const cabecera = tieneCabecera ? filas[0] : COLUMNAS_CSV.map(c => c.label);
    const mapa = crearMapaCabecera(cabecera, COLUMNAS_CSV, !tieneCabecera);
    const datos = tieneCabecera ? filas.slice(1) : filas;
    G.registros = datos.map((fila, indice) => registroDesdeFila(fila, mapa, indice));
  }
  G.fotoCache.clear();
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
  ubicacion: "fUbicacion",
  estado: "fEstado",
  comida: "fComida",
  gluten: "fGluten",
  leche: "fLeche",
  fechaPreparacion: "fFechaPreparacion"
};
function leerFiltros() {
  return Object.fromEntries(Object.entries(FILTROS).map(([key, id]) => [key, $(id).value]));
}
function coincideTexto(valor, filtro) { return !norm(filtro) || norm(valor).includes(norm(filtro)); }
function coincideFecha(valor, filtro) { return !filtro || fechaParaInput(valor) === filtro; }
function coincideSiNo(valor, filtro) {
  if (!filtro) return true;
  if (!norm(valor)) return false;
  return (esSi(valor) ? "si" : "no") === filtro;
}
function registroCoincide(registro, filtros) {
  return coincideTexto(registro.nombreTupper, filtros.nombreTupper)
    && coincideTexto(registro.comentariosTupper, filtros.comentariosTupper)
    && coincideTexto(registro.materialTapa, filtros.materialTapa)
    && coincideTexto(registro.materialTupper, filtros.materialTupper)
    && coincideFecha(registro.fechaCompraTupper, filtros.fechaCompraTupper)
    && coincideTexto(registro.fotoTupper, filtros.fotoTupper)
    && coincideSiNo(registro.ocupado, filtros.ocupado)
    && coincideTexto(registro.ubicacion, filtros.ubicacion)
    && coincideTexto(registro.estado, filtros.estado)
    && coincideTexto(registro.comida, filtros.comida)
    && coincideSiNo(registro.gluten, filtros.gluten)
    && coincideSiNo(registro.leche, filtros.leche)
    && coincideFecha(registro.fechaPreparacion, filtros.fechaPreparacion);
}
function compararUbicacion(a, b) {
  const na = norm(a.ubicacion), nb = norm(b.ubicacion);
  const ra = ORDEN_UBICACIONES.has(na) ? ORDEN_UBICACIONES.get(na) : 100;
  const rb = ORDEN_UBICACIONES.has(nb) ? ORDEN_UBICACIONES.get(nb) : 100;
  if (ra !== rb) return ra - rb;
  const porTexto = collator.compare(a.ubicacion || "", b.ubicacion || "");
  if (porTexto) return porTexto;
  const ea = ORDEN_ESTADOS.has(norm(a.estado)) ? ORDEN_ESTADOS.get(norm(a.estado)) : 50;
  const eb = ORDEN_ESTADOS.has(norm(b.estado)) ? ORDEN_ESTADOS.get(norm(b.estado)) : 50;
  if (ea !== eb) return ea - eb;
  const porEstado = collator.compare(a.estado || "", b.estado || "");
  if (porEstado) return porEstado;
  return collator.compare(a.nombreTupper || "", b.nombreTupper || "");
}
function limpiarFiltros() {
  Object.values(FILTROS).forEach(id => { $(id).value = ""; });
  render();
}

// ---------- Render ----------
function datoHtml(label, valor, fecha = false) {
  const texto = fecha ? fechaVisible(valor) : valorVisible(valor);
  return `<div class="dato"><dt>${escapeHtml(label)}</dt><dd>${escapeHtml(texto)}</dd></div>`;
}
function tarjetaHtml(registro, indice) {
  const ocupado = esSi(registro.ocupado);
  const clase = ocupado ? "ocupado" : "libre";
  const comida = ocupado ? valorVisible(registro.comida, "Contenido sin indicar") : "Tupper libre";
  const foto = normalizarRutaFoto(registro.fotoTupper);
  const marcoFoto = foto
    ? `<div class="foto-marco" data-foto-marco="${indice}"><span class="cargando-foto">Cargando foto…</span></div>`
    : `<div class="foto-marco"><span class="sin-foto">Sin foto</span></div>`;

  return `
    <article class="tupper ${clase}" data-indice="${indice}">
      ${marcoFoto}
      <div>
        <div class="cabecera-tupper">
          <div>
            <h3>${escapeHtml(valorVisible(registro.nombreTupper, "Tupper sin nombre"))}</h3>
            <p class="comida-principal ${ocupado ? "" : "vacio"}">${escapeHtml(comida)}</p>
          </div>
          <button class="editar" type="button" data-editar="${indice}">Editar contenido</button>
        </div>
        <div class="badges">
          <span class="badge">Ubicación: ${escapeHtml(valorVisible(registro.ubicacion))}</span>
          <span class="badge estado">Estado: ${escapeHtml(ocupado ? valorVisible(registro.estado) : "Libre")}</span>
          <span class="badge fecha">Preparación: ${escapeHtml(ocupado ? fechaVisible(registro.fechaPreparacion) : "—")}</span>
        </div>
        <dl class="datos-tupper">
          ${datoHtml("Comentarios tupper", registro.comentariosTupper)}
          ${datoHtml("Material tapa", registro.materialTapa)}
          ${datoHtml("Material tupper", registro.materialTupper)}
          ${datoHtml("Fecha compra tupper", registro.fechaCompraTupper, true)}
          ${datoHtml("Foto tupper", registro.fotoTupper)}
          ${datoHtml("Ocupado", canonSiNo(registro.ocupado, false))}
          ${datoHtml("Gluten", ocupado ? registro.gluten : "")}
          ${datoHtml("Leche", ocupado ? registro.leche : "")}
        </dl>
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
    .sort((a, b) => compararUbicacion(a.registro, b.registro));

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
function usoHistorialHtml(uso, posicion) {
  const ocupado = esSi(uso.ocupado);
  const url = urlHttpSegura(uso.urlCookido);
  const urlHtml = url
    ? `<a href="${escapeHtml(url)}" target="_blank" rel="noopener">Abrir receta</a>`
    : escapeHtml(valorVisible(uso.urlCookido));
  return `<article class="uso-anterior">
    <div class="uso-anterior-cabecera">
      <strong>Uso anterior ${posicion}</strong>
      <span>${escapeHtml(fechaHoraVisible(uso.guardado))}</span>
    </div>
    <dl class="datos-historial">
      ${datoHtml("Ocupado", canonSiNo(uso.ocupado, false))}
      ${datoHtml("Ubicación", uso.ubicacion)}
      ${datoHtml("Estado", ocupado ? uso.estado : "")}
      ${datoHtml("Comida", ocupado ? uso.comida : "")}
      <div class="dato"><dt>Url cookido</dt><dd>${ocupado ? urlHtml : "—"}</dd></div>
      ${datoHtml("Gluten", ocupado ? uso.gluten : "")}
      ${datoHtml("Leche", ocupado ? uso.leche : "")}
      ${datoHtml("Fecha preparación", ocupado ? uso.fechaPreparacion : "", true)}
    </dl>
  </article>`;
}
function renderHistorial(registro) {
  const historial = normalizarHistorial(registro._historial).slice(0, 3);
  $("historialUsos").innerHTML = historial.length
    ? historial.map((uso, indice) => usoHistorialHtml(uso, indice + 1)).join("")
    : `<p class="historial-vacio">Todavía no hay usos anteriores guardados.</p>`;
}
function abrirEditor(indice) {
  const registro = G.registros[indice];
  if (!registro) return;
  $("editorIndice").value = String(indice);
  $("tituloEditor").textContent = `Editar contenido · ${valorVisible(registro.nombreTupper, "Tupper")}`;
  $("datosFijos").innerHTML = CAMPOS_FIJOS.map(col => datoHtml(col.label, registro[col.key], col.key === "fechaCompraTupper")).join("");

  $("eOcupado").value = esSi(registro.ocupado) ? "si" : "no";
  const ubicacion = String(registro.ubicacion || "").trim();
  const predefinidas = ["CM", "CS", "3", "4"];
  const predefinida = predefinidas.find(v => norm(v) === norm(ubicacion));
  $("eUbicacionTipo").value = predefinida || "otra";
  $("eUbicacionLibre").value = predefinida ? "" : ubicacion;
  $("eEstado").value = ["Congelador", "Frigo", "Fuera"].find(v => norm(v) === norm(registro.estado)) || "";
  $("eComida").value = registro.comida || "";
  $("eUrlCookido").value = registro.urlCookido || "";
  $("eGluten").value = esSi(registro.gluten) ? "Sí" : (norm(registro.gluten) === "no" ? "No" : "");
  $("eLeche").value = esSi(registro.leche) ? "Sí" : (norm(registro.leche) === "no" ? "No" : "");
  $("eFechaPreparacion").value = fechaParaInput(registro.fechaPreparacion);
  renderHistorial(registro);
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
  const otraUbicacion = $("eUbicacionTipo").value === "otra";
  $("campoUbicacionLibre").hidden = !otraUbicacion;
  $("eUbicacionLibre").required = otraUbicacion;
  $("contenidoEditor").disabled = !ocupado;
  $("eEstado").required = ocupado;
  $("eComida").required = ocupado;
  $("eUrlCookido").required = ocupado;
  $("eGluten").required = ocupado;
  $("eLeche").required = ocupado;
  $("eFechaPreparacion").required = ocupado;
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
  const ubicacion = $("eUbicacionTipo").value === "otra"
    ? $("eUbicacionLibre").value.trim()
    : $("eUbicacionTipo").value;

  registro.ocupado = ocupado ? "Sí" : "No";
  registro.ubicacion = ubicacion;
  if (ocupado) {
    registro.estado = $("eEstado").value;
    registro.comida = $("eComida").value.trim();
    registro.urlCookido = $("eUrlCookido").value.trim();
    registro.gluten = $("eGluten").value;
    registro.leche = $("eLeche").value;
    registro.fechaPreparacion = $("eFechaPreparacion").value;
  } else {
    registro.estado = "";
    registro.comida = "";
    registro.urlCookido = "";
    registro.gluten = "";
    registro.leche = "";
    registro.fechaPreparacion = "";
  }

  const haCambiado = CLAVES_USO.some(key => String(copiaAnterior[key] ?? "").trim() !== String(registro[key] ?? "").trim());
  if (haCambiado) anadirAlHistorial(registro, copiaAnterior);

  render();
  try {
    if (haCambiado) await guardarDatos();
    else setEstado("No había cambios que guardar.");
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
$("eUbicacionTipo").addEventListener("change", actualizarEditor);
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
