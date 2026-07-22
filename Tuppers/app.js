// app.js — PSCM (listas de tareas con validación por check o foto, sobre OneDrive)
//
// Estructura en OneDrive:
//   ZZZ_SG_pscm/                      ← carpeta en la raíz (sin distinguir mayúsculas)
//     pscm.csv                        ← plantilla: PISO;ZONA;ACCION  (headerless, ";")
//     2607151032/                     ← una subcarpeta por ejecución (YYMMDDHHMM)
//       2607151032.csv                ← estado: PISO;ZONA;ACCION;ESTADO;FOTO;HORA;USUARIO;MARCA
//       003_Foto_armario_izquierda.jpg  ← fotos de esa ejecución
//
// Si la ACCION empieza por "foto" → se valida haciendo una foto; si no → check.

// ====== AJUSTES ======
const CARPETA_RAIZ = "ZZZ_SG_pscm";   // carpeta en la raíz de OneDrive (case-insensitive)
const FILE_PLANTILLA = "pscm.csv";    // plantilla de tareas (3 columnas, sin cabecera)
const RE_EJECUCION = /^\d{10}$/;      // nombre de ejecución: YYMMDDHHMM
// =====================

// 1. MSAL (misma arquitectura que "Revisión de cargos")
const esMovil = /Android|iPhone|iPad|iPod/i.test(navigator.userAgent) ||
                (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
const REDIRECT_PAGE = window.location.origin + window.location.pathname;
const msalConfig = {
  auth: { clientId: "7765de16-766b-4051-b73d-d7167f5897dc", redirectUri: REDIRECT_PAGE },
  cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
  system: { loggerOptions: { logLevel: msal.LogLevel.Warning, loggerCallback: (l, m, p) => { if (!p) console.log(`[MSAL][${l}]`, m); } } }
};
const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["User.Read", "Files.ReadWrite"] };
const fileScopes   = { scopes: ["Files.ReadWrite"] };

// Estado global
const G = {
  driveId: null,           // drive de la carpeta (null = mi OneDrive; con valor = carpeta compartida)
  folderId: null,          // id de ZZZ_SG_pscm
  plantilla: [],           // [[piso, zona, accion], ...] leída de pscm.csv
  listaHoy: null,          // nombre YYMMDDHHMM de la lista ya creada hoy (o null)
  ejecucion: null          // { nombre, subfolderId, stateName, rows: [[piso,zona,accion,estado,foto,hora],...] }
};
// Estado de la interfaz del árbol
const UI = {
  verHechas: true,         // selector "Ver realizadas SÍ/NO"
  editables: new Set()     // idx de tareas hechas desbloqueadas con "Editar"
};
let inicializado = false;
let guardando = false, guardarDeNuevo = false;
let fotoPendiente = null;  // idx de la fila cuya foto se está pidiendo
let syncTimer = null;      // polling de sincronización multi-dispositivo
let minTimer = null;       // refresco del "hace x min" en tareas hechas
const SYNC_MS = 20000;     // intervalo del polling (ms)

// ---------- UI helpers ----------
function setEstado(msg, err = false) {
  const el = document.getElementById("estado-carga");
  if (el) { el.textContent = msg; el.classList.toggle("error", !!err); }
}
function escapeHtml(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
function norm(v) { return (v || "").trim().toLowerCase(); }
function mostrarVista(id) {
  document.querySelectorAll(".vista").forEach(v => v.classList.toggle("activa", v.id === id));
}
function esFoto(accion) { return norm(accion).startsWith("foto"); }
// Nombre de fichero seguro a partir de la acción (para la foto)
function slug(s) {
  return (s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 60) || "foto";
}
// Nombre de ejecución: YYMMDDHHMM (hora local)
function nombreEjecucion(d = new Date()) {
  const p = n => String(n).padStart(2, "0");
  return String(d.getFullYear()).slice(2) + p(d.getMonth() + 1) + p(d.getDate()) + p(d.getHours()) + p(d.getMinutes());
}
function nombreLegible(n) { // "2607151032" → "15/07/26 10:32"
  return `${n.slice(4, 6)}/${n.slice(2, 4)}/${n.slice(0, 2)} ${n.slice(6, 8)}:${n.slice(8, 10)}`;
}
function hoyPrefijo() { // YYMMDD de hoy (hora local)
  const p = n => String(n).padStart(2, "0"); const d = new Date();
  return String(d.getFullYear()).slice(2) + p(d.getMonth() + 1) + p(d.getDate());
}
function fechaLarga(n) { // "2607151032" → "15/07/2026 10:32"
  return `${n.slice(4, 6)}/${n.slice(2, 4)}/20${n.slice(0, 2)} ${n.slice(6, 8)}:${n.slice(8, 10)}`;
}
function haceDias(n) { // días naturales transcurridos desde la ejecución
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const f = new Date(2000 + Number(n.slice(0, 2)), Number(n.slice(2, 4)) - 1, Number(n.slice(4, 6)));
  const dias = Math.round((hoy - f) / 86400000);
  return dias <= 0 ? "hoy" : dias === 1 ? "ayer" : `hace ${dias} días`;
}
// Primera palabra del nombre de la cuenta activa (fallback: parte local del email)
function nombreUsuario() {
  const a = msalInstance.getActiveAccount();
  const n = (a && (a.name || a.username)) || "";
  return n.split(/[\s@]+/)[0] || "";
}
// Marca temporal completa "YYYY-MM-DD HH:MM:SS" (hora local)
function ahoraTS() {
  const d = new Date(), p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}
// Minutos transcurridos desde la validación de una fila (usa MARCA; si falta, fecha de la lista + HORA)
function haceTexto(row, nombreLista) {
  let d = null;
  const m = (row[7] || "").match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/);
  if (m) d = new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);
  else if (row[5]) { // filas antiguas sin MARCA: fecha de la lista + HORA
    const h = row[5].split(":");
    d = new Date(2000 + +nombreLista.slice(0, 2), +nombreLista.slice(2, 4) - 1, +nombreLista.slice(4, 6), +h[0] || 0, +h[1] || 0, +h[2] || 0);
  }
  if (!d || isNaN(d)) return "";
  const mins = Math.max(0, Math.floor((Date.now() - d.getTime()) / 60000));
  if (mins < 1) return "ahora mismo";
  if (mins < 60) return `hace ${mins} min`;
  if (mins < 1440) return `hace ${Math.floor(mins / 60)} h ${mins % 60} min`;
  return `hace ${Math.floor(mins / 1440)} día(s)`;
}

// ---------- CSV (";" — mismo parser que la app hermana) ----------
function parseCSV(text, delim = ";") {
  const rows = []; let row = [], field = "", i = 0, q = false;
  while (i < text.length) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { field += '"'; i += 2; continue; } q = false; i++; continue; } field += c; i++; continue; }
    if (c === '"') { q = true; i++; continue; }
    if (c === delim) { row.push(field); field = ""; i++; continue; }
    if (c === "\r") { i++; continue; }
    if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; i++; continue; }
    field += c; i++;
  }
  if (field.length || row.length) { row.push(field); rows.push(row); }
  return rows.filter(r => r.some(f => (f || "").trim() !== ""));
}
function serialF(v, d) { v = v == null ? "" : String(v); return /["\n\r]/.test(v) || v.includes(d) ? '"' + v.replace(/"/g, '""') + '"' : v; }
function serializeCSV(rows, d = ";") { return rows.map(r => r.map(f => serialF(f, d)).join(d)).join("\r\n"); }

// ---------- Graph / OneDrive ----------
const GRAPH = "https://graph.microsoft.com/v1.0";
function itemBase() { return G.driveId ? `${GRAPH}/drives/${G.driveId}` : `${GRAPH}/me/drive`; }

// Localiza ZZZ_SG_pscm en la raíz (soporta carpeta compartida → driveId, como la app hermana)
async function localizarCarpeta(token) {
  const H = { Authorization: `Bearer ${token}` };
  const r = await fetch(`${GRAPH}/me/drive/root/children?$top=200`, { headers: H });
  if (!r.ok) {
    if (r.status === 401 || r.status === 403)
      throw new Error(`Sin permiso (HTTP ${r.status}). ¿Añadiste "Files.ReadWrite" en Azure? resetMsal() y reentra.`);
    throw new Error("raíz " + r.status + " " + (await r.text()));
  }
  const hijos = (await r.json()).value || [];
  const objetivo = CARPETA_RAIZ.toLowerCase();
  const c = hijos.find(it => (it.folder || it.remoteItem) && (it.name || "").toLowerCase() === objetivo);
  if (!c) {
    setEstado(`No existe la carpeta "${CARPETA_RAIZ}" en la raíz de tu OneDrive.`, true);
    return false;
  }
  if (c.remoteItem) { G.driveId = c.remoteItem.parentReference && c.remoteItem.parentReference.driveId; G.folderId = c.remoteItem.id; }
  else { G.driveId = null; G.folderId = c.id; }
  return true;
}

async function gFetch(url, opts = {}) {
  const token = await getToken(fileScopes);
  opts.headers = Object.assign({ Authorization: `Bearer ${token}` }, opts.headers || {});
  const res = await fetch(url, opts);
  return res;
}

// Descarga un fichero (por ruta relativa a la carpeta raíz de la app)
async function descargarTexto(ruta) {
  const res = await gFetch(`${itemBase()}/items/${G.folderId}:/${encodeURI(ruta)}:/content`);
  if (!res.ok) throw new Error(`descarga "${ruta}" → ${res.status}`);
  return new TextDecoder("utf-8").decode(await res.arrayBuffer()).replace(/^\uFEFF/, "");
}

// Sube (crea o reemplaza) un fichero pequeño por ruta relativa a la carpeta raíz
async function subirContenido(ruta, body, contentType) {
  const res = await gFetch(`${itemBase()}/items/${G.folderId}:/${encodeURI(ruta)}:/content`, {
    method: "PUT", headers: { "Content-Type": contentType }, body
  });
  if (!res.ok) throw new Error(`subida "${ruta}" → ${res.status} ${await res.text()}`);
  return await res.json();
}

// Sube un fichero grande (>3,5 MB) por sesión de subida, en trozos de 5 MiB (múltiplo de 320 KiB)
async function subirGrande(ruta, blob, contentType) {
  const rs = await gFetch(`${itemBase()}/items/${G.folderId}:/${encodeURI(ruta)}:/createUploadSession`, {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace" } })
  });
  if (!rs.ok) throw new Error(`sesión "${ruta}" → ${rs.status}`);
  const { uploadUrl } = await rs.json();
  const CHUNK = 5 * 1024 * 1024, total = blob.size;
  let pos = 0, ultimo = null;
  while (pos < total) {
    const fin = Math.min(pos + CHUNK, total);
    const res = await fetch(uploadUrl, {
      method: "PUT",
      headers: { "Content-Range": `bytes ${pos}-${fin - 1}/${total}`, "Content-Type": contentType },
      body: blob.slice(pos, fin)
    });
    if (!res.ok && res.status !== 202) throw new Error(`trozo ${pos} → ${res.status}`);
    ultimo = res; pos = fin;
  }
  return await ultimo.json();
}

async function subirFoto(ruta, blob) {
  const tipo = blob.type || "image/jpeg";
  return blob.size <= 3.5 * 1024 * 1024 ? subirContenido(ruta, blob, tipo) : subirGrande(ruta, blob, tipo);
}

// Crea la subcarpeta de la ejecución (si ya existe, la reutiliza)
async function crearSubcarpeta(nombre) {
  const res = await gFetch(`${itemBase()}/items/${G.folderId}/children`, {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: nombre, folder: {}, "@microsoft.graph.conflictBehavior": "fail" })
  });
  if (res.ok) return await res.json();
  if (res.status === 409) { // ya existía (dos ejecuciones en el mismo minuto): la reutilizamos
    const r2 = await gFetch(`${itemBase()}/items/${G.folderId}:/${encodeURI(nombre)}`);
    if (r2.ok) return await r2.json();
  }
  throw new Error(`crear carpeta "${nombre}" → ${res.status} ${await res.text()}`);
}

// Lista las subcarpetas de ejecución (YYMMDDHHMM), más recientes primero
async function listarEjecuciones() {
  const res = await gFetch(`${itemBase()}/items/${G.folderId}/children?$top=200&$orderby=name%20desc`);
  if (!res.ok) throw new Error("listar ejecuciones → " + res.status);
  return ((await res.json()).value || [])
    .filter(it => it.folder && RE_EJECUCION.test(it.name))
    .map(it => it.name)
    .sort((a, b) => b.localeCompare(a));
}

// Lista las fotos (imágenes) de la carpeta de la ejecución, con miniaturas
async function listarFotos(nombre) {
  const res = await gFetch(`${itemBase()}/items/${G.folderId}:/${encodeURI(nombre)}:/children?$expand=thumbnails&$top=200`);
  if (!res.ok) throw new Error("listar fotos → " + res.status);
  return ((await res.json()).value || [])
    .filter(it => it.file && (/^image\//.test(it.file.mimeType || "") || /\.(jpe?g|png|heic|webp)$/i.test(it.name)))
    .sort((a, b) => a.name.localeCompare(b.name));
}

// ---------- Sincronización multi-dispositivo ----------
// Cada dispositivo mantiene sus filas en memoria; sin fusión, el autosave de uno
// machacaría lo validado en el otro (last-write-wins sobre el CSV completo).
// Estrategia: e.dirty = Set de índices modificados localmente y aún no subidos.
// Al fusionar con el CSV remoto, lo local "sucio" gana; el resto adopta lo remoto.

async function descargarEstadoRemoto() {
  const e = G.ejecucion;
  const text = await descargarTexto(`${e.nombre}/${e.stateName}`);
  return parseCSV(text).map(r => {
    while (r.length < 8) r.push("");
    return r.slice(0, 8).map(f => (f || "").trim());
  });
}

// Copia ESTADO/FOTO/HORA/USUARIO/MARCA remotos en las filas no modificadas localmente. Devuelve nº de cambios.
function fusionar(remoto) {
  const e = G.ejecucion;
  let cambios = 0;
  const n = Math.min(e.rows.length, remoto.length);
  for (let i = 0; i < n; i++) {
    if (e.dirty.has(i)) continue;                 // lo modificado aquí tiene prioridad
    const l = e.rows[i], r = remoto[i];
    let dif = false;
    for (let c = 3; c < 8; c++) if (l[c] !== r[c]) { dif = true; break; }
    if (dif) { for (let c = 3; c < 8; c++) l[c] = r[c]; cambios++; }
  }
  return cambios;
}

// Polling / resync manual: baja el CSV remoto y refleja cambios del otro dispositivo.
async function sincronizar() {
  const e = G.ejecucion;
  if (!e || guardando || document.hidden) return;
  try {
    const remoto = await descargarEstadoRemoto();
    if (fusionar(remoto)) {
      renderArbol();
      setEstado("Lista actualizada con cambios de otro dispositivo.");
    }
  } catch (err) { console.warn("Sync:", err.message); }
}
function iniciarSync() {
  detenerSync();
  syncTimer = setInterval(sincronizar, SYNC_MS);
  // refresco periódico del "hace x min" de las tareas hechas
  minTimer = setInterval(() => { if (G.ejecucion && !document.hidden) renderArbol(); }, 60000);
}
function detenerSync() {
  if (syncTimer) { clearInterval(syncTimer); syncTimer = null; }
  if (minTimer) { clearInterval(minTimer); minTimer = null; }
}

// ---------- Guardado del estado de la ejecución (autosave con reintento en cola) ----------
async function guardarEstado() {
  const e = G.ejecucion;
  const modAlEmpezar = e.mod;
  // Fusiona lo remoto antes de subir para no pisar validaciones de otro dispositivo.
  try {
    const remoto = await descargarEstadoRemoto();
    if (fusionar(remoto)) renderArbol();
  } catch (err) { /* p. ej. 404 en el primer guardado de una lista recién creada */ }
  await subirContenido(`${e.nombre}/${e.stateName}`, serializeCSV(e.rows), "text/csv");
  // Solo limpiamos "dirty" si no hubo modificaciones durante la subida (si las hubo,
  // hay otro guardado en cola que las volverá a fusionar y subir).
  if (e.mod === modAlEmpezar) e.dirty.clear();
}
function programarGuardado() {
  if (guardando) { guardarDeNuevo = true; return; }
  guardando = true;
  guardarEstado()
    .then(() => setEstado(`Cambios guardados en ${G.ejecucion.nombre}.`))
    .catch(e => setEstado("Error guardando (los cambios siguen en pantalla, reintenta): " + e.message, true))
    .finally(() => { guardando = false; if (guardarDeNuevo) { guardarDeNuevo = false; programarGuardado(); } });
}

// ---------- Plantilla ----------
async function cargarPlantilla() {
  const text = await descargarTexto(FILE_PLANTILLA);
  // 3 columnas: PISO;ZONA;ACCION (se ignoran ";" finales / columnas extra)
  G.plantilla = parseCSV(text).map(r => [(r[0] || "").trim(), (r[1] || "").trim(), (r[2] || "").trim()])
    .filter(r => r[2] !== "");
  if (!G.plantilla.length) throw new Error(`"${FILE_PLANTILLA}" está vacío o no tiene 3 columnas.`);
}

// ---------- Flujos ----------
// Límite: una lista por día natural. Si ya existe, el botón pasa a "Abrir lista en marcha".
async function actualizarMenu() {
  const btn = document.getElementById("btnNueva");
  btn.disabled = true;
  try {
    const nombres = await listarEjecuciones();
    G.listaHoy = nombres.find(n => n.startsWith(hoyPrefijo())) || null;
  } catch (e) { console.warn("No pude comprobar la lista de hoy:", e); G.listaHoy = null; }
  btn.textContent = G.listaHoy
    ? `▶ Abrir lista en marcha (creada a las ${G.listaHoy.slice(6, 8)}:${G.listaHoy.slice(8, 10)})`
    : "▶ Realizar una lista completa";
  btn.disabled = false;
}
function volverAlMenu() {
  G.ejecucion = null; cerrarCarrete(); detenerSync();
  mostrarVista("menu");
  setEstado("Elige una opción.");
  actualizarMenu();
}
function onNueva() {
  if (G.listaHoy) abrirEjecucion(G.listaHoy);
  else nuevaEjecucion();
}
function resetUiArbol() {
  UI.verHechas = true;
  UI.editables.clear();
  cerrarCarrete();
}
async function nuevaEjecucion() {
  try {
    setEstado("Creando nueva lista…");
    // Doble comprobación del límite diario (por si el menú estaba desactualizado)
    const previas = await listarEjecuciones();
    const deHoy = previas.find(n => n.startsWith(hoyPrefijo()));
    if (deHoy) { G.listaHoy = deHoy; return abrirEjecucion(deHoy); }
    await cargarPlantilla();                       // siempre la plantilla actual
    const nombre = nombreEjecucion();
    const sub = await crearSubcarpeta(nombre);
    // Estado: PISO;ZONA;ACCION;ESTADO;FOTO;HORA;USUARIO;MARCA — congela la estructura con la que se ejecuta
    const rows = G.plantilla.map(t => [t[0], t[1], t[2], "", "", "", "", ""]);
    G.ejecucion = { nombre, subfolderId: sub.id, stateName: `${nombre}.csv`, rows, dirty: new Set(), mod: 0 };
    G.listaHoy = nombre;
    await guardarEstado();
    setEstado(`Lista ${nombreLegible(nombre)} creada.`);
    resetUiArbol();
    renderArbol();
    mostrarVista("arbol");
    iniciarSync();
  } catch (e) { setEstado("Error: " + e.message, true); }
}

async function verEjecuciones() {
  try {
    setEstado("Buscando listas existentes…");
    const nombres = await listarEjecuciones();
    const ul = document.getElementById("listaEjecuciones");
    ul.innerHTML = "";
    if (!nombres.length) {
      ul.innerHTML = `<li class="vacio">No hay ninguna lista todavía.</li>`;
    } else {
      nombres.forEach(n => {
        const li = document.createElement("li");
        li.innerHTML = `<button type="button" data-nombre="${n}">
                          <span class="nombre">${fechaLarga(n)}</span><span class="fecha">(${haceDias(n)})</span>
                        </button>`;
        ul.appendChild(li);
      });
    }
    setEstado(`${nombres.length} lista(s) encontrada(s).`);
    mostrarVista("ejecuciones");
  } catch (e) { setEstado("Error: " + e.message, true); }
}

async function abrirEjecucion(nombre) {
  try {
    setEstado(`Abriendo ${nombreLegible(nombre)}…`);
    const text = await descargarTexto(`${nombre}/${nombre}.csv`);
    const rows = parseCSV(text).map(r => {
      while (r.length < 8) r.push("");
      return r.slice(0, 8).map(f => (f || "").trim());
    });
    G.ejecucion = { nombre, subfolderId: null, stateName: `${nombre}.csv`, rows, dirty: new Set(), mod: 0 };
    setEstado(`Lista ${nombreLegible(nombre)} cargada (${rows.length} tareas).`);
    resetUiArbol();
    renderArbol();
    mostrarVista("arbol");
    iniciarSync();
  } catch (e) { setEstado("Error abriendo la lista: " + e.message, true); }
}

// ---------- Render del árbol piso → zona → acción ----------
function renderArbol() {
  const e = G.ejecucion;
  document.getElementById("tituloEjecucion").textContent = `Lista ${nombreLegible(e.nombre)}  (${e.nombre})`;
  document.getElementById("btnVerHechas").textContent = `Ver realizadas: ${UI.verHechas ? "SÍ" : "NO"}`;
  const cont = document.getElementById("tareas");
  cont.innerHTML = "";
  let pisoAct = null, zonaAct = null;
  let pendPiso = false, pendZona = false;   // cabeceras pendientes de pintar (solo si hay tarea visible)
  e.rows.forEach((row, idx) => {
    const [piso, zona, accion, estado, foto] = row;
    if (piso !== pisoAct) { pisoAct = piso; zonaAct = null; pendPiso = true; }
    if (zona !== zonaAct) { zonaAct = zona; pendZona = true; }
    const hecha = norm(estado) === "ok";
    if (hecha && !UI.verHechas) return;     // ocultar realizadas
    if (pendPiso) {
      const d = document.createElement("div"); d.className = "piso"; d.textContent = piso || "(sin piso)";
      cont.appendChild(d); pendPiso = false;
    }
    if (pendZona) {
      const d = document.createElement("div"); d.className = "zona"; d.textContent = zona || "(sin zona)";
      cont.appendChild(d); pendZona = false;
    }
    const conFoto = esFoto(accion);
    const editable = UI.editables.has(idx);
    const d = document.createElement("div");
    d.className = "tarea" + (hecha ? " hecha" : "");
    let botones;
    if (!hecha) {
      // Pendiente: acción directa
      botones = conFoto
        ? `<button type="button" class="foto" data-idx="${idx}" data-act="foto">📷 Foto</button>`
        : `<button type="button" class="check pendiente" data-idx="${idx}" data-act="check">Pendiente</button>`;
    } else if (!editable) {
      // Hecha y bloqueada: solo "Editar"
      botones = `<button type="button" class="edit" data-idx="${idx}" data-act="editar">✎ Editar</button>`;
    } else {
      // Hecha y desbloqueada: acción de edición + cancelar
      botones = (conFoto
        ? `<button type="button" class="foto" data-idx="${idx}" data-act="foto">📷 Repetir</button>`
        : `<button type="button" class="check deshacer" data-idx="${idx}" data-act="check">↩ Quitar</button>`)
        + `<button type="button" class="cancel" data-idx="${idx}" data-act="cancelar">✕</button>`;
    }
    const ver = hecha && foto ? ` <a class="ver" href="#" data-idx="${idx}" data-act="ver">ver foto</a>` : "";
    let meta = "";
    if (hecha) {
      const partes = [];
      if (row[6]) partes.push(escapeHtml(row[6]));
      const t = haceTexto(row, e.nombre);
      if (t) partes.push(t);
      if (row[5]) partes.push(escapeHtml(row[5]));
      if (partes.length) meta = `<span class="hora">${partes.join(" · ")}</span>`;
    }
    d.innerHTML = `<span class="texto">${escapeHtml(accion)}${ver}${meta}</span><span class="botones">${botones}</span>`;
    cont.appendChild(d);
  });
  actualizarProgreso();
}
function actualizarProgreso() {
  const rows = G.ejecucion.rows;
  const ok = rows.filter(r => norm(r[3]) === "ok").length;
  document.getElementById("progreso").textContent =
    `${ok} de ${rows.length} tareas completadas · pendientes: ${rows.length - ok}` +
    (ok === rows.length ? " · ✅ LISTA COMPLETA" : "");
}

// ---------- Acciones sobre tareas ----------
function marcar(idx, estado, foto) {
  const e = G.ejecucion;
  const row = e.rows[idx];
  const ok = estado === "OK";
  row[3] = estado;
  if (foto !== undefined) row[4] = foto;
  row[5] = ok ? new Date().toTimeString().slice(0, 8) : "";
  row[6] = ok ? nombreUsuario() : "";   // quién lo ha hecho (primera palabra del nombre)
  row[7] = ok ? ahoraTS() : "";         // marca completa para calcular "hace x min"
  e.dirty.add(idx);           // modificado localmente: gana en la fusión hasta subirse
  e.mod++;
  UI.editables.delete(idx);   // tras editar, la tarea vuelve a quedar bloqueada
  renderArbol();
  programarGuardado();
}
function onCheck(idx) {
  const row = G.ejecucion.rows[idx];
  const hecha = norm(row[3]) === "ok";
  if (hecha && !UI.editables.has(idx)) return;   // hecha y bloqueada: no editable
  if (!hecha && !confirm(`¿Confirmas que esta tarea está hecha?\n\n${row[2]}`)) return;
  marcar(idx, hecha ? "" : "OK");                // toggle solo en modo edición
}
function pedirFoto(idx) {
  const hecha = norm(G.ejecucion.rows[idx][3]) === "ok";
  if (hecha && !UI.editables.has(idx)) return;   // hecha y bloqueada: no editable
  fotoPendiente = idx;
  const inp = document.getElementById("fotoInput");
  inp.value = "";
  inp.click();
}
// Reduce la foto a máx. 1600 px (JPEG 80%) para subir rápido; si falla, sube el original
async function comprimirFoto(file) {
  try {
    const bmp = await createImageBitmap(file);
    const MAX = 1600, esc = Math.min(1, MAX / Math.max(bmp.width, bmp.height));
    const cv = document.createElement("canvas");
    cv.width = Math.round(bmp.width * esc); cv.height = Math.round(bmp.height * esc);
    cv.getContext("2d").drawImage(bmp, 0, 0, cv.width, cv.height);
    const blob = await new Promise(res => cv.toBlob(res, "image/jpeg", 0.8));
    return blob || file;
  } catch (e) { console.warn("Compresión falló, subo original:", e); return file; }
}
async function onFotoElegida(file) {
  const idx = fotoPendiente; fotoPendiente = null;
  if (idx == null || !file) return;
  const e = G.ejecucion;
  const fname = `${String(idx + 1).padStart(3, "0")}_${slug(e.rows[idx][2])}.jpg`;
  try {
    setEstado(`Subiendo foto ${fname}…`);
    const blob = await comprimirFoto(file);
    await subirFoto(`${e.nombre}/${fname}`, blob);
    marcar(idx, "OK", fname);
    setEstado(`Foto ${fname} subida.`);
  } catch (err) { setEstado("Error subiendo la foto: " + err.message, true); }
}
async function verFoto(idx) {
  const e = G.ejecucion;
  const fname = e.rows[idx][4];
  if (!fname) return;
  try {
    const res = await gFetch(`${itemBase()}/items/${G.folderId}:/${encodeURI(e.nombre + "/" + fname)}`);
    if (!res.ok) throw new Error(res.status);
    const item = await res.json();
    const url = (item["@microsoft.graph.downloadUrl"]) || item.webUrl;
    if (url) window.open(url, "_blank");
  } catch (err) { setEstado("No pude abrir la foto: " + err.message, true); }
}

// ---------- Carrete de fotos de la ejecución ----------
async function abrirCarrete() {
  const e = G.ejecucion;
  if (!e) return;
  try {
    setEstado("Cargando carrete…");
    const fotos = await listarFotos(e.nombre);
    const grid = document.getElementById("carreteGrid");
    grid.innerHTML = "";
    if (!fotos.length) {
      grid.innerHTML = `<p class="vacio">Todavía no hay fotos en esta lista.</p>`;
    } else {
      fotos.forEach(it => {
        const th = it.thumbnails && it.thumbnails[0];
        const mini = th && (th.large || th.medium || th.small);
        const full = it["@microsoft.graph.downloadUrl"] || it.webUrl || "#";
        const a = document.createElement("a");
        a.className = "fotoRoll";
        a.href = full; a.target = "_blank"; a.rel = "noopener";
        a.innerHTML = (mini ? `<img src="${mini.url}" alt="${escapeHtml(it.name)}" loading="lazy" />`
                            : `<span class="sinMini">🖼</span>`)
                    + `<span class="pie">${escapeHtml(it.name)}</span>`;
        grid.appendChild(a);
      });
    }
    document.getElementById("carrete").hidden = false;
    setEstado(`Carrete: ${fotos.length} foto(s).`);
  } catch (err) { setEstado("Error cargando el carrete: " + err.message, true); }
}
function cerrarCarrete() {
  const c = document.getElementById("carrete");
  if (c) { c.hidden = true; document.getElementById("carreteGrid").innerHTML = ""; }
}

// ---------- Inicio tras login ----------
async function iniciar() {
  if (inicializado) return;
  inicializado = true;
  setEstado(`Buscando carpeta "${CARPETA_RAIZ}"…`);
  try {
    const token = await getToken(fileScopes);
    if (!(await localizarCarpeta(token))) { inicializado = false; return; }
    await actualizarMenu();
    setEstado("Carpeta localizada. Elige una opción.");
    mostrarVista("menu");
  } catch (e) {
    setEstado("Error: " + e.message, true);
    inicializado = false;
  }
}

// ---------- Autenticación (idéntica a la app hermana) ----------
function updateAuthUI() {
  const account = msalInstance.getActiveAccount();
  const loginBtn = document.getElementById("loginBtn");
  if (!loginBtn) return;
  if (account) {
    loginBtn.textContent = "Cerrar sesión (" + (account.name || account.username) + ")";
    document.body.classList.add("logged-in");
    loginBtn.onclick = async (e) => {
      e.preventDefault(); loginBtn.disabled = true;
      try { await msalInstance.logoutPopup({ account }); msalInstance.setActiveAccount(null); }
      catch (err) { console.error(err); }
      finally {
        loginBtn.disabled = false; inicializado = false; G.ejecucion = null;
        detenerSync();
        document.getElementById("tareas").innerHTML = "";
        document.getElementById("listaEjecuciones").innerHTML = "";
        cerrarCarrete();
        mostrarVista("");
        setEstado("Inicia sesión para empezar.");
        updateAuthUI();
      }
    };
    iniciar();
  } else {
    loginBtn.textContent = "Iniciar sesión";
    document.body.classList.remove("logged-in");
    loginBtn.onclick = async (e) => { e.preventDefault(); loginBtn.disabled = true; try { await handleLoginClick(); } finally { loginBtn.disabled = false; updateAuthUI(); } };
  }
}
async function handleLoginClick() { return msalInstance.loginRedirect(loginRequest); }
async function getToken(request) {
  const active = msalInstance.getActiveAccount();
  const req = Object.assign({}, request, { account: request.account || active });
  try { return (await msalInstance.acquireTokenSilent(req)).accessToken; }
  catch (e) { console.warn("Silent falló, redirigiendo:", e && e.errorCode); msalInstance.acquireTokenRedirect(request); return; }
}
function resetMsal() {
  Object.keys(localStorage).filter(k => k.startsWith("msal.") || k.includes(msalConfig.auth.clientId)).forEach(k => localStorage.removeItem(k));
  if (location.hash) history.replaceState(null, "", location.pathname + location.search);
  console.log("Estado MSAL limpiado.");
}

// ---------- Arranque / eventos ----------
document.getElementById("btnNueva").addEventListener("click", onNueva);
document.getElementById("btnVer").addEventListener("click", verEjecuciones);
document.getElementById("btnVolver1").addEventListener("click", volverAlMenu);
document.getElementById("btnVolver2").addEventListener("click", volverAlMenu);
document.getElementById("btnVerHechas").addEventListener("click", () => { UI.verHechas = !UI.verHechas; renderArbol(); });
document.getElementById("btnCarrete").addEventListener("click", abrirCarrete);
document.getElementById("btnCerrarCarrete").addEventListener("click", cerrarCarrete);
document.getElementById("listaEjecuciones").addEventListener("click", (e) => {
  const b = e.target.closest("button[data-nombre]");
  if (b) abrirEjecucion(b.dataset.nombre);
});
document.getElementById("tareas").addEventListener("click", (e) => {
  const el = e.target.closest("[data-act]");
  if (!el) return;
  e.preventDefault();
  const idx = Number(el.dataset.idx);
  if (el.dataset.act === "check") onCheck(idx);
  else if (el.dataset.act === "foto") pedirFoto(idx);
  else if (el.dataset.act === "ver") verFoto(idx);
  else if (el.dataset.act === "editar") { UI.editables.add(idx); renderArbol(); }
  else if (el.dataset.act === "cancelar") { UI.editables.delete(idx); renderArbol(); }
});
document.getElementById("fotoInput").addEventListener("change", (e) => onFotoElegida(e.target.files[0]));
// Al volver la app/pestaña a primer plano, resincronizar de inmediato (multi-dispositivo)
document.addEventListener("visibilitychange", () => { if (!document.hidden) sincronizar(); });

(async () => {
  try {
    const rr = await msalInstance.handleRedirectPromise();
    if (rr && rr.account) msalInstance.setActiveAccount(rr.account);
    else { const a = msalInstance.getAllAccounts(); if (a.length === 1) msalInstance.setActiveAccount(a[0]); }
  } catch (e) { console.error(e); }
  updateAuthUI();
})();
window.pscmAuth = { msalInstance, getToken, resetMsal, G, UI };
