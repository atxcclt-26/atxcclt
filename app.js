// app.js — Revisión de cargos (columnas del banco + SINO_2 / CATEGORIA_2 / WARRANTY)

// ====== AJUSTES ======
const PREFIJO_CARPETA = "zzz_sg_";                 // carpeta en la raíz (sin distinguir mayúsculas)
const FILE_NAME       = "AlCaralloConLosTickets.csv";

// El CSV NO tiene fila de cabecera: TODAS las filas son datos.
// Estos son los nombres de columna (etiquetas de combos y tabla), en el orden del fichero.
const BASE_HEADER = [
  "BANCO", "FECHA CARGO", "FECHA GASTO", "COMERCIO", "DETALLE",
  "SINO", "€", "CATEGORIA", "COMENTARIO", "MONEDA", "VALOR MONEDA", "VALOR ORIGINAL",
  "FECHA", "FICHERO"
];

// Columnas
const HIDDEN_COLS = ["VALOR MONEDA", "FECHA", "FICHERO"];   // no se muestran ni se filtran
const COL_SINO = "SINO", COL_CAT = "CATEGORIA";    // columnas de origen que disparan botones/color
const COL_SINO2 = "SINO_2", COL_CAT2 = "CATEGORIA_2", COL_WAR = "WARRANTY";  // columnas nuevas
const EXTRA_COLS = [COL_SINO2, COL_CAT2, COL_WAR];
// =====================

// 1. MSAL
const esMovil = /Android|iPhone|iPad|iPod/i.test(navigator.userAgent) ||
                (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
const REDIRECT_PAGE = window.location.origin + window.location.pathname;
const msalConfig = {
  auth: { clientId: "24e9d6d3-d9ad-437e-b7f6-1a27f48c2696", redirectUri: REDIRECT_PAGE },
  cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
  system: { loggerOptions: { logLevel: msal.LogLevel.Warning, loggerCallback: (l, m, p) => { if (!p) console.log(`[MSAL][${l}]`, m); } } }
};
const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["User.Read", "Files.ReadWrite"] };
const fileScopes   = { scopes: ["Files.ReadWrite"] };

// Estado
const G = { fileId: null, driveId: null, header: [], rows: [], delim: ",", idx: {}, hidden: [] };
const filtros = {};
let revisionIniciada = false;
let guardando = false, guardarDeNuevo = false;

// ---------- UI helpers ----------
function setEstado(msg, err = false) {
  const el = document.getElementById("estado-carga");
  if (el) { el.textContent = msg; el.classList.toggle("error", !!err); }
}
function escapeHtml(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
function norm(v) { return (v || "").trim().toLowerCase(); }
function up(v) { return (v || "").trim().toUpperCase(); }
function col(nombre) { return G.header.findIndex(h => up(h) === up(nombre)); }

// ---------- CSV ----------
function detectarDelim(t) { const f = t.split(/\r?\n/)[0] || ""; return f.split(";").length > f.split(",").length ? ";" : ","; }
function parseCSV(text, delim) {
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
  return rows.filter(r => !(r.length === 1 && r[0] === ""));
}
function serialF(v, d) { v = v == null ? "" : String(v); return /["\n\r]/.test(v) || v.includes(d) ? '"' + v.replace(/"/g, '""') + '"' : v; }
function serializeCSV(rows, d) { return rows.map(r => r.map(f => serialF(f, d)).join(d)).join("\r\n"); }

// ---------- Graph / OneDrive ----------
const GRAPH = "https://graph.microsoft.com/v1.0";
function itemBase() { return G.driveId ? `${GRAPH}/drives/${G.driveId}` : `${GRAPH}/me/drive`; }
async function localizarFichero(token) {
  const H = { Authorization: `Bearer ${token}` };
  const rRoot = await fetch(`${GRAPH}/me/drive/root/children?$top=200`, { headers: H });
  if (!rRoot.ok) {
    if (rRoot.status === 401 || rRoot.status === 403)
      throw new Error(`Sin permiso (HTTP ${rRoot.status}). ¿Añadiste "Files.ReadWrite" en Azure? resetMsal() y reentra.`);
    throw new Error("raíz " + rRoot.status + " " + (await rRoot.text()));
  }
  const hijos = (await rRoot.json()).value || [];
  const prefijo = PREFIJO_CARPETA.toLowerCase();
  const carpetas = hijos.filter(it => (it.folder || it.remoteItem) && (it.name || "").toLowerCase().startsWith(prefijo));
  if (!carpetas.length) {
    console.warn("Contenido de la raíz:", hijos.map(it => ({ nombre: it.name, tipo: it.folder ? "carpeta" : it.remoteItem ? "compartida" : "fichero" })));
    setEstado(`No hay ninguna carpeta que empiece por "${PREFIJO_CARPETA}" en la raíz de tu OneDrive.`, true);
    return null;
  }
  for (const c of carpetas) {
    let driveId = null, folderId = c.id;
    if (c.remoteItem) { driveId = c.remoteItem.parentReference && c.remoteItem.parentReference.driveId; folderId = c.remoteItem.id; }
    const base = driveId ? `${GRAPH}/drives/${driveId}` : `${GRAPH}/me/drive`;
    const rf = await fetch(`${base}/items/${folderId}:/${encodeURIComponent(FILE_NAME)}`, { headers: H });
    if (rf.ok) { G.driveId = driveId; console.log(`Fichero encontrado en carpeta: ${c.name}${driveId ? " (compartida)" : ""}`); return await rf.json(); }
    if ((rf.status === 401 || rf.status === 403) && driveId) {
      setEstado(`La carpeta "${c.name}" es compartida y necesita "Files.ReadWrite.All" en Azure.`, true); return null;
    }
  }
  setEstado(`Encontré carpeta(s) "${PREFIJO_CARPETA}*" pero ninguna contiene "${FILE_NAME}".`, true);
  return null;
}
async function descargarCsv(token) {
  const res = await fetch(`${itemBase()}/items/${G.fileId}/content`, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error("descarga " + res.status);
  return new TextDecoder("utf-8").decode(await res.arrayBuffer()).replace(/^\uFEFF/, "");
}
async function guardarCsv() {
  const token = await getToken(fileScopes);
  const res = await fetch(`${itemBase()}/items/${G.fileId}/content`, {
    method: "PUT", headers: { Authorization: `Bearer ${token}`, "Content-Type": "text/csv" },
    // El fichero es headerless: se guardan solo las filas de datos (con las 3 columnas añadidas al final).
    body: serializeCSV(G.rows, G.delim)
  });
  if (!res.ok) throw new Error(res.status + " " + (await res.text()));
}
function programarGuardado() {
  if (guardando) { guardarDeNuevo = true; return; }
  guardando = true;
  guardarCsv()
    .then(() => setEstado(`Cambios guardados en ${FILE_NAME}.`))
    .catch(e => setEstado("Error guardando (los cambios siguen en pantalla, reintenta): " + e.message, true))
    .finally(() => { guardando = false; if (guardarDeNuevo) { guardarDeNuevo = false; programarGuardado(); } });
}

// ---------- Índices de columnas ----------
function prepararColumnas() {
  // La cabecera (BASE_HEADER + EXTRA_COLS) ya se fija al cargar; esto es un no-op salvo que faltara alguna.
  EXTRA_COLS.forEach(name => {
    if (col(name) === -1) { G.header.push(name); G.rows.forEach(r => r.push("")); }
  });
  // Normaliza cada fila a la longitud de la cabecera (rellena las columnas añadidas la 1ª vez).
  G.rows.forEach(r => { while (r.length < G.header.length) r.push(""); });
  G.idx = {
    SINO: col(COL_SINO), CAT: col(COL_CAT),
    SINO2: col(COL_SINO2), CAT2: col(COL_CAT2), WAR: col(COL_WAR)
  };
  G.hidden = HIDDEN_COLS.map(col).filter(i => i !== -1);
}
function esVisible(c) { return !G.hidden.includes(c); }
function esRevisable(row) { return G.idx.SINO !== -1 && norm(row[G.idx.SINO]) === "no"; }
function warSi(row) { return norm(row[G.idx.WAR]) === "si"; }

// ---------- Filtros ----------
function valoresDistintos(c) {
  const s = new Set();
  G.rows.forEach(r => { const v = r[c] || ""; if (v !== "") s.add(v); });
  return [...s].sort((a, b) => a.localeCompare(b, "es", { numeric: true }));
}
function construirFiltros() {
  const cont = document.getElementById("filtros");
  cont.innerHTML = "";
  G.header.forEach((h, c) => {
    if (!esVisible(c)) return;                      // FECHA / FICHERO no generan combo
    const wrap = document.createElement("div"); wrap.className = "f";
    const lab = document.createElement("label"); lab.textContent = h || `Col ${c + 1}`;
    const sel = document.createElement("select");
    if (c === G.idx.SINO2) {
      sel.innerHTML = `<option value="">Todos</option><option value="__pend">Pendientes</option><option value="ok">OK</option><option value="nok">NOK</option>`;
    } else if (c === G.idx.WAR) {
      sel.innerHTML = `<option value="">Todos</option><option value="si">Sí</option><option value="no">No</option>`;
    } else {
      sel.innerHTML = `<option value="">Todos</option>` + valoresDistintos(c).map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join("");
    }
    sel.value = filtros[c] || "";
    sel.addEventListener("change", () => { filtros[c] = sel.value; renderTabla(); });
    wrap.append(lab, sel); cont.appendChild(wrap);
  });
  const btn = document.createElement("button");
  btn.id = "btnLimpiar"; btn.textContent = "Limpiar";
  btn.addEventListener("click", () => { Object.keys(filtros).forEach(k => delete filtros[k]); construirFiltros(); renderTabla(); });
  cont.appendChild(btn);
}
function pasaFiltros(row) {
  for (const c in filtros) {
    const sel = filtros[c]; if (!sel) continue;
    const ci = Number(c);
    if (ci === G.idx.SINO2) {
      const v = norm(row[G.idx.SINO2]);
      if (sel === "__pend" && !(esRevisable(row) && v !== "ok" && v !== "nok")) return false;
      if (sel === "ok" && v !== "ok") return false;
      if (sel === "nok" && v !== "nok") return false;
    } else if (ci === G.idx.WAR) {
      if ((warSi(row) ? "si" : "no") !== sel) return false;
    } else if ((row[ci] || "") !== sel) return false;
  }
  return true;
}

// ---------- Tabla ----------
function renderCabecera() {
  const ths = G.header.map((h, c) => esVisible(c) ? `<th>${escapeHtml(h)}</th>` : "").join("");
  document.querySelector("#tabla thead").innerHTML = `<tr>${ths}<th>Acción</th></tr>`;
}
function celda(row, c) {
  if (c === G.idx.SINO2) {
    const v = norm(row[c]); const cls = v === "ok" ? "ok" : v === "nok" ? "nok" : "pend";
    return `<td class="sino2 ${cls}">${escapeHtml(row[c] || "—")}</td>`;
  }
  if (c === G.idx.WAR) {
    const si = warSi(row);
    return `<td class="war ${si ? "si" : "no"}">${si ? "Sí" : "No"}</td>`;
  }
  if (c === G.idx.SINO) {
    // Rojo si SINO = "No" y SINO_2 sigue sin clasificar (vacío o "no"); verde en el resto.
    const flagged = up(row[c]) === "NO";
    const s2 = norm(row[G.idx.SINO2]);
    const resuelto = s2 !== "" && s2 !== "no";
    const cls = (flagged && !resuelto) ? "cell-red" : "cell-green";
    return `<td class="${cls}">${escapeHtml(row[c] || "")}</td>`;
  }
  if (c === G.idx.CAT) {
    // Rojo si CATEGORIA = "OTROS" y CATEGORIA_2 está vacío; verde en el resto.
    const flagged = up(row[c]) === "OTROS";
    const cat2 = (row[G.idx.CAT2] || "").trim();
    const cls = (flagged && cat2 === "") ? "cell-red" : "cell-green";
    return `<td class="${cls}">${escapeHtml(row[c] || "")}</td>`;
  }
  return `<td>${escapeHtml(row[c] || "")}</td>`;
}
function botonesAccion(row, idx) {
  let h = "";
  if (esRevisable(row)) {                          // OK/NOK solo si SINO = No
    const v = norm(row[G.idx.SINO2]);
    h += `<button class="ok" data-act="sino2" data-val="ok" data-idx="${idx}"${v === "ok" ? " disabled" : ""}>OK</button>`;
    h += `<button class="nok" data-act="sino2" data-val="nok" data-idx="${idx}"${v === "nok" ? " disabled" : ""}>NOK</button>`;
  }
  if (G.idx.CAT !== -1 && norm(row[G.idx.CAT]) === "otros") {   // Categoría si CATEGORIA = OTROS (siempre visible)
    const cat = row[G.idx.CAT2] || "";
    h += `<button class="cat" data-act="cat" data-idx="${idx}">Categoría${cat ? ": " + escapeHtml(cat) : ""}</button>`;
  }
  const si = warSi(row);                           // Garantía siempre
  h += `<button class="war ${si ? "si" : "no"}" data-act="war" data-idx="${idx}">Garantía: ${si ? "Sí" : "No"}</button>`;
  return `<td class="accion">${h}</td>`;
}
function renderTabla() {
  const tbody = document.querySelector("#tabla tbody");
  const frag = document.createDocumentFragment();
  let visibles = 0;
  G.rows.forEach((row, idx) => {
    if (!pasaFiltros(row)) return;
    visibles++;
    const tr = document.createElement("tr");
    let html = "";
    G.header.forEach((_, c) => { if (esVisible(c)) html += celda(row, c); });
    html += botonesAccion(row, idx);
    tr.innerHTML = html;
    frag.appendChild(tr);
  });
  tbody.innerHTML = ""; tbody.appendChild(frag);
  actualizarProgreso(visibles);
}
function actualizarProgreso(visibles) {
  const rev = G.rows.filter(esRevisable);
  const ok = rev.filter(r => norm(r[G.idx.SINO2]) === "ok").length;
  const nok = rev.filter(r => norm(r[G.idx.SINO2]) === "nok").length;
  const war = G.rows.filter(warSi).length;
  const prog = document.getElementById("progreso");
  if (prog) prog.textContent =
    `${G.rows.length} cargos · revisables (SINO=No): ${rev.length} · OK: ${ok} · NOK: ${nok} · pendientes: ${rev.length - ok - nok} · garantía Sí: ${war}` +
    (visibles != null ? ` · mostrando: ${visibles}` : "");
}

// ---------- Acciones ----------
// Resumen del gasto (columnas visibles de origen) para mostrarlo en la ventanita de categoría.
function resumenFila(row) {
  const partes = [];
  G.header.forEach((h, c) => {
    if (!esVisible(c)) return;
    if (c === G.idx.SINO2 || c === G.idx.CAT2 || c === G.idx.WAR) return;
    const v = (row[c] || "").trim();
    if (v !== "") partes.push(`${h}: ${v}`);
  });
  return partes.join("\n");
}
function setCampo(rowIdx, colIdx, valor) {
  if (colIdx == null || colIdx < 0) return false;
  if (G.rows[rowIdx][colIdx] === valor) return false;
  G.rows[rowIdx][colIdx] = valor;
  renderTabla();
  programarGuardado();
  return true;
}
function onAccion(idx, act) {
  const row = G.rows[idx];
  if (act === "sino2-ok")  setCampo(idx, G.idx.SINO2, "ok");
  else if (act === "sino2-nok") setCampo(idx, G.idx.SINO2, "nok");
  else if (act === "war") setCampo(idx, G.idx.WAR, warSi(row) ? "No" : "Si");
  else if (act === "cat") {
    const actual = row[G.idx.CAT2] || "";
    const val = window.prompt(resumenFila(row) + "\n\nCategoría para este cargo:", actual);
    if (val === null) return;                       // cancelado
    setCampo(idx, G.idx.CAT2, val.trim());
  }
}

// ---------- Carga ----------
async function iniciarRevision() {
  if (revisionIniciada) return;
  revisionIniciada = true;
  setEstado(`Buscando "${FILE_NAME}" en carpeta "${PREFIJO_CARPETA}*"…`);
  try {
    const token = await getToken(fileScopes);
    const item = await localizarFichero(token);
    if (!item) { revisionIniciada = false; return; }
    G.fileId = item.id;
    const text = await descargarCsv(token);
    G.delim = detectarDelim(text);
    const all = parseCSV(text, G.delim);
    if (!all.length) { setEstado("El fichero está vacío.", true); return; }
    // El CSV es headerless: TODAS las filas son datos. La cabecera es sintética (etiquetas fijas).
    G.header = BASE_HEADER.concat(EXTRA_COLS);
    G.rows = all;
    prepararColumnas();
    construirFiltros();
    renderCabecera();
    renderTabla();
    setEstado(`Fichero cargado: ${item.name} (${G.rows.length} filas, separador "${G.delim}").`);
  } catch (e) {
    setEstado("Error: " + e.message, true);
    revisionIniciada = false;
  }
}

// ---------- Autenticación ----------
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
        loginBtn.disabled = false; revisionIniciada = false;
        document.querySelector("#tabla tbody").innerHTML = "";
        document.querySelector("#tabla thead").innerHTML = "";
        document.getElementById("filtros").innerHTML = "";
        setEstado("Inicia sesión para cargar el fichero.");
        updateAuthUI();
      }
    };
    iniciarRevision();
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
async function listarOneDrive(carpeta = "") {
  const token = await getToken(fileScopes);
  const url = carpeta ? `${GRAPH}/me/drive/root:/${encodeURI(carpeta)}:/children` : `${GRAPH}/me/drive/root/children?$top=200`;
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) { console.error("listar:", r.status, await r.text()); return; }
  console.table(((await r.json()).value || []).map(it => ({ nombre: it.name, tipo: it.folder ? "carpeta" : it.remoteItem ? "compartida" : "fichero" })));
}

// ---------- Arranque ----------
document.querySelector("#tabla tbody").addEventListener("click", (e) => {
  const b = e.target.closest("button[data-act]");
  if (!b) return;
  const idx = Number(b.dataset.idx);
  const act = b.dataset.act === "sino2" ? "sino2-" + b.dataset.val : b.dataset.act;
  onAccion(idx, act);
});
(async () => {
  try {
    const rr = await msalInstance.handleRedirectPromise();
    if (rr && rr.account) msalInstance.setActiveAccount(rr.account);
    else { const a = msalInstance.getAllAccounts(); if (a.length === 1) msalInstance.setActiveAccount(a[0]); }
  } catch (e) { console.error(e); }
  updateAuthUI();
})();
window.tuppersAuth = { msalInstance, getToken, resetMsal, iniciarRevision, listar: listarOneDrive, G };
