// app.js — Revisión de cargos (tabla con OK/NOK por fila + filtros por columna)
// El fichero está en TU OneDrive, dentro de una carpeta cuyo nombre empieza por un prefijo
// (p.ej. zzz_sg_1, zzz_sg_3, zzz_sg_4, zzz_sg_5). La app localiza sola la que contenga el CSV.

// ====== AJUSTES ======
const PREFIJO_CARPETA = "zzz_sg_";                 // carpeta en la raíz de tu OneDrive que empieza por esto
const FILE_NAME       = "AlCaralloConLosTickets.csv";
const STATUS_COLUMN   = "estado";
// =====================

// Detección de dispositivo y URIs de redirección (funciona en local y en GitHub Pages)
const esMovil = /Android|iPhone|iPad|iPod/i.test(navigator.userAgent) ||
                (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
const REDIRECT_POPUP = new URL("blank.html", window.location.href).href;   // para el popup (escritorio)
const REDIRECT_PAGE  = window.location.origin + window.location.pathname;  // la propia página (móvil, flujo redirect)

// 1. MSAL
const msalConfig = {
  auth: { clientId: "24e9d6d3-d9ad-437e-b7f6-1a27f48c2696", redirectUri: REDIRECT_PAGE },
  cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false },
  system: { loggerOptions: { logLevel: msal.LogLevel.Warning, loggerCallback: (l, m, p) => { if (!p) console.log(`[MSAL][${l}]`, m); } } }
};
const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["User.Read", "Files.ReadWrite"] };
const fileScopes   = { scopes: ["Files.ReadWrite"] };

// Estado
const G = { fileId: null, header: [], rows: [], delim: ",", statusIdx: -1 };
const filtros = {};
let revisionIniciada = false;
let guardando = false, guardarDeNuevo = false;

// ---------- UI helpers ----------
function setEstado(msg, err = false) {
  const el = document.getElementById("estado-carga");
  if (el) { el.textContent = msg; el.classList.toggle("error", !!err); }
}
function escapeHtml(s) {
  return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function estadoNorm(v) { return (v || "").trim().toLowerCase(); }

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
// Escanea la raíz, encuentra las carpetas que empiezan por PREFIJO_CARPETA y devuelve
// el item del fichero en la primera que lo contenga.
async function localizarFichero(token) {
  const H = { Authorization: `Bearer ${token}` };
  const rRoot = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root/children?$top=200`, { headers: H });
  if (!rRoot.ok) {
    if (rRoot.status === 401 || rRoot.status === 403)
      throw new Error(`Sin permiso para leer ficheros (HTTP ${rRoot.status}). ¿Añadiste "Files.ReadWrite" en Azure y lo consentiste? Ejecuta tuppersAuth.resetMsal() y vuelve a entrar.`);
    throw new Error("raíz " + rRoot.status + " " + (await rRoot.text()));
  }
  const hijos = (await rRoot.json()).value || [];
  const carpetas = hijos.filter(it => it.folder && it.name.startsWith(PREFIJO_CARPETA));
  if (!carpetas.length) {
    console.warn("Carpetas en la raíz:", hijos.filter(it => it.folder).map(it => it.name));
    setEstado(`No hay ninguna carpeta que empiece por "${PREFIJO_CARPETA}" en la raíz de tu OneDrive.`, true);
    return null;
  }
  for (const c of carpetas) {
    const rf = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${c.id}:/${encodeURIComponent(FILE_NAME)}`, { headers: H });
    if (rf.ok) { console.log(`Fichero encontrado en carpeta: ${c.name}`); return await rf.json(); }
  }
  console.warn("Carpetas revisadas:", carpetas.map(c => c.name));
  setEstado(`Encontré ${carpetas.length} carpeta(s) "${PREFIJO_CARPETA}*" pero ninguna contiene "${FILE_NAME}".`, true);
  return null;
}
async function descargarCsv(token) {
  const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${G.fileId}/content`, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error("descarga " + res.status);
  return new TextDecoder("utf-8").decode(await res.arrayBuffer()).replace(/^\uFEFF/, "");
}
async function guardarCsv() {
  const token = await getToken(fileScopes);
  const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${G.fileId}/content`, {
    method: "PUT", headers: { Authorization: `Bearer ${token}`, "Content-Type": "text/csv" },
    body: serializeCSV([G.header, ...G.rows], G.delim)
  });
  if (!res.ok) throw new Error(res.status + " " + (await res.text()));
}
function programarGuardado() {
  if (guardando) { guardarDeNuevo = true; return; }
  guardando = true;
  guardarCsv()
    .then(() => setEstado(`Cambios guardados en ${FILE_NAME}.`))
    .catch(e => setEstado("Error guardando (los cambios siguen en pantalla, reintenta con otra acción): " + e.message, true))
    .finally(() => { guardando = false; if (guardarDeNuevo) { guardarDeNuevo = false; programarGuardado(); } });
}

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
    const wrap = document.createElement("div"); wrap.className = "f";
    const lab = document.createElement("label"); lab.textContent = h || `Col ${c + 1}`;
    const sel = document.createElement("select");
    if (c === G.statusIdx) {
      sel.innerHTML = `<option value="">Todos</option><option value="__pend">Pendientes</option><option value="ok">OK</option><option value="nok">NOK</option>`;
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
    if (Number(c) === G.statusIdx) {
      const v = estadoNorm(row[G.statusIdx]);
      if (sel === "__pend" && (v === "ok" || v === "nok")) return false;
      if (sel === "ok" && v !== "ok") return false;
      if (sel === "nok" && v !== "nok") return false;
    } else if ((row[c] || "") !== sel) return false;
  }
  return true;
}

// ---------- Tabla ----------
function renderCabecera() {
  document.querySelector("#tabla thead").innerHTML =
    "<tr>" + G.header.map(h => `<th>${escapeHtml(h)}</th>`).join("") + "<th>Acción</th></tr>";
}
function renderTabla() {
  const tbody = document.querySelector("#tabla tbody");
  const frag = document.createDocumentFragment();
  let visibles = 0;
  G.rows.forEach((row, idx) => {
    if (!pasaFiltros(row)) return;
    visibles++;
    const v = estadoNorm(row[G.statusIdx]);
    const tr = document.createElement("tr");
    if (v === "ok") tr.className = "row-ok"; else if (v === "nok") tr.className = "row-nok";
    let html = "";
    G.header.forEach((_, c) => {
      if (c === G.statusIdx) {
        const cls = v === "ok" ? "ok" : v === "nok" ? "nok" : "pend";
        html += `<td class="estado ${cls}">${escapeHtml(row[c] || "—")}</td>`;
      } else {
        html += `<td>${escapeHtml(row[c] || "")}</td>`;
      }
    });
    html += `<td class="accion"><button class="ok" data-idx="${idx}" data-val="ok"${v === "ok" ? " disabled" : ""}>OK</button><button class="nok" data-idx="${idx}" data-val="nok"${v === "nok" ? " disabled" : ""}>NOK</button></td>`;
    tr.innerHTML = html;
    frag.appendChild(tr);
  });
  tbody.innerHTML = "";
  tbody.appendChild(frag);
  actualizarProgreso(visibles);
}
function actualizarProgreso(visibles) {
  const total = G.rows.length;
  const ok = G.rows.filter(r => estadoNorm(r[G.statusIdx]) === "ok").length;
  const nok = G.rows.filter(r => estadoNorm(r[G.statusIdx]) === "nok").length;
  const prog = document.getElementById("progreso");
  if (prog) prog.textContent = `${total} cargos · OK: ${ok} · NOK: ${nok} · pendientes: ${total - ok - nok}` + (visibles != null ? ` · mostrando: ${visibles}` : "");
}
function marcar(idx, valor) {
  const prev = G.rows[idx][G.statusIdx];
  if (estadoNorm(prev) === valor) return;
  G.rows[idx][G.statusIdx] = valor;
  renderTabla();
  programarGuardado();
}

// ---------- Carga ----------
async function iniciarRevision() {
  if (revisionIniciada) return;
  revisionIniciada = true;
  setEstado(`Buscando "${FILE_NAME}" en carpeta "${PREFIJO_CARPETA}*"…`);
  try {
    const token = await getToken(fileScopes);
    const item = await localizarFichero(token);
    if (!item) { revisionIniciada = false; return; }   // el mensaje de detalle ya lo puso localizarFichero
    G.fileId = item.id;
    const text = await descargarCsv(token);
    G.delim = detectarDelim(text);
    const all = parseCSV(text, G.delim);
    if (!all.length) { setEstado("El fichero está vacío.", true); return; }
    G.header = all[0];
    G.rows = all.slice(1);
    G.statusIdx = G.header.findIndex(h => h.trim().toLowerCase() === STATUS_COLUMN.toLowerCase());
    if (G.statusIdx === -1) { G.header.push(STATUS_COLUMN); G.statusIdx = G.header.length - 1; }
    G.rows.forEach(r => { while (r.length < G.header.length) r.push(""); });

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
async function handleLoginClick() {
  // Flujo por redirección en todos los dispositivos: evita los problemas de COOP con el popup.
  return msalInstance.loginRedirect(loginRequest);
}
async function getToken(request) {
  const active = msalInstance.getActiveAccount();
  const req = Object.assign({}, request, { account: request.account || active });
  try { return (await msalInstance.acquireTokenSilent(req)).accessToken; }
  catch (e) {
    console.warn("Silent falló, redirigiendo para renovar token:", e && e.errorCode);
    msalInstance.acquireTokenRedirect(request);   // navega y vuelve; tras el login suele bastar el silent
    return;
  }
}
function resetMsal() {
  Object.keys(localStorage).filter(k => k.startsWith("msal.") || k.includes(msalConfig.auth.clientId)).forEach(k => localStorage.removeItem(k));
  if (location.hash) history.replaceState(null, "", location.pathname + location.search);
  console.log("Estado MSAL limpiado.");
}
// Ayudante de consola: tuppersAuth.listar() lista la raíz de tu OneDrive
async function listarOneDrive(carpeta = "") {
  const token = await getToken(fileScopes);
  const url = carpeta
    ? `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURI(carpeta)}:/children`
    : `https://graph.microsoft.com/v1.0/me/drive/root/children?$top=200`;
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) { console.error("listar:", r.status, await r.text()); return; }
  console.table(((await r.json()).value || []).map(it => ({ nombre: it.name, tipo: it.folder ? "carpeta" : "fichero" })));
}

// ---------- Arranque ----------
document.querySelector("#tabla tbody").addEventListener("click", (e) => {
  const b = e.target.closest("button[data-val]");
  if (b) marcar(Number(b.dataset.idx), b.dataset.val);
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
