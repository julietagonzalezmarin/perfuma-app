// PERFUMA — servidor de administración (Cloudflare Worker)
//
// Aquí viven los secretos que antes estaban a la vista en index.html.
// Se configuran como "Secret" en el panel de Cloudflare, nunca en el código:
//   GITHUB_TOKEN    token de GitHub con permiso de escritura SOLO en este repositorio
//   GITHUB_REPO     por ejemplo: julietagonzalezmarin/perfuma-app
//   ADMIN_PASSWORD  contraseña de administración
//   GITHUB_BRANCH   opcional, por defecto "main"
//
// Rutas:
//   POST /api/login    { password }                        -> { sesion }
//   GET  /api/datos                                        -> { catalogo, descartados }  (texto CSV al día)
//   POST /api/cambios  { agregar, descartar, restaurar }   -> agrega perfumes y actualiza descartados
//   POST /api/stock    { contenido }  (xlsx en base64)     -> reemplaza relbase_actual.xlsx
//
// Todo lo que no empiece con /api/ lo sirve Cloudflare como archivo estático
// desde la carpeta public/, que arma el comando de build.

const ARCHIVO_CATALOGO = "catalogo_maestro.csv";
const ARCHIVO_DESCARTADOS = "descartados.csv";
const ARCHIVO_STOCK = "relbase_actual.xlsx";
const HORAS_SESION = 12;
const MAX_STOCK_BASE64 = 8_000_000; // ~6 MB de Excel
const MAX_FILAS = 500;

const codificador = new TextEncoder();

function responder(objeto, status = 200) {
  return new Response(JSON.stringify(objeto), {
    status,
    headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" }
  });
}

// ---------- Sesiones firmadas (sin base de datos) ----------
function base64Url(buffer) {
  let s = "";
  for (const b of new Uint8Array(buffer)) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

async function firmar(clave, texto) {
  const k = await crypto.subtle.importKey("raw", codificador.encode(clave), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  return base64Url(await crypto.subtle.sign("HMAC", k, codificador.encode(texto)));
}

function iguales(a, b) {
  if (a.length !== b.length) return false;
  let d = 0;
  for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}

// Cambiar la contraseña o el token cierra todas las sesiones abiertas.
const claveSesion = env => `${env.ADMIN_PASSWORD}::${env.GITHUB_TOKEN}`;

async function crearSesion(env) {
  const vence = Date.now() + HORAS_SESION * 3600 * 1000;
  return `${vence}.${await firmar(claveSesion(env), "sesion:" + vence)}`;
}

async function sesionValida(env, request) {
  const cab = request.headers.get("authorization") || "";
  const [vence, firma] = cab.replace(/^Bearer\s+/i, "").split(".");
  if (!vence || !firma || !(Number(vence) > Date.now())) return false;
  return iguales(firma, await firmar(claveSesion(env), "sesion:" + vence));
}

// ---------- GitHub ----------
function base64ATexto(b64) {
  const bin = atob(String(b64 || "").replace(/\s/g, ""));
  return new TextDecoder().decode(Uint8Array.from(bin, c => c.charCodeAt(0)));
}

function textoABase64(texto) {
  const bytes = codificador.encode(texto);
  let s = "";
  for (let i = 0; i < bytes.length; i += 8192) s += String.fromCharCode(...bytes.subarray(i, i + 8192));
  return btoa(s);
}

function github(env, metodo, archivo, cuerpo) {
  const rama = env.GITHUB_BRANCH || "main";
  const url = `https://api.github.com/repos/${env.GITHUB_REPO}/contents/${archivo}` +
    (metodo === "GET" ? `?ref=${encodeURIComponent(rama)}` : "");
  return fetch(url, {
    method: metodo,
    headers: {
      authorization: `Bearer ${env.GITHUB_TOKEN}`,
      accept: "application/vnd.github+json",
      "x-github-api-version": "2022-11-28",
      "user-agent": "perfuma-app",
      ...(cuerpo ? { "content-type": "application/json" } : {})
    },
    body: cuerpo ? JSON.stringify({ ...cuerpo, branch: rama }) : undefined
  });
}

function mensajeGitHub(status) {
  if (status === 401) return "El token de GitHub no es válido o venció. Crea uno nuevo y actualízalo en Cloudflare.";
  if (status === 403 || status === 404) return "El token no tiene permiso de escritura en el repositorio, o GITHUB_REPO está mal escrito.";
  if (status === 409 || status === 422) return "Alguien más guardó al mismo tiempo. Vuelve a intentarlo.";
  return `GitHub respondió con el código ${status}. Vuelve a intentarlo en un momento.`;
}

class ErrorGitHub extends Error {
  constructor(status) { super(mensajeGitHub(status)); this.status = status; }
}

async function leerArchivo(env, archivo, conTexto = true) {
  const r = await github(env, "GET", archivo);
  if (r.status === 404) return { existe: false, sha: undefined, texto: "" };
  if (!r.ok) throw new ErrorGitHub(r.status);
  const j = await r.json();
  return { existe: true, sha: j.sha, texto: conTexto ? base64ATexto(j.content) : "" };
}

async function escribirArchivo(env, archivo, contenidoBase64, sha, mensaje) {
  const r = await github(env, "PUT", archivo, {
    message: String(mensaje || "Actualización desde Perfuma").slice(0, 120),
    content: contenidoBase64,
    ...(sha ? { sha } : {})
  });
  if (!r.ok) throw new ErrorGitHub(r.status);
}

// ---------- CSV ----------
function csvCampo(v) {
  const s = String(v ?? "").slice(0, 2000);
  return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

function csvLeer(texto) {
  const filas = [];
  let fila = [], campo = "", comillas = false;
  for (let i = 0; i < texto.length; i++) {
    const c = texto[i];
    if (comillas) {
      if (c === '"') { if (texto[i + 1] === '"') { campo += '"'; i++; } else comillas = false; }
      else campo += c;
    } else if (c === '"') comillas = true;
    else if (c === ",") { fila.push(campo); campo = ""; }
    else if (c === "\n") { fila.push(campo); filas.push(fila); fila = []; campo = ""; }
    else if (c !== "\r") campo += c;
  }
  if (campo !== "" || fila.length) { fila.push(campo); filas.push(fila); }
  return filas.filter(f => f.some(x => x !== ""));
}

const claveSku = s => String(s ?? "").trim().replace(/\.0$/, "").replace(/^0+/, "");

// ---------- Fragrantica ----------
// Fragrantica bloquea la lectura directa, así que se lee a través de r.jina.ai.
// Hacerlo aquí y no en el navegador evita el bloqueo entre sitios, y la caché
// queda compartida por las tres tiendas.
const RE_FRAGRANTICA = /^https:\/\/(?:www\.)?fragrantica\.(?:es|com)\/perfume\/[^\/?#]+\/[^\/?#]*-(\d+)\.html$/i;
const DIAS_CACHE = 180;

async function rutaFragrantica(request, env) {
  const url = (new URL(request.url).searchParams.get("url") || "").trim();
  const m = url.match(RE_FRAGRANTICA);
  if (!m) return responder({ error: "Ese link no es la página de un perfume en Fragrantica." }, 400);
  const llave = "frag_" + m[1];

  if (env.CACHE_FRAG) {
    const guardado = await env.CACHE_FRAG.get(llave);
    if (guardado) return responder({ texto: guardado, deCache: true });
  }

  let texto = "";
  try {
    const r = await fetch("https://r.jina.ai/" + url, {
      headers: { accept: "text/plain", "user-agent": "perfuma-app" },
      signal: AbortSignal.timeout(22000)
    });
    if (r.ok) texto = await r.text();
  } catch (e) {}

  if (!texto || !/acordes principales|main accords/i.test(texto)) {
    return responder({ error: "Fragrantica no respondió a tiempo. Inténtalo de nuevo en unos segundos, o escribe los acordes a mano." }, 502);
  }

  texto = texto.slice(0, 120000);
  if (env.CACHE_FRAG) {
    try { await env.CACHE_FRAG.put(llave, texto, { expirationTtl: 60 * 60 * 24 * DIAS_CACHE }); } catch (e) {}
  }
  return responder({ texto });
}

// ---------- Rutas ----------
async function rutaLogin(request, env) {
  const { password } = await request.json().catch(() => ({}));
  // Se comparan las firmas (mismo largo siempre) para no filtrar información por tiempo de respuesta.
  const [a, b] = await Promise.all([
    firmar("comparar", String(password ?? "")),
    firmar("comparar", String(env.ADMIN_PASSWORD))
  ]);
  if (!iguales(a, b)) {
    await new Promise(r => setTimeout(r, 1000)); // frena los intentos de adivinar la contraseña
    return responder({ error: "Contraseña incorrecta." }, 401);
  }
  return responder({ sesion: await crearSesion(env), horas: HORAS_SESION });
}

async function rutaDatos(env) {
  const [cat, desc] = await Promise.all([leerArchivo(env, ARCHIVO_CATALOGO), leerArchivo(env, ARCHIVO_DESCARTADOS)]);
  return responder({ catalogo: cat.texto, descartados: desc.texto });
}

async function rutaCambios(request, env) {
  const cuerpo = await request.json().catch(() => ({}));
  const lista = x => (Array.isArray(x) ? x.slice(0, MAX_FILAS) : []);
  const agregar = lista(cuerpo.agregar), descartar = lista(cuerpo.descartar), restaurar = lista(cuerpo.restaurar);
  let agregados = 0;

  if (agregar.length) {
    const cat = await leerArchivo(env, ARCHIVO_CATALOGO);
    const bom = cat.texto.startsWith("﻿") ? "﻿" : "";
    const texto = cat.texto.replace(/^﻿/, "");
    const filas = csvLeer(texto);
    const cabecera = filas[0] || [];
    const iSku = cabecera.indexOf("SKU");
    if (!cat.existe || iSku === -1 || !cabecera.includes("Perfume")) {
      return responder({ error: "No se pudo leer catalogo_maestro.csv en el repositorio. No se guardó nada." }, 409);
    }
    const existentes = new Set(filas.slice(1).map(f => claveSku(f[iSku])));
    const lineas = [];
    for (const p of agregar) {
      const k = claveSku(p && p.SKU);
      if (!p || !String(p.Perfume || "").trim() || !k || existentes.has(k)) continue;
      existentes.add(k);
      lineas.push(cabecera.map(col => csvCampo(p[col])).join(","));
    }
    if (lineas.length) {
      const nuevo = bom + texto.replace(/\s+$/, "") + "\n" + lineas.join("\n") + "\n";
      const nombre = lineas.length === 1 ? String(agregar[0].Perfume) : `${lineas.length} perfumes`;
      await escribirArchivo(env, ARCHIVO_CATALOGO, textoABase64(nuevo), cat.sha, `Catálogo: agregado ${nombre}`);
      agregados = lineas.length;
    }
  }

  if (descartar.length || restaurar.length) {
    const d = await leerArchivo(env, ARCHIVO_DESCARTADOS);
    const mapa = new Map();
    csvLeer(d.texto.replace(/^﻿/, "")).slice(1).forEach(f => { if (claveSku(f[0])) mapa.set(claveSku(f[0]), f); });
    const antes = JSON.stringify([...mapa]);
    const hoy = new Date().toISOString().slice(0, 10);
    restaurar.forEach(s => mapa.delete(claveSku(s)));
    descartar.forEach(x => {
      const k = claveSku(x && x.SKU);
      if (k && !mapa.has(k)) mapa.set(k, [String(x.SKU).trim(), String(x.Producto || ""), hoy]);
    });
    if (JSON.stringify([...mapa]) !== antes) {
      const csv = ["SKU,Producto,Fecha", ...[...mapa.values()].map(f => f.slice(0, 3).map(csvCampo).join(","))].join("\n") + "\n";
      await escribirArchivo(env, ARCHIVO_DESCARTADOS, textoABase64(csv), d.sha, `Sugerencias: ${mapa.size} perfumes marcados como "no agregar"`);
    }
  }

  return responder({ ok: true, agregados });
}

async function rutaStock(request, env) {
  const { contenido } = await request.json().catch(() => ({}));
  if (typeof contenido !== "string" || !contenido.startsWith("UEsDB")) {
    return responder({ error: "El archivo no es un Excel .xlsx válido." }, 400);
  }
  if (contenido.length > MAX_STOCK_BASE64) return responder({ error: "El Excel es demasiado grande (máximo 6 MB)." }, 413);
  const actual = await leerArchivo(env, ARCHIVO_STOCK, false);
  await escribirArchivo(env, ARCHIVO_STOCK, contenido, actual.sha, "Stock: actualización desde Relbase");
  return responder({ ok: true });
}

export async function manejar(request, env) {
  const ruta = new URL(request.url).pathname.replace(/\/+$/, "").split("/api/")[1] || "";
  try {
    // La búsqueda es pública y no depende de la configuración de administración.
    if (ruta === "fragrantica" && request.method === "GET") return await rutaFragrantica(request, env);
  } catch (e) {
    return responder({ error: "No se pudo leer Fragrantica en este momento." }, 502);
  }
  if (!env.GITHUB_TOKEN || !env.GITHUB_REPO || !env.ADMIN_PASSWORD) {
    return responder({ error: "Falta configurar GITHUB_TOKEN, GITHUB_REPO y ADMIN_PASSWORD en Cloudflare (Settings → Variables and Secrets)." }, 503);
  }
  try {
    if (ruta === "login" && request.method === "POST") return await rutaLogin(request, env);
    if (!["datos", "cambios", "stock"].includes(ruta)) return responder({ error: "Ruta no encontrada." }, 404);
    if (!(await sesionValida(env, request))) return responder({ error: "La sesión venció. Vuelve a entrar." }, 401);
    if (ruta === "datos" && request.method === "GET") return await rutaDatos(env);
    if (ruta === "cambios" && request.method === "POST") return await rutaCambios(request, env);
    if (ruta === "stock" && request.method === "POST") return await rutaStock(request, env);
    return responder({ error: "Método no permitido." }, 405);
  } catch (e) {
    if (e instanceof ErrorGitHub) return responder({ error: e.message }, 502);
    return responder({ error: "Error inesperado en el servidor." }, 500);
  }
}

export default {
  async fetch(request, env) {
    const { pathname } = new URL(request.url);
    if (pathname.startsWith("/api/")) return manejar(request, env);
    // Cualquier otra dirección: devolver la página principal.
    return env.ASSETS.fetch(new URL("/index.html", request.url).toString());
  }
};
