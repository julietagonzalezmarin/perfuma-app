import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import requests as std_requests
import urllib.parse
import unicodedata
import base64
import json
import glob
import os
import re
import io

st.set_page_config(page_title="PERFUMA Chile - Equivalencias", page_icon="🧪", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Montserrat', sans-serif; }
    .perfuma-header { text-align: center; padding: 5px 0; display: flex; align-items: center; justify-content: center; gap: 15px; }
    .linea-logo { height: 4px; width: 50px; background-color: #8E24AA; border-radius: 2px; }
    .perfuma-logo { font-size: 30px; font-weight: 700; letter-spacing: 6px; color: #111; margin: 0; }
    .perfuma-subtitle { text-align: center; font-size: 11px; color: #555; letter-spacing: 2px; margin-bottom: 12px; font-weight: 600; }
    .custom-divider { border-bottom: 1px solid #e2e8f0; margin: 10px 0; }
    div[data-testid="stColumn"]:has(button[kind="secondary"]) button {
        font-size: 26px !important; font-weight: 800 !important; height: 48px !important;
        border: 2px solid #111 !important; border-radius: 10px !important; background: #fff !important; color: #111 !important;
    }
    div[data-testid="stColumn"]:has(button[kind="secondary"]) button:hover {
        background: #8E24AA !important; color: #fff !important; border-color: #8E24AA !important;
    }
    </style>
    <div class="perfuma-header"><div class="linea-logo"></div><h1 class="perfuma-logo">PERFUMA</h1><div class="linea-logo"></div></div>
    <div class="perfuma-subtitle">ASISTENTE INTELIGENTE DE EQUIVALENCIAS OLFATIVAS</div>
    <div class="custom-divider"></div>
""", unsafe_allow_html=True)

ARCHIVO_DB = "catalogo_maestro.csv"
ARCHIVO_DB_BACKUP = "catalogo_maestro_respaldo.csv"
ARCHIVO_RELBASE = "relbase_actual.xlsx"
COLUMNAS_DB = ["SKU", "Perfume", "Marca", "Categoria", "Genero", "Precio", "Acordes", "Notas", "Foto_URL", "URL_Fragrantica", "Stock"]

MAPA_BODEGAS_EXCEL = {
    "📍 Paseo Estado": "BODEGA PASEO ESTADO",
    "📍 Paseo Puente Duomo": "BODEGA PUENTE DUOMO",
    "📍 Paseo Puente Árabe": "BODEGA PUENTE ARABE"
}
LISTA_TIENDAS = list(MAPA_BODEGAS_EXCEL.keys())

def sincronizar_archivo_nube(ruta_archivo, mensaje_commit="Actualización en vivo desde PERFUMA Admin"):
    try:
        if "GITHUB_TOKEN" not in st.secrets or "GITHUB_REPO" not in st.secrets:
            return False
        token = str(st.secrets["GITHUB_TOKEN"]).strip()
        repo = str(st.secrets["GITHUB_REPO"]).strip()
        branch = str(st.secrets.get("GITHUB_BRANCH", "main")).strip()
        if not token or not repo or not os.path.exists(ruta_archivo):
            return False

        url_api = f"https://api.github.com/repos/{repo}/contents/{ruta_archivo}"
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json"
        }
        r_get = std_requests.get(f"{url_api}?ref={branch}", headers=headers, timeout=15)
        sha_actual = r_get.json().get("sha") if r_get.status_code == 200 else None

        with open(ruta_archivo, "rb") as f:
            contenido_b64 = base64.b64encode(f.read()).decode("utf-8")

        payload = {
            "message": mensaje_commit,
            "content": contenido_b64,
            "branch": branch
        }
        if sha_actual:
            payload["sha"] = sha_actual

        r_put = std_requests.put(url_api, headers=headers, data=json.dumps(payload), timeout=25)
        return r_put.status_code in [200, 201]
    except Exception:
        return False

def convertir_df_a_excel(df_exportar, nombre_hoja="Datos"):
    buffer = io.BytesIO()
    df_copia = df_exportar.copy()
    if "SKU / Código" in df_copia.columns:
        df_copia["SKU / Código"] = df_copia["SKU / Código"].astype(str)
    if "SKU" in df_copia.columns:
        df_copia["SKU"] = df_copia["SKU"].astype(str)
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_copia.to_excel(writer, index=False, sheet_name=nombre_hoja)
    return buffer.getvalue()

def quitar_tildes(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', str(texto).lower().strip()) if unicodedata.category(c) != 'Mn')

SINONIMOS_ACORDES = {
    "citrico": "citrico", "citricos": "citrico", "vainilla": "avainillado", "avainillado": "avainillado",
    "frutal": "afrutados", "frutales": "afrutados", "afrutado": "afrutados", "afrutados": "afrutados",
    "rosa": "rosa", "rosas": "rosa", "floral blanco": "floral blanco", "florales blancos": "floral blanco",
    "oud": "madera de oud", "madera de oud": "madera de oud", "especiado suave": "especiado suave",
    "suave especiado": "especiado suave", "calido especiado": "calido especiado", "fresco especiado": "fresco especiado"
}

def normalizar_acorde(ac):
    base = re.sub(r'\s+', ' ', quitar_tildes(ac))
    return SINONIMOS_ACORDES.get(base, base)

COLORES_ACORDES = {
    "citrico": ("#FFFF33", "#1f1f1f"), "verde": ("#5CB85C", "#fff"), "almizclado": ("#E8DAEF", "#1f1f1f"),
    "aromatico": ("#26A69A", "#fff"), "fresco especiado": ("#8CD964", "#1f1f1f"), "fresco": ("#CCFFFF", "#1f1f1f"),
    "afrutados": ("#FF7043", "#fff"), "atalcado": ("#F5EEF8", "#1f1f1f"), "dulce": ("#F06292", "#fff"),
    "calido especiado": ("#D84315", "#fff"), "especiado suave": ("#E59866", "#1f1f1f"), "avainillado": ("#FFF59D", "#1f1f1f"),
    "amaderado": ("#B9936C", "#fff"), "ambar": ("#FFA726", "#fff"), "madera de oud": ("#5D4037", "#fff"),
    "pachuli": ("#795548", "#fff"), "terroso": ("#8D6E63", "#fff"), "acuatico": ("#4FC3F7", "#1f1f1f"),
    "marino": ("#0288D1", "#fff"), "cafe": ("#4E342E", "#fff"), "cacao": ("#6D4C41", "#fff"),
    "cuero": ("#757575", "#fff"), "ahumado": ("#90A4AE", "#fff"), "conac": ("#A93226", "#fff"),
    "salado": ("#B3E5FC", "#1f1f1f"), "floral blanco": ("#F8F9F9", "#1f1f1f"), "nardos": ("#E0F2F1", "#1f1f1f"),
    "floral": ("#F48FB1", "#1f1f1f"), "rosa": ("#EC407A", "#fff"), "tropical": ("#FFB74D", "#1f1f1f"),
    "coco": ("#FFE082", "#1f1f1f"), "lavanda": ("#B39DDB", "#1f1f1f"), "tabaco": ("#6E2C00", "#fff"),
    "solar": ("#FFF59D", "#1f1f1f"), "metalico": ("#CFD8DC", "#1f1f1f"), "canela": ("#BF360C", "#fff"),
    "iris": ("#CE93D8", "#1f1f1f"), "almendrado": ("#F5CBA7", "#1f1f1f"), "lactonico": ("#FFF9C4", "#1f1f1f"),
    "balsamico": ("#A1887F", "#fff"), "animalico": ("#8D6E63", "#fff"), "herbal": ("#66BB6A", "#1f1f1f"),
    "mineral": ("#B0BEC5", "#1f1f1f"), "musgoso": ("#689F38", "#fff"), "jabonoso": ("#E1F5FE", "#1f1f1f"),
    "acido": ("#DCE775", "#1f1f1f"), "miel": ("#FFD54F", "#1f1f1f"), "caramelo": ("#D2691E", "#fff")
}

def limpiar_nombre_perfume(t):
    if not t or t == "Búsqueda Manual":
        return "el perfume consultado"
    return re.sub(r'\s*(para Hombres y Mujeres|para Hombres|para Mujeres|for women and men|for women|for men).*$', '', t, flags=re.IGNORECASE).strip()

def detectar_genero_web(texto):
    t = str(texto).lower()
    if "hombres y mujeres" in t or "women and men" in t or "unisex" in t:
        return "Para Hombres y Mujeres"
    elif "para mujeres" in t or "for women" in t or "mujer" in t:
        return "Para Mujeres"
    elif "para hombres" in t or "for men" in t or "hombre" in t:
        return "Para Hombres"
    return "Para Hombres y Mujeres"

def es_perfume_valido_relbase(nombre):
    n = str(nombre).upper()
    excl = ['BOLSO', 'DESODORANTE', 'MIST', 'BODY', 'TESTER', 'SET', 'BOLSA', 'ENVASADOR', 'AIR FRESHENER', 'AMBIENTAL', 'CREMA', 'AFTER SHAVE', 'SHOWER', 'LOCION']
    return not any(e in n for e in excl)

def extraer_datos_fragrantica(url):
    url = str(url).strip()
    m_id = re.search(r'-(\d+)\.html', url)
    foto_url = f"https://fimgs.net/mdimg/perfume/375x500.{m_id.group(1)}.jpg" if m_id else ""
    cabeceras = {"Accept-Language": "es-ES,es;q=0.9,en;q=0.8", "Referer": "https://www.google.com/"}
    html_texto = ""

    try:
        from curl_cffi import requests as cffi_req
        for navegador in ["chrome116", "chrome120", "chrome", "safari15_5"]:
            try:
                resp = cffi_req.get(url, impersonate=navegador, headers=cabeceras, timeout=15)
                if resp.status_code == 200 and "just a moment" not in resp.text.lower():
                    html_texto = resp.text
                    break
            except Exception:
                continue
    except Exception:
        pass

    if not html_texto:
        try:
            cab_std = {**cabeceras, "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
            r_std = std_requests.get(url, headers=cab_std, timeout=15)
            if r_std.status_code == 200 and "just a moment" not in r_std.text.lower():
                html_texto = r_std.text
        except Exception:
            pass

    if not html_texto:
        return None, "Fragrantica bloqueó momentáneamente la lectura automática. Intenta en unos segundos."

    soup = BeautifulSoup(html_texto, "html.parser")
    t_pag = soup.find("title").text if soup.find("title") else ""
    h1 = soup.find("h1")
    nom = h1.text.strip() if h1 else "Perfume"
    gen = detectar_genero_web(f"{nom} {t_pag}")

    if not foto_url:
        meta_img = soup.find("meta", property="og:image")
        if meta_img and meta_img.get("content"):
            foto_url = meta_img["content"]

    acordes = []
    palabras_basura = ["acordes", "gratis", "siempre", "100%", "votos", "fotos", "buscar", "principales"]

    for el in soup.find_all(class_=lambda v: v and "accord" in v.lower()):
        hijos = el.find_all(class_=lambda v: v and "accord-bar" in v.lower())
        if hijos:
            for h in hijos:
                tx = h.get_text(strip=True).lower()
                if tx and len(tx) < 25 and tx not in acordes and not any(b in tx for b in palabras_basura):
                    acordes.append(tx)
        else:
            tx = el.get_text(strip=True).lower()
            if tx and len(tx) < 25 and tx not in acordes and not any(b in tx for b in palabras_basura):
                acordes.append(tx)

    if not acordes:
        for caja in soup.find_all("div", style=lambda s: s and "background" in s.lower() and ("color" in s.lower() or "width" in s.lower())):
            tx = caja.get_text(strip=True).lower()
            if tx and len(tx) < 25 and tx not in acordes and not any(b in tx for b in palabras_basura):
                acordes.append(tx)

    notas = [a.get_text(strip=True) for a in soup.find_all("a", href=True) if re.search(r'/(?:notas|notes)/[^/]+-\d+\.html', a["href"])]
    notas_limpias = ", ".join(list(dict.fromkeys(notas))[:10]) if notas else "Notas aromáticas, especiadas y amaderadas"

    if not acordes:
        return None, "No se encontraron acordes visibles en este enlace de Fragrantica."

    return {"nombre": limpiar_nombre_perfume(nom), "acordes": acordes, "notas": notas_limpias, "foto": foto_url, "genero": gen}, None

def normalizar_dataframe_db(df):
    df = df.copy().fillna("")
    for col_extra in [c for c in df.columns if str(c).startswith("Unnamed") or str(c) == ""]:
        df.drop(columns=[col_extra], inplace=True, errors="ignore")
    for col in COLUMNAS_DB:
        if col not in df.columns:
            df[col] = "Si" if col == "Stock" else ""
    df["SKU"] = df["SKU"].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    df["Stock"] = df["Stock"].apply(lambda x: "No" if str(x).strip().lower() == "no" else "Si")
    return df[COLUMNAS_DB]

def cargar_db():
    for ruta in [ARCHIVO_DB, ARCHIVO_DB_BACKUP]:
        if os.path.exists(ruta):
            for sep_intento in [",", ";"]:
                try:
                    df = pd.read_csv(ruta, dtype=str, sep=sep_intento, encoding="utf-8-sig").fillna("")
                    if "Perfume" in df.columns and len(df.columns) >= 5:
                        df_norm = normalizar_dataframe_db(df)
                        if len(df_norm) > 0:
                            return df_norm
                except Exception:
                    continue

    if not os.path.exists(ARCHIVO_DB):
        df_vacio = pd.DataFrame(columns=COLUMNAS_DB)
        df_vacio.to_csv(ARCHIVO_DB, index=False, encoding="utf-8-sig")
        return df_vacio

    return pd.DataFrame(columns=COLUMNAS_DB)

def guardar_db(df, mensaje_nube="Actualización de catálogo maestro"):
    df_limpio = normalizar_dataframe_db(df)
    df_limpio.to_csv(ARCHIVO_DB, index=False, encoding="utf-8-sig")
    if len(df_limpio) > 0:
        df_limpio.to_csv(ARCHIVO_DB_BACKUP, index=False, encoding="utf-8-sig")
    sincronizar_archivo_nube(ARCHIVO_DB, mensaje_nube)

def encontrar_excel_relbase():
    if os.path.exists(ARCHIVO_RELBASE):
        return ARCHIVO_RELBASE
    archivos = [f for f in sorted(glob.glob("*.xlsx"), key=os.path.getmtime, reverse=True) if "Catalogo" not in f and "Faltantes" not in f]
    if archivos:
        return archivos[0]
    return None

@st.cache_data(show_spinner=False)
def leer_excel_relbase_cacheado(ruta_excel, fecha_modificacion):
    df_raw = pd.read_excel(ruta_excel, sheet_name='Stock actual', skiprows=5)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    df_raw['SKU_Clave'] = df_raw['Código'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True).str.lstrip('0')
    return df_raw

def filtrar_inventario_por_tienda(df_catalogo, tienda_sel):
    ruta_excel = encontrar_excel_relbase()
    if not ruta_excel or not os.path.exists(ruta_excel):
        return df_catalogo.copy()

    try:
        mtime = os.path.getmtime(ruta_excel)
        df_raw = leer_excel_relbase_cacheado(ruta_excel, mtime).copy()
        col_bodega = MAPA_BODEGAS_EXCEL.get(tienda_sel, "")

        if col_bodega not in df_raw.columns:
            return df_catalogo.copy()

        df_raw[col_bodega] = pd.to_numeric(df_raw[col_bodega], errors='coerce').fillna(0)
        stock_por_sku = df_raw.groupby('SKU_Clave')[col_bodega].sum().to_dict()
        df_filtrado = df_catalogo.copy()

        def verificar_stock(sku_cat):
            s_limpio = str(sku_cat).strip().replace('.0', '').lstrip('0')
            if s_limpio in stock_por_sku:
                return "Si" if stock_por_sku[s_limpio] > 0 else "No"
            return "Si"

        df_filtrado["Stock"] = df_filtrado["SKU"].apply(verificar_stock)
        return df_filtrado
    except Exception:
        return df_catalogo.copy()

def generar_html_escala_acordes(lista_acordes, titulo="ESCALA DE ACORDES"):
    acordes = [a.strip().lower() for a in (re.split(r'[,;/]+', lista_acordes) if isinstance(lista_acordes, str) else lista_acordes) if a.strip()]
    if not acordes:
        return ""
    anchos = [100, 82, 70, 62, 56, 52, 48, 45, 42, 40]
    barras = ""
    for i, ac in enumerate(acordes[:10]):
        bg, fg = COLORES_ACORDES.get(normalizar_acorde(ac), ("#E0E0E0", "#1f1f1f"))
        barras += (f'<div style="width:100%;display:flex;"><div style="width:{anchos[i] if i < 10 else 40}%;'
                   f'background-color:{bg};color:{fg};text-align:center;padding:5px 8px;'
                   f'border-radius:0 10px 10px 0;font-size:12px;font-weight:500;border-bottom:1px solid rgba(255,255,255,0.4);">{ac}</div></div>')
    return (f'<div style="background:#fafafa;border:1px solid #e6e8eb;border-radius:10px;padding:12px 14px;">'
            f'<div style="text-align:center;font-weight:700;font-size:11px;color:#444;margin-bottom:6px;">📊 {titulo}</div>'
            f'<div style="max-width:260px;margin:0 auto;"><div style="text-align:center;font-size:12px;color:#333;margin-bottom:5px;">acordes principales</div>{barras}</div></div>')

def clave_unica(perfume, marca=""):
    tx = quitar_tildes(f"{marca} {perfume}")
    tx = re.sub(r'\b\d{2,3}\s*ml\b|\b(?:edp|edt|edc|parfum|eau|de|toilette|hombre|mujer|man|woman|unisex|set)\b', ' ', tx)
    return " ".join(dict.fromkeys(re.findall(r'[a-z0-9]+', tx)))

def calcular_similitud(ac_cli, ac_t_str):
    if isinstance(ac_cli, str):
        lista_cli = [normalizar_acorde(a) for a in re.split(r'[,;/]+', ac_cli) if a.strip()]
    else:
        lista_cli = [normalizar_acorde(a) for a in ac_cli if str(a).strip()]

    lista_t = [normalizar_acorde(a) for a in re.split(r'[,;/]+', str(ac_t_str)) if a.strip()]
    if not lista_cli or not lista_t:
        return 0.0, []
    p_c = {a: 1.0 / (i + 1) ** 0.5 for i, a in enumerate(dict.fromkeys(lista_cli))}
    p_t = {a: 1.0 / (i + 1) ** 0.5 for i, a in enumerate(dict.fromkeys(lista_t))}
    pts = sum(v * p_t[a] for a, v in p_c.items() if a in p_t)
    com = [a.capitalize() for a in p_c if a in p_t]
    denominador = (sum(v**2 for v in p_c.values())**0.5) * (sum(v**2 for v in p_t.values())**0.5)
    if denominador == 0:
        return 0.0, []
    return round((pts / denominador) * 100, 1), com

def generar_frase_cliente(pos, nom_ref, perf_t, marca_t, comunes, notas_str):
    ac_txt = ", ".join([a.lower() for a in comunes[:3]])
    notas_l = [n.strip().lower() for n in str(notas_str).split(",") if n.strip()][:3]
    n_txt = ", ".join(notas_l) if notas_l else "excelente fijación"
    if pos == 1:
        return f"«Si te gusta el estilo de {nom_ref}, te va a encantar {perf_t} de {marca_t}. Van por el mismo camino olfativo gracias a su carácter {ac_txt}, destacando en piel sus notas de {n_txt}.»"
    elif pos == 2:
        return f"«Otra excelente alternativa en esa misma línea es {perf_t} ({marca_t}), que mantiene esa vibra {ac_txt} con un toque muy rico de {n_txt}.»"
    return f"«Como tercera opción con el mismo ADN {ac_txt}, prueba cómo huele {perf_t} de {marca_t}; se parece muchísimo por su mezcla de {n_txt}.»"

for k, v in [("acordes_buscados", []), ("nombre_referencia", ""), ("genero_referencia", "Para Hombres y Mujeres"),
             ("foto_referencia", None), ("url_referencia", ""), ("indice_rec", 0),
             ("cat_ant", "Árabe/Diseñador"), ("gen_uni_ant", "Hombre"),
             ("tienda_ant", "📍 Paseo Estado"), ("admin_autenticado", False), ("mostrar_login_admin", False)]:
    if k not in st.session_state:
        st.session_state[k] = v

df_db_bruto = cargar_db()

col_tienda_top, col_espacio_top, col_btn_admin = st.columns([2.2, 3.0, 1.2])

with col_tienda_top:
    tienda_seleccionada = st.selectbox("🏪 Sucursal / Tienda Activa:", LISTA_TIENDAS, index=0)
    if tienda_seleccionada != st.session_state.tienda_ant:
        st.session_state.indice_rec = 0
        st.session_state.tienda_ant = tienda_seleccionada

with col_btn_admin:
    st.write("")
    if not st.session_state.admin_autenticado:
        if st.button("🔐 Acceso Admin", type="primary", use_container_width=True):
            st.session_state.mostrar_login_admin = not st.session_state.mostrar_login_admin
    else:
        if st.button("🔒 Cerrar Modo Admin", type="primary", use_container_width=True):
            st.session_state.admin_autenticado = False
            st.session_state.mostrar_login_admin = False
            st.rerun()

if st.session_state.mostrar_login_admin and not st.session_state.admin_autenticado:
    with st.container(border=True):
        c_p1, c_p2 = st.columns([3, 1])
        clave_ingresada = c_p1.text_input("Ingresa la contraseña de administrador:", type="password", key="input_pass_admin")
        c_p2.write("")
        if c_p2.button("Desbloquear", type="primary", use_container_width=True):
            if clave_ingresada == "Julieta25":
                st.session_state.admin_autenticado = True
                st.session_state.mostrar_login_admin = False
                st.rerun()
            else:
                st.error("⚠️ Contraseña incorrecta.")

df_db = filtrar_inventario_por_tienda(df_db_bruto, tienda_seleccionada)
df_activos = df_db[df_db["Stock"].str.strip().str.lower() != "no"].copy()

if st.session_state.admin_autenticado:
    tab_mostrador, tab_admin = st.tabs([
        f"🔍 Buscador para Clientes - {tienda_seleccionada} ({len(df_activos)} con stock en tienda / {len(df_db_bruto)} en catálogo total)",
        "➕ Ingresar / Gestionar Base de Datos & Radar Relbase"
    ])
else:
    tab_mostrador = st.container()
    tab_admin = None

with tab_mostrador:
    modo = st.radio("Método de búsqueda:", ["🔗 Pegar Link de Fragrantica", "✍️ Ingresar acordes manualmente"], horizontal=True)
    if modo == "🔗 Pegar Link de Fragrantica":
        url_input = st.text_input("URL de Fragrantica (.es o .com):", placeholder="https://www.fragrantica.es/perfume/Xerjoff/Erba-Pura-55157.html")
    else:
        c_m1, c_m2 = st.columns([2, 1])
        acordes_man = c_m1.text_input("Acordes separados por coma:", placeholder="cítrico, afrutados, dulce, almizclado")
        gen_man = c_m2.selectbox("Género buscado:", ["Para Hombres", "Para Hombres y Mujeres", "Para Mujeres"])

    c_btn, c_flt = st.columns([1.0, 3.2])
    buscar_click = c_btn.button("Buscar en Stock", type="primary", use_container_width=True)
    filtro_categoria = c_flt.radio("Categoría:", ["Árabe/Diseñador", "Árabe", "Diseñador"], horizontal=True, label_visibility="collapsed")

    if filtro_categoria != st.session_state.cat_ant:
        st.session_state.indice_rec = 0
        st.session_state.cat_ant = filtro_categoria

    if buscar_click:
        if modo == "🔗 Pegar Link de Fragrantica" and url_input.strip():
            with st.spinner("Extrayendo acordes desde Fragrantica..."):
                datos, err = extraer_datos_fragrantica(url_input.strip())
                if err:
                    st.error(f"⚠️ {err}")
                else:
                    st.session_state.update({"nombre_referencia": datos["nombre"], "acordes_buscados": datos["acordes"], "foto_referencia": datos["foto"], "genero_referencia": datos["genero"], "url_referencia": url_input.strip(), "indice_rec": 0})
        elif modo != "🔗 Pegar Link de Fragrantica" and acordes_man.strip():
            st.session_state.update({"nombre_referencia": "Búsqueda Manual", "acordes_buscados": [a.strip().lower() for a in re.split(r'[,;/]+', acordes_man) if a.strip()], "foto_referencia": None, "genero_referencia": gen_man, "url_referencia": "", "indice_rec": 0})

    if st.session_state.acordes_buscados:
        gen_ref = detectar_genero_web(st.session_state.genero_referencia)

        if gen_ref == "Para Hombres y Mujeres":
            c_r, c_g, c_f = st.columns([3.2, 2.6, 0.8])
            op_g = c_g.radio("👤 Recomendación para:", ["Hombre", "Mujer", "Unisex"], horizontal=True)
            if op_g != st.session_state.gen_uni_ant:
                st.session_state.indice_rec = 0
                st.session_state.gen_uni_ant = op_g

            if op_g == "Hombre":
                genero_objetivo = "Para Hombres"
            elif op_g == "Mujer":
                genero_objetivo = "Para Mujeres"
            else:
                genero_objetivo = "Para Hombres y Mujeres"
        else:
            c_r, _, c_f = st.columns([4.5, 1.5, 0.8])
            genero_objetivo = gen_ref

        c_r.success(f"**Referencia:** {st.session_state.nombre_referencia}  |  **Género en Fragrantica:** `{gen_ref}`")

        agrupados = {}
        for _, fila in df_activos.iterrows():
            if st.session_state.url_referencia and fila["URL_Fragrantica"].strip() == st.session_state.url_referencia:
                continue
            if st.session_state.nombre_referencia != "Búsqueda Manual" and clave_unica(fila["Perfume"], fila["Marca"]) == clave_unica(st.session_state.nombre_referencia, fila["Marca"]):
                continue

            gen_fila = detectar_genero_web(fila["Genero"])
            if gen_fila != genero_objetivo:
                continue

            cat_f = fila["Categoria"].strip().upper()
            if filtro_categoria == "Árabe" and cat_f not in ["ÁRABE", "ARABE"]:
                continue
            elif filtro_categoria == "Diseñador" and cat_f != "DISEÑADOR":
                continue
            elif filtro_categoria == "Árabe/Diseñador" and cat_f not in ["ÁRABE", "ARABE", "DISEÑADOR"]:
                continue

            score, comunes = calcular_similitud(st.session_state.acordes_buscados, fila["Acordes"])
            if score > 0:
                cf = clave_unica(fila["Perfume"], fila["Marca"])
                if cf not in agrupados:
                    agrupados[cf] = {**fila.to_dict(), "Similitud": score, "Comunes": comunes}

        resultados = sorted(agrupados.values(), key=lambda x: x["Similitud"], reverse=True)[:3]
        with c_f:
            if len(resultados) > 1 and st.button("➡", help="Siguiente recomendación", use_container_width=True):
                st.session_state.indice_rec = (st.session_state.indice_rec + 1) % len(resultados)

        if not resultados:
            st.info(f"No se encontraron perfumes con stock en **{tienda_seleccionada}** para el género **{genero_objetivo}** dentro de la categoría **{filtro_categoria}**.")
        else:
            idx = st.session_state.indice_rec % len(resultados)
            res = resultados[idx]
            c1, c2, c3, c4 = st.columns([1.0, 1.35, 1.25, 3.2], gap="small")
            with c1:
                with st.container(border=True):
                    st.caption("📸 **Perfume buscado:**")
                    if st.session_state.foto_referencia:
                        st.image(st.session_state.foto_referencia, use_container_width=True)
            with c2:
                st.markdown(generar_html_escala_acordes(st.session_state.acordes_buscados, "PERFUME BUSCADO"), unsafe_allow_html=True)
            with c3:
                with st.container(border=True):
                    st.caption(f"🧴 **Opción #{idx+1} en {tienda_seleccionada} ({res['Categoria']}):**")
                    m_id = re.search(r'-(\d+)\.html', str(res.get("URL_Fragrantica", "")))
                    foto_f = f"https://fimgs.net/mdimg/perfume/375x500.{m_id.group(1)}.jpg" if m_id else res.get("Foto_URL", "")
                    foto_mostrar = foto_f if str(foto_f).startswith("http") else f"https://tse2.mm.bing.net/th?q={urllib.parse.quote(res['Marca'] + ' ' + res['Perfume'] + ' perfume bottle')}&w=350&h=420&c=7&rs=1&p=0"
                    st.image(foto_mostrar, caption=f"{res['Perfume']} - {res['Marca']}", use_container_width=True)
            with c4:
                with st.container(border=True):
                    cd, ce = st.columns([1.35, 1.05], gap="medium")
                    with cd:
                        st.markdown(f"### #{idx+1} - {res['Perfume']} ({res['Marca']})")
                        st.write(f"**SKU:** `{res['SKU'] or 'S/N'}` | **Línea:** `{res['Categoria']}`")
                        st.write(f"**Tienda:** `{tienda_seleccionada}`")
                        st.write(f"**Género:** {res['Genero']} | **Precio:** {res['Precio']}")
                        st.write(f"**Coinciden en:** {', '.join(res['Comunes'])}")
                        st.caption(f"**Notas clave:** {res['Notas']}")
                        st.metric("Nivel de Afinidad", f"{res['Similitud']}%")
                        st.progress(min(int(res["Similitud"]), 100))
                    with ce:
                        st.markdown(generar_html_escala_acordes(res["Acordes"], f"ACORDES DE {res['Perfume'].upper()}"), unsafe_allow_html=True)
                    st.info(f"💬 **Frase sugerida:**\n\n*{generar_frase_cliente(idx+1, st.session_state.nombre_referencia, res['Perfume'], res['Marca'], res['Comunes'], res['Notas'])}*")

if tab_admin is not None:
    with tab_admin:
        with st.expander("🛟 Restaurar / Importar Catálogo desde tu archivo exportado (.csv o .xlsx)", expanded=(len(df_db_bruto) < 5)):
            st.write("Selecciona cualquier respaldo `.csv` o `.xlsx` para restaurar y unificar tus perfumes:")
            archivo_restaurar = st.file_uploader("Seleccionar respaldo de catálogo (.csv o .xlsx)", type=["csv", "xlsx"], key="uploader_restore")
            if archivo_restaurar is not None:
                if st.button("✅ Restaurar y Unificar Perfumes Ahora", type="primary"):
                    try:
                        if archivo_restaurar.name.endswith(".csv"):
                            df_imp = pd.read_csv(archivo_restaurar, dtype=str).fillna("")
                        else:
                            df_imp = pd.read_excel(archivo_restaurar, dtype=str).fillna("")
                        df_imp_norm = normalizar_dataframe_db(df_imp)
                        df_combinado = pd.concat([df_db_bruto, df_imp_norm], ignore_index=True)
                        df_combinado = df_combinado.drop_duplicates(subset=["SKU", "Perfume"], keep="last")
                        guardar_db(df_combinado, "Restauración de catálogo desde archivo")
                        st.success(f"🎉 ¡Se restauraron y sincronizaron exitosamente {len(df_combinado)} perfumes!")
                        st.rerun()
                    except Exception as e_rest:
                        st.error(f"⚠️ Error al leer el archivo de respaldo: {e_rest}")

        st.markdown(f"### 🔍 Radar de Faltantes: Perfumes físicos en `{tienda_seleccionada}` sin registrar")
        ruta_ex_radar = encontrar_excel_relbase()
        if not ruta_ex_radar or not os.path.exists(ruta_ex_radar):
            st.warning("⚠️ Aún no has subido un archivo Excel de Relbase. Súbelo en la sección de abajo para activar el radar.")
        else:
            try:
                mtime_r = os.path.getmtime(ruta_ex_radar)
                df_ex_raw = leer_excel_relbase_cacheado(ruta_ex_radar, mtime_r).copy()
                col_b_radar = MAPA_BODEGAS_EXCEL.get(tienda_seleccionada, "")

                if col_b_radar in df_ex_raw.columns:
                    df_ex_raw = df_ex_raw[df_ex_raw['Producto'].apply(es_perfume_valido_relbase)].copy()
                    df_ex_raw[col_b_radar] = pd.to_numeric(df_ex_raw[col_b_radar], errors='coerce').fillna(0)

                    df_con_stock = df_ex_raw[df_ex_raw[col_b_radar] > 0].copy()
                    df_agrupado_stock = df_con_stock.groupby('SKU_Clave', as_index=False).agg({
                        'Producto': 'first',
                        'Categoría': 'first',
                        'Código': 'first',
                        'Precio neto': 'first',
                        col_b_radar: 'sum'
                    })

                    skus_en_db = set(df_db_bruto["SKU"].astype(str).str.strip().str.replace(r'\.0$', '', regex=True).str.lstrip('0'))
                    df_faltantes = df_agrupado_stock[~df_agrupado_stock['SKU_Clave'].isin(skus_en_db)].copy()
                    df_faltantes = df_faltantes.sort_values(by=col_b_radar, ascending=False)

                    if df_faltantes.empty:
                        st.success(f"🎉 ¡Todos los perfumes con stock físico en `{tienda_seleccionada}` ya están registrados en tu Catálogo Maestro!")
                    else:
                        tabla_mostrar = df_faltantes[['Código', 'Producto', 'Categoría', col_b_radar, 'Precio neto']].copy()
                        tabla_mostrar.columns = ['SKU / Código', 'Nombre en Relbase', 'Marca / Categoría', 'Stock en Tienda', 'Precio Neto']

                        c_info_f, c_btn_desc_f = st.columns([2.2, 1.2])
                        with c_info_f:
                            st.info(f"📋 Se encontraron **{len(df_faltantes)} perfumes con stock en `{tienda_seleccionada}`** que aún no están en tu `catalogo_maestro.csv`:")
                        with c_btn_desc_f:
                            nombre_limpio_suc = re.sub(r'[^A-Za-z0-9]+', '_', quitar_tildes(tienda_seleccionada)).strip('_')
                            st.download_button(
                                label="📥 Descargar Faltantes en Excel (.xlsx)",
                                data=convertir_df_a_excel(tabla_mostrar, "Faltantes"),
                                file_name=f"Perfumes_Faltantes_{nombre_limpio_suc}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary",
                                use_container_width=True
                            )
                        st.dataframe(tabla_mostrar, use_container_width=True, hide_index=True)
                else:
                    st.error(f"⚠️ No se encontró la columna `{col_b_radar}` en el Excel de Relbase.")
            except Exception as ex:
                st.error(f"⚠️ Error al procesar el radar de faltantes: {ex}")

        st.divider()
        st.markdown("### 📁 Subir Reporte XLSX de Relbase (Sincronización en Vivo para todas las Sucursales)")
        archivo_relbase_subido = st.file_uploader("Subir stock_actual.xlsx", type=["xlsx", "xls"])
        if archivo_relbase_subido is not None:
            if st.button("☁️ Actualizar Stock en Todas las Sucursales Ahora", type="primary"):
                with st.spinner("Actualizando y sincronizando stock con todas las sucursales..."):
                    with open(ARCHIVO_RELBASE, "wb") as f:
                        f.write(archivo_relbase_subido.getbuffer())
                    leer_excel_relbase_cacheado.clear()
                    sincronizar_archivo_nube(ARCHIVO_RELBASE, "Actualización de stock Relbase")
                st.success("✅ ¡Reporte de Relbase actualizado en vivo para todas las sucursales!")
                st.rerun()

        st.divider()
        st.markdown("### ➕ Agregar Perfume por Link de Fragrantica (`catalogo_maestro.csv`)")
        col_f1, col_f2 = st.columns(2, gap="large")
        with col_f1:
            url_nuevo = st.text_input("1. Pega el Link de Fragrantica (Lee acordes, notas y foto solo):", placeholder="https://www.fragrantica.es/perfume/Afnan/Tribute-Blue-42264.html")
            nombre_nuevo = st.text_input("2. Nombre del Perfume:", placeholder="Afnan Tribute Blue")
            marca_nueva = st.text_input("3. Marca:", placeholder="Afnan")
        with col_f2:
            sku_nuevo = st.text_input("4. Código / SKU (Copia el SKU desde el Radar de arriba):", placeholder="6290171002338")
            cat_nueva = st.selectbox("5. Categoría en el Programa:", ["Árabe", "Diseñador"])
            gen_nuevo = st.selectbox("6. Género:", ["Auto-detectar de Fragrantica", "Para Hombres", "Para Mujeres", "Para Hombres y Mujeres"])
            precio_nuevo = st.text_input("7. Precio en Tienda:", value="$29.990")

        if st.button("💾 Leer de Fragrantica y Guardar en Todas las Sucursales", type="primary", use_container_width=True):
            if not url_nuevo.strip() or not nombre_nuevo.strip() or not marca_nueva.strip():
                st.warning("⚠️ Completa el Link de Fragrantica, el Nombre del Perfume y la Marca.")
            else:
                with st.spinner("Leyendo de Fragrantica y sincronizando con todas las tiendas..."):
                    datos_f, err_f = extraer_datos_fragrantica(url_nuevo.strip())
                    if err_f:
                        st.error(f"⚠️ {err_f}")
                    else:
                        gen_guardar = datos_f["genero"] if gen_nuevo == "Auto-detectar de Fragrantica" else gen_nuevo
                        nueva_fila = pd.DataFrame([{
                            "SKU": sku_nuevo.strip() or str(len(df_db_bruto) + 1),
                            "Perfume": nombre_nuevo.strip(),
                            "Marca": marca_nueva.strip(),
                            "Categoria": cat_nueva,
                            "Genero": gen_guardar,
                            "Precio": precio_nuevo.strip(),
                            "Acordes": ", ".join(datos_f["acordes"]),
                            "Notas": datos_f["notas"],
                            "Foto_URL": datos_f["foto"],
                            "URL_Fragrantica": url_nuevo.strip(),
                            "Stock": "Si"
                        }])
                        guardar_db(pd.concat([df_db_bruto, nueva_fila], ignore_index=True), f"Agregado {nombre_nuevo.strip()}")
                        st.success(f"✅ ¡**{nombre_nuevo}** guardado y sincronizado en todas las sucursales!")
                        st.rerun()

        st.divider()
        c_tit_inv, c_btn_inv = st.columns([2.2, 1.2])
        with c_tit_inv:
            st.markdown(f"### 📋 Inventario Total en tu Base de Datos ({len(df_db_bruto)} perfumes)")
        with c_btn_inv:
            st.download_button(
                label="📥 Descargar Catálogo en Excel (.xlsx)",
                data=convertir_df_a_excel(df_db_bruto, "Catalogo Maestro"),
                file_name="Catalogo_Maestro_Perfuma.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        df_editado = st.data_editor(df_db_bruto, num_rows="dynamic", use_container_width=True, key="editor_db")
        if st.button("🔄 Guardar Cambios hechos en la Tabla"):
            with st.spinner("Guardando y sincronizando en la nube..."):
                guardar_db(df_editado, "Edición manual de tabla")
            st.success("✅ ¡Tabla actualizada y sincronizada en todas las sucursales!")
            st.rerun()
