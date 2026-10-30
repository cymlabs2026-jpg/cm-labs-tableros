# C&M LABS - Tableros Automáticos v3
# Áreas y subtableros + análisis avanzado + temas de color + IA del cliente
# (Claude, ChatGPT, Gemini o Azure OpenAI) para los insights.
# pip install streamlit pandas numpy openpyxl plotly
# streamlit run app_tableros_v2.py
import streamlit as st, pandas as pd, numpy as np, plotly.express as px, plotly.graph_objects as go
import json, re, urllib.request, urllib.error

st.set_page_config(page_title="C&M LABS · Tableros", layout="wide")

def pct(x): return f"{x:.1f}%"
def num(x): return f"{x:,.0f}".replace(",", ".")
def money(x): return "$" + num(x)
def C(d, m, campo): return d[m[campo]]                      # columna real de un campo estándar
def B(d, m, campo): return C(d, m, campo).astype(str).str.lower().isin(["1","true","si","sí","x","yes","verdadero"])

# ---------------------------------------------------------------------------
# CATÁLOGO: área -> subtablero -> {claves para reconocerlo, indicadores, gráficos}
# Indicador: (campos requeridos, cálculo, formato, meta opcional, activo por defecto)
# Gráfico:   (tipo, campo eje X, campo valor o "conteo", agregación, título)
# Para crear un subtablero nuevo, solo se agrega un bloque aquí.
# ---------------------------------------------------------------------------
CAT = {
"Logística": {
 "Exactitud de inventario": {
  "claves": ["stock_sistema","stock_fisico","conteo","diferencia","ubicacion"],
  "ind": {
   "Exactitud de inventario": (["stock_sistema","stock_fisico"], lambda d,m: (1-(C(d,m,"stock_sistema")-C(d,m,"stock_fisico")).abs().sum()/C(d,m,"stock_sistema").sum())*100, pct, "≥ 98%", True),
   "Referencias con diferencia": (["stock_sistema","stock_fisico"], lambda d,m: (C(d,m,"stock_sistema")!=C(d,m,"stock_fisico")).sum(), num, None, True),
   "Valor de la diferencia": (["stock_sistema","stock_fisico","costo_unit"], lambda d,m: ((C(d,m,"stock_fisico")-C(d,m,"stock_sistema"))*C(d,m,"costo_unit")).sum(), money, None, True),
   "Unidades faltantes": (["stock_sistema","stock_fisico"], lambda d,m: (C(d,m,"stock_sistema")-C(d,m,"stock_fisico")).clip(lower=0).sum(), num, None, False),
   "Unidades sobrantes": (["stock_sistema","stock_fisico"], lambda d,m: (C(d,m,"stock_fisico")-C(d,m,"stock_sistema")).clip(lower=0).sum(), num, None, False),
  },
  "graf": [("bar","categoria","_dif_abs","sum","Diferencia absoluta por categoría"),
           ("bar","ubicacion","_dif_abs","sum","Diferencia por ubicación"),
           ("top","sku","_dif_abs","sum","Top 10 referencias con más diferencia")],
 },
 "Averías": {
  "claves": ["averia","daño","danio","causa","merma","rotura","novedad"],
  "ind": {
   "Unidades averiadas": (["cantidad_averiada"], lambda d,m: C(d,m,"cantidad_averiada").sum(), num, None, True),
   "Costo de averías": (["cantidad_averiada","costo_unit"], lambda d,m: (C(d,m,"cantidad_averiada")*C(d,m,"costo_unit")).sum(), money, None, True),
   "Índice de avería": (["cantidad_averiada","cantidad"], lambda d,m: C(d,m,"cantidad_averiada").sum()/C(d,m,"cantidad").sum()*100, pct, "≤ 0,5%", True),
   "Eventos de avería": (["cantidad_averiada"], lambda d,m: (C(d,m,"cantidad_averiada")>0).sum(), num, None, False),
  },
  "graf": [("pie","causa","cantidad_averiada","sum","Averías por causa"),
           ("bar","categoria","cantidad_averiada","sum","Averías por categoría"),
           ("line","fecha","cantidad_averiada","sum","Tendencia mensual de averías")],
 },
 "Ausentismo": {
  "claves": ["ausencia","incapacidad","permiso","empleado","dias_ausencia","motivo"],
  "ind": {
   "Tasa de ausentismo": (["dias_ausencia","dias_laborables"], lambda d,m: C(d,m,"dias_ausencia").sum()/C(d,m,"dias_laborables").sum()*100, pct, "≤ 3%", True),
   "Días perdidos": (["dias_ausencia"], lambda d,m: C(d,m,"dias_ausencia").sum(), num, None, True),
   "Colaboradores con ausencias": (["empleado","dias_ausencia"], lambda d,m: d.loc[C(d,m,"dias_ausencia")>0, m["empleado"]].nunique(), num, None, True),
   "Promedio de días por ausente": (["empleado","dias_ausencia"], lambda d,m: C(d,m,"dias_ausencia").sum()/max(d.loc[C(d,m,"dias_ausencia")>0, m["empleado"]].nunique(),1), lambda x: f"{x:.1f} días", None, False),
  },
  "graf": [("pie","motivo","dias_ausencia","sum","Días de ausencia por motivo"),
           ("bar","cargo","dias_ausencia","sum","Ausencias por cargo o turno"),
           ("line","fecha","dias_ausencia","sum","Tendencia mensual")],
 },
 "Cumplimiento de pedidos": {
  "claves": ["pedido","a_tiempo","completo","otif","despacho","entrega"],
  "ind": {
   "OTIF": (["completo","a_tiempo"], lambda d,m: (B(d,m,"completo")&B(d,m,"a_tiempo")).mean()*100, pct, "≥ 95%", True),
   "Pedidos completos": (["completo"], lambda d,m: B(d,m,"completo").mean()*100, pct, None, True),
   "Entregas a tiempo": (["a_tiempo"], lambda d,m: B(d,m,"a_tiempo").mean()*100, pct, None, True),
   "Errores de despacho": (["error_despacho"], lambda d,m: B(d,m,"error_despacho").mean()*100, pct, "≤ 1%", True),
   "Total de pedidos": (["pedido"], lambda d,m: C(d,m,"pedido").nunique(), num, None, False),
  },
  "graf": [("bar","ciudad","_a_tiempo_pct","mean","Entregas a tiempo por ciudad (%)"),
           ("line","fecha","_otif_pct","mean","OTIF mensual (%)"),
           ("bar","cliente","pedido","count","Pedidos por cliente")],
 },
 "Productividad de bodega": {
  "claves": ["operario","alistamiento","picking","lineas","turno","minutos"],
  "ind": {
   "Tiempo promedio de alistamiento": (["min_alistamiento"], lambda d,m: C(d,m,"min_alistamiento").mean(), lambda x: f"{x:.0f} min", None, True),
   "Pedidos por operario": (["pedido","operario"], lambda d,m: C(d,m,"pedido").nunique()/C(d,m,"operario").nunique(), num, None, True),
   "Unidades por operario": (["cantidad","operario"], lambda d,m: C(d,m,"cantidad").sum()/C(d,m,"operario").nunique(), num, None, True),
  },
  "graf": [("bar","operario","pedido","count","Pedidos por operario"),
           ("bar","operario","min_alistamiento","mean","Minutos promedio por operario"),
           ("bar","turno","cantidad","sum","Unidades por turno")],
 },
 "Transporte": {
  "claves": ["flete","vehiculo","placa","ruta","transportador","km"],
  "ind": {
   "Costo de flete por pedido": (["costo_flete"], lambda d,m: C(d,m,"costo_flete").mean(), money, None, True),
   "Costo total de fletes": (["costo_flete"], lambda d,m: C(d,m,"costo_flete").sum(), money, None, True),
   "Entregas a tiempo": (["a_tiempo"], lambda d,m: B(d,m,"a_tiempo").mean()*100, pct, "≥ 95%", True),
  },
  "graf": [("bar","transportador","costo_flete","sum","Costo por transportador"),
           ("bar","ciudad","costo_flete","mean","Flete promedio por ciudad"),
           ("line","fecha","costo_flete","sum","Costo de fletes mensual")],
 },
},
"Ventas": {
 "Desempeño comercial": {
  "claves": ["venta","factura","cliente","vendedor","ticket"],
  "ind": {
   "Ventas totales": (["valor_venta"], lambda d,m: C(d,m,"valor_venta").sum(), money, None, True),
   "Ticket promedio": (["valor_venta","factura"], lambda d,m: C(d,m,"valor_venta").sum()/C(d,m,"factura").nunique(), money, None, True),
   "Clientes activos": (["cliente"], lambda d,m: C(d,m,"cliente").nunique(), num, None, True),
   "Margen bruto": (["valor_venta","costo_venta"], lambda d,m: (1-C(d,m,"costo_venta").sum()/C(d,m,"valor_venta").sum())*100, pct, None, False),
  },
  "graf": [("bar","vendedor","valor_venta","sum","Ventas por vendedor"),
           ("line","fecha","valor_venta","sum","Ventas mensuales"),
           ("top","cliente","valor_venta","sum","Top 10 clientes")],
 },
 "Cumplimiento de metas": {
  "claves": ["meta","presupuesto","cuota"],
  "ind": {
   "Cumplimiento de meta": (["valor_venta","meta"], lambda d,m: C(d,m,"valor_venta").sum()/C(d,m,"meta").sum()*100, pct, "≥ 100%", True),
   "Brecha frente a la meta": (["valor_venta","meta"], lambda d,m: C(d,m,"meta").sum()-C(d,m,"valor_venta").sum(), money, None, True),
  },
  "graf": [("bar","vendedor","valor_venta","sum","Ventas por vendedor"),
           ("bar","vendedor","meta","sum","Meta por vendedor")],
 },
},
"Recursos Humanos": {
 "Planta de personal": {
  "claves": ["empleado","cargo","ingreso","retiro","salario","contrato"],
  "ind": {
   "Total de colaboradores": (["empleado"], lambda d,m: C(d,m,"empleado").nunique(), num, None, True),
   "Rotación de personal": (["retirado"], lambda d,m: B(d,m,"retirado").mean()*100, pct, None, True),
  },
  "graf": [("bar","cargo","empleado","count","Personas por cargo"),
           ("pie","area_empresa","empleado","count","Personas por área")],
 },
 "Ausentismo": None,   # reutiliza el de Logística
 "Capacitación": {
  "claves": ["capacitacion","curso","horas","formacion"],
  "ind": {
   "Horas de capacitación": (["horas_capacitacion"], lambda d,m: C(d,m,"horas_capacitacion").sum(), num, None, True),
   "Horas por persona": (["horas_capacitacion","empleado"], lambda d,m: C(d,m,"horas_capacitacion").sum()/C(d,m,"empleado").nunique(), lambda x: f"{x:.1f} h", None, True),
  },
  "graf": [("bar","curso","horas_capacitacion","sum","Horas por curso"),
           ("bar","cargo","horas_capacitacion","sum","Horas por cargo")],
 },
},
"Finanzas": {
 "Estado de resultados": {
  "claves": ["ingreso","gasto","utilidad","cuenta","centro_costo"],
  "ind": {
   "Ingresos": (["ingresos"], lambda d,m: C(d,m,"ingresos").sum(), money, None, True),
   "Gastos": (["gastos"], lambda d,m: C(d,m,"gastos").sum(), money, None, True),
   "Margen neto": (["ingresos","gastos"], lambda d,m: (C(d,m,"ingresos").sum()-C(d,m,"gastos").sum())/C(d,m,"ingresos").sum()*100, pct, None, True),
  },
  "graf": [("line","fecha","ingresos","sum","Ingresos mensuales"),
           ("bar","centro_costo","gastos","sum","Gastos por centro de costo")],
 },
 "Cartera": {
  "claves": ["cartera","vencida","dias_mora","saldo"],
  "ind": {
   "Cartera total": (["cartera_total"], lambda d,m: C(d,m,"cartera_total").sum(), money, None, True),
   "Cartera vencida": (["cartera_vencida","cartera_total"], lambda d,m: C(d,m,"cartera_vencida").sum()/C(d,m,"cartera_total").sum()*100, pct, "≤ 10%", True),
  },
  "graf": [("top","cliente","cartera_vencida","sum","Top 10 clientes con cartera vencida"),
           ("bar","rango_mora","cartera_total","sum","Cartera por días de mora")],
 },
},
}
CAT["Recursos Humanos"]["Ausentismo"] = CAT["Logística"]["Ausentismo"]

# Variable principal de cada subtablero para el análisis avanzado: (campo, agregación, nombre)
PRINCIPAL = {
    "Exactitud de inventario": ("_dif_abs", "sum", "Diferencia de inventario (unidades)"),
    "Averías": ("cantidad_averiada", "sum", "Unidades averiadas"),
    "Ausentismo": ("dias_ausencia", "sum", "Días de ausencia"),
    "Cumplimiento de pedidos": ("_otif_pct", "mean", "OTIF (%)"),
    "Productividad de bodega": ("min_alistamiento", "mean", "Minutos de alistamiento"),
    "Transporte": ("costo_flete", "sum", "Costo de fletes"),
    "Desempeño comercial": ("valor_venta", "sum", "Ventas"),
    "Cumplimiento de metas": ("valor_venta", "sum", "Ventas"),
    "Planta de personal": ("empleado", "count", "Colaboradores"),
    "Capacitación": ("horas_capacitacion", "sum", "Horas de capacitación"),
    "Estado de resultados": ("ingresos", "sum", "Ingresos"),
    "Cartera": ("cartera_vencida", "sum", "Cartera vencida"),
}
DIMENSIONES = ["categoria","causa","motivo","cargo","ciudad","ubicacion","operario","sku","cliente","vendedor",
               "transportador","turno","curso","centro_costo","rango_mora","area_empresa","empleado"]
ANALISIS = ["Tendencia y variación mensual", "Pareto 80/20", "Valores atípicos",
            "Pronóstico 3 meses", "Correlaciones", "Estadísticas descriptivas"]

# ---------------------------------------------------------------------------
# TEMAS DE COLOR
# ---------------------------------------------------------------------------
TEMAS = {
    "Oscuro C&M":       dict(bg="#0d0f14", card="#171a22", txt="#eeeeee", sub="#a9adb8", acc="#FF6B1A", plotly="plotly_dark",  borde="#262a35"),
    "Claro ejecutivo":  dict(bg="#f4f5f7", card="#ffffff", txt="#1a1d24", sub="#5b6170", acc="#1F6FEB", plotly="plotly_white", borde="#dde1e7"),
    "Azul corporativo": dict(bg="#0b1a2e", card="#13263f", txt="#eaf1fb", sub="#9fb3cc", acc="#2EC4FF", plotly="plotly_dark",  borde="#1f3858"),
    "Verde agro":       dict(bg="#0f1a12", card="#18261b", txt="#ecf3ea", sub="#a3b79f", acc="#6BCB3D", plotly="plotly_dark",  borde="#26392a"),
}
SECUNDARIOS = ["#FFB01A", "#2EC4B6", "#8E7CFF", "#E5484D", "#9AA4B2", "#F2C94C"]
OK_COLOR, ALERTA_COLOR = "#2EB67D", "#E5484D"

# ---------------------------------------------------------------------------
# IA DEL CLIENTE
# ---------------------------------------------------------------------------
SIN_IA = "Sin IA (reglas automáticas)"
PROVEEDORES = {
    "Claude (Anthropic)":       dict(secreto="ANTHROPIC_API_KEY",    modelo="claude-sonnet-5-5"),
    "ChatGPT (OpenAI)":         dict(secreto="OPENAI_API_KEY",       modelo="gpt-4.1-mini"),
    "Gemini (Google)":          dict(secreto="GEMINI_API_KEY",       modelo="gemini-2.5-flash"),
    "Azure OpenAI (Microsoft)": dict(secreto="AZURE_OPENAI_API_KEY", modelo=""),
    "Microsoft Copilot (Copilot Studio)": dict(secreto="COPILOT_DIRECTLINE_SECRET", modelo="copilot"),
}

def secreto(nombre):
    try: return st.secrets.get(nombre, "")
    except Exception: return ""

def _post(url, headers, body, timeout=90):
    req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"),
                                 headers={"Content-Type": "application/json", **headers}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"El proveedor respondió {e.code}: {e.read().decode('utf-8', 'ignore')[:300]}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"No se pudo conectar con el proveedor: {e.reason}")

def _get(url, headers, timeout=30):
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))

def copilot(prompt, secreto_dl, dominio):
    """Conversa con un agente de Copilot Studio por el canal Direct Line y devuelve su respuesta."""
    import time
    base, h = dominio.rstrip("/") + "/v3/directline", {"Authorization": f"Bearer {secreto_dl}"}
    conv = _post(f"{base}/conversations", h, {})["conversationId"]
    envio = _post(f"{base}/conversations/{conv}/activities", h,
                  {"type": "message", "from": {"id": "tablero-cm-labs"}, "text": prompt, "locale": "es-CO"})["id"]
    textos, quietos, limite = [], 0, time.time() + 90
    while time.time() < limite:
        time.sleep(1.5)
        acts = _get(f"{base}/conversations/{conv}/activities", h).get("activities", [])
        nuevos = [a.get("text", "") for a in acts if a.get("type") == "message"
                  and a.get("from", {}).get("id") != "tablero-cm-labs" and a.get("replyToId") == envio and a.get("text")]
        if nuevos and len(nuevos) == len(textos): quietos += 1
        else: quietos = 0
        textos = nuevos
        if textos and quietos >= 2: break
    if not textos: raise RuntimeError("El agente de Copilot no respondió a tiempo. Revisa que tenga respuestas generativas activadas.")
    return "\n\n".join(textos)

def ia(prompt, max_tokens=1200):
    """Envía el prompt a la IA que configuró el cliente y devuelve el texto."""
    p, k, mod = IA["prov"], IA["key"], IA["modelo"]
    if not IA["activa"]: raise RuntimeError("No hay una IA conectada.")
    msg = [{"role": "user", "content": prompt}]
    if p.startswith("Claude"):
        r = _post("https://api.anthropic.com/v1/messages", {"x-api-key": k, "anthropic-version": "2023-06-01"},
                  {"model": mod, "max_tokens": max_tokens, "messages": msg})
        return "".join(b.get("text", "") for b in r.get("content", []))
    if p.startswith("ChatGPT"):
        r = _post("https://api.openai.com/v1/chat/completions", {"Authorization": f"Bearer {k}"}, {"model": mod, "messages": msg})
        return r["choices"][0]["message"]["content"]
    if p.startswith("Azure"):
        url = f"{IA['endpoint'].rstrip('/')}/openai/deployments/{mod}/chat/completions?api-version={IA['version']}"
        r = _post(url, {"api-key": k}, {"messages": msg})
        return r["choices"][0]["message"]["content"]
    if p.startswith("Microsoft Copilot"):
        return copilot(prompt, k, IA["endpoint"] or "https://directline.botframework.com")
    if p.startswith("Gemini"):
        r = _post(f"https://generativelanguage.googleapis.com/v1beta/models/{mod}:generateContent",
                  {"x-goog-api-key": k}, {"contents": [{"parts": [{"text": prompt}]}]})
        return "".join(x.get("text", "") for x in r["candidates"][0]["content"]["parts"])
    raise RuntimeError("Proveedor no reconocido.")

# ---------------------------------------------------------------------------
# RECONOCIMIENTO DE TIPOLOGÍA
# ---------------------------------------------------------------------------
def campos_de(sub):
    s = set()
    for req, *_ in sub["ind"].values(): s |= set(req)
    for _, x, y, *_ in sub["graf"]: s |= {x, y}
    return sorted(c for c in s if not c.startswith("_") and c not in ("conteo",))

def norm(s): return re.sub(r"[^a-z0-9]", "", str(s).lower().replace("ñ","n"))

def detectar_por_palabras(df):
    cols = " ".join(norm(c) for c in df.columns)
    puntajes = {(a, s): sum(norm(k) in cols for k in sub["claves"]) for a, subs in CAT.items() for s, sub in subs.items()}
    return max(puntajes, key=puntajes.get)

def mapa_por_nombre(df, campos):
    """Asignación de respaldo sin IA: primero coincidencias exactas, luego parciales,
    sin reutilizar columnas y sin tomar una columna que corresponde exactamente a otro campo."""
    m, usadas = {}, set()
    exactas = {norm(c) for c in campos}
    for c in campos:
        for col in df.columns:
            if norm(col) == norm(c): m[c] = col; usadas.add(col); break
    for c in sorted([c for c in campos if c not in m], key=lambda x: -len(x)):
        cand = [col for col in df.columns if col not in usadas and norm(col) not in exactas and norm(c) in norm(col)]
        if cand:
            col = min(cand, key=lambda x: len(str(x))); m[c] = col; usadas.add(col)
    return m

def json_de(t): return json.loads(t[t.find("{"): t.rfind("}")+1])

def reconocer(df):
    if IA["activa"]:
        opciones = {a: list(s.keys()) for a, s in CAT.items()}
        p = (f"Columnas y muestra de una tabla empresarial:\n{df.head(5).to_csv(index=False)}\n"
             f"Opciones de tablero: {json.dumps(opciones, ensure_ascii=False)}\n"
             "Elige el área y subtablero que mejor corresponde a estos datos. "
             'Responde SOLO JSON: {"area": "...", "subtablero": "...", "razon": "frase corta"}')
        try:
            r = json_de(ia(p, 300))
            if r["area"] in CAT and r["subtablero"] in CAT[r["area"]]: return r["area"], r["subtablero"], r.get("razon", "")
        except Exception: pass
    a, s = detectar_por_palabras(df); return a, s, "Detectado por nombres de columnas."

def mapear(df, area, subt):
    campos = campos_de(CAT[area][subt])
    if IA["activa"]:
        p = (f"Tabla:\n{df.head(5).to_csv(index=False)}\nCampos estándar del tablero '{subt}': {campos}. "
             "Asigna cada campo a la columna real equivalente (o una columna que sirva como dimensión parecida, "
             "por ejemplo 'bodega' para 'ubicacion'). Responde SOLO JSON {campo: columna}; omite los que no existan.")
        try:
            m = json_de(ia(p))
            m = {k: v for k, v in m.items() if v in df.columns and k in campos}
            if m: return m
        except Exception: pass
    return mapa_por_nombre(df, campos)

# columnas derivadas que usan algunos gráficos
def derivadas(df, m):
    d = df.copy()
    if {"stock_sistema","stock_fisico"} <= m.keys():
        d["_dif_abs"] = (d[m["stock_sistema"]] - d[m["stock_fisico"]]).abs(); m["_dif_abs"] = "_dif_abs"
    if "a_tiempo" in m:
        d["_a_tiempo_pct"] = B(d, m, "a_tiempo")*100; m["_a_tiempo_pct"] = "_a_tiempo_pct"
    if {"completo","a_tiempo"} <= m.keys():
        d["_otif_pct"] = (B(d,m,"completo") & B(d,m,"a_tiempo"))*100; m["_otif_pct"] = "_otif_pct"
    return d, m

# ---------------------------------------------------------------------------
# GRÁFICOS
# ---------------------------------------------------------------------------
def estilo(fig, titulo=None):
    fig.update_layout(template=T["plotly"], paper_bgcolor=T["card"], plot_bgcolor=T["card"],
                      font=dict(color=T["txt"]), colorway=PALETA, margin=dict(l=20, r=20, t=55, b=20),
                      legend=dict(orientation="h", y=-0.18), title_font=dict(size=16))
    if titulo: fig.update_layout(title=titulo)
    return fig

def grafico(d, m, tipo, x, y, agg, titulo):
    if x not in m or (agg != "count" and y not in m): return None
    X = m[x]
    if tipo == "line":
        g = d.assign(_mes=pd.to_datetime(d[X], errors="coerce").dt.to_period("M").astype(str))
        s = g.groupby("_mes")[m[y]].agg(agg).reset_index(); s.columns = ["Mes", titulo]
        return estilo(px.line(s, x="Mes", y=titulo, markers=True, color_discrete_sequence=[T["acc"]]), titulo)
    s = (d.groupby(X).size() if agg == "count" else d.groupby(X)[m[y]].agg(agg)).reset_index(); s.columns = [x, "valor"]
    if tipo == "top": s = s.sort_values("valor", ascending=False).head(10).iloc[::-1]
    elif s.shape[0] > 25: s = s.sort_values("valor", ascending=False).head(25)
    if tipo == "pie":
        return estilo(px.pie(s, names=x, values="valor", hole=.5, color_discrete_sequence=PALETA), titulo)
    fig = px.bar(s, x="valor" if tipo=="top" else x, y=x if tipo=="top" else "valor",
                 orientation="h" if tipo=="top" else "v", color_discrete_sequence=[T["acc"]])
    return estilo(fig, titulo)

# ---------------------------------------------------------------------------
# ANÁLISIS AVANZADO  (cada función devuelve un panel o None si no hay datos)
# ---------------------------------------------------------------------------
def serie_mensual(d, m, campo, agg):
    if "fecha" not in m or (agg != "count" and campo not in m): return None
    g = d.assign(_mes=pd.to_datetime(d[m["fecha"]], errors="coerce").dt.to_period("M")).dropna(subset=["_mes"])
    s = g.groupby("_mes").size() if agg == "count" else g.groupby("_mes")[m[campo]].agg(agg)
    return s.sort_index() if len(s) else None

def a_tendencia(d, m, pr):
    campo, agg, etq = pr
    s = serie_mensual(d, m, campo, agg)
    if s is None or len(s) < 2: return None
    meses, var = s.index.astype(str), s.pct_change() * 100
    fig = go.Figure()
    fig.add_bar(x=meses, y=s.values, name=etq, marker_color=T["acc"])
    fig.add_scatter(x=meses, y=var.values, name="Variación %", yaxis="y2", mode="lines+markers",
                    line=dict(color=SECUNDARIOS[1], width=2))
    fig.update_layout(yaxis2=dict(overlaying="y", side="right", showgrid=False, ticksuffix="%"))
    ult = var.iloc[-1]
    h = "" if pd.isna(ult) else f"{etq}: en {meses[-1]} {'subió' if ult > 0 else 'bajó'} {abs(ult):.1f}% frente al mes anterior."
    return dict(fig=estilo(fig, f"Tendencia mensual · {etq}"), nota=h, hallazgo=h)

def a_pareto(d, m, pr, dim):
    campo, agg, etq = pr
    if not dim or dim not in m or agg not in ("sum", "count") or (agg == "sum" and campo not in m): return None
    s = d.groupby(m[dim]).size() if agg == "count" else d.groupby(m[dim])[m[campo]].sum()
    s = s[s > 0].sort_values(ascending=False)
    if len(s) < 2: return None
    acum = s.cumsum() / s.sum() * 100
    n80 = int((acum < 80).sum()) + 1
    top = s.head(20); x = top.index.astype(str)
    fig = go.Figure()
    fig.add_bar(x=x, y=top.values, name=etq, marker_color=T["acc"])
    fig.add_scatter(x=x, y=acum.head(20).values, name="% acumulado", yaxis="y2", mode="lines+markers",
                    line=dict(color=SECUNDARIOS[1], width=2))
    fig.add_hline(y=80, yref="y2", line_dash="dot", line_color=T["sub"])
    fig.update_layout(yaxis2=dict(overlaying="y", side="right", range=[0, 105], showgrid=False, ticksuffix="%"))
    h = f"{n80} de {len(s)} valores de '{dim}' concentran el 80% de {etq.lower()}: {', '.join(map(str, s.index[:min(n80, 3)]))}."
    return dict(fig=estilo(fig, f"Pareto 80/20 · {etq} por {dim}"), nota=h, hallazgo=h)

def a_atipicos(d, m, pr):
    campo, agg, etq = pr
    if agg == "count" or campo not in m: return None
    v = pd.to_numeric(d[m[campo]], errors="coerce")
    if v.dropna().nunique() <= 2: return None
    q1, q3 = v.quantile(.25), v.quantile(.75); iqr = q3 - q1
    bajo, alto = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    fuera = d[(v < bajo) | (v > alto)]
    fig = go.Figure(go.Box(y=v.dropna(), name=etq, marker_color=T["acc"], boxpoints="outliers"))
    h = (f"{len(fuera)} registros atípicos en {etq.lower()} (fuera del rango {bajo:,.1f} a {alto:,.1f})."
         if len(fuera) else f"No hay valores atípicos en {etq.lower()}.")
    tabla = fuera.sort_values(m[campo], ascending=False).head(15) if len(fuera) else None
    return dict(fig=estilo(fig, f"Valores atípicos · {etq}"), nota=h, hallazgo=h, tabla=tabla, tabla_titulo="Ver registros atípicos")

def a_pronostico(d, m, pr):
    campo, agg, etq = pr
    s = serie_mensual(d, m, campo, agg)
    if s is None or len(s) < 3: return None
    y = s.values.astype(float); x = np.arange(len(y))
    a, b = np.polyfit(x, y, 1)
    fx = np.arange(len(y), len(y) + 3); fy = a * fx + b
    meses = list(s.index.astype(str)); futuros = [str(s.index[-1] + i) for i in (1, 2, 3)]
    fig = go.Figure()
    fig.add_scatter(x=meses, y=y, name="Real", mode="lines+markers", line=dict(color=T["acc"], width=3))
    fig.add_scatter(x=[meses[-1]] + futuros, y=[y[-1]] + list(fy), name="Pronóstico", mode="lines+markers",
                    line=dict(color=SECUNDARIOS[1], dash="dash", width=2))
    cifra = f"{fy[-1]:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
    h = f"Si sigue la tendencia, {etq.lower()} {'subiría' if a > 0 else 'bajaría'} a cerca de {cifra} en {futuros[-1]}."
    return dict(fig=estilo(fig, f"Pronóstico 3 meses · {etq}"), nota=h + " (Tendencia lineal simple, úsese como referencia.)", hallazgo=h)

def a_correlaciones(d, m):
    cols = []
    for c in dict.fromkeys(m.values()):
        if c in d.columns and pd.api.types.is_numeric_dtype(d[c]) and d[c].nunique() > 2: cols.append(c)
    if len(cols) < 2: return None
    c = d[cols].corr().round(2)
    fig = px.imshow(c, text_auto=True, zmin=-1, zmax=1, color_continuous_scale=[SECUNDARIOS[1], T["card"], T["acc"]])
    par = c.where(~np.eye(len(c), dtype=bool)).abs().stack()
    h = ""
    if len(par):
        (x1, x2), r = par.idxmax(), c.loc[par.idxmax()]
        h = f"La relación más fuerte es entre '{x1}' y '{x2}' (correlación {r:+.2f})."
    return dict(fig=estilo(fig, "Correlaciones entre variables"), nota=h, hallazgo=h)

def a_descriptivas(d, m):
    cols = [c for c in dict.fromkeys(m.values()) if c in d.columns and pd.api.types.is_numeric_dtype(d[c])]
    if not cols: return None
    t = d[cols].describe().T.round(2)
    t = t.rename(columns={"count": "Registros", "mean": "Promedio", "std": "Desv. estándar", "min": "Mínimo",
                          "25%": "P25", "50%": "Mediana", "75%": "P75", "max": "Máximo"})
    return dict(titulo="Estadísticas descriptivas", tabla=t.reset_index().rename(columns={"index": "Variable"}), tabla_visible=True)

def cumple(valor, meta):
    if not meta or valor is None: return None
    r = re.match(r"([≥≤])\s*([\d\.,]+)", meta)
    if not r: return None
    obj = float(r.group(2).replace(",", "."))
    return valor >= obj if r.group(1) == "≥" else valor <= obj

# ---------------------------------------------------------------------------
# DATOS DE EJEMPLO (para demostraciones sin archivo)
# ---------------------------------------------------------------------------
@st.cache_data
def ejemplo(nombre):
    rng = np.random.default_rng(7)
    cats = ["Abarrotes", "Aseo", "Bebidas", "Lácteos", "Snacks"]
    if nombre == "Exactitud de inventario":
        n = 80
        df = pd.DataFrame({"sku": [f"SKU-{i:03d}" for i in range(n)], "categoria": rng.choice(cats, n),
                           "ubicacion": rng.choice(["Bodega A", "Bodega B", "Muelle", "Mezanine"], n),
                           "stock_sistema": rng.integers(50, 500, n), "costo_unit": rng.integers(2000, 40000, n)})
        df["stock_fisico"] = df.stock_sistema - np.where(rng.random(n) < .3, rng.integers(-5, 15, n), 0)
        return df
    if nombre == "Averías":
        n = 150
        return pd.DataFrame({"fecha": pd.Timestamp("2026-04-01") + pd.to_timedelta(rng.integers(0, 180, n), "D"),
                             "sku": [f"SKU-{i:03d}" for i in rng.integers(0, 80, n)], "categoria": rng.choice(cats, n),
                             "causa": rng.choice(["Golpe en montacargas", "Humedad", "Mal estibado", "Vencimiento", "Transporte"], n, p=[.35, .15, .25, .1, .15]),
                             "cantidad": rng.integers(100, 600, n), "cantidad_averiada": rng.integers(0, 6, n),
                             "costo_unit": rng.integers(2000, 40000, n)})
    filas, cargos = [], ["Operario", "Montacarguista", "Auxiliar de inventario", "Despachador"]
    for mes in pd.date_range("2026-04-01", periods=6, freq="MS"):
        for i in range(25):
            filas.append([mes, f"Colaborador {i+1}", cargos[i % 4], rng.choice(["Incapacidad", "Permiso", "Calamidad", "Injustificada"]),
                          int(rng.choice([0, 0, 0, 1, 2, 3], p=[.5, .15, .1, .12, .08, .05])), 24])
    return pd.DataFrame(filas, columns=["fecha", "empleado", "cargo", "motivo", "dias_ausencia", "dias_laborables"])

# ---------------------------------------------------------------------------
# BARRA LATERAL (los contenedores fijan el orden visual)
# ---------------------------------------------------------------------------
sb_datos, sb_tablero, sb_ind, sb_avz, sb_filtros, sb_tema, sb_ia = [st.sidebar.container() for _ in range(7)]

with sb_tema:
    st.header("6. Apariencia")
    tema = st.selectbox("Tema", list(TEMAS.keys()))
    T = dict(TEMAS[tema])
    T["acc"] = st.color_picker("Color de acento", T["acc"], key=f"acc_{tema}")
    PALETA = [T["acc"]] + SECUNDARIOS

with sb_ia:
    st.header("7. IA para insights")
    opciones = [SIN_IA] + list(PROVEEDORES)
    por_defecto = next((i + 1 for i, p in enumerate(PROVEEDORES) if secreto(PROVEEDORES[p]["secreto"])), 0)
    prov = st.selectbox("Conectar con", opciones, index=por_defecto,
                        help="Usa la IA que ya tiene tu empresa. Sin IA, los insights se generan con reglas automáticas.")
    IA = dict(prov=prov, key="", modelo="", endpoint="", version="2024-10-21", activa=False)
    if prov != SIN_IA:
        cfg, del_servidor = PROVEEDORES[prov], secreto(PROVEEDORES[prov]["secreto"])
        clave = st.text_input("Secreto de Direct Line" if prov.startswith("Microsoft Copilot") else "Clave API", type="password",
                              placeholder="Usando la clave del servidor" if del_servidor else "Pega aquí la clave de tu empresa")
        IA["key"] = clave or del_servidor
        if prov.startswith("Azure"):
            IA["endpoint"] = st.text_input("Endpoint", value=secreto("AZURE_OPENAI_ENDPOINT"), placeholder="https://mirecurso.openai.azure.com")
            IA["modelo"] = st.text_input("Nombre del deployment", value=secreto("AZURE_OPENAI_DEPLOYMENT"))
            IA["version"] = st.text_input("Versión de la API", value=IA["version"])
        elif prov.startswith("Microsoft Copilot"):
            IA["modelo"] = "copilot"
            IA["endpoint"] = st.text_input("Dominio de Direct Line", value=secreto("COPILOT_DIRECTLINE_DOMAIN") or "https://directline.botframework.com",
                                           help="Cámbialo solo si tu agente está en una región específica (por ejemplo, Europa).")
            st.caption("Usa un agente de Copilot Studio con respuestas generativas activadas y el canal Direct Line "
                       "publicado (Configuración › Seguridad › Canales web).")
        else:
            IA["modelo"] = st.text_input("Modelo", value=cfg["modelo"], help="Escribe el modelo que tenga habilitado tu cuenta.")
        IA["activa"] = bool(IA["key"] and IA["modelo"] and (IA["endpoint"] or not prov.startswith("Azure")))
        if st.button("Probar conexión", disabled=not IA["activa"]):
            try: ia("Responde únicamente con la palabra OK.", 20); st.success("Conexión correcta.")
            except Exception as e: st.error(str(e))
        st.caption("La clave solo vive en esta sesión y no se guarda. A la IA se envían 5 filas de muestra "
                   "para reconocer las columnas y los indicadores calculados para los insights.")

st.markdown(f"""<style>
.stApp{{background:{T['bg']}}}
[data-testid="stHeader"]{{background:transparent}}
[data-testid="stMain"] h1,[data-testid="stMain"] h2,[data-testid="stMain"] h3,
[data-testid="stMain"] .stMarkdown p,[data-testid="stMain"] .stMarkdown li{{color:{T['txt']}}}
[data-testid="stMain"] [data-testid="stCaptionContainer"]{{color:{T['sub']}}}
[data-testid="stMain"] [data-testid="stVerticalBlockBorderWrapper"]{{background:{T['card']};border-color:{T['borde']}!important;border-radius:10px}}
.kick{{letter-spacing:4px;color:{T['acc']};font-size:12px;font-weight:700;margin:6px 0 10px}}
.kpi{{background:{T['card']};padding:16px;border-left:4px solid {T['acc']};border-radius:8px;margin-bottom:10px}}
.kpi b{{font-size:26px;display:block;color:{T['txt']}}} .kpi span{{color:{T['sub']}}}
.kpi small{{display:block;margin-top:4px;color:{T['sub']}}} .kpi small.ok{{color:{OK_COLOR}}} .kpi small.mal{{color:{ALERTA_COLOR}}}
.resumen{{color:{T['txt']};line-height:1.9}} .resumen b{{color:{T['acc']}}}
.resumen h4{{color:{T['txt']};margin:0 0 8px}}
</style>""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# FUNCIONES DE PANTALLA
# ---------------------------------------------------------------------------
GENERAL = "Tablero general del área"

def calcular(sub, d, m, nombres):
    res = {}
    for n in nombres:
        req, calc, fmt, meta, _ = sub["ind"][n]
        try: val = float(calc(d, m)); res[n] = dict(valor=fmt(val), meta=meta, cumple=cumple(val, meta))
        except Exception: res[n] = dict(valor="—", meta=meta, cumple=None)
    return res

def pintar_kpis(res, etiquetas=None):
    nombres = list(res)
    for i in range(0, len(nombres), 4):
        fila = nombres[i:i+4]
        for col, n in zip(st.columns(len(fila)), fila):
            r, ok = res[n], res[n]["cumple"]
            meta = "" if not r["meta"] else f'<small class="{"ok" if ok else "mal" if ok is False else ""}">{"✓" if ok else "⚠" if ok is False else ""} Meta {r["meta"]}</small>'
            col.markdown(f'<div class="kpi"><b>{r["valor"]}</b><span>{(etiquetas or {}).get(n, n)}</span>{meta}</div>', unsafe_allow_html=True)

def avanzados(d, m, subt, lista, dim):
    pr, paneles, hallazgos, sin = PRINCIPAL.get(subt), [], [], []
    if not pr: return paneles, hallazgos, list(lista)
    f = {"Tendencia y variación mensual": lambda: a_tendencia(d, m, pr),
         "Pareto 80/20": lambda: a_pareto(d, m, pr, dim),
         "Valores atípicos": lambda: a_atipicos(d, m, pr),
         "Pronóstico 3 meses": lambda: a_pronostico(d, m, pr),
         "Correlaciones": lambda: a_correlaciones(d, m),
         "Estadísticas descriptivas": lambda: a_descriptivas(d, m)}
    for nombre in lista:
        try: p = f[nombre]()
        except Exception: p = None
        if p:
            paneles.append(p)
            if p.get("hallazgo"): hallazgos.append(p["hallazgo"])
        else: sin.append(nombre)
    return paneles, hallazgos, sin

def panel_resumen(df, d, sub, mapa, n_ind, n_avz, filtros):
    cols_ok = [c for c in campos_de(sub) if c in mapa]
    usadas = [mapa[c] for c in cols_ok if mapa[c] in df.columns]
    completo = (1 - df[usadas].isna().mean().mean()) * 100 if usadas else 0
    periodo = ""
    if "fecha" in mapa:
        f = pd.to_datetime(df[mapa["fecha"]], errors="coerce").dropna()
        if len(f): periodo = f"<br>Período: <b>{f.min():%d/%m/%Y}</b> a <b>{f.max():%d/%m/%Y}</b>"
    return dict(html=f"""<div class="resumen"><h4>Resumen de los datos</h4>
        Registros analizados: <b>{num(len(d))}</b> de {num(len(df))}{periodo}<br>
        Campos reconocidos: <b>{len(cols_ok)}</b> de {len(campos_de(sub))}<br>
        Datos completos: <b>{completo:.1f}%</b><br>
        Indicadores visibles: <b>{n_ind}</b> · Análisis avanzados: <b>{n_avz}</b><br>
        Filtros: <b>{'; '.join(filtros) if filtros else 'ninguno'}</b></div>""")

def pintar(p):
    with st.container(border=True):
        if p.get("titulo"): st.markdown(f"**{p['titulo']}**")
        if p.get("fig") is not None: st.plotly_chart(p["fig"], width="stretch")
        if p.get("html"): st.markdown(p["html"], unsafe_allow_html=True)
        if p.get("tabla") is not None:
            if p.get("tabla_visible"): st.dataframe(p["tabla"], width="stretch", hide_index=True)
            else:
                with st.expander(p.get("tabla_titulo", "Ver datos")): st.dataframe(p["tabla"], width="stretch", hide_index=True)
        if p.get("nota"): st.caption(p["nota"])

def pintar_paneles(paneles, relleno):
    """Gráficos en pares y tablas a lo ancho. Si queda un hueco, se llena con el resumen."""
    anchos = [p for p in paneles if p.get("tabla_visible")]
    pares = [p for p in paneles if not p.get("tabla_visible")]
    if len(pares) % 2 == 1 or not pares: pares.append(relleno())
    for i in range(0, len(pares), 2):
        c1, c2 = st.columns(2)
        with c1: pintar(pares[i])
        if i + 1 < len(pares):
            with c2: pintar(pares[i + 1])
    for p in anchos: pintar(p)

def insights(titulo, resultados, hallazgos):
    st.markdown('<div class="kick">INSIGHTS</div>', unsafe_allow_html=True)
    def por_reglas():
        out = []
        for n, r in resultados.items():
            if r["cumple"] is True: out.append(f"- ✓ **{n}** ({r['valor']}) cumple la meta {r['meta']}.")
            elif r["cumple"] is False: out.append(f"- ⚠ **{n}** ({r['valor']}) está por fuera de la meta {r['meta']}; conviene revisarlo.")
        out += [f"- {h}" for h in hallazgos]
        return "\n".join(out) or "- Activa indicadores con meta o análisis avanzados para obtener hallazgos."
    etiqueta = f"Generar insights con {IA['prov'].split(' (')[0]}" if IA["activa"] else "Generar insights"
    if st.button(etiqueta):
        if IA["activa"]:
            datos = {n: {"valor": r["valor"], "meta": r["meta"]} for n, r in resultados.items()}
            p = (f"Eres analista de {titulo} para una pyme colombiana. Indicadores: {json.dumps(datos, ensure_ascii=False)}. "
                 f"Hallazgos del análisis avanzado: {json.dumps(hallazgos, ensure_ascii=False)}. "
                 "Da 4 a 6 insights breves y accionables en español, en viñetas, comparando con las metas y estándares del sector.")
            with st.spinner("Analizando..."):
                try: st.markdown(ia(p, 1200)); return
                except Exception as e: st.warning(f"No se pudo usar la IA ({e}). Estos son los insights automáticos:")
        st.markdown(por_reglas())
        if not IA["activa"]:
            st.caption("Conecta la IA de tu empresa en la sección 7 de la barra lateral para insights redactados por IA.")

# ---------------------------------------------------------------------------
# INTERFAZ
# ---------------------------------------------------------------------------
st.markdown('<div class="kick">C&M LABS · TABLEROS AUTOMÁTICOS</div>', unsafe_allow_html=True)
st.title("Sube tus datos y te armamos el tablero")

with sb_datos:
    st.header("1. Datos")
    archivos = st.file_uploader("Excel o CSV (puedes subir varios)", type=["xlsx", "csv"], accept_multiple_files=True)
    demo = st.selectbox("…o prueba con datos de ejemplo",
                        ["(ninguno)", "Logística completa (3 tablas)", "Averías", "Exactitud de inventario", "Ausentismo"])

@st.cache_data(show_spinner=False)
def leer(nombre, contenido):
    import io
    if nombre.lower().endswith(".csv"): return {"datos": pd.read_csv(io.BytesIO(contenido))}
    return pd.read_excel(io.BytesIO(contenido), sheet_name=None)

fuentes = {}
if archivos:
    for a in archivos:
        hojas = leer(a.name, a.getvalue())
        for h, x in hojas.items():
            if len(x) and len(x.columns): fuentes[a.name if len(hojas) == 1 else f"{a.name} › {h}"] = x
elif demo.startswith("Logística"):
    fuentes = {f"Ejemplo · {n}": ejemplo(n) for n in ["Exactitud de inventario", "Averías", "Ausentismo"]}
elif demo != "(ninguno)":
    fuentes = {f"Ejemplo · {demo}": ejemplo(demo)}
if not fuentes:
    st.info("Sube uno o varios archivos, o escoge datos de ejemplo. La app reconoce si son datos de inventario, averías, "
            "ausentismo, pedidos, ventas, cartera, etc. Si subes varias tablas de la misma área, puedes ver el tablero general. "
            "Abajo a la izquierda puedes cambiar el tema y conectar tu IA.")
    st.stop()

# Reconocer cada tabla una sola vez
recon, CLAVE = st.session_state.setdefault("recon", {}), {}
for f, x in fuentes.items():
    CLAVE[f] = f"{f}|{len(x)}|{'|'.join(map(str, x.columns))}"
    if CLAVE[f] not in recon:
        with st.spinner(f"Reconociendo {f}..."): recon[CLAVE[f]] = reconocer(x)
R = {f: recon[CLAVE[f]] for f in fuentes}

def fuentes_de(area, subt):
    return [f for f in fuentes if CAT[R[f][0]][R[f][1]] is CAT[area][subt]]

def mapa_de(f, area, subt):
    mapas = st.session_state.setdefault("mapas", {})
    k = f"{CLAVE[f]}::{subt}"
    if k not in mapas:
        with st.spinner("Asignando columnas..."): mapas[k] = mapear(fuentes[f], area, subt)
    return mapas[k]

# Sugerencia: el área más frecuente; si hay varias tablas de esa área, el tablero general
areas_det = [R[f][0] for f in fuentes]
sug_area = max(set(areas_det), key=areas_det.count)
subs_det = {R[f][1] for f in fuentes if R[f][0] == sug_area}
sug_sub = GENERAL if len(subs_det) > 1 else next(iter(subs_det))

with sb_tablero:
    st.header("2. Tablero")
    areas = list(CAT.keys())
    area = st.selectbox("Área", areas, index=areas.index(sug_area))
    subs = [GENERAL] + list(CAT[area].keys())
    subt = st.selectbox("Subtablero", subs, index=subs.index(sug_sub) if sug_sub in subs else 0)
    st.caption("Reconocido: " + "; ".join(f"{f} → {R[f][1]}" for f in fuentes))

# =========================== TABLERO GENERAL ===============================
if subt == GENERAL:
    con_datos = [s for s in CAT[area] if fuentes_de(area, s)]
    with sb_ind:
        st.header("3. Subtableros incluidos")
        incluidos = st.multiselect("Incluir en el tablero general", list(CAT[area]), default=con_datos, key=f"gen_{area}")
        sin = [s for s in CAT[area] if s not in con_datos]
        if sin: st.caption("Sube datos para incluir: " + ", ".join(sin))
    with sb_avz:
        st.header("4. Análisis avanzado")
        avz = st.multiselect("Incluir en cada subtablero", ANALISIS, default=ANALISIS[:1], key=f"avz_gen_{area}")
    with sb_filtros:
        st.header("5. Filtros")
        st.caption("El tablero general muestra todos los datos. Para filtrar, abre cada subtablero por separado.")

    st.markdown(f'<div class="kick">{area.upper()} › TABLERO GENERAL</div>', unsafe_allow_html=True)
    secciones, clave_area, todos, hall_todos = [], {}, {}, []
    for s in incluidos:
        fs = fuentes_de(area, s)
        if not fs: continue
        sub_s, f = CAT[area][s], fs[0]
        mapa_s = mapa_de(f, area, s)
        d_s, m_s = derivadas(fuentes[f], dict(mapa_s))
        nombres = [n for n, (req, *_r) in sub_s["ind"].items() if all(c in mapa_s for c in req) and sub_s["ind"][n][4]]
        res = calcular(sub_s, d_s, m_s, nombres)
        if res:
            principal = next((n for n in res if res[n]["meta"]), next(iter(res)))
            clave_area[f"{s}|{principal}"] = res[principal]
        todos.update({f"{s} · {n}": r for n, r in res.items()})
        secciones.append((s, f, sub_s, fuentes[f], d_s, m_s, mapa_s, res))

    if not secciones:
        st.warning("No hay datos para los subtableros escogidos. Sube las tablas del área o cambia la selección.")
        st.stop()

    st.markdown("### Indicadores clave del área")
    pintar_kpis(clave_area, {k: f"{k.split('|')[0]} · {k.split('|')[1]}" for k in clave_area})

    for s, f, sub_s, df_s, d_s, m_s, mapa_s, res in secciones:
        st.divider()
        st.markdown(f'<div class="kick">{s.upper()}</div>', unsafe_allow_html=True)
        st.caption(f"Fuente: {f}")
        pintar_kpis(res)
        figs = [dict(fig=g) for g in (grafico(d_s, m_s, *x) for x in sub_s["graf"]) if g is not None][:2]
        dims_s = [c for c in DIMENSIONES if c in mapa_s and c not in ("sku", "empleado")]
        pan, hall, _ = avanzados(d_s, m_s, s, avz, dims_s[0] if dims_s else None)
        hall_todos += [f"{s}: {h}" for h in hall]
        pintar_paneles(figs + pan, lambda: panel_resumen(df_s, d_s, sub_s, mapa_s, len(res), len(pan), []))

    st.divider()
    insights(f"{area} (tablero general: {', '.join(x[0] for x in secciones)})", todos, hall_todos)
    st.stop()

# =========================== SUBTABLERO ====================================
sub = CAT[area][subt]
ops = list(fuentes)
cands = fuentes_de(area, subt)
with sb_tablero:
    fuente = st.selectbox("Tabla de datos", ops, index=ops.index(cands[0]) if cands else 0, key=f"fu_{area}_{subt}") if len(ops) > 1 else ops[0]
    if not cands and len(ops) > 1: st.caption("Ninguna tabla se reconoció como este subtablero; revisa la asignación de columnas.")
df = fuentes[fuente]
mapa = mapa_de(fuente, area, subt)

with sb_tablero:
    with st.expander("Ajustar columnas"):
        for campo in campos_de(sub):
            ops_c = ["(ninguna)"] + list(df.columns)
            sel = st.selectbox(campo, ops_c, index=ops_c.index(mapa[campo]) if mapa.get(campo) in ops_c else 0, key=f"{CLAVE[fuente]}_{subt}_{campo}")
            if sel == "(ninguna)": mapa.pop(campo, None)
            else: mapa[campo] = sel

with sb_ind:
    st.header("3. Indicadores")
    disp = [n for n, (req, *_) in sub["ind"].items() if all(c in mapa for c in req)]
    elegidos = st.multiselect("Muestra", disp, default=[n for n in disp if sub["ind"][n][4]], key=f"ind_{subt}")
    falta = [n for n in sub["ind"] if n not in disp]
    if falta: st.caption("Sin datos suficientes para: " + ", ".join(falta))

dims = [c for c in DIMENSIONES if c in mapa]
with sb_avz:
    st.header("4. Análisis avanzado")
    avz = st.multiselect("Incluir", ANALISIS, default=ANALISIS[:2], key=f"avz_{subt}")
    dim_pareto = st.selectbox("Pareto por", dims, key=f"par_{subt}") if "Pareto 80/20" in avz and dims else None

with sb_filtros:
    st.header("5. Filtros")
    d, filtros = df.copy(), []
    for campo in ["categoria","ciudad","ubicacion","operario","cargo","vendedor","cliente","causa","motivo"]:
        if campo in mapa and d[mapa[campo]].nunique() <= 50:
            v = st.multiselect(campo.capitalize(), sorted(d[mapa[campo]].dropna().astype(str).unique()), key=f"f_{subt}_{campo}")
            if v: d = d[d[mapa[campo]].astype(str).isin(v)]; filtros.append(f"{campo}: {', '.join(v)}")

d, m = derivadas(d, dict(mapa))
st.markdown(f'<div class="kick">{area.upper()} › {subt.upper()}</div>', unsafe_allow_html=True)

resultados = calcular(sub, d, m, elegidos)
pintar_kpis(resultados)
paneles = [dict(fig=g) for g in (grafico(d, m, *x) for x in sub["graf"]) if g is not None]
pan, hallazgos, sin_datos = avanzados(d, m, subt, avz, dim_pareto)
pintar_paneles(paneles + pan, lambda: panel_resumen(df, d, sub, mapa, len(elegidos), len(pan), filtros))
if sin_datos: st.caption("Sin datos suficientes para: " + ", ".join(sin_datos) + ". Revisa las columnas asignadas (por ejemplo, una fecha).")
insights(f"{area}, tablero '{subt}'", resultados, hallazgos)
