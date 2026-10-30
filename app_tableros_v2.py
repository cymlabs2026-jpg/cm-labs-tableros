# C&M LABS - Tableros Automáticos v2 (áreas y subtableros)
# pip install streamlit pandas openpyxl anthropic plotly
# streamlit run app_tableros_v2.py   (requiere ANTHROPIC_API_KEY)
import streamlit as st, pandas as pd, plotly.express as px, json, re

ACC = "#FF6B1A"
st.set_page_config(page_title="C&M LABS · Tableros", layout="wide")
st.markdown(f"""<style>.stApp{{background:#0d0f14;color:#eee}}
.kick{{letter-spacing:4px;color:{ACC};font-size:12px;font-weight:700}}
.kpi{{background:#171a22;padding:16px;border-left:4px solid {ACC};border-radius:8px;margin-bottom:10px}}
.kpi b{{font-size:26px;display:block}} .kpi span{{color:#aaa}} .kpi small{{color:#777;display:block}}</style>""",
unsafe_allow_html=True)

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

def campos_de(sub):
    s = set()
    for req, *_ in sub["ind"].values(): s |= set(req)
    for _, x, y, *_ in sub["graf"]: s |= {x, y}
    return sorted(c for c in s if not c.startswith("_") and c not in ("conteo",))

# ---------------------------------------------------------------------------
# RECONOCIMIENTO DE TIPOLOGÍA
# ---------------------------------------------------------------------------
def claude(prompt, max_tokens=1500):
    from anthropic import Anthropic
    r = Anthropic().messages.create(model="claude-sonnet-4-5", max_tokens=max_tokens,
                                    messages=[{"role": "user", "content": prompt}])
    return r.content[0].text

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

def reconocer(df):
    opciones = {a: list(s.keys()) for a, s in CAT.items()}
    p = (f"Columnas y muestra de una tabla empresarial:\n{df.head(5).to_csv(index=False)}\n"
         f"Opciones de tablero: {json.dumps(opciones, ensure_ascii=False)}\n"
         "Elige el área y subtablero que mejor corresponde a estos datos. "
         'Responde SOLO JSON: {"area": "...", "subtablero": "...", "razon": "frase corta"}')
    try:
        t = claude(p, 300); r = json.loads(t[t.find("{"): t.rfind("}")+1])
        if r["area"] in CAT and r["subtablero"] in CAT[r["area"]]: return r["area"], r["subtablero"], r.get("razon","")
    except Exception: pass
    a, s = detectar_por_palabras(df); return a, s, "Detectado por nombres de columnas."

def mapear(df, area, subt):
    campos = campos_de(CAT[area][subt])
    p = (f"Tabla:\n{df.head(5).to_csv(index=False)}\nCampos estándar del tablero '{subt}': {campos}. "
         "Asigna cada campo a la columna real equivalente (o una columna que sirva como dimensión parecida, "
         "por ejemplo 'bodega' para 'ubicacion'). Responde SOLO JSON {campo: columna}; omite los que no existan.")
    try:
        t = claude(p); m = json.loads(t[t.find("{"): t.rfind("}")+1])
        return {k: v for k, v in m.items() if v in df.columns}
    except Exception:
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

def grafico(d, m, tipo, x, y, agg, titulo):
    if x not in m or (y != "pedido" and agg != "count" and y not in m): return None
    X = m[x]
    if tipo == "line":
        g = d.assign(_mes=pd.to_datetime(d[X], errors="coerce").dt.to_period("M").astype(str))
        s = g.groupby("_mes")[m[y]].agg(agg).reset_index(); s.columns = ["Mes", titulo]
        return px.line(s, x="Mes", y=titulo, markers=True, title=titulo, template="plotly_dark", color_discrete_sequence=[ACC])
    s = (d.groupby(X).size() if agg == "count" else d.groupby(X)[m[y]].agg(agg)).reset_index(); s.columns = [x, "valor"]
    if tipo == "top": s = s.sort_values("valor", ascending=False).head(10)
    elif s.shape[0] > 25: s = s.sort_values("valor", ascending=False).head(25)
    if tipo == "pie":
        return px.pie(s, names=x, values="valor", hole=.5, title=titulo, template="plotly_dark",
                      color_discrete_sequence=[ACC,"#FFB01A","#2EC4B6","#8E7CFF","#E5E5E5","#555"])
    fig = px.bar(s, x="valor" if tipo=="top" else x, y=x if tipo=="top" else "valor", orientation="h" if tipo=="top" else "v",
                 title=titulo, template="plotly_dark", color_discrete_sequence=[ACC])
    return fig

# ---------------------------------------------------------------------------
# INTERFAZ
# ---------------------------------------------------------------------------
st.markdown('<div class="kick">C&M LABS · TABLEROS AUTOMÁTICOS</div>', unsafe_allow_html=True)
st.title("Sube tus datos y te armamos el tablero")

with st.sidebar:
    st.header("1. Datos")
    archivo = st.file_uploader("Excel o CSV", type=["xlsx","csv"])

if not archivo:
    st.info("Sube un archivo. La app reconoce si son datos de inventario, averías, ausentismo, pedidos, ventas, cartera, etc.")
    st.stop()

if archivo.name.endswith(".csv"):
    hojas = {"datos": pd.read_csv(archivo)}
else:
    hojas = pd.read_excel(archivo, sheet_name=None)
with st.sidebar:
    hoja = st.selectbox("Hoja", list(hojas.keys())) if len(hojas) > 1 else list(hojas.keys())[0]
df = hojas[hoja]

clave = f"{archivo.name}|{hoja}"
if st.session_state.get("clave") != clave:
    with st.spinner("Claude está reconociendo el tipo de datos..."):
        a, s, razon = reconocer(df)
    st.session_state.update(clave=clave, area=a, subt=s, razon=razon, mapas={})

with st.sidebar:
    st.header("2. Tablero")
    areas = list(CAT.keys())
    area = st.selectbox("Área", areas, index=areas.index(st.session_state.area))
    subs = list(CAT[area].keys())
    subt = st.selectbox("Subtablero", subs, index=subs.index(st.session_state.subt) if st.session_state.subt in subs else 0)
    st.caption(f"Sugerido: {st.session_state.area} › {st.session_state.subt}. {st.session_state.razon}")

sub = CAT[area][subt]
if subt not in st.session_state.mapas:
    with st.spinner("Asignando columnas..."):
        st.session_state.mapas[subt] = mapear(df, area, subt)
mapa = st.session_state.mapas[subt]

with st.sidebar:
    with st.expander("Ajustar columnas"):
        for campo in campos_de(sub):
            ops = ["(ninguna)"] + list(df.columns)
            sel = st.selectbox(campo, ops, index=ops.index(mapa[campo]) if mapa.get(campo) in ops else 0, key=f"{subt}_{campo}")
            if sel == "(ninguna)": mapa.pop(campo, None)
            else: mapa[campo] = sel
    st.header("3. Indicadores")
    disp = [n for n, (req, *_ ) in sub["ind"].items() if all(c in mapa for c in req)]
    elegidos = st.multiselect("Muestra", disp, default=[n for n in disp if sub["ind"][n][4]], key=f"ind_{subt}")
    falta = [n for n in sub["ind"] if n not in disp]
    if falta: st.caption("Sin datos suficientes para: " + ", ".join(falta))
    # filtros
    st.header("4. Filtros")
    d = df.copy()
    for campo in ["categoria","ciudad","ubicacion","operario","cargo","vendedor","cliente"]:
        if campo in mapa and d[mapa[campo]].nunique() <= 50:
            v = st.multiselect(campo.capitalize(), sorted(d[mapa[campo]].dropna().astype(str).unique()), key=f"f_{subt}_{campo}")
            if v: d = d[d[mapa[campo]].astype(str).isin(v)]

d, m = derivadas(d, dict(mapa))
st.markdown(f'<div class="kick">{area.upper()} › {subt.upper()}</div>', unsafe_allow_html=True)

resultados = {}
cols = st.columns(4)
for k, n in enumerate(elegidos):
    req, calc, fmt, meta, _ = sub["ind"][n]
    try: txt = fmt(calc(d, m)); resultados[n] = txt
    except Exception: txt = "—"
    cols[k % 4].markdown(f'<div class="kpi"><b>{txt}</b><span>{n}</span>{f"<small>Meta {meta}</small>" if meta else ""}</div>', unsafe_allow_html=True)

figs = [f for f in (grafico(d, m, *g) for g in sub["graf"]) if f is not None]
for i in range(0, len(figs), 2):
    c1, c2 = st.columns(2)
    c1.plotly_chart(figs[i], use_container_width=True)
    if i + 1 < len(figs): c2.plotly_chart(figs[i+1], use_container_width=True)
if not figs: st.caption("Asigna columnas como categoría, causa, fecha u operario para ver gráficos.")

st.markdown('<div class="kick">INSIGHTS GENERADOS POR CLAUDE</div>', unsafe_allow_html=True)
if st.button("Generar insights") and resultados:
    p = (f"Eres analista de {area}, tablero '{subt}', para una pyme colombiana. Indicadores: "
         f"{json.dumps(resultados, ensure_ascii=False)}. Da 4 insights breves y accionables en español, en viñetas, "
         "comparando con estándares del sector y metas típicas.")
    with st.spinner("Analizando..."):
        try: st.markdown(claude(p, 800))
        except Exception as e: st.error(f"No se pudieron generar insights: {e}")
