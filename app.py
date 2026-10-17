import streamlit as st
import sqlite3
import os
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Control Financiero", page_icon="💰", layout="centered")

RUTA_CARPETA = os.path.dirname(os.path.abspath(__file__))
RUTA_DB = os.path.join(RUTA_CARPETA, "control_cuentas_dinamico.db")
RUTA_EXCEL = os.path.join(RUTA_CARPETA, "reporte_finanzas.xlsx")
def inicializar_db():
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS mes_activo (
            item TEXT PRIMARY KEY,
            monto REAL NOT NULL,
            tipo TEXT NOT NULL
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS historico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mes_ano TEXT NOT NULL,
            item TEXT NOT NULL,
            monto REAL NOT NULL,
            tipo TEXT NOT NULL DEFAULT 'Gasto'
        )
    """)
    cursor.execute("SELECT COUNT(*) FROM mes_activo WHERE tipo = 'Ingreso'")
    if cursor.fetchone() == 0:
        cursor.execute("INSERT OR IGNORE INTO mes_activo (item, monto, tipo) VALUES ('Ingresos Principales', 0.0, 'Ingreso')")
    conn.commit()
    conn.close()

def actualizar_monto_db(item, nuevo_monto):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("UPDATE mes_activo SET monto = ? WHERE item = ?", (nuevo_monto, item))
    conn.commit()
    conn.close()

def insertar_nuevo_item_db(item, tipo="Gasto"):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO mes_activo (item, monto, tipo) VALUES (?, 0.0, ?)", (item, tipo))
        conn.commit()
        exito = True
    except sqlite3.IntegrityError:
        exito = False
    conn.close()
    return exito
def cambiar_nombre_ingreso_db(viejo_nombre, nuevo_nombre):
    if viejo_nombre == nuevo_nombre:
        return True
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE mes_activo SET item = ? WHERE item = ? AND tipo = 'Ingreso'", (nuevo_nombre, viejo_nombre))
        conn.commit()
        exito = True
    except sqlite3.IntegrityError:
        exito = False
    conn.close()
    return exito

def eliminar_gasto_db(item):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM mes_activo WHERE item = ? AND tipo = 'Gasto'", (item,))
    conn.commit()
    conn.close()

def obtener_todos_los_saldos():
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT item, monto, tipo FROM mes_activo")
    datos = cursor.fetchall()
    conn.close()
    return datos

def guardar_mes_en_historico(mes_ano, datos_saldos):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM historico WHERE mes_ano = ?", (mes_ano,))
    for item, monto, tipo in datos_saldos:
        cursor.execute("INSERT INTO historico (mes_ano, item, monto, tipo) VALUES (?, ?, ?, ?)", (mes_ano, item, monto, tipo))
    conn.commit()
    conn.close()
    exportar_datos_a_excel()

def exportar_datos_a_excel():
    conn = sqlite3.connect(RUTA_DB)
    df_historico = pd.read_sql_query("SELECT mes_ano, item, monto, tipo FROM historico", conn)
    conn.close()
    if df_historico.empty:
        return
    df_desglose = df_historico.copy()
    df_desglose.columns = ["Mes-Año", "Concepto", "Monto (CLP)", "Tipo Flujo"]
    df_gastos = df_historico[df_historico['tipo'] == 'Gasto']
    if not df_gastos.empty:
        df_comparativa = df_gastos.groupby('mes_ano')['monto'].sum().reset_index()
        df_comparativa.columns = ["Periodo Contable", "Total Gastos Mensuales (CLP)"]
        df_comparativa = df_comparativa.sort_values(by="Total Gastos Mensuales (CLP)", ascending=False)
    else:
        df_comparativa = pd.DataFrame(columns=["Periodo Contable", "Total Gastos Mensuales (CLP)"])
    with pd.ExcelWriter(RUTA_EXCEL, engine='openpyxl') as writer:
        df_desglose.to_excel(writer, sheet_name="Historial Desglosado", index=False)
        df_comparativa.to_excel(writer, sheet_name="Comparativa de Gastos", index=False)

def obtener_meses_historicos():
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT mes_ano FROM historico ORDER BY id DESC")
    meses = [fila for fila in cursor.fetchall()]
    conn.close()
    return meses

def formatear_clp(monto):
    valor_absoluto = int(round(abs(monto)))
    formato_numero = f"${valor_absoluto:,}".replace(",", ".")
    return formato_numero if monto >= 0 else f"-{formato_numero}"

def limpiar_y_convertir_monto(texto_monto):
    texto_limpio = texto_monto.replace("$", "").replace(".", "").strip()
    if not texto_limpio:
        return 0.0
    try:
        return float(texto_limpio)
    except ValueError:
        return 0.0
inicializar_db()

st.title("📱 Control Financiero Móvil")

st.subheader("👤 Bienvenido al sistema")
usuario_app = st.text_input("Ingrese su nombre para comenzar:", value="", placeholder="Escriba su nombre aquí...")

st.markdown("---")
st.subheader("📅 Periodo Contable Activo")
col_mes, col_ano = st.columns(2)
with col_mes:
    mes_sel = st.selectbox("Mes:", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"], index=8)
with col_ano:
    ano_sel = st.selectbox("Año:", ["2025", "2026", "2027", "2028"], index=1)
periodo_completo = f"{mes_sel} {ano_sel}"

st.markdown("### 📥 Ingresos Consolidados")
saldos_db = obtener_todos_los_saldos()
nombre_ingreso_inicial = next((row for row in saldos_db if row == "Ingreso"), "Ingresos Principales")
monto_ingreso_inicial = next((row for row in saldos_db if row == "Ingreso"), 0.0)

nuevo_nombre_ingreso = st.text_input("Concepto de Ingreso:", value=nombre_ingreso_inicial)
if nuevo_nombre_ingreso != nombre_ingreso_inicial and nuevo_nombre_ingreso.strip():
    cambiar_nombre_ingreso_db(nombre_ingreso_inicial, nuevo_nombre_ingreso.strip())
    st.rerun()

monto_inicial_formateado = f"{int(monto_ingreso_inicial):,}".replace(",", ".") if monto_ingreso_inicial > 0 else ""
ingreso_input_text = st.text_input(f"Monto {nuevo_nombre_ingreso} (CLP $):", value=monto_inicial_formateado, placeholder="Ej: 1.500.000")

ingreso_procesado = limpiar_y_convertir_monto(ingreso_input_text)
texto_formateado_ingreso = f"{int(ingreso_procesado):,}".replace(",", ".") if ingreso_procesado > 0 else ""

if texto_formateado_ingreso != ingreso_input_text and ingreso_input_text != "":
    actualizar_monto_db(nuevo_nombre_ingreso, ingreso_procesado)
    st.rerun()
elif ingreso_input_text == "" and monto_ingreso_inicial != 0.0:
    actualizar_monto_db(nuevo_nombre_ingreso, 0.0)
    st.rerun()

monto_ingreso = ingreso_procesado
st.markdown("---")
st.markdown("### 💸 Controles Operativos de Gastos")

# Inicializar memoria temporal si no existe
if "gasto_texto_tmp" not in st.session_state:
    st.session_state["gasto_texto_tmp"] = ""

col_add, col_del = st.columns(2)

with col_add:
    # Capturar el texto vinculándolo a la memoria persistente de la sesión
    nuevo_gasto = st.text_input("Añadir nuevo concepto de gasto:", value=st.session_state["gasto_texto_tmp"], placeholder="Ej: Supermercado")
    if st.button("➕ Añadir Gasto", use_container_width=True):
        if nuevo_gasto.strip():
            if nuevo_gasto.strip() != nuevo_nombre_ingreso:
                if insertar_nuevo_item_db(nuevo_gasto.strip(), "Gasto"):
                    st.session_state["gasto_texto_tmp"] = ""  # Limpiar casilla
                    st.success(f"'{nuevo_gasto.strip()}' agregado exitosamente.")
                    st.rerun()
                else:
                    st.error("Este ítem de gasto ya se encuentra registrado.")
            else:
                st.error("El nombre coincide con la casilla activa de ingresos.")
        else:
            st.warning("Por favor, escriba un nombre de concepto válido.")

with col_del:
    gastos_existentes = [row for row in saldos_db if row == "Gasto"]
    gasto_a_eliminar = st.selectbox("Eliminar un gasto:", ["Seleccione..."] + gastos_existentes)
    if st.button("🗑️ Eliminar Gasto", use_container_width=True):
        if gasto_a_eliminar != "Seleccione...":
            eliminar_gasto_db(gasto_a_eliminar)
            st.success(f"'{gasto_a_eliminar}' eliminado exitosamente.")
            st.rerun()

st.markdown("#### Valores del Periodo")
total_ingresos = monto_ingreso
total_gastos = 0.0
gastos_filtrados = {}

for item, valor, tipo in saldos_db:
    if tipo == "Gasto":
        valor_inicial_formateado = f"{int(valor):,}".replace(",", ".") if valor > 0 else ""
        gasto_input_text = st.text_input(f"Monto {item} (CLP $):", value=valor_inicial_formateado, key=f"gasto_input_{item}", placeholder="Ej: 45.000")
        
        gasto_procesado = limpiar_y_convertir_monto(gasto_input_text)
        texto_formateado_gasto = f"{int(gasto_procesado):,}".replace(",", ".") if gasto_procesado > 0 else ""
        
        if texto_formateado_gasto != gasto_input_text and gasto_input_text != "":
            actualizar_monto_db(item, gasto_procesado)
            st.rerun()
        elif gasto_input_text == "" and valor != 0.0:
            actualizar_monto_db(item, 0.0)
            st.rerun()
            
        total_gastos += gasto_procesado
        if gasto_procesado > 0:
            gastos_filtrados[item] = gasto_procesado

saldo_disponible = total_ingresos - total_gastos
st.markdown("---")
st.markdown("### 📊 Cuadro de Mando de Analíticas")
st.metric("Total Ingresos", formatear_clp(total_ingresos))
st.metric("Total Gastos", formatear_clp(total_gastos))
st.metric("Balance Neto", formatear_clp(saldo_disponible))

fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(6, 8))
fig.patch.set_facecolor('#282C34')

ax1.set_facecolor("#282C34")
ax2.set_facecolor("#282C34")

barras = ax1.bar(["Ingresos", "Gastos"], [total_ingresos, total_gastos], color=["#4A90E2", "#E74C3C"], width=0.4)
ax1.set_title("Balance Financiero", color="#E1E4E6", fontsize=12, fontweight="bold")
ax1.tick_params(colors="#A0A5AD", labelsize=10)
for bar in barras:
    yval = bar.get_height()
    ax1.text(bar.get_x() + bar.get_width()/2, yval + 1, formatear_clp(yval), ha='center', va='bottom', color='#E1E4E6', fontsize=10, fontweight="bold")

if gastos_filtrados:
    ax2.pie(gastos_filtrados.values(), labels=gastos_filtrados.keys(), autopct='%1.1f%%', startangle=90, textprops=dict(color="#E1E4E6", fontsize=10), colors=['#4A90E2', '#357ABD', '#225E9B', '#1B4780', '#5C9CE6'])
    ax2.set_title("Distribución de Egresos", color="#E1E4E6", fontsize=12, fontweight="bold")
else:
    ax2.text(0.5, 0.5, 'Sin transacciones de egreso', ha='center', va='center', color='#8A9099', fontsize=10)

st.pyplot(fig)

st.markdown("---")
if st.button("🗄️ Archivar Periodo y Actualizar Excel", use_container_width=True):
    if total_ingresos + total_gastos > 0:
        guardar_mes_en_historico(periodo_completo, obtener_todos_los_saldos())
        conn = sqlite3.connect(RUTA_DB)
        cursor = conn.cursor()
        cursor.execute("UPDATE mes_activo SET monto = 0.0")
        conn.commit()
        conn.close()
        st.success(f"¡Periodo contable {periodo_completo} indexado con éxito!")
        st.rerun()
    else:
        st.warning("No hay transacciones para archivar.")

st.subheader("📁 Periodos Consolidados")
meses_hist = obtener_meses_historicos()
mes_hist_sel = st.selectbox("Selecciona un mes para auditar:", ["Seleccione..."] + [m for m in meses_hist])

if mes_hist_sel != "Seleccione...":
    conn = sqlite3.connect(RUTA_DB)
    df_resumen = pd.read_sql_query("SELECT item as 'Concepto', monto as 'Monto' FROM historico WHERE mes_ano = ?", conn, params=(mes_hist_sel,))
    conn.close()
    df_resumen['Monto'] = df_resumen['Monto'].apply(formatear_clp)
    st.dataframe(df_resumen, use_container_width=True)

