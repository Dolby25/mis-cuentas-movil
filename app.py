import streamlit as st
import sqlite3
import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

st.set_page_config(page_title="Control Financiero PRO", page_icon="💰", layout="centered")

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
    # Nueva tabla estratégica para las metas de ahorro con tope de día
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS metas_ahorro (
            concepto TEXT PRIMARY KEY,
            monto_meta REAL NOT NULL,
            monto_actual REAL NOT NULL,
            fecha_tope TEXT NOT NULL
        )
    """)
    # Nueva tabla para resguardar el PIN de seguridad del usuario
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS seguridad (
            llave TEXT PRIMARY KEY,
            valor TEXT NOT NULL
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
    return f"${int(round(abs(monto))):,}".replace(",", ".") if monto >= 0 else f"-${int(round(abs(monto))):,}".replace(",", ".")

def limpiar_y_convertir_monto(texto_monto):
    texto_limpio = texto_monto.replace("$", "").replace(".", "").strip()
    return float(texto_limpio) if texto_limpio.isdigit() else 0.0

# Funciones de validación del PIN
def obtener_pin_seguridad():
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT valor FROM seguridad WHERE llave = 'user_pin'")
    fila = cursor.fetchone()
    conn.close()
    return fila[0] if fila else None

def guardar_nuevo_pin(nuevo_pin):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO seguridad (llave, valor) VALUES ('user_pin', ?)", (nuevo_pin,))
    conn.commit()
    conn.close()
inicializar_db()

# Manejo del estado de autenticación en la sesión actual para recordar al usuario
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

pin_registrado = obtener_pin_seguridad()

if not st.session_state["autenticado"]:
    st.title("🔒 Acceso Protegido")
    
    if pin_registrado is None:
        st.subheader("🆕 Crear PIN de Seguridad")
        st.info("No se ha detectado un código PIN activo en este teléfono. Por favor, crea uno para proteger tus datos financieros.")
        nuevo_pin_input = st.text_input("Define tu PIN de 4 dígitos:", type="password", max_chars=4)
        confirmar_pin_input = st.text_input("Confirma tu PIN de 4 dígitos:", type="password", max_chars=4)
        
        if st.button("💾 Guardar PIN y Entrar", use_container_width=True):
            if len(nuevo_pin_input) == 4 and nuevo_pin_input.isdigit():
                if nuevo_pin_input == confirmar_pin_input:
                    guardar_nuevo_pin(nuevo_pin_input)
                    st.session_state["autenticado"] = True
                    st.success("¡PIN configurado exitosamente!")
                    st.rerun()
                else:
                    st.error("Los códigos PIN ingresados no coinciden.")
            else:
                st.error("El PIN debe ser estrictamente numérico de 4 dígitos.")
        st.stop()
        
    else:
        st.subheader("🔢 Ingrese su PIN de Acceso")
        pin_ingresado = st.text_input("Código PIN:", type="password", max_chars=4, placeholder="••••")
        
        if st.button("🔓 Desbloquear Aplicación", use_container_width=True) or (len(pin_ingresado) == 4 and pin_ingresado == pin_registrado):
            if pin_ingresado == pin_registrado:
                st.session_state["autenticado"] = True
                st.rerun()
            elif len(pin_ingresado) == 4:
                st.error("PIN incorrecto. Inténtelo nuevamente.")
        st.stop()
# --- INTERFAZ PRINCIPAL DESBLOQUEADA ---
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

if "gasto_texto_tmp" not in st.session_state:
    st.session_state["gasto_texto_tmp"] = ""

def ejecutar_alta_gasto():
    texto_alta = st.session_state["campo_alta_gasto"].strip()
    if texto_alta:
        if texto_alta != nuevo_nombre_ingreso:
            if insertar_nuevo_item_db(texto_alta, "Gasto"):
                st.session_state["gasto_texto_tmp"] = ""
                st.toast(f"'{texto_alta}' agregado.", icon="➕")
            else:
                st.error("Este ítem de gasto ya se encuentra registrado.")
        else:
            st.error("El nombre coincide con la casilla activa de ingresos.")

def ejecutar_baja_gasto():
    target_baja = st.session_state["selector_baja_gasto"]
    if target_baja != "Seleccione...":
        eliminar_gasto_db(target_baja)
        st.toast(f"'{target_baja}' eliminado.", icon="🗑️")

col_add, col_del = st.columns(2)
with col_add:
    st.text_input("Añadir nuevo concepto de gasto:", value=st.session_state["gasto_texto_tmp"], key="campo_alta_gasto", placeholder="Ej: Supermercado", on_change=ejecutar_alta_gasto)
    st.button("➕ Añadir Gasto", use_container_width=True)
with col_del:
    gastos_existentes = [row for row in saldos_db if row == "Gasto"]
    st.selectbox("Eliminar un gasto:", ["Seleccione..."] + gastos_existentes, key="selector_baja_gasto", on_change=ejecutar_baja_gasto)
    st.button("🗑️ Eliminar Gasto", use_container_width=True)

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
# --- NUEVA SECCIÓN: METAS DE AHORRO CON TOPE DE DÍA ---
st.markdown("---")
st.markdown("### 🎯 Planificador de Metas de Ahorro")

# Funciones internas para la gestión de metas
def ejecutar_alta_meta(concepto, meta, actual, fecha):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO metas_ahorro VALUES (?, ?, ?, ?)", (concepto, meta, actual, fecha))
    conn.commit()
    conn.close()

with st.expander("➕ Configurar nueva meta o actualizar de ahorro", expanded=False):
    concepto_meta = st.text_input("¿Para qué estás ahorrando?:", placeholder="Ej: Vacaciones o Pie de Auto")
    monto_meta_text = st.text_input("Monto Total Objetivo (CLP $):", placeholder="Ej: 2.000.000")
    monto_actual_text = st.text_input("¿Cuánto dinero tienes ahorrado ya? (CLP $):", value="0", placeholder="Ej: 400.000")
    
    # Campo de fecha táctil solicitado (Tope de día)
    fecha_tope_date = st.date_input("Tope de día para cumplir la meta:", value=datetime.now())
    
    if st.button("🎯 Establecer Meta de Ahorro", use_container_width=True):
        val_meta = limpiar_y_convertir_monto(monto_meta_text)
        val_actual = limpiar_y_convertir_monto(monto_actual_text)
        if concepto_meta.strip() and val_meta > 0:
            ejecutar_alta_meta(concepto_meta.strip(), val_meta, val_actual, fecha_tope_date.strftime("%Y-%m-%d"))
            st.success("¡Meta guardada exitosamente!")
            st.rerun()
        else:
            st.error("Ingresa un concepto válido y un monto objetivo mayor a cero.")

# Renderizar metas activas guardadas en base de datos
conn = sqlite3.connect(RUTA_DB)
df_metas = pd.read_sql_query("SELECT * FROM metas_ahorro", conn)
conn.close()

if not df_metas.empty:
    st.markdown("#### Tus Objetivos Activos")
    for _, fila in df_metas.iterrows():
        concepto = fila["concepto"]
        m_meta = fila["monto_meta"]
        m_actual = fila["monto_actual"]
        f_tope_str = fila["fecha_tope"]
        
        # Cálculos de tiempo restante y ratios financieros
        f_tope = datetime.strptime(f_tope_str, "%Y-%m-%d").date()
        dias_restantes = (f_tope - datetime.now().date()).days
        
        faltante = m_meta - m_actual
        progreso_pct = min(1.0, m_actual / m_meta) if m_meta > 0 else 0.0
        
        st.markdown(f"**🎯 {concepto}**")
        st.progress(progreso_pct)
        
        col1, col2 = st.columns(2)
        with col1:
            st.write(f"Llevas: **{formatear_clp(m_actual)}** de {formatear_clp(m_meta)}")
            st.write(f"Fecha Límite: **{f_tope.strftime('%d-%m-%Y')}**")
        with col2:
            if dias_restantes > 0:
                st.write(f"Días restantes: **{dias_restantes} días**")
                cuota_diaria = faltante / dias_restantes if faltante > 0 else 0.0
                st.write(f"Debes ahorrar: **{formatear_clp(cuota_diaria)} / día**")
            elif dias_restantes == 0:
                st.warning("⚠️ ¡Hoy es el día tope para cumplir este objetivo!")
            else:
                st.error(f"❌ El plazo expiró hace {abs(dias_restantes)} días.")
                
        if st.button(f"🗑️ Eliminar Meta: {concepto}", key=f"del_meta_{concepto}"):
            conn = sqlite3.connect(RUTA_DB)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM metas_ahorro WHERE concepto = ?", (concepto,))
            conn.commit()
            conn.close()
            st.rerun()

# --- ANALÍTICAS Y HISTORIAL ---
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

