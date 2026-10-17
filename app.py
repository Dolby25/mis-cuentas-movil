import streamlit as st
import sqlite3
import os
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="Control Financiero Vinculado", page_icon="💰", layout="centered")

RUTA_CARPETA = os.path.dirname(os.path.abspath(__file__))
RUTA_DB = os.path.join(RUTA_CARPETA, "control_cuentas_dinamico.db")
RUTA_EXCEL = os.path.join(RUTA_CARPETA, "reporte_finanzas.xlsx")

def inicializar_db():
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    # Tabla de credenciales de usuario
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            identificador TEXT PRIMARY KEY,
            clave TEXT NOT NULL
        )
    """)
    # Tabla de saldos modificada para incluir al usuario propietario
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS mes_activo (
            usuario_id TEXT,
            item TEXT,
            monto REAL NOT NULL,
            tipo TEXT NOT NULL,
            PRIMARY KEY (usuario_id, item)
        )
    """)
    # Tabla histórica modificada para incluir al usuario propietario
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS historico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario_id TEXT NOT NULL,
            mes_ano TEXT NOT NULL,
            item TEXT NOT NULL,
            monto REAL NOT NULL,
            tipo TEXT NOT NULL DEFAULT 'Gasto'
        )
    """)
    conn.commit()
    conn.close()

def registrar_usuario_db(identificador, clave):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO usuarios (identificador, clave) VALUES (?, ?)", (identificador, clave))
        # Crear de forma automática su casilla de ingresos iniciales
        cursor.execute("INSERT INTO mes_activo (usuario_id, item, monto, tipo) VALUES (?, 'Ingresos Principales', 0.0, 'Ingreso')", (identificador,))
        conn.commit()
        exito = True
    except sqlite3.IntegrityError:
        exito = False
    conn.close()
    return exito

def verificar_usuario_db(identificador, clave):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT cl_id FROM (SELECT clave as cl_id FROM usuarios WHERE identificador = ?)", (identificador,))
    resultado = cursor.fetchone()
    conn.close()
    if resultado and resultado[0] == clave:
        return True
    return False

def actualizar_monto_db(usuario_id, item, nuevo_monto):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("UPDATE mes_activo SET monto = ? WHERE usuario_id = ? AND item = ?", (nuevo_monto, usuario_id, item))
    conn.commit()
    conn.close()

def insertar_nuevo_item_db(usuario_id, item, tipo="Gasto"):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO mes_activo (usuario_id, item, monto, tipo) VALUES (?, ?, 0.0, ?)", (usuario_id, item, tipo))
        conn.commit()
        exito = True
    except sqlite3.IntegrityError:
        exito = False
    conn.close()
    return exito

def eliminar_item_db(usuario_id, item, tipo):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM mes_activo WHERE usuario_id = ? AND item = ? AND tipo = ?", (usuario_id, item, tipo))
    conn.commit()
    conn.close()

def obtener_todos_los_saldos(usuario_id):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT item, monto, tipo FROM mes_activo WHERE usuario_id = ?", (usuario_id,))
    datos = cursor.fetchall()
    conn.close()
    return datos

def guardar_mes_en_historico(usuario_id, mes_ano, datos_saldos):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM historico WHERE usuario_id = ? AND mes_ano = ?", (usuario_id, mes_ano))
    for item, monto, tipo in datos_saldos:
        cursor.execute("INSERT INTO historico (usuario_id, mes_ano, item, monto, tipo) VALUES (?, ?, ?, ?, ?)", (usuario_id, mes_ano, item, monto, tipo))
    conn.commit()
    conn.close()
    exportar_datos_a_excel(usuario_id)

def exportar_datos_a_excel(usuario_id):
    conn = sqlite3.connect(RUTA_DB)
    df_historico = pd.read_sql_query("SELECT mes_ano, item, monto, tipo FROM historico WHERE usuario_id = ?", conn, params=(usuario_id,))
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

def obtener_meses_historicos(usuario_id):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT mes_ano FROM historico WHERE usuario_id = ? ORDER BY id DESC", (usuario_id,))
    meses = [fila for fila in cursor.fetchall()]
    conn.close()
    return meses

def formatear_clp(monto):
    valor_absoluto = int(round(abs(monto)))
    formato_numero = f"${valor_absoluto:,}".replace(",", ".")
    return formato_numero if monto >= 0 else f"-{formato_numero}"

inicializar_db()
st.title("📱 Sistema Financiero Personalizado")

# Inicialización de estados de sesión
if "sesion_activa" not in st.session_state:
    st.session_state["sesion_activa"] = False
if "cuenta_usuario" not in st.session_state:
    st.session_state["cuenta_usuario"] = ""

# Formulario de Acceso Seguro
if not st.session_state["sesion_activa"]:
    pestana_login, pestana_registro = st.tabs(["🔒 Iniciar Sesión", "📝 Crear Cuenta"])
    
    with pestana_login:
        st.subheader("Acceso a tu Cuenta")
        usuario_input = st.text_input("Correo o Teléfono (ej: +56912345678):", key="login_user", placeholder="usuario@correo.com o celular")
        clave_input = st.text_input("Contraseña:", type="password", key="login_pass")
        
        if st.button("🔓 Entrar", use_container_width=True):
            if usuario_input.strip() and clave_input.strip():
                if verificar_usuario_db(usuario_input.strip(), clave_input.strip()):
                    st.session_state["sesion_activa"] = True
                    st.session_state["cuenta_usuario"] = usuario_input.strip()
                    st.toast("Sesión iniciada con éxito.", icon="✅")
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas o usuario no registrado.")
            else:
                st.warning("Por favor rellene todos los campos.")
                
    with pestana_registro:
        st.subheader("Enlaza una Nueva Cuenta")
        nuevo_usuario = st.text_input("Registrar Correo o Teléfono:", key="reg_user", placeholder="ejemplo@correo.com")
        nueva_clave = st.text_input("Definir Contraseña:", type="password", key="reg_pass")
        
        if st.button("🚀 Crear y Vincular Cuenta", use_container_width=True):
            if nuevo_usuario.strip() and nueva_clave.strip():
                if registrar_usuario_db(nuevo_usuario.strip(), nueva_clave.strip()):
                    st.success("¡Cuenta enlazada correctamente! Ya puedes iniciar sesión desde la otra pestaña.")
                else:
                    st.error("Ese correo o teléfono ya se encuentra vinculado a una cuenta activa.")
            else:
                st.warning("Por favor ingrese datos válidos.")
    st.stop()

# --- A PARTIR DE AQUÍ TODO SE FILTRA POR EL USUARIO LOGUEADO ---
usuario_logueado = st.session_state["cuenta_usuario"]
st.info(f"🟢 Cuenta activa vinculada: **{usuario_logueado}**")
st.subheader("📅 Periodo Contable Activo")
col_mes, col_ano = st.columns(2)
with col_mes:
    mes_sel = st.selectbox("Mes:", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"], index=8)
with col_ano:
    ano_sel = st.selectbox("Año:", ["2025", "2026", "2027", "2028"], index=1)
periodo_completo = f"{mes_sel} {ano_sel}"

st.markdown("---")
st.markdown("### 📥 Ingresos Consolidados")

if "nuevo_ingreso_tmp" not in st.session_state:
    st.session_state["nuevo_ingreso_tmp"] = ""

def ejecutar_alta_ingreso():
    texto_alta = st.session_state["campo_alta_ingreso"].strip()
    if texto_alta:
        if insertar_nuevo_item_db(usuario_logueado, texto_alta, "Ingreso"):
            st.session_state["nuevo_ingreso_tmp"] = ""
            st.toast(f"Ingreso '{texto_alta}' añadido.", icon="➕")
        else:
            st.error("Este concepto de ingreso ya está registrado.")

with st.expander("👥 Añadir casilla para sueldos compartidos"):
    st.text_input("Nombre de la nueva fuente de ingreso (Ej: Sueldo Milena):", value=st.session_state["nuevo_ingreso_tmp"], key="campo_alta_ingreso", on_change=ejecutar_alta_ingreso)
    st.button("➕ Confirmar Nuevo Ingreso", use_container_width=True)

saldos_db = obtener_todos_los_saldos(usuario_logueado)

total_ingresos = 0.0
ingresos_existentes = [row for row in saldos_db if row[2] == "Ingreso"]

for item_nom, valor_monto, tipo_flujo in ingresos_existentes:
    col_ing_txt, col_ing_del = st.columns([0.85, 0.15])
    
    with col_ing_txt:
        ingreso_num = st.number_input(f"Monto {item_nom} (CLP $):", value=int(valor_monto), min_value=0, step=1000, key=f"ingreso_input_{item_nom}")
    
    with col_ing_del:
        st.markdown("<div style='padding-top: 28px;'></div>", unsafe_allow_html=True)
        if st.button("🗑️", key=f"del_ing_{item_nom}", help=f"Eliminar {item_nom}"):
            eliminar_item_db(usuario_logueado, item_nom, "Ingreso")
            st.rerun()
            
    if ingreso_num != valor_monto:
        actualizar_monto_db(usuario_logueado, item_nom, float(ingreso_num))
        st.rerun()
        
    total_ingresos += ingreso_num

st.markdown("---")
st.markdown("### 💸 Controles Operativos de Gastos")

if "gasto_texto_tmp" not in st.session_state:
    st.session_state["gasto_texto_tmp"] = ""

def ejecutar_alta_gasto():
    texto_alta = st.session_state["campo_alta_gasto"].strip()
    if texto_alta:
        if insertar_nuevo_item_db(usuario_logueado, texto_alta, "Gasto"):
            st.session_state["gasto_texto_tmp"] = ""
            st.toast(f"'{texto_alta}' agregado.", icon="➕")
        else:
            st.error("Este ítem de gasto ya se encuentra registrado.")

def ejecutar_baja_gasto():
    target_baja = st.session_state["selector_baja_gasto"]
    if target_baja != "Seleccione...":
        eliminar_item_db(usuario_logueado, target_baja, "Gasto")
        st.toast(f"'{target_baja}' eliminado.", icon="🗑️")

col_add, col_del = st.columns(2)

with col_add:
    st.text_input("Añadir nuevo concepto de gasto:", value=st.session_state["gasto_texto_tmp"], key="campo_alta_gasto", placeholder="Ej: Supermercado", on_change=ejecutar_alta_gasto)
    st.button("➕ Añadir Gasto", use_container_width=True)

with col_del:
    saldos_db = obtener_todos_los_saldos(usuario_logueado)
    gastos_existentes = [row[0] for row in saldos_db if row[2] == "Gasto"]
    st.selectbox("Eliminar un gasto:", ["Seleccione..."] + gastos_existentes, key="selector_baja_gasto", on_change=ejecutar_baja_gasto)
    st.button("🗑️ Eliminar Gasto", use_container_width=True)

st.markdown("#### Valores del Periodo")
total_gastos = 0.0
gastos_filtrados = {}

for item, valor, tipo in saldos_db:
    if tipo == "Gasto":
        gasto_num = st.number_input(f"Monto {item} (CLP $):", value=int(valor), min_value=0, step=1000, key=f"gasto_input_{item}")
        
        if gasto_num != valor:
            actualizar_monto_db(usuario_logueado, item, float(gasto_num))
            st.rerun()
            
        total_gastos += gasto_num
        if gasto_num > 0:
            gastos_filtrados[item] = gasto_num

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
        guardar_mes_en_historico(usuario_logueado, periodo_completo, obtener_todos_los_saldos(usuario_logueado))
        conn = sqlite3.connect(RUTA_DB)
        cursor = conn.cursor()
        cursor.execute("UPDATE mes_activo SET monto = 0.0 WHERE usuario_id = ?", (usuario_logueado,))
        conn.commit()
        conn.close()
        st.success(f"¡Periodo contable {periodo_completo} indexado con éxito!")
        st.rerun()
    else:
        st.warning("No hay transacciones para archivar.")

st.subheader("📁 Periodos Consolidados")
meses_hist = obtener_meses_historicos(usuario_logueado)
mes_hist_sel = st.selectbox("Selecciona un mes para auditar:", ["Seleccione..."] + [m for m in meses_hist])

if mes_hist_sel != "Seleccione...":
    conn = sqlite3.connect(RUTA_DB)
    df_resumen = pd.read_sql_query("SELECT item as 'Concepto', monto as 'Monto' FROM historico WHERE usuario_id = ? AND mes_ano = ?", conn, params=(usuario_logueado, mes_hist_sel))
    conn.close()
    df_resumen['Monto'] = df_resumen['Monto'].apply(formatear_clp)
    st.dataframe(df_resumen, use_container_width=True)

st.markdown("---")
if st.button("🔒 Cerrar Sesión Segura", use_container_width=True):
    st.session_state["sesion_activa"] = False
    st.session_state["cuenta_usuario"] = ""
    st.rerun()







