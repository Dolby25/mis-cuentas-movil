import streamlit as st
import sqlite3
import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

st.set_page_config(page_title="Control Financiero Multi-User", page_icon="🔒", layout="centered")

RUTA_CARPETA = os.path.dirname(os.path.abspath(__file__))
RUTA_DB = os.path.join(RUTA_CARPETA, "control_cuentas_dinamico.db")
RUTA_EXCEL = os.path.join(RUTA_CARPETA, "reporte_finanzas.xlsx")

def inicializar_db():
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    
    # 1. Tabla Maestra de Usuarios
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            email TEXT PRIMARY KEY,
            password TEXT NOT NULL,
            pin TEXT NOT NULL
        )
    """)
    
    # 2. Verificar si la tabla mes_activo es antigua y necesita migración estructural
    cursor.execute("PRAGMA table_info(mes_activo)")
    columnas = [info[1] for info in cursor.fetchall()]
    
    if len(columnas) > 0 and "email" not in columnas:
        # Si la tabla existe pero no tiene la columna email, la reiniciamos de forma segura
        cursor.execute("DROP TABLE mes_activo")
        
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS mes_activo (
            email TEXT,
            item TEXT,
            monto REAL NOT NULL,
            tipo TEXT NOT NULL,
            PRIMARY KEY (email, item)
        )
    """)
    
    # 3. Verificar si la tabla historico es antigua y necesita migración
    cursor.execute("PRAGMA table_info(historico)")
    columnas_hist = [info[1] for info in cursor.fetchall()]
    if len(columnas_hist) > 0 and "email" not in columnas_hist:
        cursor.execute("DROP TABLE historico")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS historico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT NOT NULL,
            mes_ano TEXT NOT NULL,
            item TEXT NOT NULL,
            monto REAL NOT NULL,
            tipo TEXT NOT NULL DEFAULT 'Gasto'
        )
    """)
    
    # 4. Tabla de Metas de Ahorro
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS metas_ahorro (
            email TEXT,
            concepto TEXT,
            monto_meta REAL NOT NULL,
            monto_actual REAL NOT NULL,
            fecha_tope TEXT NOT NULL,
            PRIMARY KEY (email, concepto)
        )
    """)
    conn.commit()
    conn.close()
def actualizar_monto_db(email, item, nuevo_monto):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("UPDATE mes_activo SET monto = ? WHERE email = ? AND item = ?", (nuevo_monto, email, item))
    conn.commit()
    conn.close()

def insertar_nuevo_item_db(email, item, tipo="Gasto"):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO mes_activo (email, item, monto, tipo) VALUES (?, ?, 0.0, ?)", (email, item, tipo))
        conn.commit()
        exito = True
    except sqlite3.IntegrityError:
        exito = False
    conn.close()
    return exito

def cambiar_nombre_ingreso_db(email, viejo_nombre, nuevo_nombre):
    if viejo_nombre == nuevo_nombre:
        return True
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE mes_activo SET item = ? WHERE email = ? AND item = ? AND tipo = 'Ingreso'", (nuevo_nombre, email, viejo_nombre))
        conn.commit()
        exito = True
    except sqlite3.IntegrityError:
        exito = False
    conn.close()
    return exito

def eliminar_gasto_db(email, item):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM mes_activo WHERE email = ? AND item = ? AND tipo = 'Gasto'", (email, item))
    conn.commit()
    conn.close()

def obtener_todos_los_saldos(email):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT item, monto, tipo FROM mes_activo WHERE email = ?", (email,))
    datos = cursor.fetchall()
    conn.close()
    
    tiene_ingreso = any(row[2] == "Ingreso" for row in datos)
    if not tiene_ingreso:
        conn = sqlite3.connect(RUTA_DB)
        cursor = conn.cursor()
        cursor.execute("INSERT OR IGNORE INTO mes_activo (email, item, monto, tipo) VALUES (?, 'Ingresos Principales', 0.0, 'Ingreso')", (email,))
        conn.commit()
        cursor.execute("SELECT item, monto, tipo FROM mes_activo WHERE email = ?", (email,))
        datos = cursor.fetchall()
        conn.close()
    return datos
def guardar_mes_en_historico(email, mes_ano, datos_saldos):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM historico WHERE email = ? AND mes_ano = ?", (email, mes_ano))
    for item, monto, tipo in datos_saldos:
        cursor.execute("INSERT INTO historico (email, mes_ano, item, monto, tipo) VALUES (?, ?, ?, ?, ?)", (email, mes_ano, item, monto, tipo))
    conn.commit()
    conn.close()
    exportar_datos_a_excel(email)

def exportar_datos_a_excel(email):
    conn = sqlite3.connect(RUTA_DB)
    df_historico = pd.read_sql_query("SELECT mes_ano, item, monto, tipo FROM historico WHERE email = ?", conn, params=(email,))
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

def obtener_meses_historicos(email):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT mes_ano FROM historico WHERE email = ? ORDER BY id DESC", (email,))
    meses = [fila for fila in cursor.fetchall()]
    conn.close()
    return meses

def formatear_clp(monto):
    return f"${int(round(abs(monto))):,}".replace(",", ".") if monto >= 0 else f"-${int(round(abs(monto))):,}".replace(",", ".")

def limpiar_y_convertir_monto(texto_monto):
    texto_limpio = texto_monto.replace("$", "").replace(".", "").strip()
    return float(texto_limpio) if texto_limpio.isdigit() else 0.0

def buscar_usuario(email):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT email, password, pin FROM usuarios WHERE email = ?", (email.strip().lower(),))
    fila = cursor.fetchone()
    conn.close()
    return fila

def registrar_usuario(email, password, pin):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO usuarios (email, password, pin) VALUES (?, ?, ?)", (email.strip().lower(), password, pin))
        conn.commit()
        exito = True
    except sqlite3.IntegrityError:
        exito = False
    conn.close()
    return exito
inicializar_db()

if "usuario_conectado" not in st.session_state:
    st.session_state["usuario_conectado"] = None
if "primer_filtro_ok" not in st.session_state:
    st.session_state["primer_filtro_ok"] = False
if "doble_factor_ok" not in st.session_state:
    st.session_state["doble_factor_ok"] = False

# Requerimiento: Pantalla unificada que solicita Correo y Contraseña juntos al iniciar
if not st.session_state["doble_factor_ok"]:
    st.title("🔐 Acceso al Sistema de Finanzas")
    
    if not st.session_state["primer_filtro_ok"]:
        # Menú táctil para elegir entre entrar o registrar una cuenta nueva
        modo_acceso = st.radio("Seleccione una opción:", ["Iniciar Sesión", "Crear Cuenta Nueva"], horizontal=True)
        
        email_input = st.text_input("Correo Electrónico:", placeholder="ejemplo@correo.com").strip().lower()
        password_input = st.text_input("Contraseña:", type="password", placeholder="••••••••")
        
        if modo_acceso == "Iniciar Sesión":
            if st.button("🚪 Entrar al Sistema", use_container_width=True):
                if email_input and password_input:
                    cuenta = buscar_usuario(email_input)
                    if cuenta and cuenta[1] == password_input:
                        st.session_state["usuario_conectado"] = email_input
                        st.session_state["primer_filtro_ok"] = True
                        st.rerun()
                    else:
                        st.error("El correo electrónico o la contraseña son incorrectos.")
                else:
                    st.warning("Por favor, rellene todos los campos.")
        
        else:
            reg_pin = st.text_input("Crea tu PIN de doble seguridad (4 números):", type="password", max_chars=4, placeholder="1234")
            if st.button("🆕 Registrarse e Ingresar", use_container_width=True):
                if email_input and password_input and len(reg_pin) == 4 and reg_pin.isdigit():
                    if buscar_usuario(email_input) is None:
                        if registrar_usuario(email_input, password_input, reg_pin):
                            st.session_state["usuario_conectado"] = email_input
                            st.session_state["primer_filtro_ok"] = True
                            st.success("¡Cuenta creada exitosamente!")
                            st.rerun()
                        else:
                            st.error("Ocurrió un problema durante el registro.")
                    else:
                        st.error("Este correo electrónico ya se encuentra registrado.")
                else:
                    st.error("Rellene los campos. El PIN debe ser estrictamente de 4 números.")
        st.stop()
        
    else:
        # FASE 2: Pantalla de Doble Factor por PIN Táctil solicitado
        cuenta_maestra = buscar_usuario(st.session_state["usuario_conectado"])
        st.subheader("🔢 Doble Factor: Ingrese su PIN")
        st.write(f"Usuario activo: `{st.session_state['usuario_conectado']}`")
        pin_ingresado = st.text_input("Código PIN de 4 dígitos:", type="password", max_chars=4, placeholder="••••")
        
        if st.button("🔓 Desbloquear Cuentas", use_container_width=True) or (len(pin_ingresado) == 4 and pin_ingresado == cuenta_maestra[2]):
            if pin_ingresado == cuenta_maestra[2]:
                st.session_state["doble_factor_ok"] = True
                st.rerun()
            elif len(pin_ingresado) == 4:
                st.error("El código PIN de seguridad es incorrecto.")
        st.stop()
user_email = st.session_state["usuario_conectado"]

st.title("📱 Control Financiero Móvil")
st.write(f"Usuario conectado: `{user_email}`")
if st.button("🔒 Cerrar Sesión", key="btn_logout"):
    st.session_state["usuario_conectado"] = None
    st.session_state["primer_filtro_ok"] = False
    st.session_state["doble_factor_ok"] = False
    st.rerun()

st.markdown("---")
tab_periodo, tab_ahorro = st.tabs(["💰 Control de Periodo", "🎯 Metas de Ahorro"])

with tab_periodo:
    st.subheader("📅 Periodo Contable Activo")
    col_mes, col_ano = st.columns(2)
    with col_mes:
        mes_sel = st.selectbox("Mes:", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"], index=8)
    with col_ano:
        ano_sel = st.selectbox("Año:", ["2025", "2026", "2027", "2028"], index=1)
    periodo_completo = f"{mes_sel} {ano_sel}"

    st.markdown("### 📥 Ingresos Consolidados")
    saldos_db = obtener_todos_los_saldos(user_email)
    nombre_ingreso_inicial = next((row[0] for row in saldos_db if row[2] == "Ingreso"), "Ingresos Principales")
    monto_ingreso_inicial = next((row[1] for row in saldos_db if row[2] == "Ingreso"), 0.0)

    nuevo_nombre_ingreso = st.text_input("Concepto de Ingreso:", value=nombre_ingreso_inicial)
    if nuevo_nombre_ingreso != nombre_ingreso_inicial and nuevo_nombre_ingreso.strip():
        cambiar_nombre_ingreso_db(user_email, nombre_ingreso_inicial, nuevo_nombre_ingreso.strip())
        st.rerun()

    monto_inicial_formateado = f"{int(monto_ingreso_inicial):,}".replace(",", ".") if monto_ingreso_inicial > 0 else ""
    ingreso_input_text = st.text_input(f"Monto {nuevo_nombre_ingreso} (CLP \$):", value=monto_inicial_formateado, placeholder="Ej: 1.500.000")
    ingreso_procesado = limpiar_y_convertir_monto(ingreso_input_text)
    
    if f"{int(ingreso_procesado):,}".replace(",", ".") != ingreso_input_text and ingreso_input_text != "":
        actualizar_monto_db(user_email, nuevo_nombre_ingreso, ingreso_procesado)
        st.rerun()
    elif ingreso_input_text == "" and monto_ingreso_inicial != 0.0:
        actualizar_monto_db(user_email, nuevo_nombre_ingreso, 0.0)
        st.rerun()

    monto_ingreso = ingreso_procesado

    st.markdown("---")
    st.markdown("### 💸 Controles Operativos de Gastos")

    if "gasto_texto_tmp" not in st.session_state:
        st.session_state["gasto_texto_tmp"] = ""

    def ejecutar_alta_gasto():
        texto_alta = st.session_state["campo_alta_gasto"].strip()
        if texto_alta and texto_alta != nuevo_nombre_ingreso:
            if insertar_nuevo_item_db(user_email, texto_alta, "Gasto"):
                st.session_state["gasto_texto_tmp"] = ""
                st.toast(f"'{texto_alta}' agregado.", icon="➕")

    def ejecutar_baja_gasto():
        target_baja = st.session_state["selector_baja_gasto"]
        if target_baja != "Seleccione...":
            eliminar_gasto_db(user_email, target_baja)
            st.toast(f"'{target_baja}' eliminado.", icon="🗑️")

    col_add, col_del = st.columns(2)
    with col_add:
        st.text_input("Añadir nuevo concepto de gasto:", value=st.session_state["gasto_texto_tmp"], key="campo_alta_gasto", placeholder="Ej: Supermercado", on_change=ejecutar_alta_gasto)
        st.button("➕ Añadir Gasto", use_container_width=True)
    with col_del:
        gastos_existentes = [row[0] for row in saldos_db if row[2] == "Gasto"]
        st.selectbox("Eliminar un gasto:", ["Seleccione..."] + gastos_existentes, key="selector_baja_gasto", on_change=ejecutar_baja_gasto)
        st.button("🗑️ Eliminar Gasto", use_container_width=True)

    st.markdown("#### Valores del Periodo")
    total_ingresos = monto_ingreso
    total_gastos = 0.0
    gastos_filtrados = {}

    for item, valor, tipo in saldos_db:
        if tipo == "Gasto":
            valor_inicial_formateado = f"{int(valor):,}".replace(",", ".") if valor > 0 else ""
            gasto_input_text = st.text_input(f"Monto {item} (CLP \$):", value=valor_inicial_formateado, key=f"gasto_input_{item}", placeholder="Ej: 45.000")
            gasto_procesado = limpiar_y_convertir_monto(gasto_input_text)
            
            if f"{int(gasto_procesado):,}".replace(",", ".") != gasto_input_text and gasto_input_text != "":
                actualizar_monto_db(user_email, item, gasto_procesado)
                st.rerun()
            elif gasto_input_text == "" and valor != 0.0:
                actualizar_monto_db(user_email, item, 0.0)
                st.rerun()
                
            total_gastos += gasto_procesado
            if gasto_procesado > 0:
                gastos_filtrados[item] = gasto_procesado

    saldo_disponible = total_ingresos - total_gastos
with tab_ahorro:
    st.subheader("🎯 Planificador de Metas de Ahorro Privadas")

    def ejecutar_alta_meta(email, concepto, meta, actual, fecha):
        conn = sqlite3.connect(RUTA_DB)
        cursor = conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO metas_ahorro VALUES (?, ?, ?, ?, ?)", (email, concepto, meta, actual, fecha))
        conn.commit()
        conn.close()

    with st.expander("➕ Configurar nueva meta o actualizar ahorro", expanded=False):
        concepto_meta = st.text_input("¿Para qué estás ahorrando?:", placeholder="Ej: Vacaciones")
        monto_meta_text = st.text_input("Monto Total Objetivo (CLP $):", placeholder="Ej: 2.000.000")
        monto_actual_text = st.text_input("¿Cuánto dinero tienes ahorrado ya? (CLP $):", value="0", placeholder="Ej: 400.000")
        fecha_tope_date = st.date_input("Tope de día para cumplir la meta:", value=datetime.now())
        
        if st.button("🎯 Establecer Meta de Ahorro", use_container_width=True):
            val_meta = limpiar_y_convertir_monto(monto_meta_text)
            val_actual = limpiar_y_convertir_monto(monto_actual_text)
            if concepto_meta.strip() and val_meta > 0:
                ejecutar_alta_meta(user_email, concepto_meta.strip(), val_meta, val_actual, fecha_tope_date.strftime("%Y-%m-%d"))
                st.success("¡Meta guardada exitosamente!")
                st.rerun()

    conn = sqlite3.connect(RUTA_DB)
    df_metas = pd.read_sql_query("SELECT * FROM metas_ahorro WHERE email = ?", conn, params=(user_email,))
    conn.close()

    if not df_metas.empty:
        st.markdown("#### Tus Objetivos Activos")
        for _, fila in df_metas.iterrows():
            concepto = fila["concepto"]
            m_meta = fila["monto_meta"]
            m_actual = fila["monto_actual"]
            f_tope = datetime.strptime(fila["fecha_tope"], "%Y-%m-%d").date()
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
                    st.write(f"Debes ahorrar: **{formatear_clp(faltante / dias_restantes)} / día**")
                elif dias_restantes == 0:
                    st.warning("⚠️ ¡Hoy es el día tope!")
                else:
                    st.error(f"❌ Expiró hace {abs(dias_restantes)} días.")
            if st.button(f"🗑️ Eliminar Meta: {concepto}", key=f"del_meta_{concepto}"):
                conn = sqlite3.connect(RUTA_DB)
                cursor = conn.cursor()
                cursor.execute("DELETE FROM metas_ahorro WHERE email = ? AND concepto = ?", (user_email, concepto))
                conn.commit()
                conn.close()
                st.rerun()

# --- BLOQUE GENERAL DE ANALÍTICAS E HISTORIAL DEL USUARIO ---
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
        guardar_mes_en_historico(user_email, periodo_completo, obtener_todos_los_saldos(user_email))
        conn = sqlite3.connect(RUTA_DB)
        cursor = conn.cursor()
        cursor.execute("UPDATE mes_activo SET monto = 0.0 WHERE email = ?", (user_email,))
        conn.commit()
        conn.close()
        st.success(f"¡Periodo contable {periodo_completo} indexado con éxito!")
        st.rerun()
    else:
        st.warning("No hay transacciones para archivar.")

st.subheader("📁 Periodos Consolidados")
meses_hist = obtener_meses_historicos(user_email)
mes_hist_sel = st.selectbox("Selecciona un mes para auditar:", ["Seleccione..."] + [m for m in meses_hist])

if mes_hist_sel != "Seleccione...":
    conn = sqlite3.connect(RUTA_DB)
    df_resumen = pd.read_sql_query("SELECT item as 'Concepto', monto as 'Monto' FROM historico WHERE email = ? AND mes_ano = ?", conn, params=(user_email, mes_hist_sel))
    conn.close()
    df_resumen['Monto'] = df_resumen['Monto'].apply(formatear_clp)
    st.dataframe(df_resumen, use_container_width=True)

