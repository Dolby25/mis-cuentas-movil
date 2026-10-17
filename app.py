import streamlit as st
import sqlite3
import os
import pandas as pd
import matplotlib.pyplot as plt
import re
import random
import string
from datetime import datetime, timedelta

st.set_page_config(page_title="Control Financiero Familiar", page_icon="💰", layout="centered")

# RUTA DE ARCHIVOS DE LA APLICACIÓN
RUTA_CARPETA = os.path.dirname(os.path.abspath(__file__))
# Cambiamos a la versión v5 para incorporar las tablas de PIN y Metas con Logros
RUTA_DB = os.path.join(RUTA_CARPETA, "control_cuentas_v5.db")
RUTA_EXCEL = os.path.join(RUTA_CARPETA, "reporte_finanzas.xlsx")

def es_correo_valido(correo):
    """Verifica si la estructura corresponde a un correo real."""
    patron = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(patron, correo) is not None

def generar_string_aleatorio():
    """Genera un código dinámico único de 6 caracteres."""
    return "INV-" + "".join(random.choices(string.ascii_uppercase + string.digits, k=4))

def inicializar_db():
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    # Tabla de usuarios incluye la columna para el PIN de 4 dígitos
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            identificador TEXT PRIMARY KEY,
            clave TEXT NOT NULL,
            hogar_id TEXT NOT NULL,
            es_propietario INTEGER NOT NULL DEFAULT 1,
            pin TEXT DEFAULT NULL
        )
    """)
    # Tabla temporal de invitaciones de uso único
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS codigos_invitacion (
            hogar_id TEXT PRIMARY KEY,
            codigo TEXT NOT NULL UNIQUE
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS mes_activo (
            hogar_id TEXT,
            item TEXT,
            monto REAL NOT NULL,
            tipo TEXT NOT NULL,
            PRIMARY KEY (hogar_id, item)
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS historico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            hogar_id TEXT NOT NULL,
            mes_ano TEXT NOT NULL,
            item TEXT NOT NULL,
            monto REAL NOT NULL,
            tipo TEXT NOT NULL DEFAULT 'Gasto'
        )
    """)
    # Nueva tabla para las Metas de Ahorro y Logros del Hogar
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS metas_ahorro (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            hogar_id TEXT NOT NULL,
            nombre TEXT NOT NULL,
            monto_objetivo REAL NOT NULL,
            monto_actual REAL NOT NULL DEFAULT 0.0,
            fecha_inicio TEXT NOT NULL,
            plazo_dias INTEGER NOT NULL,
            completado INTEGER NOT NULL DEFAULT 0
        )
    """)
    conn.commit()
    conn.close()

def registrar_usuario_db(identificador, clave):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    nuevo_hogar = "HOGAR-" + "".join(random.choices(string.ascii_uppercase + string.digits, k=6))
    try:
        cursor.execute("INSERT INTO usuarios (identificador, clave, hogar_id, es_propietario) VALUES (?, ?, ?, 1)", (identificador, clave, nuevo_hogar))
        cursor.execute("INSERT INTO mes_activo (hogar_id, item, monto, tipo) VALUES (?, 'Ingresos Principales', 0.0, 'Ingreso')", (nuevo_hogar,))
        conn.commit()
        exito = True
    except sqlite3.IntegrityError:
        exito = False
    conn.close()
    return exito

def verificar_usuario_db(identificador, clave):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT clave, hogar_id, es_propietario, pin FROM usuarios WHERE identificador = ?", (identificador,))
    resultado = cursor.fetchone()
    conn.close()
    if resultado and resultado[0] == clave:
        return {
            "hogar_id": resultado[1],
            "es_propietario": resultado[2],
            "pin": resultado[3]
        }
    return None

def actualizar_pin_db(usuario_id, nuevo_pin):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("UPDATE usuarios SET pin = ? WHERE identificador = ?", (nuevo_pin, usuario_id))
    conn.commit()
    conn.close()

def solicitar_nuevo_codigo_db(hogar_id):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    nuevo_codigo = generar_string_aleatorio()
    cursor.execute("INSERT OR REPLACE INTO codigos_invitacion (hogar_id, codigo) VALUES (?, ?)", (hogar_id, nuevo_codigo))
    conn.commit()
    conn.close()
    return nuevo_codigo

def obtener_codigo_activo_db(hogar_id):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT codigo FROM codigos_invitacion WHERE hogar_id = ?", (hogar_id,))
    fila = cursor.fetchone()
    conn.close()
    return fila[0] if fila else None

def enlazar_y_quemar_codigo_db(usuario_id, codigo_ingresado):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT hogar_id FROM codigos_invitacion WHERE codigo = ?", (codigo_ingresado,))
    resultado = cursor.fetchone()
    if resultado:
        hogar_destino = resultado[0]
        cursor.execute("UPDATE usuarios SET hogar_id = ?, es_propietario = 0 WHERE identificador = ?", (hogar_destino, usuario_id))
        cursor.execute("DELETE FROM codigos_invitacion WHERE codigo = ?", (codigo_ingresado,))
        conn.commit()
        exito = True
    else:
        exito = False
    conn.close()
    return exito
def insertar_meta_db(hogar_id, nombre, objetivo, plazo):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    cursor.execute("""
        INSERT INTO metas_ahorro (hogar_id, nombre, monto_objetivo, monto_actual, fecha_inicio, plazo_dias, completado)
        VALUES (?, ?, ?, 0.0, ?, ?, 0)
    """, (hogar_id, nombre, objetivo, fecha_hoy, plazo))
    conn.commit()
    conn.close()

def obtener_metas_db(hogar_id):
    conn = sqlite3.connect(RUTA_DB)
    df = pd.read_sql_query("SELECT id, nombre, monto_objetivo, monto_actual, fecha_inicio, plazo_dias, completado FROM metas_ahorro WHERE hogar_id = ?", conn)
    conn.close()
    return df

def abonar_a_meta_db(meta_id, monto):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT monto_actual, monto_objetivo FROM metas_ahorro WHERE id = ?", (meta_id,))
    res = cursor.fetchone()
    if res:
        nuevo_monto = res[0] + monto
        completado = 1 if nuevo_monto >= res[1] else 0
        cursor.execute("UPDATE metas_ahorro SET monto_actual = ?, completado = ? WHERE id = ?", (nuevo_monto, completado, meta_id))
        conn.commit()
    conn.close()

def actualizar_monto_db(hogar_id, item, nuevo_monto):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("UPDATE mes_activo SET monto = ? WHERE hogar_id = ? AND item = ?", (nuevo_monto, hogar_id, item))
    conn.commit()
    conn.close()

def insertar_nuevo_item_db(hogar_id, item, tipo="Gasto"):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO mes_activo (hogar_id, item, monto, tipo) VALUES (?, ?, 0.0, ?)", (hogar_id, item, tipo))
        conn.commit()
        exito = True
    except sqlite3.IntegrityError:
        exito = False
    conn.close()
    return exito

def eliminar_item_db(hogar_id, item, tipo):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM mes_activo WHERE hogar_id = ? AND item = ? AND tipo = ?", (hogar_id, item, tipo))
    conn.commit()
    conn.close()

def obtener_todos_los_saldos(hogar_id):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT item, monto, tipo FROM mes_activo WHERE hogar_id = ?", (hogar_id,))
    datos = cursor.fetchall()
    conn.close()
    return datos

def guardar_mes_en_historico(hogar_id, mes_ano, datos_saldos):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM historico WHERE hogar_id = ? AND mes_ano = ?", (hogar_id, mes_ano))
    for item, monto, tipo in datos_saldos:
        cursor.execute("INSERT INTO historico (hogar_id, mes_ano, item, monto, tipo) VALUES (?, ?, ?, ?, ?)", (hogar_id, mes_ano, item, monto, tipo))
    conn.commit()
    conn.close()
    exportar_datos_a_excel(hogar_id)

def exportar_datos_a_excel(hogar_id):
    conn = sqlite3.connect(RUTA_DB)
    df_historico = pd.read_sql_query("SELECT mes_ano, item, monto, tipo FROM historico WHERE hogar_id = ?", conn, params=(hogar_id,))
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

def obtener_meses_historicos(hogar_id):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT mes_ano FROM historico WHERE hogar_id = ? ORDER BY id DESC", (hogar_id,))
    meses = [fila for fila in cursor.fetchall()]
    conn.close()
    return meses

def formatear_clp(monto):
    valor_absoluto = int(round(abs(monto)))
    formato_numero = f"${valor_absoluto:,}".replace(",", ".")
    return formato_numero if monto >= 0 else f"-{formato_numero}"

inicializar_db()

# --- CONTROL DE ACCESO PERSISTENTE Y BLOQUEO POR PIN ---
if "cuenta_usuario" not in st.session_state:
    st.session_state["cuenta_usuario"] = ""
if "hogar_id" not in st.session_state:
    st.session_state["hogar_id"] = ""
if "es_propietario" not in st.session_state:
    st.session_state["es_propietario"] = True
if "pin_usuario" not in st.session_state:
    st.session_state["pin_usuario"] = None
if "pin_desbloqueado" not in st.session_state:
    st.session_state["pin_desbloqueado"] = False

# Caso 1: No hay sesión de correo iniciada (Primera vez en el celular)
if not st.session_state["cuenta_usuario"]:
    st.title("🔒 Registro e Inicio de Sesión Inicial")
    tab_log, tab_reg = st.tabs(["Iniciar Sesión", "Crear Cuenta"])
    
    with tab_log:
        u_input = st.text_input("Correo Electrónico:", key="main_u").strip().lower()
        p_input = st.text_input("Contraseña:", type="password", key="main_p")
        if st.button("🔓 Conectar Dispositivo", use_container_width=True):
            res = verificar_usuario_db(u_input, p_input)
            if res:
                st.session_state["cuenta_usuario"] = u_input
                st.session_state["hogar_id"] = res["hogar_id"]
                st.session_state["es_propietario"] = True if res["es_propietario"] == 1 else False
                st.session_state["pin_usuario"] = res["pin"]
                st.rerun()
            else:
                st.error("Credenciales incorrectas.")
                
    with tab_reg:
        u_reg = st.text_input("Registrar Nuevo Correo:", key="reg_u").strip().lower()
        p_reg = st.text_input("Definir Contraseña:", type="password", key="reg_pass")
        if st.button("🚀 Crear Cuenta Familiar", use_container_width=True):
            if es_correo_valido(u_reg) and p_reg.strip():
                if registrar_usuario_db(u_reg, p_reg.strip()):
                    st.success("¡Cuenta vinculada! Ahora puedes iniciar sesión en la primera pestaña.")
                else:
                    st.error("Ese correo ya existe en el sistema.")
            else:
                st.error("Ingresa un correo verdadero válido.")
    st.stop()

# Caso 2: El correo está conectado, pero exige el PIN rápido de 4 dígitos
if st.session_state["pin_usuario"] and not st.session_state["pin_desbloqueado"]:
    st.title("🔒 Acceso de Seguridad Rápido")
    st.info(f"Sesión en curso: **{st.session_state['cuenta_usuario']}**")
    pin_ingresado = st.text_input("Introduce tu PIN numérico de 4 dígitos:", type="password", max_chars=4, placeholder="••••")
    if st.button("🔓 Desbloquear App", use_container_width=True):
        if pin_ingresado == st.session_state["pin_usuario"]:
            st.session_state["pin_desbloqueado"] = True
            st.rerun()
        else:
            st.error("PIN incorrecto. Inténtalo de nuevo.")
    st.stop()

# Caso 3: El usuario está conectado por correo pero aún no define su PIN inicial
if not st.session_state["pin_usuario"]:
    st.title("⚙️ Configuración Obligatoria del PIN")
    st.warning("Define un código PIN corto de 4 números para abrir tu aplicación en el día a día sin escribir contraseñas.")
    nuevo_pin_1 = st.text_input("Crea tu PIN rápido de 4 dígitos:", type="password", max_chars=4, placeholder="Ej: 1234")
    nuevo_pin_2 = st.text_input("Confirma tu PIN:", type="password", max_chars=4)
    if st.button("💾 Guardar PIN de Seguridad", use_container_width=True):
        if nuevo_pin_1.isdigit() and len(nuevo_pin_1) == 4 and nuevo_pin_1 == nuevo_pin_2:
            actualizar_pin_db(st.session_state["cuenta_usuario"], nuevo_pin_1)
            st.session_state["pin_usuario"] = nuevo_pin_1
            st.session_state["pin_desbloqueado"] = True
            st.success("¡PIN configurado correctamente!")
            st.rerun()
        else:
            st.error("Los códigos deben ser de 4 dígitos, numéricos y coincidir exactamente.")
    st.stop()
# --- PANEL OPERATIVO DE LA APLICACIÓN (ACCESO COMPLETAMENTE AUTORIZADO) ---
usuario_logueado = st.session_state["cuenta_usuario"]
hogar_actual = st.session_state["hogar_id"]
es_propietario = st.session_state["es_propietario"]

st.title("📱 Sistema Financiero Familiar Compartido")
st.info(f"🟢 Conectado como: **{usuario_logueado}**")

# Pestañas principales de navegación móvil en la pantalla principal
tab_cuentas, tab_metas, tab_miembros = st.tabs(["📊 Control Mensual", "🎯 Metas de Ahorro", "👥 Invitar / Compartir"])

# --- TAB DE METAS DE AHORRO Y LOGROS ---
with tab_metas:
    st.markdown("### 🎯 Tus Objetivos de Ahorro")
    
    with st.expander("➕ Añadir Nueva Meta de Ahorro"):
        nom_meta = st.text_input("¿Qué quieres lograr? (Ej: Vacaciones o Auto):", placeholder="Escribe el nombre aquí...")
        obj_meta = st.number_input("Monto total que necesitan ahorrar (CLP $):", min_value=0, step=10000)
        plazo_meta = st.number_input("¿En cuántos días deben cumplirlo?:", min_value=1, step=1)
        if st.button("🎯 Crear Meta del Hogar", use_container_width=True):
            if nom_meta.strip() and obj_meta > 0:
                insertar_meta_db(hogar_actual, nom_meta.strip(), float(obj_meta), int(plazo_meta))
                st.success(f"¡Meta '{nom_meta}' guardada con un plazo de {plazo_meta} días!")
                st.rerun()

    df_metas = obtener_metas_db(hogar_actual)
    
    if not df_metas.empty:
        st.markdown("#### Progreso de Objetivos Activos")
        for idx, row in df_metas.iterrows():
            f_inicio = datetime.strptime(row['fecha_inicio'], "%Y-%m-%d")
            f_limite = f_inicio + timedelta(days=int(row['plazo_dias']))
            dias_restantes = (f_limite - datetime.now()).days
            
            st.markdown(f"**Objetivo: {row['nombre']}**")
            porcentaje = min(float(row['monto_actual'] / row['monto_objetivo']), 1.0)
            st.progress(porcentaje)
            
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                st.write(f"Ahorrado: {formatear_clp(row['monto_actual'])} de {formatear_clp(row['monto_objetivo'])}")
            with col_m2:
                if dias_restantes >= 0 and row['completado'] == 0:
                    st.write(f"⏳ Quedan **{dias_restantes} días** de plazo.")
                elif row['completado'] == 1:
                    st.write("🏆 ¡Objetivo cumplido!")
                else:
                    st.write("❌ Plazo vencido.")
            
            if row['completado'] == 0:
                monto_abono = st.number_input(f"Monto a depositar en {row['nombre']} (CLP $):", min_value=0, step=5000, key=f"abono_{row['id']}")
                if st.button(f"💰 Guardar Ahorro para {row['nombre']}", key=f"btn_abono_{row['id']}"):
                    if monto_abono > 0:
                        abonar_a_meta_db(row['id'], float(monto_abono))
                        st.rerun()
            st.markdown("---")
            
        # Galería de Logros Permanentes del Hogar
        st.markdown("### 🏆 Logros y Trofeos del Hogar")
        metas_ganadas = df_metas[df_metas['completado'] == 1]
        if not metas_ganadas.empty:
            for idx, m in metas_ganadas.iterrows():
                st.success(f"🏅 **LOGRO DESBLOQUEADO:** ¡Cumplieron la meta **'{m['nombre']}'** ahorrando un total de {formatear_clp(m['monto_objetivo'])} con éxito! Guardado permanentemente.")
        else:
            st.caption("Aún no tienen logros guardados. ¡Completen su primera meta de ahorro para ganar una medalla!")
    else:
        st.caption("No hay metas configuradas en este hogar.")
# --- TAB DE CONTROL MENSUAL CONTABLE ---
with tab_cuentas:
    col_mes, col_ano = st.columns(2)
    with col_mes:
        mes_sel = st.selectbox("Mes:", ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"], index=9)
    with col_ano:
        ano_sel = st.selectbox("Año:", ["2025", "2026", "2027", "2028"], index=1)
    periodo_completo = f"{mes_sel} {ano_sel}"

    st.markdown("### 📥 Ingresos Consolidados")
    if "nuevo_ingreso_tmp" not in st.session_state:
        st.session_state["nuevo_ingreso_tmp"] = ""

    def ejecutar_alta_ingreso():
        texto_alta = st.session_state["campo_alta_ingreso"].strip()
        if texto_alta:
            if insertar_nuevo_item_db(hogar_actual, texto_alta, "Ingreso"):
                st.session_state["nuevo_ingreso_tmp"] = ""
                st.toast(f"Ingreso '{texto_alta}' añadido.", icon="➕")

    with st.expander("👥 Añadir casilla para sueldos compartidos"):
        st.text_input("Nombre de la fuente de ingreso (Ej: Sueldo Milena):", value=st.session_state["nuevo_ingreso_tmp"], key="campo_alta_ingreso", on_change=ejecutar_alta_ingreso)
        st.button("Confirmar Nuevo Ingreso", use_container_width=True)

    saldos_db = obtener_todos_los_saldos(hogar_actual)
    total_ingresos = 0.0
    ingresos_existentes = [row for row in saldos_db if row == "Ingreso"]

    for item_nom, valor_monto, tipo_flujo in ingresos_existentes:
        col_ing_txt, col_ing_del = st.columns([0.85, 0.15])
        with col_ing_txt:
            ingreso_num = st.number_input(f"Monto {item_nom} (CLP $):", value=int(valor_monto), min_value=0, step=1000, key=f"ingreso_input_{item_nom}")
        with col_ing_del:
            st.markdown("<div style='padding-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("🗑️", key=f"del_ing_{item_nom}"):
                eliminar_item_db(hogar_actual, item_nom, "Ingreso")
                st.rerun()
        if ingreso_num != valor_monto:
            actualizar_monto_db(hogar_actual, item_nom, float(ingreso_num))
            st.rerun()
        total_ingresos += ingreso_num

    st.markdown("### 💸 Controles Operativos de Gastos")
    if "gasto_texto_tmp" not in st.session_state:
        st.session_state["gasto_texto_tmp"] = ""

    def ejecutar_alta_gasto():
        texto_alta = st.session_state["campo_alta_gasto"].strip()
        if texto_alta:
            if insertar_nuevo_item_db(hogar_actual, texto_alta, "Gasto"):
                st.session_state["gasto_texto_tmp"] = ""
                st.toast(f"'{texto_alta}' agregado.", icon="➕")

    def ejecutar_baja_gasto():
        target_baja = st.session_state["selector_baja_gasto"]
        if target_baja != "Seleccione...":
            eliminar_item_db(hogar_actual, target_baja, "Gasto")

    col_add, col_del = st.columns(2)
    with col_add:
        st.text_input("Añadir nuevo concepto de gasto:", value=st.session_state["gasto_texto_tmp"], key="campo_alta_gasto", placeholder="Ej: Supermercado", on_change=ejecutar_alta_gasto)
    with col_del:
        gastos_existentes = [row for row in saldos_db if row == "Gasto"]
        st.selectbox("Eliminar un gasto:", ["Seleccione..."] + gastos_existentes, key="selector_baja_gasto", on_change=ejecutar_baja_gasto)

    st.markdown("#### Valores del Periodo")
    total_gastos = 0.0
    gastos_filtrados = {}

    for item, valor, tipo in saldos_db:
        if tipo == "Gasto":
            gasto_num = st.number_input(f"Monto {item} (CLP $):", value=int(valor), min_value=0, step=1000, key=f"gasto_input_{item}")
            if gasto_num != valor:
                actualizar_monto_db(hogar_actual, item, float(gasto_num))
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

    if st.button("🗄️ Archivar Periodo y Actualizar Excel", use_container_width=True):
        if total_ingresos + total_gastos > 0:
            guardar_mes_en_historico(hogar_actual, periodo_completo, obtener_todos_los_saldos(hogar_actual))
            st.success(f"¡Periodo contable {periodo_completo} indexado con éxito!")
            st.rerun()

    st.subheader("📁 Periodos Consolidados")
    meses_hist = obtener_meses_historicos(hogar_actual)
    mes_hist_sel = st.selectbox("Selecciona un mes para auditar:", ["Seleccione..."] + [m for m in meses_hist])
    if mes_hist_sel != "Seleccione...":
        conn = sqlite3.connect(RUTA_DB)
        df_resumen = pd.read_sql_query("SELECT item as 'Concepto', monto as 'Monto' FROM historico WHERE hogar_id = ? AND mes_ano = ?", conn, params=(hogar_actual, mes_hist_sel))
        conn.close()
        df_resumen['Monto'] = df_resumen['Monto'].apply(formatear_clp)
        st.dataframe(df_resumen, use_container_width=True)

# --- TAB DE INVITACIONES COMPARTIDAS ---
with tab_miembros:
    tab_env, tab_rec = st.tabs(["✉️ Emitir Invitación", "🔗 Enlazar Código"])
    with tab_env:
        if es_propietario:
            st.write("Genera un token numérico único para invitar a alguien a tu panel.")
            codigo_actual_db = obtener_codigo_activo_db(hogar_actual)
            if codigo_actual_db:
                st.markdown("Código activo para enviar:")
                st.code(codigo_actual_db, language="text")
            else:
                st.warning("No tienes ningún código activo.")
            if st.button("🔄 Generar / Renovar Código Único", key="gen_token"):
                solicitar_nuevo_codigo_db(hogar_actual)
                st.rerun()
        else:
            st.warning("Solo el propietario del hogar puede emitir códigos de invitación.")
            
    with tab_rec:
        st.write("Ingresa el código que recibiste para unirte a las cuentas de otro hogar.")
        cod_ing = st.text_input("Introduce el código de invitación (Ej: INV-A1B2):", placeholder="INV-XXXX", key="p3_token").strip().upper()
        if st.button("Confirmar Vinculación", use_container_width=True, key="p3_btn_token"):
            if cod_ing:
                if enlazar_y_quemar_codigo_db(usuario_logueado, cod_ing):
                    st.success("¡Cuentas enlazadas de forma correcta!")
                    st.session_state["pin_desbloqueado"] = False
                    st.rerun()
                else:
                    st.error("El código es inválido o ya caducó.")

# Botón inferior para re-bloquear la pantalla con PIN manualmente
st.markdown("---")
if st.button("🔒 Bloquear Pantalla (Fijar PIN)", use_container_width=True):
    st.session_state["pin_desbloqueado"] = False
    st.rerun()












