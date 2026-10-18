import streamlit as st
import sqlite3
import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import re
import random
import string
from datetime import datetime, timedelta

st.set_page_config(page_title="Administrador de finanzas", page_icon="💰", layout="centered")

# CONFIGURACIÓN DE LAS RUTAS DE LA BASE DE DATOS Y REPORTE EXCEL
RUTA_CARPETA = os.path.dirname(os.path.abspath(__file__))
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
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            identificador TEXT PRIMARY KEY,
            clave TEXT NOT NULL,
            hogar_id TEXT NOT NULL,
            es_propietario INTEGER NOT NULL DEFAULT 1,
            pin TEXT DEFAULT NULL
        )
    """)
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
            "es_propietario": True if resultado[2] == 1 else False,
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
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre, monto_objetivo, monto_actual, fecha_inicio, plazo_dias, completado FROM metas_ahorro WHERE hogar_id = ?", (hogar_id,))
    filas = cursor.fetchall()
    conn.close()
    columnas = ['id', 'nombre', 'monto_objetivo', 'monto_actual', 'fecha_inicio', 'plazo_dias', 'completado']
    return pd.DataFrame(filas, columns=columnas)

def abonar_a_meta_db(meta_id, monto):
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT monto_actual, monto_objetivo FROM metas_ahorro WHERE id = ?", (meta_id,))
    res = cursor.fetchone()
    if res:
        monto_actual = res[0]
        monto_objetivo = res[1]
        nuevo_monto = monto_actual + monto
        completado = 1 if nuevo_monto >= monto_objetivo else 0
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

def generar_grafico_torta(valores_egresos, etiquetas_leyenda):
    """Genera el gráfico de egresos de forma segura y compatible con Python 3.14."""
    fig, ax2 = plt.subplots(figsize=(6, 6))
    
    # Solución al error de Matplotlib: extracción explícita sin conversión forzada
    wedges, texts = ax2.pie(
        valores_egresos, 
        startangle=90, 
        counterclock=False,
        colors=plt.cm.Dark2.colors[:len(valores_egresos)] if len(valores_egresos) <= 8 else None
    )
    
    ax2.legend(
        handles=wedges, 
        labels=etiquetas_leyenda, 
        title="Conceptos de Egreso", 
        loc="center left", 
        bbox_to_anchor=(1, 0, 0.5, 1),
        facecolor='#1E222A', 
        edgecolor='#31333F', 
        textprops=dict(color='#E1E4E6', fontsize=9)
    )
    
    ax2.axis('equal')  
    plt.tight_layout()
    return fig
# Inicializar Base de Datos en el arranque
inicializar_db()

# Control de inicio de sesión en el estado de la sesión
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
    st.session_state["usuario_id"] = ""
    st.session_state["hogar_id"] = ""

st.title("💰 Control de Finanzas Móvil")

if not st.session_state["autenticado"]:
    st.subheader("Iniciar Sesión / Registro")
    
    usuario = st.text_input("Correo electrónico o identificador:")
    clave = st.text_input("Contraseña:", type="password")
    
    col_login, col_reg = st.columns(2)
    
    with col_login:
        if st.button("Ingresar", use_container_width=True):
            datos_usuario = verificar_usuario_db(usuario, clave)
            if datos_usuario:
                st.session_state["autenticado"] = True
                st.session_state["usuario_id"] = usuario
                st.session_state["hogar_id"] = datos_usuario["hogar_id"]
                st.success("¡Bienvenido!")
                st.rerun()
            else:
                st.error("Credenciales incorrectas.")
                
    with col_reg:
        if st.button("Registrarse como nuevo hogar", use_container_width=True):
            if usuario and clave:
                if registrar_usuario_db(usuario, clave):
                    st.success("Usuario registrado con éxito. Ya puedes ingresar.")
                else:
                    st.error("El usuario ya existe o no se pudo registrar.")
            else:
                st.warning("Por favor completa los campos.")

else:
    # Menú lateral para usuarios autenticados
    st.sidebar.title(f"Hola, {st.session_state['usuario_id']}")
    st.sidebar.info(f"ID Hogar: {st.session_state['hogar_id']}")
    
    opcion = st.sidebar.radio("Menú Principal", ["Resumen Mensual", "Añadir / Eliminar Ítems", "Metas de Ahorro"])
    
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state["autenticado"] = False
        st.rerun()
        
    # --- PANTALLA 1: RESUMEN MENSUAL Y GRÁFICOS ---
    if opcion == "Resumen Mensual":
        st.header("📊 Resumen del Mes Activo")
        
        saldos = obtener_todos_los_saldos(st.session_state["hogar_id"])
        
        if not saldos:
            st.info("No hay transacciones ni conceptos registrados para este mes.")
        else:
            df_saldos = pd.DataFrame(saldos, columns=["Concepto", "Monto", "Tipo"])
            
            st.subheader("Tus saldos y presupuestos")
            for index, row in df_saldos.iterrows():
                nuevo_val = st.number_input(f"{row['Concepto']} ({row['Tipo']})", value=float(row['Monto']), step=10.0, key=f"monto_{index}")
                if nuevo_val != row['Monto']:
                    actualizar_monto_db(st.session_state["hogar_id"], row['Concepto'], nuevo_val)
                    st.rerun()
            
            # Filtrar gastos para el gráfico de torta corregido
            df_gastos = df_saldos[df_saldos["Tipo"] == "Gasto"]
            
            if not df_gastos.empty and df_gastos["Monto"].sum() > 0:
                st.subheader("Distribución de Gastos")
                figura = generar_grafico_torta(df_gastos["Monto"].tolist(), df_gastos["Concepto"].tolist())
                st.pyplot(figura)
            else:
                st.info("Agrega montos mayores a 0 en tus gastos para ver el gráfico de distribución.")

    # --- PANTALLA 2: CONFIGURACIÓN DE CONCEPTOS ---
    elif opcion == "Añadir / Eliminar Ítems":
        st.header("⚙️ Administrar Conceptos Financieros")
        
        st.subheader("Agregar nuevo concepto")
        nuevo_item = st.text_input("Nombre del concepto (Ej: Arriendo, Supermercado, Sueldo):")
        tipo_item = st.selectbox("Tipo de flujo:", ["Gasto", "Ingreso"])
        
        if st.button("Guardar concepto"):
            if nuevo_item.strip():
                if insertar_nuevo_item_db(st.session_state["hogar_id"], nuevo_item.strip(), tipo_item):
                    st.success(f"'{nuevo_item}' agregado correctamente.")
                    st.rerun()
                else:
                    st.error("Este concepto ya está registrado.")
            else:
                st.warning("El nombre no puede estar vacío.")
                
        st.markdown("---")
        st.subheader("Eliminar conceptos existentes")
        saldos_actuales = obtener_todos_los_saldos(st.session_state["hogar_id"])
        
        if saldos_actuales:
            for item, monto, tipo in saldos_actuales:
                col_txt, col_btn = st.columns([4, 1])
                col_txt.write(f"**{item}** — {tipo}")
                if col_btn.button("❌", key=f"del_{item}_{tipo}"):
                    eliminar_item_db(st.session_state["hogar_id"], item, tipo)
                    st.rerun()
        else:
            st.info("No hay conceptos creados.")

    # --- PANTALLA 3: METAS DE AHORRO ---
    elif opcion == "Metas de Ahorro":
        st.header("🎯 Metas de Ahorro")
        
        with st.expander("➕ Crear Nueva Meta"):
            nombre_meta = st.text_input("¿Para qué estás ahorrando?")
            objetivo_meta = st.number_input("Monto Objetivo ($):", min_value=1.0, step=50.0)
            plazo_meta = st.number_input("Plazo estimado (en días):", min_value=1, step=30)
            
            if st.button("Crear Meta"):
                if nombre_meta.strip():
                    insertar_meta_db(st.session_state["hogar_id"], nombre_meta.strip(), objetivo_meta, plazo_meta)
                    st.success("¡Meta creada con éxito!")
                    st.rerun()
                else:
                    st.error("Por favor, introduce un nombre para la meta.")
                    
        df_metas = obtener_metas_db(st.session_state["hogar_id"])
        if df_metas.empty:
            st.info("Aún no tienes metas definidas.")
        else:
            for _, fila in df_metas.iterrows():
                st.subheader(f"🏆 {fila['nombre']}")
                progreso = fila['monto_actual'] / fila['monto_objetivo']
                st.progress(min(progreso, 1.0))
                st.write(f"Ahorrado: **${fila['monto_actual']:.2f}** de **${fila['monto_objetivo']:.2f}**")
                
                monto_aporte = st.number_input(f"Abonar a '{fila['nombre']}':", min_value=0.0, step=10.0, key=f"aporte_{fila['id']}")
                if monto_aporte > 0:
                    if st.button(f"Abonar ${monto_aporte}", key=f"btn_ap_{fila['id']}"):
                        abonar_a_meta_db(fila['id'], monto_aporte)
                        st.success("¡Aporte registrado!")
                        st.rerun()
