import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import os  
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd

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
    # Insertar el flujo base dinámico de ingresos si el sistema está vacío
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

def formatear_entrada_dinamica(event):
    widget = event.widget
    texto_actual = widget.get().replace("$", "").replace(".", "").strip()
    if texto_actual.isdigit():
        numero = int(texto_actual)
        texto_formateado = f"${numero:,}".replace(",", ".")
        widget.delete(0, tk.END)
        widget.insert(0, texto_formateado)

modo_borrar_activo = False  

def alternar_modo_borrado():
    global modo_borrar_activo
    modo_borrar_activo = not modo_borrar_activo
    if modo_borrar_activo:
        btn_basurero.config(bg="#E74C3C", fg="#FFFFFF", text="✕ Cancelar")
    else:
        btn_basurero.config(bg="#2E3238", fg="#E74C3C", text="🗑️ Editar")
    reconstruir_lista_gastos_visual()
def registrar_valor(item, entry_widget):
    try:
        monto_raw = entry_widget.get().replace("$", "").replace(".", "").strip()
        valor_ingresado = float(monto_raw) if monto_raw else 0.0
        if valor_ingresado < 0:
            raise ValueError
    except ValueError:
        messagebox.showerror("Aviso del Sistema", "Por favor, ingrese un monto numérico válido.")
        return
    actualizar_monto_db(item, valor_ingresado)
    actualizar_pantalla()

def registrar_cambio_nombre_ingreso():
    nuevo_nombre = ent_nombre_ingreso.get().strip()
    viejos_datos = obtener_todos_los_saldos()
    nombre_anterior = next((row for row in viejos_datos if row == "Ingreso"), "Ingresos Principales")
    if not nuevo_nombre:
        messagebox.showwarning("Aviso del Sistema", "El campo no puede quedar vacío.")
        ent_nombre_ingreso.insert(0, nombre_anterior)
        return
    if cambiar_nombre_ingreso_db(nombre_anterior, nuevo_nombre):
        actualizar_pantalla()
    else:
        messagebox.showerror("Duplicado", "Ese concepto ya está en uso.")
def añadir_nuevo_gasto_interfaz():
    nuevo_nombre = entry_crear_gasto.get().strip()
    if not nuevo_nombre:
        messagebox.showwarning("Aviso del Sistema", "La casilla de concepto está vacía.")
        return
    if nuevo_nombre == ent_nombre_ingreso.get().strip():
        messagebox.showerror("Restricción", "Coincide con tu casilla activa de ingresos.")
        return
    if insertar_nuevo_item_db(nuevo_nombre, "Gasto"):
        entry_crear_gasto.delete(0, tk.END)
        reconstruir_lista_gastos_visual()
        actualizar_pantalla()
        canvas_scroll.update_idletasks()
        canvas_scroll.yview_moveto(1.0)
    else:
        messagebox.showerror("Duplicado", "Este ítem ya existe.")

def borrar_gasto_interfaz(item):
    if messagebox.askyesno("Confirmar Acción", f"¿Desea eliminar permanente el ítem '{item}'?"):
        eliminar_gasto_db(item)
        reconstruir_lista_gastos_visual()
        actualizar_pantalla()

def guardar_registro_mes():
    nombre_mes_completo = f"{combo_mes.get()} {combo_ano.get()}"
    datos_completos = obtener_todos_los_saldos()
    if sum(row for row in datos_completos) == 0:
        messagebox.showwarning("Aviso del Sistema", "No hay transacciones registradas.")
        return
    if messagebox.askyesno("Archivar Periodo", f"¿Desea consolidar el periodo {nombre_mes_completo}?"):
        guardar_mes_en_historico(nombre_mes_completo, datos_completos)
        for item, _, _ in datos_completos:
            actualizar_monto_db(item, 0.0)
        for entry in todas_las_entradas.values():
            entry.delete(0, tk.END)
        messagebox.showinfo("Éxito", "El periodo mensual ha sido guardado.")
        actualizar_pantalla()
        actualizar_lista_historica()
def revisar_mes_guardado(event):
    seleccion = lista_historial.curselection()
    if not seleccion:
        return
    mes_ano = lista_historial.get(seleccion).replace(" 📁 ", "").strip()
    conn = sqlite3.connect(RUTA_DB)
    cursor = conn.cursor()
    cursor.execute("SELECT item, monto FROM historico WHERE mes_ano = ?", (mes_ano,))
    registros = cursor.fetchall()
    conn.close()
    if not registros:
        return
    sub_ventana = tk.Toplevel(ventana)
    sub_ventana.title(f"Consulta de Datos // {mes_ano}")
    sub_ventana.geometry("450x500")
    sub_ventana.configure(bg="#1E2229")
    tk.Label(sub_ventana, text=f"Resumen Consolidado: {mes_ano}", bg="#1E2229", fg="#E1E4E6", font=("Segoe UI", 11, "bold")).pack(pady=15)
    tabla_rev = ttk.Treeview(sub_ventana, columns=("Item", "Monto"), show="headings", height=15)
    tabla_rev.heading("Item", text="CONCEPTO O ÍTEM")
    tabla_rev.heading("Monto", text="MONTO GUARDADO")
    tabla_rev.column("Item", width=220, anchor="w")
    tabla_rev.column("Monto", width=160, anchor="center")
    tabla_rev.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
    for item, monto in registros:
        tabla_rev.insert("", tk.END, values=(item, formatear_clp(monto)))

def actualizar_lista_historica():
    lista_historial.delete(0, tk.END)
    for m in obtener_meses_historicos():
        lista_historial.insert(tk.END, f" 📁 {m}")

def actualizar_pantalla():
    datos = obtener_todos_los_saldos()
    total_ingresos = 0
    total_gastos = 0
    gastos_filtrados = {}
    for item, monto, tipo in datos:
        if tipo == "Ingreso":
            total_ingresos += monto
        else:
            total_gastos += monto
            if monto > 0:
                gastos_filtrados[item] = monto
    saldo_disponible = total_ingresos - total_gastos
    lbl_total_ingresos.config(text=f"Total Ingresos: {formatear_clp(total_ingresos)}")
    lbl_total_gastos.config(text=f"Total Gastos: {formatear_clp(total_gastos)}")
    lbl_sobra_falta.config(text=f"Balance Neto: {formatear_clp(saldo_disponible)}", fg="#2ECC71" if saldo_disponible >= 0 else "#E74C3C")
    ax1.clear()
    ax2.clear()
    ax1.set_facecolor("#282C34")
    ax2.set_facecolor("#282C34")
    barras = ax1.bar(["Ingresos", "Gastos"], [total_ingresos, total_gastos], color=["#4A90E2", "#E74C3C"], width=0.35, edgecolor="#3A3F47")
    ax1.set_title("Balance Financiero del Periodo", color="#E1E4E6", fontsize=10, fontweight="bold", family="Segoe UI", pad=12)
    ax1.tick_params(colors="#A0A5AD", labelsize=8)
    ax1.spines['bottom'].set_color('#3A3F47')
    ax1.spines['left'].set_color('#3A3F47')
    ax1.spines['top'].set_visible(False)
    ax1.spines['right'].set_visible(False)
    for bar in barras:
        yval = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2, yval + (max(total_ingresos, total_gastos) * 0.02 if max(total_ingresos, total_gastos) > 0 else 1), formatear_clp(yval), ha='center', va='bottom', color='#E1E4E6', fontsize=8, fontweight="bold", family="Segoe UI")
    if gastos_filtrados:
        colores_ejecutivos = ['#4A90E2', '#357ABD', '#225E9B', '#1B4780', '#5C9CE6', '#7FB0EB']
        ax2.pie(gastos_filtrados.values(), labels=gastos_filtrados.keys(), autopct='%1.1f%%', startangle=90, textprops=dict(color="#E1E4E6", fontsize=8, family="Segoe UI"), colors=colores_ejecutivos, wedgeprops={"edgecolor":"#282C34", 'linewidth': 1, 'antialiased': True})
        ax2.set_title("Distribución Porcentual de Egresos", color="#E1E4E6", fontsize=10, fontweight="bold", family="Segoe UI", pad=12)
    else:
        ax2.text(0.5, 0.5, 'Sin transacciones de egreso', ha='center', va='center', color='#8A9099', fontsize=9, family="Segoe UI")
        ax2.set_title("Distribución Porcentual de Egresos", color="#E1E4E6", fontsize=10, fontweight="bold", family="Segoe UI", pad=12)
    canvas.draw()
def reconstruir_lista_gastos_visual():
    global todas_las_entradas
    for widget in scroll_frame.winfo_children():
        widget.destroy()
    datos = obtener_todos_los_saldos()
    nombre_ingreso = next((row for row in datos if row == "Ingreso"), "Ingresos Principales")
    todas_las_entradas = {nombre_ingreso: ent_monto_ingreso}
    for item, valor, tipo in datos:
        if tipo == "Gasto":
            f_g = tk.Frame(scroll_frame, bg="#282C34", padx=6, pady=3)
            f_g.pack(fill=tk.X, pady=3, padx=4)
            if modo_borrar_activo:
                btn_del = tk.Button(f_g, text="✕", bg="#E74C3C", fg="#FFFFFF", font=("Segoe UI", 8, "bold"), relief="flat", activebackground="#C0392B", activeforeground="white", cursor="hand2", bd=0, width=2, command=lambda i=item: borrar_gasto_interfaz(i))
                btn_del.pack(side=tk.LEFT, padx=(0, 6))
            lbl = tk.Label(f_g, text=item, bg="#282C34", fg="#E1E4E6", font=("Segoe UI", 9))
            lbl.pack(side=tk.LEFT, padx=2 if not modo_borrar_activo else 0)
            ent_gasto = tk.Entry(f_g, **in_opts)
            ent_gasto.pack(side=tk.RIGHT, padx=2)
            todas_las_entradas[item] = ent_gasto
            if valor > 0:
                ent_gasto.insert(0, str(int(valor)))
                class EventDummy: widget = ent_gasto
                formatear_entrada_dinamica(EventDummy)
            ent_gasto.bind("<KeyRelease>", formatear_entrada_dinamica)
            ent_gasto.bind("<FocusOut>", lambda e, i=item, w=ent_gasto: registrar_valor(i, w))
            ent_gasto.bind("<Return>", lambda e, i=item, w=ent_gasto: registrar_valor(i, w))

inicializar_db()
ventana = tk.Tk()
ventana.title("Control Financiero de Cuentas Corporativo")
ventana.geometry("1300x760")
ventana.resizable(False, False)
ventana.configure(bg="#1E2229")

lbl_opts = {"bg": "#1E2229", "fg": "#E1E4E6", "font": ("Segoe UI", 10, "bold")}
in_opts = {"bg": "#1E2229", "fg": "#E1E4E6", "insertbackground": "#E1E4E6", "relief": "solid", "highlightcolor": "#4A90E2", "bd": 1, "font": ("Segoe UI", 10), "width": 14, "justify": "right"}
todas_las_entradas = {}
frame_interfaz_completa = tk.Frame(ventana, bg="#1E2229", bd=0)
frame_interfaz_completa.pack(fill=tk.BOTH, expand=True)

frame_izquierda = tk.Frame(frame_interfaz_completa, bg="#1E2229", padx=15, pady=15, width=430)
frame_izquierda.pack(side=tk.LEFT, fill=tk.Y, expand=False, padx=15, pady=15)
frame_izquierda.pack_propagate(False)

frame_mes_selector = tk.Frame(frame_izquierda, bg="#1E2229", pady=5)
frame_mes_selector.pack(fill=tk.X, pady=(0, 5))
tk.Label(frame_mes_selector, text="Periodo Activo:", **lbl_opts).pack(side=tk.LEFT, padx=5)

estilo_combo = ttk.Style()
estilo_combo.theme_use('clam')
estilo_combo.configure("TCombobox", fieldbackground="#FFFFFF", background="#FFFFFF", foreground="#000000", darkcolor="#FFFFFF", lightcolor="#FFFFFF", insertcolor="#000000")
ventana.option_add("*TCombobox*Listbox.background", "#FFFFFF")
ventana.option_add("*TCombobox*Listbox.foreground", "#000000")

combo_mes = ttk.Combobox(frame_mes_selector, values=["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"], width=11, state="readonly")
combo_mes.set("Septiembre")
combo_mes.pack(side=tk.LEFT, padx=5)

combo_ano = ttk.Combobox(frame_mes_selector, values=["2025", "2026", "2027", "2028"], width=6, state="readonly")
combo_ano.set("2026")
combo_ano.pack(side=tk.LEFT, padx=5)

frame_ingresos_caja = tk.LabelFrame(frame_izquierda, text=" Ingresos Consolidados ", bg="#282C34", fg="#4A90E2", font=("Segoe UI", 9, "bold"), padx=12, pady=10, bd=1, relief="solid")
frame_ingresos_caja.pack(fill=tk.X, pady=5)

# Fila superior de identificación del usuario de la App
f_usuario = tk.Frame(frame_ingresos_caja, bg="#282C34")
f_usuario.pack(fill=tk.X, pady=(2, 6))
tk.Label(f_usuario, text="Usuario App:", bg="#282C34", fg="#8A9099", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=2)
ent_usuario_app = tk.Entry(f_usuario, bg="#1E2229", fg="#FFFFFF", insertbackground="#FFFFFF", relief="solid", bd=1, font=("Segoe UI", 9), width=18)
ent_usuario_app.pack(side=tk.RIGHT, padx=2)
ent_usuario_app.insert(0, "Francisco")

f_ingreso_dinamico = tk.Frame(frame_ingresos_caja, bg="#282C34")
f_ingreso_dinamico.pack(fill=tk.X, pady=4)

ent_nombre_ingreso = tk.Entry(f_ingreso_dinamico, bg="#282C34", fg="#4A90E2", insertbackground="#E1E4E6", relief="flat", font=("Segoe UI", 9, "bold"), width=16)
ent_nombre_ingreso.pack(side=tk.LEFT, padx=2)
ent_nombre_ingreso.bind("<FocusOut>", lambda e: registrar_cambio_nombre_ingreso())
ent_nombre_ingreso.bind("<Return>", lambda e: registrar_cambio_nombre_ingreso())

ent_monto_ingreso = tk.Entry(f_ingreso_dinamico, **in_opts)
ent_monto_ingreso.pack(side=tk.RIGHT, padx=2)

saldos_db_actuales = obtener_todos_los_saldos()
nombre_ingreso_inicial = next((row for row in saldos_db_actuales if row == "Ingreso"), "Ingresos Principales")
monto_ingreso_inicial = next((row for row in saldos_db_actuales if row == "Ingreso"), 0.0)

ent_nombre_ingreso.insert(0, nombre_ingreso_inicial)
if monto_ingreso_inicial > 0:
    ent_monto_ingreso.insert(0, str(int(monto_ingreso_inicial)))
    class EventDummy: widget = ent_monto_ingreso
    formatear_entrada_dinamica(EventDummy)

ent_monto_ingreso.bind("<KeyRelease>", formatear_entrada_dinamica)
ent_monto_ingreso.bind("<FocusOut>", lambda e: registrar_valor(ent_nombre_ingreso.get().strip(), ent_monto_ingreso))
ent_monto_ingreso.bind("<Return>", lambda e: registrar_valor(ent_nombre_ingreso.get().strip(), ent_monto_ingreso))

lbl_total_ingresos = tk.Label(frame_ingresos_caja, text="Total Ingresos: $0", bg="#282C34", fg="#2ECC71", font=("Segoe UI", 10, "bold"))
lbl_total_ingresos.pack(anchor="w", pady=(6, 2))

# Bloque limpio de control de egresos
frame_gastos_caja = tk.LabelFrame(frame_izquierda, text=" Controles Operativos de Gastos ", bg="#282C34", fg="#A0A5AD", font=("Segoe UI", 9, "bold"), padx=12, pady=10, bd=1, relief="solid")
frame_gastos_caja.pack(fill=tk.BOTH, expand=True, pady=5)

frame_creador = tk.Frame(frame_gastos_caja, bg="#282C34", pady=2)
frame_creador.pack(fill=tk.X)
entry_crear_gasto = tk.Entry(frame_creador, bg="#1E2229", fg="#E1E4E6", insertbackground="#E1E4E6", relief="solid", font=("Segoe UI", 9), width=18, bd=1)
entry_crear_gasto.pack(side=tk.LEFT, padx=(0, 6))

btn_add = tk.Button(frame_creador, text="Añadir", bg="#4A90E2", fg="#FFFFFF", font=("Segoe UI", 8, "bold"), relief="flat", padx=6, cursor="hand2", command=añadir_nuevo_gasto_interfaz)
btn_add.pack(side=tk.LEFT, padx=(0, 4))

btn_basurero = tk.Button(frame_creador, text="🗑️ Editar", bg="#2E3238", fg="#E74C3C", font=("Segoe UI", 8, "bold"), relief="flat", padx=6, cursor="hand2", command=alternar_modo_borrado)
btn_basurero.pack(side=tk.LEFT)

canvas_scroll = tk.Canvas(frame_gastos_caja, bg="#282C34", bd=0, highlightthickness=0)
scrollbar = ttk.Scrollbar(frame_gastos_caja, orient="vertical", command=canvas_scroll.yview)
scroll_frame = tk.Frame(canvas_scroll, bg="#282C34")

scroll_frame.bind("<Configure>", lambda e: canvas_scroll.configure(scrollregion=canvas_scroll.bbox("all")))
canvas_scroll.create_window((0, 0), window=scroll_frame, anchor="nw", width=365)
canvas_scroll.configure(yscrollcommand=scrollbar.set)
canvas_scroll.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=4)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

def desplazar_gastos_mouse(event):
    canvas_scroll.yview_scroll(int(-1 * (event.delta / 120)), "units")
canvas_scroll.bind_all("<MouseWheel>", desplazar_gastos_mouse)

lbl_total_gastos = tk.Label(frame_gastos_caja, text="Total Gastos: $0", bg="#282C34", fg="#E74C3C", font=("Segoe UI", 10, "bold"))
lbl_total_gastos.pack(side=tk.BOTTOM, anchor="w", pady=(2, 0))

frame_fijo_inferior = tk.Frame(frame_izquierda, bg="#1E2229")
frame_fijo_inferior.pack(fill=tk.X, side=tk.BOTTOM, pady=(5, 0))

frame_balance = tk.Frame(frame_fijo_inferior, bg="#282C34", padx=12, pady=10, bd=1, relief="solid")
frame_balance.pack(fill=tk.X, pady=(0, 5))
lbl_sobra_falta = tk.Label(frame_balance, text="Balance Neto: $0", bg="#282C34", fg="#E1E4E6", font=("Segoe UI", 10, "bold"))
lbl_sobra_falta.pack(side=tk.LEFT)

btn_guardar_mes = tk.Button(frame_fijo_inferior, text="Archivar Registro del Periodo", command=guardar_registro_mes, bg="#2E3238", fg="#4A90E2", font=("Segoe UI", 9, "bold"), relief="solid", bd=1, activebackground="#4A90E2", activeforeground="white", pady=8, cursor="hand2")
btn_guardar_mes.pack(fill=tk.X)

frame_centro = tk.LabelFrame(frame_interfaz_completa, text=" Cuadro de Mando de Analíticas ", bg="#282C34", fg="#E1E4E6", font=("Segoe UI", 9, "bold"), padx=12, pady=12, bd=1, relief="solid")
frame_centro.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=15)

fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(5, 6))
fig.patch.set_facecolor('#282C34')
fig.tight_layout(pad=4.0)

canvas = FigureCanvasTkAgg(fig, master=frame_centro)
canvas.get_tk_widget().config(bg="#282C34", bd=0, highlightthickness=0)
canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

frame_derecha = tk.LabelFrame(frame_interfaz_completa, text=" Periodos Consolidados ", bg="#282C34", fg="#E1E4E6", font=("Segoe UI", 9, "bold"), padx=12, pady=12, bd=1, relief="solid")
frame_derecha.pack(side=tk.RIGHT, fill=tk.Y, padx=15, pady=15)

tk.Label(frame_derecha, text="Doble clic para auditar periodo:", bg="#282C34", fg="#8A9099", font=("Segoe UI", 8, "italic")).pack(anchor="w", pady=(0, 6))

lista_historial = tk.Listbox(frame_derecha, bg="#1E2229", fg="#E1E4E6", font=("Segoe UI", 9, "bold"), width=24, relief="flat", highlightthickness=0, selectbackground="#4A90E2", selectforeground="white")
lista_historial.pack(fill=tk.BOTH, expand=True, pady=5)
lista_historial.bind("<Double-1>", revisar_mes_guardado)

reconstruir_lista_gastos_visual()
actualizar_pantalla()
actualizar_lista_historica()
ventana.mainloop()

