import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry  
from docxtpl import DocxTemplate
from docxcompose.composer import Composer
from docx import Document as DocumentoWord
from num2words import num2words
from datetime import datetime
from dateutil.relativedelta import relativedelta
import locale
import os
import platform

# Configuración regional
try:
    locale.setlocale(locale.LC_ALL, 'es_PY.UTF-8')
except:
    locale.setlocale(locale.LC_ALL, '')

# --- VARIABLES GLOBALES PARA CUOTAS PERSONALIZADAS ---
cuotas_custom_data = {} # Diccionario: {numero_cuota: monto}

# --- FUNCIONES ---

def aplicar_formato_miles(event):
    widget = event.widget
    valor_actual = widget.get().replace('.', '') 
    if valor_actual.isdigit() and valor_actual:
        valor_formateado = "{:,.0f}".format(int(valor_actual)).replace(",", ".")
        widget.delete(0, tk.END)
        widget.insert(0, valor_formateado)
    elif valor_actual == "":
        pass

def toggle_codeudor():
    if var_tiene_codeudor.get():
        frame_codeudor.pack(fill="x", padx=15, pady=5, after=frame_check)
    else:
        frame_codeudor.pack_forget()

def limpiar_todo():
    # Limpia campos de texto
    campos_texto = [
        entry_acreedor_nombre, entry_acreedor_nac, entry_acreedor_ci, entry_acreedor_dom,
        entry_deudor_nombre, entry_deudor_ci, entry_deudor_dom,
        entry_cod_nombre, entry_cod_ci, entry_cod_dom,
        entry_monto, entry_cuotas 
    ]
    for campo in campos_texto:
        campo.delete(0, tk.END)
    
    combo_acreedor_sexo.current(0)
    combo_acreedor_est.current(0)
    combo_deudor_sexo.current(0)
    var_tiene_codeudor.set(False)
    toggle_codeudor()
    
    # Limpiar datos custom
    cuotas_custom_data.clear()
    lbl_status_custom.config(text="Ninguna cuota personalizada.", foreground="gray")
    
    entry_acreedor_nombre.focus()

def abrir_plantilla_word():
    archivo = "plantilla_pagare.docx"
    if os.path.exists(archivo):
        os.startfile(archivo)
    else:
        messagebox.showerror("Error", "No encuentro 'plantilla_pagare.docx'")

# --- GESTOR DE CUOTAS (VENTANA EMERGENTE) ---
def abrir_gestor_cuotas():
    # Validar que haya cuotas definidas primero
    str_cuotas = entry_cuotas.get()
    if not str_cuotas.isdigit() or int(str_cuotas) < 1:
        messagebox.showwarning("Atención", "Primero ingrese la 'Cant. Cuotas' en el formulario principal.")
        return
    
    total_cuotas = int(str_cuotas)
    
    # Crear ventana emergente
    top = tk.Toplevel(root)
    top.title("Administrar Cuotas Diferentes")
    top.geometry("400x450")
    top.grab_set() # Bloquea la ventana principal hasta cerrar esta

    # Frame entrada
    f_input = ttk.Frame(top)
    f_input.pack(pady=10)

    ttk.Label(f_input, text="N° Cuota:").pack(side="left", padx=5)
    combo_n = ttk.Combobox(f_input, values=[str(i) for i in range(1, total_cuotas+1)], width=5, state="readonly")
    combo_n.pack(side="left", padx=5)
    if total_cuotas > 0: combo_n.current(0)

    ttk.Label(f_input, text="Monto:").pack(side="left", padx=5)
    ent_m = ttk.Entry(f_input, width=12)
    ent_m.pack(side="left", padx=5)
    ent_m.bind("<FocusOut>", aplicar_formato_miles)
    ent_m.bind("<Return>", aplicar_formato_miles)

    # Lista (Treeview)
    columns = ("cuota", "monto")
    tree = ttk.Treeview(top, columns=columns, show="headings", height=10)
    tree.heading("cuota", text="Cuota N°")
    tree.heading("monto", text="Monto Diferente (Gs)")
    tree.column("cuota", width=80, anchor="center")
    tree.column("monto", width=150, anchor="e")
    tree.pack(pady=10, padx=10, fill="both", expand=True)

    def refrescar_lista():
        # Limpiar
        for item in tree.get_children():
            tree.delete(item)
        # Llenar
        # Ordenamos por numero de cuota
        for k in sorted(cuotas_custom_data.keys()):
            v = cuotas_custom_data[k]
            v_fmt = "{:,.0f}".format(v).replace(",", ".")
            tree.insert("", "end", values=(str(k), v_fmt))
        
        # Actualizar label ventana principal
        cant = len(cuotas_custom_data)
        if cant > 0:
            lbl_status_custom.config(text=f"✅ {cant} cuotas personalizadas configuradas.", foreground="green")
        else:
            lbl_status_custom.config(text="Ninguna cuota personalizada.", foreground="gray")

    def agregar():
        n = combo_n.get()
        m_str = ent_m.get().replace('.', '')
        
        if not m_str.isdigit(): return
        m = int(m_str)
        if m <= 0: return
        
        idx = int(n)
        cuotas_custom_data[idx] = m
        refrescar_lista()
        ent_m.delete(0, tk.END)

    def borrar():
        selected = tree.selection()
        if not selected: return
        item = tree.item(selected[0])
        cuota_idx = int(item['values'][0])
        del cuotas_custom_data[cuota_idx]
        refrescar_lista()

    btn_add = ttk.Button(f_input, text="Agregar", command=agregar)
    btn_add.pack(side="left", padx=5)

    btn_del = ttk.Button(top, text="Borrar Seleccionado", command=borrar)
    btn_del.pack(pady=5)
    
    ttk.Button(top, text="Cerrar / Guardar", command=top.destroy).pack(pady=10)

    # Cargar datos existentes si reabre la ventana
    refrescar_lista()


def generar_documento_unico():
    try:
        # --- 1. DATOS ---
        acreedor_data = {
            'nombre': entry_acreedor_nombre.get().upper(),
            'nac': entry_acreedor_nac.get(),
            'est': combo_acreedor_est.get(),
            'ci': entry_acreedor_ci.get(),
            'dom': entry_acreedor_dom.get(),
            'sexo': combo_acreedor_sexo.get()
        }
        deudor_data = {
            'nombre': entry_deudor_nombre.get().upper(),
            'ci': entry_deudor_ci.get(),
            'dom': entry_deudor_dom.get(),
            'sexo': combo_deudor_sexo.get()
        }
        usa_codeudor = var_tiene_codeudor.get()
        cod_data = {'nombre': '', 'ci': '', 'dom': ''}
        if usa_codeudor:
            cod_data['nombre'] = entry_cod_nombre.get().upper()
            cod_data['ci'] = entry_cod_ci.get()
            cod_data['dom'] = entry_cod_dom.get()
            if not cod_data['nombre']:
                messagebox.showwarning("Falta dato", "Falta nombre del Codeudor.")
                return

        # --- 2. CÁLCULOS ---
        m_total_str = entry_monto.get().replace('.', '')
        if not m_total_str.isdigit(): messagebox.showerror("Error", "Monto Total inválido"); return
        monto_total = int(m_total_str)

        cuotas_str = entry_cuotas.get()
        if not cuotas_str.isdigit(): messagebox.showerror("Error", "Cant. Cuotas inválida"); return
        cantidad_cuotas = int(cuotas_str)

        # Lógica Multicuotas
        suma_asignada = sum(cuotas_custom_data.values())
        
        # Validar que las cuotas custom no superen el total de cuotas
        for k in cuotas_custom_data.keys():
            if k > cantidad_cuotas:
                messagebox.showerror("Error", f"Tienes configurada la cuota {k}, pero el total de cuotas es {cantidad_cuotas}.")
                return

        if suma_asignada > monto_total:
            messagebox.showerror("Error Matemático", "La suma de las cuotas especiales supera el Monto Total.")
            return

        indices_ocupados = len(cuotas_custom_data)
        cuotas_restantes = cantidad_cuotas - indices_ocupados
        saldo_restante_para_regulares = monto_total - suma_asignada
        
        cuota_regular = 0
        if cuotas_restantes > 0:
            cuota_regular = saldo_restante_para_regulares / cuotas_restantes
        elif cuotas_restantes == 0 and saldo_restante_para_regulares != 0:
             messagebox.showwarning("Atención", "Has personalizado TODAS las cuotas, pero los montos no suman exactamente el Monto Total. Revisa tus números.")
             # Continuamos igual, asumiendo que el usuario sabe lo que hace

        # --- 3. CONFIG ---
        fecha_inicio_obj = entry_fecha.get_date()
        frecuencia = combo_frecuencia.get()
        moneda_sel = combo_moneda.get()
        
        if "Guaraníes" in moneda_sel: m_s, m_p, m_sg = "GS.", "GUARANIES", "GUARANI"
        elif "Dólares" in moneda_sel: m_s, m_p, m_sg = "USD", "DOLARES AMERICANOS", "DOLAR AMERICANO"
        elif "Reales" in moneda_sel: m_s, m_p, m_sg = "R$", "REALES", "REAL"
        elif "Euros" in moneda_sel: m_s, m_p, m_sg = "€", "EUROS", "EURO"
        elif "Pesos Arg" in moneda_sel: m_s, m_p, m_sg = "ARS", "PESOS ARGENTINOS", "PESO ARGENTINO"
        else: m_s, m_p, m_sg = "$", "MONEDA", "MONEDA"

        acr_tit = "del señor" if acreedor_data['sexo'] == "Masculino" else "de la señora"
        acr_dom = "domiciliado" if acreedor_data['sexo'] == "Masculino" else "domiciliada"
        deu_dom = "domiciliado" if deudor_data['sexo'] == "Masculino" else "domiciliada"

        plantilla_path = "plantilla_pagare.docx"
        if not os.path.exists(plantilla_path):
            messagebox.showerror("Error", "No encuentro 'plantilla_pagare.docx'")
            return

        # --- 4. GENERACIÓN ---
        archivo_maestro = None
        composer = None
        temp_files = []
        fecha_venc = fecha_inicio_obj

        for i in range(1, cantidad_cuotas + 1):
            
            monto_actual = 0
            
            # VERIFICAR SI ESTA CUOTA ES PERSONALIZADA
            if i in cuotas_custom_data:
                monto_actual = cuotas_custom_data[i]
            else:
                monto_actual = cuota_regular
            
            doc = DocxTemplate(plantilla_path)
            txt_mon = m_sg if monto_actual == 1 else m_p

            context = {
                'acreedor_nombre': acreedor_data['nombre'],
                'acreedor_nacionalidad': acreedor_data['nac'],
                'acreedor_estado': acreedor_data['est'],
                'acreedor_ci': acreedor_data['ci'],
                'acreedor_dom': acreedor_data['dom'],
                'acreedor_titulo': acr_tit,       
                'acreedor_dom_texto': acr_dom,
                'deudor_nombre': deudor_data['nombre'],
                'deudor_ci': deudor_data['ci'],
                'deudor_dom': deudor_data['dom'],
                'deudor_dom_texto': deu_dom,
                'hay_codeudor': usa_codeudor,
                'codeudor_nombre': cod_data['nombre'],
                'codeudor_ci': cod_data['ci'],
                'codeudor_dom': cod_data['dom'],
                'cuota_actual': f"{i:02d}",
                'cuota_total': f"{cantidad_cuotas:02d}",
                'moneda_simbolo': m_s,
                'moneda_texto': txt_mon,
                'monto_num': "{:,.0f}".format(monto_actual).replace(",", "."),
                'monto_letras': num2words(monto_actual, lang='es').upper(),
                'fecha_venc': fecha_venc.strftime("%d/%m/%Y"),
                'fecha_emision': datetime.now().strftime("%d/%m/%Y"),
                'lugar': "Bella Vista"
            }
            
            doc.render(context)
            tmp = f"temp_{i}.docx"
            doc.save(tmp)
            temp_files.append(tmp)

            if i == 1:
                archivo_maestro = DocumentoWord(tmp)
                composer = Composer(archivo_maestro)
            else:
                doc_s = DocumentoWord(tmp)
                archivo_maestro.add_page_break() 
                composer.append(doc_s)

            if frecuencia == 'Mensual': fecha_venc += relativedelta(months=1)
            elif frecuencia == 'Bimestral': fecha_venc += relativedelta(months=2)
            elif frecuencia == 'Trimestral': fecha_venc += relativedelta(months=3)
            elif frecuencia == 'Cuatrimestral': fecha_venc += relativedelta(months=4)
            elif frecuencia == 'Semestral': fecha_venc += relativedelta(months=6)
            elif frecuencia == 'Anual': fecha_venc += relativedelta(years=1)

        nombre_file = f"Pagares_{deudor_data['nombre'].replace(' ', '_')}.docx"
        path_guardado = filedialog.asksaveasfilename(initialfile=nombre_file, defaultextension=".docx", filetypes=[("Word", "*.docx")])

        if path_guardado:
            composer.save(path_guardado)
            for f in temp_files:
                try: os.remove(f)
                except: pass
            if messagebox.askyesno("Listo", "Documento creado.\n¿Abrir ahora?"):
                os.startfile(path_guardado)

    except Exception as e:
        messagebox.showerror("Error Crítico", f"Detalle: {str(e)}")

# --- GUI ---
root = tk.Tk()
root.title("Gestor de Pagarés v10 - Multi Cuotas")
root.geometry("650x900")
root.configure(bg="#ececec")

style = ttk.Style()
style.theme_use('clam')
style.configure("TLabel", background="#ececec", font=("Segoe UI", 9))
style.configure("TLabelframe", background="#ececec", relief="solid", borderwidth=1)
style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"), background="#ececec", foreground="#333")
style.configure("TButton", font=("Segoe UI", 10, "bold"))
style.configure("TCheckbutton", background="#ececec", font=("Segoe UI", 10))

LISTA_CIVIL = ["Soltero", "Soltera", "Casado", "Casada", "Divorciado", "Divorciada", "Viudo", "Viuda"]
LISTA_MONEDAS = ["Guaraníes (PYG)", "Dólares (USD)", "Reales (BRL)", "Euros (EUR)", "Pesos Arg. (ARS)"]
LISTA_FRECUENCIAS = ["Mensual", "Bimestral", "Trimestral", "Cuatrimestral", "Semestral", "Anual"]

# 1. ACREEDOR
frame_acreedor = ttk.LabelFrame(root, text=" 1. Acreedor (Quien cobra) ")
frame_acreedor.pack(fill="x", padx=15, pady=5)
ttk.Label(frame_acreedor, text="Nombre:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
entry_acreedor_nombre = ttk.Entry(frame_acreedor, width=35)
entry_acreedor_nombre.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="w")
ttk.Label(frame_acreedor, text="Sexo:").grid(row=0, column=3, padx=5, pady=5, sticky="e")
combo_acreedor_sexo = ttk.Combobox(frame_acreedor, values=["Femenino", "Masculino"], state="readonly", width=10)
combo_acreedor_sexo.current(0)
combo_acreedor_sexo.grid(row=0, column=4, padx=5, pady=5, sticky="w")
ttk.Label(frame_acreedor, text="Nacionalidad:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
entry_acreedor_nac = ttk.Entry(frame_acreedor, width=15)
entry_acreedor_nac.grid(row=1, column=1, padx=5, pady=5, sticky="w")
ttk.Label(frame_acreedor, text="Estado Civil:").grid(row=1, column=2, padx=5, pady=5, sticky="e")
combo_acreedor_est = ttk.Combobox(frame_acreedor, values=LISTA_CIVIL, state="readonly", width=13)
combo_acreedor_est.current(0)
combo_acreedor_est.grid(row=1, column=3, columnspan=2, padx=5, pady=5, sticky="w")
ttk.Label(frame_acreedor, text="C.I. / RUC:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
entry_acreedor_ci = ttk.Entry(frame_acreedor, width=20)
entry_acreedor_ci.grid(row=2, column=1, columnspan=4, padx=5, pady=5, sticky="w")
ttk.Label(frame_acreedor, text="Domicilio:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
entry_acreedor_dom = ttk.Entry(frame_acreedor, width=45)
entry_acreedor_dom.grid(row=3, column=1, columnspan=4, padx=5, pady=5, sticky="w")

# 2. DEUDOR
frame_deudor = ttk.LabelFrame(root, text=" 2. Deudor (Quien firma) ")
frame_deudor.pack(fill="x", padx=15, pady=5)
ttk.Label(frame_deudor, text="Nombre:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
entry_deudor_nombre = ttk.Entry(frame_deudor, width=35)
entry_deudor_nombre.grid(row=0, column=1, columnspan=2, padx=5, pady=5)
ttk.Label(frame_deudor, text="Sexo:").grid(row=0, column=3, padx=5, pady=5, sticky="e")
combo_deudor_sexo = ttk.Combobox(frame_deudor, values=["Masculino", "Femenino"], state="readonly", width=10)
combo_deudor_sexo.current(0)
combo_deudor_sexo.grid(row=0, column=4, padx=5, pady=5, sticky="w")
ttk.Label(frame_deudor, text="C.I. N°:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
entry_deudor_ci = ttk.Entry(frame_deudor, width=20)
entry_deudor_ci.grid(row=1, column=1, columnspan=4, padx=5, pady=5, sticky="w")
ttk.Label(frame_deudor, text="Domicilio:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
entry_deudor_dom = ttk.Entry(frame_deudor, width=45)
entry_deudor_dom.grid(row=2, column=1, columnspan=4, padx=5, pady=5)

# CHECKBOX CODEUDOR
frame_check = tk.Frame(root, bg="#ececec")
frame_check.pack(fill="x", padx=15, pady=5)
var_tiene_codeudor = tk.BooleanVar(value=False) 
chk_codeudor = ttk.Checkbutton(frame_check, text="¿Tiene Codeudor / Garante?", variable=var_tiene_codeudor, 
                               command=toggle_codeudor, style="TCheckbutton")
chk_codeudor.pack(anchor="w", padx=5)

# CODEUDOR
frame_codeudor = ttk.LabelFrame(root, text=" Datos del Codeudor (Garante) ")
ttk.Label(frame_codeudor, text="Nombre:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
entry_cod_nombre = ttk.Entry(frame_codeudor, width=45)
entry_cod_nombre.grid(row=0, column=1, columnspan=3, padx=5, pady=5)
ttk.Label(frame_codeudor, text="C.I. N°:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
entry_cod_ci = ttk.Entry(frame_codeudor, width=20)
entry_cod_ci.grid(row=1, column=1, sticky="w", padx=5, pady=5)
ttk.Label(frame_codeudor, text="Domicilio:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
entry_cod_dom = ttk.Entry(frame_codeudor, width=45)
entry_cod_dom.grid(row=2, column=1, columnspan=3, padx=5, pady=5)

# 3. PRÉSTAMO
frame_prestamo = ttk.LabelFrame(root, text=" 3. Datos del Préstamo ")
frame_prestamo.pack(fill="x", padx=15, pady=5)
ttk.Label(frame_prestamo, text="Moneda:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
combo_moneda = ttk.Combobox(frame_prestamo, values=LISTA_MONEDAS, state="readonly", width=20)
combo_moneda.current(0)
combo_moneda.grid(row=0, column=1, padx=5, pady=5, sticky="w")

ttk.Label(frame_prestamo, text="Monto Total:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
entry_monto = ttk.Entry(frame_prestamo, width=20)
entry_monto.grid(row=1, column=1, padx=5, pady=5, sticky="w")
entry_monto.bind("<FocusOut>", aplicar_formato_miles)
entry_monto.bind("<Return>", aplicar_formato_miles)

ttk.Label(frame_prestamo, text="Cant. Cuotas:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
entry_cuotas = ttk.Entry(frame_prestamo, width=10)
entry_cuotas.grid(row=2, column=1, padx=5, pady=5, sticky="w")

ttk.Label(frame_prestamo, text="1er Vencim.:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
entry_fecha = DateEntry(frame_prestamo, width=12, background='#444444', foreground='white', borderwidth=2, date_pattern='dd/mm/y', locale='es_PY', headersbackground='#222222')
entry_fecha.grid(row=3, column=1, padx=5, pady=5, sticky="w")
ttk.Label(frame_prestamo, text="Frecuencia:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
combo_frecuencia = ttk.Combobox(frame_prestamo, values=LISTA_FRECUENCIAS, state="readonly", width=10)
combo_frecuencia.current(0)
combo_frecuencia.grid(row=4, column=1, padx=5, pady=5, sticky="w")

# 4. PERSONALIZACIÓN (GESTOR AVANZADO)
frame_custom = ttk.LabelFrame(root, text=" 4. Configuración Avanzada ")
frame_custom.pack(fill="x", padx=15, pady=5)

btn_gestionar_cuotas = ttk.Button(frame_custom, text="CONFIGURAR CUOTAS DISTINTAS", command=abrir_gestor_cuotas)
btn_gestionar_cuotas.pack(side="left", padx=10, pady=5)

lbl_status_custom = ttk.Label(frame_custom, text="Ninguna cuota personalizada.", foreground="gray", font=("Segoe UI", 9))
lbl_status_custom.pack(side="left", padx=5)


# BOTONERA FINAL
frame_botones = tk.Frame(root, bg="#ececec")
frame_botones.pack(pady=15, fill="x", padx=40)

btn_limpiar = tk.Button(frame_botones, text="LIMPIAR TODO", command=limpiar_todo, bg="#777", fg="white", font=("Segoe UI", 9, "bold"), width=15, height=2)
btn_limpiar.pack(side="left", padx=5)

btn_plantilla = tk.Button(frame_botones, text="EDITAR PLANTILLA", command=abrir_plantilla_word, bg="#005a9e", fg="white", font=("Segoe UI", 9, "bold"), width=18, height=2)
btn_plantilla.pack(side="left", padx=5)

btn_generar = tk.Button(frame_botones, text="GENERAR PAGARES", command=generar_documento_unico, bg="#28a745", fg="white", font=("Segoe UI", 11, "bold"), height=2)
btn_generar.pack(side="left", padx=5, fill="x", expand=True)

root.mainloop()