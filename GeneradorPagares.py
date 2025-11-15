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

def formatear_campo(variable, entry_widget):
    valor_raw = variable.get().replace('.', '')
    if valor_raw.isdigit() and valor_raw:
        valor_fmt = "{:,.0f}".format(int(valor_raw)).replace(",", ".")
        if variable.get() != valor_fmt:
            variable.set(valor_fmt)
            entry_widget.icursor(tk.END)
    elif valor_raw == "":
        pass
    else:
        variable.set(valor_raw[:-1])

def toggle_codeudor():
    if var_tiene_codeudor.get():
        frame_codeudor.pack(fill="x", padx=15, pady=5, after=frame_check)
    else:
        frame_codeudor.pack_forget()

def limpiar_todo():
    entry_acreedor_nombre.delete(0, tk.END)
    entry_acreedor_nac.delete(0, tk.END)
    entry_acreedor_ci.delete(0, tk.END)
    entry_acreedor_dom.delete(0, tk.END)
    combo_acreedor_sexo.current(0)
    combo_acreedor_est.current(0)

    entry_deudor_nombre.delete(0, tk.END)
    entry_deudor_ci.delete(0, tk.END)
    entry_deudor_dom.delete(0, tk.END)
    combo_deudor_sexo.current(0)
    
    var_tiene_codeudor.set(False)
    toggle_codeudor()
    entry_cod_nombre.delete(0, tk.END)
    entry_cod_ci.delete(0, tk.END)
    entry_cod_dom.delete(0, tk.END)
    
    monto_var.set("")
    entry_cuotas.delete(0, tk.END)
    monto_primera_var.set("")
    monto_ultima_var.set("")
    
    entry_acreedor_nombre.focus()

def abrir_plantilla_word():
    archivo = "plantilla_pagare.docx"
    if os.path.exists(archivo):
        os.startfile(archivo)
    else:
        messagebox.showerror("Error", "No encuentro el archivo 'plantilla_pagare.docx' en esta carpeta.")

def generar_documento_unico():
    try:
        # --- 1. DATOS ---
        acreedor_nombre = entry_acreedor_nombre.get().upper()
        acreedor_nac = entry_acreedor_nac.get()
        acreedor_est = combo_acreedor_est.get()
        acreedor_ci = entry_acreedor_ci.get()
        acreedor_dom = entry_acreedor_dom.get()
        acreedor_sexo = combo_acreedor_sexo.get()

        deudor_nombre = entry_deudor_nombre.get().upper()
        deudor_ci = entry_deudor_ci.get()
        deudor_dom = entry_deudor_dom.get()
        deudor_sexo = combo_deudor_sexo.get()

        usa_codeudor = var_tiene_codeudor.get()
        cod_nombre = ""
        cod_ci = ""
        cod_dom = ""
        if usa_codeudor:
            cod_nombre = entry_cod_nombre.get().upper()
            cod_ci = entry_cod_ci.get()
            cod_dom = entry_cod_dom.get()
            if not cod_nombre:
                messagebox.showwarning("Falta dato", "Falta el nombre del Codeudor.")
                return

        # --- 2. CÁLCULOS ---
        monto_str = monto_var.get().replace('.', '')
        if not monto_str.isdigit():
            messagebox.showerror("Error", "El monto total debe ser numérico.")
            return
        monto_total = int(monto_str)
        
        try:
            cantidad_cuotas = int(entry_cuotas.get())
            if cantidad_cuotas < 1: raise ValueError
        except:
            messagebox.showerror("Error", "Cantidad de cuotas inválida.")
            return

        monto_primera_str = monto_primera_var.get().replace('.', '')
        monto_primera = int(monto_primera_str) if monto_primera_str else 0
        monto_ultima_str = monto_ultima_var.get().replace('.', '')
        monto_ultima_objetivo = int(monto_ultima_str) if monto_ultima_str else 0

        if (monto_primera + monto_ultima_objetivo) > monto_total:
            messagebox.showerror("Error", "La suma de 1ra y Última supera el total.")
            return

        cant_reg = cantidad_cuotas
        saldo_calc = monto_total
        if monto_primera > 0:
            saldo_calc -= monto_primera
            cant_reg -= 1
        if monto_ultima_objetivo > 0:
            saldo_calc -= monto_ultima_objetivo
            cant_reg -= 1
        
        cuota_regular = 0
        if cant_reg > 0:
            cuota_regular = saldo_calc / cant_reg

        # --- 3. CONFIG ---
        fecha_inicio_obj = entry_fecha.get_date()
        frecuencia = combo_frecuencia.get()
        moneda_sel = combo_moneda.get()
        
        if "Guaraníes" in moneda_sel: m_sim, m_plu, m_sing = "GS.", "GUARANIES", "GUARANI"
        elif "Dólares" in moneda_sel: m_sim, m_plu, m_sing = "USD", "DOLARES AMERICANOS", "DOLAR AMERICANO"
        elif "Reales" in moneda_sel: m_sim, m_plu, m_sing = "R$", "REALES", "REAL"
        elif "Euros" in moneda_sel: m_sim, m_plu, m_sing = "€", "EUROS", "EURO"
        elif "Pesos Arg" in moneda_sel: m_sim, m_plu, m_sing = "ARS", "PESOS ARGENTINOS", "PESO ARGENTINO"
        else: m_sim, m_plu, m_sing = "$", "MONEDA", "MONEDA"

        acr_tit = "del señor" if acreedor_sexo == "Masculino" else "de la señora"
        acr_dom_txt = "domiciliado" if acreedor_sexo == "Masculino" else "domiciliada"
        deu_dom_txt = "domiciliado" if deudor_sexo == "Masculino" else "domiciliada"

        lugar_emision = "Bella Vista"
        fecha_emision_hoy = datetime.now().strftime("%d/%m/%Y")
        plantilla_path = "plantilla_pagare.docx"

        if not os.path.exists(plantilla_path):
            messagebox.showerror("Error", "No encuentro 'plantilla_pagare.docx'")
            return

        # --- 4. GENERACIÓN ---
        archivo_maestro = None
        composer = None
        temp_files = []
        
        fecha_venc = fecha_inicio_obj
        saldo_restante = monto_total

        for i in range(1, cantidad_cuotas + 1):
            
            es_ultima = (i == cantidad_cuotas)
            es_primera = (i == 1)
            monto_actual = 0

            if es_ultima: monto_actual = saldo_restante
            elif es_primera and monto_primera > 0: monto_actual = monto_primera
            else: monto_actual = cuota_regular

            saldo_restante -= monto_actual

            doc = DocxTemplate(plantilla_path)
            txt_mon = m_sing if monto_actual == 1 else m_plu

            context = {
                'acreedor_nombre': acreedor_nombre,
                'acreedor_nacionalidad': acreedor_nac,
                'acreedor_estado': acreedor_est,
                'acreedor_ci': acreedor_ci,
                'acreedor_dom': acreedor_dom,
                'acreedor_titulo': acr_tit,       
                'acreedor_dom_texto': acr_dom_txt,
                
                'deudor_nombre': deudor_nombre,
                'deudor_ci': deudor_ci,
                'deudor_dom': deudor_dom,
                'deudor_dom_texto': deu_dom_txt,

                'hay_codeudor': usa_codeudor,
                'codeudor_nombre': cod_nombre,
                'codeudor_ci': cod_ci,
                'codeudor_dom': cod_dom,

                'cuota_actual': f"{i:02d}",
                'cuota_total': f"{cantidad_cuotas:02d}",
                'moneda_simbolo': m_sim,
                'moneda_texto': txt_mon,
                'monto_num': "{:,.0f}".format(monto_actual).replace(",", "."),
                'monto_letras': num2words(monto_actual, lang='es').upper(),
                'fecha_venc': fecha_venc.strftime("%d/%m/%Y"),
                'fecha_emision': fecha_emision_hoy,
                'lugar': lugar_emision
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

            # --- LÓGICA DE FRECUENCIAS ACTUALIZADA ---
            if frecuencia == 'Mensual':
                 fecha_venc += relativedelta(months=1)
            elif frecuencia == 'Bimestral':
                 fecha_venc += relativedelta(months=2)
            elif frecuencia == 'Trimestral':
                 fecha_venc += relativedelta(months=3)
            elif frecuencia == 'Cuatrimestral':
                 fecha_venc += relativedelta(months=4)
            elif frecuencia == 'Semestral':
                 fecha_venc += relativedelta(months=6)
            elif frecuencia == 'Anual':
                 fecha_venc += relativedelta(years=1)

        nombre_file = f"Pagares_{deudor_nombre.replace(' ', '_')}.docx"
        path_guardado = filedialog.asksaveasfilename(initialfile=nombre_file, defaultextension=".docx", filetypes=[("Word", "*.docx")])

        if path_guardado:
            composer.save(path_guardado)
            for f in temp_files:
                try: os.remove(f)
                except: pass
            
            if messagebox.askyesno("Listo", "Documento creado con éxito.\n¿Desea abrir el archivo ahora?"):
                os.startfile(path_guardado)

    except Exception as e:
        messagebox.showerror("Error Crítico", f"Detalle: {str(e)}")

# --- GUI ---
root = tk.Tk()
root.title("Sistema Gestor de Pagarés")
root.geometry("640x950")
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
monto_var = tk.StringVar()
monto_var.trace("w", lambda *args: formatear_campo(monto_var, entry_monto))
entry_monto = ttk.Entry(frame_prestamo, textvariable=monto_var, width=20)
entry_monto.grid(row=1, column=1, padx=5, pady=5, sticky="w")
ttk.Label(frame_prestamo, text="Cant. Cuotas:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
entry_cuotas = ttk.Entry(frame_prestamo, width=10)
entry_cuotas.grid(row=2, column=1, padx=5, pady=5, sticky="w")
ttk.Label(frame_prestamo, text="1er Vencim.:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
entry_fecha = DateEntry(frame_prestamo, width=12, background='#444444', foreground='white', borderwidth=2, date_pattern='dd/mm/y', locale='es_PY', headersbackground='#222222')
entry_fecha.grid(row=3, column=1, padx=5, pady=5, sticky="w")
ttk.Label(frame_prestamo, text="Frecuencia:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
combo_frecuencia = ttk.Combobox(frame_prestamo, values=LISTA_FRECUENCIAS, state="readonly", width=10)
combo_frecuencia.current(0) # Default Mensual
combo_frecuencia.grid(row=4, column=1, padx=5, pady=5, sticky="w")

# 4. PERSONALIZACIÓN
frame_custom = ttk.LabelFrame(root, text=" 4. Cuotas Personalizadas (Opcional) ")
frame_custom.pack(fill="x", padx=15, pady=5)
ttk.Label(frame_custom, text="Monto 1ra Cuota:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
monto_primera_var = tk.StringVar()
monto_primera_var.trace("w", lambda *args: formatear_campo(monto_primera_var, entry_monto_primera))
entry_monto_primera = ttk.Entry(frame_custom, textvariable=monto_primera_var, width=15)
entry_monto_primera.grid(row=0, column=1, padx=5, pady=5, sticky="w")
ttk.Label(frame_custom, text="Monto ÚLTIMA Cuota:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
monto_ultima_var = tk.StringVar()
monto_ultima_var.trace("w", lambda *args: formatear_campo(monto_ultima_var, entry_monto_ultima))
entry_monto_ultima = ttk.Entry(frame_custom, textvariable=monto_ultima_var, width=15)
entry_monto_ultima.grid(row=0, column=3, padx=5, pady=5, sticky="w")

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