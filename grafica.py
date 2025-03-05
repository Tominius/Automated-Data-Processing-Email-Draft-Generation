import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import os
import csv
from email.message import EmailMessage
import mimetypes

ruta_archivo = ""  # Variable global para almacenar la ruta del archivo
carpeta_excels=""  # Variable global para almacenar la ruta del carpeta excels
carpeta_mails = "" # Variable global para almacenar la ruta del carpeta mails
ruta_archivo_csv = ""  # Variable global para almacenar la ruta del archivo .csv
ruta_cliente_mail= ""

def seleccionar_archivo():
    global ruta_archivo
    ruta_archivo = filedialog.askopenfilename(title="Selecciona un archivo Excel", filetypes=[("Archivos Excel", "*.xlsx;*.xls")])
    if ruta_archivo:
        etiqueta.config(text=f"📂 Archivo seleccionado:\n{ruta_archivo}")

def seleccionar_carpeta():
    global carpeta_excels
    carpeta_excels = filedialog.askdirectory(title="Selecciona la carpeta para los archivos Excel")
    if carpeta_excels:
        etiqueta_carpeta.config(text=f"📁 Carpeta seleccionada:\n{carpeta_excels}")

def seleccionar_archivo_csv():
    global ruta_archivo_csv
    ruta_archivo_csv = filedialog.askopenfilename(title="Selecciona un archivo CSV", filetypes=[("Archivos CSV", "*.csv")])
    if ruta_archivo_csv:
        etiqueta_csv.config(text=f"📄 Archivo CSV seleccionado:\n{ruta_archivo_csv}")

def seleccionar_carpeta_mail():
    global carpeta_mails
    carpeta_mails = filedialog.askdirectory(title="Selecciona la carpeta para los mails")
    if carpeta_mails:
        etiqueta_mails.config(text=f"📧 Carpeta seleccionada:\n{carpeta_mails}")

def seleccionar_archivo_cliente_mail():
    global ruta_cliente_mail
    ruta_cliente_mail = filedialog.askopenfilename(title="Selecciona un archivo CSV", filetypes=[("Archivos CSV", "*.csv")])
    if ruta_cliente_mail:
        etiqueta_cliente_mail.config(text=f"📜 Archivo seleccionado:\n{ruta_cliente_mail}")

def convertir_xlsx_a_csv():
    """Función para convertir el archivo Excel a CSV"""
    global ruta_archivo  # Aseguramos que usamos la variable global
    if not ruta_archivo:  
        etiqueta_resultado.config(text="⚠️ Selecciona un archivo primero.")
        return

    archivo_csv = os.path.splitext(ruta_archivo)[0] + ".csv"

    try:
        df = pd.read_excel(ruta_archivo, skiprows=2)  
        df.dropna(axis=1, how='all', inplace=True)  # Eliminar columnas vacías
        df.dropna(axis=0, how='all', inplace=True)  # Eliminar filas vacías
        df.to_csv(archivo_csv, index=False, encoding='utf-8', lineterminator="\n")

        etiqueta_resultado.config(text=f"✅ Archivo convertido con éxito:\n{archivo_csv}")
    except Exception as e:
        etiqueta_resultado.config(text=f"❌ Error al convertir archivo:\n{str(e)}")

def excels():
    global carpeta_excels, ruta_archivo_csv

    if not ruta_archivo_csv:
        print("⚠️ Debes seleccionar un archivo CSV antes de continuar.")
        return

    if not carpeta_excels:
        print("⚠️ Debes seleccionar una carpeta de destino.")
        return

    # Obtener valores de entrada
    try:
        cliente_texto = entrada_cliente.get().strip()
        columnas_texto = entrada_columnas.get().strip()
        nombres_texto = entrada_nombres.get().strip()

        if not cliente_texto or not columnas_texto or not nombres_texto:
            raise ValueError("Hay campos vacíos.")

        entrada_cliente_valor = int(cliente_texto)
        entrada_columnas_valor = list(map(int, columnas_texto.split(",")))
        entrada_nombres_valor = nombres_texto.split(",")

        print(f"✅ Columnas seleccionadas: {entrada_columnas_valor}")
        print(f"✅ Nombre de columnas: {entrada_nombres_valor}")
        print(f"✅ Columna de clientes: {entrada_cliente_valor}")

    except ValueError as e:
        print(f"⚠️ Error en los valores de entrada: {str(e)}")
        return

    clientes = set()

    # Leer el archivo CSV y obtener la lista de clientes
    with open(ruta_archivo_csv, "r", encoding="utf-8") as archivo:
        reader = csv.reader(archivo)
        next(reader, None)  # Saltar la primera fila si es un encabezado
        for linea in reader:
            if len(linea) > entrada_cliente_valor:
                cliente_nombre = linea[entrada_cliente_valor].strip()
                if cliente_nombre:
                    clientes.add(cliente_nombre)

    print(f"📋 Clientes detectados: {clientes}")

    # Procesar cada cliente y generar su archivo Excel
    for cliente in clientes:
        if not cliente.strip():
            continue

        archivo_excel = os.path.join(carpeta_excels, f"{cliente}.xlsx")

        if os.path.exists(archivo_excel):
            df = pd.read_excel(archivo_excel)
            print(f"📂 Archivo {archivo_excel} encontrado, agregando datos.")
        else:
            df = pd.DataFrame(columns=entrada_nombres_valor)
            print(f"🆕 Creando nuevo archivo: {archivo_excel}")

        datos_nuevos = []

        # Leer CSV nuevamente y filtrar por cliente
        with open(ruta_archivo_csv, "r", encoding="utf-8") as archivo:
            reader = csv.reader(archivo)
            next(reader, None)  # Saltar encabezado
            for fila in reader:
                if len(fila) > max(entrada_columnas_valor):  
                    datos_fila = [fila[i].strip() for i in entrada_columnas_valor]

                    print(f"🔎 Analizando fila: {datos_fila}")

                    if fila[entrada_cliente_valor].strip() == cliente:
                        print(f"✅ Fila agregada para {cliente}: {datos_fila}")
                        while len(datos_fila) < len(entrada_nombres_valor):
                            datos_fila.append("")  # Rellenar con valores vacíos

                        datos_nuevos.append(datos_fila[:len(entrada_nombres_valor)])

        if datos_nuevos:
            df_nuevo = pd.DataFrame(datos_nuevos, columns=entrada_nombres_valor)
            df = pd.concat([df, df_nuevo], ignore_index=True)
            df.to_excel(archivo_excel, index=False)
            print(f"✅ Archivo generado correctamente para {cliente}: {archivo_excel}")
        else:
            print(f"⚠️ No se encontraron datos para {cliente}, archivo vacío.")




def correos():
    if not ruta_cliente_mail or not carpeta_mails:
        messagebox.showwarning("Advertencia", "Selecciona la carpeta de mails y el archivo de clientes.")
        return
    emails_dict = {}
    try:
        with open(ruta_cliente_mail, "r", encoding="utf-8") as csv_file:
            reader = csv.reader(csv_file)
            for row in reader:
                if len(row) >= 2:
                    emails_dict[row[0].strip()] = row[1].strip()
        print("📧 Diccionario de correos cargado:", emails_dict)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el archivo de clientes:\n{str(e)}")
        return
    for archivo in os.listdir(carpeta_excels):
        if archivo.endswith(".xlsx"):
            ruta_archivo = os.path.join(carpeta_excels, archivo)
            email_destinatario = emails_dict.get(os.path.splitext(archivo)[0])
            if not email_destinatario:
                print(f"No se encontró email para {archivo}, omitido.")
                continue
            msg = EmailMessage()
            msg['Subject'] = 'Deliveries'
            msg['From'] = 'tarteach@uade.edu.ar'
            msg['To'] = email_destinatario
            msg['X-Unsent'] = '1'
            msg.add_alternative("""
            <html>
            <body>
                <p>Estimado/a,</p>
                <p>Adjunto encontrará los DELIVERIES.</p>
                <p>Saludos cordiales,<br>Tominius</p>
            </body>
            </html>
            """, subtype='html')
            if os.path.exists(ruta_archivo):
                tipo_mime, _ = mimetypes.guess_type(ruta_archivo)
                if tipo_mime is None:
                    tipo_mime = 'application/octet-stream'
                with open(ruta_archivo, 'rb') as adjunto:
                    msg.add_attachment(adjunto.read(), maintype=tipo_mime.split('/')[0], subtype=tipo_mime.split('/')[1], filename=archivo)
                ruta_eml = os.path.join(carpeta_mails, archivo + ".eml")
                with open(ruta_eml, 'wb') as f:
                    f.write(msg.as_bytes())
                print(f"Correo borrador guardado: {ruta_eml}")
            else:
                print(f"Archivo no encontrado: {ruta_archivo}")


# Configuración de la ventana Tkinter
app = tk.Tk()
app.geometry("600x900")
app.config(background="#bf4342")
app.title("AdminTool by H")

# Botones y etiquetas
boton_seleccionar = tk.Button(app, text="📂 Seleccionar Excel", command=seleccionar_archivo, bg="#e7d7c1", fg="black")
boton_seleccionar.pack(pady=5)

etiqueta = tk.Label(app, text="Ningún archivo seleccionado", bg="#bf4342", fg="white")
etiqueta.pack()


boton_convertir = tk.Button(app, text="📄 Convertir a CSV", command=convertir_xlsx_a_csv, bg="#e7d7c1", fg="black")
boton_convertir.pack(pady=5)

etiqueta_resultado = tk.Label(app, text="", bg="#bf4342", fg="white")
etiqueta_resultado.pack()


# Etiqueta para mostrar el resultado
etiqueta_resultado = tk.Label(app, text="", bg="#bf4342", fg="white", wraplength=500)
etiqueta_resultado.pack(pady=10)

# Botón para seleccionar CSV
boton_csv = tk.Button(app, text="Seleccionar archivo CSV", command=seleccionar_archivo_csv, bg="#e7d7c1", fg="black")
boton_csv.pack(pady=5)
etiqueta_csv = tk.Label(app, text="Ningún archivo seleccionado", bg="#bf4342", fg="white", wraplength=600)
etiqueta_csv.pack(pady=5)

# Botón para seleccionar carpeta de salida
boton_carpeta = tk.Button(app, text="Seleccionar carpeta Excels", command=seleccionar_carpeta, bg="#e7d7c1", fg="black")
boton_carpeta.pack(pady=5)
etiqueta_carpeta = tk.Label(app, text="Ninguna carpeta seleccionada", bg="#bf4342", fg="white", wraplength=600)
etiqueta_carpeta.pack(pady=5)

# Entrada para posiciones de columnas
tk.Label(app, text="Nombres de columnas (separadas por coma):", bg="#bf4342", fg="white").pack(pady=5)
entrada_columnas = tk.Entry(app, width=50)
entrada_columnas.pack(pady=5)
entrada_columnas.insert(0, "0,1,2")  # Valores predeterminados

# Entrada para posición del cliente
tk.Label(app, text="Posición de la columna de clientes (filas a agrupar):", bg="#bf4342", fg="white").pack(pady=5)
entrada_cliente = tk.Entry(app, width=10)
entrada_cliente.pack(pady=5)
entrada_cliente.insert(0, "8")  # Valor predeterminado

# Entrada para nombres de columnas
tk.Label(app, text="Nombres de las nuevas columnas (separados por coma):", bg="#bf4342", fg="white").pack(pady=5)
entrada_nombres = tk.Entry(app, width=50)
entrada_nombres.pack(pady=5)
entrada_nombres.insert(0, "Fila A, Fila B, Fila C")

# Botón para generar archivos Excel
boton_generar = tk.Button(app, text="Generar Excels por Cliente", command=excels, bg="#e7d7c1", fg="black")
boton_generar.pack(pady=10)

# Botón para seleccionar carpeta de salida
boton_carpeta = tk.Button(app, text="Seleccionar carpeta Mails", command=seleccionar_carpeta_mail, bg="#e7d7c1", fg="black")
boton_carpeta.pack(pady=5)
etiqueta_mails = tk.Label(app, text="Ninguna carpeta seleccionada", bg="#bf4342", fg="white", wraplength=600)
etiqueta_mails.pack(pady=5)

# Botón para seleccionar CSV
boton_csv = tk.Button(app, text="Seleccionar archivo cliente_mail.csv", command=seleccionar_archivo_cliente_mail, bg="#e7d7c1", fg="black")
boton_csv.pack(pady=5)
etiqueta_cliente_mail = tk.Label(app, text="Ningún archivo seleccionado", bg="#bf4342", fg="white", wraplength=600)
etiqueta_cliente_mail.pack(pady=5)
boton_generar_eml = tk.Button(app, text="📧 Generar mails", command=correos, bg="#e7d7c1", fg="black")
boton_generar_eml.pack(pady=5)

app.mainloop()