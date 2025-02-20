import pandas as pd
import csv
import os
from email.message import EmailMessage
import mimetypes

#------------------------------------------------------------------------Convertidor--------------------------------------------------------------------------

def convertir_xlsx_a_csv(archivo_xlsx, archivo_csv):
    # Cargar el archivo Excel
    df = pd.read_excel(archivo_xlsx, skiprows=2)  # Ajusta según el formato del archivo

    # Eliminar columnas completamente vacías
    df.dropna(axis=1, how='all', inplace=True)

    # Eliminar filas completamente vacías
    df.dropna(axis=0, how='all', inplace=True)

    # Guardar como archivo CSV sin incluir líneas en blanco
    df.to_csv(archivo_csv, index=False, encoding='utf-8', lineterminator="\n")

    print(f"Archivo convertido sin filas vacías: {archivo_csv}")

# Ejecutar la función
convertir_xlsx_a_csv("ruta.xlsx", "archivo.csv")

#------------------------------------------------------------------------Convertidor---------------------------------------------------------------------------

#------------------------------------------------------------------------Creacion de Excels--------------------------------------------------------------------


def excels(arch_csv, carpeta_Excels, pos, nombresColumnas, posClientes):
    # Lista de clientes únicos
    clientes = set()

    # Leer el archivo CSV y obtener la lista de clientes
    with open(arch_csv, "r", encoding="utf-8") as arch:
        reader = csv.reader(arch)
        for lineaL in reader:
            if len(lineaL) > posClientes:  # Asegurar que hay suficientes columnas
                clientes.add(lineaL[posClientes].strip())  # Limpiar espacios en blanco

    # Procesar cada cliente y generar su archivo Excel
    for cliente in clientes:
        if not cliente.strip():  # Evita nombres vacíos
            print("⚠️ Cliente vacío encontrado, saltando...")
            continue
        else:
            archivo_excel = os.path.join(carpeta_Excels, f"{cliente}.xlsx")

            # Intentar cargar el archivo Excel existente
            if os.path.exists(archivo_excel):
                df = pd.read_excel(archivo_excel)
                print(f"📂 Archivo {archivo_excel} encontrado, agregando datos.")
            else:
                df = pd.DataFrame(columns=nombresColumnas)  # Crear nuevo DataFrame con encabezados
                print(f"🆕 Archivo {archivo_excel} no encontrado, creando nuevo.")

            # Procesar y filtrar datos para Excel
            with open(archivo_csv, "r", encoding="utf-8") as arch:
                reader = csv.reader(arch)
                datos_nuevos = []
                
                for datos in reader:
                    if len(datos) <= max(pos):  
                        continue  # Saltar filas con menos columnas de las necesarias
                    
                    # Filtrar columnas por índice
                    datosExcel = [datos[i].strip() for i in pos]

                    if datosExcel[1] == cliente:  # Verificar que el cliente coincide
                        # Asegurar que los datos tengan el mismo número de columnas que el DataFrame
                        while len(datosExcel) < len(columnas_fijas):
                            datosExcel.append("")  # Rellenar con valores vacíos

                        datos_nuevos.append(datosExcel[:len(columnas_fijas)])

                if datos_nuevos:
                    df_nuevo = pd.DataFrame(datos_nuevos, columns=columnas_fijas)
                    df = pd.concat([df, df_nuevo], ignore_index=True)

            # Guardar el archivo Excel actualizado
            df.to_excel(archivo_excel, index=False)
            print(f"✅ Datos insertados correctamente en {archivo_excel}.")


# Ruta del archivo CSV fuente
archivo_csv = "ruta.csv" #al archivo creado en el paso de convercion
carpeta_excels = "AdminTool\\Excels"

# Lista de posiciones a mantener (columnas que queremos en el nuevo excel de los clientes individuales)
posiciones = [7, 8, 9, 10, 11, 12, 13, 14]

posClientes = 8  #Posicion de columna donde se encuentran los clientes 

# Nombres de columnas fijos
columnas_fijas = [
    "Fecha puesta dis. Mat", "Nombre del Solicitante", "Material",
    "Denominacion", "Peso Total", "Ruta", "Lugar-Solicitante-Pedido",
    "Fecha de Carga", "Numero de Transporte"
]

excels(archivo_csv,carpeta_excels,posiciones,columnas_fijas,posClientes)


#------------------------------------------------------------------------Creacion de Excels--------------------------------------------------------------------

#------------------------------------------------------------------------Creacion de .eml----------------------------------------------------------------------

#Necesario archivo.csv que contenga (nombre de la empresa, mail del destinatario) en este caso clientes_mail.csv

def correos (CARPETA_EXCELS,CARPETA_EML,ARCHIVO_CSV ):
    emails_dict = {}
    with open(ARCHIVO_CSV, "r", encoding="utf-8") as csv_file:
        reader = csv.reader(csv_file)
        for row in reader:
            if len(row) >= 2:
                emails_dict[row[0].strip()] = row[1].strip()

    for archivo in os.listdir(CARPETA_EXCELS):
        if archivo.endswith(".xlsx"):
            ruta_archivo = os.path.join(CARPETA_EXCELS, archivo)
            email_destinatario = emails_dict.get(archivo)

            if not email_destinatario:
                print(f"⚠️ No se encontró email para {archivo}, archivo omitido.")
                continue


            # Crear el mensaje de correo
            msg = EmailMessage()
            msg['Subject'] = 'Deliveries'
            msg['From'] = 'tarteach@uade.edu.ar'
            msg['To'] = email_destinatario
            msg['X-Unsent'] = '1'  # Indica que es un borrador
            msg.add_alternative(f"""
            <html>
            <body>
                <p>Estimado/a,</p>
                <p>Espero que este mensaje le encuentre bien. Me pongo en contacto con usted para enviarle los DELIVERIES.</p>
                <p>Quedo atento/a a su respuesta y agradezco de antemano su tiempo y consideración.</p>
                <p>Saludos cordiales,<br>Tominius</p>
            </body>
            </html>
            """, subtype='html')

            # Adjuntar un archivo
            archivo_adjunto = ruta_archivo
            tipo_mime, _ = mimetypes.guess_type(archivo_adjunto)
            if tipo_mime is None:
                tipo_mime = 'application/octet-stream'

            with open(archivo_adjunto, 'rb') as adjunto:
                msg.add_attachment(adjunto.read(), maintype=tipo_mime.split('/')[0], subtype=tipo_mime.split('/')[1], filename=archivo)

            # Guardar el archivo .eml en modo binario
            ruta_eml = os.path.join(CARPETA_EML, archivo + ".eml")
            with open(ruta_eml, 'wb') as f:
                f.write(msg.as_bytes())
                
            print("Borrador guardado como 'correo_borrador.eml'")


# Rutas de carpetas
CARPETA_EXCELS = r"AdminTool\\Excels"
CARPETA_EML = r"AdminTool\\Mails"
ARCHIVO_CSV = r"AdminTool\\clientes_mails.csv" #clientes_mails

correos (CARPETA_EXCELS,CARPETA_EML,ARCHIVO_CSV )

#------------------------------------------------------------------------Creacion de .eml----------------------------------------------------------------------
