# V2.0 Alexis G

# pip install pyinstaller pandas cryptography smartsheet-python-sdk
# dependencias

# Importar librerías
import re
import glob
import os  
import html
import pandas as pd
import time
import pickle
import locale
import shutil
import tkinter as tk
import threading
import configparser
import json
from tkinter import scrolledtext
from datetime import datetime

#Librerías locales
from security import security


# Rutas de archivos y variables necesarias
config = configparser.ConfigParser()
config.read('config.cfg')

user_input = None
lock = threading.Lock()
global t
t = None
locale.setlocale(locale.LC_TIME, 'es_ES')
current_dir = os.getcwd()
pickle_file = os.path.join(current_dir, 'OCS\\facturas_procesadas.pickle')
ruta_terceros_csv = os.path.join(current_dir, 'terceros.csv')

# Funciones principales
def ha_cambiado():
    carpeta_backup = os.path.join(current_dir, 'backups')
    if not os.path.exists(carpeta_backup):
        os.makedirs(carpeta_backup, exist_ok=True)
        return True  # Si no existe la carpeta, asumimos que necesitamos hacer un backup

    backups_pickles = sorted([f for f in os.listdir(carpeta_backup) if f.endswith('.pickle')])
    if not backups_pickles:
        return True  # Si no hay backups, asumimos que necesitamos hacer uno

    ultimo_backup = os.path.join(carpeta_backup, backups_pickles[-1])
    try:
        with open(pickle_file, 'rb') as f_actual, open(ultimo_backup, 'rb') as f_ultimo:
            datos_actuales = pickle.load(f_actual)
            datos_ultimo_backup = pickle.load(f_ultimo)
    except FileNotFoundError:
        return True  # Si alguno de los archivos no existe, asumimos que hay un cambio

    return datos_actuales != datos_ultimo_backup

def guardar_backup_si_ha_cambiado():
    if ha_cambiado():  # Llamada sin argumentos
        fecha_actual = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        carpeta_backup = os.path.join(current_dir, 'backups')
        os.makedirs(carpeta_backup, exist_ok=True)
        
        # Definir las rutas de los archivos de backup
        archivo_backup_pickle = os.path.join(carpeta_backup, f'facturas_procesadas_{fecha_actual}.pickle')
        archivo_backup_excel = os.path.join(carpeta_backup, f'salida_{fecha_actual}.xlsx')
        archivo_backup_csv_oc_pendientes = os.path.join(carpeta_backup, f'OCs_Pendientes_{fecha_actual}.csv')
        
        # Copiar los archivos a la carpeta de backups
        shutil.copy(pickle_file, archivo_backup_pickle)
        
        # Mantenimiento de la cantidad de backups
        backups = sorted([f for f in os.listdir(carpeta_backup) if f.endswith('.pickle') or f.endswith('.xlsx') or f.endswith('.csv')], reverse=True)
        while len(backups) > 20:  # Asumiendo 20 versiones de cada tipo de archivo
            os.remove(os.path.join(carpeta_backup, backups.pop()))

        print("Backup realizado con éxito.")

def cargar_o_inicializar_registros():
    if not os.path.exists(pickle_file):
        with open(pickle_file, 'wb') as f:  # Crea un archivo pickle vacío si no existe.
            pickle.dump({"facturas_procesadas": {}, "carpetas_vacias": {}}, f)
    
    try:
        with open(pickle_file, 'rb') as f:  # Intenta abrir el archivo pickle en modo lectura binaria.
            return pickle.load(f)  # Retorna el diccionario cargado desde el archivo pickle.
    except (FileNotFoundError, EOFError):  # Captura errores si el archivo no existe o está vacío.
        return {"facturas_procesadas": {}, "carpetas_vacias": {}}  # Retorna un nuevo diccionario con estructuras iniciales vacías.

def registrar_carpetas_vacias():
    registros = cargar_o_inicializar_registros()  # Carga o inicializa los registros.
    subcarpetas = [d for d in glob.glob(os.path.join(current_dir, 'OCS\\*\\'))]  # Lista todas las subcarpetas.
    
    for carpeta in subcarpetas:
        if not any(f.endswith('.xml') for f in os.listdir(carpeta)):  # Si no hay archivos .xml en la carpeta.
            oc = os.path.basename(carpeta.rstrip('\\'))  # Extrae el nombre de la OC basado en el nombre de la carpeta.
            registros["carpetas_vacias"][oc] = True  # Marca la OC como carpeta vacía en los registros.

    if not os.path.exists(pickle_file):
        with open(pickle_file, 'wb') as f:  # Crea un archivo pickle vacío si no existe.
            pickle.dump({"facturas_procesadas": {}, "carpetas_vacias": {}}, f)

    print("Registrando carpetas")

def limpiar_registros_carpetas():
    registros = cargar_o_inicializar_registros()  # Carga los registros actuales.
    carpetas_a_eliminar = [carpeta for carpeta in registros["carpetas_vacias"] if any(f.endswith('.xml') for f in os.listdir(os.path.join(current_dir, 'OCS', carpeta)))] #Identifica carpetas a limpiar

    for carpeta in carpetas_a_eliminar:  # Elimina las entradas de carpetas que ya no están vacías.
        del registros["carpetas_vacias"][carpeta]

    with open(pickle_file, 'wb') as f:  # Guarda los registros actualizados en el archivo pickle.
        pickle.dump(registros, f)

def normalizar_ruc(ruc, longitud_estandar=13):
    # Asegura que el RUC tenga la longitud estándar, añadiendo ceros al inicio si es necesario
    return ruc.zfill(longitud_estandar)

def extraer_informacion_de_archivo(ruta_archivo):

    # Verificar si existe un archivo PDF con el mismo nombre que el archivo XML en la misma carpeta
    global user_input  # Declarar user_input como global
    ruta_pdf = os.path.splitext(ruta_archivo)[0] + '.pdf'
    nombre_carpeta = os.path.basename(os.path.dirname(ruta_archivo))

    # Iniciar un bucle que continúa hasta que el archivo PDF exista
    while not os.path.exists(ruta_pdf):
        print(f"Por favor verifique la OC en {nombre_carpeta}. Una vez solventado ingrese OK para continuar.")
        while True:
            with lock:
                if user_input is not None:
                    entrada = user_input.strip().lower()
                    print(f"Entrada recibida: {user_input}")  # Mostrar la entrada recibida
                    user_input = None  # Restablecer user_input para evitar repeticiones
                    
                    if entrada == "ok":
                        if os.path.exists(ruta_pdf):
                            print("Archivo PDF encontrado. Continuando con el proceso.")
                            break  # Salir del bucle si el archivo PDF ya existe
                        else:
                            print("El archivo PDF aún no existe. Por favor, verifique y vuelva a intentarlo.")
                    elif entrada == "":
                        print("No se detectó ninguna entrada. Por favor, ingrese OK cuando esté listo.")
                    else:
                        print("Entrada no reconocida. Ingrese OK para confirmar que el archivo PDF está listo.")
            time.sleep(1)  # Pequeña pausa para evitar saturación de CPU

    try:
        with open(ruta_archivo, 'r', encoding='utf-8') as archivo:
            contenido = archivo.read()
    except UnicodeDecodeError:
        with open(ruta_archivo, 'r', encoding='ISO-8859-1', errors='replace') as archivo:
            contenido = archivo.read()

    contenido = html.unescape(contenido)

    patrones = {
        'autorizacion': r'<numeroAutorizacion>(.*?)<\/numeroAutorizacion>',
        'claveAcceso' : r'<claveAcceso>(.*?)<\/claveAcceso>',
        'ruc': r'<ruc>(.*?)<\/ruc>',
        'estab': r'<estab>(.*?)<\/estab>',
        'ptoEmi': r'<ptoEmi>(.*?)<\/ptoEmi>',
        'secuencial': r'<secuencial>(.*?)<\/secuencial>',
        'fecha_emision': r'<fechaEmision>(.*?)<\/fechaEmision>',
        'nombre_comercial': r'<razonSocial>(.*?)<\/razonSocial>',
        'compania': r'<razonSocialComprador>(.*?)<\/razonSocialComprador>'
    }

    datos_extraidos = {}
    for clave, patron in patrones.items():
        coincidencia = re.search(patron, contenido, re.DOTALL)
        if coincidencia:
            datos_extraidos[clave] = coincidencia.group(1)
        else:
            datos_extraidos[clave] = 'No Disponible'

    if datos_extraidos['fecha_emision'] != 'No Disponible':
        datos_extraidos['fecha_formateada'] = datos_extraidos['fecha_emision']
    else:
        datos_extraidos['fecha_formateada'] = 'No Disponible'

    if datos_extraidos['autorizacion'] == 'No Disponible':
        datos_extraidos['autorizacion'] = datos_extraidos['claveAcceso']
        datos_extraidos.pop('claveAcceso', '')
    else:
        datos_extraidos.pop('claveAcceso', '')


    datos_extraidos['OC'] = os.path.basename(os.path.dirname(ruta_archivo))

    ruc = datos_extraidos.get('ruc')
    ruc_normalizado = normalizar_ruc(ruc)
    mapeo_terceros = cargar_y_mapear_terceros(ruta_terceros_csv)
    datos_extraidos['Tercero'] = mapeo_terceros.get(ruc_normalizado, {}).get('TERCERO', 'No Disponible')
    datos_extraidos['Centro de Costo'] = mapeo_terceros.get(ruc_normalizado, {}).get('CC', 'No Disponible')
    datos_extraidos['Nombre Farmacia'] = mapeo_terceros.get(ruc_normalizado, {}).get('NOMBRE FARMACIA', 'No Disponible')
    datos_extraidos['Frecuencia facturación'] = mapeo_terceros.get(ruc_normalizado, {}).get('FACTURA SEMESTRAL/MENSUAL', 'No Disponible')

    descripcion_tags = re.findall(r'<descripcion>(.*?)<\/descripcion>', contenido)
    precio_unitario_tags = re.findall(r'<precioUnitario>(.*?)<\/precioUnitario>', contenido)
    cantidad_tags = re.findall(r'<cantidad>(.*?)<\/cantidad>', contenido)

    precio_unitario_redondeado = [f"{float(precio):.2f}" for precio in precio_unitario_tags]
    cantidad_redondeado = [f"{float(cantidad):.2f}" for cantidad in cantidad_tags]

    descriptions_with_prices = [
        f"{descripcion} ({precio_unitario} x {cantidad})"
        for descripcion, precio_unitario, cantidad in zip(descripcion_tags, precio_unitario_redondeado, cantidad_redondeado)
    ]
    datos_extraidos['descripciones'] = " - ".join(descriptions_with_prices)
    
    subtotales = {}
    iva_values = {}

    # Buscar en la etiqueta "detalle" -> "impuestos" -> "impuesto"
    detalle_pattern = r'<detalle>(.*?)<\/detalle>'
    detalle_matches = re.findall(detalle_pattern, contenido, re.DOTALL)

    for detalle in detalle_matches:
        impuestos_pattern = r'<impuesto>(.*?)<\/impuesto>'
        impuestos_matches = re.findall(impuestos_pattern, detalle, re.DOTALL)

        for impuesto in impuestos_matches:
            base_imponible_pattern = r'<baseImponible>(.*?)<\/baseImponible>'
            tarifa_pattern = r'<tarifa>(.*?)<\/tarifa>'
            valor_pattern = r'<valor>(.*?)<\/valor>'

            base_imponible = re.search(base_imponible_pattern, impuesto)
            tarifa = re.search(tarifa_pattern, impuesto)
            valor = re.search(valor_pattern, impuesto)

            if base_imponible and tarifa and valor:
                base_imponible = float(base_imponible.group(1))
                tarifa = float(tarifa.group(1))
                valor = float(valor.group(1))

                if tarifa not in subtotales:
                    subtotales[tarifa] = 0
                if tarifa not in iva_values:
                    iva_values[tarifa] = 0

                subtotales[tarifa] += base_imponible
                iva_values[tarifa] += valor

    if not subtotales and not iva_values:
        total_impuestos_pattern = r'<totalImpuesto>(.*?)<\/totalImpuesto>'
        total_impuestos_matches = re.findall(total_impuestos_pattern, contenido, re.DOTALL)

        for total_impuesto in total_impuestos_matches:
            base_imponible_pattern = r'<baseImponible>(.*?)<\/baseImponible>'
            tarifa_pattern = r'<tarifa>(.*?)<\/tarifa>'
            valor_pattern = r'<valor>(.*?)<\/valor>'

            base_imponible = re.search(base_imponible_pattern, total_impuesto)
            tarifa = re.search(tarifa_pattern, total_impuesto)
            valor = re.search(valor_pattern, total_impuesto)

            if base_imponible and tarifa and valor:
                base_imponible = float(base_imponible.group(1))
                tarifa = float(tarifa.group(1))
                valor = float(valor.group(1))

                if tarifa not in subtotales:
                    subtotales[tarifa] = 0
                if tarifa not in iva_values:
                    iva_values[tarifa] = 0

                subtotales[tarifa] += base_imponible
                iva_values[tarifa] += valor

    return datos_extraidos, subtotales, iva_values

def copiar_carpetas_oc(ruta_carpeta_consolidada):
    # Obtener las subcarpetas de OC en la carpeta OCS
    subcarpetas_oc = [d for d in glob.glob(os.path.join(current_dir, 'OCS\\*\\'))]

    for carpeta in subcarpetas_oc:
        # Buscar archivos XML y PDF dentro de cada carpeta
        archivos_xml = glob.glob(os.path.join(carpeta, '*.xml'))
        archivos_pdf = glob.glob(os.path.join(carpeta, '*.pdf'))
        
        # Consolidar archivos XML
        for archivo_xml in archivos_xml:
            nombre_archivo = os.path.basename(archivo_xml)
            ruta_destino_xml = os.path.join(ruta_carpeta_consolidada, nombre_archivo)
            shutil.copy2(archivo_xml, ruta_destino_xml)
        
        # Consolidar archivos PDF
        for archivo_pdf in archivos_pdf:
            nombre_archivo = os.path.basename(archivo_pdf)
            ruta_destino_pdf = os.path.join(ruta_carpeta_consolidada, nombre_archivo)
            shutil.copy2(archivo_pdf, ruta_destino_pdf)

    print("Copiado de archivos completado en la carpeta")

def cargar_y_mapear_terceros(ruta_terceros_csv):
    # Intenta leer el archivo CSV con diferentes codecs
    try:
        terceros_df = pd.read_csv(ruta_terceros_csv, encoding='utf-8')
    except UnicodeDecodeError:
        terceros_df = pd.read_csv(ruta_terceros_csv, encoding='latin1')  # Prueba con el codec latin1

    # Normaliza la columna RUC
    terceros_df['RUC'] = terceros_df['RUC'].apply(lambda x: normalizar_ruc(str(x)))

    # Elimina duplicados en la columna 'RUC'
    terceros_df.drop_duplicates(subset='RUC', inplace=True)

    # Reemplaza valores NaN por la cadena 'NaN'
    terceros_df.fillna('NaN', inplace=True)

    # Crea el mapeo de los terceros
    mapeo_terceros = terceros_df.set_index('RUC')[['TERCERO', 'CC', 'NOMBRE FARMACIA', 'FACTURA SEMESTRAL/MENSUAL']].to_dict(orient='index')
    
    return mapeo_terceros

def generar_json_consolidado():
    ubicacion_destino = config['ubicacion']['ruta']
    regional = config['ubicacion']['regional']

    # Crear la carpeta consolidada con el formato OCS-{regional}-{fecha-hora actual}
    fecha_hora_actual = datetime.now().strftime('%Y%m%d-%H%M%S')
    nombre_carpeta_consolidada = f"OCS-{regional}-{fecha_hora_actual}"
    ruta_carpeta_consolidada = os.path.join(ubicacion_destino, nombre_carpeta_consolidada)

    # Si la carpeta consolidada no existe, la creamos
    if not os.path.exists(ruta_carpeta_consolidada):
        os.makedirs(ruta_carpeta_consolidada)

    # Intentar cargar el registro de facturas procesadas desde el archivo pickle
    try:
        with open(pickle_file, 'rb') as f:
            facturas_procesadas = pickle.load(f)
    except (FileNotFoundError, EOFError):
        # Si el archivo pickle no existe o está vacío, procesar todas las OCs
        facturas_procesadas = {}

    # Obtener todos los archivos XML en la carpeta OCS
    archivos_xml = glob.glob(os.path.join(current_dir, 'OCS', '**', '*.xml'), recursive=True)
    
    # Crear una lista para almacenar los datos de todas las OCs procesadas en esta ejecución
    lista_oc_consolidada = []


    for ruta_archivo in archivos_xml:
        # Extraer información del archivo XML
        informacion, subtotales, iva_values = extraer_informacion_de_archivo(ruta_archivo)
        oc = informacion['OC']

        # Verificar si la OC ya ha sido procesada
        if oc not in facturas_procesadas:
            # Preparar los datos para añadir al JSON consolidado
            factura = f"{informacion['estab']}-{informacion['ptoEmi']}-{informacion['secuencial']}"
            datos_para_json = {
                'Autorizacion': informacion["autorizacion"],
                'RUC': informacion['ruc'],
                'Tercero': informacion['Tercero'],
                'Nombre Comercial': informacion['nombre_comercial'],
                'Compañía': informacion['compania'],
                'Centro de Costo': informacion['Centro de Costo'],
                'Nombre Farmacia': informacion['Nombre Farmacia'],
                'OC': oc,
                'Factura': factura,
                'Fecha': informacion['fecha_formateada'],
                'Descripcion': informacion['descripciones'],
                'Subtotal 0%': str(subtotales.get(0, 0)),
                'Tarifa': "-".join(str(tarifa) for tarifa in subtotales if tarifa != 0),
                'Subtotales Impuesto': "-".join(str(subtotales[tarifa]) for tarifa in subtotales if tarifa != 0),
                'IVA': "-".join(str(iva_values[tarifa]) for tarifa in iva_values if tarifa != 0),
                'Frecuencia facturación': informacion['Frecuencia facturación']
            }
            lista_oc_consolidada.append(datos_para_json)

            # Renombrar los archivos XML y PDF utilizando el número de autorización
            numero_autorizacion = informacion.get('autorizacion', 'NoDisponible')
            carpeta_actual = config['ubicacion']['ruta']
            nuevo_nombre_xml = f"{numero_autorizacion}.xml"
            nuevo_nombre_pdf = f"{numero_autorizacion}.pdf"
            ruta_nuevo_xml = os.path.join(carpeta_actual, nombre_carpeta_consolidada, nuevo_nombre_xml)
            ruta_nuevo_pdf = os.path.join(carpeta_actual, nombre_carpeta_consolidada, nuevo_nombre_pdf)
            ruta_pdf = os.path.splitext(ruta_archivo)[0] + '.pdf'

            # Renombrar el archivo XML si no existe ya con el nuevo nombre
            if not os.path.exists(ruta_nuevo_xml):
                os.rename(ruta_archivo, ruta_nuevo_xml)

            # Renombrar el archivo PDF si no existe ya con el nuevo nombre
            if os.path.exists(ruta_pdf) and not os.path.exists(ruta_nuevo_pdf):
                os.rename(ruta_pdf, ruta_nuevo_pdf)

    # Copiar las carpetas OC
    copiar_carpetas_oc(ruta_carpeta_consolidada)  # OJO

    # Crear la carpeta consolidada
    ruta_carpeta_consolidada = os.path.join(ubicacion_destino, nombre_carpeta_consolidada)

    # Generar el nombre del archivo JSON consolidado
    nombre_archivo_json = os.path.join(ruta_carpeta_consolidada, 'OCS_consolidado.json')

    # Guardar todos los datos en un único archivo JSON
    if lista_oc_consolidada:  # Solo guardamos si hay nuevas OCs procesadas
        os.makedirs(ruta_carpeta_consolidada, exist_ok=True)  # Crear la carpeta si no existe
        with open(nombre_archivo_json, 'w', encoding='utf-8') as archivo_json:
            json.dump(lista_oc_consolidada, archivo_json, ensure_ascii=False, indent=4)
        
        print("Archivo JSON consolidado generado")

def main():
    try:
        registrar_carpetas_vacias()
        limpiar_registros_carpetas()
        generar_json_consolidado() #OJO
        guardar_backup_si_ha_cambiado()
    except Exception as e:
        print(f"Error durante la ejecución de tareas: {e}")
    finally:
        print("Finalizando la ejecución de tareas.")

# Código interfaz gráfica
def iniciar_tareas():
    print("Iniciando gestión de facturas, por favor espere.")
    t = threading.Thread(target=main)
    t.start()

def enviar_input():
    global user_input
    entrada = entry_box.get()
    entry_box.delete(0, tk.END)
    user_input = entrada
    print(f"Entrada recibida")
    if validar_clave(entrada):
        print("Clave válida. Puede iniciar las tareas.")
        start_button.config(state=tk.NORMAL)
    else:
        print("Clave inválida. Inténtelo de nuevo.")
        start_button.config(state=tk.DISABLED)

def validar_clave(clave_ingresada):
    datos = security()
    return clave_ingresada == datos['clave']

window = tk.Tk()
window.title("Gestión de Facturas GPF")

text_area = scrolledtext.ScrolledText(window, wrap=tk.WORD, width=45, height=14)
text_area.grid(column=0, row=0, columnspan=3, pady=10, padx=10, sticky="nsew")

window.grid_columnconfigure(0, weight=1)
window.grid_columnconfigure(1, weight=1)
window.grid_columnconfigure(2, weight=1)
window.grid_rowconfigure(0, weight=1)

def print(*args, **kwargs):
    text_area.insert(tk.END, ' '.join(map(str, args)) + '\n')
    text_area.see(tk.END)

print("Bienvenido. Por favor, ingrese la clave para iniciar.")

entry_box = tk.Entry(window, width=35)
entry_box.grid(column=0, row=1, pady=10)

input_button = tk.Button(window, text="Enviar", command=enviar_input)
input_button.grid(column=1, row=1)

start_button = tk.Button(window, text="Iniciar", command=iniciar_tareas)
start_button.grid(column=0, row=2, pady=10)
start_button.config(state=tk.DISABLED)

stop_button = tk.Button(window, text="Terminar", command=lambda: os._exit(0))
stop_button.grid(column=1, row=2)

window.mainloop()