import os
import pickle
import smartsheet
import random
import string
from datetime import datetime
from cryptography.fernet import Fernet

# Conexión a Smartsheet
ACCESS_TOKEN = 'xxx'
SHEET_ID = 'xxx'
ROW_ID = xxx
PICKLE_FILE = 'appkey.pkl'
CLAVE_CIF = b'ZLqGI4S4zc_qOYzozyf3WV9Mo4lINSe0PZUSjoKMRS0='

def escribirSmartsheet(NEW_VALUE, COLUMN_ID):
    # Inicializar el cliente de Smartsheet
    ss_client = smartsheet.Smartsheet(ACCESS_TOKEN)

    # Crear la actualización para la celda específica
    cell_update = ss_client.models.Cell()
    cell_update.column_id = COLUMN_ID
    cell_update.value = NEW_VALUE

    # Crear la fila que contiene la actualización de la celda
    row_update = ss_client.models.Row()
    row_update.id = ROW_ID
    row_update.cells.append(cell_update)

    # Realizar la actualización en la hoja
    updated_row = ss_client.Sheets.update_rows(SHEET_ID, [row_update])

# Cifrar una cadena
def cifrar_cadena(cadena, clave):
    f = Fernet(clave)
    cadena_cifrada = f.encrypt(cadena.encode())
    return cadena_cifrada

# Descifrar una cadena
def descifrar_cadena(cadena_cifrada, clave):
    f = Fernet(clave)
    cadena_descifrada = f.decrypt(cadena_cifrada).decode()
    return cadena_descifrada

def generar_clave(seed: str):
    random.seed(seed)
    caracteres = string.ascii_letters + string.digits
    clave = random.sample(string.ascii_letters, 2) + random.sample(string.digits, 2) + random.choices(caracteres, k=3)
    random.shuffle(clave)
    clave_str = ''.join(clave)
    generar_clave.cifrada = cifrar_cadena(clave_str, CLAVE_CIF)
    return clave_str

def generar_cadena_fecha():
    # Obtener la fecha y hora actual
    ahora = datetime.now()
    # Formatear la cadena en el formato deseado
    cadena = f"INMO-{ahora.month:02d}-{ahora.year}"
    return cadena

def guardar_datos(seed: str, app_key: str):
    with open(PICKLE_FILE, 'wb') as f:
        pickle.dump({'seed': seed, 'app_key': app_key}, f)

def cargar_datos():
    if os.path.exists(PICKLE_FILE):
        with open(PICKLE_FILE, 'rb') as f:
            return pickle.load(f)
    return None

def security():
    cadena = None
    seed_actual = generar_cadena_fecha()
    datos_guardados = cargar_datos()

    if datos_guardados:
        seed_guardado = datos_guardados['seed']
        app_key_guardada = datos_guardados['app_key']
        
        if seed_guardado == seed_actual:
            cadena = generar_cadena_fecha()
            clave_descifrada = descifrar_cadena(datos_guardados['app_key'],CLAVE_CIF)
            return {"cadena":cadena, "clave_cifrada": app_key_guardada, "clave": clave_descifrada}
    
    nueva_clave = generar_clave(seed_actual)
    nueva_clave_cifrada = generar_clave.cifrada
    guardar_datos(seed_actual, nueva_clave_cifrada)
    escribirSmartsheet(nueva_clave, 6253555087527812)
    cadena = generar_cadena_fecha()
    return {"cadena":cadena, "clave_cifrada": nueva_clave_cifrada, "clave": nueva_clave}