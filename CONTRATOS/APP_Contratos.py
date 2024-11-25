import tkinter as tk
from tkinter import scrolledtext, messagebox
import sys
import smartsheet
from docxtpl import DocxTemplate
from num2words import num2words
from datetime import datetime
from dateutil.relativedelta import relativedelta  # Para sumar años sin alterar el día
import locale
from jinja2 import Template
import re

# Configura tu acceso a Smartsheet
ACCESS_TOKEN = 'xxx'
SHEET_ID = 'xxx'

ss_client = smartsheet.Smartsheet(ACCESS_TOKEN)
sheet = ss_client.Sheets.get_sheet(SHEET_ID)

# Definir los IDs de las columnas basados en los proporcionados
COLUMN_ID_FOR_PROJECT_NAME = 527835979796356
COLUMN_ID_FOR_DATE = 8461911885827972
COLUMN_ID_FOR_CONTACT = 2779635793481604
COLUMN_ID_FOR_PRONOUN = 7121570343636868
COLUMN_ID_FOR_COMPANY = 7283235420852100
COLUMN_ID_FOR_PROVINCE = 1653735886638980
COLUMN_ID_FOR_CANTON = 6157335514009476
COLUMN_ID_FOR_PARISH = 3905535700324228
COLUMN_ID_FOR_ADDRESS = 8409135327694724
COLUMN_ID_FOR_METRAJE = 3342585746902916
COLUMN_ID_FOR_WATER_METER = 8972085281116036
COLUMN_ID_FOR_WATER_PAYMENT_METHOD = 6720285467430788
COLUMN_ID_FOR_LIGHT_METER = 4521262211878788
#COLUMN_ID_FOR_LIGHT_PAYMENT_METHOD = 17662584508292
COLUMN_ID_FOR_LANDLORD_NAME = 6773062025564036
COLUMN_ID_FOR_LANDLORD_ID = 3955193751490436
COLUMN_ID_FOR_LANDLORD_MARITAL_STATUS = 1143562491350916
COLUMN_ID_FOR_LANDLORD_COMPANY_NAME = 5647162118721412
COLUMN_ID_FOR_EMAIL = 3395362305036164
#COLUMN_ID_FOR_PHONE = 7898961932406660
COLUMN_ID_FOR_RENT_AMOUNT = 6210112072142724
COLUMN_ID_FOR_RENT_INCREMENT = 4802737188589444
COLUMN_ID_FOR_DATE_INCREMENT = 4474124958388100
COLUMN_ID_FOR_SUBLEASE_CONTRACT = 8743386862538628
#COLUMN_ID_FOR_RENT_TAXES = 158400072863620
COLUMN_ID_FOR_GUARANTEE_AMOUNT = 3958312258457476
COLUMN_ID_FOR_CONTRACT_END_DATE = 299137561218948
COLUMN_ID_FOR_ALIQUOT_AMOUNT = 2550937374904196
COLUMN_ID_FOR_WHO_INVOICE = 1425037468061572
COLUMN_ID_FOR_BENEFICIARY_BANK = 5420177937354628
COLUMN_ID_FOR_ACCOUNT_TYPE = 6229487214874500
COLUMN_ID_FOR_BANK_ACCOUNT = 4575894741143428
COLUMN_ID_FOR_SUBLEASE_PERMISSION = 2410199886548868
COLUMN_ID_FOR_WATER_SUBCONDITION = 5787899607076740
COLUMN_ID_FOR_JURISDICTION = 4661999700234116
COLUMN_ID_FOR_WATER_CONDITION = 2216685840060292
COLUMN_ID_FOR_LIGHT_CONDITION = 4468485653745540
COLUMN_ID_FOR_TYPE_PERSON = 2269462398193540
COLUMN_ID_FOR_CONTRACT_DURATION = 2973149839970180
#REPRESENTANTE LEGAL
COLUMN_ID_FOR_REPRESENTATIVE = 4580843801759620 #Si o No
COLUMN_ID_FOR_REPRESENTATIVE_NAME = 2126845517713284
COLUMN_ID_FOR_REPRESENTATIVE_ID = 7984675781037956
COLUMN_ID_FOR_REPRESENTATIVE_TITLE = 1071365894655876
COLUMN_ID_FOR_REPRESENTATIVE_PRONOUN = 1844395583426436

# Cargar la plantilla de contrato
template = DocxTemplate("templated.docx")

# Funciones de ayuda para reemplazar variables

locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

def get_date_parts(date_str):
    if date_str:
        date = datetime.strptime(date_str.split('T')[0], '%Y-%m-%d')
        return date.day, date.strftime('%B'), date.year
    return None, None, None

def format_currency(amount):
    if amount is None or amount == '':
        return "N/A"
    try:
        amount = float(amount)
        words = num2words(amount, to='currency', lang='es').replace('euros', 'dólares').replace('euro', 'dólar').replace('céntimos', 'centavos')
        return f"{words} ({amount:,.2f} USD)"
    except ValueError:
        return amount

def get_title(pronoun, marital= None):
    if marital == "SOLTERO/A":
        titles = {
            'el': 'señor',
            'la': 'señorita',
            'los': 'señores',
            'las': 'señoritas'
        }
        return titles.get(pronoun, '')
    else:
        titles = {
            'el': 'señor',
            'la': 'señora',
            'los': 'señores',
            'las': 'señoras'
        }
        return titles.get(pronoun, '')


def get_arr(pronoun):
    arr = {
        'el': 'EL ARRENDADOR',
        'la': 'LA ARRENDADORA',
        'los': 'LOS ARRENDADORES',
        'las': 'LAS ARRENDADORAS'
    }
    return arr.get(pronoun, '')

def get_MayuscTitle(pronoun):
    arr = {
        'el': 'El',
        'la': 'La',
        'los': 'Los',
        'las': 'Las'
    }
    return arr.get(pronoun, '')

def get_deARR(pronoun):
    deARR = {
        'el': 'del ARRENDADOR',
        'la': 'de la ARRENDADORA',
        'los': 'de los ARRENDADORES',
        'las': 'de las ARRENDADORAS'
    }
    return deARR.get(pronoun, '')

def get_sinPronARR(pronoun):
    arr = {
        'el': 'ARRENDADOR',
        'la': 'ARRENDADORA',
        'los': 'ARRENDADORES',
        'las': 'ARRENDADORAS'
    }
    return arr.get(pronoun, '')

def get_quien(pronoun):
    return 'quien' if pronoun in ['el', 'la'] else 'quienes'

def get_oas(pronoun):
    arr = {
        'el': 'o',
        'la': 'a',
        'los': 'os',
        'las': 'as'
    }
    return arr.get(pronoun, '')

def get_es(pronoun):
    arr = {
        'el': '',
        'la': 'a',
        'los': 'es',
        'las': 'as'
    }
    return arr.get(pronoun, '')

def get_eses(pronoun):
    arr = {
        'el': '',
        'la': '',
        'los': 's',
        'las': 's'
    }
    return arr.get(pronoun, '')

def get_ns(pronoun):
    return '' if pronoun in ['el', 'la'] else 'n'

def get_company_description(company):
    descriptions = {
        'ECONOFARM': 'ECONOFARM S.A. con RUC número 1791715772001',
        'FARCOMED': 'FARMACIAS Y COMISARIATOS DE MEDICINAS S.A. FARCOMED con RUC número 1790710319001',
        'OKIDOKI': 'FARMACIAS Y COMISARIATOS DE MEDICINAS S.A. FARCOMED con RUC número 1790710319001'
    }
    return descriptions.get(company, '')

def get_purpose(company):
    purposes = {
        'ECONOFARM': 'CONSTITUIR UNA FARMACIA, a través de la cual se comercializará productos farmacéuticos, medicinales y de aseo',
        'FARCOMED': 'CONSTITUIR UNA FARMACIA, a través de la cual se comercializará productos farmacéuticos, medicinales y de aseo',
        'OKIDOKI': 'LA IMPLEMENTACIÓN DE UN MINIMARKET, a través del cual se comercializará comida y bebida'
    }
    return purposes.get(company, '')

def get_purpose2(company):
    purposes = {
        'ECONOFARM': 'productos farmacéuticos',
        'FARCOMED': 'productos farmacéuticos',
        'OKIDOKI': 'comida y/o bebida'
    }
    return purposes.get(company, '')

def get_purpose3(company):
    purposes = {
        'ECONOFARM': 'de la farmacia',
        'FARCOMED': 'de la farmacia',
        'OKIDOKI': 'del minimarket'
    }
    return purposes.get(company, '')

def get_bank_prefix(bank_name):
    if not isinstance(bank_name, str):
        return ""
    bank_name_lower = bank_name.lower()
    if "banco" in bank_name_lower:
        return "del"
    elif "produbanco" in bank_name_lower:
        return "de"
    else:
        return "de la"

def get_delARR(pronoun):
    delARR = {
        'el': 'del',
        'la': 'de la',
        'los': 'de los',
        'las': 'de las'
    }
    return delARR.get(pronoun, '')

def procesar_cuenta(cadena):
    # Convertir la cadena a mayúsculas para estandarizar
    if cadena:
        cadena = cadena.upper()
    else:
        cadena = ""
    # Separar la cadena por el guión
    partes = cadena.split('-')
    # Verificar si el formato es CC o CU
    if len(partes) == 2:
        tipo = partes[0]
        numero = partes[1]
        if tipo == 'CC':
            return "la cuenta contrato", numero
        elif tipo == 'CU':
            return "el código único", numero
    # Si no es ninguno de los formatos anteriores, retornar un string vacío y la cadena original
    return "", cadena

def procesar_agua(cadena): #pend
    # Convertir la cadena a mayúsculas para estandarizar
    cadena = cadena.upper()
    # Separar la cadena por el guión
    partes = cadena.split('-')
    # Verificar si el formato es CC o CU
    if len(partes) == 2:
        tipo = partes[0]
        numero = partes[1]
        if tipo == 'P':
            return f"2.- Por el consumo de agua potable, EL ARRENDATARIO cancelará el {numero} mensual del valor total de la planilla durante la vigencia del contrato." 
        elif tipo == 'F':
            return  f"2.- Por el consumo de agua potable, EL ARRENDATARIO cancelará el valor de {numero} mensuales durante la vigencia del contrato."
    # Si no es ninguno de los formatos anteriores, retornar un string vacío y la cadena original
    elif cadena == "A":
        return "2.- El consumo de agua potable está incluido en el pago de la alícuota"
    
    else:
        return cadena
    
# Generar tabla de renta
def generar_tabla_renta(incremento, anioIncremento, canon, fecha_inicio_contrato, plazo):
    # Validación de la fecha
    try:
        fecha_inicio = datetime.strptime(fecha_inicio_contrato, "%Y-%m-%d")
    except ValueError:
        return [f"Error: La fecha de inicio del contrato '{fecha_inicio_contrato}' no tiene el formato correcto o es inválida."]

    # Calcular "hasta" inicial, un año menos un día después
    fecha_hasta_inicial = fecha_inicio + relativedelta(years=1) - relativedelta(days=1)

    # Cálculo del valor inicial del canon
    valor_inicial = canon
    tabla = []

    incremento = incremento.upper()

    if plazo == 0:
        tabla.append({"Desde": fecha_inicio.strftime("%Y-%m-%d"),
                    "Hasta":f"Error: La fecha de fin del contrato es inválida.",
                    "Valor": round(valor_inicial, 2)
            })

    # Si es un incremento fijo
    if incremento.startswith("F-"):
        try:
            porcentaje_incremento = float(incremento.split("-")[1]) / 100  # Obtenemos el porcentaje
        except (ValueError, IndexError):
            return [f"Error: Formato de incremento '{incremento}' es incorrecto."]
        
        anio_inicio_incremento = None if anioIncremento == "NA" else int(anioIncremento)

        for i in range(plazo):
            # Mantener el mismo día y mes, sumando años usando relativedelta
            desde = fecha_inicio + relativedelta(years=i)
            hasta = fecha_hasta_inicial + relativedelta(years=i)

            if anio_inicio_incremento and i < (anio_inicio_incremento - 1):
                valor = valor_inicial
            else:
                # Aplicar el incremento
                valor = valor_inicial * ((1 + porcentaje_incremento) ** (i if anio_inicio_incremento is None else i - anio_inicio_incremento + 2))

            tabla.append({
                "Desde": desde.strftime("%Y-%m-%d"),
                "Hasta": hasta.strftime("%Y-%m-%d"),
                "Valor": round(valor, 2)
            })

    # Si el incremento es por inflación (INEC)
    elif incremento.startswith("I"):
        if anioIncremento == "NA":
            return [f"El canon de arrendamiento será revisado anualmente y será incrementado por acuerdo de las partes en un monto que no exceda el índice de inflación publicado por Instituto Nacional de Estadísticas y Censos (INEC)."]
        else:
            return [f"El canon de arrendamiento será revisado anualmente y será incrementado por acuerdo de las partes en un monto que no exceda el índice de inflación publicado por Instituto Nacional de Estadísticas y Censos (INEC) a partir del Año {int(anioIncremento)}."]

    # Si el incremento es variable (V-)
    elif incremento.startswith("V-"):
        detalle = incremento.split("-", 1)[1]
        detalle = detalle.capitalize()
        return [detalle]  # Retorna el detalle en una lista

    # Para otros casos de error
    else:
        return [f"Valores se deben corregir, revisar de acuerdo a: incremento={incremento}, anioIncremento={anioIncremento}, canon={canon}, fecha_inicio_contrato={fecha_inicio_contrato}"]

    # Retornar la tabla en formato de lista de diccionarios (compatible con docxtpl)
    return tabla

def separar_por_comas_y(texto):
    # Remover espacios innecesarios y convertir a minúsculas para evitar problemas
    texto = texto.lower().strip()
    
    # Reemplazar " y " (espacio y espacio) por una coma
    texto = re.sub(r'\s+y\s+', ',', texto)
    
    # Dividir el texto por comas
    return [nombre.strip().title() for nombre in texto.split(",") if nombre.strip()]

# Variables para datos
data = {}

# Función para redirigir la salida estándar (print) al cuadro de texto en la interfaz
class RedirectText:
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)

    def flush(self):
        pass

def obtener_datos_smartsheet():
    try:
        ss_client = smartsheet.Smartsheet(ACCESS_TOKEN)
        sheet = ss_client.Sheets.get_sheet(SHEET_ID)
        return sheet.rows
    except Exception as e:
        print(f"Error al obtener los datos de Smartsheet: {e}")
        return []

def iniciar_proceso():
    rows = obtener_datos_smartsheet()
    if not rows:
        print("No se encontraron filas en la hoja de Smartsheet.")
        return

    print("Filas disponibles:")
    for idx, row in enumerate(rows):
        for cell in row.cells:
            if cell.column_id == COLUMN_ID_FOR_PROJECT_NAME:
                print(f"Fila {idx + 1}: {cell.value}")

    inicio_btn.config(state=tk.DISABLED)
    fila_input.config(state=tk.NORMAL)
    procesar_btn.config(state=tk.NORMAL)

def procesar_fila():
    project_name = ""
    pronoun = None
    estado_civil = None
    representante = None
    representante_id = None
    representante_cargo = None
    tipo_persona = None
    representante_del = None
    representante_as = None
    canon = None
    medidor_agua = None
    condicion_compartido = None
    incremento = None
    anioIncremento = None
    fecha_inicio_contrato = None
    fecha_fin_contrato = None
    plazo = None
    ArrendadoresCedulas = None
    ArrendadoresNombres = None
    ArrendadorCedula = None
    ArrNombre = None
    CuentaCodigo = ""
    NroLuz = ""
    subleaseContract = ""
    contactoEnvio = None
    texto_repte_tpl2 = ""
    
    try:
        last_row_num = int(fila_input.get()) - 1
        rows = obtener_datos_smartsheet()

        if last_row_num < 0 or last_row_num >= len(rows):
            raise ValueError("Número de fila fuera del rango disponible")

        last_row = rows[last_row_num]
        data = {}

        for cell in last_row.cells:
            if cell.column_id == COLUMN_ID_FOR_CANTON:
                data['Ciudad'] = cell.value
            elif cell.column_id == COLUMN_ID_FOR_DATE:
                fecha_inicio_contrato = cell.value
                data['FechaInicioContrato'] = fecha_inicio_contrato
                dia, mes, anno = get_date_parts(cell.value)
                data['Dia'] = dia
                data['Mes'] = mes
                data['Anno'] = anno
            elif cell.column_id == COLUMN_ID_FOR_CONTACT:
                contactoEnvio = cell.value
            elif cell.column_id == COLUMN_ID_FOR_PRONOUN:
                pronoun = cell.value
                data['PronombreArrendador'] = pronoun
                data['Quien'] = get_quien(pronoun)
                data['ARR'] = get_arr(pronoun)
                data['OAS'] = get_oas(pronoun)
                data['Ns'] = get_ns(pronoun)
                data['ENS'] = (get_ns(pronoun)).upper()
                data['deARR'] = get_deARR(pronoun)
                data['MayuscTitle'] = get_MayuscTitle(pronoun)
                data['sinPronARR'] = get_sinPronARR(pronoun)
                data['arrendador_del'] = get_delARR(pronoun)
                data['es'] = get_es(pronoun)
                data['eses'] = get_eses(pronoun)
            elif cell.column_id == COLUMN_ID_FOR_PROJECT_NAME:
                project_name = cell.value
            elif cell.column_id == COLUMN_ID_FOR_COMPANY:
                data['CIA'] = get_company_description(cell.value)
                data['Proposito'] = get_purpose(cell.value)
                data['Purpose2'] = get_purpose2(cell.value)
                data['Purpose3'] = get_purpose3(cell.value)
            elif cell.column_id == COLUMN_ID_FOR_PROVINCE:
                data['Provincia'] = cell.value.title()
            elif cell.column_id == COLUMN_ID_FOR_PARISH:
                data['Parroquia'] = cell.value.title()
            elif cell.column_id == COLUMN_ID_FOR_ADDRESS:
                data['Direccion'] = cell.value.title()
            elif cell.column_id == COLUMN_ID_FOR_METRAJE:
                data['m2'] = cell.value
            elif cell.column_id == COLUMN_ID_FOR_LANDLORD_NAME:
                if cell.value is not None:
                    data['ArrendadorNombre'] = cell.value.title()
                    ArrNombre = cell.value.title()
                    ArrendadoresNombres = separar_por_comas_y(cell.value)
            elif cell.column_id == COLUMN_ID_FOR_LANDLORD_ID:
                if cell.value is not None:
                    data['ArrendadorCedula'] = str(cell.value).replace('.0', '')
                    ArrendadorCedula = str(cell.value).replace('.0', '')
                    ArrendadoresCedulas = separar_por_comas_y(str(cell.value))
            elif cell.column_id == COLUMN_ID_FOR_LANDLORD_MARITAL_STATUS:
                estado_civil = cell.value #ojo
                data['TituloArrendador'] = get_title(pronoun, estado_civil)
            elif cell.column_id == COLUMN_ID_FOR_TYPE_PERSON:
                tipo_persona = cell.value
                data['tipoPersona'] = cell.value
            elif cell.column_id == COLUMN_ID_FOR_LANDLORD_COMPANY_NAME:
                if cell.value is not None:
                    data['NombreEmpresa'] = cell.value
                else:
                    data['NombreEmpresa'] = ""
            elif cell.column_id == COLUMN_ID_FOR_EMAIL:
                data['Correo'] = cell.value
            elif cell.column_id == COLUMN_ID_FOR_SUBLEASE_CONTRACT:
                subleaseContract = cell.value
            elif cell.column_id == COLUMN_ID_FOR_RENT_AMOUNT:
                canon = float(cell.value)
                data['CANON'] = format_currency(canon)
            elif cell.column_id == COLUMN_ID_FOR_RENT_INCREMENT:
                incremento = cell.value
            elif cell.column_id == COLUMN_ID_FOR_DATE_INCREMENT:
                anioIncremento = cell.value
                if type(cell.value) == float:
                    data['anioIncremento'] = int(cell.value)
                else:
                    data['anioIncremento'] = ""
            elif cell.column_id == COLUMN_ID_FOR_GUARANTEE_AMOUNT:
                data['Garantia'] = format_currency(cell.value)
            elif cell.column_id == COLUMN_ID_FOR_CONTRACT_END_DATE:
                fecha_fin_contrato = cell.value
                data['FechaFinContrato'] = fecha_fin_contrato
            elif cell.column_id == COLUMN_ID_FOR_BANK_ACCOUNT:
                data['CtaBancaria'] = str(cell.value).replace('.0', '')
            elif cell.column_id == COLUMN_ID_FOR_SUBLEASE_PERMISSION:
                data['Subarriendo'] = cell.value
            elif cell.column_id == COLUMN_ID_FOR_LIGHT_CONDITION:
                medidor_luz = cell.value
                if medidor_luz == "Uso medidor que se encuentra en PDV de manera provisional hasta gestionar nuevo medidor":
                    texto_luz_template = ("1.- Para el servicio básico de electricidad se gestionará un nuevo medidor por parte del ARRENDATARIO"
                                        " y lo cancelará mensualmente, así mismo una vez termine la vigencia del contrato "
                                        "deberá realizar el trámite respectivo para su anulación.")
                else:
                    texto_luz_template = ("1.- El servicio básico de electricidad con {{CuentaCodigo}} No. {{NroLuz}}, EL ARRENDATARIO lo cancelará mensualmente.")
            elif cell.column_id == COLUMN_ID_FOR_LIGHT_METER:
                CuentaCodigo, NroLuz = procesar_cuenta(cell.value)

            elif cell.column_id == COLUMN_ID_FOR_WATER_CONDITION:
                medidor_agua = cell.value

            elif cell.column_id == COLUMN_ID_FOR_WATER_METER:
                data['NroAgua'] = str(cell.value).replace('.0', '')
            
            elif cell.column_id == COLUMN_ID_FOR_ALIQUOT_AMOUNT:
                data['Alicuota'] = cell.value
                if isinstance(cell.value, (int, float)) and cell.value != 0:
                    texto_alicuota_template = ("3.- El pago del valor de la alícuota, será de $ {{ValorAlicuota}} que serán cancelados directamente por EL ARRENDATARIO a la administración del edificio.")
                else:
                    texto_alicuota_template = ""
            elif cell.column_id == COLUMN_ID_FOR_WHO_INVOICE: 
                data['NombreFactura'] = cell.value.title()
            elif cell.column_id == COLUMN_ID_FOR_BENEFICIARY_BANK:
                banco = cell.value
                data['dBanco'] = get_bank_prefix(banco)
                data['Banco'] = banco
            elif cell.column_id == COLUMN_ID_FOR_ACCOUNT_TYPE: 
                data['TipoCuenta'] = cell.value
            elif cell.column_id == COLUMN_ID_FOR_JURISDICTION: 
                data['Jurisdiccion'] = cell.value
            elif cell.column_id == COLUMN_ID_FOR_WATER_SUBCONDITION:
                condicion_compartido = cell.value
                if medidor_agua == "Uso medidor que se encuentra en PDV de manera provisional hasta gestionar nuevo medidor":
                    texto_agua_template = ("2.- Para el servicio básico de agua potable se gestionará un nuevo medidor por parte del ARRENDATARIO"
                                        " y lo cancelará mensualmente, así mismo una vez termine la vigencia del contrato "
                                        "deberá realizar el trámite respectivo para su anulación.")
                    template_agua = Template(texto_agua_template)
                    data['TextoAgua'] = template_agua.render(NroAgua=data.get('NroAgua', ''))
                elif medidor_agua == "Propio exclusivo del PDV":
                    texto_agua_template = ("2.- El consumo de agua potable con la cuenta No. {{NroAgua}}, EL ARRENDATARIO lo cancelará mensualmente.")
                    template_agua = Template(texto_agua_template)
                    data['TextoAgua'] = template_agua.render(NroAgua=data.get('NroAgua', ''))
                else:
                    data['TextoAgua']= procesar_agua(condicion_compartido)
            elif cell.column_id == COLUMN_ID_FOR_CONTRACT_DURATION:
                plazo = int(cell.value)
                data['Plazo'] = int(cell.value)
                decimal_plazo = round((cell.value - int(cell.value)), 1)
                if int(cell.value) == 0:
                    data['plazo_correcto'] = "POR FAVOR SE DEBE VALIDAR LOS PLAZOS INGRESADOS EN EL CONTRATO"
                elif decimal_plazo > 0:
                    data['plazo_correcto'] = "POR FAVOR SE DEBE VALIDAR LOS PLAZOS INGRESADOS EN EL CONTRATO" #durac
                else:
                    data['plazo_correcto'] = ""
            elif cell.column_id == COLUMN_ID_FOR_REPRESENTATIVE:
                representante = cell.value
                data['TieneRepresentante'] = cell.value
            elif cell.column_id == COLUMN_ID_FOR_REPRESENTATIVE_PRONOUN:
                data['representante_pronoun'] = cell.value    
                data['representante_titulo'] = get_title(cell.value)
                data['ARR_REP'] = get_arr(cell.value)
                data['Quien_REP'] = get_quien(cell.value) #OJO pendiente definir si se va a usar
                data['OAS_REP'] = get_oas(cell.value)
                data['Ns_REP'] = get_ns(cell.value)
                data['deARR_REP'] = get_deARR(cell.value)
                data['sinPronARR_REP'] = get_sinPronARR(cell.value)
                data['eses_REP'] = get_eses(cell.value)
                data['ENS_REP'] = get_ns(cell.value).upper()
                representante_del = get_delARR(cell.value)
                representante_as = get_es(cell.value)
            elif cell.column_id == COLUMN_ID_FOR_REPRESENTATIVE_NAME:
                if cell.value:
                    data['representante_nombre'] = cell.value.title()
                if ArrNombre == None:
                    ArrNombre = cell.value
            elif cell.column_id == COLUMN_ID_FOR_REPRESENTATIVE_ID:
                representante_id = str(cell.value).replace('.0', '')
                data['representante_id'] = str(cell.value).replace('.0', '')
            elif cell.column_id == COLUMN_ID_FOR_REPRESENTATIVE_TITLE:
                representante_cargo = cell.value
                if representante == "Si":
                    if tipo_persona == "Jurídica":
                        texto_repte_tpl = "{{representante_pronoun}} {{representante_titulo}} {{representante_nombre}} de nacionalidad ecuatoriana, mayor de edad, portador{{representante_as}} de la cédula de ciudadanía No. {{representante_id}} en calidad de {{representante_cargo}} de {{NombreEmpresa}}"
                        texto_repte_tpl2 = "{{NombreEmpresa}}"
                    else:
                        texto_repte_tpl = "{{representante_pronoun}} {{representante_titulo}} {{representante_nombre}} de nacionalidad ecuatoriana, mayor de edad, portador{{representante_as}} de la cédula de ciudadanía No. {{representante_id}} en calidad de {{representante_cargo}} {{arrendador_del}} {{TituloArrendador}} {{ArrendadorNombre}}"
                        texto_repte_tpl2 = "{{MayuscTitle}} {{TituloArrendador}} {{ArrendadorNombre}}"
                else:
                    texto_repte_tpl = "{{PronombreArrendador}} {{TituloArrendador}} {{ArrendadorNombre}} de nacionalidad ecuatoriana, mayor{{es}} de edad, portador{{es}} de la{{s}} cédula{{s}} de ciudadanía No. {{ArrendadorCedula}}"
                    texto_repte_tpl2 = "{{ARR}}"
        
        
        print("Procesando la información de la fila seleccionada...")

        data['TestTabla'] = ""


        if incremento != "S":
            data['TestTabla'] = generar_tabla_renta(incremento, anioIncremento, canon, fecha_inicio_contrato, plazo)

        if ArrendadoresNombres and ArrendadoresCedulas:
            data['Varios'] = [{'Nombre': nombre, 'Cedula': cedula} for nombre, cedula in zip(ArrendadoresNombres, ArrendadoresCedulas)]
        else:
            data['Varios'] = [{'Nombre': ArrNombre, 'Cedula': ArrendadorCedula}]

        data['incremento'] = incremento
        data['TipoPersonaNat'] = tipo_persona

        #Jinja2 para variables anidadas
        template_luz = Template(texto_luz_template)
        data['TextoLuz'] = template_luz.render(NroLuz = NroLuz, CuentaCodigo = CuentaCodigo)


        template_persona = Template(texto_repte_tpl)
        data["TipoPersona"] = template_persona.render(PronombreArrendador=data.get("PronombreArrendador", ""), TituloArrendador=data.get("TituloArrendador", ""), ArrendadorNombre=data.get("ArrendadorNombre",""), ArrendadorCedula=data.get("ArrendadorCedula", ""), NombreEmpresa=data.get("NombreEmpresa", ""),representante_del = representante_del, representante_pronoun= data.get("representante_pronoun"), arrendador_del = data.get("arrendador_del", ""), representante_titulo = data.get("representante_titulo"), representante_nombre= data.get("representante_nombre"), es = data.get("es", ""), s = data.get("eses", ""), representante_cargo = representante_cargo, representante_id = representante_id, representante_as = representante_as)

        template_persona2 = Template(texto_repte_tpl2)        
        data["TipoPersona2"] = template_persona2.render( ArrendadorNombre = data.get("ArrendadorNombre",""), NombreEmpresa=data.get("NombreEmpresa", ""), ARR = data.get("ARR", ""), TituloArrendador = data.get("TituloArrendador", ""), MayuscTitle = data.get("MayuscTitle", ""))

        template_alicuota = Template(texto_alicuota_template)
        data['TextoAlicuota'] = template_alicuota.render(ValorAlicuota = data.get("Alicuota", ""))

        # Reemplazar valores en el documento
        template.render(data)
        output_filename = f"ADMINISTRACION DE CONTRATO {project_name.upper()} - {contactoEnvio}.docx"
        if subleaseContract == "NO":
            template.save(output_filename)
            print(f"Documento generado: {project_name}")
        else:
            print(f"Documento no generado por Subarriendo")

        fila_input.delete(0, tk.END)    
    except ValueError as e:
        messagebox.showerror("Error", f"Entrada inválida: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"Error al procesar la fila: {e}")

def validar_entrada(event):
    if not fila_input.get().isdigit():
        messagebox.showwarning("Advertencia", "Por favor, ingrese un número válido.")

# Crear la ventana principal
root = tk.Tk()
root.title("Generador de Contratos Borrador")

# Crear widgets
output_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=60, height=20)
output_text.pack(pady=10)

# Crear un frame para contener el input y los botones
frame = tk.Frame(root)
frame.pack(pady=5)  # Empaquetar el frame con un margen en Y

# Crear la caja de entrada (input) y el botón en el mismo nivel
fila_input = tk.Entry(frame, width=40)
fila_input.pack(side=tk.LEFT, padx=5)  # Alinear a la izquierda con un margen en X
fila_input.bind("<FocusOut>", validar_entrada)

procesar_btn = tk.Button(frame, text="Procesar Fila", command=procesar_fila, state=tk.DISABLED)  # Aquí iría tu función de procesamiento
procesar_btn.pack(side=tk.LEFT, padx=5)  # Alinear a la izquierda con un margen en X

inicio_btn = tk.Button(frame, text="Inicio", command=iniciar_proceso)  # Aquí iría tu función de inicio
inicio_btn.pack(side=tk.LEFT, padx=5)  # Alinear a la izquierda con un margen en X

# Redirigir la salida estándar al cuadro de texto
redirect_text = RedirectText(output_text)
sys.stdout = redirect_text

# Ejecutar el bucle de eventos de Tkinter
root.mainloop()