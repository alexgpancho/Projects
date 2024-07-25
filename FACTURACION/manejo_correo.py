#Test GH
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os
import time
import configparser
from azure.identity import DeviceCodeCredential
import httpx
import asyncio

#emisor = 'facturas_gpf_sierra@outlook.com'
#contraseña = 'cnvzpbgggmtdqiry'

emisor = 'facturas_gpf@outlook.com'
contraseña = 'lleibtocysmvsnko'

def enviar_correo(asunto, cuerpo, destinatario, cc, adjuntos=[], print_func=print, max_reintentos=3):

    mensaje = MIMEMultipart()
    mensaje['From'] = emisor
    mensaje['To'] = destinatario
    mensaje['Cc'] = cc
    mensaje['Subject'] = asunto
    mensaje.attach(MIMEText(cuerpo, 'plain'))
    destinatarios = [destinatario] + cc.split(',')

    for archivo in adjuntos:
        parte = MIMEBase('application', 'octet-stream')
        with open(archivo, 'rb') as file:
            parte.set_payload(file.read())
        encoders.encode_base64(parte)
        parte.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(archivo)}")
        mensaje.attach(parte)

    intentos = 0
    enviado = False
    while intentos < max_reintentos and not enviado:
        try:
            server = smtplib.SMTP('smtp.office365.com', 587)
            server.starttls()
            server.login(emisor, contraseña)
            text = mensaje.as_string()
            server.sendmail(emisor, destinatarios, text)
            server.quit()
            enviado = True
            print_func(f"Correo enviado a {destinatarios} en el intento {intentos + 1}")
        except Exception as e:
            intentos += 1
            print_func(f"Error al enviar correo: {e}. Reintentando... ({intentos}/{max_reintentos})")
            time.sleep(5)

    if not enviado:
        print_func(f"Fallo al enviar el correo a {destinatarios} después de {max_reintentos} intentos.")

async def autenticar(print_func):
    try:
        config = configparser.ConfigParser()
        config.read(['config.cfg', 'config.dev.cfg'])
        azure_settings = config['azure']

        client_id = azure_settings['clientId']
        tenant_id = azure_settings['tenantId']
        graph_scopes = azure_settings['graphUserScopes'].split(' ')

        # Callback to handle the device code presentation
        def print_code_callback(verification_uri, user_code, expires_on):
            print_func(f"Para iniciar sesión ve a {verification_uri} y entra el código {user_code} para autenticarte.")
            print_func(f"El código expira en: {expires_on.strftime('%Y-%m-%d %H:%M:%S')}")

        device_code_credential = DeviceCodeCredential(
            client_id=client_id,
            tenant_id=tenant_id,
            prompt_callback=print_code_callback
        )

        # Authenticate and obtain the access token
        access_token = device_code_credential.get_token(*graph_scopes).token
        return access_token
    except Exception as e:
        print_func(f"Error durante la autenticación: {str(e)}")
        raise

async def gestionar_correos_enviados(print_func, access_token):
    try:
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Accept': 'application/json'
        }

        async with httpx.AsyncClient() as client:
            # Obtener carpetas de correo
            response = await client.get('https://graph.microsoft.com/v1.0/me/mailFolders', headers=headers)
            response.raise_for_status()
            folders = response.json()['value']

            sent_items_folder_id = None
            for folder in folders:
                if folder['displayName'] == 'Elementos enviados':
                    sent_items_folder_id = folder['id']
                    break

            if sent_items_folder_id:
                # Listar mensajes en la carpeta de elementos enviados
                messages_response = await client.get(f'https://graph.microsoft.com/v1.0/me/mailFolders/{sent_items_folder_id}/messages', headers=headers)
                messages_response.raise_for_status()
                messages = messages_response.json()['value']

                for message in messages:
                    # Eliminar cada mensaje
                    delete_response = await client.delete(f'https://graph.microsoft.com/v1.0/me/messages/{message["id"]}', headers=headers)
                    delete_response.raise_for_status()

                print_func("Correos enviados eliminados.")
            else:
                print_func("No se encontró la carpeta de Elementos enviados.")
    except httpx.HTTPStatusError as e:
        print_func(f"Error HTTP: {e.response.status_code} - {e.response.text}")
    except Exception as e:
        print_func(f"Error: {str(e)}")