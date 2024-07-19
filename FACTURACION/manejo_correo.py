import smtplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os
import time
import imaplib

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


def gestionar_correos_enviados(imap_server='imap.outlook.com', email_user='emisor', email_pass='contraseña', print_func=print):
    def guardar_en_archivo(linea, nombre_archivo='correos_enviados.txt'):
        try:
            with open(nombre_archivo, 'a') as archivo:
                archivo.write(linea + '\n')
        except Exception as e:
            print_func(f"Error al guardar en el archivo: {e}")

    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_user, email_pass)
        mail.select("/Sent")  # Carpeta de correos enviados en Outlook

        result, data = mail.search(None, 'ALL')
        correo_ids = data[0].split()

        for correo_id in correo_ids:
            result, mensaje_data = mail.fetch(correo_id, '(RFC822)')
            mensaje = email.message_from_bytes(mensaje_data[0][1])
            asunto = mensaje['subject']
            mensaje_guardar = f"Título del correo enviado: {asunto}"
            print_func(mensaje_guardar)
            guardar_en_archivo(mensaje_guardar)

            # Elimina el correo
            mail.store(correo_id, '+FLAGS', '\\Deleted')

        mail.expunge()  # Borra físicamente los correos marcados para eliminación
        mail.logout()
    except Exception as e:
        error_message = f"Error al gestionar correos enviados: {e}"
        print_func(error_message)
        guardar_en_archivo(error_message)

"""def gestionar_correos_enviados(imap_server = 'imap.outlook.com', email_user = emisor, email_pass = contraseña, print_func=print):
    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_user, email_pass)
        mail.select("/Sent")  # Carpeta de correos enviados en Outlook

        result, data = mail.search(None, 'ALL')
        correo_ids = data[0].split()

        for correo_id in correo_ids:
            result, mensaje_data = mail.fetch(correo_id, '(RFC822)')
            mensaje = email.message_from_bytes(mensaje_data[0][1])
            asunto = mensaje['subject']
            print_func(f"Título del correo enviado: {asunto}")

            # Elimina el correo
            mail.store(correo_id, '+FLAGS', '\\Deleted')

        mail.expunge()  # Borra físicamente los correos marcados para eliminación
        mail.logout()
    except Exception as e:
        print_func(f"Error al gestionar correos enviados: {e}")"""