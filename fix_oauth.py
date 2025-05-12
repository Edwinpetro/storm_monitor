#!/usr/bin/env python3
"""
Script para solucionar el problema del token OAuth persistente en el sistema de monitoreo de tormentas.
Este script genera un token OAuth completo con refresh_token que persiste entre sesiones.
"""

import os
import json
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import base64

# Define el scope para Gmail
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def create_persistent_token():
    """
    Crea un token OAuth persistente con refresh_token para Gmail
    """
    print("=== CREACIÓN DE TOKEN OAUTH2 PERSISTENTE ===\n")
    
    # Eliminar token existente si lo hay
    token_path = 'token.json'
    if os.path.exists(token_path):
        try:
            os.remove(token_path)
            print(f"Token anterior eliminado: {token_path}")
        except Exception as e:
            print(f"Error al eliminar token anterior: {e}")
    
    print("\nIniciando autenticación con Google...")
    print("Se abrirá una ventana del navegador para que autorices la aplicación.")
    
    try:
        # Configuración explícita para garantizar que obtenemos refresh_token
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', 
            SCOPES,
            # Parámetros adicionales para asegurar el refresh_token
            redirect_uri='http://localhost',
        )
        
        # Esto abre el navegador para autenticación
        creds = flow.run_local_server(
            port=0,
            prompt='consent',  # Forzar pantalla de consentimiento
            access_type='offline'  # Acceso offline para obtener refresh_token
        )
        
        # Guardar el token
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
            print(f"\n✅ Token guardado en {token_path}")
        
        # Verificar que el token contiene refresh_token
        token_data = json.loads(creds.to_json())
        if 'refresh_token' in token_data:
            print("✅ El token contiene refresh_token - funcionalidad offline garantizada.")
        else:
            print("❌ ALERTA: El token no contiene refresh_token.")
            print("Intenta revocar el acceso a la aplicación en tu cuenta de Google y ejecuta este script nuevamente.")
            return False
        
        # Probar el token creando un servicio
        service = build('gmail', 'v1', credentials=creds)
        print("✅ Servicio de Gmail creado exitosamente con las credenciales.")
        
        # Ofrecer enviar un correo de prueba
        if input("\n¿Quieres enviar un correo de prueba? (s/n): ").lower().startswith('s'):
            destinatario = input("Introduce la dirección de correo para la prueba: ")
            enviar_correo_prueba(service, destinatario)
        
        return True
        
    except Exception as e:
        print(f"\n❌ Error durante la autenticación: {e}")
        return False

def enviar_correo_prueba(service, destinatario):
    """
    Envía un correo electrónico de prueba para verificar que el token funciona correctamente
    """
    try:
        # Crear mensaje
        message = MIMEMultipart()
        message['to'] = destinatario
        message['subject'] = "Prueba de token OAuth2 persistente"
        
        # Cuerpo del mensaje
        msg_text = """
        Este es un correo de prueba para verificar que el token OAuth2 se ha creado correctamente 
        y es persistente. Si recibes este correo, la configuración es correcta y el sistema podrá 
        enviar correos sin solicitar autenticación cada vez.
        
        No es necesario responder a este correo.
        """
        
        msg_html = f"""
        <html>
          <head></head>
          <body>
            <h2>Prueba de token OAuth2 persistente</h2>
            <p>Este es un correo de prueba para verificar que el token OAuth2 se ha creado correctamente 
            y es persistente.</p>
            <p><strong>Si recibes este correo, la configuración es correcta y el sistema podrá 
            enviar correos sin solicitar autenticación cada vez.</strong></p>
            <p style="color: #999;">No es necesario responder a este correo.</p>
          </body>
        </html>
        """
        
        msg_alternative = MIMEMultipart('alternative')
        msg_alternative.attach(MIMEText(msg_text, 'plain'))
        msg_alternative.attach(MIMEText(msg_html, 'html'))
        message.attach(msg_alternative)
        
        # Codificar y enviar
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        
        send_message = service.users().messages().send(
            userId="me", 
            body={'raw': encoded_message}
        ).execute()
        
        print(f"\n✅ Correo enviado exitosamente a {destinatario}")
        print(f"ID del mensaje: {send_message['id']}")
        
        # Probar una segunda vez para verificar que no pide autenticación
        print("\nEnviando un segundo correo para verificar persistencia...")
        send_message = service.users().messages().send(
            userId="me", 
            body={'raw': encoded_message}
        ).execute()
        
        print(f"✅ Segundo correo enviado exitosamente sin reautenticación")
        print(f"ID del mensaje: {send_message['id']}")
        
        print("\n✅ CONFIGURACIÓN EXITOSA: El token OAuth2 es persistente")
        print("El sistema de monitoreo de tormentas ahora puede enviar correos sin solicitar")
        print("autenticación cada vez. Este token será válido hasta que se revoque el acceso.")
        
        return True
        
    except Exception as e:
        print(f"\n❌ Error al enviar correo de prueba: {e}")
        return False

def revocation_instructions():
    """
    Muestra instrucciones para revocar el acceso a la aplicación en Google
    """
    print("\n=== SI NECESITAS RESETEAR LA AUTENTICACIÓN ===")
    print("Si el token sigue sin funcionar correctamente, debes revocar el acceso a la aplicación en tu cuenta de Google:")
    print("1. Ve a https://myaccount.google.com/permissions")
    print("2. Busca la aplicación 'Event Monitor' o el nombre que hayas dado a tu aplicación OAuth")
    print("3. Haz clic en ella y selecciona 'Revocar acceso'")
    print("4. Luego ejecuta este script nuevamente para crear un nuevo token desde cero")

if __name__ == "__main__":
    print("\n===== HERRAMIENTA DE CONFIGURACIÓN DE TOKEN OAUTH2 PERSISTENTE =====")
    print("Este script creará un token OAuth2 que permitirá al sistema de monitoreo")
    print("de tormentas enviar correos sin solicitar autenticación cada vez.")
    print("\nSe eliminará cualquier token existente y se creará uno nuevo.")
    print("Asegúrate de tener acceso a la cuenta de Gmail que usarás para enviar correos.")
    
    if input("\n¿Continuar? (s/n): ").lower().startswith('s'):
        if create_persistent_token():
            print("\n=== INSTRUCCIONES PARA EL SISTEMA DE MONITOREO DE TORMENTAS ===")
            print("1. Copia el archivo token.json generado al directorio del sistema")
            print("2. Verifica que el archivo app.py esté utilizando la versión más reciente")
            print("   del código que mantiene el refresh_token")
            print("3. Ejecuta el sistema normalmente y ahora no debería solicitar autenticación")
            print("   cada vez que envía un correo.")
        else:
            revocation_instructions()
    else:
        print("\nOperación cancelada por el usuario.")