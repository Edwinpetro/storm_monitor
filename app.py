import geopandas as gpd
import pandas as pd
from shapely.geometry import Point
from shapely.ops import unary_union
from datetime import datetime, timedelta, timezone
import os
import plotly.graph_objects as go
import plotly.io as pio
import json
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from dotenv import load_dotenv

# Cargar variables de entorno desde archivo .env
load_dotenv()

import app_functions as f  # Import custom functions from app_functions module

# Define the Gmail API scopes needed
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def get_gmail_service():
    """
    Authenticate and return Gmail API service using OAuth 2.0
    """
    creds = None
    token_path = 'token.json'
    
    # Load existing token if available
    if os.path.exists(token_path):
        try:
            with open(token_path, 'r') as f:
                creds = Credentials.from_authorized_user_info(json.load(f))
        except (ValueError, json.JSONDecodeError) as e:
            # Si el token existe pero no es válido, lo eliminamos para regenerarlo
            os.remove(token_path)
            creds = None
            print(f"Token inválido, se generará uno nuevo. Error: {e}")
    
    # Check if credentials are invalid or don't exist
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                # Guardar el token actualizado
                with open(token_path, 'w') as token:
                    token.write(creds.to_json())
            except Exception as e:
                print(f"Error al refrescar el token: {e}")
                # Si falla el refresh, forzamos nueva autenticación
                creds = None
        
        if not creds:
            # Get credentials from the client secrets file
            try:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', 
                    SCOPES,
                    redirect_uri='http://localhost',
                )
                creds = flow.run_local_server(
                    port=0,
                    prompt='consent',  # Forzar pantalla de consentimiento
                    access_type='offline'  # Acceso offline para obtener refresh_token
                )
                
                # Save credentials for future use
                with open(token_path, 'w') as token:
                    token.write(creds.to_json())
                
                # Verificar que el token contiene refresh_token
                token_data = json.loads(creds.to_json())
                if 'refresh_token' not in token_data:
                    print("ADVERTENCIA: El token no contiene refresh_token. La autenticación se solicitará en cada ejecución.")
                    print("Ejecuta fix_oauth.py para resolver este problema.")
            except Exception as e:
                print(f"Error en el proceso de autenticación: {e}")
                raise
    
    # Build and return the Gmail service
    return build('gmail', 'v1', credentials=creds)

def send_email_oauth(to, subject, body_html, attachment_df=None, cc=None, bcc=None, sender=None):
    """
    Send an email using Gmail API with OAuth 2.0 authentication
    """
    try:
        # Get Gmail service
        service = get_gmail_service()
        
        # Get sender email from environment variables or use provided sender
        sender_email = os.getenv('SENDER_EMAIL_DISPLAY', 'Lockton Storm Monitor <analistadatoslockton@gmail.com>')
        
        # Create message container
        message = MIMEMultipart('related')
        message['to'] = ', '.join(to) if isinstance(to, list) else to
        message['subject'] = subject
        message['from'] = sender_email
        
        if cc:
            message['cc'] = ', '.join(cc) if isinstance(cc, list) else cc
        if bcc:
            message['bcc'] = ', '.join(bcc) if isinstance(bcc, list) else bcc
        
        # Add HTML body
        msgAlternative = MIMEMultipart('alternative')
        message.attach(msgAlternative)
        msgAlternative.attach(MIMEText(body_html, 'html'))
        
        # Attach DataFrame as CSV if provided
        if attachment_df is not None and not attachment_df.empty:
            attachment = MIMEApplication(attachment_df.to_csv(index=False))
            attachment['Content-Disposition'] = 'attachment; filename="data.csv"'
            message.attach(attachment)
        
        # Encode message
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        
        # Send message
        send_message = service.users().messages().send(userId='me', body={'raw': raw_message}).execute()
        
        print(f"Email sent. Message ID: {send_message['id']}")
        return send_message
    
    except HttpError as error:
        print(f"An error occurred: {error}")
        return None

def send_map_via_email_oauth(fig, recipients, subject, body_html, cc=None, bcc=None, sender=None):
    """
    Send a Plotly figure map via email using OAuth
    """
    try:
        # Get Gmail service
        service = get_gmail_service()
        
        # Get sender email from environment variables or use provided sender
        sender_email = os.getenv('SENDER_EMAIL_DISPLAY', 'Lockton Storm Monitor <analistadatoslockton@gmail.com>')
        
        # Create message container
        message = MIMEMultipart('related')
        message['to'] = ', '.join(recipients) if isinstance(recipients, list) else recipients
        message['subject'] = subject
        message['from'] = sender_email
        
        if cc:
            message['cc'] = ', '.join(cc) if isinstance(cc, list) else cc
        if bcc:
            message['bcc'] = ', '.join(bcc) if isinstance(bcc, list) else bcc
        
        # Add HTML body
        msgAlternative = MIMEMultipart('alternative')
        message.attach(msgAlternative)
        msgAlternative.attach(MIMEText(body_html, 'html'))
        
        # Save figure to HTML file temporarily
        temp_html_path = "temp_map.html"
        fig.write_html(temp_html_path)
        
        # Attach HTML file
        with open(temp_html_path, 'rb') as f:
            attachment = MIMEApplication(f.read(), _subtype='html')
            attachment['Content-Disposition'] = f'attachment; filename="storm_map.html"'
            message.attach(attachment)
        
        # Remove temporary file
        os.remove(temp_html_path)
        
        # Encode message
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        
        # Send message
        send_message = service.users().messages().send(userId='me', body={'raw': raw_message}).execute()
        
        print(f"Map email sent. Message ID: {send_message['id']}")
        return send_message
    
    except HttpError as error:
        print(f"An error occurred: {error}")
        return None

def load_html_template(template_name):
    """
    Load HTML template from file and return as string
    """
    try:
        # Intentar cargar desde la carpeta templates
        template_path = os.path.join(os.path.dirname(__file__), 'templates', template_name)
        
        with open(template_path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        print(f"Error loading template from templates folder: {e}")
        try:
            # Intentar cargar directamente desde el directorio actual
            template_path = os.path.join(os.path.dirname(__file__), template_name)
            with open(template_path, 'r', encoding='utf-8') as f:
                return f.read()
        except Exception as e2:
            print(f"Error loading template from current directory: {e2}")
            # Fallback to basic HTML if template not found
            if 'no_storms' in template_name:
                return """
                <html><body>
                <h1>No Active Storms</h1>
                <p>Team,</p>
                <p>There are currently no recently active storms.</p>
                <p>Best regards,<br>Lockton Storm Monitor</p>
                </body></html>
                """
            elif 'no_impacts' in template_name:
                return """
                <html><body>
                <h1>No Portfolio Impacts</h1>
                <p>Team,</p>
                <p>There are recently active storms: {STORMS}, but none whose cone of uncertainty come within ~70km of intercepting any geometry in the client portfolio.</p>
                <p>Best regards,<br>Lockton Storm Monitor</p>
                </body></html>
                """
            elif 'forecast_map' in template_name:
                return """
                <html><body>
                <h1>Forecast Map - {STORM}</h1>
                <p>Team,</p>
                <p>Please find the attached forecast map for <b>{CLIENT_NAME}</b> for <b>{STORM}</b> based on the Official Forecast made on <b>{TIMESTAMP}</b>.</p>
                <p>Best regards,<br>Lockton Storm Monitor</p>
                </body></html>
                """
            elif 'errors_email' in template_name:
                return """
                <html><body>
                <h1>Processing Errors</h1>
                <p>Alvaro,</p>
                <p>Some errors where encountered while processing ({STORMS}) Please check the data attached.</p>
                <p>Error messages: {ERROR_MESSAGES}</p>
                <p>Best regards,<br>Lockton Storm Monitor</p>
                </body></html>
                """
            else:  # error_email
                return """
                <html><body>
                <h1>Error Alert</h1>
                <p>Alvaro,</p>
                <p>An error occurred in the Storm Monitor script:</p>
                <p>{ERROR}</p>
                <p>Best regards,<br>Lockton Storm Monitor</p>
                </body></html>
                """

def send_no_storms_email(recipients):
    """
    Send email notification when there are no active storms
    """
    subject = 'No Active Storms - Lockton Event Monitor'
    
    # Load HTML template
    html_template = load_html_template('no_storms_email.html')
    
    # Send email
    send_email_oauth(recipients, subject, html_template)

def send_errors_email(recipients, storms, error_msgs, error_data):
    """
    Send email notification about errors during processing
    """
    storms_str = ', '.join(storms)
    error_msgs_str = ', '.join(error_msgs)
    
    subject = 'Error Alert - Lockton Storm Processing'
    
    # Load HTML template
    html_template = load_html_template('errors_email.html')
    
    # Replace placeholders
    html_body = html_template.replace('{STORMS}', storms_str).replace('{ERROR_MESSAGES}', error_msgs_str)
    
    # Send email with attachment
    send_email_oauth(recipients, subject, html_body, error_data)

def send_no_impacts_email(recipients, storms):
    """
    Send email notification when no client locations are impacted
    """
    storms_str = ', '.join(storms)
    subject = 'No Portfolio Impacts - Lockton Event Monitor'
    
    # Load HTML template
    html_template = load_html_template('no_impacts_email.html')
    
    # Replace placeholders
    html_body = html_template.replace('{STORMS}', storms_str)
    
    # Send email
    send_email_oauth(recipients, subject, html_body)

def send_forecast_map_email(fig, storm, client_name, timestamp_utc, recipients):
    """
    Send forecast map via email
    """
    subject = f'Storm Alert: {storm} - {client_name}'
    
    # Load HTML template
    html_template = load_html_template('forecast_map_email.html')
    
    # Replace placeholders
    html_body = html_template.replace('{STORM}', storm) \
                            .replace('{CLIENT_NAME}', client_name) \
                            .replace('{TIMESTAMP}', timestamp_utc.strftime('%Y-%m-%d %H:%M UTC'))
    
    # Send email with map attachment
    send_map_via_email_oauth(
        fig, recipients, subject, html_body, cc=[], bcc=[]
    )

def send_error_email(error, recipients):
    """
    Send email notification about general script errors
    """
    subject = 'System Error - Lockton Event Monitor'
    
    # Load HTML template
    html_template = load_html_template('error_email.html')
    
    # Replace placeholder
    html_body = html_template.replace('{ERROR}', str(error))
    
    # Send email
    send_email_oauth(recipients, subject, html_body)

def main():
    try:
        # Set the root directory to the current working directory (using C drive instead of Z)
        root_dir = "C:\\strom_monitor\\"
        os.chdir(root_dir)

        # Set the current timestamp in UTC (update this as needed for testing)
        timestamp_utc = datetime(2024, 10, 6, 18, 0, 0, tzinfo=timezone.utc)
        
        # Set the current timestamp in UTC
        #timestamp_utc = datetime.now(timezone.utc) #COMMENT TO TEST A TIME ENTERED ABOVE
        
        # Get the current UTC time
        now_utc = datetime.now(timezone.utc)

        # Define the forecast issuance hours
        forecast_hours = [0, 6, 12, 18]

        # Adjust timestamp to the closest previous forecast hour if needed
        current_hour = timestamp_utc.hour
        valid_hours = [hour for hour in forecast_hours if hour <= current_hour]

        if valid_hours:
            latest_forecast_hour = max(valid_hours)
            timestamp_utc = timestamp_utc.replace(hour=latest_forecast_hour, minute=0, second=0, microsecond=0)
        else:
            latest_forecast_hour = 18
            timestamp_utc = (timestamp_utc - timedelta(days=1)).replace(hour=latest_forecast_hour, minute=0, second=0, microsecond=0)

        print(f"Adjusted timestamp_utc: {timestamp_utc}")

        # Read in the Areas of Interest (AOIs) GeoJSON file
        AOIs = gpd.read_file(
            os.path.join(root_dir, 'Areas_Of_Interest_For_ALERT', '2024_Policies.geojson')
        )

        # Create a buffered version of the AOIs with a 0.7-degree buffer
        AOIs_buff = gpd.GeoDataFrame(
            data=AOIs, geometry=AOIs.buffer(0.7), crs=AOIs.crs
        )

        # Define the list of email recipients from environment variables
        email_recipients = os.getenv('EMAIL_RECIPIENTS', 'edwin.petro@lockton.com')
        debug_email_recipients = os.getenv('DEBUG_EMAIL_RECIPIENTS', 'edwin.petro@lockton.com')
        
        # Convert to lists if not empty
        email_list = email_recipients.split(',') if email_recipients else []
        debug_email_list = debug_email_recipients.split(',') if debug_email_recipients else []

        # Get recent storm forecast paths from the adeck directory CSV file
        recent_adeck_paths = f.get_recent_adeck_paths(
            os.path.join(root_dir, 'storm_adeck_directory.csv'),
            timestamp_utc
        )

        # Remove 'Invest' storms from the list (storms not yet named)
        recent_adeck_paths = recent_adeck_paths[
            recent_adeck_paths['Storm_Name'] != 'Invest'
        ]

        # Check if there are any recent storms
        if recent_adeck_paths.empty:
            send_no_storms_email(email_list)
            return

        # Create a GeoDataFrame to hold the uncertainty cones for active storms
        cones = gpd.GeoDataFrame()

        #Create a list with error messages for each storm to diagnose and a dataframe with the data for the problem storms (if any)
        error_msgs=[]
        error_data=pd.DataFrame()

        # Iterate over each recent storm and create its uncertainty cone
        for index, row in recent_adeck_paths.iterrows():
            try:
                # Read the adeck data file for the storm into a GeoDataFrame
                storm_dat = gpd.read_parquet(row['adeck_path_parquet'])
                
                # Filter storm data to the specific timestamp
                print(storm_dat['DateTime'].unique())
                print('##')
                print(timestamp_utc)
                storm_dat = storm_dat[storm_dat['DateTime'] == timestamp_utc]
                
                if storm_dat.empty:
                    err_msg=f"No data found for {timestamp_utc} for storm {row['Storm_Name']}. Skipping."
                    print(err_msg)
                    error_msgs.append(err_msg)

                    storm_dat = gpd.read_parquet(row['adeck_path_parquet'])
                    storm_dat['Storm_Name']=row['Storm_Name']

                    error_data=pd.concat([error_data,storm_dat],ignore_index=True)

                    continue

                # Create the uncertainty cone for the storm
                cone = f.create_uncertainty_cone(storm_dat)

                # Add storm metadata to the cone GeoDataFrame
                cone['Basin'] = row['Basin']
                cone['Storm_Number'] = row['Storm_Number']
                cone['Storm_Name'] = row['Storm_Name']
                cone['last_update'] = row['Storm_End_Date']
                cone['adeck_path'] = row['adeck_path']
                cone['adeck_path_parquet'] = row['adeck_path_parquet']

                # Append the cone to the cones GeoDataFrame
                cones = pd.concat([cone, cones], ignore_index=True)
            except Exception as e:
                err_msg=f'Error processing storm {row["Storm_Name"]}: {e} when creating its cone'
                print(err_msg)
                
                error_msgs.append(err_msg) 

                continue

        # Check if there are any errors
        if len(error_msgs) > 0:
            try:
                send_errors_email(debug_email_list, recent_adeck_paths['Storm_Name'].unique(), error_msgs, error_data)
            except Exception as e:
                print(f"Error sending errors email: {e}")
                # Continuar con la ejecución incluso si falla el envío de correo

        # Intercept the cones with the buffered AOIs to find potential impacts
        intercepts = gpd.GeoDataFrame()

        for index, row in cones.iterrows():
            try:
                # Create a GeoDataFrame for the current cone
                row_gdf = gpd.GeoDataFrame([row], geometry='geometry', crs=cones.crs)

                # Perform a spatial overlay to find intersections with AOIs
                intercept = gpd.overlay(AOIs_buff, row_gdf)

                # Append the intercepts to the intercepts GeoDataFrame
                intercepts = pd.concat([intercept, intercepts], ignore_index=True)

                # Clean up the 'Name' and 'ClientName' fields in intercepts and AOIs
                intercepts = f.clean_strings(intercepts, 'Name', 'ClientName')
                AOIs = f.clean_strings(AOIs, 'Name', 'ClientName')

                # Get the list of affected clients from the intercepts
                affected_clients = intercepts['ClientName'].unique()

            except Exception as e:
                print(f'Error processing cone for storm {row["Storm_Name"]}: {e}')

        # If there are no affected clients, send an email and exit
        if 'affected_clients' not in locals() or len(affected_clients) == 0:
            send_no_impacts_email(email_list, recent_adeck_paths['Storm_Name'].unique())
            return

        # Generate the forecast map and send an email for each affected client
        for affc in affected_clients:
            try:
                rel_intercept = intercepts[intercepts['ClientName'] == affc]

                # For each storm affecting the client
                for storm in rel_intercept["Storm_Name"].unique():
                    try:
                        # Get the specific path for the storm data
                        adeck_path = rel_intercept.loc[rel_intercept['Storm_Name'] == storm, 'adeck_path_parquet'].iloc[0]
                        
                        # Read and filter the storm data
                        storm_dat = gpd.read_parquet(adeck_path)
                        storm_dat = storm_dat[storm_dat['DateTime'] == timestamp_utc]
                        storm_dat = f.filter_adeck_gdf_for_official_model(storm_dat)

                        # Get the cone for the current storm
                        cone = cones[cones['Storm_Name'] == storm]

                        # Define the title for the map
                        title = f"Lockton Alert for {storm}"

                        print(storm_dat.columns)

                        # Create the forecast map figure
                        fig = f.create_forecast_map_with_cone_for_AOIs(
                            storm_dat,
                            cone,
                            AOIs[AOIs['ClientName'] == affc],
                            AOI_Name='Name',
                            title=title
                        )

                        # Define the filename for the HTML map
                        filename = f"client_storm_alerts/{affc}-{storm}-forecast-{timestamp_utc.strftime('%Y-%m-%d %H-%M')}.html"
                        filename = filename.replace(':', '-')

                        # Ensure the directory exists
                        os.makedirs(os.path.dirname(filename), exist_ok=True)

                        # Save the figure as an HTML file
                        fig.write_html(filename)

                        # Prepare email parameters and send the map
                        send_forecast_map_email(
                            fig, storm, affc, timestamp_utc, email_list
                        )

                    except Exception as e:
                        print(f'Error processing storm {storm} for client {affc}: {e}')
            except Exception as e:
                print(f'Error processing client {affc}: {e}')

    except Exception as e:
        print(f'An error occurred: {e}')
        try:
            send_error_email(e, debug_email_list)
        except Exception as email_error:
            print(f"Error al enviar correo de error: {email_error}")
            # Si falla el envío de correo, al menos registramos el error original en los logs

if __name__ == "__main__":
    main()