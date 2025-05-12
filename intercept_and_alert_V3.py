import geopandas as gpd
import pandas as pd
from shapely.geometry import Point
from shapely.ops import unary_union
from datetime import datetime, timedelta, timezone
import os
import plotly.graph_objects as go
import plotly.io as pio

import app_functions as f  # Import custom functions from app_functions module

def main():
    try:
        # Set the root directory to the current working directory
        root_dir = "Z:\\Event_Monitor\\"
        os.chdir(root_dir)

        # Set the current timestamp in UTC (update this as needed for testing)
        timestamp_utc = datetime(2024, 10, 6, 18, 0, 0, tzinfo=timezone.utc)
        
        # Set the current timestamp in UTC
        #timestamp_utc = datetime.now(timezone.utc) #COMMENT TO TEST A TIME ENTERED ABOVE
#
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

        # Define the list of email recipients
        email_list = ['alvaro.farias@lockton.com']#,'matt.cohen@lockton.com','BWierenga@lockton.com','Maria.Rodriguez@lockton.com','Alexis.Pedraza@lockton.com','Henry.Bellwood@lockton.com','jm.ramos@lockton.com','BMcCann@lockton.com']
        debug_email_list = ['alvaro.farias@lockton.com']

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
                #storm_dat = f.format_adeck_parquet(storm_dat)
                
                #storm_dat.to_csv(f'./{row['Storm_Name']}_testing_stormdat.csv')
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
                    #storm_dat = f.format_adeck_parquet(storm_dat)
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
        if len(error_msgs)>0:
            send_errors_email(debug_email_list, recent_adeck_paths['Storm_Name'].unique(), error_msgs,error_data)
            

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
        if len(affected_clients) == 0:
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
                        #storm_dat = f.format_adeck_parquet(storm_dat)
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
        send_error_email(e, ['alvaro.farias@lockton.com'])

def send_no_storms_email(recipients):
    subject = 'There are currently no recently active storms'
    body = """
    <html>
    <head></head>
    <body>
        <p>Team,</p>
        <p>There are currently no recently active storms.</p>
        <p>Best regards,<br>Lockton Storm Monitor</p>
    </body>
    </html>
    """
    f.send_email(recipients, subject, body)

def send_errors_email(recipients,storms,error_msgs,error_data):
    storms_str = ', '.join(storms)
    error_msgs_str=', '.join(error_msgs)

    subject = 'Some errors occurred while processing storms'
    body = f"""
    <html>
    <head></head>
    <body>
        <p>Alvaro,</p>
        <p>Some errors where encountered while processning ({storms_str}) Please check the data attached.</p>
        <p>Here are some error messages from the processing: {error_msgs_str}
        <p>Best regards,<br>Lockton Storm Monitor</p>
    </body>
    </html>
    """
    f.send_email(recipients, subject, body, error_data)

def send_no_impacts_email(recipients, storms):
    storms_str = ', '.join(storms)
    subject = 'No Recently Active Storms Intercepting the Client Portfolio'
    body = f"""
    <html>
    <head></head>
    <body>
        <p>Team,</p>
        <p>There are recently active storms: {storms_str}, but none whose cone of uncertainty come within ~70km of intercepting any geometry in the client portfolio.</p>
        <p>Best regards,<br>Lockton Storm Monitor</p>
    </body>
    </html>
    """
    f.send_email(recipients, subject, body)

def send_forecast_map_email(fig, storm, client_name, timestamp_utc, recipients):
    subject = f'Forecast Map - {storm}'
    body = f"""
    <html>
    <head></head>
    <body>
        <p>Team,</p>
        <p>Please find the attached forecast map for <b>{client_name} </b> for <b>{storm}</b> based on the Official Forecast made on <b>{timestamp_utc.strftime('%Y-%m-%d %H:%M')} UTC</b>.</p>
        <p>Best regards,<br>Lockton Storm Monitor</p>
    </body>
    </html>
    """
    f.send_map_via_email(
        fig, recipients, subject, body, cc=[], bcc=[], sender='alvaro.farias@lockton.com'
    )

def send_error_email(error, recipients):
    subject = 'Error in Storm Monitor Script'
    body = f"""
    <html>
    <head></head>
    <body>
        <p>Alvaro,</p>
        <p>An error occurred in the Storm Monitor script:</p>
        <p>{error}</p>
        <p>Please check the logs for more details.</p>
        <p>Best regards,<br>Lockton Storm Monitor</p>
    </body>
    </html>
    """
    f.send_email(recipients, subject, body)

if __name__ == "__main__":
    main()