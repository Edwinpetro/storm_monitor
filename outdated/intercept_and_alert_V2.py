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
        # Add this at the beginning of your main function to ensure the correct working directory
        os.chdir("Z:\\Event_Monitor\\")

        # Set the current timestamp in UTC
        timestamp_utc = datetime.now(timezone.utc)

        # Get the current UTC time
        now_utc = datetime.now(timezone.utc)

        # for testing as-if it what a different time. Comment these 2 lines when done testing
        timestamp_utc = datetime(2024, 10, 7, 12, 0, 0, tzinfo=timezone.utc)
        now_utc = timestamp_utc

        # Define the forecast issuance hours
        forecast_hours = [0, 6, 12, 18]

        # Get the current hour
        current_hour = now_utc.hour

        # Find the latest forecast hour less than or equal to the current hour
        valid_hours = [hour for hour in forecast_hours if hour <= current_hour]

        if valid_hours:
            # If there is a valid forecast hour, use the latest one
            latest_forecast_hour = max(valid_hours)
            # Set timestamp_utc to today at the latest forecast hour
            timestamp_utc = now_utc.replace(hour=latest_forecast_hour, minute=0, second=0, microsecond=0)
        else:
            # If current hour is before the first forecast hour (shouldn't happen but included for completeness)
            # Go back to the previous day's last forecast hour
            latest_forecast_hour = 18
            timestamp_utc = (now_utc - timedelta(days=1)).replace(hour=latest_forecast_hour, minute=0, second=0, microsecond=0)

        print(f"Current UTC time: {now_utc}")
        print(f"Adjusted timestamp_utc: {timestamp_utc}")

        # Read in the Areas of Interest (AOIs) GeoJSON file
        AOIs = gpd.read_file(
            os.path.join(root_dir, 'Areas_Of_Interest_For_ALERT', '2024_Policies.geojson')
        )

        # Create a buffered version of the AOIs with a 0.7-degree buffer
        # This buffer expands the AOIs slightly for interception calculations
        AOIs_buff = gpd.GeoDataFrame(
            data=AOIs, geometry=AOIs.buffer(0.7), crs=AOIs.crs
        )

        # Define the list of email recipients
        email_list = ['alvaro.farias@lockton.com']  # Add additional recipients as needed

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
            # If there are no recent storms, send email and exit
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
            recipients = ['alvaro.farias@lockton.com']
            cc = []
            bcc = []
            sender = 'alvaro.farias@lockton.com'

            # Send the email
            f.send_email(recipients, subject, body, cc=cc, bcc=bcc, sender=sender)

            # Exit the script
            return

        # Create a GeoDataFrame to hold the uncertainty cones for active storms
        cones = gpd.GeoDataFrame()

        # Iterate over each recent storm and create its uncertainty cone
        for index, row in recent_adeck_paths.iterrows():
            try:
                # Initialize timestamp for retry logic
                attempts = 1
                                
                while attempts > 0:
                    # Read the adeck data file for the storm into a GeoDataFrame
                    #storm_dat = f.read_adeck_dat_file_to_gdf(
                    #    row['adeck_path_parquet'], forecast_datetime=timestamp_utc
                    #)
                    storm_dat=gpd.read_parquet(row['adeck_path_parquet'])
                    storm_dat=f.format_adeck_parquet(storm_dat)
                    
                    storm_dat = storm_dat[storm_dat['DateTime'] == timestamp_utc]
                    
                    # Check if storm_dat is empty
                    if storm_dat is not None and not storm_dat.empty:
                        break  # Stop if valid data is found
                    else:
                        print(f"No data found for {timestamp_utc}. Retrying with 6-hour earlier timestamp.")
                        # Decrease the timestamp by 6 hours
                        timestamp_utc -= pd.Timedelta(hours=6)
                        attempts -= 1
                if storm_dat is None or storm_dat.empty:
                    print(f"Skipping storm {row['storm_id']} after 5 failed attempts.")
                    continue  # Skip this storm and move to the next iteration in the loop

                # Check if there are forecast hours greater than 0
                if len(storm_dat[storm_dat['ForecastHour'] > 0]) > 0:
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
                else:
                    # There are no forecasts for this storm
                    print(f'There are no forecasts for storm {row["Storm_Name"]}')
            except Exception as e:
                print(f'Error processing storm {row["Storm_Name"]}: {e}')
                continue

        # Check if there are any cones generated
        if cones.empty:
            # No cones were generated, which might indicate an issue
            subject = 'No uncertainty cones generated for recent storms'
            body = """
            <html>
            <head></head>
            <body>
                <p>Alvaro,</p>
                <p>No uncertainty cones were generated for recent storms. Please check the data.</p>
                <p>Best regards,<br>Lockton Storm Monitor</p>
            </body>
            </html>
            """
            recipients = ['alvaro.farias@lockton.com']
            cc = []
            bcc = []
            sender = 'alvaro.farias@lockton.com'

            # Send the email
            f.send_email(recipients, subject, body, cc=cc, bcc=bcc, sender=sender)

            # Exit the script
            return

        # Intercept the cones with the buffered AOIs to find potential impacts
        intercepts = gpd.GeoDataFrame()

        for index, row in cones.iterrows():
            try:
                # Create a GeoDataFrame for the current cone
                row_gdf = gpd.GeoDataFrame(
                    [row], geometry='geometry', crs=cones.crs
                )

                # Perform a spatial overlay to find intersections with AOIs
                intercept = gpd.overlay(AOIs_buff, row_gdf)

                # Append the intercepts to the intercepts GeoDataFrame
                intercepts = pd.concat([intercept, intercepts], ignore_index=True)
                
            except Exception as e:
                print(f'Error processing cone for storm {row["Storm_Name"]}: {e}')

        # Clean up the 'Name' and 'ClientName' fields in intercepts and AOIs
        intercepts = f.clean_strings(intercepts, 'Name', 'ClientName')
        AOIs = f.clean_strings(AOIs, 'Name', 'ClientName')

        # Get the list of affected clients from the intercepts
        affected_clients = intercepts['ClientName'].unique()

        # If there are no affected clients, send an email and exit
        if len(affected_clients) == 0:
            print("No clients affected by currently active storms")

            # Prepare the list of recent storms
            recent_storms_list = recent_adeck_paths['Storm_Name'].unique()
            storms_str = ', '.join(recent_storms_list)

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
            recipients = ['alvaro.farias@lockton.com']
            cc = []
            bcc = []
            sender = 'alvaro.farias@lockton.com'

            # Send the email
            f.send_email(recipients, subject, body, cc=cc, bcc=bcc, sender=sender)

            # Exit the script
            return

        # For each affected client, generate the forecast map and send an email
        for affc in affected_clients:
            try:
                # Filter intercepts for the current client
                rel_intercept = intercepts[intercepts['ClientName'] == affc]

                # For each storm affecting the client
                for storm in rel_intercept["Storm_Name"].unique():
                    try:
                        # Filter intercepts for the current storm
                        rel_storm_intercept = rel_intercept[rel_intercept['Storm_Name'] == storm]
                        # Read the adeck data for the storm
                        storm_dat=gpd.read_parquet(rel_storm_intercept['adeck_path_parquet'].iloc[0])
                        storm_dat=f.format_adeck_parquet(storm_dat)
                        print("#### before time filter")
                        print(timestamp_utc)
                        print(storm_dat)
                        storm_dat = storm_dat[storm_dat['DateTime'] == timestamp_utc]
                        print("ACA")
                        print(storm_dat)
                        # Filter for the official model
                        storm_dat = f.filter_adeck_gdf_for_official_model(storm_dat)

                        # Get the cone for the current storm
                        cone = cones[cones['Storm_Name'] == storm]

                        # Define the path to the logo image
                        logo_path = 'Z:\\Event_Monitor\\logos\\LOCKTON_logo-white-footer.svg'

                        # Define the title for the map
                        Title = f"Lockton Alert for {storm}"
                        

                        # Create the forecast map figure
                        fig = f.create_forecast_map_with_cone_for_AOIs(
                            storm_dat,
                            cone,
                            AOIs[AOIs['ClientName'] == affc],
                            AOI_Name='Name',
                            title=Title
                            #logo_url=logo_path
                        )

                        # Define the filename for the HTML map
                        filename = f"client_storm_alerts/{affc}-{storm}-forecast-{timestamp_utc.strftime('%Y-%m-%d %H-%M')}.html"
                        filename = filename.replace(':', '-')

                        # Ensure the directory exists
                        os.makedirs(os.path.dirname(filename), exist_ok=True)

                        # Save the figure as an HTML file
                        fig.write_html(filename)

                        # Prepare email parameters
                        recipients = email_list
                        subject = f'Forecast Map - {storm}'
                        body = f"""
                        <html>
                        <head></head>
                        <body>
                            <p>Team,</p>
                            <p>Please find the attached forecast map for <b>{storm}</b> based on the Official Forecast made on <b>{timestamp_utc.strftime('%Y-%m-%d %H:%M')} UTC</b>.</p>
                            <p>Best regards,<br>Lockton Storm Monitor</p>
                        </body>
                        </html>
                        """

                        # Optional parameters
                        cc = ['']
                        bcc = ['']
                        sender = 'alvaro.farias@lockton.com'  # Optional: specify if needed

                        # Send the email with the map
                        f.send_map_via_email(
                            fig, recipients, subject, body, cc=cc, bcc=bcc, sender=sender
                        )

                    except Exception as e:
                        print(f'Error processing storm {storm} for client {affc}: {e}')
            except Exception as e:
                print(f'Error processing client {affc}: {e}')

    except Exception as e:
        print(f'An error occurred: {e}')
        # Optionally, send an email notification about the error
        subject = 'Error in Storm Monitor Script'
        body = f"""
        <html>
        <head></head>
        <body>
            <p>Alvaro,</p>
            <p>An error occurred in the Storm Monitor script:</p>
            <p>{e}</p>
            <p>Please check the logs for more details.</p>
            <p>Best regards,<br>Lockton Storm Monitor</p>
        </body>
        </html>
        """
        recipients = ['alvaro.farias@lockton.com']
        cc = []
        bcc = []
        sender = 'alvaro.farias@lockton.com'

        # Send the error email
        f.send_email(recipients, subject, body, cc=cc, bcc=bcc, sender=sender)


if __name__ == "__main__":
    main()