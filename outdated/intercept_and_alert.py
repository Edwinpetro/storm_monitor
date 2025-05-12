import geopandas as gpd
import pandas as pd
from shapely.geometry import Point
from shapely.ops import unary_union
from datetime import datetime, timedelta, timezone
import os
import win32com.client as win32
import plotly.graph_objects as go
import tempfile
import plotly.io as pio
import plotly.graph_objects as go

import app_functions as f


root_dir=os.getcwd()
timestamp_utc = pd.to_datetime("2024-10-07 12:00").tz_localize('UTC') #Enter a value for testing
now_utc = datetime.now(timezone.utc)
timestamp_utc=now_utc #Comment this line in order to test for an as-if date

AOIs=gpd.read_file(root_dir+'\\Areas_Of_Interest_For_ALERT\\2024_Policies.geojson')
AOIs_buff=gpd.GeoDataFrame(data=AOIs,geometry=AOIs.buffer(0.7),crs=AOIs.crs)

email_list=['alvaro.farias@lockton.com']#,'DMonsalve@lockton.com']

print(root_dir+'\\storm_adeck_directory.csv')
recent_adeck_paths = f.get_recent_adeck_paths(root_dir+'\\storm_adeck_directory.csv',)
print(recent_adeck_paths)
recent_adeck_paths=recent_adeck_paths[recent_adeck_paths['Storm_Name']!='Invest'] #Remove Invests from list



#Create a cone dataframe for the active storms
cones=gpd.GeoDataFrame()
for index, row in recent_adeck_paths.iterrows():
    storm_dat=f.read_adeck_dat_file_to_gdf(row['adeck_path'],forecast_datetime=timestamp_utc)
    if len(storm_dat[storm_dat['ForecastHour']>0])>0:
        cone=f.create_uncertainty_cone(storm_dat)
        cone['Basin']=row['Basin']
        cone['Storm_Number']=row['Storm_Number']
        cone['Storm_Name']=row['Storm_Name']
        cone['last_update']=row['Storm_End_Date']
        cone['adeck_path']=row['adeck_path']
        cones=pd.concat([cone,cones])
    else: #There are no forecasts for this storm
        print('There are no forecasts for storm ' + row['Storm_Name'])

#Intercept the cones with the buffered AOIs
intercepts=gpd.GeoDataFrame()
for index, row in cones.iterrows():
    row_gdf = gpd.GeoDataFrame([row], geometry='geometry', crs=cones.crs)
    intercept=gpd.overlay(AOIs_buff,row_gdf)
    intercepts=pd.concat([intercept,intercepts])

intercepts=f.clean_strings(intercepts,'Name','ClientName')
AOIs=f.clean_strings(AOIs,'Name','ClientName')

# For intercepts between the storms cone and the clients buffered trigger geometry (AOI), make the map and email it
affected_clients=intercepts['ClientName'].unique()
if len(affected_clients)==0:
    print("No clients affected by currently active storms")
else:
    for affc in affected_clients:
        rel_intercept=intercepts[intercepts['ClientName']==affc]
        
        for storm in rel_intercept["Storm_Name"].unique():
            rel_intercept=rel_intercept[rel_intercept['Storm_Name']==storm]
            adeck_gdf=f.read_adeck_dat_file_to_gdf(rel_intercept['adeck_path'].unique()[0],forecast_datetime=timestamp_utc)#)
            adeck_gdf=f.filter_adeck_gdf_for_official_model(adeck_gdf)
            cone=cones[cones['Storm_Name']==storm]

            logo_path='Z:\\Event_Monitor\\logos\\LOCKTON_logo-white-footer.svg'
            Title=f"Lockton Alert for {storm}"

            fig=f.create_forecast_map_with_cone_for_AOIs(adeck_gdf,cone,AOIs[AOIs['ClientName']==affc],AOI_Name='Name',title=Title,logo_url=logo_path)

            filename=f"client_storm_alerts/{affc}-{storm}-forecast-{timestamp_utc.strftime('%Y-%m-%d %H:%M')}.html"
            filename= filename.replace(':', '-')

            fig.write_html(f'./{filename}')

            #Send the email

            # Define email parameters
            recipients = email_list

            subject = f'Forecast Map - {storm}'
            body = f"""
            <html>
            <head></head>
            <body>
                <p>Dear Team,</p>
                <p>Please find the attached forecast map for <b>{storm}</b> based on the Official Forecast made on <b>{timestamp_utc} UTC</b> </p>
                <p>Best regards,<br>Lockton Storm Monitor</p>
            </body>
            </html>
            """

            # Optional parameters
            cc = ['']
            bcc = ['']
            sender = 'alvaro.farias@lockton.com'  # Optional: specify if needed

            # Send the email
            f.send_map_via_email(fig, recipients, subject, body, cc=cc, bcc=bcc, sender=sender)



