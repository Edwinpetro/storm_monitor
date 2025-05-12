import pandas as pd
import plotly.graph_objects as go
import datetime
import os
import geopandas as gpd
from shapely.geometry import Point
from shapely.ops import unary_union
import base64 
import numpy as np
import plotly.io as pio
from datetime import datetime, timedelta, timezone
import tempfile
import re
import win32com.client as win32


# Dictionary to map forecast period (hours) to 2/3 probability circle radii (in nautical miles)
radius_lookup = {
    'AL': {12: 26, 24: 41, 36: 55, 48: 70, 60: 88, 72: 102, 96: 151, 120: 220},
    'EP': {12: 26, 24: 39, 36: 53, 48: 65, 60: 76, 72: 92, 96: 119, 120: 152},
    'CP': {12: 34, 24: 49, 36: 66, 48: 81, 60: 95, 72: 120, 96: 137, 120: 156},
    'WP': {12: 26, 24: 41, 36: 55, 48: 70, 60: 88, 72: 102, 96: 151, 120: 220}, #No data provided by NHC. Using same as AL
    'IO': {12: 26, 24: 41, 36: 55, 48: 70, 60: 88, 72: 102, 96: 151, 120: 220}, #No data provided by NHC. Using same as AL
    'SH': {12: 26, 24: 41, 36: 55, 48: 70, 60: 88, 72: 102, 96: 151, 120: 220}, #No data provided by NHC. Using same as AL
    'SL': {12: 26, 24: 41, 36: 55, 48: 70, 60: 88, 72: 102, 96: 151, 120: 220}  #No data provided by NHC. Using same as AL
}

def get_radius(basin, forecast_hour):
    """
    Returns the 2/3 probability circle radius (nautical miles) for the given basin and forecast hour.
    """
    if basin in radius_lookup and forecast_hour in radius_lookup[basin]:
        return radius_lookup[basin][forecast_hour]
    return None

def filter_adeck_gdf_for_official_model(adeck_gdf):
    """
    Given an a deck geoDataFrame returns the official forecast. It varies by Basin
    """

    models_to_try = ['OFCL', 'OFCI', 'AEMN', 'AVNO', 'AVNI', 'AVNX', 'EGRR', 'EGRI', 'EGR2', 
                     'UKM', 'UKX', 'UKMI', 'UKXI', 'UKM2', 'UKX2', 'HWRF', 'HWFI', 'HWF2']  # Ordered by preference
    for model in models_to_try:
        ofcl_gdf = adeck_gdf[adeck_gdf['ModelName'] == model]
        if not ofcl_gdf.empty:
            #print(f"found data for {model}")
            break  # Stop if we find data for the current model
        

    if ofcl_gdf.empty:
        print("No relevant model data found. Returning empty cone.")
        ofcl_gdf = gpd.GeoDataFrame()  # Return empty GeoDataFrame

    return ofcl_gdf

def create_uncertainty_cone(input_gdf):
    """
    Given an input GeoDataFrame with forecast points, generates the uncertainty cone for the OFCL model
    based on forecast hour and basin, ensuring the buffers are created in date order.

    Parameters:
    - input_gdf: GeoDataFrame with forecast points, containing lat/lons, forecast hour, and basin.

    Returns:
    - cone_gdf: GeoDataFrame containing the merged polygon representing the uncertainty cone.
    """

    ofcl_gdf=filter_adeck_gdf_for_official_model(input_gdf)

    # Sort by DateTime to ensure the buffers are created in the correct order
    ofcl_gdf = ofcl_gdf.sort_values(by='ForecastHour')

    # Convert input GeoDataFrame to UTM for accurate buffering in meters
    ofcl_gdf = ofcl_gdf.to_crs(ofcl_gdf.estimate_utm_crs())

    # List to store buffered geometries
    buffered_geometries = []

    # Iterate over each point and buffer according to its forecast hour and basin
    for _, row in ofcl_gdf.iterrows():
        forecast_hour = row['ForecastHour']
        basin = row['Basin']

        # Get the radius in nautical miles for the forecast hour and basin
        radius_nm = get_radius(basin, forecast_hour)
        if radius_nm is None:
            #print(f"No radius found for Basin: {basin}, Forecast Hour: {forecast_hour}")
            continue

        # Convert radius from nautical miles to meters (1 NM = 1852 meters)
        radius_m = radius_nm * 1852

        # Apply buffer (in meters) and store it
        buffer_geom = row['geometry'].buffer(radius_m)
        buffered_geometries.append(buffer_geom)

    if not buffered_geometries:
        print("No buffered geometries created.")
        return gpd.GeoDataFrame()

    # Generate the convex hull between each pair of consecutive buffers
    convex_hull_segments = []
    for first, second in zip(buffered_geometries, buffered_geometries[1:]):
        # Apply convex hull directly between the current and next buffer
        convex_hull_segment = first.convex_hull.union(second.convex_hull).convex_hull
        convex_hull_segments.append(convex_hull_segment)

    # Perform unary union on all convex hull segments to merge them into one polygon
    final_cone_polygon = unary_union(convex_hull_segments)

    # Create a GeoDataFrame with the final merged polygon
    cone_gdf = gpd.GeoDataFrame(geometry=gpd.GeoSeries([final_cone_polygon]), crs=ofcl_gdf.crs)

    # Reproject back to EPSG:4326 for lat/lon coordinates
    cone_gdf = cone_gdf.to_crs("EPSG:4326")

    return cone_gdf

def create_forecast_map_with_cone(adeck_gdf, selected_forecast_datetime):
    """
    Reads the gdf of the forecasts file and creates a Plotly map of the forecast tracks, including the uncertainty cone.

    Parameters:
    - adeck_gdf: GeoDatFrame with a point geometry column of the target dat file (a deck)
    - selected_forecast_datetime: str, the forecast date and time in YYYYMMDDHH format.

    Returns:
    - fig: Plotly figure object.
    """
    mapbox_access_token = 'pk.eyJ1IjoiYWx2YXJvZmFyaWFzIiwiYSI6ImNtMXptbm9iaDA4OHMybG9vc3VqdW1vZ3oifQ.ZJ8d6gNAiR1htIYxESOYuQ'
    
    # Filter by forecast datetime if provided
    if selected_forecast_datetime:
        adeck_gdf = adeck_gdf[adeck_gdf['DateTime'] == selected_forecast_datetime]
    
    # Filter out data points where ValidTime is more than 5 days (120 hours) from the minimum ValidTime
    min_valid_time = adeck_gdf['ValidTime'].min()
    max_time_limit = min_valid_time + timedelta(hours=120)
    adeck_gdf = adeck_gdf[adeck_gdf['ValidTime'] <= max_time_limit]

    # List of models to keep (prioritized models)
    models_to_keep = ['OFCL', 'OFCI', 'AVNO', 'AVNI', 'AVNX', 'EGRR', 'EGRI', 'EGR2', 
                      'UKM', 'UKX', 'UKMI', 'UKXI', 'UKM2', 'UKX2', 'HWRF', 'HWFI', 'HWF2']
    
    adeck_gdf = adeck_gdf[adeck_gdf['ModelName'].isin(models_to_keep)]

    # Define category colors in the desired order
    category_colors = {
        'Tropical Depression': '#A0A0A0',  # Light gray
        'Tropical Storm': '#4682B4',  # Steel blue
        'Category 1': '#FFD700',  # Gold/yellow
        'Category 2': '#FFA500',  # Orange
        'Category 3': '#FF4500',  # Orange-red
        'Category 4': '#FF0000',  # Red
        'Category 5': '#B22222',  # Dark red/magenta
        'Unknown': '#808080'  # Medium gray
    }

    # Sort the DataFrame by ModelName and ValidTime to ensure lines connect correctly per model
    adeck_gdf = adeck_gdf.sort_values(['ModelName', 'ValidTime'])

    # Initialize the Plotly figure
    fig = go.Figure()

    # Add a line trace for each model (slightly lighter blue for the storm paths)
    line_color = '#ADD8E6'  # Light blue
    models = adeck_gdf['ModelName'].unique()
    for model in models:
        model_gdf = adeck_gdf[adeck_gdf['ModelName'] == model].sort_values('ValidTime')
        if model_gdf.empty:
            continue
        
        # Default trace for the storm path (light blue line)
        fig.add_trace(go.Scattermapbox(
            lat=model_gdf['Latitude'],
            lon=model_gdf['Longitude'],
            mode='lines',
            line=dict(color=line_color, width=2.5),  
            hoverinfo='none',  # No hover info for this line
            showlegend=False  # Do not show in the legend
        ))
    
        # Group data by Category in the specified order and plot the markers 
    for category, color in category_colors.items():
        category_gdf = adeck_gdf[adeck_gdf['Category'] == category] 

        if not category_gdf.empty:
            fig.add_trace(go.Scattermapbox(
                lat=category_gdf['Latitude'],
                lon=category_gdf['Longitude'],
                mode='markers',
                name=category,
                marker=dict(
                    size=12,  
                    color=color,  # Single color per category
                ),
                text=category_gdf.apply(lambda row: f"Time: {row['ValidTime']:%Y-%m-%d %H:%M UTC}<br>"
                                                    f"Wind Speed (mph): {row['MaxWindSpeed_mph']:.1f}<br>"
                                                    #f"Central Pressure (mb): {row['MinPressure']}<br>"
                                                    f"Category: {row['Category']}<br>"
                                                    f"Model: {row['ModelName']}", axis=1),
                hoverinfo='text'
            ))

    # Generate and simplify the uncertainty cone geometry to reduce render time
    cone_gdf = create_uncertainty_cone(adeck_gdf)
    if not cone_gdf.empty:
        cone_gdf['geometry'] = cone_gdf['geometry'].simplify(tolerance=0.01)  # Simplify geometry

    # If the cone exists, add it to the map
    if not cone_gdf.empty:
        # Extract lat/lon coordinates for plotting
        cone_coords = cone_gdf.geometry[0].exterior.coords.xy
        cone_lats, cone_lons = list(cone_coords[1]), list(cone_coords[0])  # Convert to lists

        # Add uncertainty cone polygon to the map (darker grey for the cone and border)
        fig.add_trace(go.Scattermapbox(
            lat=cone_lats,
            lon=cone_lons,
            mode='lines',
            fill='toself',
            fillcolor='rgba(105, 105, 105, 0.7)',  # Darker grey for the cone
            line=dict(color='darkgrey', width=3),  # Darker grey border
            name='Uncertainty Cone',
            hoverinfo='skip',  # No hover info for the cone
            visible='legendonly'  # Let the user toggle this layer in the legend
        ))

    # Add a logo (SVG) to the map layout
    fig.update_layout(
        images=[
            dict(
                source='./assets/LOCKTON_logo-white-footer.svg',  # Path to the SVG image
                xref="paper", yref="paper",  # Relative to the plot area
                x=0.01, y=0.99,  # Position of the logo
                sizex=0.1, sizey=0.12,  # Size of the logo
                xanchor="left", yanchor="top"  # Anchor the image position
            )
        ],
                # Set up the map layout with Mapbox Satellite Streets
        mapbox=dict(
            accesstoken=mapbox_access_token,
            style='mapbox://styles/mapbox/satellite-streets-v12',  # Use full style URL
            zoom=5,
            center={"lat": adeck_gdf['Latitude'].mean(), "lon": adeck_gdf['Longitude'].mean()},
        ),
        margin={"r": 0, "t": 0, "l": 0, "b": 0},
        
        legend=dict( #legend settings
            title=dict(
                text="LEGEND",  # Legend title
                font=dict(
                    color="white",  # Title font color
                    size=16,  # Title font size
                    family="Arial",  # Set font family
                    weight="bold"  # Make only the title bold
                ),
                side="top"  # Position the title at the top and center
            ),
            font=dict(
                color="white",  # Set legend text color to white
                size=14  # Increase font size to 14 for legend items (normal weight)
            ),
            bgcolor="black",  # Set legend background to black
            yanchor="top",
            y=0.85,  # Move down
            x=0.01,  # Align to the left
            xanchor="left",
            itemsizing='constant'
        ),
            uirevision=True,  # Preserve the zoom level and state of the map
        )

    return fig

def create_forecast_map(adeck_gdf,selected_forecast_datetime):
    """
    Reads the gdf of the forecasts file and creates a Plotly map of the forecast tracks.

    Parameters:
    - adeck_gdf: GeoDatFrame with a point geometry column of the target dat file (a deck)
    - selected_forecast_datetime: str, the forecast date and time in YYYYMMDDHH format.

    Returns:
    - fig: Plotly figure object.
    """
    mapbox_access_token = 'pk.eyJ1IjoiYWx2YXJvZmFyaWFzIiwiYSI6ImNtMXptbm9iaDA4OHMybG9vc3VqdW1vZ3oifQ.ZJ8d6gNAiR1htIYxESOYuQ'


        # Filter by forecast datetime if provided
    if selected_forecast_datetime:
        adeck_gdf = adeck_gdf[adeck_gdf['DateTime'] == selected_forecast_datetime]

    # Define category colors in the desired order
    category_colors = {
        'Tropical Depression': 'green',
        'Tropical Storm': 'blue',
        'Category 1': 'yellow',
        'Category 2': 'orange',
        'Category 3': 'red',
        'Category 4': 'orange',
        'Category 5': 'magenta',
        'Unknown': 'gray'
    }

    # Define the desired order for the legend
    category_order = [
        'Tropical Depression',
        'Tropical Storm',
        'Category 1',
        'Category 2',
        'Category 3',
        'Category 4',
        'Category 5',
        'Unknown'
    ]

    # Map category to colors based on the updated category_colors
    adeck_gdf['Color'] = adeck_gdf['Category'].apply(lambda x: category_colors.get(x, 'gray'))

    # Sort the DataFrame by ModelName and ValidTime to ensure lines connect correctly per model
    adeck_gdf = adeck_gdf.sort_values(['ModelName', 'ValidTime'])

    # Initialize the Plotly figure
    fig = go.Figure()

    # Add a line trace for each model (white lines connecting points of the same model)
    models = adeck_gdf['ModelName'].unique()
    for model in models:
        model_gdf = adeck_gdf[adeck_gdf['ModelName'] == model].sort_values('ValidTime')
        if model_gdf.empty:
            continue
        fig.add_trace(go.Scattermapbox(
            lat=model_gdf['Latitude'],
            lon=model_gdf['Longitude'],
            mode='lines',
            line=dict(color='white', width=2),
            hoverinfo='none',  # No hover info for the lines
            showlegend=False    # Do not show lines in the legend
        ))

    # Group data by Category in the specified order and plot the markers
    for category in category_order:
        category_gdf = adeck_gdf[adeck_gdf['Category'] == category]
        if not category_gdf.empty:
            fig.add_trace(go.Scattermapbox(
                lat=category_gdf['Latitude'],
                lon=category_gdf['Longitude'],
                mode='markers',
                name=category,
                marker=dict(
                    size=6,
                    color=category_colors[category],  # Single color per category
                ),
                text=category_gdf.apply(lambda row: f"Time: {row['ValidTime']:%Y-%m-%d %H:%M UTC}<br>"
                                               f"Wind Speed (mph): {row['MaxWindSpeed_mph']:.1f}<br>"
                                               f"Central Pressure (mb): {row['MinPressure']}<br>"
                                               f"Category: {row['Category']}<br>"
                                               f"Model: {row['ModelName']}", axis=1),
                hoverinfo='text'
            ))

    # Add a logo (SVG) to the map layout
    fig.update_layout(
        images=[
            dict(
                source='./assets/LOCKTON_logo-white-footer.svg',  # Path to the SVG image
                xref="paper", yref="paper",  # Relative to the plot area
                x=0.01, y=0.99,  # Position of the logo
                sizex=0.12, sizey=0.12,  # Size of the logo
                xanchor="left", yanchor="top"  # Anchor the image position
            )
        ],
        # Set up the map layout with a dark style
        
        mapbox=dict(
            accesstoken=mapbox_access_token,
            style='mapbox://styles/mapbox/satellite-streets-v12',  # Use full style URL
            zoom=5,
            center={"lat": adeck_gdf['Latitude'].mean(), "lon": adeck_gdf['Longitude'].mean()},
        ),
        margin={"r": 0, "t": 0, "l": 0, "b": 0},

        legend_title_text='Storm Category',
        legend=dict(
            itemsizing='constant'
        )
    )
    return fig

# Helper function to parse latitude and longitude
def parse_lat_lon(value):
    """
    Parses latitude or longitude value from ATCF format to decimal degrees.
    """
    if isinstance(value, str):
        value = value.strip()
        if value and value[-1] in ['N', 'S', 'E', 'W']:
            direction = value[-1]
            try:
                degrees = float(value[:-1]) / 10.0  # Assuming value is in tenths of degrees
            except ValueError:
                print(f"Error parsing degrees from value: {value}")
                return None
            if direction in ['S', 'W']:
                degrees = -degrees
            return degrees
        elif value:
            # If no direction indicator, try to convert directly
            try:
                return float(value) / 10.0
            except ValueError:
                print(f"Error parsing lat/lon value '{value}': cannot convert to float.")
                return None
    return None

# Wind speed thresholds in knots
ws_categories = [
    (0, 34, 'Tropical Depression', -1),
    (34, 63, 'Tropical Storm', 0),
    (64, 82, 'Category 1', 1),
    (83, 95, 'Category 2', 2),
    (96, 112, 'Category 3', 3),
    (113, 136, 'Category 4', 4),
    (137, float('inf'), 'Category 5', 5)
]

# Helper function to categorize wind speed
def wind_speed_to_category(wind_speed):
    """
    Converts wind speed in knots to hurricane category.
    """
    for lower_bound, upper_bound, category_name, category_number in ws_categories:
        if lower_bound <= wind_speed < upper_bound:
            return category_name, category_number
    return 'Unknown' ,-999
    
# Central pressure thresholds in mb
cp_categories = [
    (np.inf, 990, 'Tropical Storm', 0),
    (990, 980, 'Category 1', 1),
    (980, 965, 'Category 2', 2),
    (965, 945, 'Category 3', 3),
    (945, 920, 'Category 4', 4),
    (920, 0, 'Category 5', 5)
]

# Helper function to categorize central pressure
def pressure_to_category(central_pressure):
    """
    Converts central pressure in mb to hurricane category.
    """
    for upper_bound, lower_bound, category_name, category_number in cp_categories:
        if upper_bound > central_pressure >= lower_bound:
            return category_name, category_number
    return 'Unknown', -999

# Helper function to map categories to colors
def category_to_color(category):
    """
    Maps hurricane category to a color.
    """
    category_colors = {
        'Tropical Depression': 'green',
        'Tropical Storm': 'blue',
        'Category 1': 'yellow',
        'Category 2': 'orange',
        'Category 3': 'red',
        'Category 4': 'darkred',
        'Category 5': 'magenta',
        'Unknown': 'gray'
    }
    return category_colors.get(category, 'gray')

def read_adeck_dat_file_to_gdf(dat_file_path, model_name=None, forecast_datetime=None):

# Check if the file exists
    if not os.path.exists(dat_file_path):
        print(f"File not found: {dat_file_path}")
        return None

    # Define column names based on ATCF A-deck format
    columns = [
        'Basin', 'CycloneNumber', 'DateTime', 'ModelNumber', 'ModelName', 'ForecastHour',
        'Latitude', 'Longitude', 'MaxWindSpeed', 'MinPressure',
        'WindRad1', 'WindRad2', 'WindRad3', 'WindRad4', 'StormType',
        'Quadrant1', 'Quadrant2', 'Quadrant3', 'Quadrant4',
        'Radius1', 'Radius2', 'Radius3', 'Radius4',
        'StormName', 'Unused1', 'Unused2'
    ]
    data = []
    # Read the .dat file line by line
    try:
        with open(dat_file_path, 'r') as file:
            for line_number, line in enumerate(file, start=1):
                # Remove leading/trailing whitespace
                line = line.strip()
                # Skip empty lines
                if not line:
                    continue
                # Split the line into fields
                fields = line.split(',')
                # Strip whitespace from each field
                fields = [field.strip() for field in fields]
                # Handle variable number of fields
                # Create a dictionary for this line
                record = {}
                num_fields = len(fields)
                for i in range(min(num_fields, len(columns))):
                    record[columns[i]] = fields[i]
                # If there are extra fields, add them as 'ExtraField1', 'ExtraField2', etc.
                if num_fields > len(columns):
                    for j in range(len(columns), num_fields):
                        record[f'ExtraField{j - len(columns) + 1}'] = fields[j]
                data.append(record)
    except Exception as e:
        print(f"Error reading .dat file at line {line_number}: {e}")
        return None

    # Convert the list of dictionaries to a DataFrame
    df = pd.DataFrame(data)

    # check for bad column types
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                # Attempt to convert to numeric if possible
                df[col] = pd.to_numeric(df[col])
            except Exception as e:
                print(f"Could not convert column {col}, which is an object to numeric: {e}, TRYING TO CONVERT TO STRING NOW")

                try: 
                    # Attempt to convert to numeric if possible
                    df[col] = df[col].astype(str)
                except Exception as e:
                    print(f"Could not convert column {col}, which is an object to string or numeric: {e}")
        #elif col == 'Category':
        #    df['Category'] = df['Category'].astype(str)

    if df.empty:
        print("No data found for the selected forecast date and time.")
        return None

    # Convert columns to appropriate data types
    df['ForecastHour'] = pd.to_numeric(df['ForecastHour'], errors='coerce')
    df['MaxWindSpeed'] = pd.to_numeric(df['MaxWindSpeed'], errors='coerce')
    df['MaxWindSpeed_mph'] = df['MaxWindSpeed'] * 1.15078  # Knots to miles per hour
    df['Latitude'] = df['Latitude'].apply(parse_lat_lon)
    df['Longitude'] = df['Longitude'].apply(parse_lat_lon)
    df['DateTime'] = pd.to_datetime(df['DateTime'], format='%Y%m%d%H', errors='coerce')
    df['DateTime'] =  df['DateTime'].dt.tz_localize('UTC')

    # Filter by forecast datetime if provided
    if forecast_datetime:
        df = df[df['DateTime'] == forecast_datetime]

    # Drop rows with missing data
    before_dropna = len(df)
    df.dropna(subset=['Latitude', 'Longitude', 'MaxWindSpeed', 'DateTime', 'ForecastHour'], inplace=True)
    after_dropna = len(df)
    dropped_na = before_dropna - after_dropna
    if dropped_na > 0:
        print(f"Dropped {dropped_na} rows due to missing data.")

    if df.empty:
        print("All data has been dropped after removing rows with missing values.")
        return None

    # Remove any rows where Latitude or Longitude is exactly 0
    before_zero_filter = len(df)
    df = df[(df['Latitude'] != 0) & (df['Longitude'] != 0)]
    after_zero_filter = len(df)
    removed_zero = before_zero_filter - after_zero_filter
    if removed_zero > 0:
        print(f"Removed {removed_zero} rows with Latitude or Longitude equal to 0.")

    if df.empty:
        print("All data has been dropped after removing rows with Latitude or Longitude equal to 0.")
        return None

    # Create 'ValidTime' as DateTime + ForecastHour
    df['ValidTime'] = df['DateTime'] + pd.to_timedelta(df['ForecastHour'], unit='h')

    # Categorize intensity and CP into storm categories

    # Apply the function and expand the result into two separate columns
    df[['Category_based_on_pressure', 'Category_number_pressure_based']] = df['MinPressure'].apply(
        pressure_to_category
    ).apply(pd.Series)

    df[['Category_based_on_windspeed', 'Category_number_windspeed_based']] = df['MaxWindSpeed'].apply(
        wind_speed_to_category
    ).apply(pd.Series)

    df['maximum_category']=df[['Category_number_windspeed_based','Category_number_pressure_based']].max(axis=1)

    # Create the 'geometry' column in df
    df['geometry'] = df.apply(lambda row: Point(row['Longitude'], row['Latitude']), axis=1)

    df['Category'] = df['Category_based_on_windspeed'] #Category for coloring is by default based on windspeed. 
    

    # Create the GeoDataFrame, specifying the geometry column and CRS
    gdf = gpd.GeoDataFrame(df, geometry='geometry', crs="EPSG:4326")

    return gdf

#############################################################################################################
## FUNCTIONS THAT ARE ONLY USED BY INTERCEPT AND ALERT (intercept_and_alert_VX.py file):
#############################################################################################################

def get_recent_adeck_paths(file_path,timestamp_utc=datetime.now(timezone.utc)):
    """
    Get the adeck_path of storms whose end date is within the last 24 hours (UTC time).

    Parameters:
    - file_path: str, path to the storm_adeck_directory CSV file.
    - timestamp_utc: timestamp object, the time from which a storm is condered active (updated 24hrs prior to this time and date)

    Returns:
    - recent_adeck_paths: list of str, paths of adeck files updated in the last 24 hours.
    """
    try:
        # Read the CSV file into a DataFrame
        df = pd.read_csv(file_path)

        # Parse 'Storm_End_Date' using the custom function
        df['Storm_End_Date'] = df['Storm_End_Date'].apply(parse_datetime)

        # Localize to UTC only if 'Storm_End_Date' is valid
        df['Storm_End_Date'] = df['Storm_End_Date'].apply(lambda x: x.tz_localize('UTC') if pd.notna(x) else x)

        # Time 24 hours ago in UTC
        time_threshold = timestamp_utc - timedelta(hours=0)

        # Filter rows whose 'Storm_End_Date' is within the last 24 hours
        recent_storms = df[df['Storm_End_Date'] > time_threshold]

        return recent_storms

    except Exception as e:
        print(f"Error occurred: {e}")
        return []

def parse_datetime(value):
    """
    Attempt to parse a datetime value with multiple formats.
    """
    # Ensure value is a string; if it's NaN or float, return NaT
    if not isinstance(value, str):
        return pd.NaT
    
    datetime_formats = [
        "%Y-%m-%d %H:%M:%S",  # 2024-10-08 18:00:00
        "%d/%m/%Y %H:%M",      # 08/10/2024 18:00
        "%Y/%m/%d %H:%M:%S",   # 2024/10/08 18:00:00 (additional possible format)
    ]

    for fmt in datetime_formats:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    
    # If none of the formats work, return None (or NaT)
    return pd.NaT
    
def clean_strings(gdf, column_name, new_column_name):
    """
    Removes any words containing numbers and all following words from the specified column 
    of the GeoDataFrame unless the word is the first word in the string.
    
    Parameters:
    gdf (GeoDataFrame): The input GeoDataFrame.
    column_name (str): The name of the column to be processed.
    new_column_name (str): The name of the new column where the cleaned strings will be stored.
    
    Returns:
    GeoDataFrame: The input GeoDataFrame with an additional column containing the cleaned strings.
    """
    
    def process_string(s):
        words = s.split()
        cleaned_words = []
        for i, word in enumerate(words):
            if i == 0:
                # Always keep the first word
                cleaned_words.append(word)
            else:
                if any(char.isdigit() for char in word):
                    # Remove this word and stop processing further words
                    break
                else:
                    cleaned_words.append(word)
        return ' '.join(cleaned_words).strip()
    
    # Apply the process_string function to the specified column
    gdf[new_column_name] = gdf[column_name].apply(process_string)
    
    return gdf

def create_forecast_map_with_cone_for_AOIs(adeck_gdf, cone_polygon, aoi_gdf, selected_forecast_datetime=None, AOI_Name=None,title=None):
    """
    Creates a Plotly map of the forecast tracks, including the uncertainty cone and areas of interest.

    Parameters:
    - adeck_gdf: GeoDataFrame with point geometries of the hurricane tracks.
    - cone_polygon: Polygon or MultiPolygon geometry representing the uncertainty cone.
    - aoi_gdf: GeoDataFrame with polygon geometries representing areas of interest.
    - selected_forecast_datetime: str, the forecast date and time in 'YYYYMMDDHH' format (optional).
    - AOI_Name: str, name of the column in aoi_gdf containing labels for AOIs, or a single label (str).

    Returns:
    - fig: Plotly figure object.
    """
    mapbox_access_token = 'pk.eyJ1IjoiYWx2YXJvZmFyaWFzIiwiYSI6ImNtMXptbm9iaDA4OHMybG9vc3VqdW1vZ3oifQ.ZJ8d6gNAiR1htIYxESOYuQ'
    pio.templates.default = 'plotly_dark'  # Optional to use dark theme

    # Filter by forecast datetime if provided
    if selected_forecast_datetime:
        adeck_gdf = adeck_gdf[adeck_gdf['DateTime'] == selected_forecast_datetime]

    # Determine AOI labels
    if AOI_Name and AOI_Name in aoi_gdf.columns:
        # Use the specified column for labels
        AOI_labels = aoi_gdf[AOI_Name].astype(str).tolist()
    else:
        # Use the provided AOI_Name as a single label, or default label
        AOI_label = AOI_Name if AOI_Name else 'Area(s) of Interest'
        AOI_labels = [AOI_label] * len(aoi_gdf)

    # Define category colors in the desired order
    category_colors = {
        'Tropical Depression': 'green',
        'Tropical Storm': 'blue',
        'Category 1': 'yellow',
        'Category 2': 'orange',
        'Category 3': 'red',
        'Category 4': 'purple',
        'Category 5': 'magenta',
        'Unknown': 'gray'
    }

    # Define the desired order for the legend
    category_order = [
        'Tropical Depression',
        'Tropical Storm',
        'Category 1',
        'Category 2',
        'Category 3',
        'Category 4',
        'Category 5',
        'Unknown'
    ]

    # Map category to colors
    adeck_gdf['Color'] = adeck_gdf['Category'].apply(lambda x: category_colors.get(x, 'gray'))

    # Sort the DataFrame by ModelName and ValidTime to ensure lines connect correctly per model
    adeck_gdf = adeck_gdf.sort_values(['ModelName', 'ValidTime'])

    # Initialize the Plotly figure
    fig = go.Figure()

    # Add a line trace for each model (white lines connecting points of the same model)
    models = adeck_gdf['ModelName'].unique()
    for model in models:
        model_gdf = adeck_gdf[adeck_gdf['ModelName'] == model].sort_values('ValidTime')
        if model_gdf.empty:
            continue
        fig.add_trace(go.Scattermapbox(
            lat=model_gdf.geometry.y,
            lon=model_gdf.geometry.x,
            mode='lines',
            line=dict(color='white', width=2),
            hoverinfo='none',  # No hover info for the lines
            showlegend=False    # Do not show lines in the legend
        ))

    # Group data by Category in the specified order and plot the markers
    for category in category_order:
        category_gdf = adeck_gdf[adeck_gdf['Category'] == category]
        if not category_gdf.empty:
            fig.add_trace(go.Scattermapbox(
                lat=category_gdf.geometry.y,
                lon=category_gdf.geometry.x,
                mode='markers',
                name=category,
                marker=dict(
                    size=10,
                    color=category_colors[category],  # Single color per category
                ),
                text=category_gdf.apply(
                    lambda row: (
                        f"Time: {row['ValidTime']:%Y-%m-%d %H:%M UTC}<br>"
                        f"Wind Speed (mph): {row['MaxWindSpeed_mph']:.1f}<br>"
                        f"Central Pressure (mb): {row['MinPressure']}<br>"
                        f"Category: {row['Category']}<br>"
                        f"Model: {row['ModelName']}"
                    ), axis=1),
                hoverinfo='text'
            ))

    # Add uncertainty cone polygon to the map
    if cone_polygon is not None and not cone_polygon.empty:
        # Ensure the cone_polygon is in WGS84 coordinate system
        if cone_polygon.crs != 'EPSG:4326':
            cone_polygon = cone_polygon.to_crs('EPSG:4326')
        
        for geom in cone_polygon.geometry:
            if geom.geom_type == 'Polygon':
                cone_coords = geom.exterior.coords.xy
                cone_lats, cone_lons = list(cone_coords[1]), list(cone_coords[0])  # Convert to lists
                fig.add_trace(go.Scattermapbox(
                    lat=cone_lats,
                    lon=cone_lons,
                    mode='lines',
                    fill='toself',
                    fillcolor='rgba(128, 128, 128, 0.5)',  # Grey with transparency
                    line=dict(color='lightgrey', width=2),  # Light grey outline
                    name='Uncertainty Cone',
                    hoverinfo='skip',  # No hover info for the cone
                    # Un-comment the next line if you want the cone to be off by default
                    # visible='legendonly'  
                ))
            elif geom.geom_type == 'MultiPolygon':
                for poly in geom.geoms:
                    cone_coords = poly.exterior.coords.xy
                    cone_lats, cone_lons = list(cone_coords[1]), list(cone_coords[0])  # Convert to lists
                    fig.add_trace(go.Scattermapbox(
                        lat=cone_lats,
                        lon=cone_lons,
                        mode='lines',
                        fill='toself',
                        fillcolor='rgba(128, 128, 128, 0.5)',  # Grey with transparency
                        line=dict(color='lightgrey', width=2),  # Light grey outline
                        name='Uncertainty Cone',
                        hoverinfo='skip',
                    ))

    # Add Areas of Interest polygons to the map
    if aoi_gdf is not None and not aoi_gdf.empty:
        # Ensure the aoi_gdf is in WGS84 coordinate system
        if aoi_gdf.crs != 'EPSG:4326':
            aoi_gdf = aoi_gdf.to_crs('EPSG:4326')

        # Iterate over each AOI polygon and their labels
        for idx, (aoi, label) in enumerate(zip(aoi_gdf.itertuples(), AOI_labels)):
            geometry = aoi.geometry
            if geometry.geom_type == 'Polygon':
                aoi_coords = geometry.exterior.coords.xy
                aoi_lats, aoi_lons = list(aoi_coords[1]), list(aoi_coords[0])  # Convert to lists
                fig.add_trace(go.Scattermapbox(
                    lat=aoi_lats,
                    lon=aoi_lons,
                    mode='lines',
                    fill='toself',
                    fillcolor='rgba(0, 255, 0, 0.2)',  # Green with transparency
                    line=dict(color='green', width=2),
                    name=label,
                    hoverinfo='skip'
                ))
            elif geometry.geom_type == 'MultiPolygon':
                for poly in geometry.geoms:
                    aoi_coords = poly.exterior.coords.xy
                    aoi_lats, aoi_lons = list(aoi_coords[1]), list(aoi_coords[0])  # Convert to lists
                    fig.add_trace(go.Scattermapbox(
                        lat=aoi_lats,
                        lon=aoi_lons,
                        mode='lines',
                        fill='toself',
                        fillcolor='rgba(0, 255, 0, 0.2)',  # Green with transparency
                        line=dict(color='green', width=2),
                        name=label,
                        hoverinfo='skip'
                    ))

    # Calculate the centroid of the AOIs and extract lat/lon for centering the map
    if not aoi_gdf.empty:
        aoi_centroid = aoi_gdf.centroid  # This gives a GeoSeries of centroids
        center_lat = aoi_centroid.y.mean()  # Calculate the mean latitude of the centroids
        center_lon = aoi_centroid.x.mean()  # Calculate the mean longitude of the centroids
    else:
        # If no AOIs, fallback to the adeck_gdf centroid
        center_lat = adeck_gdf.geometry.y.mean()
        center_lon = adeck_gdf.geometry.x.mean()

    # Add a logo (SVG) to the map layout #
    
    with open('./assets/LOCKTON_logo-white-footer.svg', 'rb') as image_file:
        encoded_image = base64.b64encode(image_file.read()).decode()# Read and encode the image in base64

    fig.update_layout(
        images=[
            dict(
                source='data:image/svg+xml;base64,{}'.format(encoded_image),
                xref="paper", yref="paper",  # Relative to the plot area
                x=0.01, y=0.99,  # Position of the logo
                sizex=0.1, sizey=0.12,  # Size of the logo
                xanchor="left", yanchor="top"  # Anchor the image position
            )
        ],
        # Set up the map layout with a dark style
        mapbox=dict(
            accesstoken=mapbox_access_token,
            style='mapbox://styles/mapbox/satellite-streets-v12',  # Use full style URL
            zoom=5,
            center={"lat": center_lat, "lon": center_lon},
        ),
        mapbox_zoom=5,
        mapbox_center={"lat": center_lat, "lon": center_lon},
        margin={"r": 0, "t": 0, "l": 0, "b": 0},

        legend=dict( #legend settings
            title=dict(
                text="LEGEND",  # Legend title
                font=dict(
                    color="white",  # Title font color
                    size=16,  # Title font size
                    family="Arial",  # Set font family
                    weight="bold"  # Make only the title bold
                ),
                side="top"  # Position the title at the top and center
            ),
            font=dict(
                color="white",  # Set legend text color to white
                size=14  # Increase font size to 14 for legend items (normal weight)
            ),
            bgcolor="black",  # Set legend background to black
            yanchor="top",
            y=0.85,  # Move down
            x=0.01,  # Align to the left
            xanchor="left",
            itemsizing='constant'
        ),
            uirevision=True,  # Preserve the zoom level and state of the map
        )
    
    return fig

def send_map_via_email(fig, recipients, subject, body, cc=None, bcc=None, sender=None, send_as=False):
    """
    Sends a Plotly map as an HTML attachment via Outlook email.

    Parameters:
    - fig (plotly.graph_objects.Figure): The Plotly figure to send.
    - recipients (list or str): List of recipient email addresses or a single email address.
    - subject (str): Subject of the email.
    - body (str): Body content of the email. Supports HTML.
    - cc (list or str, optional): List of CC email addresses or a single email address.
    - bcc (list or str, optional): List of BCC email addresses or a single email address.
    - sender (str, optional): Email address to send on behalf of. Must have permissions.
    - send_as (bool, optional): If True, attempts to set the 'From' property instead of 'SentOnBehalfOfName'.

    Returns:
    - None

    Raises:
    - ValueError: If no valid recipients are provided.
    - Exception: If Outlook is not installed or an error occurs during email sending.
    """
    try:
        # Function to validate email addresses using a simple regex
        def is_valid_email(email):
            regex = r'^[\w\.-]+@[\w\.-]+\.\w+$'
            return re.match(regex, email) is not None

        # Helper function to ensure input is a list
        def ensure_list(input_field):
            if input_field is None:
                return []
            if isinstance(input_field, str):
                input_field = input_field.strip()
                return [input_field] if input_field else []
            if isinstance(input_field, list):
                # Remove any empty strings and strip whitespace
                return [email.strip() for email in input_field if email.strip()]
            raise ValueError("Email fields must be either a string or a list of strings.")

        # Convert recipients, cc, bcc to lists
        recipients = ensure_list(recipients)
        cc = ensure_list(cc)
        bcc = ensure_list(bcc)

        # Combine all recipients to ensure at least one is present
        all_recipients = recipients + cc + bcc

        if not all_recipients:
            raise ValueError("At least one recipient must be provided in To, Cc, or Bcc.")

        # Validate all email addresses
        invalid_emails = [email for email in all_recipients if not is_valid_email(email)]
        if invalid_emails:
            raise ValueError(f"The following email addresses are invalid: {invalid_emails}")

        # Create a temporary directory to store the HTML file
        with tempfile.TemporaryDirectory() as tmpdirname:
            # Define the HTML file path
            html_file_path = os.path.join(tmpdirname, 'forecast_map.html')
            
            # Save the Plotly figure as an HTML file
            fig.write_html(html_file_path, full_html=True, config={'scrollZoom': True})
            
            # Initialize Outlook application
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)  # 0: olMailItem
            
            # Set email parameters
            mail.Subject = subject
            mail.Body = body  # Plain text body
            mail.HTMLBody = body  # HTML body
            
            # Set sender if specified
            if sender:
                try:
                    if send_as:
                        mail.From = sender  # Requires 'Send As' permissions
                    else:
                        mail.SentOnBehalfOfName = sender  # Requires 'Send on Behalf Of' permissions
                except Exception as e:
                    print(f"Error setting sender to '{sender}': {e}")
                    raise

            # Add recipients to To
            for recipient in recipients:
                mail.Recipients.Add(recipient)
            
            # Add CC recipients
            for c in cc:
                mail.Recipients.Add(c)
            
            # Add BCC recipients
            for b in bcc:
                mail.Recipients.Add(b)
            
            # Resolve all recipients
            if not mail.Recipients.ResolveAll():
                unresolved = [rec.Name for rec in mail.Recipients if not rec.Resolved]
                raise ValueError(f"Some recipients could not be resolved: {unresolved}")
            
            # Attach the HTML file
            mail.Attachments.Add(Source=html_file_path)
            
            # Send the email
            mail.Send()
            
            print("Email sent successfully.")
    
    except ValueError as ve:
        print(f"ValueError: {ve}")
        raise
    except Exception as e:
        print(f"An error occurred while sending the email: {e} {html_file_path} ")
        raise

def send_email(recipients, subject, body, df=None, cc=None, bcc=None, sender=None, send_as=False):
    """
    Sends an email via Outlook using win32com.client with an optional DataFrame attachment.

    Parameters:
    - recipients (list or str): List of recipient email addresses or a single email address.
    - subject (str): Subject of the email.
    - body (str): Body content of the email. Supports HTML.
    - df (pd.DataFrame, optional): DataFrame to send as a .csv attachment.
    - cc (list or str, optional): List of CC email addresses or a single email address.
    - bcc (list or str, optional): List of BCC email addresses or a single email address.
    - sender (str, optional): Email address to send on behalf of. Must have permissions.
    - send_as (bool, optional): If True, attempts to set the 'From' property instead of 'SentOnBehalfOfName'.

    Returns:
    - None

    Raises:
    - ValueError: If no valid recipients are provided.
    - Exception: If Outlook is not installed or an error occurs during email sending.
    """
    try:
        def is_valid_email(email):
            regex = r'^[\w\.-]+@[\w\.-]+\.\w+$'
            return re.match(regex, email) is not None

        def ensure_list(input_field):
            if input_field is None:
                return []
            if isinstance(input_field, str):
                input_field = input_field.strip()
                return [input_field] if input_field else []
            if isinstance(input_field, list):
                return [email.strip() for email in input_field if email.strip()]
            raise ValueError("Email fields must be either a string or a list of strings.")

        recipients = ensure_list(recipients)
        cc = ensure_list(cc)
        bcc = ensure_list(bcc)

        all_recipients = recipients + cc + bcc
        if not all_recipients:
            raise ValueError("At least one recipient must be provided in To, Cc, or Bcc.")

        invalid_emails = [email for email in all_recipients if not is_valid_email(email)]
        if invalid_emails:
            raise ValueError(f"The following email addresses are invalid: {invalid_emails}")

        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mail.Subject = subject
        mail.Body = body
        mail.HTMLBody = body

        if sender:
            try:
                if send_as:
                    mail.SentOnBehalfOfName = sender
                else:
                    mail.SentOnBehalfOfName = sender
            except Exception as e:
                print(f"Error setting sender to '{sender}': {e}")
                raise

        for recipient in recipients:
            recipient_obj = mail.Recipients.Add(recipient)
            recipient_obj.Type = 1

        for c in cc:
            recipient_obj = mail.Recipients.Add(c)
            recipient_obj.Type = 2

        for b in bcc:
            recipient_obj = mail.Recipients.Add(b)
            recipient_obj.Type = 3

        if not mail.Recipients.ResolveAll():
            unresolved = [rec.Address for rec in mail.Recipients if not rec.Resolved]
            raise ValueError(f"Some recipients could not be resolved: {unresolved}")

        #if df is not None:
            #with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp_file:
                #csv_path = tmp_file.name
                #df.to_csv(csv_path, index=False)
                #attachment = mail.Attachments.Add(csv_path)
                #attachment.DisplayName = "data.csv"

        mail.Send()
        print("Email sent successfully.")

        #if df is not None:
        #    os.remove(csv_path)

    except ValueError as ve:
        print(f"ValueError: {ve}")
        raise
    except Exception as e:
        print(f"An error occurred while sending the email: {e}")
        raise

    except ValueError as ve:
        print(f"ValueError: {ve}")
        raise
    except Exception as e:
        print(f"An error occurred while sending the email: {e}")
        raise

#def format_adeck_parquet(df):
#    # Convert columns to appropriate data types
#    df['ForecastHour'] = pd.to_numeric(df['ForecastHour'], errors='coerce')
#    df['MaxWindSpeed'] = pd.to_numeric(df['MaxWindSpeed'], errors='coerce')
#    df['MinPressure'] = pd.to_numeric(df['MinPressure'],errors='coerce')
#    df['MaxWindSpeed_mph'] = df['MaxWindSpeed'] * 1.15078  # Knots to miles per hour
#    #df['Latitude'] = df['Latitude'].apply(parse_lat_lon)
#    #df['Longitude'] = df['Longitude'].apply(parse_lat_lon)
#    df['DateTime'] = pd.to_datetime(df['DateTime'], format='%Y%m%d%H', errors='coerce')
#    try:
#        df['DateTime'] =  df['DateTime'].dt.tz_localize('UTC')
#    except:
#        print('')
#
#    # Categorize intensity into storm categories
#    # Apply the function and expand the result into two separate columns
#    df[['Category_based_on_pressure', 'Category_number_pressure_based']] = df['MinPressure'].apply(
#        pressure_to_category
#    ).apply(pd.Series)
#
#    df[['Category_based_on_windspeed', 'Category_number_windspeed_based']] = df['MaxWindSpeed'].apply(
#        wind_speed_to_category
#    ).apply(pd.Series)
#
#    df['maximum_category']=df[['Category_number_windspeed_based','Category_number_pressure_based']].max(axis=1)
#
#    df['Category'] = df['Category_based_on_windspeed'] #Category for coloring is by default based on windspeed. 
#    
#    return(df)