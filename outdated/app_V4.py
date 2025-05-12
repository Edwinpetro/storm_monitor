from dash import Dash, dcc, html, Input, Output, State, no_update
from datetime import datetime
import pandas as pd
import dash_bootstrap_components as dbc
import os
import re
import app_functions as f  # Contains the create_forecast_map function
import plotly.graph_objects as go  # Import Plotly graph objects
import geopandas as gpd

# Initialize the Dash app with Bootstrap theme and enable suppress_callback_exceptions
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], suppress_callback_exceptions=True)

# Load Areas of Interest for email alert service
AOIs=gpd.read_file('./Areas_Of_Interest_For_ALERT/2024_Policies.geojson')
AOIs=gpd.GeoDataFrame(data=AOIs,geometry=AOIs.buffer(0.7),crs=AOIs.crs)

# Get the directory of the current script to handle relative paths correctly
script_dir = os.path.dirname(os.path.abspath(__file__))

# Path to Storm Adeck Directory CSV
storm_adeck_path = os.path.join(script_dir, 'storm_adeck_directory.csv')

# Load Storm Adeck data
storm_adeck_data = pd.read_csv(storm_adeck_path)

# Remove rows where 'Storm_Year' is missing
storm_adeck_data = storm_adeck_data.dropna(subset=['Storm_Year'])

# If the data is not empty, extract the basins and years
if not storm_adeck_data.empty:
    # Convert 'Storm_Year' and 'Storm_Number' columns to integers
    storm_adeck_data['Storm_Year'] = storm_adeck_data['Storm_Year'].astype(int)
    storm_adeck_data['Storm_Number'] = storm_adeck_data['Storm_Number'].astype(int)

    # Map basin codes to basin names (ensure codes are uppercase)
    basin_mapping = {
        'AL': 'Atlantic',
        'EP': 'Eastern Pacific',
        'CP': 'Central Pacific',
        'WP': 'Western Pacific',
        'IO': 'Indian Ocean',
        'SH': 'Southern Hemisphere'
    }
    # Ensure basin codes are uppercase
    storm_adeck_data['Basin'] = storm_adeck_data['Basin'].str.upper()
    # Extract unique basin codes from the data
    forecast_basins_codes = storm_adeck_data['Basin'].unique().tolist()
    # Create basin options for the dropdown
    forecast_basins = [{'label': basin_mapping.get(code, code), 'value': code} for code in forecast_basins_codes]
    # Extract unique years from the data
    forecast_years = sorted(storm_adeck_data['Storm_Year'].unique().tolist())
else:
    forecast_basins = []
    forecast_years = []

# Current year
current_year = datetime.now().year

# Initialize the layout with Tabs
app.layout = dbc.Container(
    [
        # Title
        dbc.Row(
            dbc.Col(html.H1("Lockton Storm Dashboard", className="text-center mb-4", style={'color': 'white'}))
        ),

        ## Tabs for Best Track and Forecasts ##
        dcc.Tabs(
            id='tabs',
            value='forecasts-tab',  # Set the default tab to Forecasts
            children=[
                dcc.Tab(label='Forecasts - A Decks', value='forecasts-tab'),
                dcc.Tab(label='Historical - Best Track', value='best-track-tab')
            ]
        ),

        # Tab content will be rendered here
        html.Div(id='tab-content')
    ],
    fluid=True,
    style={'backgroundColor': 'black', 'color': 'grey'}
)

# Callback to render content for each tab
@app.callback(
    Output('tab-content', 'children'),
    [Input('tabs', 'value')]
)
def render_tab_content(tab):
    if tab == 'forecasts-tab':
        # Forecast Tab content
        return html.Div([
            dbc.Row(
                [
                    # Basin Dropdown
                    dbc.Col(
                        html.Div([
                            html.P("Select a Basin:", style={'fontWeight': 'bold', 'color': 'white'}),
                            dcc.Dropdown(
                                id='basin-dropdown-forecast',
                                options=forecast_basins,
                                value='AL' if 'AL' in [option['value'] for option in forecast_basins] else None,  # Default to 'AL' if available
                                placeholder='Select Basin',
                                style={'width': '100%', 'color': 'black', 'fontWeight': 'bold'}
                            )
                        ]),
                        width=3
                    ),

                    # Year dropdown for Forecasts
                    dbc.Col(
                        html.Div([
                            html.P("Select Year:", style={'fontWeight': 'bold', 'color': 'white'}),
                            dcc.Dropdown(
                                id='year-dropdown-forecast',
                                options=[{'label': str(year), 'value': year} for year in forecast_years],
                                value=current_year if current_year in forecast_years else forecast_years[-1] if forecast_years else None,
                                placeholder='Select year',
                                style={'width': '100%', 'color': 'black', 'fontWeight': 'bold'},
                                multi=False  # Single selection only
                            )
                        ]),
                        width=3
                    ),
                    # Static definition for the Storm Dropdown for Forecasts
                    dbc.Col(
                        html.Div([
                            html.P("Select Storm ID:", style={'fontWeight': 'bold', 'color': 'white'}),
                            dcc.Dropdown(
                                id='storm-dropdown-forecast',
                                options=[],  # Empty initially, to be populated later
                                placeholder='Select storm ID',
                                style={'width': '100%', 'color': 'black', 'fontWeight': 'bold'}
                            )
                        ]),
                        width=3
                    ),
                    # Forecast Date Dropdown (include it in the initial layout)
                    dbc.Col(
                        html.Div(id='forecast-date-dropdown-div', children=[
                            html.P("Select Forecast Date and Time:", style={'fontWeight': 'bold', 'color': 'white'}),
                            dcc.Dropdown(
                                id='forecast-date-dropdown',
                                options=[],
                                placeholder='Select forecast date and time',
                                style={'width': '100%', 'color': 'black', 'fontWeight': 'bold'}
                            )
                        ]),
                        width=3
                    )
                ],
                justify='center',
                className="mb-4"
            ),
            # Add the Graph component here
            dcc.Graph(id='forecast-map',  style={'height': '800px'})
        ])

    elif tab == 'best-track-tab':
        # Best Track Tab content (unchanged)
        return html.Div([
            # Best track code
        ])

# Callback to handle the Forecasts year and storm selection
@app.callback(
    Output('storm-dropdown-forecast', 'options'),
    Output('storm-dropdown-forecast', 'value'),
    [Input('basin-dropdown-forecast', 'value'),
     Input('year-dropdown-forecast', 'value')]#,
    #prevent_initial_call=True  # Prevent the callback from firing initially
)

def update_forecast_storms(selected_basin, selected_year):
    # Log selected basin and year
    print(f"Selected Basin: {selected_basin}, Selected Year: {selected_year}")
    
    if not selected_basin or not selected_year or storm_adeck_data.empty:
        return [], None  # Return empty options and no selected value if inputs are not valid

    # Print the data types to check for any issues
    print("Data types in storm_adeck_data:", storm_adeck_data.dtypes)

    # Ensure 'Storm_Number' is treated as an integer
    storm_adeck_data['Storm_Number'] = pd.to_numeric(storm_adeck_data['Storm_Number'], errors='coerce')

    # Filter storm_adeck_data based on selected basin and year
    filtered_data = storm_adeck_data[
        (storm_adeck_data['Basin'] == selected_basin) &
        (storm_adeck_data['Storm_Year'] == selected_year)
    ].sort_values('Storm_Number')

    # Log the filtered data for debugging
    print(f"Filtered data for Basin: {selected_basin}, Year: {selected_year}")
    print(filtered_data)  # This will print the filtered data to the console

    # If no data matches, return empty dropdown
    if filtered_data.empty:
        print("No data found for the selected basin and year.")
        return [], None

    # Create storm options (including all storm numbers, even >= 90)
    storm_options = []
    for _, row in filtered_data.iterrows():
        # Create a label for the storm, e.g., "Storm_Name (Storm_Number)"
        label = f"{row['Storm_Name'].title()} ({row['Storm_Number']:02d})"
        value = f"{row['Basin']}{row['Storm_Number']:02d}"
        storm_options.append({'label': label, 'value': value})

    # Select the highest storm number less than 90 as default, if any exist
    storm_numbers_below_90 = filtered_data[filtered_data['Storm_Number'] < 90]

    if not storm_numbers_below_90.empty:
        latest_storm_number_below_90 = storm_numbers_below_90['Storm_Number'].max()
        default_storm_value = f"{selected_basin}{latest_storm_number_below_90:02d}"
    else:
        # If no storms have Storm_Number < 90, set the default to the highest storm number available
        latest_storm_number = filtered_data['Storm_Number'].max()
        default_storm_value = f"{selected_basin}{latest_storm_number:02d}"

    print(f"Returning storm options: {storm_options}")
    print(f"Default selected storm: {default_storm_value}")

    return storm_options, default_storm_value

# Callback to update the Forecast Date Dropdown based on selected storm ID and year
@app.callback(
    [Output('forecast-date-dropdown', 'options'),
     Output('forecast-date-dropdown', 'value')],
    [Input('storm-dropdown-forecast', 'value'),
     Input('year-dropdown-forecast', 'value')]#,
    #prevent_initial_call=True  # Prevent the callback from firing initially
)
def update_forecast_dates(selected_storm_id, selected_year):
    if not selected_storm_id or not selected_year or storm_adeck_data.empty:
        return [], None

    # Extract basin and storm number from selected_storm_id
    selected_basin = selected_storm_id[:2]
    selected_storm_number = int(selected_storm_id[2:])

    # Find the .dat file path from storm_adeck_data using basin, storm number, and year
    storm_row = storm_adeck_data[
        (storm_adeck_data['Basin'] == selected_basin) &
        (storm_adeck_data['Storm_Number'] == selected_storm_number) &
        (storm_adeck_data['Storm_Year'] == selected_year)
    ]

    if storm_row.empty:
        return [], None

    adeck_path = storm_row.iloc[0]['adeck_path']

    # Construct the full path to the .dat file
    adeck_full_path = os.path.abspath(os.path.join(script_dir, adeck_path.strip('./\\')))
    if not os.path.exists(adeck_full_path):
        return [], None

    # Read the .dat file and extract forecast dates
    try:
        dates = set()
        with open(adeck_full_path, 'r') as file:
            for line in file:
                # Remove leading/trailing whitespace
                line = line.strip()
                # Skip empty lines
                if not line:
                    continue
                # Split the line into fields
                fields = line.split(',')
                # Strip whitespace from each field
                fields = [field.strip() for field in fields]
                # The 'DateTime' is in the third field (index 2)
                if len(fields) >= 3:
                    date_str = fields[2]
                    date = pd.to_datetime(date_str, format='%Y%m%d%H', errors='coerce')
                    if pd.notna(date):
                        dates.add(date)
        # Get unique dates
        unique_dates = sorted(dates)
        if not unique_dates:
            return [], None
        # Create dropdown options
        date_options = [{'label': date.strftime("%Y-%m-%d %H:%M UTC"), 'value': date.strftime("%Y%m%d%H")} for date in unique_dates]
        # Default to most recent date
        default_date_value = date_options[-1]['value']
        # Return options and default value
        return date_options, default_date_value
    except Exception as e:
        print(f"Error reading file {adeck_full_path}: {e}")
        return [], None

# Callback to update the forecast map based on selected forecast date, storm ID, and year
@app.callback(
    Output('forecast-map', 'figure'),
    [Input('forecast-date-dropdown', 'value'),
     Input('storm-dropdown-forecast', 'value'),
     Input('year-dropdown-forecast', 'value')],
    prevent_initial_call=True  # Prevent the callback from firing initially
)
def update_forecast_map(selected_forecast_datetime, selected_storm_id, selected_year):
    if not selected_forecast_datetime or not selected_storm_id or not selected_year:
        return go.Figure()  # Return an empty figure

    # Extract basin and storm number from selected_storm_id
    selected_basin = selected_storm_id[:2]
    selected_storm_number = int(selected_storm_id[2:])

    # Retrieve the adeck_path from your storm_adeck_data DataFrame using basin, storm number, and year
    storm_row = storm_adeck_data[
        (storm_adeck_data['Basin'] == selected_basin) &
        (storm_adeck_data['Storm_Number'] == selected_storm_number) &
        (storm_adeck_data['Storm_Year'] == selected_year)
    ]

    if storm_row.empty:
        return go.Figure()

    adeck_path = storm_row.iloc[0]['adeck_path']
    adeck_full_path = os.path.abspath(os.path.join(script_dir, adeck_path.strip('./\\')))
    print(adeck_full_path, selected_forecast_datetime)

    selected_forecast_datetime = pd.to_datetime(selected_forecast_datetime, format='%Y%m%d%H')
    selected_forecast_datetime=selected_forecast_datetime.tz_localize('UTC')

    # Create the map figure using the create_forecast_map function
    adeck_gdf = f.read_adeck_dat_file_to_gdf(adeck_full_path)
    print(selected_forecast_datetime)
    print(adeck_gdf)
    fig = f.create_forecast_map_with_cone(adeck_gdf, selected_forecast_datetime)

    if fig is None:
        return go.Figure()  # Return an empty figure

    return fig

if __name__ == '__main__':
    app.run_server(debug=True)