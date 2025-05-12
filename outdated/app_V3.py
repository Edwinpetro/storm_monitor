from dash import Dash, dcc, html, Input, Output, State, no_update
from datetime import datetime
import pandas as pd
import dash_bootstrap_components as dbc
import os
import re
import app_functions as f  # Contains the create_forecast_map function
import plotly.graph_objects as go  # Import Plotly graph objects

# Initialize the Dash app with Bootstrap theme and enable suppress_callback_exceptions
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP], suppress_callback_exceptions=True)

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
                    # Storm Dropdown for Forecasts
                    dbc.Col(
                        html.Div(id='storm-dropdown-forecast-div'),
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
    Output('storm-dropdown-forecast-div', 'children'),
    [Input('basin-dropdown-forecast', 'value'),
     Input('year-dropdown-forecast', 'value')]
)
def update_forecast_storms(selected_basin, selected_year):
    if not selected_basin or not selected_year or storm_adeck_data.empty:
        return no_update

    # Filter storm_adeck_data based on selected basin and year
    filtered_data = storm_adeck_data[
        (storm_adeck_data['Basin'] == selected_basin) &
        (storm_adeck_data['Storm_Year'] == selected_year)
    ].sort_values('Storm_Number')

    # Create storm options
    storm_options = []
    for _, row in filtered_data.iterrows():
        # Create a label for the storm, e.g., "Storm_Name (Storm_Number)"
        label = f"{row['Storm_Name'].title()} ({row['Storm_Number']:02d})"
        value = f"{row['Basin']}{row['Storm_Number']:02d}"
        storm_options.append({'label': label, 'value': value})

    # Select the storm with the highest number by default
    if not filtered_data.empty:
        latest_storm_number = filtered_data['Storm_Number'].max()
        default_storm_value = f"{selected_basin}{latest_storm_number:02d}"
    else:
        default_storm_value = None  # No storms available

    # Return Storm ID dropdown for Forecasts
    return html.Div([
        html.P("Select Storm ID:", style={'fontWeight': 'bold', 'color': 'white'}),
        dcc.Dropdown(
            id='storm-dropdown-forecast',
            options=storm_options,
            value=default_storm_value,
            placeholder='Select storm ID',
            style={'width': '100%', 'color': 'black', 'fontWeight': 'bold'}
        )
    ])

# Callback to update the Forecast Date Dropdown based on selected storm ID and year
@app.callback(
    [Output('forecast-date-dropdown', 'options'),
     Output('forecast-date-dropdown', 'value')],
    [Input('storm-dropdown-forecast', 'value'),
     Input('year-dropdown-forecast', 'value')]
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
     Input('year-dropdown-forecast', 'value')]
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
    print(adeck_full_path,selected_forecast_datetime)
    # Create the map figure using the create_forecast_map function
    fig = f.create_forecast_map(adeck_full_path, selected_forecast_datetime)

    if fig is None:
        return go.Figure()  # Return an empty figure

    return fig

if __name__ == '__main__':
    app.run_server(debug=True)