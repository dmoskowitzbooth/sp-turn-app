import pandas as pd
from datetime import datetime, timedelta, time as dt_time
import re
from pulp import *
import xlsxwriter
import plotly.express as px
from dash import Dash, dcc, html, Input, Output, State, callback_context
from dash.dependencies import ALL
import dash_bootstrap_components as dbc
import json
import base64
import io
import uuid
import dash

def clean_time_string(time_str):
    # Ensure the input is treated as a string
    if isinstance(time_str, dt_time):
        time_str = time_str.strftime('%I:%M %p')
    elif isinstance(time_str, datetime):
        time_str = time_str.strftime('%I:%M %p')
    else:
        time_str = str(time_str)

    # Attempt to parse the time string in various expected formats
    for fmt in ['%I:%M:%S %p', '%I:%M %p', '%I:%M%p', '%I:%M:%S%p']:
        try:
            return datetime.strptime(time_str, fmt).strftime('%I:%M%p')
        except ValueError:
            pass

    # Handle the '5:00A' format
    cleaned_time_str = re.sub(r'^[^\d]*', '', time_str)
    cleaned_time_str = re.sub(r'(\d+:\d+)([AP])$', r'\1\2M', cleaned_time_str)

    return cleaned_time_str

def convert_to_minutes_since_midnight(time_str):
    cleaned_time_str = clean_time_string(time_str)
    if not cleaned_time_str:
        return None
    try:
        time_obj = datetime.strptime(cleaned_time_str, '%I:%M%p')
        minutes_since_midnight = time_obj.hour * 60 + time_obj.minute
        return minutes_since_midnight
    except ValueError:
        return None

def minutes_to_hhmm(minutes):
    hours = minutes // 60
    minutes = minutes % 60
    return f'{hours:02d}:{minutes:02d}'

def minutes_to_datetime(minutes):
    return (datetime(2024, 1, 1) + timedelta(minutes=minutes)).strftime('%Y-%m-%d %H:%M')

# Initialize the Dash app
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.title="Train Schedule Optimizer"

# Layout
app.layout = dbc.Container([
    dbc.Row([
        dbc.Col([
            html.H1("Train Schedule Optimizer"),
            html.Div([
                html.H3("Instructions:"),
                html.Ol([
                    html.Li([
                        "Download the ", 
                        html.A("template file", href="#", id="download-template-link"),
                        "."
                    ]),
                    html.Li("Detailed instructions and troubleshooting steps can be found in the 'Instructions' sheet of this document."),
                    html.Li("Fill in the 'turn' sheet's fields for train numbers (highlighted in green), stations (under the station header), and times. Times can be input as hh:mm in 24-hour format (e.g. 15:00) or hh:mm A/P in 12-hour format (e.g. 3:00 P). They will auto-format to the necessary formatting so long as they're input in one of these formats."),
                    html.Li("Once you've filled in the excel document, save it as an Excel Workbook (.xlsx) file (the name doesn't matter as long as it's the correct filed type). Upload it to the app. "),
                    html.Li("Input the number of sets and default minimum turn time. If you need to set turn times for specific stations, you can do so after this by selecting a station from the dropdown and inputting the turn time for that station only."),
                    html.Li("The app will run the optimizer to determine the best solution given the parameters. This may take up to 60 seconds depending on how large the dataset is. You can check the status by looking at the title of the tab in the browser."),
                    html.Li("Once complete, the chart will provide a visual of the solution. If there are overlapping trains, this means there is not a viable solution with the parameters you entered (see the troubleshooting part of the 'Instructions' sheet in the template). You can download an excel file with the detailed results."),
                ])
            ]),
            dcc.Upload(
                id='upload-data',
                children=html.Div([
                    'Drag and Drop or ',
                    html.A('Select Files')
                ], id='upload-text'),
                style={
                    'width': '100%',
                    'height': '60px',
                    'lineHeight': '60px',
                    'borderWidth': '1px',
                    'borderStyle': 'dashed',
                    'borderRadius': '5px',
                    'textAlign': 'center',
                    'margin': '10px'
                },
                multiple=False
            ),
            dbc.Input(id='sets-input', placeholder='Enter Number of Sets', type='number'),
            dbc.Input(id='minturntime-input', placeholder='Enter Minimum Turn Time (minutes)', type='number'),
            html.Div(id='station-turn-time-inputs', children=[
                html.Label('Specify minimum turn times for selected stations:'),
                dcc.Dropdown(id='station-dropdown', placeholder='Select a Station'),
                dbc.Input(id='station-turntime-input', placeholder='Enter Turn Time (minutes)', type='number'),
                html.Button('Add Turn Time Override', id='add-turntime-button', n_clicks=0),
                html.Div(id='turntime-overrides')
            ]),
            html.Button('Update Schedule', id='update-button', n_clicks=0),
            html.Hr(),
            dcc.Graph(id='gantt-chart'),
            dcc.Store(id='schedule-data'),
            dcc.Store(id='turntime-overrides-store', data=[]),  # Store to keep track of overrides
            html.Button('Download Excel', id='download-button', n_clicks=0),
            dcc.Download(id='download-data'),
            dcc.Download(id='download-template')  # Add hidden download component
        ])
    ])
])

@app.callback(
    Output('download-template', 'data'),
    [Input('download-template-link', 'n_clicks')],
    prevent_initial_call=True
)
def download_template_link(n_clicks):
    # Path to your template file
    template_path = 'template.xltx'
    
    # Read the template file
    with open(template_path, 'rb') as f:
        data = f.read()
    
    return dcc.send_bytes(data, 'template.xltx')

@app.callback(
    Output('upload-text', 'children'),
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def update_upload_box(contents, filename):
    if contents is not None:
        return html.Div([
            'File uploaded: ',
            html.Span(filename, style={'fontWeight': 'bold'})
        ])
    return 'Drag and Drop or Select Files'

# Callback to parse the uploaded file and process the data
@app.callback(
    Output('station-dropdown', 'options'),
    Input('upload-data', 'contents'),
    State('upload-data', 'filename'),
    State('upload-data', 'last_modified')
)
def update_station_dropdown(contents, filename, last_modified):
    if contents is None:
        return []

    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    df = pd.read_excel(io.BytesIO(decoded), sheet_name='turn', header=None)

    # Extract unique stations
    table_start_indices = df.index[df.iloc[:, 0].str.contains('Station', na=False, case=False)].tolist()
    table_start_indices.append(len(df))
    stations = set()
    for i in range(len(table_start_indices) - 1):
        start_index = table_start_indices[i]
        end_index = table_start_indices[i + 1]
        table_df = df.iloc[start_index:end_index].reset_index(drop=True)
        column_names = table_df.iloc[0]
        table_df.columns = column_names
        table_df = table_df.drop(0).reset_index(drop=True)
        table_df = table_df.dropna(axis=1, how='all')

        if 'Station' in table_df.columns:
            stations.update(table_df['Station'].dropna().unique())

    # Create dropdown options
    station_options = [{'label': station, 'value': station} for station in sorted(stations)]
    return station_options

@app.callback(
    Output('turntime-overrides-store', 'data'),
    Output('turntime-overrides', 'children'),
    Input('add-turntime-button', 'n_clicks'),
    Input({'type': 'remove-button', 'index': ALL}, 'n_clicks'),
    State('station-dropdown', 'value'),
    State('station-turntime-input', 'value'),
    State('turntime-overrides-store', 'data')
)
def update_turntime_overrides(add_clicks, remove_clicks, station, turn_time, overrides):
    ctx = dash.callback_context

    if not ctx.triggered:
        return overrides, []

    button_id = ctx.triggered[0]['prop_id'].split('.')[0]

    if 'add-turntime-button' in button_id and station and turn_time:
        new_override = {'id': str(uuid.uuid4()), 'station': station, 'turn_time': turn_time}
        overrides.append(new_override)
    else:
        remove_index = eval(button_id)['index']
        overrides = [override for override in overrides if override['id'] != remove_index]

    children = [
        html.Div([
            f"{override['station']}: {override['turn_time']} minutes",
            html.Button('Remove', id={'type': 'remove-button', 'index': override['id']}, n_clicks=0)
        ], style={'margin-top': '10px'}, id=override['id']) for override in overrides
    ]

    return overrides, children

@app.callback(
    Output('schedule-data', 'data'),
    Input('upload-data', 'contents'),
    Input('minturntime-input', 'value'),
    Input('sets-input', 'value'),
    Input('update-button', 'n_clicks'),
    State('upload-data', 'filename'),
    State('upload-data', 'last_modified'),
    State('turntime-overrides-store', 'data')
)
def update_output(contents, minturntime, sets, n_clicks, filename, last_modified, turntime_overrides):
    if contents is None:
        return {}

    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    df = pd.read_excel(io.BytesIO(decoded), sheet_name='turn', header=None)

    minturntime = int(minturntime) - 1 if minturntime else 54

    # Process turn time overrides
    station_turntime_dict = {}
    for override in turntime_overrides:
        station_turntime_dict[override['station']] = override['turn_time']

    df = pd.read_excel(io.BytesIO(decoded), sheet_name='turn', header=None)
    table_start_indices = df.index[df.iloc[:, 0].str.contains('Station', na=False, case=False)].tolist()
    table_start_indices.append(len(df))

    SCHEDULE = {}
    for i in range(len(table_start_indices) - 1):
        start_index = table_start_indices[i]
        end_index = table_start_indices[i + 1]
        table_df = df.iloc[start_index:end_index].reset_index(drop=True)
        column_names = table_df.iloc[0]
        table_df.columns = column_names
        table_df = table_df.drop(0).reset_index(drop=True)
        table_df = table_df.dropna(axis=1, how='all')

        if 'Station' not in table_df.columns:
            continue

        for train in table_df.columns[1:]:
            train_data = table_df[['Station', train]].dropna()

            if not train_data.empty:
                clean_train = train
                departure_time = train_data[train].iloc[0]
                arrival_time = train_data[train].iloc[-1]
                departure_time_minutes = convert_to_minutes_since_midnight(departure_time)
                arrival_time_minutes = convert_to_minutes_since_midnight(arrival_time)

                if arrival_time_minutes is not None and departure_time_minutes is not None and arrival_time_minutes < departure_time_minutes:
                    arrival_time_minutes += 1440

                departure_station = train_data['Station'].iloc[0]
                arrival_station = train_data['Station'].iloc[-1]

                if departure_time_minutes is not None and arrival_time_minutes is not None:
                    SCHEDULE[int(clean_train)] = [departure_station, arrival_station, departure_time_minutes, arrival_time_minutes]

    STATIONS = list(set([SCHEDULE[train][0] for train in SCHEDULE] + [SCHEDULE[train][1] for train in SCHEDULE]))
    setcount = int(sets)
    SETS = ["set" + str(i + 1) for i in range(setcount)]
    train_to_set = {}

    TURNTIMEDICT = {t: {u: -1000000 for u in SCHEDULE} for t in SCHEDULE}
    for t in SCHEDULE:
        for u in SCHEDULE:
            tempminturntime = minturntime
            if SCHEDULE[t][1] in station_turntime_dict:
                tempminturntime = station_turntime_dict[SCHEDULE[t][1]]
            if SCHEDULE[t][1] == SCHEDULE[u][0] and SCHEDULE[t][3] < -tempminturntime + SCHEDULE[u][2]:
                TURNTIMEDICT[t][u] = SCHEDULE[u][2] - SCHEDULE[t][3]

    turnprob = LpProblem("VitoCapuanoTurns", LpMaximize)
    turns_vars = LpVariable.dicts("Turns", [(t, u) for t in SCHEDULE for u in SCHEDULE], 0, 1, LpBinary)
    starts_vars = LpVariable.dicts("Starts", [(u, s) for u in SCHEDULE for s in SETS], 0, 1, LpBinary)
    sum_train_starts_vars = LpVariable.dicts("SumTrainStarts", [u for u in SCHEDULE], 0, 1, LpBinary)

    turnprob += lpSum(turns_vars[(t, u)] * TURNTIMEDICT[t][u] for t in SCHEDULE for u in SCHEDULE)

    for u in SCHEDULE:
        turnprob += lpSum(starts_vars[(u, s)] for s in SETS) == sum_train_starts_vars[u]

    for s in SETS:
        turnprob += lpSum(starts_vars[(u, s)] for u in SCHEDULE) <= 1

    for u in SCHEDULE:
        turnprob += lpSum(turns_vars[(t, u)] for t in SCHEDULE) + sum_train_starts_vars[u] == 1

    for t in SCHEDULE:
        turnprob += lpSum(turns_vars[(t, u)] for u in SCHEDULE) <= 1

    turnprob.solve()

    for s in SETS:
        for u in SCHEDULE:
            if starts_vars[(u, s)].varValue > .01:
                current_train = u
                while current_train:
                    train_to_set[current_train] = s
                    next_trains = [v for v in SCHEDULE if turns_vars[(current_train, v)].varValue > .01]
                    current_train = next_trains[0] if next_trains else None

    gantt_data = []
    for train, details in SCHEDULE.items():
        gantt_data.append({
            'Task': f'Train {train}',
            'Start': minutes_to_datetime(details[2]),
            'Finish': minutes_to_datetime(details[3]),
            'Resource': train_to_set.get(train, ''),
            'Station': details[0],  # Departure station for coloring
            'Label': f'{details[0]} > {train} > {details[1]}'  # Departure station, Train number, Arrival station
        })

    return gantt_data

# Callback to update the Gantt chart
@app.callback(
    Output('gantt-chart', 'figure'),
    [Input('schedule-data', 'data')]
)
def update_gantt_chart(data):
    if not data:
        return px.timeline()

    df = pd.DataFrame(data)
    df['Resource_numeric'] = df['Resource'].str.extract(r'(\d+)$').astype(int)  # Extract numeric part and convert to int
    df = df.sort_values(by='Resource_numeric')  # Sort by numeric part of 'Resource' column

    # Define category order based on sorted numeric values
    category_order = df.sort_values(by='Resource_numeric')['Resource'].unique()

    fig = px.timeline(df, x_start='Start', x_end='Finish', y='Resource', color='Station', text='Label', 
                      category_orders={'Resource': category_order})
    fig.update_yaxes(title='Train Sets')
    fig.update_xaxes(title='Time')
    fig.update_traces(textposition='inside', textfont_size=12)
    return fig

# Callback to generate and download the Excel file
@app.callback(
    Output('download-data', 'data'),
    [Input('download-button', 'n_clicks')],
    [State('schedule-data', 'data')],
    prevent_initial_call=True
)
def generate_excel(n_clicks, schedule_data):
    if not schedule_data:
        return dcc.send_data_frame(pd.DataFrame().to_excel, 'schedule.xlsx')

    # Process the data into the required format
    df = pd.DataFrame(schedule_data)
    df['Resource_numeric'] = df['Resource'].str.extract(r'(\d+)$').astype(int)  # Extract numeric part and convert to int
    df = df.sort_values(by='Resource_numeric')  # Sort by numeric part of 'Resource' column

    # Convert 'Start' and 'Finish' columns to datetime
    df['Start'] = pd.to_datetime(df['Start'])
    df['Finish'] = pd.to_datetime(df['Finish'])

    # Sheet 1: "turns"
    turns_data = []
    for set_ in df['Resource'].unique():
        set_trains = df[df['Resource'] == set_].sort_values(by='Start')
        start_station = set_trains.iloc[0]['Station']
        last_train_label = set_trains.iloc[-1]['Label']
        end_station = last_train_label.split('>')[-1].strip()
        turns = ' > '.join(set_trains['Task'].str.extract(r'Train (\d+)')[0].tolist())
        turns_data.append({'Set': set_, 'Start Station': start_station, 'Turns': turns, 'End Station': end_station})

    turns_df = pd.DataFrame(turns_data)

    # Sheet 2: "stations"
    start_station_counts = turns_df['Start Station'].value_counts().reset_index()
    start_station_counts.columns = ['Station', 'Trains Start']
    end_station_counts = turns_df['End Station'].value_counts().reset_index()
    end_station_counts.columns = ['Station', 'Trains End']

    # Merge the start and end counts into one DataFrame
    stations_df = pd.merge(start_station_counts, end_station_counts, on='Station', how='outer').fillna(0)
    stations_df['Trains Start'] = stations_df['Trains Start'].astype(int)
    stations_df['Trains End'] = stations_df['Trains End'].astype(int)

    # Sheet 3: "schedule"
    schedule_data = []
    max_trains_per_set = df.groupby('Resource').size().max()  # Find the maximum number of trains in any set

    for set_ in df['Resource'].unique():
        set_trains = df[df['Resource'] == set_].sort_values(by='Start')
        schedule_row = [set_]
        for i in range(max_trains_per_set):
            if i < len(set_trains):
                train = set_trains.iloc[i]
                label_parts = train['Label'].split(' > ')
                departure_station = label_parts[0]
                train_number = label_parts[1]
                arrival_station = label_parts[-1]
                departure_time = train['Start'].strftime('%H:%M')  # Format time without date
                arrival_time = train['Finish'].strftime('%H:%M')  # Format time without date
                if i < len(set_trains) - 1:
                    next_departure_time = set_trains.iloc[i + 1]['Start']
                    turn_time = (next_departure_time - train['Finish']).total_seconds() / 60  # in minutes
                else:
                    turn_time = None
                schedule_row.extend([departure_station, departure_time, train_number, arrival_station, arrival_time, turn_time])
            else:
                # Fill with empty values if there are fewer trains than the max
                schedule_row.extend([None, None, None, None, None, None])

        schedule_data.append(schedule_row)

    schedule_columns = ['Set']
    for i in range(max_trains_per_set):
        schedule_columns.extend([
            f'Departure Station {i+1}', f'Departure Time {i+1}', f'Train Number {i+1}', f'Arrival Station {i+1}', f'Arrival Time {i+1}', f'Turn Time {i+1}'
        ])

    schedule_df = pd.DataFrame(schedule_data, columns=schedule_columns)

    # Create the Excel file
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    turns_df.to_excel(writer, index=False, sheet_name='Turns')
    schedule_df.to_excel(writer, index=False, sheet_name='Schedule')
    stations_df.to_excel(writer, index=False, sheet_name='Stations')

    writer.close()
    output.seek(0)

    return dcc.send_bytes(output.read(), "schedule.xlsx")


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8050))
    app.run_server(debug=False, host='0.0.0.0', port=port)
