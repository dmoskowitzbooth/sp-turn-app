import pandas as pd
from datetime import datetime, timedelta
import re
from pulp import *
import xlsxwriter
import plotly.express as px
from dash import Dash, dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import json
import base64
import io
import dash_draggable

def clean_time_string(time_str):
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

# Layout
app.layout = dbc.Container([
    dbc.Row([
        dbc.Col([
            html.H1("Train Schedule Optimizer"),
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
            html.Button('Update Schedule', id='update-button', n_clicks=0),
            html.Hr(),
            dcc.Graph(id='gantt-chart'),
            dcc.Store(id='schedule-data')
        ])
    ])
])

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
    Output('schedule-data', 'data'),
    [Input('upload-data', 'contents'),
     Input('minturntime-input', 'value'),
     Input('sets-input', 'value'),
     Input('update-button', 'n_clicks')],
    [State('upload-data', 'filename'),
     State('upload-data', 'last_modified')]
)
def update_output(contents, minturntime, sets, n_clicks, filename, last_modified):
    if contents is None:
        return {}

    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    df = pd.read_excel(io.BytesIO(decoded), sheet_name='turn', header=None)

    minturntime = int(minturntime)-1 if minturntime else 54

    # Rest of the logic to process the Excel file and compute the schedule
    # Load the Excel file from the uploaded content
    df = pd.read_excel(io.BytesIO(decoded), sheet_name='turn', header=None)

    # Find the row indices where the tables start (look for the "Station" header)
    table_start_indices = df.index[df.iloc[:, 0].str.contains('Station', na=False, case=False)].tolist()

    # Add the index after the last row for proper slicing
    table_start_indices.append(len(df))

    SCHEDULE = {}

    # Process each table separately
    for i in range(len(table_start_indices) - 1):
        start_index = table_start_indices[i]
        end_index = table_start_indices[i + 1]
        table_df = df.iloc[start_index:end_index].reset_index(drop=True)
        
        # Manually set the column names using the first row
        column_names = table_df.iloc[0]
        table_df.columns = column_names
        table_df = table_df.drop(0).reset_index(drop=True)  # Drop the first row (which is now the header row)

        # Remove columns that are entirely NaN
        table_df = table_df.dropna(axis=1, how='all')

        if 'Station' not in table_df.columns:
            continue

        # Process each train separately
        for train in table_df.columns[1:]:
            train_data = table_df[['Station', train]].dropna()

            if not train_data.empty:
                clean_train = train
                
                # First and last non-null times
                departure_time = train_data[train].iloc[0]
                arrival_time = train_data[train].iloc[-1]
                
                # Convert to minutes since midnight
                departure_time_minutes = convert_to_minutes_since_midnight(departure_time)
                arrival_time_minutes = convert_to_minutes_since_midnight(arrival_time)
                
                # Adjust for overnight trips
                if arrival_time_minutes is not None and departure_time_minutes is not None and arrival_time_minutes < departure_time_minutes:
                    arrival_time_minutes += 1440  # Add 24 hours worth of minutes
                
                # Corresponding stations
                departure_station = train_data['Station'].iloc[0]
                arrival_station = train_data['Station'].iloc[-1]
                
                # Store in the dictionary only if times are valid
                if departure_time_minutes is not None and arrival_time_minutes is not None:
                    SCHEDULE[int(clean_train)] = [departure_station, arrival_station, departure_time_minutes, arrival_time_minutes]

    # Define stations based on SCHEDULE
    STATIONS = []
    onlyouts = []
    turninbound = []
    turnoutbound = []
    finaloutbounds = []

    for train in SCHEDULE:
        if SCHEDULE[train][0] not in STATIONS:
            STATIONS.append(SCHEDULE[train][0])
        if SCHEDULE[train][1] not in STATIONS:
            STATIONS.append(SCHEDULE[train][1])

    # Show Failed Solution; to suppress use =0, to show use =1
    showfail = 0

    # Add constraints to refuse a particular start --- use = [#,#,...]
    nostart = []

    # Add constraints to force a particular start --- use = [#,#,...]
    addstart = []

    # Force turns --- use structure = [[#,#],[#,#],...]
    forceturn = []

    # Refuse turns --- use structure = [[#,#],[#,#],...]
    refuseturn = []

    # Define general and specific minimum turn time in minutes
    stationturntimemindict = {}

    # Define number of available sets
    setcount = int(sets)
    SETS = ["set" + str(i + 1) for i in range(setcount)]
    train_to_set = {}

    # Initialize high negative value for any turn, next loop will reset legal turns to positive values
    TURNTIMEDICT = {t: {u: -1000000 for u in SCHEDULE} for t in SCHEDULE}

    for t in SCHEDULE:
        for u in SCHEDULE:
            if SCHEDULE[t][1] in stationturntimemindict:
                tempminturntime = stationturntimemindict[SCHEDULE[t][1]]
            else:
                tempminturntime = minturntime
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

    for train in nostart:
        for noset in SETS:
            turnprob += starts_vars[(train, noset)] == 0

    setit = 0
    for train in addstart:
        turnprob += starts_vars[(train, SETS[setit])] == 1
        setit += 1

    for fturn in forceturn:
        turnprob += turns_vars[(fturn[0], fturn[1])] == 1

    for rturn in refuseturn:
        turnprob += turns_vars[(rturn[0], rturn[1])] == 0

    # All computation happens here
    turnprob.solve()

    for s in SETS:
        for u in SCHEDULE:
            if starts_vars[(u, s)].varValue > .01:
                current_train = u
                while current_train:
                    train_to_set[current_train] = s
                    next_trains = [v for v in SCHEDULE if turns_vars[(current_train, v)].varValue > .01]
                    current_train = next_trains[0] if next_trains else None

    # Prepare the Gantt chart data
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
    fig = px.timeline(df, x_start='Start', x_end='Finish', y='Resource', color='Station', text='Label')
    fig.update_yaxes(title='Train Sets')
    fig.update_xaxes(title='Time')
    fig.update_traces(textposition='inside', textfont_size=12)
    return fig

if __name__ == '__main__':
    app.run_server(debug=True)
