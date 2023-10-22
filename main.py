# Import necessary libraries
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import pandas as pd
import io
import base64
from collections import Counter

# Setting display options to None means showing all rows/columns
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)

# Create Dash App
app = dash.Dash(__name__)

# Layout of the app
app.layout = html.Div([
    # Two columns for Headers and Upload Components
    html.Div([
        # First column (XLS Upload Component)
        html.Div([
            html.H1("Spectra ankomstliste", style={'textAlign': 'center'}),
            dcc.Upload(
                id='upload-data',
                children=html.Div([
                    'Drag and Drop or ',
                    html.A('Select File')
                ]),
                style={
                    'width': '100%',
                    'height': '60px',
                    'lineHeight': '60px',
                    'borderWidth': '1px',
                    'borderStyle': 'dashed',
                    'borderRadius': '5px',
                    'textAlign': 'center',
                },
                multiple=False
            )
        ], style={'width': '48%', 'display': 'inline-block'}),
        
        # Second column (CSV Upload Component)
        html.Div([
            html.H1("SiteMinder ankomstliste", style={'textAlign': 'center'}),
            dcc.Upload(
                id='upload-csv',
                children=html.Div([
                    'Drag and Drop or ',
                    html.A('Select CSV File')
                ]),
                style={
                    'width': '100%',
                    'height': '60px',
                    'lineHeight': '60px',
                    'borderWidth': '1px',
                    'borderStyle': 'dashed',
                    'borderRadius': '5px',
                    'textAlign': 'center',
                },
                multiple=False
            )
        ], style={'width': '48%', 'display': 'inline-block'})
    ], style={"display": "flex", "justify-content": "space-between"}),
    
    # Div to show the difference in names
    html.Div(id='difference-output', style={"marginTop": "50px", "width": "100%"}),
    
    # Two columns for displaying XLS and CSV contents
    html.Div([
        # First column for XLS contents
        html.Div(id='output-data-upload', style={'width': '48%', 'display': 'inline-block'}),
        
        # Second column for CSV contents
        html.Div(id='output-csv-upload', style={'width': '48%', 'display': 'inline-block'})
    ], style={"display": "flex", "justify-content": "space-between", "marginTop": "50px"})
], style={'margin': '20px'})


@app.callback(
    [Output('output-data-upload', 'children'),
     Output('output-csv-upload', 'children'),
     Output('difference-output', 'children')],
    [Input('upload-data', 'contents'),
     Input('upload-data', 'filename'),
     Input('upload-csv', 'contents'),
     Input('upload-csv', 'filename')]
)
def update_output(contents_xls, filename_xls, contents_csv, filename_csv):
    xls_div, csv_div = [], []
    df_xls, df_csv = None, None

    # Processing XLS
    if contents_xls is not None:
        content_string = contents_xls.split(',')[1]
        decoded = base64.b64decode(content_string)
        try:
            if filename_xls and (filename_xls.lower().endswith('.xls') or filename_xls.lower().endswith('.xlsx')):
                df_xls = pd.read_excel(io.BytesIO(decoded), skiprows=1, engine='xlrd')
                df_xls = df_xls.drop(columns=['Ankomst', 'Værelsenummer', 'Vip', 'Tekst100', 'GrpVirk', 'Land', 'Antal', 'VærelseType', 'AntalVoksen', 'AnkomstTid', 'PrisKode', 'DPris', 'Markedgr.', 'Arrangement', 'Udlån', 'Bemærkning1', 'Beskeder', 'Afrejse'])
                
                def format_name(name):
                    if not isinstance(name, str):
                        return name
                    name_parts = name.strip(',').split(',')
                    formatted_name = ' '.join(reversed(name_parts)).strip()
                    return formatted_name

                df_xls['Navn'] = df_xls['Navn'].apply(format_name)
                xls_div = [html.H5("AnkostListe file Contents:"), html.Pre(df_xls.to_string())]
            else:
                xls_div = ['Only XLS files are accepted.']
        except Exception as e:
            xls_div = ['There was an error processing the XLS file: ', str(e)]
    else:
        xls_div = ['Ingen XLS fil uploadet.']

    # Processing CSV
    if contents_csv is not None:
        content_string = contents_csv.split(',')[1]
        decoded = base64.b64decode(content_string)
        try:
            if filename_csv and filename_csv.lower().endswith('.csv'):
                df_csv = pd.read_csv(io.BytesIO(decoded))
                df_csv = df_csv.drop(columns=['Affiliated Channel', 'Booking reference', 'Check-in', 'Check-out', 'Booked-on date'])
                df_csv['Guest names'] = df_csv['Guest names'].str.split(',').apply(lambda x: [i.strip() for i in x])
                df_csv = df_csv.explode('Guest names').reset_index(drop=True)
                csv_div = [html.H5("SiteMinder file Contents:"), html.Pre(df_csv.to_string())]
            else:
                csv_div = ['Only CSV files are accepted.']
        except Exception as e:
            csv_div = ['There was an error processing the CSV file: ', str(e)]
    else:
        csv_div = ['Ingen CSV fil uploadet.']

    # Find difference in names and counts
    xls_names = df_xls['Navn'].tolist() if df_xls is not None and 'Navn' in df_xls else []
    csv_names = df_csv['Guest names'].tolist() if df_csv is not None and 'Guest names' in df_csv else []

    xls_counter = Counter(xls_names)
    csv_counter = Counter(csv_names)
    difference_names = []

    for name, count in csv_counter.items():
        if name not in xls_counter:
            difference_names.extend([name] * count)
        elif count > xls_counter[name]:
            difference_names.extend([name] * (count - xls_counter[name]))

    # Check if files are uploaded
    if df_xls is None or df_csv is None:
        difference_div = html.Div(['Upload venligst begge ankomstlister'])
    elif not difference_names:
        difference_div = html.Div(['All names from CSV are in XLS.'])
    else:
        difference_div = html.Div([
            html.H5("Navne i 'SiteMinder' der ikke findes i 'AnkomstListe'"),
            html.Pre(', '.join(difference_names))
        ])

    return xls_div, csv_div, difference_div

if __name__ == '__main__':
    app.run_server(debug=True)
