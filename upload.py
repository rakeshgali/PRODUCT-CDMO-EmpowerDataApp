import dash
import base64
import io
import datetime
import unicodecsv
from dash.dependencies import Input, Output, State
import dash_design_kit as ddk
import dash_core_components as dcc
import dash_html_components as html
import dash_bootstrap_components as dbc
import dash_table
import textwrap
import pandas as pd
import Compiler
import csv
import xlsxwriter
import flask
import numpy as np
import plotly.express as px
from openpyxl import Workbook
from openpyxl.styles import Font, Color, Alignment, Border, Side, colors, numbers
from openpyxl.comments import Comment
from openpyxl.utils import units
import json

# external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']
# app = dash.Dash(__name__, external_stylesheets=external_stylesheets)
app = dash.Dash(__name__)
server = app.server  # expose server variable for Procfile
# upload files button on the left side
controls1 = [
    ddk.ControlItem(
        html.Div([
            dcc.Upload(
                id="upload-data",
                children=html.Div([
                    'Drag and drop or',
                    html.A('click to select a file to upload')
                ]),
                multiple=True,
            ),
        ],
            style={"max-width": "500px"},
        ),
        label='Files Upload',
        style={
            'borderStyle': 'dotted',
            'textAlign': 'Center'
        },
    )
]
# Text filter on the left side
controls2 = [
    ddk.ControlItem(
        dcc.Input(
            id="input_QL",
            value=0,
            type='number'
        ),
        label='QL:',
    )
]
app.layout = ddk.App([
    ddk.Header([
        ddk.Logo(src=app.get_asset_url('gilead.png')),
        ddk.Title('CDMO Empower Data', style={'textAlign': 'Center'})
    ]),
    ddk.Block(width=20, children=[
        ddk.Card(children=dcc.Markdown(textwrap.dedent(
            '''
            This application allows CDMO team to analyze and output Empower data as a formatted file.
            ''')
        )),
        ddk.ControlCard(controls1),
        ddk.ControlCard(controls2),
    ]),
    ddk.Block(width=40, children=[
        ddk.Card(children=[
            html.H3("Files uploaded"),
            html.Div(id='output'),
        ])

    ]),
    ddk.Block(width=40, children=[
        ddk.Card(children=[
            html.H3("Entered QL Value"),
            html.Ul(id="out-all-types")
        ]
        )
    ]),
    ddk.Block(width=20, children=[
        ddk.Card(children=[
            html.Button('Start Processing', id='start', n_clicks=0),
            html.Button('download', id='btn_xlsx', n_clicks=0, style={'textAlign': 'center'}),
        ]),
    ]
              ),

])

def parse_contents(contents, filename):
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string).decode('utf-8')
    rows = []
    try:
        if 'txt' in filename:
            # Assume that the user uploaded a txt file
            rows = [line.split('\t') for line in decoded.split('\r') if line.strip()]
        elif 'xls' in filename:
            # Assume that the user uploaded an excel file
            df = pd.read_excel(io.BytesIO(decoded))
            rows = df.values
    except Exception as e:
        print(e)
        return [
            'There was an error processing this file.'
        ]
    return filename, rows

@app.callback(Output('output', 'children'),
              # add dependency of start button
              Input('upload-data', 'contents'),
              State('upload-data', 'filename'),
              State('upload-data', 'last_modified'))
def update_output(list_of_contents, list_of_names, list_of_dates):
    if list_of_contents is not None:
        children = [
            parse_contents(c, n) for c, n in
            zip(list_of_contents, list_of_names)]
        df = Compiler.process_file(children, 0.1)
        return dash_table.DataTable(
            id='table_output',
            columns=[{"name": str(i), "id": str(i)} for i in df.columns],
            data=df.to_dict('records'),
        )

@app.callback(
    Output("out-all-types", "children"),
    [Input("input_QL", "value")],
)
def float_input(*vals):
    return " | ".join((str(val) for val in vals if val))


if __name__ == '__main__':
    app.run_server(debug=True)