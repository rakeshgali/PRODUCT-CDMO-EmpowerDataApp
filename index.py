import dash
import base64
import io
# import datetime
# import unicodecsv
from dash.dependencies import Input, Output, State
import dash_design_kit as ddk
import dash_core_components as dcc
import dash_html_components as html
# import dash_bootstrap_components as dbc
import dash_table
import textwrap
import pandas as pd
import Compiler
import csv
import xlsxwriter

# import flask
# import numpy as np
# import plotly.express as px
# from openpyxl import Workbook
# from openpyxl.styles import Font, Color, Alignment, Border, Side, colors, numbers
# from openpyxl.comments import Comment
# from openpyxl.utils import units
# import json

app = dash.Dash(__name__)
server = app.server  # expose server variable for Procfile
# upload files button on the left side
controls1 = [
    ddk.ControlItem(
        html.Div([
            dcc.Upload(
                id="upload-data",
                children=html.Div([
                    'Drag and drop or ',
                    html.A('click to select a file to upload'),

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
            value=0.1,
            type='number',
            min=0.0,
            max=0.2

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
    ddk.Block(width=80, children=[
        html.Div(children=[
            html.H3("Files uploaded", style={'textAlign': 'Center'}),
            html.Div(id='output'),
        ])
    ]),
    ddk.Block(width=20, children=[
        ddk.Card(children=[
            html.Button('Start Processing', id='start', n_clicks=0),
        ]),
    ]),

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
              Input('input_QL', 'value'),
              Input('start', 'n_clicks'),
              State('upload-data', 'filename'),
              State('upload-data', 'last_modified'))
def update_output(list_of_contents, QL, submit, list_of_names, list_of_dates):
    if list_of_contents is not None and submit > 0:
        children = [
            parse_contents(c, n) for c, n in
            zip(list_of_contents, list_of_names)]
        df = Compiler.process_file(children, QL)
        print(df)
        return dash_table.DataTable(
            data=df.to_dict('records'),
            id='table_output',
            columns=[{"name": str(i), "id": str(i)} for i in df.columns],
            fixed_rows={'headers': True},
            export_format="xlsx",
            style_header={'backgroundColor': 'rgb(30, 30, 30)'},
            style_cell={
                'whiteSpace': 'normal',
                'height': 'auto',
                'width': 'auto'
            },
        )


if __name__ == '__main__':
    app.run_server(debug=True)
