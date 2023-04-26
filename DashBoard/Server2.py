import dash

import plotly.express as px
import pandas as pd
import xlwings as xl
import datetime
from dash import html, dcc
from dash.dependencies import Input, Output
import numpy as np


WB = xl.Book('D:/PyProject2023/DashBoard/DashBoardDataTest.xlsb')
SH1 = WB.sheets['Sample1']
# Prepare sample data


app = dash.Dash(__name__)

app.layout = html.Div([
    html.H1("Real-time Dash App with Graph"),
    html.P("This Dash app updates a line chart in real-time."),
    dcc.Graph(id='realtime-graph', animate=True),
    dcc.Interval(id='interval-component', interval=3000, n_intervals=0)  # Update every 1000ms (1 second)
])

time_start = datetime.datetime.now()

@app.callback(
    Output('realtime-graph', 'figure'),
    Input('interval-component', 'n_intervals')
)
def update_graph(n_intervals):
    current_time = datetime.datetime.now() - time_start
    current_time_seconds = current_time.total_seconds()

    # Create a sine wave with a frequency of 1 Hz and an amplitude of 1
    # table1 = SH1.range('b5').expand('table').value
    # df = pd.DataFrame(table1, columns=SH1.range('b4:c4').value)
    # print(df)
    x = SH1.range('b5:b12').value
    y = SH1.range('c5:c12').value
    fig = px.line(x, y)

    # fig.update_layout(width=800, height=600)
    return fig

if __name__ == '__main__':
    app.run_server(debug=True)
