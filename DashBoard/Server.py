import dash
import dash_core_components as dcc
import dash_html_components as html
import plotly.express as px
import pandas as pd
import xlwings as xl

from dash.dependencies import Event, Input, Output

WB = xl.Book('/Users/kwangryeollee/PyProject/DashBoard/DashBoardDataTest.xlsb')
SH1 = WB.sheets['Sample1']
# Prepare sample data

table1 = SH1.range('b5').expand('table').value
df = pd.DataFrame(table1, columns=SH1.range('b4:c4').value)
data = df.values

print(data)

print(df)

# Create a bar chart using Plotly Express
# Express
# wide_df = px.data.medals_wide()
# print('wide table:', wide_df)

fig = px.line(df, df.x, df.y)

# Initialize Dash app
app = dash.Dash(__name__)

# Define app layout
app.layout = html.Div(
    dcc.Interval(id='refresh', interval=200),
    html.Div(id='content', className="container"),

    children=[
        html.H1("Lee's First Dash Board"),
        dcc.Graph(figure=fig, id="bar-chart"),
    ] )



# Run the Dash app
if __name__ == "__main__":
    app.run_server(debug=True)
