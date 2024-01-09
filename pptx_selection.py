import dash
from dash import dcc
from dash import html
from dash import html, dcc
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.graph_objects as go
import time
import dash
from dash import html, dcc, callback, Output, Input, State
import plotly.express as px
import pandas as pd
from dash.exceptions import PreventUpdate
import os
import base64
from pptx import Presentation 
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import plotly.graph_objects as go
import plotly.io as pio
from pptx import Presentation
from pptx.util import Inches
from PIL import Image
import io
import os
import os
# Create a Dash web application
app = dash.Dash(__name__)
data = px.data.iris()
fig = px.scatter(data, x='sepal_width', y='sepal_length', color='species', title='Scatter Plot')

# Define the layout of the app
app.layout = html.Div(
    children=[
        html.H1("test App"),
        dbc.Row([
            dbc.Col([
                html.Img(src='assets/template1.png', alt='Image'),
                dbc.Button("Choose this template", id="template"),
                html.Div(id='notice', children=[]),
                dcc.Graph(
                    id='example-graph',
                    figure=fig
                ),
                dbc.Button('Add to pptx', id='add_btn'),
            ])
        ])
    ]
)

    
if __name__ == '__main__':
    app.run_server(debug=True)
