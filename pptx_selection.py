import dash
from dash import html, dcc, Input, Output
import dash_bootstrap_components as dbc
import plotly.express as px
import pandas as pd
import plotly.graph_objects as go
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# Create a Dash web application
app = dash.Dash(__name__)
data = px.data.iris()
fig = px.scatter(data, x='sepal_width', y='sepal_length',
                 color='species', title='Scatter Plot')

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
                html.Div(id='output')
            ])
        ])
    ]
)


@app.callback(
    Output('output', 'children'),
    Input('add_btn', 'n_clicks'),
)
def toggle(n_clicks):
    filename = "test2.pptx"

    pr = Presentation("template1.pptx")

    slides = pr.slides
    slides._sldIdLst.clear()
    slide_layout = pr.slide_layouts[5]
    slide = pr.slides.add_slide(slide_layout)
    image_path = "newplot.png"
    left = Inches(.1)
    top = Inches(.1)
    # width = Inches(13)
    height = Inches(7)
    # max height and width
    slide.shapes.add_picture(image_path, left, top, None, height)

    pr.save(filename)
    return html.Div("printed")


if __name__ == '__main__':
    app.run_server(debug=True)
