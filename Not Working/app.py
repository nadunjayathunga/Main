import dash
import dash_auth
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output
from dash import html, dcc
from data import company_info, db_info,USER_MAPPING
from datetime import datetime,date
from sqlalchemy import create_engine
import pandas as pd

app = dash.Dash(name=__name__,  external_stylesheets=[
                dbc.themes.PULSE], use_pages=True,title='Dashboard')

dash_auth.BasicAuth(app, USER_MAPPING)

header_row = dbc.Row(
    children=[
        html.H1('Corporate Dashboard', className='text-center text-primary')
    ],
    id='main-heading'
)

secondory_row = dbc.Row(
    children=[
        dbc.Col(
            [
                dcc.Dropdown(
                    options=[
                        {'label': i['data']['long_name'], 'value': i['data']['database']} for i in company_info
                    ],
                    value='elite_security',
                    id='company-name',
                    className='mt-1',
                    optionHeight=35,
                    clearable=False
                )
            ], width={'size': 2}

        ),
        dbc.Col(
            [
                dcc.DatePickerRange(
                    id='dt-pkr-range',
                    min_date_allowed=None,
                    max_date_allowed=None,
                    updatemode='bothdates',
                    start_date=None,
                    end_date=None,  # dt(2023, 8, 31)
                )
            ], width={'size': 3},
            style={'margin-top': 2},
        ),
        dbc.Tooltip('The date range need to be first day and last day of any given period',
                    target='dt-pkr-range',
                    placement='top'),
        dbc.Col(
            [
                dbc.Nav(
                    children=[], vertical=False, pills=True, justified=True, id='menu-items'
                )
            ], width={'size': 7},
        )
    ]
)

def check_date_format(date_str):
    try:
        return datetime.strptime(date_str, "%Y-%m-%dT%H:%M:%S").date()
    except ValueError:
        try:
            return datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            raise ValueError("Invalid date format")

app.layout = html.Div(children=[
    dcc.Store(id='start-date', data={}),
    dcc.Store(id='end-date', data={}),
    dcc.Store(id='database', data={}),
    header_row,
    secondory_row,
    html.Hr(),
    dash.page_container])


@app.callback(
    Output(component_id='menu-items', component_property='children'),
    [Input(component_id='company-name', component_property='value')]
)
def create_menu_item(item):
    menu_items = [i['data']['nav_links'] for i in company_info if i['data']
                  ['database'] == item]  # 'nav_links': ['Finance','Operations','Sales']
    nav_links = []
    for item in menu_items[0]:  # Finance
        menu_item = dbc.NavLink(item.upper(),
                                href='/' if item == 'Finance' else f'/{item.lower()}',
                                active='exact', disabled=False)
        nav_links.append(menu_item)
    return nav_links


@app.callback(
    [Output(component_id='dt-pkr-range', component_property='min_date_allowed'),
     Output(component_id='dt-pkr-range',
            component_property='max_date_allowed'),
     Output(component_id='dt-pkr-range', component_property='end_date'),
     Output(component_id='dt-pkr-range', component_property='start_date')],
    [Input(component_id='company-name', component_property='value'),]
)
def set_dates(company_db):
    engine = create_engine(
        f'postgresql://{db_info["USERNAME"]}:{db_info["PWD"]}@{db_info["HOSTNAME"]}:{db_info["PORT_ID"]}/{company_db}')
    query = 'SELECT voucher_date FROM "fGL"'
    df_fGl = pd.read_sql_query(query, engine)

    earliest_date = df_fGl['voucher_date'].min()
    closest_date = df_fGl['voucher_date'].max()
    current_year :str = str(date.today().year)
    cy_start_date = current_year+'-01-01'
    cy_start_date = datetime.strptime(cy_start_date,'%Y-%m-%d').date()

    return [earliest_date,
            closest_date,
            closest_date,
            cy_start_date]


@app.callback(
    [
        Output(component_id='start-date', component_property='data'),
        Output(component_id='end-date', component_property='data'),
        Output(component_id='database', component_property='data')
    ],
    [
        Input(component_id='dt-pkr-range', component_property='start_date'),
        Input(component_id='dt-pkr-range', component_property='end_date'),
        Input(component_id='company-name', component_property='value')
    ],prevent_initial_call = True
)
def output_data(start_date, end_date, database):
    start_date = check_date_format(start_date)
    end_date = check_date_format(end_date)

    return [start_date, end_date, database]


if __name__ == '__main__':
    app.run(debug=True, port=3000)
