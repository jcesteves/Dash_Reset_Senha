import dash
import dash_core_components as dcc
import dash_html_components as html
import pandas as pd
from flask import send_file
import io

# Carrega o arquivo com os dados do reset de senha
df = pd.read_excel('reset.xlsx')
df['senha_expira'] = df['senha_expira'].dt.date


#Função pra efetuar o calculo dataa expirar - data atual
def update_data():
   df['expira_em'] = df['senha_expira'] - pd.datetime.now().date()
   df['expira_em'] = df['expira_em'].dt.days.astype('int16')
   return df['expira_em']



# Converte e efetua o cálculo entre o prazo a expirar e subtrai com a data atual

# Arquivo externo de css do gráfico
external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

# Instancia o abjeto da aplicação
app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

# Design e plotagem do gráfico
app.layout = html.Div(className='div', style={'backgroundColor': 'blue'}, children=[
    # Este comando é o responsável pelo refresh na página
    dcc.Interval(
        interval=1 * 25200000,
        n_intervals=0
    ),
    html.Nav(
        className='nav-dash',
        children=[html.Img(className='nav-img', src='/assets/logo.png'),
                  html.A(html.Button('Base Excel!', className='button'),
                         href='/download_excel/'),
                  html.A(html.Button('Manual!', className='button1'),
                         href='/dash/urlToDownload'),
                  html.P(className='nav-title', children='CONSULTA DE SENHAS A EXPIRAR'),

                  ]),

    html.P(className='p1', children='Matrículas "Menores que 0", já estão com as senhas expiradas - \
           Matrículas "Igual a 0", expiram hoje - Demais Matrículas a expirar'),

# Código do gráfico a expirar global
    dcc.Graph(
        id='update_data',
        figure={
            'data': [
                {
                    'x': update_data(),
                    'y': df.home_office,
                    'text': df.status+ ' - ' + df.matricula + ' - ' + 'Home-Office',
                    'name': 'Teste1',
                    'mode': 'markers',
                    'marker': {'size': 11,
                               'color': '#4682B4'
                               },

                },

            ],
            'layout': {
                'clickmode': 'event+select',
                'height': 550,
                'xaxis': {'title': 'DIAS A EXPIRAR POR USUARIO HOME-OFFICE'},
                'grid': 'grid'


            }
        }
    ),

    html.P(className='linha'),
# Código do gráfico total a expirar por dia
    dcc.Graph(
        style={'backgroundColor': 'blue'},
        id='total_user',
        figure={
            'data': [
                {
                    'x': df.groupby(["senha_expira"])['home_office'].count().index,
                    'y': df.groupby(["expira_em", "senha_expira"])['home_office'].count().values,
                    'name': 'Trace 1',
                    'mode': 'markers',
                    'marker': {'size': 8,
                               'color': '#4682B4'
                               },
                    'type': 'bar',




                },

            ],
            'layout': {
                'clickmode': 'event+select',
                'height': 550,
                'xaxis': {'title': 'TOTAL DE USUARIOS HOME-OFFICE POR DIA'},
                'grid': 'grid'
            }
        }
    ),



    dcc.Graph(
            id='example-graph',
            figure={
                'data': [
                    {'x': df.home_office.value_counts(), 'y':  df.home_office.value_counts(), 'type': 'bar', 'name': 'Fixo'},
                    {'x': df.home_office.value_counts(), 'y': df.home_office.value_counts(), 'type': 'bar', 'name': 'Home-Office'},
                ],
                'layout': {
                    'title': 'Total de usuários fixo e home-office'
                }
            }
    ),














])


# Função para gerar relatorio em excel com base no dataframe do reset.xlsx
@app.server.route('/download_excel/')
def download_excel():
    # Converte o Data Frame
    strIO = io.BytesIO()
    excel_writer = pd.ExcelWriter(strIO, engine="xlsxwriter")
    df.to_excel(excel_writer, sheet_name="sheet1")
    excel_writer.save()
    excel_data = strIO.getvalue()
    strIO.seek(0)

    return send_file(strIO,
                     attachment_filename='Base_Reset.xlsx',
                     as_attachment=True)


@app.server.route('/dash/urlToDownload')
def download_manual():
    return send_file('Orientação para visualização e entendimento de senhas expiradas.docx',
                     mimetype='text/csv',
                     attachment_filename='manual.docx',
                     as_attachment=True)

if __name__ == '__main__':
    app.run_server(debug=True)
