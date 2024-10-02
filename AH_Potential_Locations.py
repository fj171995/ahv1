from flask import Flask, request, redirect, url_for, render_template_string, session
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'secret_key'  # Necesario para sesiones

@app.route('/')
def home():
    if 'uploaded_wb' not in session:
        return redirect(url_for('upload_file'))

    wb = load_workbook(filename=BytesIO(session['uploaded_wb']), data_only=True)
    sheet = wb.active

    # Leer los datos de la hoja de cálculo
    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        data.append(row)

    # Construcción de HTML para la tabla con los hipervínculos y el botón "+info"
    html_table = '''
    <style>
        body {
            font-family: 'Roboto', sans-serif;
            background-color: #f4f4f4;
            padding: 20px;
        }
        .container {
            max-width: 1200px;
            margin: auto;
            background: #ffffff;
            padding: 20px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
            border-radius: 10px;
        }
        h1, h2 {
            color: #333;
            font-weight: 700;
            text-align: center;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }
        th, td {
            padding: 15px;
            text-align: center;
            border-bottom: 1px solid #ddd;
        }
        th {
            background-color: #ff6200;
            color: white;
        }
        tr:hover {
            background-color: #f1f1f1;
        }
        a {
            color: #ff6200;
            text-decoration: none;
            font-weight: bold;
        }
        a:hover {
            text-decoration: underline;
        }
        .info-button {
            cursor: pointer;
            color: #007bff;
            text-decoration: underline;
            background: none;
            border: none;
            font-size: 1em;
            font-weight: bold;
        }
    </style>
    <div class="container">
        <h1>Report AH Potential Locations</h1>
        <h2>Fecha: ''' + datetime.now().strftime("%d/%m/%Y") + '''</h2>
        <table>
            <tr>
                <th>Location</th>
                <th>Google map</th>
                <th>Pictures</th>
                <th>Real estate ad</th>
                <th>Net rent / month</th>
                <th>Total sqm outdoor + indoor</th>
                <th>Total parking spots</th>
                <th>Rent/sqm</th>
                <th>€/parking spot</th>
                <th>+info</th>
            </tr>
    '''

    for row in data:
        html_table += f'''
            <tr>
                <td>{row[0]}</td>
                <td><a href="{row[1]}" target="_blank">Link</a></td>
                <td><a href="{row[2]}" target="_blank">Link</a></td>
                <td><a href="{row[3]}" target="_blank">Link</a></td>
                <td>{row[4]}</td>
                <td>{row[5]}</td>
                <td>{row[6]}</td>
                <td>{row[7]}</td>
                <td>{row[8]}</td>
                <td>
                    <button class="info-button" onclick="alert('Flexicar: {row[9]}\\nOcasionPlus: {row[10]}\\nCTC: {row[11]}')">
                        + Info
                    </button>
                </td>
            </tr>
        '''

    html_table += '''
        </table>
        <form action="/new_report" method="get">
            <input type="submit" value="New Report" style="background-color: #ff6200; color: #fff; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer;">
        </form>
    </div>
    '''

    return render_template_string(html_table)

@app.route('/new_report', methods=['GET', 'POST'])
def new_report():
    if request.method == 'POST':
        password = request.form['password']
        if password == 'desde54':
            return redirect(url_for('upload_file'))
        else:
            return 'Contraseña incorrecta, intenta nuevamente.'

    return '''
    <form method="post">
        <label for="password">Contraseña:</label>
        <input type="password" name="password" required>
        <input type="submit" value="Submit">
    </form>
    '''

@app.route('/upload')
def upload_file():
    return render_template_string('''
    <!doctype html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Subir archivo Excel - AH Potential Locations</title>
        <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap" rel="stylesheet">
        <style>
            body {
                font-family: 'Roboto', sans-serif;
                background-color: #f4f4f4;
                margin: 0;
                padding: 0;
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
            }
            .container {
                background: #ffffff;
                padding: 30px;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
                border-radius: 10px;
                text-align: center;
            }
            h1 {
                color: #333;
                font-weight: 700;
                margin-bottom: 20px;
            }
            form {
                margin-top: 20px;
            }
            input[type=file] {
                margin-bottom: 20px;
            }
            input[type=submit] {
                background-color: #ff6200;
                color: #ffffff;
                padding: 10px 20px;
                border: none;
                border-radius: 5px;
                cursor: pointer;
                font-weight: bold;
            }
            input[type=submit]:hover {
                background-color: #e55b00;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Sube el archivo Excel</h1>
            <form method="post" enctype="multipart/form-data" action="/upload_excel">
                <input type="file" name="file" required>
                <br>
                <input type="submit" value="Subir archivo">
            </form>
        </div>
    </body>
    </html>
    ''')

@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    if file:
        # Cargar el archivo Excel en memoria
        session['uploaded_wb'] = file.read()
        return redirect(url_for('home'))

if __name__ == '__main__':
    app.run()
