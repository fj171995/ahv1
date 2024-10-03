import os
from flask import Flask, request, redirect, url_for, render_template_string
import pandas as pd  # type: ignore
from datetime import datetime

app = Flask(__name__)

# Ruta principal redirigida a la página de carga del archivo
@app.route('/')
def index():
    return redirect(url_for('upload_file'))

@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file part'
        file = request.files['file']
        if file.filename == '':
            return 'No selected file'
        if file:
            # Leer el archivo Excel directamente en Pandas
            df = pd.read_excel(file, engine='openpyxl')
            # Guardar el DataFrame en una variable global para su uso en la página de reporte
            global df_selected
            df_selected = df.copy()
            return redirect(url_for('report'))
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
            <form method="post" enctype="multipart/form-data">
                <input type="file" name="file" required>
                <br>
                <input type="submit" value="Subir archivo">
            </form>
        </div>
    </body>
    </html>
    ''')

@app.route('/report', methods=['GET'])
def report():
    # Verificar si el dataframe existe
    if 'df_selected' not in globals():
        return redirect(url_for('upload_file'))

    # Seleccionar las columnas requeridas
    columns = [
        'LOCATION', 'Google map', 'Pictures', 'Real Estate ad', 'net rent / month', 
        'TOTAL SQM OUTDOOR + INDOOR', 'Estimated # parking spots outdoor', 
        '# parking spaces for showroom', 'Rent/sqm', '€/parking spot', 
        'Flexicar Around? Insert in comments which one and driving time', 
        'OcasionPlus Around? Insert in comments which one and driving time', 
        'CTC >10 min?'
    ]
    df = df_selected[columns]

    # Crear nueva columna para el total de plazas de aparcamiento
    df['Total Parking spots'] = df['Estimated # parking spots outdoor'] + df['# parking spaces for showroom']

    # Construir HTML para la tabla
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
        .details {
            display: none;
            padding-top: 10px;
        }
        .toggle-button {
            cursor: pointer;
            color: #ff6200;
            font-weight: bold;
            text-decoration: underline;
        }
        @media screen and (max-width: 768px) {
            table, th, td {
                font-size: 14px;
            }
        }
    </style>
    <div class="container">
        <h1>Report AH Potential Locations</h1>
        <h2>Fecha: ''' + datetime.now().strftime("%d/%m/%Y") + '''</h2>
        <table>
            <tr>
                <th>#</th>
                <th>Location</th>
                <th>Google map</th>
                <th>Pictures</th>
                <th>Real estate ad</th>
                <th>Net rent / month</th>
                <th>Total sqm outdoor + indoor</th>
                <th>Total parking spots</th>
                <th>Rent/sqm</th>
                <th>€/parking spot</th>
                <th style="width: 15%;">+ info</th>
            </tr>
    '''

    for index, row in df.iterrows():
        row_number = index + 1
        google_map_link = f'<a href="{row["Google map"]}" target="_blank">Link</a>' if pd.notna(row["Google map"]) else 'N/A'
        pictures_link = f'<a href="{row["Pictures"]}" target="_blank">Link</a>' if pd.notna(row["Pictures"]) else 'N/A'
        real_estate_link = f'<a href="{row["Real Estate ad"]}" target="_blank">Link</a>' if pd.notna(row["Real Estate ad"]) else 'N/A'

        flexicar_link = f'<a href="{row["Flexicar Around? Insert in comments which one and driving time"]}" target="_blank">Link</a>' if pd.notna(row["Flexicar Around? Insert in comments which one and driving time"]) else 'N/A'
        ocasionplus_link = f'<a href="{row["OcasionPlus Around? Insert in comments which one and driving time"]}" target="_blank">Link</a>' if pd.notna(row["OcasionPlus Around? Insert in comments which one and driving time"]) else 'N/A'
        ctc_info = row['CTC >10 min?'] if pd.notna(row['CTC >10 min?']) else 'N/A'

        html_table += f'''
            <tr>
                <td>{row_number}</td>
                <td>{row['LOCATION']}</td>
                <td>{google_map_link}</td>
                <td>{pictures_link}</td>
                <td>{real_estate_link}</td>
                <td>{row['net rent / month']}</td>
                <td>{row['TOTAL SQM OUTDOOR + INDOOR']}</td>
                <td>{row['Total Parking spots']}</td>
                <td>{row['Rent/sqm']}</td>
                <td>{row['€/parking spot']}</td>
                <td>
                    <span class="toggle-button" onclick="toggleDetails({index})">+ Info</span>
                    <div id="details-{index}" class="details">
                        <p><strong>Flexicar:</strong> {flexicar_link}</p>
                        <p><strong>OcasionPlus:</strong> {ocasionplus_link}</p>
                        <p><strong>CTC >10 min?</strong> {ctc_info}</p>
                    </div>
                </td>
            </tr>
        '''

    html_table += '''
        </table>
    </div>
    <script>
        function toggleDetails(index) {
            var details = document.getElementById('details-' + index);
            if (details.style.display === 'none' || details.style.display === '') {
                details.style.display = 'block';
            } else {
                details.style.display = 'none';
            }
        }
    </script>
    '''

    return render_template_string(html_table)

if __name__ == '__main__':
    app.run(debug=True)

