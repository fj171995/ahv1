from flask import Flask, request, redirect, url_for, render_template_string
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime

app = Flask(__name__)

@app.route('/')
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
            <form method="post" enctype="multipart/form-data" action="/upload">
                <input type="file" name="file" required>
                <br>
                <input type="submit" value="Subir archivo">
            </form>
        </div>
    </body>
    </html>
    ''')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    if file:
        # Cargar el archivo Excel en memoria
        wb = load_workbook(filename=BytesIO(file.read()), data_only=True)
        # Guardar el workbook en una variable global para su uso posterior
        global uploaded_wb
        uploaded_wb = wb
        return redirect(url_for('display_data'))

@app.route('/data', methods=['GET', 'POST'])
def display_data():
    # Verificar si el archivo ha sido subido
    if 'uploaded_wb' not in globals():
        return redirect(url_for('upload_file'))

    wb = uploaded_wb
    sheet = wb.active

    # Leer los datos de la hoja de cálculo
    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):  # Asumiendo que la primera fila es el encabezado
        location = row[9]  # Columna J
        google_map = f'<a href="{row[14]}">Link</a>' if row[14] else 'Missing'  # Columna O
        pictures = f'<a href="{row[15]}">Link</a>' if row[15] else 'Missing'  # Columna P
        real_estate_ad = f'<a href="{row[16]}">Link</a>' if row[16] else 'Missing'  # Columna Q
        net_rent_month = row[21]  # Columna V
        total_sqm = row[27]  # Columna AB
        estimated_parking_spots = row[41]  # Columna AP
        rent_per_sqm = row[39]  # Columna AN
        cost_per_parking_spot = row[1]  # Columna B
        flexicar = f'<a href="{row[3]}">Link</a>' if row[3] else 'Missing'  # Columna D
        ocasionplus = f'<a href="{row[4]}">Link</a>' if row[4] else 'Missing'  # Columna E
        ctc = f'<a href="{row[5]}">Link</a>' if row[5] else 'Missing'  # Columna F
        comments = row[40] if row[40] else 'No comments available'  # Columna AO

        data.append((
            location, google_map, pictures, real_estate_ad,
            net_rent_month, total_sqm, estimated_parking_spots,
            rent_per_sqm, cost_per_parking_spot,
            flexicar, ocasionplus, ctc, comments
        ))

    # Aplicar filtro si se envía desde el formulario
    location_filter = request.args.get('location_filter', '')
    if location_filter:
        data = [row for row in data if row[0] and location_filter.lower() in row[0].lower()]

    # Determinar si se muestra el botón "Generate Report"
    show_generate_report = request.args.get('show_generate_report', 'true').lower() == 'true'

    # Construcción de HTML para mostrar la tabla con diseño responsive y filtro
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
        .filter-form {
            text-align: left;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 10px;
            background-color: #007bff;
            padding: 10px;
            border-radius: 5px;
            color: white;
        }
        .filter-form label {
            font-weight: bold;
        }
        .filter-form input[type=text] {
            padding: 5px;
            border: none;
            border-radius: 3px;
        }
        .filter-form input[type=submit] {
            background-color: #ff6200;
            color: #ffffff;
            padding: 5px 10px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-weight: bold;
        }
        .filter-form input[type=submit]:hover {
            background-color: #e55b00;
        }
        .expandable-row {
            display: none;
            background-color: #f9f9f9;
        }
        .expanded-content {
            padding: 15px;
            text-align: left;
            color: #333; /* Color oscuro para mejor visibilidad */
        }
        .clickable-info {
            color: #ff6200;
            font-weight: bold;
            cursor: pointer;
        }
        .generate-report-button {
            margin-top: 20px;
            padding: 10px 20px;
            background-color: #007bff;
            color: #ffffff;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-weight: bold;
        }
        .generate-report-button:hover {
            background-color: #0056b3;
        }
    </style>
    <script>
        function toggleRow(id) {
            var row = document.getElementById(id);
            if (row.style.display === "none" || row.style.display === "") {
                row.style.display = "table-row";
            } else {
                row.style.display = "none";
            }
        }
        
        function generateReport() {
            const urlParams = new URLSearchParams(window.location.search);
            urlParams.set('location_filter', document.getElementById('location_filter').value);
            urlParams.set('show_generate_report', 'false');
            const reportUrl = window.location.origin + window.location.pathname + '?' + urlParams.toString();
            prompt('Shareable Report URL:', reportUrl);
        }
    </script>
    <div class="container">
        <form method="get" class="filter-form">
            <label for="location_filter">Filter Location:</label>
            <input type="text" name="location_filter" id="location_filter" placeholder="Location">
            <input type="submit" value="Apply">
        </form>
        <table>
            <tr>
                <th>#</th>
                <th>Location</th>
                <th>Google map</th>
                <th>Pictures</th>
                <th>Real estate ad</th>
                <th>Net rent / month</th>
                <th>Total sqm (outdoor + indoor)</th>
                <th>Estimated # parking spots</th>
                <th>Rent/sqm</th>
                <th>€/parking spot</th>
                <th>+Info</th>
            </tr>
    '''

    for index, row in enumerate(data):
        expandable_row_id = f"expandable-row-{index}"
        html_table += f'''
            <tr>
                <td>{index + 1}</td>
                <td>{row[0]}</td>
                <td>{row[1]}</td>
                <td>{row[2]}</td>
                <td>{row[3]}</td>
                <td>{row[4]}</td>
                <td>{row[5]}</td>
                <td>{row[6]}</td>
                <td>{row[7]}</td>
                <td>{row[8]}</td>
                <td class="clickable-info" onclick="toggleRow('{expandable_row_id}')">+info</td>
            </tr>
            <tr id="{expandable_row_id}" class="expandable-row">
                <td colspan="11" class="expanded-content">
                    <p><strong>Flexicar:</strong> {row[9]}</p>
                    <p><strong>OcasionPlus:</strong> {row[10]}</p>
                    <p><strong>CTC:</strong> {row[11]}</p>
                    <p><strong>Comments:</strong> {row[12]}</p>
                </td>
            </tr>
        '''

    html_table += '''
        </table>
    '''

    if show_generate_report:
        html_table += '''
        <button class="generate-report-button" onclick="generateReport()">Generate Report</button>
        '''

    html_table += '''
    </div>
    '''

    # Obtención de la fecha actual en formato dd/mm/yyyy
    current_date = datetime.now().strftime("%d/%m/%Y")

    return render_template_string('''
    <!doctype html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Report AH Potential Locations</title>
        <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&display=swap" rel="stylesheet">
    </head>
    <body>
        <div class="container">
            <h1>Report AH Potential Locations</h1>
            <h2>Fecha: {{ date }}</h2>
            <div>{{ table|safe }}</div>
        </div>
    </body>
    </html>
    ''', table=html_table, date=current_date)

# Configuración para Vercel
app = app

if __name__ == '__main__':
    app.run()
