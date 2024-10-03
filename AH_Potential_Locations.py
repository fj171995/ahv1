from flask import Flask, request, redirect, url_for, render_template_string, send_file
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime
from xhtml2pdf import pisa

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
                max-width: 90%;
                width: 400px;
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
        wb = load_workbook(filename=BytesIO(file.read()), data_only=True)
        global uploaded_wb
        uploaded_wb = wb
        return redirect(url_for('display_data'))

@app.route('/data', methods=['GET', 'POST'])
def display_data():
    if 'uploaded_wb' not in globals():
        return redirect(url_for('upload_file'))

    wb = uploaded_wb
    sheet = wb.active

    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        location = row[9]  # Columna J
        google_map = f'<a href="{row[14]}" target="_blank">Link</a>' if row[14] else 'Missing'  # Columna O
        pictures = f'<a href="{row[15]}" target="_blank">Link</a>' if row[15] else 'Missing'  # Columna P
        real_estate_ad = f'<a href="{row[16]}" target="_blank">Link</a>' if row[16] else 'Missing'  # Columna Q
        net_rent_month = row[21]  # Columna V
        total_sqm = row[27]  # Columna AB
        estimated_parking_spots = row[41]  # Columna AP
        rent_per_sqm = row[39]  # Columna AN
        cost_per_parking_spot = row[1]  # Columna B
        flexicar = f'<a href="{row[3]}" target="_blank">Link</a>' if row[3] else 'Missing'  # Columna D
        ocasionplus = f'<a href="{row[4]}" target="_blank">Link</a>' if row[4] else 'Missing'  # Columna E
        ctc = f'<a href="{row[5]}" target="_blank">Link</a>' if row[5] else 'Missing'  # Columna F
        comments = row[40] if row[40] else 'No comments available'  # Columna AO

        data.append((location, google_map, pictures, real_estate_ad, net_rent_month,
                     total_sqm, estimated_parking_spots, rent_per_sqm, cost_per_parking_spot,
                     flexicar, ocasionplus, ctc, comments))

    location_filter = request.args.get('location_filter', '')
    if location_filter:
        data = [row for row in data if row[0] and location_filter.lower() in row[0].lower()]

    show_generate_report = request.args.get('show_generate_report', 'true').lower() == 'true'

    html_table = '''
    <style>
        body {
            font-family: 'Roboto', sans-serif;
            background-color: #f4f4f4;
            padding: 20px;
        }
        .container {
            max-width: 100%;
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
            overflow-x: auto;
            display: block;
        }
        th, td {
            padding: 15px;
            text-align: center;
            border-bottom: 1px solid #ddd;
            font-size: 0.9em;
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
    </style>
    <div class="container">
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
                <td>+info</td>
            </tr>
        '''

    html_table += '''
        </table>
    '''

    if show_generate_report:
        html_table += '''
        <button class="generate-report-button" onclick="window.location.href='/generate_pdf?location_filter=' + document.getElementById('location_filter').value">Generate Report as PDF</button>
        '''

    html_table += '''
    </div>
    '''

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

@app.route('/generate_pdf')
def generate_pdf():
    location_filter = request.args.get('location_filter', '')
    html_content = display_data().data.decode('utf-8')

    pdf_file = BytesIO()
    pisa_status = pisa.CreatePDF(BytesIO(html_content.encode('utf-8')), dest=pdf_file)

    if pisa_status.err:
        return 'Error generating PDF', 500

    pdf_file.seek(0)
    return send_file(pdf_file, mimetype='application/pdf', as_attachment=True, download_name='report.pdf')

# Configuración para Vercel
app = app

if __name__ == '__main__':
    app.run()
