import os
import pandas as pd
from flask import Flask, request
import zipfile
from datetime import datetime


app = Flask(__name__)

# Función para validar y parsear la fecha en diferentes formatos
def parse_date(date_str):
    date_formats = ['%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y%m%d'] 
    for fmt in date_formats:
        try:
            return datetime.strptime(date_str, fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    return None 

@app.route('/upload', methods=['POST'])
def upload_files():
    
    # Verificar que todos los archivos y la ruta están presentes
    if 'agent_chat_log_file' not in request.files or 'zip_file' not in request.files or 'saving_path' not in request.form:
        return "Debes proporcionar un archivo Excel, un archivo ZIP y una ruta de salida", 400

    # Obtener los archivos y la ruta de salida
    agent_chat_log_file = request.files['agent_chat_log_file']
    zip_file = request.files['zip_file']
    saving_path = request.form['saving_path']

    # Verificar que la ruta de salida exista y sea un directorio
    if not os.path.exists(saving_path) or not os.path.isdir(saving_path):
        return "La ruta de salida no es válida", 400

    #  crear la carpeta con la fecha
    date_str = request.form.get('date', None)  # El parámetro 'date' puede ser opcional
    if date_str:
        folder_date = parse_date(date_str)
        if not folder_date:
            return "El formato de la fecha no es válido. Usa formatos como 'YYYY-MM-DD', 'DD/MM/YYYY', 'DD-MM-YYYY', o 'YYYYMMDD'", 400
    else:
        folder_date = datetime.now().strftime('%Y-%m-%d') 

    # Crear la carpeta con la fecha dentro de la ruta de salida
    output_folder = os.path.join(saving_path, folder_date)
    os.makedirs(output_folder, exist_ok=True)


    if not agent_chat_log_file.filename.endswith(('.xlsx', '.xls')):
        return "El archivo proporcionado no es un archivo Excel válido", 400


    if not zip_file.filename.endswith('.zip'):
        return "El archivo proporcionado no es un archivo ZIP válido", 400

    try:
        excel_data = pd.read_excel(agent_chat_log_file.stream)
    except Exception as e:
        return f"Error al leer el archivo Excel: {str(e)}", 400


    excel_data.columns = excel_data.columns.str.strip().str.lower()


    required_columns = ["agent", "customer id", "account name"]
    if not all(column in excel_data.columns for column in required_columns):
        return "El archivo Excel no contiene las columnas requeridas: 'Agent', 'Customer ID', 'Account Name'", 400

    # paso siguiente...
    html_files = []
    with zipfile.ZipFile(zip_file.stream, 'r') as zip_ref:
        for file_info in zip_ref.infolist():
            if file_info.filename.endswith('.html'):
                with zip_ref.open(file_info) as html_file:
                    html_content = html_file.read()
                    html_files.append(html_content)  

    # Generar y guardar los archivos de salida en la carpeta creada
    output_filenames = []
    for index, row in excel_data.iterrows():
        agent = row["agent"] if pd.notna(row["agent"]) else "None"
        customer_id = row["customer id"] if pd.notna(row["customer id"]) else "None"
        account_name = row["account name"] if pd.notna(row["account name"]) else "None"


        output_filename = f"{agent}_{customer_id}_{account_name}.xlsx"

        if agent != "None" or customer_id != "None" or account_name != "None":
            output_filenames.append(output_filename)
            output_filepath = os.path.join(output_folder, output_filename)
            output_df = pd.DataFrame() 
            output_df.to_excel(output_filepath, index=False)

    print("Archivos de salida generados")
    for filename in output_filenames:
        print(filename)

    return f"", 200

# Iniciar la app
if __name__ == '__main__':
    app.run(debug=True)
