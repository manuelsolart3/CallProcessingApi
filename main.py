import os
import pandas as pd
from tkinter import Tk, filedialog, Label, Button, Entry, messagebox
import zipfile
from datetime import datetime

# Función para validar y parsear la fecha en diferentes formatos
def parse_date(date_str):
    date_formats = ['%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%Y%m%d'] 
    for fmt in date_formats:
        try:
            return datetime.strptime(date_str, fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    return None 

def select_excel_file():
    global excel_file_path
    excel_file_path = filedialog.askopenfilename(
        title="Selecciona el archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )
    if excel_file_path:
        label_excel_file.config(text=os.path.basename(excel_file_path))

def select_zip_file():
    global zip_file_path
    zip_file_path = filedialog.askopenfilename(
        title="Selecciona el archivo ZIP",
        filetypes=[("Archivos ZIP", "*.zip")]
    )
    if zip_file_path:
        label_zip_file.config(text=os.path.basename(zip_file_path))

def select_output_folder():
    global output_folder_path
    output_folder_path = filedialog.askdirectory(title="Selecciona la carpeta de salida")
    if output_folder_path:
        label_output_folder.config(text=os.path.basename(output_folder_path))

def process_files():
    if not excel_file_path or not zip_file_path or not output_folder_path:
        messagebox.showerror("Error", "Debes seleccionar el archivo Excel, el archivo ZIP y la carpeta de salida")
        return
    
    try:
        # Verificar que la ruta de salida sea válida
        if not os.path.exists(output_folder_path) or not os.path.isdir(output_folder_path):
            messagebox.showerror("Error", "La ruta de salida no es válida")
            return

        # Preguntar por la fecha
        date_str = date_entry.get()  
        if date_str:
            folder_date = parse_date(date_str)
            if not folder_date:
                messagebox.showerror("Error", "El formato de la fecha no es válido. Usa formatos como 'YYYY-MM-DD', 'DD/MM/YYYY', 'DD-MM-YYYY', o 'YYYYMMDD'")
                return
        else:
            folder_date = datetime.now().strftime('%Y-%m-%d')

        # Crear la carpeta con la fecha dentro de la ruta de salida
        output_folder = os.path.join(output_folder_path, folder_date)
        os.makedirs(output_folder, exist_ok=True)

        # Leer el archivo Excel
        excel_data = pd.read_excel(excel_file_path)
        excel_data.columns = excel_data.columns.str.strip().str.lower()

        # Verificar las columnas requeridas
        required_columns = ["agent", "customer id", "account name", "chat time"]
        if not all(column in excel_data.columns for column in required_columns):
            messagebox.showerror("Error", "El archivo Excel no contiene las columnas requeridas: 'Agent', 'Customer ID', 'Account Name', 'Chat Time'")
            return

        # Leer el archivo ZIP y extraer archivos HTML
        html_files = []
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            for file_info in zip_ref.infolist():
                if file_info.filename.endswith('.html'):
                    with zip_ref.open(file_info) as html_file:
                        html_content = html_file.read()
                        html_files.append(html_content)

        # Variables para contar los archivos procesados y los que no se pudieron procesar
        output_filenames = []
        unprocessed_count = 0

        # Generar y guardar los archivos de salida en la carpeta creada
        for index, row in excel_data.iterrows():
            agent = row["agent"] if pd.notna(row["agent"]) else "None"
            customer_id = row["customer id"] if pd.notna(row["customer id"]) else "None"
            account_name = row["account name"] if pd.notna(row["account name"]) else "None"
            chat_time = row["chat time"] if pd.notna(row["chat time"]) else "None"

            # Validación de 'Chat Time': Si es "00:00:00", no se procesa esa fila
            if chat_time == "00:00:00" or chat_time == "None":
                unprocessed_count += 1
                continue  # No crear archivo para esta fila

            # Generar el nombre del archivo
            output_filename = f"{agent}_{customer_id}_{account_name}.xlsx"
            output_filenames.append(output_filename)

            # Crear el archivo de salida
            output_filepath = os.path.join(output_folder, output_filename)
            output_df = pd.DataFrame()  # Aquí puedes añadir los datos correspondientes a la lógica de tu proyecto
            output_df.to_excel(output_filepath, index=False)

        # Mostrar alerta con la cantidad de archivos no procesados
        messagebox.showinfo("Éxito", f"Archivos generados correctamente.\nNo se pudieron procesar {unprocessed_count} archivos.")
        
        print("Archivos de salida generados:")
        for filename in output_filenames:
            print(filename)

    except Exception as e:
        messagebox.showerror("Error", f"Hubo un problema procesando los archivos: {str(e)}")

# Crear la interfaz gráfica usando Tkinter
root = Tk()
root.title("Carga de Archivos")

label_excel_file = Label(root, text="Selecciona el archivo Excel")
label_excel_file.pack()

btn_select_excel = Button(root, text="Seleccionar Excel", command=select_excel_file)
btn_select_excel.pack()

label_zip_file = Label(root, text="Selecciona el archivo ZIP")
label_zip_file.pack()

btn_select_zip = Button(root, text="Seleccionar ZIP", command=select_zip_file)
btn_select_zip.pack()

label_output_folder = Label(root, text="Selecciona la carpeta de salida")
label_output_folder.pack()

btn_select_output_folder = Button(root, text="Seleccionar Carpeta", command=select_output_folder)
btn_select_output_folder.pack()

Label(root, text="Fecha (opcional, formato: YYYY-MM-DD)").pack()

date_entry = Entry(root)
date_entry.pack()

btn_process_files = Button(root, text="Procesar Archivos", command=process_files)
btn_process_files.pack()

root.mainloop()
