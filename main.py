import os
import pandas as pd
from tkinter import Tk, filedialog, Label, Button, Entry, messagebox, Frame, StringVar
from tkinter.ttk import Style
from datetime import datetime

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
        excel_file_var.set(os.path.basename(excel_file_path))

def select_zip_file():
    global zip_file_path
    zip_file_path = filedialog.askopenfilename(
        title="Selecciona el archivo ZIP",
        filetypes=[("Archivos ZIP", "*.zip")]
    )
    if zip_file_path:
        zip_file_var.set(os.path.basename(zip_file_path))

def select_output_folder():
    global output_folder_path
    output_folder_path = filedialog.askdirectory(title="Selecciona la carpeta de salida")
    if output_folder_path:
        output_folder_var.set(os.path.basename(output_folder_path))

def process_files():
    if not excel_file_path or not zip_file_path or not output_folder_path:
        messagebox.showerror("Error", "Debes seleccionar el archivo Excel, el archivo ZIP y la carpeta de salida")
        return
    
    try:
        if not os.path.exists(output_folder_path) or not os.path.isdir(output_folder_path):
            messagebox.showerror("Error", "La ruta de salida no es válida")
            return

        date_str = date_entry.get()
        if date_str:
            folder_date = parse_date(date_str)
            if not folder_date:
                messagebox.showerror("Error", "El formato de la fecha no es válido. Usa formatos como 'YYYY-MM-DD', 'DD/MM/YYYY', 'DD-MM-YYYY', o 'YYYYMMDD'")
                return
        else:
            folder_date = datetime.now().strftime('%Y-%m-%d')

        output_folder = os.path.join(output_folder_path, folder_date)
        os.makedirs(output_folder, exist_ok=True)

        excel_data = pd.read_excel(excel_file_path)
        excel_data.columns = excel_data.columns.str.strip().str.lower()

        required_columns = ["agent", "customer id", "account name", "chat time"]
        if not all(column in excel_data.columns for column in required_columns):
            messagebox.showerror("Error", "El archivo Excel no contiene las columnas requeridas: 'Agent', 'Customer ID', 'Account Name', 'Chat Time'")
            return

        processed_count = 0
        zero_time_count = 0

        for index, row in excel_data.iterrows():
            agent = str(row["agent"]) if pd.notna(row["agent"]) else "[None]"
            customer_id = str(row["customer id"]) if pd.notna(row["customer id"]) else "[None]"
            account_name = str(row["account name"]) if pd.notna(row["account name"]) else "[None]"
            chat_time = str(row["chat time"]) if pd.notna(row["chat time"]) else "[None]"

            if agent == "[None]" and customer_id == "[None]" and account_name == "[None]":
                continue

            if chat_time == "00:00:00" or chat_time == "[None]":
                zero_time_count += 1
                continue

            output_filename = f"{agent}_{customer_id}_{account_name}.xlsx"
            output_filepath = os.path.join(output_folder, output_filename)

            output_df = pd.DataFrame({'Data': ['Processed data for this call would go here']})
            output_df.to_excel(output_filepath, index=False)
            processed_count += 1

        message = f"Procesamiento completado.\n\n"
        message += f"Archivos procesados: {processed_count}\n"
        message += f"Archivos no procesado con Chat Time igual a 00:00:00 o nulo: {zero_time_count}"
        
        messagebox.showinfo("Éxito", message)

    except Exception as e:
        messagebox.showerror("Error", f"Hubo un problema procesando los archivos: {str(e)}")

# GUI setup
root = Tk()
root.title("Procesador de Archivos de Llamadas")
root.geometry("500x400")  # Ajustar tamaño de la ventana

style = Style()
style.theme_use('clam')

main_frame = Frame(root, padx=10, pady=10)  # Reducir márgenes
main_frame.pack(fill='both', expand=True)

excel_file_var = StringVar()
zip_file_var = StringVar()
output_folder_var = StringVar()

Label(main_frame, text="Archivo Excel:").pack(anchor='w')
Label(main_frame, textvariable=excel_file_var, width=22, relief="sunken", padx=2).pack(fill='x', pady=(0, 2))  # Reducir el ancho
Button(main_frame, text="Seleccionar Excel", command=select_excel_file).pack(pady=(0, 2))  # Reducir el espacio entre botones

Label(main_frame, text="Archivo ZIP:").pack(anchor='w')
Label(main_frame, textvariable=zip_file_var, width=22, relief="sunken", padx=2).pack(fill='x', pady=(0, 2))
Button(main_frame, text="Seleccionar ZIP", command=select_zip_file).pack(pady=(0, 2))

Label(main_frame, text="Carpeta de Salida:").pack(anchor='w')
Label(main_frame, textvariable=output_folder_var, width=22, relief="sunken", padx=2).pack(fill='x', pady=(0, 2))
Button(main_frame, text="Seleccionar Carpeta", command=select_output_folder).pack(pady=(0, 2))

Label(main_frame, text="Fecha (opcional, formato: YYYY-MM-DD):").pack(anchor='w')
date_entry = Entry(main_frame)
date_entry.pack(fill='x', pady=(0, 2))

Button(main_frame, text="Procesar Archivos", command=process_files, bg='#4CAF20', fg='white', padx=2, pady=3).pack(pady=10)  # Reducir padding

root.mainloop()
