from tkinter import Tk, Label, Button, filedialog, messagebox
import pandas as pd
from docx import Document
import os
from datetime import datetime


def generar_contratos():
    if not plantilla_path or not excel_path or not output_folder:
        messagebox.showwarning(
            "Faltan archivos", "Debes seleccionar la plantilla, el archivo Excel y la carpeta de destino.")
        return

    try:
        df = pd.read_excel(excel_path)
        df.columns = df.columns.str.lower()

        # Crear la subcarpeta "Contratos Generados" dentro de la carpeta de destino seleccionada
        contratos_folder = os.path.join(output_folder, "Contratos Generados")
        os.makedirs(contratos_folder, exist_ok=True)

        log_path = os.path.join(contratos_folder, "log_contratos.txt")
        total_generados = 0

        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write(
                f"\n--- Contratos generados el {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n")

            for _, fila in df.iterrows():
                doc = Document(plantilla_path)
                for parrafo in doc.paragraphs:
                    for clave, valor in fila.items():
                        if pd.notna(valor):
                            parrafo.text = parrafo.text.replace(
                                f"{{{{{clave}}}}}", str(valor))

                nombre_base = str(
                    fila.get("nombre", "desconocido")).strip().replace(" ", "_")
                nombre_archivo = f"Contrato_{nombre_base}.docx"
                ruta_archivo = os.path.join(contratos_folder, nombre_archivo)
                doc.save(ruta_archivo)

                log_file.write(f"✅ {nombre_archivo} generado correctamente.\n")
                total_generados += 1

        messagebox.showinfo(
            "Éxito", f"Se generaron {total_generados} contratos en:\n{contratos_folder}")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")


def seleccionar_plantilla():
    global plantilla_path
    plantilla_path = filedialog.askopenfilename(
        filetypes=[("Documentos Word", "*.docx")])
    label_plantilla.config(
        text=f"Plantilla: {os.path.basename(plantilla_path)}")


def seleccionar_excel():
    global excel_path
    excel_path = filedialog.askopenfilename(
        filetypes=[("Archivos Excel", "*.xlsx")])
    label_excel.config(text=f"Excel: {os.path.basename(excel_path)}")


def seleccionar_carpeta_destino():
    global output_folder
    output_folder = filedialog.askdirectory()
    if output_folder:
        label_carpeta.config(text=f"Carpeta de destino: {output_folder}")
    else:
        label_carpeta.config(text="Carpeta de destino: No seleccionada")


# Inicialización
plantilla_path = ""
excel_path = ""
output_folder = ""

ventana = Tk()
ventana.title("Generador de Contratos")
ventana.geometry("400x300")

Label(ventana, text="Generador de Contratos",
      font=("Helvetica", 16)).pack(pady=10)
Button(ventana, text="Seleccionar plantilla Word",
       command=seleccionar_plantilla).pack(pady=5)
label_plantilla = Label(ventana, text="Plantilla: No seleccionada")
label_plantilla.pack()

Button(ventana, text="Seleccionar archivo Excel",
       command=seleccionar_excel).pack(pady=5)
label_excel = Label(ventana, text="Excel: No seleccionado")
label_excel.pack()

Button(ventana, text="Seleccionar carpeta de destino",
       command=seleccionar_carpeta_destino).pack(pady=5)
label_carpeta = Label(ventana, text="Carpeta de destino: No seleccionada")
label_carpeta.pack()

Button(ventana, text="Generar contratos", command=generar_contratos,
       bg="#4CAF50", fg="white").pack(pady=20)

ventana.mainloop()
