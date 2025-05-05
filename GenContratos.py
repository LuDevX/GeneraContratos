from tkinter import Tk, Label, Button, filedialog, messagebox
import pandas as pd
from docx import Document
import os
from datetime import datetime


def generar_contratos():
    if not plantilla_path or not excel_path:
        messagebox.showwarning(
            "Faltan archivos", "Debes seleccionar tanto la plantilla como el archivo Excel.")
        return

    try:
        df = pd.read_excel(excel_path)
        df.columns = df.columns.str.lower()
        output_folder = os.path.join(
            os.path.dirname(excel_path), "Contratos generados")
        os.makedirs(output_folder, exist_ok=True)

        log_path = os.path.join(output_folder, "log_contratos.txt")
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
                ruta_archivo = os.path.join(output_folder, nombre_archivo)
                doc.save(ruta_archivo)

                log_file.write(f"✅ {nombre_archivo} generado correctamente.\n")
                total_generados += 1

        messagebox.showinfo(
            "Éxito", f"Se generaron {total_generados} contratos en:\n{output_folder}")
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


# Inicialización
plantilla_path = ""
excel_path = ""

ventana = Tk()
ventana.title("Generador de Contratos")
ventana.geometry("400x250")

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

Button(ventana, text="Generar contratos", command=generar_contratos,
       bg="#4CAF50", fg="white").pack(pady=20)

ventana.mainloop()
