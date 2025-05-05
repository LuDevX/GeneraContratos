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
        output_folder = os.path.join(os.path.dirname(
            excel_path), "contratos_generados_gui")
        os.makedirs(output_folder, exist_ok=True)

        for _, fila in df.iterrows():
            doc = Document(plantilla_path)
            for parrafo in doc.paragraphs:
                for clave, valor in fila.items():
                    if pd.notna(valor):
                        parrafo.text = parrafo.text.replace(
                            f"{{{{{clave}}}}}", str(valor))
            nombre = f"Contrato_{fila.iloc[0].replace(' ', '_')}_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
            doc.save(os.path.join(output_folder, nombre))

        messagebox.showinfo(
            "Éxito", f"Contratos generados en:\n{output_folder}")
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
