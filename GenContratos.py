from tkinter import Tk, Label, Button, filedialog, messagebox, PhotoImage
import pandas as pd
from docx import Document
import os
from datetime import datetime

# Centrar la ventana


def centrar_ventana(ventana, ancho, alto):
    pantalla_ancho = ventana.winfo_screenwidth()
    pantalla_alto = ventana.winfo_screenheight()
    x = int((pantalla_ancho / 2) - (ancho / 2))
    y = int((pantalla_alto / 2) - (alto / 2))
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")

# Funciones de hover


def on_enter(e):
    e.widget.config(bg="#45a049")


def on_leave(e):
    e.widget.config(bg="#4CAF50")


def generar_contratos():
    if not plantilla_path or not excel_path or not output_folder:
        messagebox.showwarning(
            "Faltan archivos", "Debes seleccionar la plantilla, el archivo Excel y la carpeta de destino.")
        return

    try:
        df = pd.read_excel(excel_path)

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
                    fila.get("NOMBRE", "desconocido")).strip().replace(" ", "_")
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
ventana.configure(bg="#f2f2f2")
ventana.resizable(False, False)
centrar_ventana(ventana, 400, 360)

# Icono (opcional, asegurate que esté en esa ruta o cambiala)
try:
    icono = PhotoImage(
        file=r"C:\Users\usuario\Desktop\Proyecto Personal\GeneraContratos\icono.png")
    ventana.iconphoto(True, icono)
except Exception as e:
    print("No se pudo cargar el ícono:", e)

# Título
Label(ventana, text="Generador de Contratos",
      font=("Helvetica", 16, "bold"), bg="#f2f2f2", fg="#333333").pack(pady=10)

# Botones y etiquetas
Button(ventana, text="Seleccionar plantilla Word",
       command=seleccionar_plantilla, bg="#4CAF50", fg="white",
       activebackground="#45a049").pack(pady=5)
label_plantilla = Label(
    ventana, text="Plantilla: No seleccionada", bg="#f2f2f2", fg="#555555")
label_plantilla.pack()

Button(ventana, text="Seleccionar archivo Excel",
       command=seleccionar_excel, bg="#4CAF50", fg="white",
       activebackground="#45a049").pack(pady=5)
label_excel = Label(ventana, text="Excel: No seleccionado",
                    bg="#f2f2f2", fg="#555555")
label_excel.pack()

Button(ventana, text="Seleccionar carpeta de destino",
       command=seleccionar_carpeta_destino, bg="#4CAF50", fg="white",
       activebackground="#45a049").pack(pady=5)
label_carpeta = Label(
    ventana, text="Carpeta de destino: No seleccionada", bg="#f2f2f2", fg="#555555")
label_carpeta.pack()

# Botón de generar contratos
boton_generar = Button(ventana, text="Generar contratos", command=generar_contratos,
                       bg="#4CAF50", fg="white", activebackground="#45a049", font=("Helvetica", 10, "bold"))
boton_generar.pack(pady=20)
boton_generar.bind("<Enter>", on_enter)
boton_generar.bind("<Leave>", on_leave)

ventana.mainloop()
