from docxtpl import DocxTemplate
import pandas as pd
import os
from datetime import datetime

df = pd.read_excel("datos_contratos.xlsx")
plantilla = "plantilla_contrato.docx"
output_dir = "contratos_generados"
os.makedirs(output_dir, exist_ok=True)

for i, row in df.iterrows():
    doc = DocxTemplate(plantilla)
    datos = row.to_dict()

    # Formatear automáticamente todas las fechas
    for key, value in datos.items():
        if isinstance(value, (datetime, pd.Timestamp)):
            datos[key] = value.strftime("%d-%m-%Y")

    nombre_archivo = f"Contrato_{datos['NOMBRE'].replace(' ', '_')}.docx"
    doc.render(datos)
    doc.save(os.path.join(output_dir, nombre_archivo))

print("¡Contratos generados!")
