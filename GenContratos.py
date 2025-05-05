from docxtpl import DocxTemplate
import pandas as pd
import os

# Cargar Excel y plantilla
df = pd.read_excel("datos_contratos.xlsx")
plantilla = "plantilla_contrato.docx"
output_dir = "contratos_generados"
os.makedirs(output_dir, exist_ok=True)

# Generar contratos uno por uno
for i, row in df.iterrows():
    doc = DocxTemplate(plantilla)
    doc.render(row.to_dict())
    nombre_archivo = f"Contrato_{row['NOMBRE'].replace(' ', '_')}.docx"
    doc.save(os.path.join(output_dir, nombre_archivo))

print("Contratos generados exitosamente!")
