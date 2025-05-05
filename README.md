# 📝 GenContratos

Generador de contratos masivos en Word a partir de una plantilla y un Excel. Ideal para contratistas, contadores, abogados o cualquier persona que necesite automatizar la creación de contratos.

---

## 🚀 ¿Qué hace?

- Carga una plantilla de contrato en Word (`.docx`)
- Reemplaza campos como `{{NOMBRE}}`, `{{RUT}}`, `{{FECHA}}`, etc.
- Usa una planilla Excel con los datos de cada persona
- Genera un archivo `.docx` por cada contrato completo

---

## 📂 Estructura del proyecto

```
GenContratos/
├── plantilla_contrato.docx         # Documento base con etiquetas
├── datos_contratos.xlsx            # Datos para generar contratos
├── GenContratos.py                 # Script principal
├── contratos_generados/           # Carpeta con contratos ya creados
└── .gitignore                      # Archivos a ignorar por Git
```

---

## ▶️ Cómo se usa

1. Abrí el Excel `datos_contratos.xlsx` y llenalo con tus datos
2. Asegurate que el Word tenga etiquetas como `{{NOMBRE}}`, `{{RUT}}`, etc.
3. Corre el script:

```bash
python GenContratos.py
```

4. ¡Listo! Los contratos se guardan en la carpeta `contratos_generados`

---

## 🛠 Requisitos

- Python 3.8 o superior
- Librerías:
  - `pandas`
  - `python-docx`

Instalalas con:

```bash
pip install pandas python-docx
```

---

## 📌 Ejemplo de uso

| NOMBRE     | RUT           | CARGO      | FECHA      |
|------------|---------------|------------|------------|
| Pedro Pérez | 12.345.678-9 | Eléctrico  | 2025-05-05 |

> Y el contrato se genera mágicamente con esos datos en un Word 🪄

---

## 🤝 Licencia

Licencia de uso personal. Si querís una versión con más funciones o soporte, hablame 😉
