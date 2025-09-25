import os
import pandas as pd

# Configuración de rutas
pptx_dir = r"C:\Users\manue\Downloads\OneDrive_2024-11-29\Generador de Constancias\Output2"
pdf_dir = r"C:\Users\manue\Downloads\OneDrive_2024-11-29\Generador de Constancias\Output2\pdfs"
output_excel = r"C:\Users\manue\Downloads\OneDrive_2024-11-29\Generador de Constancias\lista_archivos.xlsx"

# Listar nombres de archivos PPTX sin extensión
pptx_files = [os.path.splitext(f)[0] for f in os.listdir(pptx_dir) if f.endswith('.pptx')]

# Listar nombres de archivos PDF sin extensión
pdf_files = [os.path.splitext(f)[0] for f in os.listdir(pdf_dir) if f.endswith('.pdf')]

# Igualar el número de filas entre PPTX y PDF rellenando con cadenas vacías
max_length = max(len(pptx_files), len(pdf_files))
pptx_files += [''] * (max_length - len(pptx_files))
pdf_files += [''] * (max_length - len(pdf_files))

# Crear DataFrame
data = {'PPTS': pptx_files, 'PDFS': pdf_files}
df = pd.DataFrame(data)

# Exportar a Excel
df.to_excel(output_excel, index=False)

print(f"Lista de archivos exportada a: {output_excel}")
