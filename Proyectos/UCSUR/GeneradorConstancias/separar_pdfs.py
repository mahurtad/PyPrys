import os
import pdfplumber

# Ruta del directorio que contiene los PDFs
input_dir = "C:\\Users\\manue\\Downloads\\Generador\\CAPACITACIÓN DISEÑO DE EXPERIENCIAS DE APRENDIZAJE CON INTELIGENCIA ARTIFICIAL\\Certificados"

try:
    # Listar todos los archivos PDF en el directorio
    for filename in os.listdir(input_dir):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(input_dir, filename)
            
            # Abrir el archivo PDF para extraer texto
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[0]  # Suponiendo que el nombre está en la primera página
                text = page.extract_text()

                # Suponiendo que el nombre aparece después de "constancia a:" y termina antes de una coma
                if "constancia a:" in text:
                    start = text.find('constancia a:') + len('constancia a:')
                    end = text.find(',', start)
                    name = text[start:end].strip()

                    # Limpiar el nombre para que sea apto como nombre de archivo
                    new_filename = "".join(char for char in name if char.isalnum() or char in " _-").rstrip() + ".pdf"
                    new_path = os.path.join(input_dir, new_filename)

                    # Renombrar el archivo
                    os.rename(pdf_path, new_path)
                    print(f"Renamed {filename} to {new_filename}")
except Exception as e:
    print(f"An error occurred: {e}")
