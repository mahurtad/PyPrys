import fitz  # PyMuPDF
import os

# === CONFIGURACIÓN ===
pdf_path = r"G:\My Drive\Shared GyM\UCSUR\Archivos para tutores virtuales\Silabo\SILABO - QUÍMICA ORGÁNICA.pdf"
output_folder = r"G:\My Drive\Shared GyM\UCSUR\Archivos para tutores virtuales\Silabo\Capturas\QUÍMICA ORGÁNICA"
palabras_clave = ["sesión", "actividad", "semana", "interacción", "autónomo"]  # puedes ajustar estas
tolerancia_paginas_vacias = 2  # cuántas páginas sin clave se permiten antes de detenerse

# === CREAR CARPETA DESTINO ===
os.makedirs(output_folder, exist_ok=True)

# === ABRIR PDF ===
doc = fitz.open(pdf_path)

# === BUSCAR INICIO DEL APARTADO 8 ===
pagina_inicio = None
for i, page in enumerate(doc):
    texto = page.get_text().lower()
    if "actividades principales" in texto:
        pagina_inicio = i
        print(f"🔍 Título '8. ACTIVIDADES PRINCIPALES' encontrado en la página {i + 1}")
        break

if pagina_inicio is None:
    print("❌ No se encontró el apartado 8 en el documento.")
else:
    # Captura dinámica
    paginas_vacias = 0
    page_number = pagina_inicio

    while page_number < len(doc):
        texto = doc[page_number].get_text().lower()

        if any(palabra in texto for palabra in palabras_clave):
            paginas_vacias = 0  # reset si hay actividad relacionada
        else:
            paginas_vacias += 1
            if paginas_vacias >= tolerancia_paginas_vacias:
                print(f"🛑 Fin del contenido detectado en la página {page_number + 1}")
                break

        # Guardar imagen
        pix = doc[page_number].get_pixmap(dpi=300)
        output_image = os.path.join(output_folder, f"pagina_{page_number + 1}.png")
        pix.save(output_image)
        print(f"📸 Capturada: {output_image}")

        page_number += 1

    print("✅ Capturas finalizadas.")
