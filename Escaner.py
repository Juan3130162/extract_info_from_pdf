import os
import fitz  # PyMuPDF
import pandas as pd

def extract_text_from_pdf(pdf_path):
    """Extrae el texto de un PDF usando PyMuPDF."""
    doc = fitz.open(pdf_path)
    text = "\n".join(page.get_text("text") for page in doc)
    doc.close()
    return text

def process_pdfs_in_folder(folder_path, output_excel):
    """Procesa todos los PDFs en una carpeta y guarda cada palabra en una fila en un archivo Excel."""
    pdf_data = []
    
    for file_name in os.listdir(folder_path):
        if file_name.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder_path, file_name)
            print(f"Procesando: {file_name}")
            text = extract_text_from_pdf(pdf_path)
            words = text.split()  # Divide el texto en palabras
            for word in words:
                pdf_data.append([file_name, word])
    
    # Guardar en Excel
    df = pd.DataFrame(pdf_data, columns=["Archivo", "Palabra"])
    df.to_excel(output_excel, index=False)
    print(f"\nInformación guardada en {output_excel}")

# Ruta de la carpeta con PDFs y nombre del archivo Excel
folder_path = "./pdfs" 
output_excel = "agrupado.xlsx"

process_pdfs_in_folder(folder_path, output_excel)

