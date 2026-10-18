# Documentos a Markdown con Docling

Conversor para Windows de PDF, DOCX y PPTX a Markdown. Incluye interfaz gráfica, procesamiento por lotes, OCR opcional, extracción de tablas, tratamiento de SmartArt y aislamiento del procesamiento de cada documento.

## Instalación

1. Instala Python y crea un entorno en esta carpeta: `py -m venv venv`.
2. Instala dependencias: `venv\Scripts\python.exe -m pip install -r requirements.txt`.
3. Para OCR, instala Tesseract OCR. La ruta predeterminada es `C:\Program Files\Tesseract-OCR\tesseract.exe`.
4. Ejecuta `iniciar_docling.bat` y elige consola o interfaz gráfica.

El procesamiento puede usar GPU o CPU según la configuración y el hardware disponible. El archivo principal es `docling_pdf2md.py`.
