import io
import os
import re
import subprocess
import tempfile
import zipfile
from docx import Document
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
import openpyxl
import pandas as pd
import requests

app = Flask(__name__)
CORS(app)


def convertir_enlace_drive(url):
    """Convierte enlaces de Google Drive / Docs para descarga directa."""
    if not url:
        return url
    match_docs = re.search(r'/document/d/([a-zA-Z0-9_-]+)', url)
    if match_docs:
        return (
            f'https://docs.google.com/document/d/{match_docs.group(1)}/export?format=docx'
        )
    match_drive = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
    if match_drive:
        return (
            f'https://drive.google.com/uc?export=download&id={match_drive.group(1)}'
        )
    return url


def descargar_machote_docx(url):
    url_directa = convertir_enlace_drive(url)
    response = requests.get(url_directa)
    if response.status_code != 200:
        raise Exception(
            'No se pudo descargar el machote desde Google Drive. Asegúrate de que sea público.'
        )
    return io.BytesIO(response.content)


def extraer_etiquetas_docx(doc_path):
    """Extrae etiquetas basadas en texto plano tipo {{etiqueta}} o <<etiqueta>>."""
    doc = Document(doc_path)
    etiquetas = set()
    
    for p in doc.paragraphs:
        texto = ''.join([run.text for run in p.runs])
        matches = re.findall(r'<<(.*?)>>|\{\{(.*?)\}\}', texto)
        for m in matches:
            tag = m[0] if m[0] else m[1]
            if tag.strip():
                etiquetas.add(tag.strip())
                
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    texto = ''.join([run.text for run in p.runs])
                    matches = re.findall(r'<<(.*?)>>|\{\{(.*?)\}\}', texto)
                    for m in matches:
                        tag = m[0] if m[0] else m[1]
                        if tag.strip():
                            etiquetas.add(tag.strip())

    return list(etiquetas)


def reemplazar_etiquetas_en_doc(doc, row_dict):
    """Reemplaza etiquetas del tipo {{tag}} o <<tag>> respetando el formato de los runs."""
    for p in doc.paragraphs:
        texto_parrafo = ''.join([run.text for run in p.runs])
        if "{{" in texto_parrafo or "<<" in texto_parrafo:
            for key, val in row_dict.items():
                val_str = str(val) if val is not None else ""
                texto_parrafo = texto_parrafo.replace(f"{{{{{key}}}}}", val_str)
                texto_parrafo = texto_parrafo.replace(f"<<{key}>>", val_str)
            
            if p.runs:
                p.runs[0].text = texto_parrafo
                for run in p.runs[1:]:
                    run.text = ""

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    texto_parrafo = ''.join([run.text for run in p.runs])
                    if "{{" in texto_parrafo or "<<" in texto_parrafo:
                        for key, val in row_dict.items():
                            val_str = str(val) if val is not None else ""
                            texto_parrafo = texto_parrafo.replace(f"{{{{{key}}}}}", val_str)
                            texto_parrafo = texto_parrafo.replace(f"<<{key}>>", val_str)
                        if p.runs:
                            p.runs[0].text = texto_parrafo
                            for run in p.runs[1:]:
                                run.text = ""


@app.route('/api/plantilla-datos', methods=['POST'])
def plantilla_datos():
    try:
        data = request.get_json()
        url_machote = data.get('url')
        if not url_machote:
            return jsonify({'error': 'Falta la URL del machote'}), 400

        machote_io = descargar_machote_docx(url_machote)

        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
            tmp.write(machote_io.getvalue())
            tmp_path = tmp.name

        tags = extraer_etiquetas_docx(tmp_path)
        os.unlink(tmp_path)

        if not tags:
            tags = ['nombre', 'cedula', 'curso', 'fecha']

        columnas = ['nombre_archivo_pdf'] + [
            t for t in tags if t != 'nombre_archivo_pdf'
        ]

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Plantilla Datos'
        ws.append(columnas)

        ejemplo = ['certificado_ejemplo'] + [f'Dato {t}' for t in columnas[1:]]
        ws.append(ejemplo)

        excel_io = io.BytesIO()
        wb.save(excel_io)
        excel_io.seek(0)

        return send_file(
            excel_io,
            mimetype=(
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            ),
            as_attachment=True,
            download_name='plantilla_datos_certificado.xlsx',
        )
    except Exception as e:
        import traceback
        return jsonify({'error': f'{str(e)} | Trace: {traceback.format_exc()}'}), 500


@app.route('/api/generar-certificados', methods=['POST'])
def generar_certificados():
    temp_dir = tempfile.mkdtemp()
    try:
        url_machote = request.form.get('url_machote')
        file_datos = request.files.get('archivo_datos')

        if not url_machote or not file_datos:
            return jsonify({'error': 'Faltan parámetros o archivos'}), 400

        machote_bytes = descargar_machote_docx(url_machote).getvalue()

        machote_path = os.path.join(temp_dir, 'machote.docx')
        with open(machote_path, 'wb') as f:
            f.write(machote_bytes)

        if file_datos.filename.endswith('.xlsx'):
            df = pd.read_excel(file_datos)
        else:
            df = pd.read_csv(file_datos)

        # PASO 1: Generar todos los archivos .docx
        archivos_docx = []
        for index, row in df.iterrows():
            row_dict = {}
            for k, v in row.items():
                row_dict[str(k)] = '' if pd.isna(v) else str(v)

            nombre_pdf = row_dict.get('nombre_archivo_pdf', '').strip()
            if not nombre_pdf:
                nombre_pdf = (
                    str(row.iloc[1]).replace(' ', '_')
                    if len(row) > 1
                    else f'certificado_{index}'
                )
            nombre_pdf = ''.join(
                c for c in nombre_pdf if c.isalnum() or c in ('_', '-', '.')
            )

            document = Document(machote_path)
            reemplazar_etiquetas_en_doc(document, row_dict)

            doc_out_path = os.path.join(temp_dir, f'{nombre_pdf}.docx')
            document.save(doc_out_path)
            archivos_docx.append(doc_out_path)

        # PASO 2: Convertir cada archivo asegurando la ruta exacta
        for doc_path in archivos_docx:
            try:
                subprocess.run(
                    [
                        'libreoffice',
                        '--headless',
                        '--convert-to',
                        'pdf',
                        '--outdir',
                        temp_dir,
                        doc_path,
                    ],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    check=True
                )
            except subprocess.CalledProcessError as sub_err:
                raise Exception(f"Error convirtiendo a PDF con LibreOffice: {sub_err.stderr}")

        # PASO 3: Empaquetar todos los PDFs generados en el archivo ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
            for filename in os.listdir(temp_dir):
                if filename.endswith('.pdf'):
                    pdf_path = os.path.join(temp_dir, filename)
                    with open(pdf_path, 'rb') as pf:
                        zip_file.writestr(filename, pf.read())

        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='certificados_pdf.zip',
        )

    except Exception as e:
        import traceback
        error_completo = traceback.format_exc()
        print(error_completo)
        return jsonify({'error': f'{str(e)} | Detalle: {error_completo}'}), 500
    finally:
        for root, dirs, files in os.walk(temp_dir, topdown=False):
            for name in files:
                try:
                    os.remove(os.path.join(root, name))
                except:
                    pass
            try:
                os.rmdir(root)
            except:
                pass



if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
