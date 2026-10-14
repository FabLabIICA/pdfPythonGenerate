import io
import re
import zipfile
import openpyxl
import pandas as pd
import requests
from docx import Document
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS

app = Flask(__name__)
CORS(app)


def convertir_enlace_drive(url):
  """Convierte un enlace de Google Drive o Google Docs en un enlace de descarga directa."""
  if not url:
    return url

  # Si es un documento de Google Docs, exportarlo como docx
  match_docs = re.search(r'/document/d/([a-zA-Z0-9_-]+)', url)
  if match_docs:
    file_id = match_docs.group(1)
    return f'https://docs.google.com/document/d/{file_id}/export?format=docx'

  # Si es un archivo estándar en Google Drive
  match_drive = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
  if match_drive:
    file_id = match_drive.group(1)
    return f'https://drive.google.com/uc?export=download&id={file_id}'

  return url


def descargar_machote_docx(url):
  """Descarga el archivo docx desde el enlace procesado de Google Drive."""
  url_directa = convertir_enlace_drive(url)
  response = requests.get(url_directa)
  if response.status_code != 200:
    raise Exception(
        'No se pudo descargar el machote desde Google Drive. Asegúrate de que'
        ' el archivo sea público.'
    )
  return io.BytesIO(response.content)


def extraer_etiquetas_docx(doc):
  """Extrae todas las etiquetas <<etiqueta>> o {{etiqueta}} del documento Word."""
  etiquetas = set()

  # Buscar en párrafos
  for p in doc.paragraphs:
    texto = ''.join([run.text for run in p.runs])
    matches = re.findall(r'<<(.*?)>>|\{\{(.*?)\}\}', texto)
    for m in matches:
      tag = m[0] if m[0] else m[1]
      if tag.strip():
        etiquetas.add(tag.strip())

  # Buscar en tablas
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


def reemplazar_etiquetas_en_doc(doc, row):
  """Reemplaza las etiquetas <<etiqueta>> o {{etiqueta}} con los datos de la fila."""
  for paragraph in doc.paragraphs:
    texto_completo = ''.join([run.text for run in paragraph.runs])
    for key, val in row.items():
      val_str = '' if pd.isna(val) else str(val)
      texto_completo = texto_completo.replace(f'<<{key}>>', val_str)
      texto_completo = texto_completo.replace(f'{{{{{key}}}}}', val_str)

    if paragraph.runs:
      paragraph.runs[0].text = texto_completo
      for run in paragraph.runs[1:]:
        run.text = ''

  for table in doc.tables:
    for row_table in table.rows:
      for cell in row_table.cells:
        for p in cell.paragraphs:
          texto_celda = ''.join([run.text for run in p.runs])
          for key, val in row.items():
            val_str = '' if pd.isna(val) else str(val)
            texto_celda = texto_celda.replace(f'<<{key}>>', val_str)
            texto_celda = texto_celda.replace(f'{{{{{key}}}}}', val_str)

          if p.runs:
            p.runs[0].text = texto_celda
            for run in p.runs[1:]:
              run.text = ''


@app.route('/api/plantilla-datos', methods=['POST'])
def plantilla_datos():
  """Endpoint que lee el machote de Drive, extrae las etiquetas y devuelve un Excel con ellas."""
  try:
    data = request.get_json()
    url_machote = data.get('url')

    if not url_machote:
      return jsonify({'error': 'Falta la URL del machote'}), 400

    # Descargar y analizar el Word
    machote_io = descargar_machote_docx(url_machote)
    doc = Document(machote_io)
    tags = extraer_etiquetas_docx(doc)

    if not tags:
      tags = [
          'nombre',
          'cedula',
          'curso',
          'fecha',
      ]  # Fallback si no encuentra etiquetas

    # Crear Excel en memoria con openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Plantilla Datos'
    ws.append(tags)

    excel_io = io.BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)

    return send_file(
        excel_io,
        mimetype=(
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ),
        as_attachment=True,
        download_name='plantilla_datos.xlsx',
    )
  except Exception as e:
    return jsonify({'error': str(e)}), 500


@app.route('/api/generar-certificados', methods=['POST'])
def generar_certificados():
  """Endpoint que descarga el machote del catálogo y procesa el archivo de datos subido."""
  try:
    url_machote = request.form.get('url_machote')
    file_datos = request.files.get('archivo_datos')

    if not url_machote or not file_datos:
      return jsonify({'error': 'Faltan parámetros o archivos'}), 400

    # Descargar machote desde Drive
    machote_bytes = descargar_machote_docx(url_machote).getvalue()

    # Leer archivo de datos
    if file_datos.filename.endswith('.xlsx'):
      df = pd.read_excel(file_datos)
    else:
      df = pd.read_csv(file_datos)

    # Crear ZIP con los certificados generados
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
      for index, row in df.iterrows():
        doc = Document(io.BytesIO(machote_bytes))
        reemplazar_etiquetas_en_doc(doc, row)

        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)

        nombre_val = (
            str(row.iloc[0]).replace(' ', '_')
            if len(row) > 0
            else f'user_{index}'
        )
        nombre_archivo = f'certificado_{nombre_val}.docx'
        zip_file.writestr(nombre_archivo, doc_io.read())

    zip_buffer.seek(0)
    return send_file(
        zip_buffer,
        mimetype='application/zip',
        as_attachment=True,
        download_name='certificados_generados.zip',
    )
  except Exception as e:
    return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
  app.run(host='0.0.0.0', port=5000)
