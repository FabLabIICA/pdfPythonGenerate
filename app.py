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
        'No se pudo descargar el machote desde Google Drive. Verifica que el'
        ' archivo sea público.'
    )
  return io.BytesIO(response.content)


def extraer_etiquetas_docx(doc):
  """Extrae todas las etiquetas <<etiqueta>> o {{etiqueta}} del documento Word."""
  etiquetas = set()
  patron = re.compile(r'<<([^>]+>>)|\{\{([^}]+)\}\}')

  # Buscar en párrafos
  for p in doc.paragraphs:
    texto = ''.join([run.text for run in p.runs])
    matches = re.findall(r'<<(.*?)>>|\{\{(.*?)\}\}', texto)
    for m in matches:
      tag = m[0] if m[0] else m[1]
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
            etiquetas.add(tag.strip())

  return list(etiquetas)


def reemplazar_etiquetas_en_doc(doc, row):
  """Reemplaza las etiquetas en el documento Word con los datos de la fila."""
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
def descargar_plantilla_datos():
  """endpoint para generar y descargar un Excel con las columnas exactas extraídas del machote."""
  try:
    data = request.get_json()
    url_machote = data.get('url')

    if not url_machote:
      return jsonify({'error': 'No se proporcionó la URL del machote'}), 400

    # Descargar y leer el Word
    archivo_io = descargar_machote_docx(url_machote)
    doc = Document(archivo_io)
    tags = extraer_etiquetas_docx(doc)

    if not tags:
      tags = [
          'nombre',
          'cedula',
          'curso',
          'fecha',
      ]  # Fallback por si no encuentra etiquetas explícitas

    # Crear archivo Excel en memoria con openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Datos Certificados'

    # Escribir cabeceras con las etiquetas encontradas
    ws.append(tags)

    # Agregar una fila de ejemplo
    ejemplo = [f'Ejemplo {t}' for t in tags]
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
    return jsonify({'error': str(e)}), 500


@app.route('/api/generar-certificados', methods=['POST'])
def generar_certificados():
  """Endpoint que procesa el machote del catálogo y la hoja de datos subida."""
  try:
    url_machote = request.form.get('url_machote')
    file_datos = request.files.get('archivo_datos')

    if not url_machote or not file_datos:
      return jsonify({'error': 'Faltan datos o URL del machote'}), 400

    # Descargar machote base desde Google Drive
    machote_io = descargar_machote_docx(url_machote)

    # Leer archivo de datos (Excel o CSV)
    if file_datos.filename.endswith('.xlsx'):
      df = pd.read_excel(file_datos)
    else:
      df = pd.read_csv(file_datos)

    # Generar ZIP con certificados individuales
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
      for index, row in df.iterrows():
        doc = Document(io.BytesIO(machote_io.getvalue()))
        reemplazar_etiquetas_en_doc(doc, row)

        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)

        # Nombre de archivo basado en la primera columna o índice
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
