import io
import zipfile
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from docx import Document
import pandas as pd

# 1. Inicializar la aplicación Flask y configurar CORS para permitir peticiones desde GitHub Pages
app = Flask(__name__)
CORS(app)


def reemplazar_etiquetas(doc, row):
  """Función auxiliar para buscar y reemplazar etiquetas {{etiqueta}}

  en párrafos y tablas de un archivo Word (.docx).
  """
  # Reemplazar en párrafos normales
  for paragraph in doc.paragraphs:
    texto_completo = ''.join([run.text for run in paragraph.runs])
    for key, val in row.items():
      etiqueta = f'{{{{{key}}}}}'  # Busca el formato {{etiqueta}}
      if etiqueta in texto_completo:
        texto_completo = texto_completo.replace(etiqueta, str(val))
        if paragraph.runs:
          paragraph.runs[0].text = texto_completo
          for run in paragraph.runs[1:]:
            run.text = ''  # Vacía los runs sobrantes para conservar el formato original

  # Reemplazar dentro de tablas (común en diseños de certificados)
  for table in doc.tables:
    for row_table in table.rows:
      for cell in row_table.cells:
        for p in cell.paragraphs:
          texto_celda = ''.join([run.text for run in p.runs])
          for key, val in row.items():
            etiqueta = f'{{{{{key}}}}}'
            if etiqueta in texto_celda:
              texto_celda = texto_celda.replace(etiqueta, str(val))
              if p.runs:
                p.runs[0].text = texto_celda
                for run in p.runs[1:]:
                  run.text = ''


# 2. Endpoint principal que recibe la plantilla de Word y el Excel con los datos
@app.route('/generar-certificados', methods=['POST'])
def generar_certificados():
  try:
    # Validar que se hayan enviado ambos archivos obligatorios
    if 'plantilla' not in request.files or 'datos' not in request.files:
      return (
          jsonify({'error': 'Faltan archivos (plantilla o datos)'}),
          400,
      )

    file_plantilla = request.files['plantilla']
    file_datos = request.files['datos']

    # Leer el archivo de datos según su extensión (.xlsx o .csv)
    if file_datos.filename.endswith('.xlsx'):
      df = pd.read_excel(file_datos)
    else:
      df = pd.read_csv(file_datos)

    # 3. Crear un archivo ZIP en memoria para empaquetar todos los certificados generados
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
      for index, row in df.iterrows():
        # Cargar una copia fresca de la plantilla para cada iteración
        doc = Document(file_plantilla)
        reemplazar_etiquetas(doc, row)

        # Guardar el documento modificado en memoria temporal
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)

        # Asignar un nombre único al archivo basado en la primera columna del registro
        nombre_val = str(row.iloc[0]).replace(' ', '_')
        nombre_archivo = f'certificado_{nombre_val}.docx'

        # Añadir el certificado individual al archivo ZIP
        zip_file.writestr(nombre_archivo, doc_io.read())

    zip_buffer.seek(0)

    # 4. Enviar el archivo ZIP comprimido como respuesta de descarga al navegador
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
