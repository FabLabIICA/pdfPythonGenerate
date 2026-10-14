FROM python:3.11-slim

# Instalar LibreOffice y dependencias del sistema
RUN apt-get update && apt-get install -y \
    libreoffice \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Instalar dependencias de Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el código del proyecto
COPY . .

# Comando para iniciar la aplicación usando el puerto asignado por Render
CMD gunicorn app:app --bind 0.0.0.0:$PORT
