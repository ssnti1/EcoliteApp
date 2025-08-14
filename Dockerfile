FROM python:3.11-slim

# instalar dependencias de sistema para tkinter y entorno gráfico
RUN apt-get update && apt-get install -y \
    python3-tk \
    xvfb \
    x11vnc \
    fluxbox \
    websockify \
    novnc \
    wget \
    && rm -rf /var/lib/apt/lists/*

# carpeta de trabajo en el contenedor
WORKDIR /app

# copiar e instalar dependencias de python
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# copiar el resto de la app
COPY . /app

# dar permisos de ejecución al script de arranque
RUN chmod +x /app/start.sh

# puerto donde se servirá noVNC
EXPOSE 8080

# comando de arranque
CMD ["/app/start.sh"]
