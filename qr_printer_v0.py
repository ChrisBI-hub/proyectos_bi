import os
import io
import json
import pickle
import tempfile
import shutil
from pathlib import Path
from typing import List, Dict, Set

# Google Drive
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# Imágenes y PDF
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas

# ================= CONFIGURACIÓN =================
# Carpeta de Google Drive a monitorear (ID o 'root')
DRIVE_FOLDER_ID = '1dyPhza4IcbsTCrhjOVe50arLK39GYziK'   # Cambia por el ID de tu carpeta si no es la raíz

# Patrón de nombre de archivo (case-insensitive)
NAME_PREFIX = 'QR'

# Extensiones de imagen válidas
VALID_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff'}

# Dimensiones (en cm)
QR_SIZE_CM = 2.5
SPACING_CM = 0.5
MARGIN_CM = 0.5

# Archivo de registro
RECORDS_FILE = 'printed_records.json'

# Carpeta temporal para imágenes descargadas
TEMP_DIR = 'temp_qr_images'

# SCOPES de Google Drive (solo lectura de archivos)
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

# ================= AUTENTICACIÓN GOOGLE DRIVE =================
def get_drive_service():
    """Autentica y devuelve el servicio de Google Drive."""
    creds = None
    # El archivo token.pickle almacena los tokens de acceso
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # Si no hay credenciales válidas, iniciar flujo OAuth
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Guardar las credenciales para la próxima ejecución
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('drive', 'v3', credentials=creds)

# ================= MANEJO DE REGISTRO DE IMPRESOS =================
def load_printed_records() -> Set[str]:
    """Carga el conjunto de IDs de archivos ya impresos."""
    if os.path.exists(RECORDS_FILE):
        with open(RECORDS_FILE, 'r') as f:
            data = json.load(f)
            return set(data.get('printed_ids', []))
    return set()

def save_printed_records(printed_ids: Set[str]):
    """Guarda el conjunto de IDs en el archivo de registro."""
    with open(RECORDS_FILE, 'w') as f:
        json.dump({'printed_ids': list(printed_ids)}, f)

# ================= LISTAR IMÁGENES QR EN DRIVE =================
def list_qr_images(service, folder_id: str) -> List[Dict]:
    """
    Obtiene todos los archivos de imagen cuyo nombre comienza con NAME_PREFIX
    en la carpeta especificada (y subcarpetas opcionalmente).
    """
    query = f"name contains '{NAME_PREFIX}' and trashed = false"
    if folder_id != 'root':
        query += f" and '{folder_id}' in parents"

    results = []
    page_token = None

    while True:
        response = service.files().list(
            q=query,
            spaces='drive',
            fields='nextPageToken, files(id, name, mimeType)',
            pageToken=page_token
        ).execute()
        
        for file in response.get('files', []):
            # Filtrar solo imágenes por extensión o mimeType
            name = file['name']
            ext = os.path.splitext(name)[1].lower()
            if ext in VALID_EXTENSIONS or file['mimeType'].startswith('image/'):
                results.append(file)
        
        page_token = response.get('nextPageToken')
        if not page_token:
            break

    return results

def download_image(service, file_id: str, destination_path: str):
    """Descarga un archivo de Drive a una ruta local."""
    request = service.files().get_media(fileId=file_id)
    with open(destination_path, 'wb') as f:
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            # print(f"Descargando {destination_path}: {int(status.progress() * 100)}%")

# ================= GENERACIÓN DE PDF =================
def generate_pdf(image_paths: List[str], output_pdf_path: str):
    """
    Crea un PDF tamaño carta con las imágenes organizadas en cuadrícula
    según las medidas definidas.
    """
    # Conversión cm -> puntos (1 cm = 28.3465 puntos)
    qr_size_pt = QR_SIZE_CM * cm
    spacing_pt = SPACING_CM * cm
    margin_pt = MARGIN_CM * cm

    cell_width = qr_size_pt + spacing_pt
    cell_height = qr_size_pt + spacing_pt

    page_width, page_height = letter  # 612 x 792 puntos

    # Calcular cuántas columnas y filas caben
    usable_width = page_width - 2 * margin_pt
    usable_height = page_height - 2 * margin_pt

    cols = int((usable_width + spacing_pt) // cell_width)
    rows = int((usable_height + spacing_pt) // cell_height)

    if cols <= 0 or rows <= 0:
        raise ValueError("El tamaño de imagen y márgenes no permiten colocar ninguna imagen.")

    print(f"Dimensiones: {cols} columnas x {rows} filas por página.")
    print(f"Total imágenes por página: {cols * rows}")

    c = canvas.Canvas(output_pdf_path, pagesize=letter)
    
    # Calcular offset para centrar el bloque de imágenes (opcional)
    total_grid_width = cols * cell_width - spacing_pt
    total_grid_height = rows * cell_height - spacing_pt
    start_x = margin_pt + (usable_width - total_grid_width) / 2
    start_y = page_height - margin_pt - qr_size_pt  # ReportLab y=0 es abajo

    img_index = 0
    total_images = len(image_paths)

    while img_index < total_images:
        for row in range(rows):
            y = start_y - row * cell_height
            for col in range(cols):
                if img_index >= total_images:
                    break
                x = start_x + col * cell_width
                
                # Dibujar la imagen
                try:
                    c.drawImage(image_paths[img_index], x, y, 
                                width=qr_size_pt, height=qr_size_pt, 
                                preserveAspectRatio=True, anchor='c')
                except Exception as e:
                    print(f"Error al insertar {image_paths[img_index]}: {e}")
                
                img_index += 1
            if img_index >= total_images:
                break
        # Nueva página si aún quedan imágenes
        if img_index < total_images:
            c.showPage()
            # Reiniciar posición para nueva página
            start_y = page_height - margin_pt - qr_size_pt

    c.save()
    print(f"PDF generado: {output_pdf_path}")

# ================= IMPRESIÓN DEL PDF =================
def print_pdf(pdf_path: str):
    """
    Envía el PDF a la impresora predeterminada.
    Funciona en Windows, Linux y macOS.
    """
    if os.name == 'nt':  # Windows
        os.startfile(pdf_path, "print")
        print("Documento enviado a la impresora (se abrirá el diálogo de impresión).")
    elif os.uname().sysname == 'Darwin':  # macOS
        os.system(f'lpr "{pdf_path}"')
    else:  # Linux
        os.system(f'lpr "{pdf_path}"')
    # Nota: En algunos sistemas puede requerir configuración adicional.

# ================= FLUJO PRINCIPAL =================
def main():
    print("Conectando con Google Drive...")
    service = get_drive_service()

    print(f"Buscando imágenes que empiezan con '{NAME_PREFIX}' en la carpeta ID: {DRIVE_FOLDER_ID}")
    qr_files = list_qr_images(service, DRIVE_FOLDER_ID)
    print(f"Encontradas {len(qr_files)} imágenes en total.")

    if not qr_files:
        print("No hay imágenes nuevas para procesar.")
        return

    # Cargar registro de impresos
    printed_ids = load_printed_records()
    
    # Filtrar las que no han sido impresas
    new_files = [f for f in qr_files if f['id'] not in printed_ids]
    print(f"Imágenes nuevas por procesar: {len(new_files)}")

    if not new_files:
        print("Todas las imágenes ya han sido impresas anteriormente.")
        return

    # Crear carpeta temporal
    temp_path = Path(TEMP_DIR)
    temp_path.mkdir(exist_ok=True)

    downloaded_paths = []
    new_ids = []

    try:
        # Descargar las nuevas imágenes
        for file in new_files:
            file_name = file['name']
            # Sanitizar nombre para sistema de archivos
            safe_name = "".join(c for c in file_name if c.isalnum() or c in " .-_()").rstrip()
            local_path = temp_path / f"{file['id']}_{safe_name}"
            
            print(f"Descargando: {file_name}")
            download_image(service, file['id'], str(local_path))
            downloaded_paths.append(str(local_path))
            new_ids.append(file['id'])

        # Generar PDF con las imágenes descargadas
        pdf_output = "qr_para_imprimir.pdf"
        generate_pdf(downloaded_paths, pdf_output)

        # Enviar a imprimir
        print("Enviando PDF a la impresora...")
        print_pdf(pdf_output)

        # Actualizar registro solo si la impresión fue exitosa (opcional: podrías agregar confirmación)
        printed_ids.update(new_ids)
        save_printed_records(printed_ids)
        print("Registro actualizado.")

    except Exception as e:
        print(f"Ocurrió un error: {e}")
    finally:
        # Limpiar archivos temporales
        if temp_path.exists():
            shutil.rmtree(temp_path)
            print("Carpeta temporal eliminada.")

if __name__ == '__main__':
    main()