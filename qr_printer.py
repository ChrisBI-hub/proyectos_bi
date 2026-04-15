"""
qr_printer.py
=============
Monitorea una carpeta de Google Drive cada 1 hora.
Descarga las imágenes QR nuevas (no impresas aún), genera un PDF
con varias imágenes por hoja y lo envía a la impresora.

Los PDFs se nombran con un contador incremental para diferenciarlos:
    qr_para_imprimir_001.pdf
    qr_para_imprimir_002.pdf
    ...

Uso standalone:
    python3 qr_printer.py          # corre indefinidamente (cada 1 hora)
    python3 qr_printer.py --once   # ejecuta un solo ciclo y sale

Cuando se usa desde main.py, este módulo expone:
    ejecutar_ciclo()   → un solo ciclo
    main_loop()        → bucle infinito cada INTERVALO_SEGUNDOS
"""

import os
import io
import json
import pickle
import shutil
import argparse
import time
from pathlib import Path
from typing import List, Dict, Set
from datetime import datetime

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

# ══════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════

# Carpeta de Google Drive a monitorear (ID o 'root')
DRIVE_FOLDER_ID  = '1dyPhza4IcbsTCrhjOVe50arLK39GYziK'

# Prefijo de nombre de archivo a buscar
NAME_PREFIX      = 'QR'

# Extensiones de imagen válidas
VALID_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff'}

# Dimensiones (en cm)
QR_SIZE_CM  = 2.5
SPACING_CM  = 0.5
MARGIN_CM   = 0.5

# Archivos de estado
RECORDS_FILE  = 'printed_records.json'   # IDs ya impresos
COUNTER_FILE  = 'pdf_counter.json'       # contador de PDFs generados

# Carpeta de salida para los PDFs generados
PDF_OUTPUT_DIR = 'pdfs_generados'

# Carpeta temporal para imágenes descargadas
TEMP_DIR      = 'temp_qr_images'

# Intervalo del bucle (en segundos) — 300 s = 5 minutos
INTERVALO_SEGUNDOS = 300

# OAuth scopes
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']


# ══════════════════════════════════════════════
# CONTADOR DE PDFs
# ══════════════════════════════════════════════
def cargar_contador() -> int:
    """Devuelve el valor actual del contador (0 si no existe o está vacío/corrupto)."""
    if os.path.exists(COUNTER_FILE):
        try:
            with open(COUNTER_FILE, 'r') as f:
                contenido = f.read().strip()
                if not contenido:
                    return 0
                return json.loads(contenido).get('contador', 0)
        except (json.JSONDecodeError, ValueError):
            print(f"   ⚠️  {COUNTER_FILE} estaba corrupto — reiniciando contador a 0.")
            return 0
    return 0


def guardar_contador(valor: int):
    with open(COUNTER_FILE, 'w') as f:
        json.dump({'contador': valor}, f)


def siguiente_numero_pdf() -> tuple[int, str]:
    """
    Incrementa el contador y devuelve (numero, ruta_pdf).
    Ejemplo: (7, 'pdfs_generados/qr_para_imprimir_007.pdf')
    """
    n = cargar_contador() + 1
    guardar_contador(n)
    os.makedirs(PDF_OUTPUT_DIR, exist_ok=True)
    nombre = f"qr_para_imprimir_{n:03d}.pdf"
    ruta   = os.path.join(PDF_OUTPUT_DIR, nombre)
    return n, ruta


# ══════════════════════════════════════════════
# AUTENTICACIÓN GOOGLE DRIVE
# ══════════════════════════════════════════════
def get_drive_service():
    """Autentica y devuelve el servicio de Google Drive."""
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('drive', 'v3', credentials=creds)


# ══════════════════════════════════════════════
# REGISTRO DE IMPRESOS
# ══════════════════════════════════════════════
def load_printed_records() -> Set[str]:
    if os.path.exists(RECORDS_FILE):
        try:
            with open(RECORDS_FILE, 'r') as f:
                contenido = f.read().strip()
                if not contenido:
                    return set()
                return set(json.loads(contenido).get('printed_ids', []))
        except (json.JSONDecodeError, ValueError):
            print(f"   ⚠️  {RECORDS_FILE} estaba corrupto — se procesarán todas las imágenes.")
            return set()
    return set()


def save_printed_records(printed_ids: Set[str]):
    with open(RECORDS_FILE, 'w') as f:
        json.dump({'printed_ids': list(printed_ids)}, f)


# ══════════════════════════════════════════════
# LISTAR IMÁGENES QR EN DRIVE
# ══════════════════════════════════════════════
def list_qr_images(service, folder_id: str) -> List[Dict]:
    query = f"name contains '{NAME_PREFIX}' and trashed = false"
    if folder_id != 'root':
        query += f" and '{folder_id}' in parents"

    results    = []
    page_token = None

    while True:
        response = service.files().list(
            q=query,
            spaces='drive',
            fields='nextPageToken, files(id, name, mimeType)',
            pageToken=page_token
        ).execute()

        for file in response.get('files', []):
            ext = os.path.splitext(file['name'])[1].lower()
            if ext in VALID_EXTENSIONS or file['mimeType'].startswith('image/'):
                results.append(file)

        page_token = response.get('nextPageToken')
        if not page_token:
            break

    return results


def download_image(service, file_id: str, destination_path: str):
    request = service.files().get_media(fileId=file_id)
    with open(destination_path, 'wb') as f:
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()


# ══════════════════════════════════════════════
# GENERACIÓN DE PDF
# ══════════════════════════════════════════════
def generate_pdf(image_paths: List[str], output_pdf_path: str):
    qr_size_pt = QR_SIZE_CM * cm
    spacing_pt = SPACING_CM * cm
    margin_pt  = MARGIN_CM  * cm
    cell_width  = qr_size_pt + spacing_pt
    cell_height = qr_size_pt + spacing_pt

    page_width, page_height = letter
    usable_width  = page_width  - 2 * margin_pt
    usable_height = page_height - 2 * margin_pt

    cols = int((usable_width  + spacing_pt) // cell_width)
    rows = int((usable_height + spacing_pt) // cell_height)

    if cols <= 0 or rows <= 0:
        raise ValueError("El tamaño de imagen y márgenes no permiten colocar ninguna imagen.")

    print(f"   📐 Cuadrícula: {cols} col × {rows} fil  ({cols * rows} QR por página)")

    total_grid_width  = cols * cell_width  - spacing_pt
    start_x = margin_pt + (usable_width - total_grid_width) / 2
    start_y = page_height - margin_pt - qr_size_pt

    c         = canvas.Canvas(output_pdf_path, pagesize=letter)
    img_index = 0
    total     = len(image_paths)

    while img_index < total:
        for row in range(rows):
            y = start_y - row * cell_height
            for col in range(cols):
                if img_index >= total:
                    break
                x = start_x + col * cell_width
                try:
                    c.drawImage(
                        image_paths[img_index], x, y,
                        width=qr_size_pt, height=qr_size_pt,
                        preserveAspectRatio=True, anchor='c'
                    )
                except Exception as e:
                    print(f"   ⚠️  Error insertando {image_paths[img_index]}: {e}")
                img_index += 1
            if img_index >= total:
                break
        if img_index < total:
            c.showPage()
            start_y = page_height - margin_pt - qr_size_pt

    c.save()


# ══════════════════════════════════════════════
# IMPRESIÓN
# ══════════════════════════════════════════════
def print_pdf(pdf_path: str):
    if os.name == 'nt':
        os.startfile(pdf_path, "print")
    elif os.uname().sysname == 'Darwin':
        os.system(f'lpr "{pdf_path}"')
    else:
        os.system(f'lpr "{pdf_path}"')


# ══════════════════════════════════════════════
# CICLO ÚNICO (llamado por main.py o --once)
# ══════════════════════════════════════════════
def ejecutar_ciclo():
    """
    Un ciclo completo:
      1. Conectar a Drive
      2. Listar QRs nuevos
      3. Descargar → generar PDF numerado → imprimir
      4. Actualizar registros
    """
    hora = datetime.now().strftime('%H:%M:%S')
    print(f"\n[{hora}] 🖨️  QR_PRINTER — Iniciando ciclo...")

    service = get_drive_service()
    print(f"   🔍 Buscando imágenes con prefijo '{NAME_PREFIX}' en Drive...")
    qr_files = list_qr_images(service, DRIVE_FOLDER_ID)
    print(f"   📁 Total en Drive: {len(qr_files)}")

    if not qr_files:
        print("   ℹ️  Sin imágenes en la carpeta.")
        return

    printed_ids = load_printed_records()
    new_files   = [f for f in qr_files if f['id'] not in printed_ids]
    print(f"   🆕 Nuevas por imprimir: {len(new_files)}")

    if not new_files:
        print("   ✅ Todas las imágenes ya fueron impresas.")
        return

    temp_path = Path(TEMP_DIR)
    temp_path.mkdir(exist_ok=True)
    downloaded_paths = []
    new_ids          = []

    try:
        for file in new_files:
            safe_name  = "".join(c for c in file['name'] if c.isalnum() or c in " .-_()").rstrip()
            local_path = temp_path / f"{file['id']}_{safe_name}"
            print(f"   ⬇️  Descargando: {file['name']}...", end=" ", flush=True)
            download_image(service, file['id'], str(local_path))
            print("✅")
            downloaded_paths.append(str(local_path))
            new_ids.append(file['id'])

        # Generar PDF con número incremental
        num, pdf_path = siguiente_numero_pdf()
        print(f"   📄 Generando PDF #{num:03d}: {pdf_path}")
        generate_pdf(downloaded_paths, pdf_path)
        print(f"   ✅ PDF guardado → {pdf_path}")

        # Imprimir
        print(f"   🖨️  Enviando a impresora...")
        print_pdf(pdf_path)
        print(f"   ✅ Enviado.")

        # Actualizar registro de impresos
        printed_ids.update(new_ids)
        save_printed_records(printed_ids)
        print(f"   💾 Registro actualizado ({len(new_ids)} nuevo(s) marcado(s)).")

    except Exception as e:
        print(f"   ❌ Error en el ciclo: {e}")
    finally:
        if temp_path.exists():
            shutil.rmtree(temp_path)


# ══════════════════════════════════════════════
# BUCLE CONTINUO (llamado por main.py)
# ══════════════════════════════════════════════
def main_loop():
    """Bucle infinito que llama a ejecutar_ciclo() cada INTERVALO_SEGUNDOS."""
    print("▶️  QR_PRINTER corriendo en modo continuo.")
    while True:
        ejecutar_ciclo()
        mins = INTERVALO_SEGUNDOS // 60
        print(f"   ⏳ Próximo ciclo en {mins} min ({INTERVALO_SEGUNDOS}s)...\n")
        time.sleep(INTERVALO_SEGUNDOS)


# ══════════════════════════════════════════════
# ENTRADA STANDALONE
# ══════════════════════════════════════════════
def main():
    parser = argparse.ArgumentParser(description="qr_printer — Drive → PDF numerado → Impresora")
    parser.add_argument("--once", action="store_true", help="Ejecutar un solo ciclo y salir")
    args = parser.parse_args()

    print("=" * 52)
    print("  QR_PRINTER — Google Drive → PDF → Impresora")
    print(f"  Carpeta Drive : {DRIVE_FOLDER_ID}")
    print(f"  Intervalo     : {INTERVALO_SEGUNDOS // 60} min")
    print(f"  Salida PDFs   : {PDF_OUTPUT_DIR}/")
    print("=" * 52)

    if args.once:
        ejecutar_ciclo()
    else:
        main_loop()


if __name__ == "__main__":
    main()