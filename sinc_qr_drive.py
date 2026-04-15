"""
sync_qr_drive.py
================
Corre en bucle cada 60 segundos. Revisa el buzón IMAP buscando
correos con el asunto exacto definido en ASUNTO_TRIGGER.
Por cada correo nuevo encontrado:
  1. Descarga los adjuntos de imagen (PNG, JPG, JPEG)
  2. Los sube directamente a la carpeta raíz de Google Drive
     (sin crear subcarpetas por fecha)
  3. Marca el correo como leído para no procesarlo dos veces
  4. Guarda las imágenes localmente como respaldo

Uso:
    python3 sync_qr_drive.py           # corre indefinidamente
    python3 sync_qr_drive.py --once    # revisa una sola vez y sale

Configuración inicial (hacer UNA sola vez):
    Ver instrucciones de OAuth en los comentarios de CONFIGURACIÓN.
"""

import sys
import subprocess

# ──────────────────────────────────────────────
# AUTO-INSTALACIÓN DE DEPENDENCIAS
# ──────────────────────────────────────────────
def instalar_dependencias():
    requeridos = {
        "google.oauth2.credentials":   "google-auth",
        "google_auth_oauthlib.flow":   "google-auth-oauthlib",
        "googleapiclient.discovery":   "google-api-python-client",
    }
    print("🐍 Verificando dependencias de Drive...")
    faltantes = []
    for modulo, paquete in requeridos.items():
        try:
            __import__(modulo)
            print(f"   ✅ {paquete}")
        except ImportError:
            print(f"   ⬇️  {paquete} → se instalará")
            faltantes.append(paquete)

    if faltantes:
        for paquete in faltantes:
            print(f"   Instalando {paquete}...", end=" ", flush=True)
            r = subprocess.run(
                [sys.executable, "-m", "pip", "install", "--quiet", paquete],
                capture_output=True, text=True
            )
            print("✅" if r.returncode == 0 else f"❌ {r.stderr.strip()}")
            if r.returncode != 0:
                sys.exit(1)
        print()

instalar_dependencias()

# ──────────────────────────────────────────────
# IMPORTS
# ──────────────────────────────────────────────
import os
import re
import json
import imaplib
import email
import mimetypes
import argparse
import time
from datetime import datetime
from pathlib import Path

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.errors import HttpError
import io

# ──────────────────────────────────────────────
# CONFIGURACIÓN  ← editar estos valores
# ──────────────────────────────────────────────

# ── Correo IMAP ───────────────────────────────
IMAP_SERVER    = "imap.gmail.com"
IMAP_PORT      = 993
EMAIL_USER     = "reportes.bi@abcsc.mx"
EMAIL_PASS     = "jwvjdrvmprzrwzxy"

# Asunto exacto que dispara la subida de imágenes
ASUNTO_TRIGGER = "ACTUALIZACION_BASE_DE_DATOS"   # ← cambia al asunto de tus correos con QR

# ── Google Drive OAuth ────────────────────────
# JSON descargado de Google Cloud Console (Aplicación de escritorio)
OAUTH_CLIENT_FILE = "/home/christian/Documentos/proyectos_bi/oauth_client.json"

# Token generado automáticamente tras el primer login (no editar)
TOKEN_FILE        = "/home/christian/Documentos/proyectos_bi/token_qr.json"

# ID de la carpeta destino en Drive (solo el ID, no la URL completa)
DRIVE_FOLDER_ID   = "1dyPhza4IcbsTCrhjOVe50arLK39GYziK"

# ── Rutas locales ─────────────────────────────
# Carpeta donde se guardan las imágenes como respaldo
CARPETA_RESPALDO  = "/home/christian/Documentos/proyectos_bi/respaldo_qr"

# Registro de archivos ya subidos a Drive (evita duplicados entre ciclos)
REGISTRO_SUBIDAS  = "/home/christian/Documentos/proyectos_bi/drive_qr_subidas.json"

# ── Otras opciones ────────────────────────────
EXTENSIONES_VALIDAS = {".png", ".jpg", ".jpeg"}
INTERVALO_SEGUNDOS  = 60
SCOPES              = ["https://www.googleapis.com/auth/drive.file"]


# ──────────────────────────────────────────────
# HELPERS — REGISTRO LOCAL
# ──────────────────────────────────────────────
def cargar_registro() -> dict:
    if os.path.exists(REGISTRO_SUBIDAS):
        with open(REGISTRO_SUBIDAS, "r") as f:
            return json.load(f)
    return {}


def guardar_registro(registro: dict):
    with open(REGISTRO_SUBIDAS, "w") as f:
        json.dump(registro, f, indent=2)


def extraer_folder_id(valor: str) -> str:
    """Acepta URL completa o ID puro, devuelve solo el ID."""
    match = re.search(r"/folders/([a-zA-Z0-9_-]+)", valor)
    return match.group(1) if match else valor.strip()


# ──────────────────────────────────────────────
# GOOGLE DRIVE — AUTENTICACIÓN Y SUBIDA
# ──────────────────────────────────────────────
def autenticar() -> Credentials:
    """
    Carga el token guardado si existe y es válido.
    Si expiró lo renueva. Si no existe abre el navegador (solo la 1ª vez).
    """
    creds = None

    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("   🔄 Token expirado, renovando...")
            creds.refresh(Request())
        else:
            if not os.path.exists(OAUTH_CLIENT_FILE):
                print(f"\n❌ No se encontró: {OAUTH_CLIENT_FILE}")
                print("   Descarga el JSON OAuth desde Google Cloud Console.")
                sys.exit(1)
            print("\n🌐 Primera vez: se abrirá el navegador para autorizar Drive.")
            flow = InstalledAppFlow.from_client_secrets_file(OAUTH_CLIENT_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())
        print(f"   ✅ Token guardado en {TOKEN_FILE}")

    return creds


def conectar_drive():
    creds   = autenticar()
    service = build("drive", "v3", credentials=creds)
    return service


def subir_imagen_desde_bytes(service, nombre: str, imagen_bytes: bytes, folder_id: str) -> str | None:
    """
    Sube una imagen (en bytes) directamente a folder_id en Drive.
    NO crea subcarpetas. Devuelve el file_id o None si hubo error.
    """
    mime_type = mimetypes.guess_type(nombre)[0] or "application/octet-stream"
    metadata  = {"name": nombre, "parents": [folder_id]}
    media     = MediaIoBaseUpload(io.BytesIO(imagen_bytes), mimetype=mime_type, resumable=True)

    try:
        archivo = service.files().create(
            body=metadata, media_body=media, fields="id"
        ).execute()
        return archivo.get("id")
    except HttpError as e:
        print(f"      ❌ Error subiendo {nombre}: {e}")
        return None


# ──────────────────────────────────────────────
# CORREO IMAP
# ──────────────────────────────────────────────
def conectar_imap():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_USER, EMAIL_PASS)
        return mail
    except Exception as e:
        print(f"   ❌ Error IMAP: {e}")
        return None


def buscar_correos_nuevos(mail: imaplib.IMAP4_SSL) -> list:
    """Devuelve UIDs de correos no leídos con el asunto trigger."""
    mail.select("INBOX")
    _, data = mail.search(None, f'(UNSEEN SUBJECT "{ASUNTO_TRIGGER}")')
    ids = data[0].split()
    return ids


def extraer_imagenes_de_correo(mail: imaplib.IMAP4_SSL, uid: bytes) -> list[tuple[bytes, str]]:
    """
    Extrae todos los adjuntos de imagen de un correo.
    Devuelve lista de (bytes, nombre_archivo).
    """
    _, data = mail.fetch(uid, "(RFC822)")
    msg     = email.message_from_bytes(data[0][1])
    imagenes = []

    for part in msg.walk():
        filename = part.get_filename()
        if not filename:
            continue
        ext = Path(filename).suffix.lower()
        if ext in EXTENSIONES_VALIDAS:
            contenido = part.get_payload(decode=True)
            if contenido:
                imagenes.append((contenido, filename))

    return imagenes


def marcar_como_leido(mail: imaplib.IMAP4_SSL, uid: bytes):
    mail.store(uid, "+FLAGS", "\\Seen")


# ──────────────────────────────────────────────
# PROCESAMIENTO DE UN CORREO
# ──────────────────────────────────────────────
def procesar_correo(mail: imaplib.IMAP4_SSL, uid: bytes, service, folder_id: str, registro: dict):
    print(f"\n   📩 Procesando correo UID {uid.decode()}...")

    imagenes = extraer_imagenes_de_correo(mail, uid)

    if not imagenes:
        print("   ⚠️  Sin adjuntos de imagen — correo ignorado.")
        marcar_como_leido(mail, uid)
        return

    print(f"   🖼️  {len(imagenes)} imagen(es) encontrada(s).")

    os.makedirs(CARPETA_RESPALDO, exist_ok=True)
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    subidos  = 0
    errores  = 0

    for imagen_bytes, nombre in imagenes:

        # Evitar duplicados: si ya fue subida en un ciclo anterior, omitir
        if nombre in registro:
            print(f"      ♻️  {nombre} ya estaba en Drive, omitiendo.")
            continue

        # Subir a Drive (directo a la carpeta raíz, sin subcarpetas)
        print(f"      ⬆️  {nombre}...", end=" ", flush=True)
        file_id = subir_imagen_desde_bytes(service, nombre, imagen_bytes, folder_id)

        if file_id:
            # Guardar respaldo local
            ruta_respaldo = os.path.join(CARPETA_RESPALDO, f"{ts}_{nombre}")
            with open(ruta_respaldo, "wb") as f:
                f.write(imagen_bytes)

            # Actualizar registro
            registro[nombre] = {
                "file_id":   file_id,
                "subido_en": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            guardar_registro(registro)
            print("✅")
            subidos += 1
        else:
            errores += 1

    # Marcar correo como leído solo si al menos una imagen fue procesada sin error
    if errores == 0:
        marcar_como_leido(mail, uid)
        print(f"   ✉️  Correo marcado como leído.")
    else:
        print(f"   ⚠️  Hubo {errores} error(es) — el correo se reintentará en el próximo ciclo.")

    return subidos, errores


# ──────────────────────────────────────────────
# CICLO PRINCIPAL
# ──────────────────────────────────────────────
def revisar_una_vez(service, folder_id: str, registro: dict):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 🔍 Revisando bandeja...")

    mail = conectar_imap()
    if not mail:
        print("   Sin conexión IMAP. Se reintentará en el próximo ciclo.")
        return

    try:
        uids = buscar_correos_nuevos(mail)
        if not uids:
            print("   Sin correos nuevos con el asunto trigger.")
        else:
            print(f"   📬 {len(uids)} correo(s) encontrado(s).")
            for uid in uids:
                procesar_correo(mail, uid, service, folder_id, registro)
    finally:
        try:
            mail.logout()
        except Exception:
            pass


def main():
    parser = argparse.ArgumentParser(description="sync_qr_drive — Sube imágenes QR de correo a Drive")
    parser.add_argument("--once", action="store_true", help="Revisar una sola vez y salir")
    args = parser.parse_args()

    folder_id = extraer_folder_id(DRIVE_FOLDER_ID)
    registro  = cargar_registro()

    print("=" * 57)
    print("  SYNC_QR_DRIVE — Correo → Imágenes QR → Google Drive")
    print(f"  Trigger   : {ASUNTO_TRIGGER}")
    print(f"  Drive ID  : {folder_id}")
    print(f"  Intervalo : {INTERVALO_SEGUNDOS}s")
    print("=" * 57)

    # Conexión a Drive (una sola vez, se reutiliza en todos los ciclos)
    print("\n🔗 Conectando a Google Drive...")
    service = conectar_drive()
    print("   ✅ Conexión exitosa\n")

    if args.once:
        revisar_una_vez(service, folder_id, registro)
        return

    print("▶️  Corriendo en modo continuo. Ctrl+C para detener.\n")
    try:
        while True:
            revisar_una_vez(service, folder_id, registro)
            print(f"   ⏳ Próxima revisión en {INTERVALO_SEGUNDOS}s...\n")
            time.sleep(INTERVALO_SEGUNDOS)
    except KeyboardInterrupt:
        print("\n🛑 Detenido por el usuario.")


if __name__ == "__main__":
    main()