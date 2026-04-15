"""
upload_drive.py
===============
Sube automáticamente todas las imágenes QR y códigos de barras
a una carpeta de Google Drive usando OAuth 2.0 (tu cuenta personal).

A diferencia de las cuentas de servicio, OAuth usa la cuota de
almacenamiento de tu propia cuenta de Google, por lo que no hay
límite de espacio más allá del de tu Drive.

──────────────────────────────────────────────────────────────────
CONFIGURACIÓN INICIAL (hacer UNA sola vez en tu PC/Kali):
──────────────────────────────────────────────────────────────────
1. Ve a https://console.cloud.google.com
2. Crea o selecciona un proyecto → activa "Google Drive API"
3. Ve a "APIs y servicios" → "Credenciales"
4. Clic en "Crear credenciales" → "ID de cliente OAuth 2.0"
5. Tipo de aplicación: "Aplicación de escritorio"
6. Descarga el JSON → renómbralo  oauth_client.json
7. Colócalo en la misma carpeta que este script
8. En DRIVE_FOLDER_ID pega el ID de tu carpeta de Drive
   (la parte final de la URL: /folders/ESTE_ID)

La primera vez que ejecutes el script abrirá el navegador para
que aceptes los permisos. Después guarda un token.json y ya no
vuelve a pedir autorización.
──────────────────────────────────────────────────────────────────
"""

import sys
import subprocess
import os

# ──────────────────────────────────────────────
# AUTO-INSTALACIÓN
# ──────────────────────────────────────────────
def instalar_dependencias():
    requeridos = {
        "google.oauth2.credentials":       "google-auth",
        "google_auth_oauthlib.flow":       "google-auth-oauthlib",
        "googleapiclient.discovery":       "google-api-python-client",
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
import json
import re
import mimetypes
import argparse
from datetime import datetime
from pathlib import Path

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError

# ──────────────────────────────────────────────
# CONFIGURACIÓN  ← editar estos valores
# ──────────────────────────────────────────────
# Ruta al JSON OAuth descargado de Google Cloud Console
OAUTH_CLIENT_FILE = "/home/christian/Documentos/proyectos_bi/oauth_client.json"

# Token que se genera automáticamente tras el primer login (no editar)
TOKEN_FILE        = "/home/christian/Documentos/proyectos_bi/token.json"

# ID de la carpeta destino en Drive (solo el ID, no la URL completa)
DRIVE_FOLDER_ID   = "1dyPhza4IcbsTCrhjOVe50arLK39GYziK"

# Carpeta local con las imágenes
CARPETA_LOCAL     = "/home/christian/Documentos/proyectos_bi/codigos_generados"

# Extensiones a subir
EXTENSIONES_VALIDAS = {".png", ".jpg", ".jpeg"}

# Registro local de archivos ya subidos (evita duplicados)
REGISTRO_SUBIDAS  = "/home/christian/Documentos/proyectos_bi/drive_subidas.json"

# Permisos mínimos necesarios
SCOPES = ["https://www.googleapis.com/auth/drive.file"]


# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────
def extraer_folder_id(valor: str) -> str:
    """Acepta URL completa o ID puro, devuelve solo el ID."""
    match = re.search(r"/folders/([a-zA-Z0-9_-]+)", valor)
    return match.group(1) if match else valor.strip()


def cargar_registro() -> dict:
    if os.path.exists(REGISTRO_SUBIDAS):
        with open(REGISTRO_SUBIDAS, "r") as f:
            return json.load(f)
    return {}


def guardar_registro(registro: dict):
    with open(REGISTRO_SUBIDAS, "w") as f:
        json.dump(registro, f, indent=2)


# ──────────────────────────────────────────────
# AUTENTICACIÓN OAuth 2.0
# ──────────────────────────────────────────────
def autenticar() -> Credentials:
    """
    Carga el token guardado si existe y es válido.
    Si expiró lo renueva automáticamente.
    Si no existe abre el navegador para pedir autorización (solo la 1ª vez).
    """
    creds = None

    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    # Si no hay token o está vencido
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("   🔄 Token expirado, renovando automáticamente...")
            creds.refresh(Request())
        else:
            if not os.path.exists(OAUTH_CLIENT_FILE):
                print(f"\n❌ No se encontró: {OAUTH_CLIENT_FILE}")
                print("   Descarga el JSON OAuth desde Google Cloud Console")
                print("   y guárdalo como oauth_client.json en la carpeta del proyecto.")
                sys.exit(1)
            print("\n🌐 Primera vez: se abrirá el navegador para autorizar acceso a Drive.")
            print("   Inicia sesión con tu cuenta de Google y acepta los permisos.\n")
            flow = InstalledAppFlow.from_client_secrets_file(OAUTH_CLIENT_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        # Guardar token para usos futuros
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())
        print(f"   ✅ Token guardado en {TOKEN_FILE}")

    return creds


def conectar_drive():
    print("\n🔗 Conectando a Google Drive (OAuth)...")
    creds   = autenticar()
    service = build("drive", "v3", credentials=creds)
    print("   ✅ Conexión exitosa\n")
    return service


# ──────────────────────────────────────────────
# GESTIÓN DE SUBCARPETAS POR FECHA
# ──────────────────────────────────────────────
def obtener_o_crear_subcarpeta(service, nombre: str, parent_id: str) -> str:
    query = (
        f"name='{nombre}' and "
        f"'{parent_id}' in parents and "
        f"mimeType='application/vnd.google-apps.folder' and "
        f"trashed=false"
    )
    results = service.files().list(q=query, fields="files(id)").execute()
    items   = results.get("files", [])
    if items:
        return items[0]["id"]

    metadata = {
        "name":     nombre,
        "mimeType": "application/vnd.google-apps.folder",
        "parents":  [parent_id],
    }
    carpeta = service.files().create(body=metadata, fields="id").execute()
    print(f"   📁 Subcarpeta creada: {nombre}")
    return carpeta["id"]


# ──────────────────────────────────────────────
# SUBIDA DE UN ARCHIVO
# ──────────────────────────────────────────────
def subir_archivo(service, ruta_local: str, folder_id: str) -> str | None:
    nombre    = os.path.basename(ruta_local)
    mime_type = mimetypes.guess_type(ruta_local)[0] or "application/octet-stream"
    metadata  = {"name": nombre, "parents": [folder_id]}
    media     = MediaFileUpload(ruta_local, mimetype=mime_type, resumable=True)

    try:
        archivo = service.files().create(
            body=metadata, media_body=media, fields="id"
        ).execute()
        return archivo.get("id")
    except HttpError as e:
        print(f"\n   ❌ Error subiendo {nombre}: {e}")
        return None


# ──────────────────────────────────────────────
# PROCESO PRINCIPAL
# ──────────────────────────────────────────────
def subir_imagenes(carpeta_local: str = CARPETA_LOCAL, archivos_especificos: list = None):
    folder_id = extraer_folder_id(DRIVE_FOLDER_ID)

    print("\n" + "=" * 52)
    print("  UPLOAD DRIVE — Subiendo imágenes de inventario")
    print(f"  Carpeta Drive ID : {folder_id}")
    print("=" * 52)

    if not os.path.isdir(carpeta_local):
        print(f"❌ Carpeta local no encontrada: {carpeta_local}")
        return

    service  = conectar_drive()
    registro = cargar_registro()

    # Determinar archivos a procesar
    if archivos_especificos:
        candidatos = [Path(p) for p in archivos_especificos
                      if Path(p).suffix.lower() in EXTENSIONES_VALIDAS]
    else:
        candidatos = [p for p in Path(carpeta_local).iterdir()
                      if p.is_file() and p.suffix.lower() in EXTENSIONES_VALIDAS]

    pendientes = [p for p in candidatos if p.name not in registro]

    if not pendientes:
        print(f"✅ Sin archivos nuevos ({len(registro)} ya estaban en Drive).")
        return

    print(f"📂 Archivos pendientes: {len(pendientes)}")

    # Agrupar por fecha de modificación
    grupos: dict[str, list] = {}
    for ruta in pendientes:
        fecha = datetime.fromtimestamp(ruta.stat().st_mtime).strftime("%Y-%m-%d")
        grupos.setdefault(fecha, []).append(ruta)

    subidos = 0
    errores = 0

    for fecha, archivos in sorted(grupos.items()):
        print(f"\n   📅 {fecha}  ({len(archivos)} archivos)")
        sub_id = obtener_o_crear_subcarpeta(service, fecha, folder_id)

        for ruta in archivos:
            print(f"      ⬆️  {ruta.name}...", end=" ", flush=True)
            file_id = subir_archivo(service, str(ruta), sub_id)

            if file_id:
                registro[ruta.name] = {
                    "file_id":   file_id,
                    "fecha":     fecha,
                    "subido_en": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
                guardar_registro(registro)
                print("✅")
                subidos += 1
            else:
                errores += 1

    print(f"\n{'='*52}")
    print(f"  ✅ Subidos   : {subidos}")
    print(f"  ❌ Errores   : {errores}")
    print(f"  ♻️  En Drive  : {len(registro) - subidos}")
    print(f"\n🔗 https://drive.google.com/drive/folders/{folder_id}")
    print("=" * 52)


# ──────────────────────────────────────────────
# ENTRADA
# ──────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Sube imágenes QR/barras a Google Drive")
    parser.add_argument("--carpeta",   default=CARPETA_LOCAL)
    parser.add_argument("--archivos",  nargs="+")
    args = parser.parse_args()

    subir_imagenes(
        carpeta_local=args.carpeta,
        archivos_especificos=args.archivos,
    )