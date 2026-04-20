"""
CodigoQR.py  (Termux) v2.0
==========================
Sistema mejorado de captura de inventario con soporte para:
  - CPU (con software y periféricos)
  - Equipos individuales (sillas, mesas, micrófonos, etc.)
  - Periféricos individuales

Flujo:
  1. Menú inicial: seleccionar tipo de equipamiento
  2. Si CPU → formulario completo (hardware + software + periféricos)
  3. Si otro → formulario simplificado
  4. Genera QR único para CPU, código de barras para cada objeto
  5. Envía xlsx con asunto trigger a sync_bd.py
"""

import sys
import subprocess
import os

# ──────────────────────────────────────────────
# AUTO-INSTALACIÓN (Termux)
# ──────────────────────────────────────────────
def instalar_paquetes_sistema():
    paquetes = ["libjpeg-turbo", "libpng", "zlib", "python"]
    print("📦 Verificando paquetes del sistema Termux...")
    for paquete in paquetes:
        try:
            r = subprocess.run(["pkg", "install", "-y", paquete], capture_output=True, text=True)
            print(f"   {'✅' if r.returncode == 0 else '⚠️ '} {paquete}")
        except FileNotFoundError:
            break

def instalar_pip_paquetes():
    requeridos = {
        "pandas":   "pandas",
        "qrcode":   "qrcode[pil]",
        "barcode":  "python-barcode",
        "openpyxl": "openpyxl",
        "Pillow":   "Pillow",
    }
    print("\n🐍 Verificando paquetes Python...")
    faltantes = []
    for modulo, paquete in requeridos.items():
        try:
            __import__(modulo)
            print(f"   ✅ {modulo}")
        except ImportError:
            print(f"   ⬇️  {modulo} → se instalará")
            faltantes.append(paquete)

    if faltantes:
        print(f"\n⚙️  Instalando {len(faltantes)} paquete(s)...")
        for paquete in faltantes:
            print(f"   {paquete}...", end=" ", flush=True)
            r = subprocess.run(
                [sys.executable, "-m", "pip", "install", "--quiet", paquete],
                capture_output=True, text=True
            )
            print("✅" if r.returncode == 0 else f"❌ {r.stderr.strip()}")
            if r.returncode != 0:
                sys.exit(1)
        print("\n✅ Dependencias listas.\n")
    else:
        print("   Sin instalaciones pendientes.\n")

instalar_paquetes_sistema()
instalar_pip_paquetes()

# ──────────────────────────────────────────────
# IMPORTS
# ──────────────────────────────────────────────
import smtplib
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
import json

import pandas as pd
import qrcode
import barcode
from barcode.writer import ImageWriter

# ──────────────────────────────────────────────
# CONFIGURACIÓN
# ──────────────────────────────────────────────
archivo_excel  = '/home/christian/Documentos/proyectos_bi/inventario.xlsx'
carpeta_salida = '/home/christian/Documentos/proyectos_bi/codigos_generados'

FECHA_HOY      = datetime.now().strftime("%Y%m%d")
ARCHIVO_SESION = f"INVENTARIO{FECHA_HOY}.xlsx"

# ── Correo ────────────────────────────────────
SMTP_SERVER    = "smtp.gmail.com"
SMTP_PORT      = 465
EMAIL_USER     = "reportes.bi@abcsc.mx"
EMAIL_PASS     = "jwvjdrvmprzrwzxy"

ASUNTO_TRIGGER = "ACTUALIZACION_BASE_DE_DATOS"

DESTINATARIOS  = [
    "ccarbajal@abcsc.mx",
#    "myanez@abcsc.mx",
#    "sgonzalez@abcsc.mx",
#    "ymontoya@abcsc.mx",
    "reportes.bi@abcsc.mx",
]

# Directorios base
EDIFICIOS = ["NORTE 180", "NORTE 182", "PATIO SEC"]
ESTADOS = ["1. Mal estado", "2. Futuro mantenimiento", "3. Buen estado", "4. Equipo nuevo"]
TIPOS_DISCO = ["HDD", "SSD"]
PERIFERICOS_DISPONIBLES = ["Monitor", "Teclado", "Mouse", "Webcam", "Auriculares", "Micrófono"]
TIPOS_EQUIPAMIENTO = [
    "CPU",
    "Silla",
    "Mesa",
    "Micrófono",
    "Periférico Individual",
    "Finalizar Inventario"
]

os.makedirs(carpeta_salida, exist_ok=True)


# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────
def extraer_tres_letras(nombre: str) -> str:
    """Extrae 3 letras representativas del nombre."""
    limpio = nombre.replace(" ", "").upper()
    if len(limpio) == 0: 
        return "XXX"
    if len(limpio) == 1: 
        return limpio * 3
    if len(limpio) == 2: 
        return limpio + "X"
    return f"{limpio[0]}{limpio[len(limpio) // 2]}{limpio[-1]}"


def seleccionar_opcion(lista: list, titulo: str = None, permitir_numeros=True) -> str:
    """
    Menú de selección mejorado.
    Si permitir_numeros=True, devuelve el item; si no, devuelve el índice como string.
    """
    if titulo:
        print(f"\n{titulo}")
    for i, item in enumerate(lista, 1):
        print(f"   {i}. {item}")
    while True:
        try:
            idx = int(input("\n   Selecciona número: ")) - 1
            if 0 <= idx < len(lista):
                return lista[idx] if permitir_numeros else str(idx)
            print(f"   ⚠️  Elige entre 1 y {len(lista)}.")
        except ValueError:
            print("   ⚠️  Ingresa un número válido.")


def input_seguro(prompt: str, permitir_vacio=False) -> str:
    """Input con validación."""
    while True:
        valor = input(prompt).strip()
        if valor or permitir_vacio:
            return valor
        print("   ⚠️  Campo requerido.")


def numero_entero(prompt: str, minimo=1) -> int:
    """Input que garantiza un número entero."""
    while True:
        try:
            valor = int(input(prompt))
            if valor >= minimo:
                return valor
            print(f"   ⚠️  Debe ser >= {minimo}.")
        except ValueError:
            print("   ⚠️  Ingresa un número válido.")


# ──────────────────────────────────────────────
# CAPTURA POR TIPO DE EQUIPAMIENTO
# ──────────────────────────────────────────────

def capturar_cpu(empresas: list, areas: list) -> dict:
    """Captura datos completos de una CPU con software y periféricos."""
    print("\n" + "=" * 55)
    print("  REGISTRO CPU — HARDWARE + SOFTWARE + PERIFÉRICOS")
    print(f"  {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 55)

    # ─────────────────────────────────────
    # SECCIÓN 1: DATOS BÁSICOS CPU
    # ─────────────────────────────────────
    print("\n🖥️  HARDWARE CPU")
    print("─" * 55)
    
    host = input_seguro("Host: ")
    no_serie = input_seguro("No. de Serie: ")
    empresa = seleccionar_opcion(empresas, "🏢 Empresa:")
    edificio = seleccionar_opcion(EDIFICIOS, "🏗️  Edificio:")
    area = seleccionar_opcion(areas, "📍 Área:")
    
    print("\n📊 Estado Físico:")
    estado_idx = seleccionar_opcion(ESTADOS, permitir_numeros=False)
    estado_map = {
        "0": "Mal estado",
        "1": "Futuro mantenimiento",
        "2": "Buen estado",
        "3": "Equipo nuevo"
    }
    estado = estado_map.get(estado_idx, "Buen estado")
    
    marca = input_seguro("Marca: ")
    modelo = input_seguro("Modelo: ")
    procesador = input_seguro("Procesador (ej: Intel i7): ")
    ram = input_seguro("RAM (ej: 16GB): ")
    capacidad_disco = input_seguro("Capacidad Disco (ej: 512GB): ")
    
    print("\n💾 Tipo de Disco Duro:")
    tipo_disco = seleccionar_opcion(TIPOS_DISCO)
    
    observaciones_hw = input_seguro("Observaciones Hardware: ", permitir_vacio=True)

    # ─────────────────────────────────────
    # SECCIÓN 2: SOFTWARE
    # ─────────────────────────────────────
    print("\n\n📦 SOFTWARE INSTALADO")
    print("─" * 55)
    
    so = input_seguro("Sistema Operativo (ej: Windows 10): ")
    office = input_seguro("Office (ej: Office 2021, LibreOffice, Ninguno): ")
    antivirus = input_seguro("Antivirus (ej: Windows Defender, Norton, Ninguno): ")
    lector_pdf = input_seguro("Lector de PDF (ej: Adobe Reader, Foxit, Ninguno): ")
    erp = input_seguro("ERP (ej: SAP, Oracle, Ninguno): ")
    
    print("\n📝 Otros Software:")
    otro1 = input_seguro("Otro Software 1: ", permitir_vacio=True)
    otro2 = input_seguro("Otro Software 2: ", permitir_vacio=True)
    otro3 = input_seguro("Otro Software 3: ", permitir_vacio=True)

    software = {
        "SO": so,
        "Office": office,
        "Antivirus": antivirus,
        "Lector_PDF": lector_pdf,
        "ERP": erp,
        "Otro1": otro1,
        "Otro2": otro2,
        "Otro3": otro3,
    }

    # ─────────────────────────────────────
    # SECCIÓN 3: PERIFÉRICOS
    # ─────────────────────────────────────
    print("\n\n🖱️  PERIFÉRICOS")
    print("─" * 55)
    
    perifericos = []
    perifericos_capturados = set()

    # Monitor
    print("\n📺 MONITOR:")
    agregarmon = input("¿Registrar monitor? (S/N): ").strip().upper()
    if agregarmon == "S":
        monitor = {
            "tipo": "Monitor",
            "modelo": input_seguro("  Modelo: "),
            "no_serie": input_seguro("  No. de Serie: "),
            "marca": input_seguro("  Marca: "),
            "estado": seleccionar_opcion(ESTADOS, "  Estado Físico:"),
            "observaciones": input_seguro("  Observaciones: ", permitir_vacio=True),
        }
        perifericos.append(monitor)
        perifericos_capturados.add("Monitor")

    # Teclado
    print("\n⌨️  TECLADO:")
    agrega_teclado = input("¿Registrar teclado? (S/N): ").strip().upper()
    if agrega_teclado == "S":
        teclado = {
            "tipo": "Teclado",
            "modelo": input_seguro("  Modelo: "),
            "no_serie": input_seguro("  No. de Serie: ", permitir_vacio=True),
            "marca": input_seguro("  Marca: "),
            "estado": seleccionar_opcion(ESTADOS, "  Estado Físico:"),
            "observaciones": input_seguro("  Observaciones: ", permitir_vacio=True),
        }
        perifericos.append(teclado)
        perifericos_capturados.add("Teclado")

    # Mouse
    print("\n🖱️  MOUSE:")
    agrega_mouse = input("¿Registrar mouse? (S/N): ").strip().upper()
    if agrega_mouse == "S":
        mouse = {
            "tipo": "Mouse",
            "modelo": input_seguro("  Modelo: "),
            "no_serie": input_seguro("  No. de Serie: ", permitir_vacio=True),
            "marca": input_seguro("  Marca: "),
            "estado": seleccionar_opcion(ESTADOS, "  Estado Físico:"),
            "observaciones": input_seguro("  Observaciones: ", permitir_vacio=True),
        }
        perifericos.append(mouse)
        perifericos_capturados.add("Mouse")

    # Periféricos adicionales
    while True:
        perifericos_faltantes = [p for p in PERIFERICOS_DISPONIBLES if p not in perifericos_capturados]
        if not perifericos_faltantes:
            break
        
        agrega_mas = input("\n¿Agregar más periféricos? (S/N): ").strip().upper()
        if agrega_mas != "S":
            break
        
        print("\n📌 Periféricos disponibles:")
        tipo_periferico = seleccionar_opcion(perifericos_faltantes)
        
        periferico = {
            "tipo": tipo_periferico,
            "modelo": input_seguro(f"  Modelo de {tipo_periferico}: "),
            "no_serie": input_seguro(f"  No. de Serie: ", permitir_vacio=True),
            "marca": input_seguro(f"  Marca: "),
            "estado": seleccionar_opcion(ESTADOS, f"  Estado Físico de {tipo_periferico}:"),
            "observaciones": input_seguro(f"  Observaciones: ", permitir_vacio=True),
        }
        perifericos.append(periferico)
        perifericos_capturados.add(tipo_periferico)

    # ─────────────────────────────────────
    # ARMAR ESTRUCTURA RETORNO
    # ─────────────────────────────────────
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    estructura_cpu = {
        "tipo_inventario": "CPU",
        "timestamp": timestamp,
        "cpu": {
            "host": host,
            "no_serie": no_serie,
            "empresa": empresa,
            "edificio": edificio,
            "area": area,
            "estado": estado,
            "marca": marca,
            "modelo": modelo,
            "procesador": procesador,
            "ram": ram,
            "capacidad_disco": capacidad_disco,
            "tipo_disco": tipo_disco,
            "observaciones": observaciones_hw,
        },
        "software": software,
        "perifericos": perifericos,
    }

    print(f"\n✅ CPU '{host}' registrada con {len(perifericos)} periférico(s).")
    return estructura_cpu


def capturar_equipo_simple(empresas: list, areas: list) -> dict:
    """Captura datos simplificados para otros tipos de equipamiento."""
    
    tipo = seleccionar_opcion(TIPOS_EQUIPAMIENTO[1:], "¿Qué tipo de equipamiento?")
    
    print("\n" + "=" * 45)
    print(f"  REGISTRO: {tipo.upper()}")
    print(f"  {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 45)

    nombre = input_seguro(f"\nNombre/Descripción del {tipo}: ")
    no_serie = input_seguro("No. de Serie: ", permitir_vacio=True)
    marca = input_seguro("Marca: ", permitir_vacio=True)
    modelo = input_seguro("Modelo: ", permitir_vacio=True)
    
    empresa = seleccionar_opcion(empresas, "🏢 Empresa:")
    edificio = seleccionar_opcion(EDIFICIOS, "🏗️  Edificio:")
    area = seleccionar_opcion(areas, "📍 Área:")
    
    print("\n📊 Estado Físico:")
    estado_idx = seleccionar_opcion(ESTADOS, permitir_numeros=False)
    estado_map = {
        "0": "Mal estado",
        "1": "Futuro mantenimiento",
        "2": "Buen estado",
        "3": "Equipo nuevo"
    }
    estado = estado_map.get(estado_idx, "Buen estado")
    
    observaciones = input_seguro("Observaciones: ", permitir_vacio=True)

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    estructura = {
        "tipo_inventario": tipo,
        "timestamp": timestamp,
        "datos": {
            "nombre": nombre,
            "no_serie": no_serie,
            "marca": marca,
            "modelo": modelo,
            "empresa": empresa,
            "edificio": edificio,
            "area": area,
            "estado": estado,
            "observaciones": observaciones,
        }
    }

    print(f"\n✅ {tipo} '{nombre}' registrado.")
    return estructura


# ──────────────────────────────────────────────
# GENERACIÓN DE CÓDIGOS (QR y BARRAS)
# ──────────────────────────────────────────────

def generar_qr_cpu(estructura_cpu: dict, numero_correlativo: int) -> tuple[str, str]:
    """
    Genera UN SOLO QR para toda la estructura CPU.
    Devuelve (ruta_qr, contenido_qr_json)
    """
    # Crear contenido JSON comprimido
    contenido_json = json.dumps({
        "tipo": "CPU",
        "host": estructura_cpu["cpu"]["host"],
        "no_serie": estructura_cpu["cpu"]["no_serie"],
        "empresa": estructura_cpu["cpu"]["empresa"],
        "procesador": estructura_cpu["cpu"]["procesador"],
        "ram": estructura_cpu["cpu"]["ram"],
        "timestamp": estructura_cpu["timestamp"],
    })

    codigo_id = f"CPU{numero_correlativo:04d}"
    ruta_qr = os.path.join(carpeta_salida, f"qr_{codigo_id}.png")
    
    qrcode.make(contenido_json).save(ruta_qr)
    print(f"   📷 QR CPU    → {ruta_qr}")
    
    return ruta_qr, contenido_json, codigo_id


def generar_barras_objeto(tipo: str, no_serie: str, numero_correlativo: int) -> str:
    """Genera un código de barras para cada objeto (CPU, periférico, etc)."""
    codigo_id = f"{extraer_tres_letras(tipo)}{numero_correlativo:04d}"
    ruta_barras_base = os.path.join(carpeta_salida, f"barras_{codigo_id}")
    
    bn_class = barcode.get_barcode_class("code128")
    ruta_barras_final = bn_class(codigo_id, writer=ImageWriter()).save(ruta_barras_base)
    
    print(f"   📊 {tipo:15} → {ruta_barras_final}")
    return ruta_barras_final, codigo_id


def generar_barras_simple(nombre: str, numero_correlativo: int) -> str:
    """Genera código de barras para equipamiento simple."""
    abrev = extraer_tres_letras(nombre)
    codigo_id = f"{abrev}{numero_correlativo:04d}"
    ruta_barras_base = os.path.join(carpeta_salida, f"barras_{codigo_id}")
    
    bn_class = barcode.get_barcode_class("code128")
    ruta_barras_final = bn_class(codigo_id, writer=ImageWriter()).save(ruta_barras_base)
    
    print(f"   📊 Barras → {ruta_barras_final}")
    return ruta_barras_final, codigo_id


# ──────────────────────────────────────────────
# GUARDAR EN EXCEL
# ──────────────────────────────────────────────

def guardar_estructura_cpu_en_excel(estructura_cpu: dict, codigo_qr: str, 
                                    codigo_barras_cpu: str, codigos_perifericos: dict):
    """
    Guarda la estructura CPU en varias hojas de Excel:
      - CPU: datos de la CPU
      - Periféricos: datos de cada periférico
      - Software: software instalado
      - Relaciones: vinculación entre código QR, barras CPU, y barras periféricos
    """
    
    excel_path = ARCHIVO_SESION
    
    # Crear diccionarios para cada hoja
    fila_cpu = {
        "Tipo": "CPU",
        "Host": estructura_cpu["cpu"]["host"],
        "No_Serie": estructura_cpu["cpu"]["no_serie"],
        "Empresa": estructura_cpu["cpu"]["empresa"],
        "Edificio": estructura_cpu["cpu"]["edificio"],
        "Area": estructura_cpu["cpu"]["area"],
        "Estado": estructura_cpu["cpu"]["estado"],
        "Marca": estructura_cpu["cpu"]["marca"],
        "Modelo": estructura_cpu["cpu"]["modelo"],
        "Procesador": estructura_cpu["cpu"]["procesador"],
        "RAM": estructura_cpu["cpu"]["ram"],
        "Capacidad_Disco": estructura_cpu["cpu"]["capacidad_disco"],
        "Tipo_Disco": estructura_cpu["cpu"]["tipo_disco"],
        "Observaciones": estructura_cpu["cpu"]["observaciones"],
        "Codigo_QR": codigo_qr,
        "Codigo_Barras_CPU": codigo_barras_cpu,
        "Timestamp": estructura_cpu["timestamp"],
    }
    
    df_cpu = pd.DataFrame([fila_cpu])
    
    # DataFrame para Software
    fila_software = {
        "Host_CPU": estructura_cpu["cpu"]["host"],
        "SO": estructura_cpu["software"]["SO"],
        "Office": estructura_cpu["software"]["Office"],
        "Antivirus": estructura_cpu["software"]["Antivirus"],
        "Lector_PDF": estructura_cpu["software"]["Lector_PDF"],
        "ERP": estructura_cpu["software"]["ERP"],
        "Otro_1": estructura_cpu["software"]["Otro1"],
        "Otro_2": estructura_cpu["software"]["Otro2"],
        "Otro_3": estructura_cpu["software"]["Otro3"],
        "Timestamp": estructura_cpu["timestamp"],
    }
    
    df_software = pd.DataFrame([fila_software])
    
    # DataFrame para Periféricos
    filas_perifericos = []
    for periferico in estructura_cpu["perifericos"]:
        tipo_periferico = periferico["tipo"]
        codigo_barras = codigos_perifericos.get(tipo_periferico, {}).get("barras", "")
        codigo_id = codigos_perifericos.get(tipo_periferico, {}).get("codigo_id", "")
        
        fila = {
            "Host_CPU": estructura_cpu["cpu"]["host"],
            "Tipo": tipo_periferico,
            "Modelo": periferico["modelo"],
            "No_Serie": periferico["no_serie"],
            "Marca": periferico["marca"],
            "Estado": periferico["estado"],
            "Observaciones": periferico["observaciones"],
            "Codigo_Barras": codigo_barras,
            "Codigo_ID": codigo_id,
            "Timestamp": estructura_cpu["timestamp"],
        }
        filas_perifericos.append(fila)
    
    df_perifericos = pd.DataFrame(filas_perifericos) if filas_perifericos else pd.DataFrame()
    
    # DataFrame para Relaciones (vinculación QR-Barras)
    filas_relaciones = [{"Codigo_QR": codigo_qr, "Codigo_Barras_CPU": codigo_barras_cpu}]
    for tipo_periferico, datos in codigos_perifericos.items():
        filas_relaciones.append({
            "Codigo_QR": codigo_qr,
            "Tipo_Periferico": tipo_periferico,
            "Codigo_Barras_Periferico": datos.get("barras", "")
        })
    df_relaciones = pd.DataFrame(filas_relaciones)
    
    # Escribir a Excel
    try:
        if os.path.exists(excel_path):
            # Si el archivo existe, leer las hojas existentes y concatenar
            with pd.ExcelWriter(excel_path, mode="a", engine="openpyxl", 
                              if_sheet_exists="overlay") as writer:
                # CPU
                if "CPU" in writer.book.sheetnames:
                    df_existente = pd.read_excel(excel_path, sheet_name="CPU")
                    df_cpu = pd.concat([df_existente, df_cpu], ignore_index=True)
                df_cpu.to_excel(writer, sheet_name="CPU", index=False)
                
                # Software
                if not df_software.empty:
                    if "Software" in writer.book.sheetnames:
                        df_existente = pd.read_excel(excel_path, sheet_name="Software")
                        df_software = pd.concat([df_existente, df_software], ignore_index=True)
                    df_software.to_excel(writer, sheet_name="Software", index=False)
                
                # Periféricos
                if not df_perifericos.empty:
                    if "Perifericos" in writer.book.sheetnames:
                        df_existente = pd.read_excel(excel_path, sheet_name="Perifericos")
                        df_perifericos = pd.concat([df_existente, df_perifericos], ignore_index=True)
                    df_perifericos.to_excel(writer, sheet_name="Perifericos", index=False)
                
                # Relaciones
                if "Relaciones" in writer.book.sheetnames:
                    df_existente = pd.read_excel(excel_path, sheet_name="Relaciones")
                    df_relaciones = pd.concat([df_existente, df_relaciones], ignore_index=True)
                df_relaciones.to_excel(writer, sheet_name="Relaciones", index=False)
        else:
            # Si no existe, crear nuevo
            with pd.ExcelWriter(excel_path, mode="w", engine="openpyxl") as writer:
                df_cpu.to_excel(writer, sheet_name="CPU", index=False)
                
                if not df_software.empty:
                    df_software.to_excel(writer, sheet_name="Software", index=False)
                
                if not df_perifericos.empty:
                    df_perifericos.to_excel(writer, sheet_name="Perifericos", index=False)
                
                df_relaciones.to_excel(writer, sheet_name="Relaciones", index=False)
    except Exception as e:
        print(f"   ❌ Error guardando en Excel: {e}")
        import traceback
        traceback.print_exc()


def guardar_estructura_simple_en_excel(estructura: dict, codigo_barras: str, codigo_id: str):
    """Guarda equipamiento simple en Excel."""
    
    excel_path = ARCHIVO_SESION
    
    fila = {
        "Tipo": estructura["tipo_inventario"],
        "Nombre": estructura["datos"]["nombre"],
        "No_Serie": estructura["datos"]["no_serie"],
        "Marca": estructura["datos"]["marca"],
        "Modelo": estructura["datos"]["modelo"],
        "Empresa": estructura["datos"]["empresa"],
        "Edificio": estructura["datos"]["edificio"],
        "Area": estructura["datos"]["area"],
        "Estado": estructura["datos"]["estado"],
        "Observaciones": estructura["datos"]["observaciones"],
        "Codigo_Barras": codigo_barras,
        "Codigo_ID": codigo_id,
        "Timestamp": estructura["timestamp"],
    }
    
    df = pd.DataFrame([fila])
    
    try:
        if os.path.exists(excel_path):
            # Leer hoja "Otros" si existe
            try:
                df_existente = pd.read_excel(excel_path, sheet_name="Otros")
                df_final = pd.concat([df_existente, df], ignore_index=True)
            except:
                # Si la hoja no existe, solo usar el nuevo df
                df_final = df
            
            # Escribir usando overlay para no perder otras hojas
            with pd.ExcelWriter(excel_path, mode="a", engine="openpyxl", 
                              if_sheet_exists="overlay") as writer:
                df_final.to_excel(writer, sheet_name="Otros", index=False)
        else:
            # Si no existe, crear nuevo
            with pd.ExcelWriter(excel_path, mode="w", engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Otros", index=False)
    except Exception as e:
        print(f"   ❌ Error guardando en Excel: {e}")
        import traceback
        traceback.print_exc()


# ──────────────────────────────────────────────
# ENVÍO POR CORREO
# ──────────────────────────────────────────────

def enviar_correo(archivos_adjuntos: list, total_registros: int):
    """Envía el archivo xlsx con el asunto trigger."""
    print("\n📧 Enviando correo con datos de sesión...")
    hoy = datetime.now().strftime("%d/%m/%Y %H:%M")

    msg = EmailMessage()
    msg["Subject"] = ASUNTO_TRIGGER
    msg["From"] = EMAIL_USER
    msg["To"] = ", ".join(DESTINATARIOS)
    msg.set_content(
        f"Sesión de inventario: {hoy}\n"
        f"Registros capturados: {total_registros}\n"
        f"Archivo adjunto: {ARCHIVO_SESION}\n\n"
        f"Este correo será procesado automáticamente por sync_bd.py.\n"
        f"— Sistema BI · Inventario v2.0"
    )

    for archivo in archivos_adjuntos:
        if not os.path.exists(archivo):
            print(f"   ⚠️  No encontrado, omitido: {archivo}")
            continue
        with open(archivo, "rb") as f:
            datos = f.read()
        nombre = os.path.basename(archivo)
        ext = nombre.rsplit(".", 1)[-1].lower()
        
        if ext in ("png", "jpg", "jpeg"):
            msg.add_attachment(datos, maintype="image", subtype=ext, filename=nombre)
        elif ext == "xlsx":
            msg.add_attachment(
                datos, maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=nombre,
            )
        else:
            msg.add_attachment(datos, maintype="application", subtype="octet-stream", filename=nombre)

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.login(EMAIL_USER, EMAIL_PASS)
            smtp.send_message(msg)
        print(f"✅ Correo enviado con asunto [{ASUNTO_TRIGGER}]")
        print(f"   Destinatarios: {', '.join(DESTINATARIOS)}")
        return True
    except smtplib.SMTPAuthenticationError:
        print("❌ Error de autenticación SMTP.")
        return False
    except Exception as e:
        print(f"❌ Error al enviar: {e}")
        return False


# ──────────────────────────────────────────────
# FLUJO PRINCIPAL
# ──────────────────────────────────────────────

def ejecutar_sistema():
    """Flujo principal del sistema de captura."""
    
    archivos_sesion = []
    total_registros = 0
    numero_correlativo = 1

    print("\n" + "=" * 60)
    print("  SISTEMA DE INVENTARIO QR v2.0")
    print("  Termux — Captura de Equipamiento")
    print("=" * 60)

    # Cargar datos de Excel
    try:
        xls = pd.ExcelFile(archivo_excel)
        empresas = pd.read_excel(xls, "Empresa")["Empresa"].dropna().tolist()
        areas = pd.read_excel(xls, "Areas")["Area"].dropna().tolist()
    except FileNotFoundError:
        print("❌ No se encontró inventario.xlsx. Crea uno primero con las hojas:")
        print("   - Empresa")
        print("   - Areas")
        return
    except Exception as e:
        print(f"❌ Error cargando Excel: {e}")
        return

    # Bucle de captura
    while True:
        try:
            # Menú principal
            tipo_seleccionado = seleccionar_opcion(
                TIPOS_EQUIPAMIENTO,
                "\n🎯 ¿Qué deseas inventariar?"
            )

            if tipo_seleccionado == "CPU":
                # Flujo CPU
                estructura = capturar_cpu(empresas, areas)
                
                # Generar QR único para toda la estructura
                ruta_qr, contenido_qr, codigo_qr = generar_qr_cpu(estructura, numero_correlativo)
                archivos_sesion.append(ruta_qr)
                
                # Generar código de barras para la CPU
                ruta_barras_cpu, codigo_barras_cpu = generar_barras_objeto(
                    "CPU", estructura["cpu"]["no_serie"], numero_correlativo
                )
                archivos_sesion.append(ruta_barras_cpu)
                
                # Generar códigos de barras para cada periférico
                codigos_perifericos = {}
                for periferico in estructura["perifericos"]:
                    ruta_barras_periferico, codigo_id = generar_barras_objeto(
                        periferico["tipo"], 
                        periferico["no_serie"],
                        numero_correlativo
                    )
                    archivos_sesion.append(ruta_barras_periferico)
                    codigos_perifericos[periferico["tipo"]] = {
                        "barras": ruta_barras_periferico,
                        "codigo_id": codigo_id
                    }
                
                # Guardar en Excel
                guardar_estructura_cpu_en_excel(
                    estructura,
                    codigo_qr,
                    codigo_barras_cpu,
                    codigos_perifericos
                )
                
                total_registros += 1 + len(estructura["perifericos"])
                numero_correlativo += 1
                
            elif tipo_seleccionado == "Finalizar Inventario":
                # Opción para terminar
                if total_registros == 0:
                    print("\n⚠️  No hay registros capturados aún.")
                    continue
                
                print(f"\n{'=' * 60}")
                print(f"  ✅ FINALIZANDO SESIÓN")
                print(f"  Total de objetos inventariados: {total_registros}")
                print(f"{'=' * 60}")
                break
                
            else:
                # Flujo equipamiento simple
                estructura = capturar_equipo_simple(empresas, areas)
                
                # Generar código de barras
                ruta_barras, codigo_id = generar_barras_simple(
                    estructura["datos"]["nombre"],
                    numero_correlativo
                )
                archivos_sesion.append(ruta_barras)
                
                # Guardar en Excel
                guardar_estructura_simple_en_excel(estructura, ruta_barras, codigo_id)
                
                total_registros += 1
                numero_correlativo += 1

        except (IndexError, ValueError) as e:
            print(f"⚠️  Entrada inválida: {e}. Intenta de nuevo.")
        except KeyboardInterrupt:
            print("\n⚠️  Interrumpido.")
            break
        except Exception as e:
            print(f"❌ Error inesperado: {e}")
            import traceback
            traceback.print_exc()

    # ─────────────────────────────────
    # CIERRE Y ENVÍO
    # ─────────────────────────────────
    
    if total_registros == 0:
        print("\nℹ️  Sin registros capturados. Fin.")
        return

    print(f"\n✅ Preparando para enviar {total_registros} registros...")
    
    # Esperar a que el usuario confirme antes de enviar
    confirmacion = input("\n¿Deseas enviar los datos? (S/N): ").strip().upper()
    
    if confirmacion != "S":
        print(f"\n⚠️  Envío cancelado.")
        print(f"   Los datos se guardaron en: {ARCHIVO_SESION}")
        return

    # Enviar correo
    archivos_adjuntos = archivos_sesion + [ARCHIVO_SESION]
    exito = enviar_correo(archivos_adjuntos, total_registros)

    if exito:
        print(f"\n{'=' * 60}")
        print(f"  ✅ INVENTARIO ENVIADO EXITOSAMENTE")
        print(f"  Total de objetos: {total_registros}")
        print(f"  Archivo: {ARCHIVO_SESION}")
        print(f"{'=' * 60}")
        print("\n🏁 ¡Proceso completado!")
        print("   sync_bd.py procesará los datos automáticamente.")
    else:
        print("\n⚠️  El archivo se guardó pero no se pudo enviar el correo.")
        print(f"   Guarda el archivo {ARCHIVO_SESION} para enviarlo manualmente.")


# ──────────────────────────────────────────────
# ENTRY POINT
# ──────────────────────────────────────────────

if __name__ == "__main__":
    try:
        ejecutar_sistema()
    except KeyboardInterrupt:
        print("\n\n🛑 Sistema interrumpido por el usuario.")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Error fatal: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)