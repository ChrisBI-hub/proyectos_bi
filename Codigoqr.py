"""
CodigoQR.py  (Termux) v2.1
==========================
Sistema mejorado de captura de inventario con:
  - CPU (con opciones: hardware+software, solo hardware, solo software)
  - Equipos individuales (sillas, mesas, micrófonos, etc.)
  - Periféricos individuales
  
MEJORAS v2.1:
  ✅ Botón "Volver" en menús para correcciones
  ✅ 3 opciones para registro de CPU (HW+SW, solo HW, solo SW con Host)
  ✅ Sin repetición de menús innecesarios
  ✅ Múltiples del mismo periférico (2 monitores, 2 auriculares, etc.)
  ✅ Resumen de CPU antes de enviar con opción de editar

Uso:
    python3 Codigoqr.py (corre normalmente)
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
archivo_excel  = '/sdcard/Download/inventario.xlsx'
carpeta_salida = '/sdcard/Download/codigos_generados'

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
    "myanez@abcsc.mx",
    "sgonzalez@abcsc.mx",
    "ymontoya@abcsc.mx",
    "reportes.bi@abcsc.mx",
]

EDIFICIOS = ["NORTE 180", "NORTE 182", "PATIO SEC"]
ESTADOS = ["Mal estado", "Futuro mantenimiento", "Buen estado", "Equipo nuevo"]
TIPOS_DISCO = ["HDD", "SSD"]
PERIFERICOS_DISPONIBLES = ["Monitor", "Teclado", "Mouse", "Webcam", "Auriculares", "Micrófono", "Bocinas"]
TIPOS_SENSORES = ["Sensor de humo", "Sensor de movimiento", "Sensor de wifi y red"]
TIPOS_EQUIPAMIENTO = [
    "CPU",
    "Aire Acondicionado",
    "Archivero",
    "Escritorio",
    "Impresora",
    "Laptop",
    "Librero",
    "Mesa",
    "Micrófono",
    "Mini Split",
    "No-Break",
    "Pizarra",
    "Refrigerador",
    "Sensores",
    "Silla",
    "Telefono fijo",
    "Ventilador",
    "Otros"
]

os.makedirs(carpeta_salida, exist_ok=True)


# ──────────────────────────────────────────────
# HELPERS CON BOTÓN VOLVER
# ──────────────────────────────────────────────

def seleccionar_opcion(lista: list, titulo: str = None, permitir_volver=True) -> str | None:
    """
    Menú de selección con opción de volver.
    Retorna None si selecciona volver.
    """
    opciones_mostrar = lista.copy()
    if permitir_volver:
        opciones_mostrar.append("← Volver")
    
    if titulo:
        print(f"\n{titulo}")
    for i, item in enumerate(opciones_mostrar, 1):
        print(f"   {i}. {item}")
    
    while True:
        try:
            idx = int(input("\n   Selecciona número: ")) - 1
            if idx == len(lista) and permitir_volver:
                return None  # Usuario seleccionó "Volver"
            if 0 <= idx < len(lista):
                return lista[idx]
            print(f"   ⚠️  Elige entre 1 y {len(opciones_mostrar)}.")
        except ValueError:
            print("   ⚠️  Ingresa un número válido.")


def input_seguro(prompt: str, permitir_vacio=False) -> str | None:
    """Input con validación. Retorna None si cancela."""
    while True:
        valor = input(prompt).strip()
        if valor == "":
            if permitir_vacio:
                return valor
            if input("   ¿Dejar vacío? (S/N): ").strip().upper() == "S":
                return valor
            continue
        return valor


# ──────────────────────────────────────────────
# CAPTURA DE CPU (3 OPCIONES)
# ──────────────────────────────────────────────

def capturar_cpu(empresas: list, areas: list) -> dict | None:
    """
    Captura datos de CPU con 3 opciones:
    1. Hardware + Software
    2. Solo Hardware
    3. Software (con Host)
    """
    
    # Seleccionar qué registrar
    while True:
        print("\n" + "=" * 55)
        print("  REGISTRO CPU — OPCIONES")
        print("=" * 55)
        
        opcion = seleccionar_opcion(
            ["Hardware + Software", "Solo Hardware", "Solo Software (con Host)"],
            "\n¿Qué deseas registrar?",
            permitir_volver=True
        )
        
        if opcion is None:
            return None  # Volver al menú principal
        
        break
    
    print("\n" + "=" * 55)
    print(f"  REGISTRO CPU — {opcion.upper()}")
    print(f"  {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 55)

    estructura_cpu = {
        "tipo_inventario": "CPU",
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "cpu": {},
        "software": {},
        "perifericos": [],
        "opcion": opcion,
    }

    # ─────────────────────────────────
    # SECCIÓN 1: HARDWARE
    # ─────────────────────────────────
    if "Hardware" in opcion or "Solo Hardware" == opcion:
        print("\n🖥️  HARDWARE CPU")
        print("─" * 55)
        
        host = input_seguro("Host: ")
        if host is None:
            return None
        
        no_serie = input_seguro("No. de Serie: ")
        if no_serie is None:
            return None
        
        empresa = seleccionar_opcion(empresas, "🏢 Empresa:", permitir_volver=False)
        if empresa is None:
            return None
            
        edificio = seleccionar_opcion(EDIFICIOS, "🏗️  Edificio:", permitir_volver=False)
        if edificio is None:
            return None
            
        area = seleccionar_opcion(areas, "📍 Área:", permitir_volver=False)
        if area is None:
            return None
        
        print("\n📊 Estado Físico:")
        estado_idx = seleccionar_opcion(ESTADOS, permitir_volver=False)
        if estado_idx is None:
            return None
        
        marca = input_seguro("Marca: ")
        if marca is None:
            return None
            
        modelo = input_seguro("Modelo: ")
        if modelo is None:
            return None
            
        procesador = input_seguro("Procesador (ej: Intel i7): ")
        if procesador is None:
            return None
            
        ram = input_seguro("RAM (ej: 16GB): ")
        if ram is None:
            return None
            
        capacidad_disco = input_seguro("Capacidad Disco (ej: 512GB): ")
        if capacidad_disco is None:
            return None
        
        print("\n💾 Tipo de Disco Duro:")
        tipo_disco = seleccionar_opcion(TIPOS_DISCO, permitir_volver=False)
        if tipo_disco is None:
            return None
        
        observaciones_hw = input_seguro("Observaciones Hardware: ", permitir_vacio=True)
        if observaciones_hw is None:
            return None

        estructura_cpu["cpu"] = {
            "host": host,
            "no_serie": no_serie,
            "empresa": empresa,
            "edificio": edificio,
            "area": area,
            "estado": estado_idx,
            "marca": marca,
            "modelo": modelo,
            "procesador": procesador,
            "ram": ram,
            "capacidad_disco": capacidad_disco,
            "tipo_disco": tipo_disco,
            "observaciones": observaciones_hw,
        }

    # ─────────────────────────────────
    # SECCIÓN 2: SOFTWARE
    # ─────────────────────────────────
    if "Software" in opcion:
        print("\n📦 SOFTWARE INSTALADO")
        print("─" * 55)
        
        # Si es solo software, pedir el Host
        if "Solo Software" in opcion:
            host_cpu = input_seguro("Host de la CPU: ")
            if host_cpu is None:
                return None
            estructura_cpu["cpu"]["host"] = host_cpu
        
        so = input_seguro("Sistema Operativo (ej: Windows 10): ")
        if so is None:
            return None
            
        office = input_seguro("Office (ej: Office 2021, LibreOffice, Ninguno): ")
        if office is None:
            return None
            
        antivirus = input_seguro("Antivirus (ej: Windows Defender, Norton, Ninguno): ")
        if antivirus is None:
            return None
            
        lector_pdf = input_seguro("Lector de PDF (ej: Adobe Reader, Foxit, Ninguno): ")
        if lector_pdf is None:
            return None
            
        erp = input_seguro("ERP (ej: SAP, Oracle, Ninguno): ")
        if erp is None:
            return None
        
        print("\n📝 Otros Software:")
        otro1 = input_seguro("Otro Software 1: ", permitir_vacio=True)
        if otro1 is None:
            return None
            
        otro2 = input_seguro("Otro Software 2: ", permitir_vacio=True)
        if otro2 is None:
            return None
            
        otro3 = input_seguro("Otro Software 3: ", permitir_vacio=True)
        if otro3 is None:
            return None

        estructura_cpu["software"] = {
            "SO": so,
            "Office": office,
            "Antivirus": antivirus,
            "Lector_PDF": lector_pdf,
            "ERP": erp,
            "Otro1": otro1,
            "Otro2": otro2,
            "Otro3": otro3,
        }

    # ─────────────────────────────────
    # SECCIÓN 3: PERIFÉRICOS
    # ─────────────────────────────────
    if "Hardware" in opcion and "Solo Software" not in opcion:
        print("\n🖱️  PERIFÉRICOS")
        print("─" * 55)
        
        perifericos_registrados = {}
        
        while True:
            print("\n¿Qué periférico deseas registrar?")
            perifericos_faltantes = [
                p for p in PERIFERICOS_DISPONIBLES 
                if p not in perifericos_registrados or perifericos_registrados[p] < 2
            ]
            
            if not perifericos_faltantes:
                print("   📌 Ya han sido registrados todos los periféricos disponibles.")
                break
            
            perifericos_faltantes.append("← Terminar periféricos")
            
            for i, item in enumerate(perifericos_faltantes, 1):
                print(f"   {i}. {item}")

            tipo_periferico = None
            terminar = False          # ← bandera de salida del bucle externo

            while True:
                try:
                    idx = int(input("\n   Selecciona número: ")) - 1
                    if idx == len(perifericos_faltantes) - 1:
                        confirmar = input("\n   ¿Estás seguro? (S/N): ").strip().upper()
                        if confirmar == "S":
                            terminar = True   # ← confirma salida
                        # Si dice N: terminar=False, volvemos a mostrar el menú
                        break               # siempre sale del bucle de input
                    elif 0 <= idx < len(perifericos_faltantes) - 1:
                        tipo_periferico = perifericos_faltantes[idx]
                        break
                    else:
                        print(f"   ⚠️  Elige entre 1 y {len(perifericos_faltantes)}.")
                except ValueError:
                    print("   ⚠️  Ingresa un número válido.")

            if terminar:
                break                 # ← ahora sí sale del bucle externo de periféricos

            if tipo_periferico is None:
                continue              # usuario dijo N, vuelve a mostrar el menú

            # Registrar el periférico seleccionado
            print(f"\n📌 {tipo_periferico.upper()}:")
            
            modelo = input_seguro(f"  Modelo: ")
            if modelo is None:
                continue
                
            no_serie = input_seguro(f"  No. de Serie: ", permitir_vacio=True)
            if no_serie is None:
                continue
                
            marca = input_seguro(f"  Marca: ")
            if marca is None:
                continue
            
            print(f"  Estado Físico de {tipo_periferico}:")
            estado = seleccionar_opcion(ESTADOS, permitir_volver=False)
            if estado is None:
                continue
                
            observaciones = input_seguro(f"  Observaciones: ", permitir_vacio=True)
            if observaciones is None:
                continue

            periferico = {
                "tipo": tipo_periferico,
                "modelo": modelo,
                "no_serie": no_serie,
                "marca": marca,
                "estado": estado,
                "observaciones": observaciones,
            }
            estructura_cpu["perifericos"].append(periferico)
            
            # Contar cuántos de este tipo se han registrado
            if tipo_periferico not in perifericos_registrados:
                perifericos_registrados[tipo_periferico] = 0
            perifericos_registrados[tipo_periferico] += 1
            
            print(f"   ✅ {tipo_periferico} registrado ({perifericos_registrados[tipo_periferico]})")

    return estructura_cpu


def mostrar_resumen_cpu(estructura_cpu: dict) -> bool:
    """
    Muestra resumen de CPU registrada.
    Retorna True si usuario quiere continuar, False si quiere editar.
    """
    print("\n" + "=" * 60)
    print("  📋 RESUMEN DE REGISTRO CPU")
    print("=" * 60)

    if estructura_cpu["cpu"]:
        print("\n🖥️  HARDWARE:")
        for clave, valor in estructura_cpu["cpu"].items():
            print(f"   • {clave:20} : {valor}")

    if estructura_cpu["software"]:
        print("\n📦 SOFTWARE:")
        for clave, valor in estructura_cpu["software"].items():
            if valor:  # Solo mostrar si tiene valor
                print(f"   • {clave:20} : {valor}")

    if estructura_cpu["perifericos"]:
        print("\n🖱️  PERIFÉRICOS:")
        for i, per in enumerate(estructura_cpu["perifericos"], 1):
            print(f"   {i}. {per['tipo']}")
            for clave, valor in per.items():
                if clave != "tipo" and valor:
                    print(f"      • {clave:18} : {valor}")

    print("\n" + "=" * 60)
    opcion = input("¿Deseas continuar? (S/N): ").strip().upper()
    return opcion == "S"


def capturar_equipo_simple(empresas: list, areas: list) -> dict | None:
    """Captura datos simplificados para otros tipos de equipamiento."""
    
    tipo = seleccionar_opcion(TIPOS_EQUIPAMIENTO[1:], "¿Qué tipo de equipamiento?", permitir_volver=True)
    if tipo is None:
        return None
    
    print("\n" + "=" * 45)
    print(f"  REGISTRO: {tipo.upper()}")
    print(f"  {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 45)

    nombre = input_seguro(f"\nNombre/Descripción del {tipo}: ")
    if nombre is None:
        return None
        
    no_serie = input_seguro("No. de Serie: ", permitir_vacio=True)
    if no_serie is None:
        return None
        
    marca = input_seguro("Marca: ", permitir_vacio=True)
    if marca is None:
        return None
        
    modelo = input_seguro("Modelo: ", permitir_vacio=True)
    if modelo is None:
        return None
    
    empresa = seleccionar_opcion(empresas, "🏢 Empresa:", permitir_volver=True)
    if empresa is None:
        return None
        
    edificio = seleccionar_opcion(EDIFICIOS, "🏗️  Edificio:", permitir_volver=True)
    if edificio is None:
        return None
        
    area = seleccionar_opcion(areas, "📍 Área:", permitir_volver=True)
    if area is None:
        return None
    
    tipo_sensor = ""
    ubicacion_en_edificio = ""
    resolucion_pantalla = ""
    sistema_operativo = ""

    if tipo == "Laptop":
        print("\n💻 DATOS DE LAPTOP:")
        resolucion_pantalla = input_seguro("Resolución de pantalla (ej: 1920x1080): ", permitir_vacio=True)
        if resolucion_pantalla is None:
            return None

        sistema_operativo = input_seguro("Sistema Operativo (ej: Windows 11): ", permitir_vacio=True)
        if sistema_operativo is None:
            return None

    if tipo == "Sensores":
        print("\n🛰️  DATOS DE SENSOR:")
        tipo_sensor = seleccionar_opcion(TIPOS_SENSORES, "Tipo de sensor:", permitir_volver=True)
        if tipo_sensor is None:
            return None

        ubicacion_en_edificio = input_seguro("Ubicación en el edificio (ej: Pasillo PB, Site, Sala de juntas): ")
        if ubicacion_en_edificio is None:
            return None
    
    print("\n📊 Estado Físico:")
    estado = seleccionar_opcion(ESTADOS, permitir_volver=False)
    if estado is None:
        return None
    
    observaciones = input_seguro("Observaciones: ", permitir_vacio=True)
    if observaciones is None:
        return None

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
            "ubicacion_en_edificio": ubicacion_en_edificio,
            "estado": estado,
            "tipo_sensor": tipo_sensor,
            "resolucion_pantalla": resolucion_pantalla,
            "sistema_operativo": sistema_operativo,
            "observaciones": observaciones,
        }
    }

    print(f"\n✅ {tipo} '{nombre}' registrado.")
    return estructura


# ──────────────────────────────────────────────
# GENERACIÓN DE CÓDIGOS
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


def generar_qr_cpu(estructura_cpu: dict, numero_correlativo: int) -> tuple:
    """Genera UN SOLO QR para toda la estructura CPU."""
    contenido_json = json.dumps({
        "tipo": "CPU",
        "host": estructura_cpu["cpu"].get("host", ""),
        "no_serie": estructura_cpu["cpu"].get("no_serie", ""),
        "empresa": estructura_cpu["cpu"].get("empresa", ""),
        "timestamp": estructura_cpu["timestamp"],
    })

    codigo_id = f"CPU{numero_correlativo:04d}"
    ruta_qr = os.path.join(carpeta_salida, f"qr_{codigo_id}.png")
    
    qrcode.make(contenido_json).save(ruta_qr)
    print(f"   📷 QR CPU    → {ruta_qr}")
    
    return ruta_qr, contenido_json, codigo_id


def generar_barras_objeto(tipo: str, numero_correlativo: int) -> tuple:
    """Genera un código de barras para cada objeto."""
    codigo_id = f"{extraer_tres_letras(tipo)}{numero_correlativo:04d}"
    ruta_barras_base = os.path.join(carpeta_salida, f"barras_{codigo_id}")
    
    bn_class = barcode.get_barcode_class("code128")
    ruta_barras_final = bn_class(codigo_id, writer=ImageWriter()).save(ruta_barras_base)
    
    print(f"   📊 {tipo:15} → {ruta_barras_final}")
    return ruta_barras_final, codigo_id


def generar_barras_simple(nombre: str, numero_correlativo: int) -> tuple:
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
    """Guarda la estructura CPU en múltiples hojas."""
    
    excel_path = ARCHIVO_SESION
    
    fila_cpu = {
        "Tipo": "CPU",
        "Host": estructura_cpu["cpu"].get("host", ""),
        "No_Serie": estructura_cpu["cpu"].get("no_serie", ""),
        "Empresa": estructura_cpu["cpu"].get("empresa", ""),
        "Edificio": estructura_cpu["cpu"].get("edificio", ""),
        "Area": estructura_cpu["cpu"].get("area", ""),
        "Estado": estructura_cpu["cpu"].get("estado", ""),
        "Marca": estructura_cpu["cpu"].get("marca", ""),
        "Modelo": estructura_cpu["cpu"].get("modelo", ""),
        "Procesador": estructura_cpu["cpu"].get("procesador", ""),
        "RAM": estructura_cpu["cpu"].get("ram", ""),
        "Capacidad_Disco": estructura_cpu["cpu"].get("capacidad_disco", ""),
        "Tipo_Disco": estructura_cpu["cpu"].get("tipo_disco", ""),
        "Observaciones": estructura_cpu["cpu"].get("observaciones", ""),
        "Codigo_QR": codigo_qr,
        "Codigo_Barras_CPU": codigo_barras_cpu,
        "Timestamp": estructura_cpu["timestamp"],
    }
    
    df_cpu = pd.DataFrame([fila_cpu])
    
    fila_software = {
        "Host_CPU": estructura_cpu["cpu"].get("host", ""),
        "SO": estructura_cpu["software"].get("SO", ""),
        "Office": estructura_cpu["software"].get("Office", ""),
        "Antivirus": estructura_cpu["software"].get("Antivirus", ""),
        "Lector_PDF": estructura_cpu["software"].get("Lector_PDF", ""),
        "ERP": estructura_cpu["software"].get("ERP", ""),
        "Otro_1": estructura_cpu["software"].get("Otro1", ""),
        "Otro_2": estructura_cpu["software"].get("Otro2", ""),
        "Otro_3": estructura_cpu["software"].get("Otro3", ""),
        "Timestamp": estructura_cpu["timestamp"],
    }
    
    df_software = pd.DataFrame([fila_software])
    
    filas_perifericos = []
    contadores_tipo = {}
    for periferico in estructura_cpu["perifericos"]:
        tipo_periferico = periferico["tipo"]
        contadores_tipo[tipo_periferico] = contadores_tipo.get(tipo_periferico, 0) + 1
        sufijo = contadores_tipo[tipo_periferico]
        clave = f"{tipo_periferico}_{sufijo}"

        datos_codigo = codigos_perifericos.get(clave, {})
        codigo_barras = datos_codigo.get("barras", "")
        codigo_id     = datos_codigo.get("codigo_id", "")
        
        fila = {
            "Host_CPU":      estructura_cpu["cpu"].get("host", ""),
            "Tipo":          tipo_periferico,
            "Modelo":        periferico.get("modelo", ""),
            "No_Serie":      periferico.get("no_serie", ""),
            "Marca":         periferico.get("marca", ""),
            "Estado":        periferico.get("estado", ""),
            "Observaciones": periferico.get("observaciones", ""),
            "Codigo_Barras": codigo_barras,
            "Codigo_ID":     codigo_id,
            "Timestamp":     estructura_cpu["timestamp"],
        }
        filas_perifericos.append(fila)
    
    df_perifericos = pd.DataFrame(filas_perifericos) if filas_perifericos else pd.DataFrame()
    
    filas_relaciones = [{"Codigo_QR": codigo_qr, "Codigo_Barras_CPU": codigo_barras_cpu,
                         "Tipo_Periferico": None, "Codigo_Barras_Periferico": None}]
    for clave, datos in codigos_perifericos.items():
        filas_relaciones.append({
            "Codigo_QR":               codigo_qr,
            "Codigo_Barras_CPU":        codigo_barras_cpu,
            "Tipo_Periferico":          datos.get("tipo", ""),
            "Codigo_Barras_Periferico": datos.get("barras", ""),
        })
    df_relaciones = pd.DataFrame(filas_relaciones)
    
    try:
        if os.path.exists(excel_path):
            with pd.ExcelWriter(excel_path, mode="a", engine="openpyxl", 
                              if_sheet_exists="overlay") as writer:
                if "CPU" in writer.book.sheetnames:
                    df_existente = pd.read_excel(excel_path, sheet_name="CPU")
                    df_cpu = pd.concat([df_existente, df_cpu], ignore_index=True)
                df_cpu.to_excel(writer, sheet_name="CPU", index=False)
                
                if not df_software.empty:
                    if "Software" in writer.book.sheetnames:
                        df_existente = pd.read_excel(excel_path, sheet_name="Software")
                        df_software = pd.concat([df_existente, df_software], ignore_index=True)
                    df_software.to_excel(writer, sheet_name="Software", index=False)
                
                if not df_perifericos.empty:
                    if "Perifericos" in writer.book.sheetnames:
                        df_existente = pd.read_excel(excel_path, sheet_name="Perifericos")
                        df_perifericos = pd.concat([df_existente, df_perifericos], ignore_index=True)
                    df_perifericos.to_excel(writer, sheet_name="Perifericos", index=False)
                
                if "Relaciones" in writer.book.sheetnames:
                    df_existente = pd.read_excel(excel_path, sheet_name="Relaciones")
                    df_relaciones = pd.concat([df_existente, df_relaciones], ignore_index=True)
                df_relaciones.to_excel(writer, sheet_name="Relaciones", index=False)
        else:
            with pd.ExcelWriter(excel_path, mode="w", engine="openpyxl") as writer:
                df_cpu.to_excel(writer, sheet_name="CPU", index=False)
                if not df_software.empty:
                    df_software.to_excel(writer, sheet_name="Software", index=False)
                if not df_perifericos.empty:
                    df_perifericos.to_excel(writer, sheet_name="Perifericos", index=False)
                df_relaciones.to_excel(writer, sheet_name="Relaciones", index=False)
    except Exception as e:
        print(f"   ❌ Error guardando en Excel: {e}")


def guardar_estructura_simple_en_excel(estructura: dict, codigo_barras: str, codigo_id: str):
    """Guarda equipamiento simple en Excel."""
    
    excel_path = ARCHIVO_SESION
    
    fila = {
        "Tipo": estructura["tipo_inventario"],
        "Nombre": estructura["datos"].get("nombre", ""),
        "No_Serie": estructura["datos"].get("no_serie", ""),
        "Marca": estructura["datos"].get("marca", ""),
        "Modelo": estructura["datos"].get("modelo", ""),
        "Empresa": estructura["datos"].get("empresa", ""),
        "Edificio": estructura["datos"].get("edificio", ""),
        "Area": estructura["datos"].get("area", ""),
        "Ubicacion_En_Edificio": estructura["datos"].get("ubicacion_en_edificio", ""),
        "Estado": estructura["datos"].get("estado", ""),
        "Tipo_Sensor": estructura["datos"].get("tipo_sensor", ""),
        "Resolucion_Pantalla": estructura["datos"].get("resolucion_pantalla", ""),
        "Sistema_Operativo": estructura["datos"].get("sistema_operativo", ""),
        "Observaciones": estructura["datos"].get("observaciones", ""),
        "Codigo_Barras": codigo_id,
        "Codigo_ID": codigo_id,
        "Timestamp": estructura["timestamp"],
    }
    
    df = pd.DataFrame([fila])
    
    try:
        if os.path.exists(excel_path):
            try:
                df_existente = pd.read_excel(excel_path, sheet_name="Otros")
                df_final = pd.concat([df_existente, df], ignore_index=True)
            except:
                df_final = df
            
            with pd.ExcelWriter(excel_path, mode="a", engine="openpyxl", 
                              if_sheet_exists="overlay") as writer:
                df_final.to_excel(writer, sheet_name="Otros", index=False)
        else:
            with pd.ExcelWriter(excel_path, mode="w", engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Otros", index=False)
    except Exception as e:
        print(f"   ❌ Error guardando en Excel: {e}")


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
        f"— Sistema BI · Inventario v2.1"
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
        return True
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
    print("  SISTEMA DE INVENTARIO QR v2.1")
    print("  Termux — Captura de Equipamiento")
    print("=" * 60)

    try:
        xls = pd.ExcelFile(archivo_excel)
        empresas = pd.read_excel(xls, "Empresa")["Empresa"].dropna().tolist()
        areas = pd.read_excel(xls, "Areas")["Area"].dropna().tolist()
    except FileNotFoundError:
        print("❌ No se encontró inventario.xlsx. Crea uno primero.")
        return
    except Exception as e:
        print(f"❌ Error cargando Excel: {e}")
        return

    # Bucle principal
    while True:
        tipo_seleccionado = seleccionar_opcion(
            TIPOS_EQUIPAMIENTO,
            "\n🎯 ¿Qué deseas inventariar?"
        )
        
        if tipo_seleccionado is None:
            # Usuario seleccionó "Volver" (que no existe en menú principal)
            break

        if tipo_seleccionado == "CPU":
            # Flujo CPU
            while True:
                estructura = capturar_cpu(empresas, areas)
                
                if estructura is None:
                    break  # Volver al menú principal
                
                # Mostrar resumen
                if not mostrar_resumen_cpu(estructura):
                    continue  # Volver a registrar esta CPU
                
                # Generar QR
                ruta_qr, contenido_qr, codigo_qr = generar_qr_cpu(estructura, numero_correlativo)
                archivos_sesion.append(ruta_qr)
                
                # Generar código de barras para la CPU
                ruta_barras_cpu, codigo_barras_cpu = generar_barras_objeto(
                    "CPU", numero_correlativo
                )
                archivos_sesion.append(ruta_barras_cpu)
                
                # Generar códigos de barras para cada periférico
                # FIX: Usamos lista de (tipo, datos) para soportar 2 del mismo tipo (ej: 2 monitores)
                # Antes (bug): dict con clave=tipo → el segundo monitor sobreescribía al primero
                # Ahora (fix): contador por tipo para generar IDs únicos (MON0001, MON0002, etc.)
                codigos_perifericos = {}
                contadores_tipo = {}
                for periferico in estructura["perifericos"]:
                    tipo_p = periferico["tipo"]
                    contadores_tipo[tipo_p] = contadores_tipo.get(tipo_p, 0) + 1
                    # Código único: tipo + correlativo_global + instancia del tipo
                    sufijo = contadores_tipo[tipo_p]
                    ruta_barras_periferico, codigo_id = generar_barras_objeto(
                        tipo_p,
                        numero_correlativo * 100 + sufijo  # garantiza unicidad
                    )
                    archivos_sesion.append(ruta_barras_periferico)
                    # Guardamos por clave única tipo+sufijo para relacionar con la fila del Excel
                    clave = f"{tipo_p}_{sufijo}"
                    codigos_perifericos[clave] = {
                        "tipo": tipo_p,
                        "barras": ruta_barras_periferico,
                        "codigo_id": codigo_id,
                        "instancia": sufijo,
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
                print(f"\n✅ CPU registrada y guardada.")
                break  # Volver al menú principal

        else:
            # Flujo equipamiento simple
            estructura = capturar_equipo_simple(empresas, areas)
            
            if estructura is None:
                continue  # Volver al menú principal
            
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
            print(f"✅ {estructura['tipo_inventario']} registrado y guardado.")

        # Preguntar si continuar
        continuar = input("\n¿Deseas registrar más equipamiento? (S/N): ").strip().upper()
        if continuar != "S":
            break

    # ─────────────────────────────
    # CIERRE Y ENVÍO
    # ─────────────────────────────
    
    if total_registros == 0:
        print("\nℹ️  Sin registros capturados.")
        return

    print(f"\n{'=' * 60}")
    print(f"  ✅ SESIÓN COMPLETADA")
    print(f"  Total de objetos inventariados: {total_registros}")
    print(f"  Archivo: {ARCHIVO_SESION}")
    print(f"{'=' * 60}")

    confirmacion = input("\n¿Deseas enviar los datos? (S/N): ").strip().upper()
    
    if confirmacion != "S":
        print(f"\n⚠️  Envío cancelado.")
        print(f"   Los datos se guardaron en: {ARCHIVO_SESION}")
        return

    archivos_adjuntos = archivos_sesion + [ARCHIVO_SESION]
    exito = enviar_correo(archivos_adjuntos, total_registros)

    if exito:
        print(f"\n{'=' * 60}")
        print(f"  ✅ INVENTARIO ENVIADO EXITOSAMENTE")
        print(f"{'=' * 60}")
        print("\n🏁 ¡Proceso completado!")
        print("   sync_bd.py procesará los datos automáticamente.")
    else:
        print("\n⚠️  El archivo se guardó pero no se pudo enviar el correo.")


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