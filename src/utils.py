"""
Funciones utilitarias generales
"""

import os
import re


def extraer_codigo(nombre_carpeta):
    """
    Extrae código del nombre de carpeta
    Soporta dos formatos:
    - "01 - Introducción - Parte 1" -> "01 - Introducción"
    - "CAL-05-Patata" -> "CAL-05"
    
    Args:
        nombre_carpeta (str): Nombre de la carpeta
    
    Returns:
        str: Código extraído o nombre completo si no tiene formato reconocido
    """
    # Caso 1: Formato con espacios " - " (ej: "01 - Introducción - Parte 1")
    if " - " in nombre_carpeta:
        partes = nombre_carpeta.split(" - ")
        if len(partes) >= 2:
            return f"{partes[0]} - {partes[1]}"
    
    # Caso 2: Formato con guiones sin espacios (ej: "CAL-05-Patata")
    elif "-" in nombre_carpeta:
        partes = nombre_carpeta.split("-")
        if len(partes) >= 2:
            return f"{partes[0]}-{partes[1]}"
    
    # Si no coincide con ningún formato, devolver el nombre completo
    return nombre_carpeta


def extraer_raiz_archivo(nombre_archivo):
    """
    Extrae la raíz del nombre de archivo detectando automáticamente el patrón.
    El patrón se identifica por la presencia de DOS GUIONES (con o sin espacios).
    Todo lo que viene DESPUÉS del segundo guión es la raíz.
    
    Ejemplos:
    - "DOC-13-Mi vida salvaje.docx" -> "Mi vida salvaje"
    - "CAL-87R1-Tareas de cálculo.pdf" -> "Tareas de cálculo"
    - "AABBCC-192X11-Base de Datos.docx" -> "Base de Datos"
    - "01 - Intro - Documento.docx" -> "Documento"
    
    Args:
        nombre_archivo (str): Nombre completo del archivo con extensión
    
    Returns:
        tuple: (raiz_detectada, patron_encontrado)
               raiz_detectada (str): Nombre raíz extraído o None si no se detectó
               patron_encontrado (bool): True si se encontró patrón automático
    """
    # Separar nombre y extensión
    nombre_sin_ext, extension = os.path.splitext(nombre_archivo)
    
    # Caso 1: Formato con espacios " - " (ej: "01 - Intro - Mi vida salvaje")
    if " - " in nombre_sin_ext:
        partes = nombre_sin_ext.split(" - ")
        if len(partes) >= 3:
            # Tomar todo después del segundo " - "
            raiz = " - ".join(partes[2:])
            return (raiz, True)
    
    # Caso 2: Formato con guiones sin espacios
    # Buscar el SEGUNDO guión y tomar todo después de él
    guiones_encontrados = 0
    for i, char in enumerate(nombre_sin_ext):
        if char == '-':
            guiones_encontrados += 1
            if guiones_encontrados == 2:
                # Encontramos el segundo guión, la raíz es todo lo que viene después
                raiz = nombre_sin_ext[i+1:]
                if raiz:  # Asegurar que hay algo después del segundo guión
                    return (raiz, True)
                break
    
    # No se encontró patrón reconocido (menos de 2 guiones o vacío después del segundo)
    return (None, False)


def construir_nombre_con_codigo(codigo, raiz, extension):
    """
    Construye el nuevo nombre de archivo combinando código + raíz + extensión.
    Adapta el formato según el tipo de código detectado.
    
    Args:
        codigo (str): Código de la carpeta (ej: "CAL-05" o "01 - Introducción")
        raiz (str): Raíz del nombre del archivo
        extension (str): Extensión del archivo (incluye el punto)
    
    Returns:
        str: Nuevo nombre completo del archivo
    """
    # Detectar formato del código para usar el separador adecuado
    if " - " in codigo:
        # Formato con espacios: "01 - Introducción"
        return f"{codigo} - {raiz}{extension}"
    else:
        # Formato con guiones: "CAL-05"
        return f"{codigo}-{raiz}{extension}"


def renombrar_archivo_con_codigo(ruta_archivo, codigo):
    """
    Renombra un archivo añadiendo el código de carpeta al principio.
    
    Si el archivo ya tiene un código (detectado por patrón de dos guiones),
    este será REEMPLAZADO por el nuevo código de carpeta.
    
    Proceso:
    1. Intenta detectar automáticamente la raíz del nombre (todo después del 2º guión)
    2. Si detecta patrón, renombra automáticamente
    3. Si NO detecta patrón, retorna información para que el controlador pida input
    
    Args:
        ruta_archivo (str): Ruta completa del archivo original
        codigo (str): Código de la carpeta a añadir
    
    Returns:
        tuple: (exito, nueva_ruta, mensaje, necesita_input)
               exito (bool): True si se renombró correctamente
               nueva_ruta (str): Nueva ruta del archivo o None si falló
               mensaje (str): Mensaje descriptivo o nombre de archivo si necesita input
               necesita_input (bool): True si necesita input del usuario
    """
    try:
        directorio = os.path.dirname(ruta_archivo)
        nombre_completo = os.path.basename(ruta_archivo)
        nombre_sin_ext, extension = os.path.splitext(nombre_completo)
        
        # Paso 1: Intentar detección automática
        raiz, patron_encontrado = extraer_raiz_archivo(nombre_completo)
        
        if patron_encontrado:
            # ✅ Patrón detectado automáticamente
            nuevo_nombre = construir_nombre_con_codigo(codigo, raiz, extension)
            nueva_ruta = os.path.join(directorio, nuevo_nombre)
            
            # Verificar que no exista ya (o que sea el mismo archivo)
            if os.path.exists(nueva_ruta) and os.path.abspath(ruta_archivo) != os.path.abspath(nueva_ruta):
                return (False, None, f"⚠️ Ya existe archivo con nombre: {nuevo_nombre}", False)
            
            # Si el nombre ya es correcto, no renombrar
            if nombre_completo == nuevo_nombre:
                return (True, ruta_archivo, f"✓ Ya tiene el código correcto: {nombre_completo}", False)
            
            # Renombrar
            os.rename(ruta_archivo, nueva_ruta)
            return (True, nueva_ruta, f"✓ Renombrado automáticamente: {nombre_completo} → {nuevo_nombre}", False)
        
        # No se detectó patrón - NECESITA INPUT DEL USUARIO
        # Retornar información para que el controlador maneje esto
        return (False, ruta_archivo, nombre_completo, True)  # necesita_input = True
        
    except Exception as e:
        return (False, None, f"❌ Error al renombrar {os.path.basename(ruta_archivo)}: {str(e)}", False)


def obtener_palabras_prohibidas_lista(texto):
    """
    Convierte texto multilínea en lista de palabras prohibidas
    
    Args:
        texto (str): Texto con palabras separadas por líneas
    
    Returns:
        list: Lista de palabras prohibidas en minúsculas
    """
    if not texto.strip():
        return []
    
    # Dividir por líneas, limpiar y convertir a minúsculas
    palabras = [linea.strip().lower() for linea in texto.split('\n') if linea.strip()]
    return palabras


def archivo_contiene_prohibida(nombre_archivo, palabras_prohibidas):
    """
    Verifica si el nombre del archivo contiene alguna palabra prohibida
    
    Args:
        nombre_archivo (str): Nombre del archivo a verificar
        palabras_prohibidas (list): Lista de palabras prohibidas
    
    Returns:
        bool: True si contiene alguna palabra prohibida, False en caso contrario
    """
    nombre_lower = nombre_archivo.lower()
    for palabra in palabras_prohibidas:
        if palabra in nombre_lower:
            return True
    return False