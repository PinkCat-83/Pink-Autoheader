"""
Funciones utilitarias generales
"""

import os


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