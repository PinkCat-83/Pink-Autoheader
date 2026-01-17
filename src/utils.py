"""
Funciones utilitarias generales
"""

import os


def extraer_codigo(nombre_carpeta):
    """
    Extrae código tipo '01 - Introducción' del nombre de carpeta
    
    Args:
        nombre_carpeta (str): Nombre de la carpeta (ej: "01 - Introducción - Parte 1")
    
    Returns:
        str: Código extraído (ej: "01 - Introducción") o nombre completo si no tiene formato
    """
    partes = nombre_carpeta.split(" - ")
    if len(partes) >= 2:
        return f"{partes[0]} - {partes[1]}"
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