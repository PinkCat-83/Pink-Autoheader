"""
Gestión de archivos y carpetas
"""

import os
import shutil


class FileManager:
    """Maneja operaciones de archivos y carpetas"""
    
    @staticmethod
    def _obtener_lista_exclusiones(texto_exclusiones):
        """
        Convierte un texto (separado por comas o saltos de línea) en una lista limpia.
        """
        if not texto_exclusiones:
            return []
        # Reemplazar saltos de línea por comas para procesar todo igual
        texto_limpio = texto_exclusiones.replace('\n', ',')
        return [e.strip().lower() for e in texto_limpio.split(',') if e.strip()]

    @staticmethod
    def _contiene_exclusion(nombre, lista_exclusiones):
        """
        Verifica si el nombre contiene alguna de las subcadenas de exclusión.
        """
        nombre_lower = nombre.lower()
        for excl in lista_exclusiones:
            if excl in nombre_lower:
                return True
        return False
    
    @staticmethod
    def contar_archivos(carpetas, extensiones, exclusiones_procesar):
        """
        Cuenta el total de archivos que coinciden con las extensiones permitidas
        en la raíz de las carpetas seleccionadas (no recursivo).
        """
        total = 0
        ext_tuple = tuple(ext.strip().lower() for ext in extensiones.split(','))
        excl_list = FileManager._obtener_lista_exclusiones(exclusiones_procesar)
        
        for carpeta in carpetas:
            if not os.path.exists(carpeta):
                continue
            
            # Solo archivos en la raíz de la carpeta seleccionada
            for f in os.listdir(carpeta):
                ruta_f = os.path.join(carpeta, f)
                if os.path.isfile(ruta_f):
                    if f.lower().endswith(ext_tuple):
                        if not FileManager._contiene_exclusion(f, excl_list):
                            total += 1
        return total

    @staticmethod
    def copiar_archivo(ruta_origen, ruta_destino, log_callback=None):
        """
        Copia un archivo individual de origen a destino.
        ruta_destino debe ser la ruta completa del archivo destino (incluyendo nombre).
        """
        try:
            # Crear carpetas de destino si no existen
            dir_destino = os.path.dirname(ruta_destino)
            if dir_destino:
                os.makedirs(dir_destino, exist_ok=True)
            
            # Evitar copiar sobre sí mismo
            if os.path.abspath(ruta_origen) != os.path.abspath(ruta_destino):
                shutil.copy2(ruta_origen, ruta_destino)
                return True
            return False
        except Exception as e:
            if log_callback:
                log_callback(f"  ERROR copiando {os.path.basename(ruta_origen)}: {e}")
            return False

    @staticmethod
    def copiar_archivos_excepto_word(carpeta_origen, carpeta_destino, extensiones_word, exclusiones_copiar, log_callback=None):
        """
        Copia archivos que no son de Word respetando la estructura y exclusiones.
        Si una carpeta está excluida, se ignora ella y todo su contenido.
        """
        ext_word = tuple(ext.strip().lower() for ext in extensiones_word.split(','))
        excl_list = FileManager._obtener_lista_exclusiones(exclusiones_copiar)

        for root, dirs, files in os.walk(carpeta_origen):
            # Filtrar directorios excluidos (usando "contiene")
            dirs[:] = [d for d in dirs if not FileManager._contiene_exclusion(d, excl_list)]
            
            # Calcular ruta relativa para replicar estructura
            rel_path = os.path.relpath(root, carpeta_origen)
            dest_root = os.path.join(carpeta_destino, rel_path) if rel_path != '.' else carpeta_destino
            
            for f in files:
                # Saltar archivos Word o excluidos (usando "contiene")
                if f.lower().endswith(ext_word) or FileManager._contiene_exclusion(f, excl_list):
                    continue
                
                ruta_f_origen = os.path.join(root, f)
                ruta_f_destino = os.path.join(dest_root, f)
                
                if FileManager.copiar_archivo(ruta_f_origen, ruta_f_destino, log_callback):
                    if log_callback:
                        log_callback(f"  Copiado: {f}")