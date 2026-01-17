"""
Controlador principal de la aplicación
Coordina la lógica de negocio entre la GUI y los procesadores
"""

import os
import threading
import traceback
import win32com.client
import psutil
from tkinter import filedialog, messagebox

from src.word_processor import WordProcessor
from src.file_manager import FileManager
from src.utils import extraer_codigo, archivo_contiene_prohibida
from src.config_manager import ConfigManager


class AppController:
    """Controlador principal que coordina toda la lógica de la aplicación"""

    def __init__(self):
        """Inicializa el controlador"""
        self.gui = None
        self.carpetas_a_procesar = []
        self.carpeta_destino = ""
        self.ruta_logo = ""
        self.procesando = False
        self.total_archivos = 0
        self.archivos_procesados = 0
        self.config_manager = ConfigManager()

    def set_gui(self, gui):
        """Establece la referencia a la GUI e inicia la carga de configuración"""
        self.gui = gui
        self._cargar_configuracion_inicial()

    def _cargar_configuracion_inicial(self):
        """Carga los valores guardados en el config.ini a la interfaz"""
        try:
            # Cargar Autor y Logo
            author = self.config_manager.get_str('USER', 'author')
            if author:
                self.gui.entry_autor.delete(0, 'end')
                self.gui.entry_autor.insert(0, author)

            last_logo = self.config_manager.get_str('USER', 'last_logo')
            if last_logo and os.path.exists(last_logo):
                self._establecer_logo(last_logo, origen="config")

            # Cargar carpeta de último destino utilizado
            last_destination = self.config_manager.get_str('USER', 'last_destination')
            if last_destination and os.path.exists(last_destination):
                self.carpeta_destino = last_destination
                self.gui.establecer_carpeta_destino(last_destination)

            # Cargar Opciones: Encabezado y Pie
            self.gui.var_add_logo.set(self.config_manager.get_bool('HEADER_FOOTER', 'add_logo', True))
            self.gui.var_add_folder_code.set(self.config_manager.get_bool('HEADER_FOOTER', 'add_folder_code', True))
            self.gui.var_add_header_line.set(self.config_manager.get_bool('HEADER_FOOTER', 'add_header_line', True))
            self.gui.var_add_footer_line.set(self.config_manager.get_bool('HEADER_FOOTER', 'add_footer_line', True))
            self.gui.var_add_author.set(self.config_manager.get_bool('HEADER_FOOTER', 'add_author', True))
            self.gui.var_add_page_number.set(self.config_manager.get_bool('HEADER_FOOTER', 'add_page_number', True))

            # Cargar Opciones: Copia
            self.gui.var_respect_structure.set(self.config_manager.get_bool('COPY_OPTIONS', 'respect_structure', True))
            self.gui.var_copy_attachments.set(self.config_manager.get_bool('COPY_OPTIONS', 'copy_attachments', True))
            self.gui.var_save_modified_dest.set(self.config_manager.get_bool('COPY_OPTIONS', 'save_modified_in_dest', True))
            self.gui.var_copy_as_pdf.set(self.config_manager.get_bool('COPY_OPTIONS', 'copy_as_pdf', True))

            # Cargar Opciones: Extensiones
            self.gui.var_process_docx.set(self.config_manager.get_bool('PROCESS_EXTENSIONS', 'process_docx', True))
            self.gui.var_process_docm.set(self.config_manager.get_bool('PROCESS_EXTENSIONS', 'process_docm', False))

            # Cargar Exclusiones
            no_process = self.config_manager.get_str('EXCLUSIONS', 'no_process_names')
            if no_process:
                self.gui.text_no_process.insert('1.0', no_process)

            no_copy = self.config_manager.get_str('EXCLUSIONS', 'no_copy_names')
            if no_copy:
                self.gui.text_no_copy.insert('1.0', no_copy)

            self.log("✓ Configuración cargada correctamente")
        except Exception as e:
            self.log(f"⚠ Error al cargar configuración: {e}")

    def word_esta_abierto(self):
        """Verifica si hay alguna instancia de Word abierta"""
        try:
            for proceso in psutil.process_iter(['name']):
                if proceso.info['name'] and proceso.info['name'].lower() == 'winword.exe':
                    return True
            return False
        except Exception:
            return False

    def examinar_logo(self):
        """Abre diálogo para seleccionar archivo de logo"""
        archivo = filedialog.askopenfilename(
            title="Seleccionar Logo",
            filetypes=[("Imágenes", "*.png *.jpg *.jpeg"), ("Todos los archivos", "*.*")]
        )
        if archivo:
            self._establecer_logo(archivo, origen="seleccionado")

    def drop_logo(self, event):
        """Maneja el evento de arrastrar y soltar logo"""
        rutas = event.widget.tk.splitlist(event.data)
        if rutas:
            ruta = rutas[0].strip('{}')
            if os.path.isfile(ruta) and ruta.lower().endswith(('.png', '.jpg', '.jpeg')):
                self._establecer_logo(ruta, origen="arrastrado")
            else:
                self.gui.mostrar_error("Error", "Arrastra un archivo de imagen válido (PNG/JPG)")

    def _establecer_logo(self, ruta, origen="seleccionado"):
        self.ruta_logo = os.path.normpath(ruta)
        self.gui.mostrar_preview_logo(self.ruta_logo)
        self.log(f"✓ Logo {origen}: {self.ruta_logo}")

    def agregar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Seleccionar Carpeta a Procesar")
        if carpeta:
            ruta = os.path.normpath(carpeta)
            if ruta not in self.carpetas_a_procesar:
                self.carpetas_a_procesar.append(ruta)
                self.gui.agregar_carpeta_a_lista(ruta)
                self.log(f"✓ Carpeta agregada: {ruta}")

    def drop_carpeta(self, event):
        rutas = self.gui.listbox_carpetas.tk.splitlist(event.data)
        for r in rutas:
            ruta = os.path.normpath(r.strip('{}'))
            if os.path.isdir(ruta) and ruta not in self.carpetas_a_procesar:
                self.carpetas_a_procesar.append(ruta)
                self.gui.agregar_carpeta_a_lista(ruta)
                self.log(f"✓ Carpeta agregada: {ruta}")

    def quitar_carpeta(self):
        index = self.gui.obtener_seleccion_carpeta()
        if index is not None:
            self.carpetas_a_procesar.pop(index)
            self.gui.quitar_carpeta_de_lista(index)

    def seleccionar_destino(self):
        carpeta = filedialog.askdirectory(title="Seleccionar Carpeta Destino")
        if carpeta:
            self.carpeta_destino = os.path.normpath(carpeta)
            self.gui.establecer_carpeta_destino(self.carpeta_destino)

    def log(self, mensaje):
        if self.gui: self.gui.log(mensaje)

    def actualizar_progreso(self, texto=None):
        if self.total_archivos > 0:
            porcentaje = (self.archivos_procesados / self.total_archivos) * 100
            self.gui.actualizar_progreso(porcentaje, texto or f"Procesados: {self.archivos_procesados}/{self.total_archivos}")

    def empezar_proceso(self):
        """Valida, guarda configuración y lanza el proceso"""
        if not self.ruta_logo and self.gui.var_add_logo.get():
            self.gui.mostrar_error("Error", "Debes seleccionar un logo si la opción está activa")
            return
        if not self.carpetas_a_procesar:
            self.gui.mostrar_error("Error", "Agrega carpetas a procesar")
            return

        self.carpeta_destino = os.path.normpath(self.gui.entry_destino.get().strip())
        if not os.path.isdir(self.carpeta_destino):
            self.gui.mostrar_error("Error", "Carpeta destino no válida")
            return

        if self.word_esta_abierto():
            self.gui.mostrar_error("Word Abierto", "Cierra Word antes de empezar")
            return

        # Guardar Configuración
        self.config_manager.set_val('USER', 'author', self.gui.entry_autor.get())
        self.config_manager.set_val('USER', 'last_logo', self.ruta_logo)
        self.config_manager.set_val('USER', 'last_destination', self.carpeta_destino)

        self.config_manager.set_val('HEADER_FOOTER', 'add_logo', self.gui.var_add_logo.get())
        self.config_manager.set_val('HEADER_FOOTER', 'add_folder_code', self.gui.var_add_folder_code.get())
        self.config_manager.set_val('HEADER_FOOTER', 'add_header_line', self.gui.var_add_header_line.get())
        self.config_manager.set_val('HEADER_FOOTER', 'add_footer_line', self.gui.var_add_footer_line.get())
        self.config_manager.set_val('HEADER_FOOTER', 'add_author', self.gui.var_add_author.get())
        self.config_manager.set_val('HEADER_FOOTER', 'add_page_number', self.gui.var_add_page_number.get())

        self.config_manager.set_val('COPY_OPTIONS', 'respect_structure', self.gui.var_respect_structure.get())
        self.config_manager.set_val('COPY_OPTIONS', 'copy_attachments', self.gui.var_copy_attachments.get())
        self.config_manager.set_val('COPY_OPTIONS', 'save_modified_in_dest', self.gui.var_save_modified_dest.get())
        self.config_manager.set_val('COPY_OPTIONS', 'copy_as_pdf', self.gui.var_copy_as_pdf.get())

        self.config_manager.set_val('PROCESS_EXTENSIONS', 'process_docx', self.gui.var_process_docx.get())
        self.config_manager.set_val('PROCESS_EXTENSIONS', 'process_docm', self.gui.var_process_docm.get())

        self.config_manager.set_val('EXCLUSIONS', 'no_process_names', self.gui.text_no_process.get('1.0', 'end-1c'))
        self.config_manager.set_val('EXCLUSIONS', 'no_copy_names', self.gui.text_no_copy.get('1.0', 'end-1c'))

        self.procesando = True
        self.gui.deshabilitar_boton_empezar()
        self.gui.limpiar_log()
        self.archivos_procesados = 0

        threading.Thread(target=self.procesar_archivos, daemon=True).start()

    def procesar_archivos(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            self.log("=== INICIANDO PROCESO ===")

            # Extensiones permitidas
            exts = []
            if self.gui.var_process_docx.get(): exts.append('.docx')
            if self.gui.var_process_docm.get(): exts.append('.docm')

            if not exts:
                self.log("⚠ No hay extensiones seleccionadas para procesar")
                return

            # Exclusiones - parsear por comas Y saltos de línea
            texto_exc_process = self.gui.text_no_process.get('1.0', 'end-1c')
            texto_exc_copy = self.gui.text_no_copy.get('1.0', 'end-1c')

            # Dividir por comas y saltos de línea, limpiar espacios
            exc_process = []
            for item in texto_exc_process.replace('\n', ',').split(','):
                item = item.strip().lower()
                if item:
                    exc_process.append(item)

            exc_copy = []
            for item in texto_exc_copy.replace('\n', ',').split(','):
                item = item.strip().lower()
                if item:
                    exc_copy.append(item)

            self.log(f"Exclusiones de proceso: {exc_process}")
            self.log(f"Exclusiones de copia: {exc_copy}")

            # Contar archivos
            self.total_archivos = 0
            for c in self.carpetas_a_procesar:
                for root, dirs, files in os.walk(c):
                    # Filtrar carpetas excluidas
                    dirs[:] = [d for d in dirs if not any(exc in d.lower() for exc in exc_process)]
                    # Contar solo archivos Word no excluidos
                    for f in files:
                        if any(f.lower().endswith(e) for e in exts):
                            if not any(exc in f.lower() for exc in exc_process):
                                self.total_archivos += 1

            self.log(f"Total archivos a procesar: {self.total_archivos}")

            word = win32com.client.Dispatch('Word.Application')
            word.Visible = True

            processor = WordProcessor(self.ruta_logo, self.gui.entry_autor.get())

            for carpeta_origen in self.carpetas_a_procesar:
                # Obtener el nombre de la carpeta raíz que se está procesando
                nombre_carpeta_raiz = os.path.basename(carpeta_origen)
                
                for root, dirs, files in os.walk(carpeta_origen):
                    # Filtrar carpetas excluidas de proceso
                    dirs[:] = [d for d in dirs if not any(exc in d.lower() for exc in exc_process)]

                    # Calcular ruta relativa para replicar estructura
                    rel_path = os.path.relpath(root, carpeta_origen)
                    codigo = extraer_codigo(os.path.basename(root))

                    for f in files:
                        f_lower = f.lower()
                        es_word = any(f_lower.endswith(e) for e in exts)

                        # Determinar ruta de destino
                        if self.gui.var_respect_structure.get():
                            # Incluir el nombre de la carpeta raíz + estructura interna
                            if rel_path == '.':
                                # Estamos en la raíz de carpeta_origen
                                ruta_dest_final = os.path.join(self.carpeta_destino, nombre_carpeta_raiz, f)
                            else:
                                # Estamos en una subcarpeta
                                ruta_dest_final = os.path.join(self.carpeta_destino, nombre_carpeta_raiz, rel_path, f)
                        else:
                            ruta_dest_final = os.path.join(self.carpeta_destino, f)

                        # 1. Si es Word y está excluido de proceso
                        if es_word and any(exc in f_lower for exc in exc_process):
                            self.log(f"⊗ Excluido de proceso: {f}")
                            # Copiar si NO está en exclusiones de copia
                            if not any(exc in f_lower for exc in exc_copy):
                                FileManager.copiar_archivo(os.path.join(root, f), ruta_dest_final, self.log)
                            continue

                        # 2. Si es Word y NO está excluido -> PROCESAR
                        if es_word:
                            dest_folder_final = os.path.dirname(ruta_dest_final)
                            if processor.procesar_docx(word, os.path.join(root, f), f, codigo, dest_folder_final, self.log, self.gui.obtener_opciones_completas()):
                                self.archivos_procesados += 1
                                self.actualizar_progreso()

                        # 3. Si NO es Word -> es anexo
                        else:
                            # Verificar si está excluido de copia
                            if any(exc in f_lower for exc in exc_copy):
                                self.log(f"⊗ Excluido de copia: {f}")
                            else:
                                # Copiar si está activado
                                if self.gui.var_copy_attachments.get():
                                    FileManager.copiar_archivo(os.path.join(root, f), ruta_dest_final, self.log)
                                    
            word.Quit()
            self.log("\n=== ✅ COMPLETADO ===")
            self.gui.mostrar_info("Completado", "Proceso finalizado con éxito")

        except Exception as e:
            self.log(f"❌ ERROR: {e}")
            self.log(traceback.format_exc())
            self.gui.mostrar_error("Error", str(e))
        finally:
            pythoncom.CoUninitialize()
            self.procesando = False
            self.gui.habilitar_boton_empezar()