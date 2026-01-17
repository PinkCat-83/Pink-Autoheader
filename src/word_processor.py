"""
Procesamiento de documentos Word
Maneja la apertura, modificación y conversión de archivos DOCX
"""

import os
import time
import traceback
from src.config import *


class WordProcessor:
    """Procesa documentos Word añadiendo encabezados, pies de página y convirtiéndolos a PDF"""
    
    def __init__(self, ruta_logo, autor):
        """
        Inicializa el procesador de Word
        
        Args:
            ruta_logo (str): Ruta al archivo de imagen del logo
            autor (str): Nombre del autor para el pie de página
        """
        self.ruta_logo = ruta_logo
        self.autor = autor
    
    def procesar_docx(self, word, ruta_completa, archivo, codigo_ejercicio, carpeta_destino, log_callback, opciones):
        """
        Procesa un archivo DOCX: abre, modifica, guarda copia y convierte a PDF
        
        Args:
            word: Instancia de Word Application (win32com)
            ruta_completa (str): Ruta completa al archivo DOCX
            archivo (str): Nombre del archivo
            codigo_ejercicio (str): Código del ejercicio para el encabezado
            carpeta_destino (str): Carpeta donde guardar los resultados
            log_callback (callable): Función para escribir en el log
            opciones (dict): Diccionario con las opciones de procesamiento del GUI
        
        Returns:
            bool: True si el procesamiento fue exitoso, False en caso contrario
        """
        doc = None
        try:
            # Crear carpeta destino si no existe
            os.makedirs(carpeta_destino, exist_ok=True)
            
            ruta_normalizada = os.path.normpath(os.path.abspath(ruta_completa))
            log_callback(f"\n>>> {archivo}")
            
            # Abrir documento
            doc = word.Documents.Open(ruta_normalizada)
            
            # Insertar encabezado y pie de página con opciones
            self.insertar_encabezado(doc, codigo_ejercicio, log_callback, opciones)
            self.insertar_pie_pagina(doc, log_callback, opciones)
            
            time.sleep(1)
            
            # --- GUARDADO ---
            # Guardar copia del DOCX modificado (si está activado)
            if opciones.get('save_modified_dest', True):
                # Detectar extensión original
                ext = '.docm' if archivo.lower().endswith('.docm') else '.docx'
                docx_copia_nombre = archivo.replace(ext, f' - COPIA{ext}')
                docx_copia_ruta = os.path.normpath(os.path.join(carpeta_destino, docx_copia_nombre))
                
                # Determinar formato de guardado
                file_format = WD_FORMAT_XML_DOCUMENT_MACRO if ext == '.docm' else WD_FORMAT_XML_DOCUMENT
                
                doc.SaveAs(docx_copia_ruta, FileFormat=file_format)
                log_callback(f"    ✓ Copia Word guardada")
            
            # Guardar como PDF (si está activado)
            if opciones.get('copy_as_pdf', True):
                pdf_nombre = (archivo.rsplit('.', 1)[0]) + ".pdf"
                pdf_ruta = os.path.normpath(os.path.join(carpeta_destino, pdf_nombre))
                doc.SaveAs(pdf_ruta, FileFormat=WD_FORMAT_PDF)
                log_callback(f"  ✓ PDF generado")
            
            # Cerrar sin guardar cambios en el original
            doc.Close(SaveChanges=False)
            time.sleep(0.5)
            return True
            
        except Exception as e:
            log_callback(f"  ✗ ERROR: {e}")
            log_callback(traceback.format_exc())
            try:
                if doc:
                    doc.Close(SaveChanges=False)
            except:
                pass
            return False
    
    def insertar_encabezado(self, doc, codigo_ejercicio, log_callback, opciones):
        """
        Inserta logo, código y línea en el encabezado del documento
        
        Args:
            doc: Documento de Word
            codigo_ejercicio (str): Código del ejercicio
            log_callback (callable): Función para escribir en el log
            opciones (dict): Opciones de configuración
        """
        try:
            for section in doc.Sections:
                header = section.Headers(WD_HEADER_FOOTER_PRIMARY)
                header_range = header.Range
                
                # Limpiar encabezado existente
                header_range.Delete()
                
                # 1. Insertar párrafo vacío inicial (para anclar el logo)
                header_range.InsertParagraphAfter()
                
                # 2. Insertar código del ejercicio (si está activado)
                if opciones.get('add_folder_code', True):
                    header_range.InsertAfter(codigo_ejercicio)
                    
                    # Configurar el párrafo del código
                    codigo_para = header.Range.Paragraphs(2)
                    codigo_para.Range.Font.Name = HEADER_FONT_NAME
                    codigo_para.Range.Font.Size = HEADER_FONT_SIZE
                    codigo_para.Range.Font.Bold = True
                    codigo_para.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_RIGHT
                    codigo_para.Range.ParagraphFormat.SpaceAfter = HEADER_SPACE_AFTER
                
                # 3. Insertar línea separadora (si está activado)
                if opciones.get('add_header_line', True):
                    self._insertar_linea_horizontal(header, doc, LINE_POSITION_Y_HEADER)
                
                # 4. Insertar logo flotante (si está activado y existe)
                if opciones.get('add_logo', True) and self.ruta_logo and os.path.exists(self.ruta_logo):
                    self._insertar_logo(header, doc)
                
        except Exception as e:
            log_callback(f"    ⚠ Error encabezado: {e}")
    
    def insertar_pie_pagina(self, doc, log_callback, opciones):
        """
        Inserta línea separadora, autor y número de página en el pie de página
        
        Args:
            doc: Documento de Word
            log_callback (callable): Función para escribir en el log
            opciones (dict): Opciones de configuración
        """
        try:
            for section in doc.Sections:
                footer = section.Footers(WD_HEADER_FOOTER_PRIMARY)
                footer_range = footer.Range
                
                # Limpiar pie de página existente
                footer_range.Delete()
                
                # Construcción del texto del pie (de atrás hacia adelante)
                
                # 1. Número de página (si está activado)
                if opciones.get('add_page_number', True):
                    # Insertar campo de número total de páginas (último elemento)
                    temp_range1 = footer_range.Duplicate
                    temp_range1.Collapse(WD_COLLAPSE_START)
                    numpages_field = footer_range.Fields.Add(
                        Range=temp_range1,
                        Type=WD_FIELD_NUM_PAGES
                    )
                    
                    # Insertar " de " antes
                    footer_range.InsertBefore(" de ")
                    
                    # Insertar campo de número de página antes
                    temp_range2 = footer_range.Duplicate
                    temp_range2.Collapse(WD_COLLAPSE_START)
                    page_field = footer_range.Fields.Add(
                        Range=temp_range2,
                        Type=WD_FIELD_PAGE
                    )
                    
                    # Insertar "Página " antes
                    footer_range.InsertBefore("Página ")
                
                # 2. Autor (si está activado y existe)
                if opciones.get('add_author', True) and self.autor:
                    # Añadir guión separador si ya hay número de página
                    prefijo = " — " if opciones.get('add_page_number', True) else ""
                    footer_range.InsertBefore(self.autor + prefijo)
                
                # Formatear todo el contenido
                footer_range.Font.Name = FOOTER_FONT_NAME
                footer_range.Font.Size = FOOTER_FONT_SIZE
                footer_range.Font.Bold = False
                footer_range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_RIGHT
                footer_range.ParagraphFormat.SpaceBefore = 0
                footer_range.ParagraphFormat.SpaceAfter = 0
                
                # Poner en negrita solo los números de página
                if opciones.get('add_page_number', True):
                    for field in footer.Range.Fields:
                        if field.Type in [WD_FIELD_PAGE, WD_FIELD_NUM_PAGES]:
                            field.Result.Font.Bold = True
                
                # 3. Insertar línea separadora (si está activado)
                if opciones.get('add_footer_line', True):
                    page_setup = doc.PageSetup
                    altura_pagina = page_setup.PageHeight
                    posicion_y = altura_pagina - LINE_POSITION_Y_FOOTER_OFFSET
                    self._insertar_linea_horizontal(footer, doc, posicion_y)
                
        except Exception as e:
            log_callback(f"    ⚠ Error pie: {e}")
    
    def _insertar_linea_horizontal(self, container, doc, posicion_y):
        """
        Inserta una línea horizontal con bolas en los extremos
        
        Args:
            container: Contenedor (header o footer)
            doc: Documento de Word
            posicion_y (float): Posición vertical de la línea
        """
        page_setup = doc.PageSetup
        margen_izq = page_setup.LeftMargin
        margen_der = page_setup.RightMargin
        ancho_pagina = page_setup.PageWidth
        
        inicio_x = margen_izq
        fin_x = ancho_pagina - margen_der
        
        # Crear la línea
        linea_shape = container.Shapes.AddLine(
            BeginX=inicio_x,
            BeginY=posicion_y,
            EndX=fin_x,
            EndY=posicion_y
        )
        
        # Configurar estilo de la línea
        linea_shape.Line.Weight = LINE_WEIGHT
        linea_shape.Line.ForeColor.RGB = LINE_COLOR_RGB
        
        # Añadir terminaciones en bola (pequeñas)
        linea_shape.Line.BeginArrowheadStyle = LINE_ARROWHEAD_STYLE
        linea_shape.Line.BeginArrowheadWidth = LINE_ARROWHEAD_WIDTH
        linea_shape.Line.BeginArrowheadLength = LINE_ARROWHEAD_LENGTH
        
        linea_shape.Line.EndArrowheadStyle = LINE_ARROWHEAD_STYLE
        linea_shape.Line.EndArrowheadWidth = LINE_ARROWHEAD_WIDTH
        linea_shape.Line.EndArrowheadLength = LINE_ARROWHEAD_LENGTH
    
    def _insertar_logo(self, header, doc):
        """
        Inserta el logo flotante centrado en el encabezado
        
        Args:
            header: Encabezado del documento
            doc: Documento de Word
        """
        # Anclar al primer párrafo
        parrafo_ancla = header.Range.Paragraphs(1)
        
        # Insertar imagen
        logo_shape = header.Shapes.AddPicture(
            FileName=self.ruta_logo,
            LinkToFile=False,
            SaveWithDocument=True,
            Anchor=parrafo_ancla.Range
        )
        
        # Configurar tamaño manteniendo proporción
        logo_shape.LockAspectRatio = True
        logo_shape.Height = LOGO_HEIGHT_POINTS
        
        # Hacer que flote detrás del texto (no desplaza contenido)
        logo_shape.WrapFormat.Type = WD_WRAP_BEHIND_TEXT
        
        # Centrar horizontalmente
        logo_shape.RelativeHorizontalPosition = WD_RELATIVE_HORIZONTAL_POSITION_MARGIN
        page_setup = doc.PageSetup
        ancho_disponible = page_setup.PageWidth - page_setup.LeftMargin - page_setup.RightMargin
        logo_shape.Left = (ancho_disponible - logo_shape.Width) / 2
        
        # Posicionar verticalmente
        logo_shape.RelativeVerticalPosition = WD_RELATIVE_VERTICAL_POSITION_PARAGRAPH
        logo_shape.Top = LOGO_TOP_POSITION
