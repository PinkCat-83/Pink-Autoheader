"""
Configuración y constantes de la aplicación
"""

# ============================================
# CONFIGURACIÓN DE VENTANA
# ============================================
WINDOW_TITLE = "Convertidor DOCX a PDF con Encabezado y Pie de Página"
# WINDOW_SIZE = "850x1000"

# ============================================
# PALABRAS PROHIBIDAS POR DEFECTO
# ============================================
PALABRAS_PROHIBIDAS_DEFAULT = ["solución", "solucion"]

# ============================================
# CONFIGURACIÓN DE UI - PREVIEW DE LOGO
# ============================================
LOGO_PREVIEW_CONTAINER_HEIGHT = 120  # Altura fija del contenedor de preview en píxeles
LOGO_PREVIEW_MAX_WIDTH = 750         # Ancho máximo para la imagen del logo
LOGO_PREVIEW_MAX_HEIGHT = 80         # Alto máximo para la imagen del logo

# ============================================
# CONSTANTES DE WORD (win32com)
# ============================================
# Headers y Footers
WD_HEADER_FOOTER_PRIMARY = 1

# Formatos de archivo
WD_FORMAT_XML_DOCUMENT = 16  # .docx
WD_FORMAT_XML_DOCUMENT_MACRO = 13 # .docm
WD_FORMAT_PDF = 17           # .pdf

# Alineación
WD_ALIGN_PARAGRAPH_RIGHT = 2

# Wrap de imágenes
WD_WRAP_BEHIND_TEXT = 3

# Posicionamiento
WD_RELATIVE_HORIZONTAL_POSITION_MARGIN = 0
WD_RELATIVE_VERTICAL_POSITION_PARAGRAPH = 1

# Campos (Fields)
WD_FIELD_PAGE = 33          # Número de página actual
WD_FIELD_NUM_PAGES = 26     # Número total de páginas

# Collapse
WD_COLLAPSE_START = 1

# ============================================
# ESTILOS DE ENCABEZADO
# ============================================
HEADER_FONT_NAME = "Calibri"
HEADER_FONT_SIZE = 14
HEADER_SPACE_AFTER = 12  # Espaciado posterior en puntos

# ============================================
# ESTILOS DE PIE DE PÁGINA
# ============================================
FOOTER_FONT_NAME = "Calibri"
FOOTER_FONT_SIZE = 12

# ============================================
# CONFIGURACIÓN DE LOGO
# ============================================
LOGO_HEIGHT_POINTS = 33.7  # 1.19cm en puntos
LOGO_TOP_POSITION = 15     # Posición vertical desde arriba

# ============================================
# CONFIGURACIÓN DE LÍNEAS
# ============================================
LINE_WEIGHT = 1.5                        # Grosor de línea
LINE_COLOR_RGB = 0                       # Negro
LINE_ARROWHEAD_STYLE = 6                 # msoArrowheadOval (bola)
LINE_ARROWHEAD_WIDTH = 1                 # msoArrowheadWidthNarrow
LINE_ARROWHEAD_LENGTH = 1                # msoArrowheadLengthShort

# Posiciones de líneas
LINE_POSITION_Y_HEADER = 68              # Posición Y en encabezado (antes 68)
LINE_POSITION_Y_FOOTER_OFFSET = 60       # Offset desde el final de página

# ============================================
# COLORES DE INTERFAZ
# ============================================
COLOR_SUCCESS = "#4CAF50"
COLOR_ERROR = "#f44336"
COLOR_INFO = "#2196F3"
COLOR_NEUTRAL = "#9E9E9E"
COLOR_DISABLED = "#cccccc"
COLOR_LOGO_BG = "#f0f0f0"
COLOR_LOGO_SUCCESS = "#c8e6c9"

# ============================================
# BARRA DE PROGRESO (estilo rosa)
# ============================================
PROGRESS_BAR_STYLE = "pink.Horizontal.TProgressbar"
PROGRESS_BAR_COLORS = {
    'troughcolor': '#f0f0f0',
    'background': '#FF69B4',
    'bordercolor': '#cccccc',
    'lightcolor': '#FFB6D9',
    'darkcolor': '#FF1493'
}