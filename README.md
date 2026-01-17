# Pink Autoheader

Aplicación de escritorio para Windows que automatiza la conversión masiva de documentos `.docx` (o `.docm`) a PDF, añadiendo encabezados personalizados (logo y código) y pies de página (autor y numeración).

![Interfaz de Pink Autoheader](img/002.png)

## Características

- ✅ Procesamiento masivo de archivos `.docx` y `.docm`
- ✅ Inserción automática de logo en encabezado
- ✅ Añadir código de carpeta y autor
- ✅ Numeración automática de páginas
- ✅ Líneas decorativas en encabezado y pie de página
- ✅ Filtrado por palabras prohibidas
- ✅ Opción de mantener estructura de carpetas original
- ✅ Interfaz gráfica intuitiva con soporte drag & drop

## Requisitos

- Windows 10/11
- Microsoft Word (Versión de escritorio de Office 365)
- Python 3.7+

## Dependencias

```bash
pip install pywin32 psutil Pillow tkinterdnd2
```

## Estructura del Proyecto

```
autopdf/
├── main.py                 # Punto de entrada
├── config.ini              # Configuración persistente
└── src/
    ├── config.py           # Constantes y estilos
    ├── config_manager.py   # Gestión de config.ini
    ├── utils.py            # Utilidades generales
    ├── file_manager.py     # Operaciones de archivos
    ├── word_processor.py   # Procesamiento Word/PDF
    ├── gui.py              # Interfaz gráfica (Tkinter)
    └── controller.py       # Lógica de negocio (MVC)
```

## Uso

1. **Configurar opciones** de formato mediante checkboxes:
   - **Formato de página**:  [ ] Logo,  [ ] líneas decorativas,  [ ] autor,  [ ] código,  [ ] numeración
   - **Opciones de carpeta**:  [ ] Respetar estructura,  [ ] copiar anexos,  [ ] convertir a PDF
   - **Extensiones**: Archivos `.docx` y `.docm`

2. **Seleccionar carpetas de origen** mediante drag & drop o botón de selección

3. **Especificar carpeta de destino** para archivos procesados

4. Presionar **EMPEZAR** para iniciar el procesamiento

> **⚠️ Importante**: La aplicación requiere que Microsoft Word esté cerrado antes de iniciar el procesamiento.

## Flujo de Trabajo

1. Usuario configura logo, autor, palabras prohibidas y carpetas
2. Al pulsar EMPEZAR, la app verifica que Word esté cerrado
3. Procesa cada carpeta según las opciones seleccionadas
4. Genera resultados en carpeta destino

## Roadmap / TO DO

### Alta Prioridad
- [ ] **Refactorización**: Dividir archivos complejos (`word_processor.py`, `gui.py`, `controller.py`) en funciones modulares y subcarpetas

### Funcionalidades Pendientes
- [X] **Config.ini**: Guardar y cargar última carpeta destino automáticamente
- [ ] **Sistema de renombre automático**: Basado en código de carpeta para reorganización masiva
- [ ] **Exportar log**: Generar archivo `.txt` externo con historial de procesamiento
- [ ] **Arrastrar** carpeta de destino además del botón de buscarlo.

### Bugs Conocidos
- [ ] **Código de carpeta**: Se copia nombre completo en lugar de solo el código extraído

## Licencia

Proyecto en desarrollo.

---

**Pink Autoheader** - Automatización de documentos simplificada