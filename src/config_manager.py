import configparser
import os

class ConfigManager:
    """Gestiona la lectura y escritura del archivo config.ini en inglés"""

    def __init__(self, config_file='config.ini'):
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        self._load_defaults()
        self.load()

    def _load_defaults(self):
        """Define los valores por defecto en inglés"""
        self.config['USER'] = {
            'author': '',
            'last_logo': '',
            'last_destination': ''
        }
        self.config['HEADER_FOOTER'] = {
            'add_logo': 'True',
            'add_folder_code': 'True',
            'add_header_line': 'True',
            'add_footer_line': 'True',
            'add_author': 'True',
            'add_page_number': 'True'
        }
        self.config['COPY_OPTIONS'] = {
            'copy_to_destination': 'True',
            'respect_structure': 'True',
            'copy_attachments': 'True',
            'save_modified_in_dest': 'True',
            'copy_as_pdf': 'True'
        }
        self.config['PROCESS_EXTENSIONS'] = {
            'process_docx': 'True',
            'process_docm': 'False'
        }
        self.config['EXCLUSIONS'] = {
            'no_process_names': '',
            'no_copy_names': ''
        }

    def load(self):
        """Carga la configuración desde el archivo"""
        if os.path.exists(self.config_file):
            self.config.read(self.config_file, encoding='utf-8')
        else:
            self.save()

    def save(self):
        """Guarda la configuración actual en el archivo"""
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)

    # Getters y Setters genéricos
    def get_bool(self, section, key, default=False):
        return self.config.getboolean(section, key, fallback=default)

    def set_val(self, section, key, value):
        self.config.set(section, key, str(value))
        self.save()

    def get_str(self, section, key, default=''):
        return self.config.get(section, key, fallback=default)