import configparser
import os
from typing import Tuple

class ConfigManager:
    """
    Manages configuration settings stored in an INI file format.
    """
    def __init__(self, config_file: str = 'settings.ini'):
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        self.defaults = {
            'folder_path': '',
            'auto_save': 'False',
            'auto_save_interval': '30',
            'backup_clean': 'False',
            'backup_clean_interval': '20'
        }
        self.load_config()

    def load_config(self) -> None:
        """
        Loads the configuration from the file, or initializes it with default values.
        """
        if not os.path.exists(self.config_file):
            self.config['Settings'] = self.defaults
            self.save_config()
        else:
            self.config.read(self.config_file)
            # Ensure all default settings are present
            for key, value in self.defaults.items():
                if not self.config.has_option('Settings', key):
                    self.config.set('Settings', key, value)
            self.save_config()

    def get_folder_path(self) -> str:
        return self.config.get('Settings', 'folder_path')

    def set_folder_path(self, path: str) -> None:
        self.config.set('Settings', 'folder_path', path)
        self.save_config()

    def get_auto_save_settings(self) -> Tuple[bool, int]:
        enabled = self.config.getboolean('Settings', 'auto_save')
        interval = self.config.getint('Settings', 'auto_save_interval')
        return enabled, interval

    def set_auto_save_settings(self, enabled: bool, interval: int) -> None:
        self.config.set('Settings', 'auto_save', str(enabled))
        self.config.set('Settings', 'auto_save_interval', str(interval))
        self.save_config()

    def get_backup_clean_settings(self) -> Tuple[bool, int]:
        enabled = self.config.getboolean('Settings', 'backup_clean')
        interval = self.config.getint('Settings', 'backup_clean_interval')
        return enabled, interval

    def set_backup_clean_settings(self, enabled: bool, interval: int) -> None:
        self.config.set('Settings', 'backup_clean', str(enabled))
        self.config.set('Settings', 'backup_clean_interval', str(interval))
        self.save_config()

    def save_config(self) -> None:
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)