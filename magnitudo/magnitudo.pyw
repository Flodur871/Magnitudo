# -*- coding: utf-8 -*-
import configparser
import sys
import threading
import winreg
from pathlib import Path

from PySide2 import QtCore, QtWidgets, QtGui

from cntrl import Kc


class UI:
    def __init__(self, main_window):
        self.main_window = main_window

        self.gridLayout = QtWidgets.QGridLayout(main_window)

        self.Startup = QtWidgets.QCheckBox(main_window)

        self.primary_label = QtWidgets.QLabel(main_window)
        self.Primary = QtWidgets.QComboBox(main_window)

        self.secondary_label = QtWidgets.QLabel(main_window)
        self.Secondary = QtWidgets.QComboBox(main_window)

        self.buttonBox = QtWidgets.QDialogButtonBox(main_window)

        self.controller = Kc()

        self.config = configparser.ConfigParser()
        self.options = None
        self.new_options = None

        self.file_path = Path(__file__).resolve()
        self.dir_path = self.file_path.parent
        self.icon_path = f'{self.dir_path}\\icon.png'

        if getattr(sys, 'frozen', False):   # check if the app is running from as an executable
            self.file_path = Path(sys.executable).resolve()
            self.dir_path = self.file_path.parent

    def build(self):
        """
        setups the ui buttons
        :return:
        """
        self.main_window.setObjectName('Magnitudo')
        self.main_window.resize(100, 200)
        self.main_window.setWindowTitle('Magnitudo')

        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Ok)

        self.gridLayout.addWidget(self.Startup, 0, 0, 1, 2)

        self.gridLayout.addWidget(self.primary_label, 1, 0, 1, 1)
        self.gridLayout.addWidget(self.Primary, 1, 1, 1, 1)

        self.gridLayout.addWidget(self.secondary_label, 2, 0, 1, 1)
        self.gridLayout.addWidget(self.Secondary, 2, 1, 1, 1)

        self.gridLayout.addWidget(self.buttonBox, 3, 0, 1, 2)

        self.get_config()

        self.add_text()
        self.set_hooks()

        self.load_config()

    def load_config(self):
        """
        loads the config file settings
        :return:
        """
        self.Startup.setChecked(self.options.getboolean('startup'))  # sets the checkbox state based on the config

        self.Primary.setCurrentText(self.options['primary'])
        self.controller.set_primary(self.options['primary'])

        self.Secondary.setCurrentText(self.options['secondary'])
        self.controller.set_secondary(self.options['secondary'])

    def input_validator(self):
        """
        validates the config file
        :return:
        """
        if 'startup' not in self.options:
            self.options['startup'] = '0'
            self.update_config()

        elif self.options['startup'] != '1' and self.options['startup'] != '0':
            self.options['startup'] = '0'
            self.update_config()

        if 'primary' not in self.options:
            self.options['primary'] = list(self.controller.audio_devices)[0]
            self.update_config()

        elif self.options['primary'] not in list(self.controller.audio_devices):
            self.options['primary'] = list(self.controller.audio_devices)[0]
            self.update_config()

        if 'secondary' not in self.options:
            self.options['secondary'] = list(self.controller.audio_devices)[1]
            self.update_config()

        elif self.options['secondary'] not in list(self.controller.audio_devices):
            self.options['secondary'] = list(self.controller.audio_devices)[1]
            self.update_config()

    def get_config(self):
        """
        load the options from the config file if exists or creates a new file with default values
        :return:
        """
        settings = f'{self.dir_path}\\settings.ini'

        if Path(settings).is_file():
            self.config.read(settings)
            self.options = self.config['options']
            self.input_validator()

        else:
            self.new_config_file()

        self.new_options = dict(self.options)  # create a copy of current options to enable reversing changes

    def new_config_file(self):
        """
        create a config file from scratch with default values
        :return:
        """
        self.config['options'] = {}
        self.options = self.config['options']

        self.options['startup'] = '0'
        self.options['primary'] = list(self.controller.audio_devices)[0]
        self.options['secondary'] = list(self.controller.audio_devices)[1]

        self.update_config()

    def add_text(self):
        """
        adds the text to the buttons
        :return:
        """
        self.Startup.setText('Run On Startup')

        self.primary_label.setText('Primary Audio Source')
        self.secondary_label.setText('Secondary Audio Source')

        self.Primary.addItems(list(self.controller.audio_devices))

        self.Secondary.addItems(list(self.controller.audio_devices))

    def set_hooks(self):
        """
        hook each button to it's function
        :return:
        """
        self.Startup.toggled.connect(self.change_startup)  # startup got toggled

        self.Primary.activated.connect(self.change_primary)  # primary audio device changed
        self.Secondary.activated.connect(self.change_secondary)  # secondary audio device changed

        self.buttonBox.accepted.connect(self.accepted)  # ok button got pressed
        self.buttonBox.rejected.connect(self.rejected)  # cancel button got pressed

    def accepted(self):
        """
        save changes to the config file
        :return:
        """
        if self.new_options['startup'] != self.options['startup']:
            self.set_startup(self.new_options['startup'])

        self.controller.set_primary(self.Primary.currentText())
        self.controller.set_secondary(self.Secondary.currentText())

        self.options['startup'] = self.new_options['startup']
        self.options['primary'] = self.new_options['primary']
        self.options['secondary'] = self.new_options['secondary']

        self.update_config()
        self.main_window.close()  # close to background

    def rejected(self):
        """
        reloads settings from the config file
        :return:
        """
        self.new_options = dict(self.options)
        self.load_config()
        self.main_window.close()

    def update_config(self):
        """
        writes the changes to the config file
        :return:
        """
        with open(f'{self.dir_path}\\settings.ini', 'w') as configfile:
            self.config.write(configfile)

    def change_primary(self):
        """
        change primary audio source
        :return:
        """
        self.new_options['primary'] = self.Primary.currentText()

    def change_secondary(self):
        """
        changes secondary audio source
        :return:
        """
        self.new_options['secondary'] = self.Secondary.currentText()

    def change_startup(self, new_state):
        """
        loads new startup settings
        :param new_state: bool
        :return:
        """
        if new_state:
            self.new_options['startup'] = '1'
        else:
            self.new_options['startup'] = '0'

    def set_startup(self, new_state):
        """
        enables or disables startup based on the user selection
        :param new_state: '1' or '0'
        :return:
        """
        key_value = r'Software\Microsoft\Windows\CurrentVersion\Run'

        # open the key to make changes to
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_value, 0, winreg.KEY_ALL_ACCESS) as key:
            if new_state == '1':  # if the checkbox was enabled
                winreg.SetValueEx(key, 'Magnitudo', 0, winreg.REG_SZ,
                                  f'{self.file_path}')  # adds the file to startup

            else:
                winreg.DeleteValue(key, 'Magnitudo')


class SystemTrayIcon(QtWidgets.QSystemTrayIcon):
    def __init__(self, icon, parent):
        self.parent = parent
        QtWidgets.QSystemTrayIcon.__init__(self, icon, self.parent)
        self.setToolTip('Magnitudo')

        menu = QtWidgets.QMenu()
        menu.addAction('Options', self.options)
        menu.addSeparator()

        menu.addAction('Exit', self.exit)

        self.setContextMenu(menu)  # add the menu to the icon

        # Restore the window when the tray icon is double clicked.
        self.activated.connect(self.on_active)

    def on_active(self, press_type):
        """
        show the main window if the icon was double clicked
        :param press_type: the type of press on the icon
        :return:
        """
        if press_type == QtWidgets.QSystemTrayIcon.DoubleClick:  # if the icon got double clicked
            self.parent.show()  # show main window

    def options(self):
        """
        show the main window if the options button was pressed
        :return:
        """
        self.parent.show()

    @staticmethod
    def exit():
        """
        exit the program if the exit button was pressed
        :return:
        """
        QtCore.QCoreApplication.exit()


def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setAttribute(QtCore.Qt.AA_DisableWindowContextHelpButton)  # disables the question mark button
    app.setQuitOnLastWindowClosed(False)  # stop the app from quitting when it's closed

    main_window = QtWidgets.QDialog()

    ui = UI(main_window)
    ui.build()

    app.lastWindowClosed.connect(ui.accepted)  # make the x button apply changes

    thread = threading.Thread(target=ui.controller.run)
    thread.setDaemon(True)  # allow for smooth termination process
    thread.start()

    icon = QtGui.QIcon(ui.icon_path)

    main_window.setWindowIcon(icon)  # adds the icon to the gui

    tray_icon = SystemTrayIcon(icon, main_window)  # adds the icon to the tray
    tray_icon.show()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
