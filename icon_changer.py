import os
import sys
import ctypes
import shutil
import win32com.client
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import QFileDialog, QMessageBox

def get_shortcuts(folder):
    try:
        return [f for f in os.listdir(folder) if f.endswith(".lnk")]
    except FileNotFoundError:
        return []

def change_icon(shortcut_path, icon_path):
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(shortcut_path)
    shortcut.IconLocation = icon_path
    shortcut.Save()

class IconChanger(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Icon Changer")
        self.setGeometry(100, 100, 600, 400)
        self.setWindowIcon(QtGui.QIcon.fromTheme("applications-system"))

        layout = QtWidgets.QVBoxLayout(self)

        self.label = QtWidgets.QLabel("Select a shortcut:")
        layout.addWidget(self.label)

        self.list_widget = QtWidgets.QListWidget()
        layout.addWidget(self.list_widget)

        self.icon_path_input = QtWidgets.QLineEdit()
        self.icon_path_input.setPlaceholderText("Select an .ico file...")
        layout.addWidget(self.icon_path_input)

        self.browse_button = QtWidgets.QPushButton("Browse for Icon")
        self.browse_button.clicked.connect(self.browse_icon)
        layout.addWidget(self.browse_button)

        self.change_button = QtWidgets.QPushButton("Change Icon")
        self.change_button.clicked.connect(self.apply_icon_change)
        layout.addWidget(self.change_button)

        self.load_shortcuts()

    def get_desktop_path(self):
        try:
            return os.path.join(os.environ["USERPROFILE"], "Desktop")
        except KeyError:
            return os.getcwd()  # Fallback to current dir

    def load_shortcuts(self):
        desktop_path = self.get_desktop_path()
        shortcuts = get_shortcuts(desktop_path)
        self.list_widget.addItems(shortcuts)

    def browse_icon(self):
        icon_file, _ = QFileDialog.getOpenFileName(self, "Select Icon", "", "Icons (*.ico)")
        if icon_file:
            self.icon_path_input.setText(icon_file)

    def apply_icon_change(self):
        selected_items = self.list_widget.selectedItems()
        icon_path = self.icon_path_input.text()

        if not selected_items:
            QMessageBox.warning(self, "No Shortcut Selected", "Please select a shortcut from the list.")
            return

        if not os.path.isfile(icon_path):
            QMessageBox.warning(self, "Invalid Icon", "Please select a valid .ico file.")
            return

        desktop_path = self.get_desktop_path()
        for item in selected_items:
            shortcut_file = os.path.join(desktop_path, item.text())
            change_icon(shortcut_file, icon_path)

        QMessageBox.information(self, "Success", "Icon changed successfully.")

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = IconChanger()
    window.show()
    sys.exit(app.exec_())
