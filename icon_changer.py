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

        layout = QtWidgets.QVBoxLayout()

        self.folder_label = QtWidgets.QLabel("Shortcut Folder:")
        self.folder_path = QtWidgets.QLineEdit()
        self.folder_browse = QtWidgets.QPushButton("Browse")
        self.folder_browse.clicked.connect(self.browse_folder)

        folder_layout = QtWidgets.QHBoxLayout()
        folder_layout.addWidget(self.folder_path)
        folder_layout.addWidget(self.folder_browse)

        self.icon_label = QtWidgets.QLabel("Icon File (.ico):")
        self.icon_path = QtWidgets.QLineEdit()
        self.icon_browse = QtWidgets.QPushButton("Browse")
        self.icon_browse.clicked.connect(self.browse_icon)

        icon_layout = QtWidgets.QHBoxLayout()
        icon_layout.addWidget(self.icon_path)
        icon_layout.addWidget(self.icon_browse)

        self.apply_button = QtWidgets.QPushButton("Apply Icon to All Shortcuts")
        self.apply_button.clicked.connect(self.apply_icons)

        layout.addWidget(self.folder_label)
        layout.addLayout(folder_layout)
        layout.addWidget(self.icon_label)
        layout.addLayout(icon_layout)
        layout.addWidget(self.apply_button)

        self.setLayout(layout)

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Shortcut Folder")
        if folder:
            self.folder_path.setText(folder)

    def browse_icon(self):
        icon_file, _ = QFileDialog.getOpenFileName(self, "Select Icon File", "", "Icon Files (*.ico)")
        if icon_file:
            self.icon_path.setText(icon_file)

    def apply_icons(self):
        folder = self.folder_path.text()
        icon = self.icon_path.text()

        if not os.path.isdir(folder):
            QMessageBox.warning(self, "Error", "Invalid shortcut folder.")
            return

        if not os.path.isfile(icon) or not icon.lower().endswith(".ico"):
            QMessageBox.warning(self, "Error", "Invalid .ico file.")
            return

        shortcuts = get_shortcuts(folder)
        if not shortcuts:
            QMessageBox.information(self, "Info", "No shortcuts found in the selected folder.")
            return

        for shortcut in shortcuts:
            path = os.path.join(folder, shortcut)
            try:
                change_icon(path, icon)
            except Exception as e:
                print(f"Failed to change icon for {path}: {e}")

        QMessageBox.information(self, "Success", "Icons applied to all shortcuts.")

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = IconChanger()
    window.show()
    sys.exit(app.exec_())
