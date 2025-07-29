import os
import sys
import shutil
import ctypes
import win32com.client

from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton,
    QVBoxLayout, QFileDialog, QListWidget, QHBoxLayout, QMessageBox
)
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import Qt

DESKTOP = os.path.join(os.environ["USERPROFILE"], "Desktop")
ICONS_DIR = os.path.join(os.getcwd(), "icons")

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

class IconChanger(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Windows Icon Changer + Icons")
        self.setGeometry(100, 100, 400, 500)

        self.shortcut_list = QListWidget()
        self.shortcut_list.addItems(get_shortcuts(DESKTOP))

        self.icon_list = QListWidget()
        self.icon_list.addItems([f for f in os.listdir(ICONS_DIR) if f.endswith(".ico")])
        self.icon_list.setFixedHeight(120)

        self.preview_label = QLabel("Preview will appear here")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setFixedHeight(100)

        choose_button = QPushButton("Choose .ico File")
        choose_button.clicked.connect(self.choose_icon)

        self.tint_button = QPushButton("Tint Icon")
        self.tint_button.setEnabled(False)

        apply_button = QPushButton("Apply Icon to Shortcut")
        apply_button.clicked.connect(self.apply_icon)

        layout = QVBoxLayout()
        layout.addWidget(QLabel("Select a shortcut:"))
        layout.addWidget(self.shortcut_list)
        layout.addWidget(QLabel("Select an icon from folder:"))
        layout.addWidget(self.icon_list)
        layout.addWidget(self.preview_label)
        layout.addWidget(choose_button)
        layout.addWidget(self.tint_button)
        layout.addWidget(apply_button)
        self.setLayout(layout)

        self.icon_list.currentItemChanged.connect(self.show_preview)
        self.selected_icon_path = ""

    def choose_icon(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choose Icon", "", "Icon Files (*.ico)")
        if path:
            self.selected_icon_path = path
            self.show_icon(path)

    def show_preview(self):
        item = self.icon_list.currentItem()
        if item:
            icon_path = os.path.join(ICONS_DIR, item.text())
            self.selected_icon_path = icon_path
            self.show_icon(icon_path)

    def show_icon(self, path):
        pixmap = QPixmap(path)
        self.preview_label.setPixmap(pixmap.scaled(64, 64, Qt.KeepAspectRatio))

    def apply_icon(self):
        item = self.shortcut_list.currentItem()
        if not item or not self.selected_icon_path:
            QMessageBox.warning(self, "Error", "Please select both a shortcut and an icon.")
            return
        shortcut_path = os.path.join(DESKTOP, item.text())
        change_icon(shortcut_path, self.selected_icon_path)
        QMessageBox.information(self, "Success", "Icon applied successfully.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = IconChanger()
    window.show()
    sys.exit(app.exec_())
