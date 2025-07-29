iimport os
import sys
import ctypes
import shutil
import win32com.client

from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QTabWidget, QLabel, QPushButton, QListWidget

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
        self.setAcceptDrops(True)

        self.tabs = QTabWidget()
        self.tab_main = QtWidgets.QWidget()
        self.tab_download = QtWidgets.QWidget()
        self.tabs.addTab(self.tab_main, "Change Icons")
        self.tabs.addTab(self.tab_download, "Download Icons")

        self.icon_folder = ''
        self.shortcut_folder = ''

        self.setup_main_tab()
        self.setup_download_tab()

        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.tabs)
        self.setLayout(layout)

    def setup_main_tab(self):
        layout = QtWidgets.QVBoxLayout()

        self.shortcut_list = QListWidget()
        self.icon_list = QListWidget()

        browse_shortcuts = QPushButton("Choose Shortcut Folder")
        browse_shortcuts.clicked.connect(self.select_shortcut_folder)

        browse_icons = QPushButton("Choose Icon Folder")
        browse_icons.clicked.connect(self.select_icon_folder)

        apply_button = QPushButton("Apply Selected Icon to All Shortcuts")
        apply_button.clicked.connect(self.apply_icon_to_all_shortcuts)

        layout.addWidget(browse_shortcuts)
        layout.addWidget(self.shortcut_list)
        layout.addWidget(browse_icons)
        layout.addWidget(self.icon_list)
        layout.addWidget(apply_button)

        self.tab_main.setLayout(layout)

    def setup_download_tab(self):
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(QLabel("Online downloader coming soon..."))
        self.tab_download.setLayout(layout)

    def select_shortcut_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Shortcut Folder")
        if folder:
            self.shortcut_folder = folder
            self.shortcut_list.clear()
            for file in get_shortcuts(folder):
                self.shortcut_list.addItem(file)

    def select_icon_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Icon Folder")
        if folder:
            self.icon_folder = folder
            self.icon_list.clear()
            for file in os.listdir(folder):
                if file.endswith(".ico"):
                    self.icon_list.addItem(file)

    def apply_icon_to_all_shortcuts(self):
        icon_item = self.icon_list.currentItem()
        if not icon_item:
            QMessageBox.warning(self, "No Icon Selected", "Please select an icon to apply.")
            return

        icon_path = os.path.join(self.icon_folder, icon_item.text())
        for i in range(self.shortcut_list.count()):
            shortcut_path = os.path.join(self.shortcut_folder, self.shortcut_list.item(i).text())
            change_icon(shortcut_path, icon_path)

        QMessageBox.information(self, "Success", "Icons updated.")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.endswith(".ico") and os.path.isfile(path):
                dest = os.path.join(os.getcwd(), "icons", os.path.basename(path))
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                shutil.copy(path, dest)
                self.icon_list.addItem(os.path.basename(dest))

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = IconChanger()
    window.show()
    sys.exit(app.exec_())

