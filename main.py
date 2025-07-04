import sys
import os
import shutil
import glob
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem, QTabWidget,
    QTextEdit, QToolBar, QAction, QFileDialog, QWidget, QHBoxLayout, QVBoxLayout,
    QStatusBar, QTabBar, QPushButton, QComboBox, QLabel, QScrollArea, QGridLayout
)
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtCore import Qt, QSize
from ppadb.client import Client as AdbClient
import qdarkstyle

class FixedWidthTabBar(QTabBar):
    def tabSizeHint(self, index):
        return QSize(200, self.sizeHint().height())

class VSCodeStyleApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Mobile Forensic Triage Tool")
        self.setGeometry(200, 100, 1100, 650)
        self.setFont(QFont("Segoe UI", 10))

        self.device = None
        self.devices_map = {}

        self.temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "TempData")
        os.makedirs(self.temp_dir, exist_ok=True)

        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)
        content_layout = QHBoxLayout()
        main_layout.addLayout(content_layout)

        self.device_dropdown = QComboBox()
        self.device_dropdown.setPlaceholderText("Select a device")
        main_layout.addWidget(self.device_dropdown, alignment=Qt.AlignCenter)
        self.device_dropdown.setVisible(False)

        self.connect_button = QPushButton("Connect")
        self.connect_button.clicked.connect(self.connect_device)
        main_layout.addWidget(self.connect_button, alignment=Qt.AlignCenter)

        self.sidebarTree = QTreeWidget()
        self.sidebarTree.setHeaderHidden(True)
        self.sidebarTree.setMaximumWidth(220)
        self.sidebarTree.itemClicked.connect(self.open_or_focus_tab)
        content_layout.addWidget(self.sidebarTree)

        self.previewTabs = QTabWidget()
        self.previewTabs.setTabsClosable(True)
        self.previewTabs.tabCloseRequested.connect(self.previewTabs.removeTab)
        self.previewTabs.setTabBar(FixedWidthTabBar())
        content_layout.addWidget(self.previewTabs)

        self.toolbar = QToolBar("Main Toolbar")
        self.toolbar.setMovable(False)
        self.addToolBar(Qt.TopToolBarArea, self.toolbar)

        open_action = QAction("Open File", self)
        open_action.triggered.connect(self.open_file)
        export_action = QAction("Export", self)
        export_action.triggered.connect(self.export_data)
        settings_action = QAction("Settings", self)
        disconnect_action = QAction("Disconnect", self)
        disconnect_action.triggered.connect(self.disconnect_device)

        self.toolbar.addAction(open_action)
        self.toolbar.addAction(export_action)
        self.toolbar.addAction(settings_action)
        self.toolbar.addAction(disconnect_action)

        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Not Connected")

        self.sidebarTree.setVisible(False)
        self.previewTabs.setVisible(False)
        self.toolbar.setVisible(False)
        self.statusBar.setVisible(False)

        self.setCentralWidget(central_widget)
        self.populate_device_list()

    def populate_device_list(self):
        try:
            client = AdbClient(host="127.0.0.1", port=5037)
            devices = client.devices()
            self.devices_map = {d.serial: d for d in devices}
            self.device_dropdown.clear()

            if not devices:
                self.statusBar.setVisible(True)
                self.statusBar.showMessage("No devices found.")
                self.device_dropdown.setVisible(False)
                return

            self.device_dropdown.addItems([d.serial for d in devices])
            self.device_dropdown.setVisible(True)
            self.statusBar.setVisible(True)
            self.statusBar.showMessage("Select a device and click Connect")

        except Exception as e:
            self.statusBar.setVisible(True)
            self.statusBar.showMessage(f"Error fetching devices: {e}")

    def connect_device(self):
        deviceselect = self.device_dropdown.currentText().strip()
        if not deviceselect:
            self.statusBar.showMessage("Please select a device first.")
            return

        try:
            client = AdbClient(host="127.0.0.1", port=5037)
            live_devices = {d.serial: d for d in client.devices()}

            if deviceselect not in live_devices:
                self.statusBar.showMessage("Selected device is no longer connected.")
                self.populate_device_list()
                return

            self.device = live_devices[deviceselect]
            self.statusBar.showMessage(f"Connected to {self.device.serial}")

            info = {
                "Model": self.device.shell("getprop ro.product.model").strip(),
                "Manufacturer": self.device.shell("getprop ro.product.manufacturer").strip(),
                "Android Version": self.device.shell("getprop ro.build.version.release").strip(),
                "Device Name": self.device.shell("getprop ro.product.device").strip(),
                "Serial Number": self.device.shell("getprop ro.serialno").strip(),
                "CPU ABI": self.device.shell("getprop ro.product.cpu.abi").strip()
            }

            self.sidebarTree.clear()
            device_root = QTreeWidgetItem(["Connected Device"])
            for key, value in info.items():
                child = QTreeWidgetItem([f"{key}: {value}"])
                device_root.addChild(child)
            self.sidebarTree.addTopLevelItem(device_root)

            for section in ["Call Logs", "SMS", "WhatsApp", "Photos", "Videos", "Contacts"]:
                item = QTreeWidgetItem([section])
                self.sidebarTree.addTopLevelItem(item)

            self.sidebarTree.setVisible(True)
            self.previewTabs.setVisible(True)
            self.toolbar.setVisible(True)
            self.connect_button.setVisible(False)
            self.device_dropdown.setVisible(False)

        except Exception as e:
            self.statusBar.setVisible(True)
            self.statusBar.showMessage(f"Connection failed: {e}")

    def disconnect_device(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)
        os.makedirs(self.temp_dir, exist_ok=True)
        self.sidebarTree.clear()
        self.previewTabs.clear()
        self.sidebarTree.setVisible(False)
        self.previewTabs.setVisible(False)
        self.toolbar.setVisible(False)
        self.statusBar.setVisible(False)
        self.connect_button.setVisible(True)
        self.device_dropdown.setVisible(True)
        self.statusBar.showMessage("Disconnected")
        self.populate_device_list()

    def open_or_focus_tab(self, item):
        title = item.text(0)
        parent = item.parent()

        if parent and parent.text(0) == "Photos":
            img_path = os.path.join(self.temp_dir, "Photos", title)
            if os.path.exists(img_path):
                self.open_image_preview(title, img_path)
            return

        for i in range(self.previewTabs.count()):
            if self.previewTabs.tabText(i) == title:
                self.previewTabs.setCurrentIndex(i)
                return

        if title == "Photos":
            try:
                photo_dir = os.path.join(self.temp_dir, "Photos")
                os.makedirs(photo_dir, exist_ok=True)
                for f in glob.glob(os.path.join(photo_dir, "*")):
                    try:
                        os.remove(f)
                    except:
                        pass

                raw = self.device.shell("ls -R /sdcard")
                lines = raw.splitlines()
                current_dir = "/sdcard"
                file_paths = []

                for line in lines:
                    line = line.strip()
                    if not line:
                        continue
                    if line.endswith(":"):
                        current_dir = line[:-1].strip()
                        if not current_dir.startswith("/"):
                            current_dir = f"/{current_dir}"
                        continue
                    for part in line.split():
                        if any(part.lower().endswith(ext) for ext in [".jpg", ".jpeg", ".png"]):
                            file_paths.append(f"{current_dir}/{part}")

                downloaded_files = []
                for path in file_paths:
                    filename = os.path.basename(path)
                    dest = os.path.join(photo_dir, filename)
                    try:
                        self.device.pull(path, dest)
                        downloaded_files.append(dest)
                    except:
                        pass

                for i in range(self.sidebarTree.topLevelItemCount()):
                    item = self.sidebarTree.topLevelItem(i)
                    if item.text(0) == "Photos":
                        item.takeChildren()
                        for img_path in downloaded_files:
                            child = QTreeWidgetItem([os.path.basename(img_path)])
                            item.addChild(child)
                        item.setExpanded(True)
                        break

            except Exception as e:
                self.open_tab(title, f"Error loading photos: {e}")

        elif title == "Call Logs":
            content = self.device.shell("content query --uri content://call_log/calls")
            self.open_tab(title, content)

        elif title == "SMS":
            content = self.device.shell("content query --uri content://sms/")
            self.open_tab(title, content)

        elif title == "Contacts":
            content = self.device.shell("content query --uri content://contacts/phones/")
            self.open_tab(title, content)

        elif title == "WhatsApp":
            self.open_tab(title, "Preview not available. Use 'Export' to pull /sdcard/WhatsApp/Media/.")

        elif title == "Videos":
            self.open_tab(title, "Preview not available. Use 'Export' to pull /sdcard/DCIM/Camera/.")

        else:
            self.open_tab(title, f"Preview content for: {title}")

    def open_image_preview(self, title, img_path):
        for i in range(self.previewTabs.count()):
            if self.previewTabs.tabText(i) == title:
                self.previewTabs.setCurrentIndex(i)
                return

        label = QLabel()
        pixmap = QPixmap(img_path).scaled(800, 600, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        label.setPixmap(pixmap)

        scroll = QScrollArea()
        scroll.setWidget(label)
        scroll.setWidgetResizable(True)

        index = self.previewTabs.addTab(scroll, title)
        self.previewTabs.setCurrentIndex(index)

    def open_tab(self, title, content):
        for i in range(self.previewTabs.count()):
            if self.previewTabs.tabText(i) == title:
                self.previewTabs.setCurrentIndex(i)
                return
        editor = QTextEdit()
        editor.setText(content)
        editor.setReadOnly(True)
        index = self.previewTabs.addTab(editor, title)
        self.previewTabs.setCurrentIndex(index)

    def open_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open File")
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
                self.open_tab(os.path.basename(file_path), content)
            except Exception as e:
                self.open_tab("Error", f"Could not open file: {e}")

    def export_data(self):
        current_tab = self.previewTabs.tabText(self.previewTabs.currentIndex())
        folder = QFileDialog.getExistingDirectory(self, "Select export folder")
        if not folder:
            return

        try:
            if current_tab == "Photos":
                self.device.shell("ls -R /sdcard/Pictures/")
                self.device.pull("/sdcard/Pictures/", folder)
            elif current_tab == "WhatsApp":
                self.device.pull("/sdcard/WhatsApp/Media/", folder)
            elif current_tab == "Call Logs":
                data = self.device.shell("content query --uri content://call_log/calls")
                with open(os.path.join(folder, "call_logs.txt"), "w", encoding="utf-8") as f:
                    f.write(data)
            elif current_tab == "SMS":
                data = self.device.shell("content query --uri content://sms/")
                with open(os.path.join(folder, "sms.txt"), "w", encoding="utf-8") as f:
                    f.write(data)
            elif current_tab == "Contacts":
                data = self.device.shell("content query --uri content://contacts/phones/")
                with open(os.path.join(folder, "contacts.txt"), "w", encoding="utf-8") as f:
                    f.write(data)
            else:
                self.statusBar.showMessage("No export action for this tab.")
                return

            self.statusBar.showMessage(f"Exported {current_tab} to {folder}")
        except Exception as e:
            self.statusBar.showMessage(f"Export failed: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    window = VSCodeStyleApp()
    window.show()
    sys.exit(app.exec_())
