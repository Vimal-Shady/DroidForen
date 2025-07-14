import sys
import os
import shutil
import glob
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem, QTabWidget,
    QTextEdit, QToolBar, QAction, QFileDialog, QWidget, QHBoxLayout, QVBoxLayout,QTableWidgetItem,
    QStatusBar, QTabBar, QPushButton, QComboBox, QLabel, QScrollArea,QSplitter,QTableWidget
)
import piexif
from PIL import Image
from PyQt5.QtGui import QIcon, QFont, QPixmap
from PyQt5.QtCore import Qt, QSize
from ppadb.client import Client as AdbClient
import qdarkstyle
from datetime import datetime

class FixedWidthTabBar(QTabBar):
    def tabSizeHint(self, index):
        return QSize(200, self.sizeHint().height())

class DroidForen(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Mobile Forensic Triage Tool")
        self.setGeometry(200, 100, 1100, 650)
        self.setFont(QFont("Segoe UI", 10))
        self.device = None
        self.devices_map={}

        #ExtensionMapping
        self.ext_map = {
            "Photos": {".jpg",".jpeg",".png"},
            "Documents": { ".pdf",".doc",".docx",".txt",".xls",".xlsx",".ppt",".pptx"},
            "Videos": { ".mp4",".avi",".mov",".mkv",".flv",".wmv"},
            "Audio": { ".mp3",".wav",".aac",".flac",".ogg",".m4a"},
            "Archives": {
                ".zip",".rar",".7z",".tar",".gz",".bz2",".xz",
                ".lzh",".lha",".ace",".alz",".cab",".arj",".cfs",".dmg",
                ".xar",".zst",".wim",".iso",".shar",".uue",".b1",".kgb",".afa"}
        }
        #Section List
        self.SectionList=["Call Logs", "SMS", "WhatsApp", "Photos", "Videos", "Audio", "Documents", "Contacts","Archives"]

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
        self.previewTabs.setMovable(True)
        self.previewTabs.setTabBar(FixedWidthTabBar())
        content_layout.addWidget(self.previewTabs)

        self.toolbar = QToolBar("Main Toolbar")
        self.toolbar.setMovable(False)
        self.addToolBar(Qt.TopToolBarArea, self.toolbar)

        open_act = QAction("Open File", self)
        open_act.triggered.connect(self.open_file)
        export_act = QAction("Export", self)
        export_act.triggered.connect(self.export_data)
        settings_act = QAction("Settings", self)
        disconnect_act = QAction("Disconnect", self)
        disconnect_act.triggered.connect(self.disconnect_device)

        self.toolbar.addAction(open_act)
        self.toolbar.addAction(export_act)
        self.toolbar.addAction(settings_act)
        self.toolbar.addAction(disconnect_act)

        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Not Connected")

        self.sidebarTree.setVisible(False)
        self.previewTabs.setVisible(False)
        self.toolbar.setVisible(False)
        self.statusBar.setVisible(False)

        self.setCentralWidget(central_widget)
        self.populate_list()

    # def closeEvent(self, event):
    #     shutil.rmtree(self.temp_dir, ignore_errors=True)
    #     event.accept()

    #Initial Single Threaded
    
    def populate_list(self):
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
                self.populate_list()
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
            self.sidebarTree.addTopLevelItems(QTreeWidgetItem([section]) for section in self.SectionList)
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
        self.populate_list()
    def Extract(self,title):
            try:
                temp_sub_dir=os.path.join(self.temp_dir, title)
                os.makedirs(temp_sub_dir, exist_ok=True)
                for f in glob.glob(os.path.join(temp_sub_dir, "*")):
                    try:
                        os.remove(f)
                    except:
                        pass
                raw=self.device.shell("ls -R /sdcard")
                lines=raw.splitlines()
                current_dir="/sdcard"
                file_paths=[]
                for line in lines:
                    line=line.strip()
                    if not line:
                        continue
                    if line.endswith(":"):
                        current_dir=line[:-1].strip()
                        if not current_dir.startswith("/"):
                            current_dir=f"/{current_dir}"
                        continue
                    for part in line.split():
                        if any(part.lower().endswith(ext) for ext in self.ext_map[title]):
                            file_paths.append(f"{current_dir}/{part}")

                download_files=list()
                for path in file_paths:
                    filename=os.path.basename(path)
                    dest=os.path.join(temp_sub_dir, filename)
                    try:
                        self.device.pull(path, dest)
                        download_files.append(dest)
                    except:
                        pass
                for i in range(self.sidebarTree.topLevelItemCount()):
                    item=self.sidebarTree.topLevelItem(i)
                    if item.text(0) == title:
                        item.takeChildren()
                        for file_path in download_files:
                            child=QTreeWidgetItem([os.path.basename(file_path)])
                            item.addChild(child)
                        item.setExpanded(True)
                        break
            except Exception as e:
                self.open_tab(title, f"Error loading {title}: {e}")
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
        if title == "Call Logs":
            self.show_call_logs()
        elif title == "SMS":
            content = self.device.shell("content query --uri content://sms/")
            self.open_tab(title, content)
        elif title == "Contacts":
            content = self.device.shell("content query --uri content://contacts/phones/")
            self.open_tab(title, content)
        elif title == "WhatsApp":
            self.open_tab(title, "Preview not available. Use 'Export' to pull /sdcard/WhatsApp/Media/.")
        else:
            self.Extract(title)

    def open_image_preview(self, title, img_path):
        for i in range(self.previewTabs.count()):
            if self.previewTabs.tabText(i) == title:
                self.previewTabs.setCurrentIndex(i)
                return

        pixmap = QPixmap(img_path)
        image_label = QLabel()
        image_label.setAlignment(Qt.AlignCenter)
        image_label.setPixmap(pixmap.scaled(600, 600, Qt.KeepAspectRatio, Qt.SmoothTransformation))

        exif_data = self.img_exif(img_path)
        exif_label = QLabel()
        exif_label.setAlignment(Qt.AlignTop)
        exif_label.setWordWrap(True)
        exif_text = "\n".join(f"{key}: {value}" for key, value in exif_data.items())
        exif_label.setText(exif_text)

        image_container = QWidget()
        image_layout = QVBoxLayout(image_container)
        image_layout.addWidget(image_label, alignment=Qt.AlignCenter)

        exif_container = QWidget()
        exif_layout = QVBoxLayout(exif_container)
        exif_layout.addWidget(QLabel("EXIF Metadata:"))
        exif_layout.addWidget(exif_label)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(image_container)
        splitter.addWidget(exif_container)
        splitter.setSizes([700, 400])

        scroll = QScrollArea()
        scroll.setWidget(splitter)
        scroll.setWidgetResizable(True)

        index = self.previewTabs.addTab(scroll, title)
        self.previewTabs.setCurrentIndex(index)

    def show_call_logs(self):
        try:
            raw = self.device.shell("content query --uri content://call_log/calls")
            entries = [e.strip() for e in raw.split("Row") if e.strip()]
            headers = ["Name", "Number", "Type", "Date", "Duration"]

            table = QTableWidget()
            table.setColumnCount(len(headers))
            table.setHorizontalHeaderLabels(headers)
            table.setRowCount(len(entries))

            for row_idx, entry in enumerate(entries):
                entry_dict = {}
                for part in entry.split(","):
                    if "=" in part:
                        key, val = part.strip().split("=", 1)
                        val = val.strip()
                        if val in ("NULL", ""):
                            val = "N/A"
                        entry_dict[key.strip()] = val

                name = entry_dict.get("name", "N/A")
                number = entry_dict.get("number", "N/A")
                call_type = self.call_type(entry_dict.get("type", "0"))
                date = self.format_date(entry_dict.get("date", "0"))
                duration = f"{entry_dict.get('duration', '0')} sec"

                table.setItem(row_idx, 0, QTableWidgetItem(name))
                table.setItem(row_idx, 1, QTableWidgetItem(number))
                table.setItem(row_idx, 2, QTableWidgetItem(call_type))
                table.setItem(row_idx, 3, QTableWidgetItem(date))
                table.setItem(row_idx, 4, QTableWidgetItem(duration))

            table.resizeColumnsToContents()
            index = self.previewTabs.addTab(table, "Call Logs")
            self.previewTabs.setCurrentIndex(index)

        except Exception as e:
            self.open_tab("Call Logs", f"Failed to load call logs: {e}")
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
    def call_type(self, call_type):
        mapping = {
            "1": "Incoming",
            "2": "Outgoing",
            "3": "Missed",
            "4": "Voicemail",
            "5": "Rejected",
            "6": "Blocked",
            "7": "Answered Externally"
        }
        return mapping.get(call_type, "Unknown")


    def format_date(self, timestamp):
        try:
            ts = int(timestamp)
            return datetime.fromtimestamp(ts / 1000).strftime("%Y-%m-%d %H:%M:%S")
        except:
            return "Invalid"

    def open_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open File")
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
                self.open_tab(os.path.basename(file_path), content)
            except Exception as e:
                self.open_tab("Error", f"Could not open file: {e}")
    def img_exif(self, image_path):
        try:
            image = Image.open(image_path)
            exif_bytes = image.info.get("exif", None)
            if not exif_bytes:
                return {"EXIF": "No EXIF data found."}

            exif_data = piexif.load(exif_bytes)
            exif_dict = {}
            exif_dict["Make"] = exif_data["0th"].get(piexif.ImageIFD.Make, b"").decode(errors="ignore")
            exif_dict["Model"] = exif_data["0th"].get(piexif.ImageIFD.Model, b"").decode(errors="ignore")
            exif_dict["Software"] = exif_data["0th"].get(piexif.ImageIFD.Software, b"").decode(errors="ignore")
            exif_dict["DateTime"] = exif_data["0th"].get(piexif.ImageIFD.DateTime, b"").decode(errors="ignore")
            exif_dict["Orientation"] = exif_data["0th"].get(piexif.ImageIFD.Orientation, "N/A")
            exif_dict["ExposureTime"] = exif_data["Exif"].get(piexif.ExifIFD.ExposureTime, "N/A")
            exif_dict["FNumber"] = exif_data["Exif"].get(piexif.ExifIFD.FNumber, "N/A")
            exif_dict["ISOSpeedRatings"] = exif_data["Exif"].get(piexif.ExifIFD.ISOSpeedRatings, "N/A")
            exif_dict["DateTimeOriginal"] = exif_data["Exif"].get(piexif.ExifIFD.DateTimeOriginal, b"").decode(errors="ignore")

            gps_data = exif_data.get("GPS", {})
            exif_dict["GPS"] = gps_data if gps_data else "Not Available"

            return exif_dict
        except Exception as e:
            return {"Error": str(e)}
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
    with open(r"DroidDark.qss", "r") as f:
        app.setStyleSheet(f.read())
    # app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    window = DroidForen()
    window.show()
    sys.exit(app.exec_())
