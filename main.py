import sys
import os
import shutil
import glob
import re
import json
from datetime import datetime

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem, QTabWidget,
    QTextEdit, QToolBar, QAction, QFileDialog, QWidget, QHBoxLayout, QVBoxLayout,
    QTableWidgetItem, QStatusBar, QTabBar, QPushButton, QComboBox, QLabel,
    QScrollArea, QSplitter, QTableWidget, QLineEdit, QTableView, QRadioButton,
    QButtonGroup, QGroupBox, QMessageBox, QListWidget, QListWidgetItem
)
import piexif
from PIL import Image
from PyQt5.QtGui import QFont, QPixmap, QStandardItemModel, QStandardItem
from PyQt5.QtCore import Qt, QSize, QSortFilterProxyModel, QPropertyAnimation, QRect, QEasingCurve

from ppadb.client import Client as AdbClient
import qdarkstyle
from langchain_openai import ChatOpenAI
from dotenv import load_dotenv

load_dotenv()

# ---------------------------- Usage Stats helpers ----------------------------
USAGE_EVENT_RE = re.compile(r'time="([^"]+)"\s+type=([A-Z_]+)\s+package=([\w\.\d]+)(.*)')


class UsageStatsWidget(QWidget):
    """Embedded widget for the 'Usage Stats' tab. Auto-pulls via ADB and shows a filterable table.
       Saves usage_dump.txt under TempData/UsageStats."""
    def __init__(self, adb_device, temp_root):
        super().__init__()
        self.device = adb_device
        self.out_dir = os.path.join(temp_root, "UsageStats")  # <<< inside TempData
        os.makedirs(self.out_dir, exist_ok=True)
        self.local_file = os.path.join(self.out_dir, "usage_dump.txt")

        self.model = QStandardItemModel()
        self.model.setHorizontalHeaderLabels(["Time", "Event Type", "Package", "Extra Info"])

        self.proxy = ParameterFilterProxyModel()
        self.proxy.setSourceModel(self.model)

        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.table.setSortingEnabled(True)

        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search (e.g. event=RESUMED package=whatsapp time=2025-08)")
        self.search_bar.textChanged.connect(self.apply_filters)

        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.clicked.connect(self.refresh_usage_stats)

        top = QHBoxLayout()
        top.addWidget(QLabel("Filter:"))
        top.addWidget(self.search_bar, stretch=1)
        top.addWidget(self.refresh_btn)

        layout = QVBoxLayout(self)
        layout.addLayout(top)
        layout.addWidget(self.table)

        # initial load
        self.refresh_usage_stats()

    def refresh_usage_stats(self):
        try:
            # dump and pull
            self.device.shell('sh -c "dumpsys usagestats > /sdcard/usage_dump.txt"')
            self.device.pull("/sdcard/usage_dump.txt", self.local_file)
            events = parse_usage_events(self.local_file)
            self.populate_table(events)
        except Exception as e:
            self.model.removeRows(0, self.model.rowCount())
            self.model.appendRow([
                QStandardItem("ERROR"),
                QStandardItem("PULL_FAILED"),
                QStandardItem("adb/usagestats"),
                QStandardItem(str(e))
            ])

    def populate_table(self, events):
        self.model.removeRows(0, self.model.rowCount())
        for ev in events:
            self.model.appendRow([
                QStandardItem(ev["time"]),
                QStandardItem(ev["event_type"]),
                QStandardItem(ev["package"]),
                QStandardItem(ev["extra_info"])
            ])

    def apply_filters(self, text):
        filters = {}
        for part in text.split():
            if "=" in part:
                k, v = part.split("=", 1)
                filters[k.strip().lower()] = v.strip()
        self.proxy.set_filters(filters)


# ---------------------------- Chat Sidebar (fixed sliding panel) ----------------------------
class ChatSidebar(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_app = parent  
        self.setFixedWidth(0)  # hidden initially; width is managed by main window
        self.setStyleSheet(
            """
            QWidget { background: #f6f7fb; }
            QTextEdit { background: white; border: 1px solid #ddd; }
            QLineEdit { background: white; border: 1px solid #ddd; padding: 4px; }
            QPushButton { padding: 6px 10px; }
            """
        )
        self.root = QVBoxLayout(self)
        self.root.setContentsMargins(10, 10, 10, 10)
        self.root.setSpacing(8)

        # Create both UIs and toggle between them
        self._build_not_configured_card()
        self._build_chat_ui()

        # default to not configured until main app updates
        self.show_not_configured()

    def _build_not_configured_card(self):
        self.not_configured_widget = QWidget()
        v = QVBoxLayout(self.not_configured_widget)
        v.setSpacing(8)
        v.addWidget(QLabel("<b>Agent not configured</b>"))
        msg = QLabel("The AI agent is not configured. Please configure it in Settings.")
        msg.setWordWrap(True)
        v.addWidget(msg)
        btn = QPushButton("Open Settings")
        btn.clicked.connect(self._open_settings_request)
        v.addWidget(btn)
        v.addStretch(1)

    def _build_chat_ui(self):
        self.chat_widget = QWidget()
        v = QVBoxLayout(self.chat_widget)
        v.setContentsMargins(0, 0, 0, 0)
        title = QLabel("Chat")
        title.setStyleSheet("font-weight: 600;")
        self.chat_history = QTextEdit()
        self.chat_history.setReadOnly(True)

        self.input_line = QLineEdit()
        self.input_line.setPlaceholderText("Type a message...")
        self.send_btn = QPushButton("Send")
        self.send_btn.clicked.connect(self.echo_message)

        row = QHBoxLayout()
        row.addWidget(self.input_line, 1)
        row.addWidget(self.send_btn)

        v.addWidget(title)
        v.addWidget(self.chat_history, 1)
        v.addLayout(row)

    def _open_settings_request(self):
        if self.main_app:
            self.main_app.show_settings_page()

    def echo_message(self):
        text = self.input_line.text().strip()
        if not text:
            return
        self.chat_history.append(f"<b>You:</b> {text}")
        self.chat_history.append(f"<b>AI:</b> :) Received: {text}")
        self.input_line.clear()

    def show_not_configured(self):
        self._set_content_widget(self.not_configured_widget)

    def show_chat_ui(self):
        self._set_content_widget(self.chat_widget)

    def _set_content_widget(self, widget):
        while self.root.count():
            item = self.root.takeAt(0)
            w = item.widget()
            if w:
                w.setParent(None)
        # Add requested widget
        self.root.addWidget(widget)


# ---------------------------- Main app helpers ----------------------------
def parse_usage_events(file_path):
    events = []
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            for line in f:
                m = USAGE_EVENT_RE.search(line)
                if m:
                    events.append({
                        "time": m.group(1),
                        "event_type": m.group(2),
                        "package": m.group(3),
                        "extra_info": (m.group(4) or "").strip()
                    })
    except Exception:
        pass
    return events


class ParameterFilterProxyModel(QSortFilterProxyModel):
    def __init__(self):
        super().__init__()
        self.filters = {}

    def set_filters(self, filters):
        self.filters = filters
        self.invalidateFilter()

    def filterAcceptsRow(self, source_row, source_parent):
        model = self.sourceModel()
        colm = {"time": 0, "event": 1, "package": 2, "extra": 3}
        for key, value in self.filters.items():
            if key in colm:
                idx = model.index(source_row, colm[key], source_parent)
                c_txt = model.data(idx) or ""
                if value.lower() not in c_txt.lower():
                    return False
        return True


class FixedWidthTabBar(QTabBar):
    def tabSizeHint(self, index):
        return QSize(200, self.sizeHint().height())


# ---------------------------- Settings Page (main content) ----------------------------
class SettingsPage(QWidget):
    """
    Settings page shown in the central content area (replaces previewTabs).
    - Theme selector (from QSS/ folder)
    - API configuration (Online / Local)
    - Save / Back
    """
    def __init__(self, main_app):
        super().__init__()
        self.main_app = main_app
        self.project_root = main_app.project_root
        self.config_path = main_app.config_path

        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(12, 12, 12, 12)

        # Theme selector
        theme_box = QGroupBox("Theme")
        tlay = QHBoxLayout()
        self.theme_dropdown = QComboBox()
        self._refresh_theme_list()
        self.apply_theme_btn = QPushButton("Apply Theme")
        self.apply_theme_btn.clicked.connect(self.apply_theme_clicked)
        tlay.addWidget(self.theme_dropdown, 1)
        tlay.addWidget(self.apply_theme_btn)
        theme_box.setLayout(tlay)
        layout.addWidget(theme_box)

        # API config
        api_box = QGroupBox("AI Agent Configuration")
        apil = QVBoxLayout()

        # mode radio
        mode_row = QHBoxLayout()
        self.online_radio = QRadioButton("Online (OpenAI API Key)")
        self.local_radio = QRadioButton("Local (LM Studio)")
        self.mode_group = QButtonGroup(self)
        self.mode_group.addButton(self.online_radio)
        self.mode_group.addButton(self.local_radio)
        mode_row.addWidget(self.online_radio)
        mode_row.addWidget(self.local_radio)
        apil.addLayout(mode_row)

        # online fields
        self.online_key_input = QLineEdit()
        self.online_key_input.setPlaceholderText("Enter OpenAI API key (sk-...)")
        apil.addWidget(self.online_key_input)

        # local fields
        local_group = QWidget()
        local_layout = QVBoxLayout(local_group)
        self.local_url_input = QLineEdit()
        self.local_url_input.setPlaceholderText("LM Studio base URL (e.g. http://127.0.0.1:8080)")
        self.local_model_input = QLineEdit()
        self.local_model_input.setPlaceholderText("Model name (e.g. ggml-model.bin)")
        local_layout.addWidget(self.local_url_input)
        local_layout.addWidget(self.local_model_input)
        apil.addWidget(local_group)

        # status label
        self.status_label = QLabel("")
        apil.addWidget(self.status_label)

        api_box.setLayout(apil)
        layout.addWidget(api_box)

        # Save / Back buttons
        btn_row = QHBoxLayout()
        self.save_btn = QPushButton("Save & Test Connection")
        self.save_btn.clicked.connect(self.save_and_test)
        self.back_btn = QPushButton("Back")
        self.back_btn.clicked.connect(self.on_back)
        btn_row.addStretch(1)
        btn_row.addWidget(self.save_btn)
        btn_row.addWidget(self.back_btn)
        layout.addLayout(btn_row)

        # Populate with current config if present
        self.load_config_to_ui()

        # Wire radio toggles to show/hide appropriate fields
        self.online_radio.toggled.connect(self._update_field_visibility)
        self._update_field_visibility()

    def _refresh_theme_list(self):
        qss_dir = os.path.join(self.project_root, "QSS")
        files = []
        if os.path.isdir(qss_dir):
            for f in sorted(os.listdir(qss_dir)):
                if f.lower().endswith(".qss"):
                    files.append(f)
        self.theme_dropdown.clear()
        self.theme_dropdown.addItem("Default (None)")
        for f in files:
            self.theme_dropdown.addItem(f)

    def apply_theme_clicked(self):
        sel = self.theme_dropdown.currentText()
        if sel == "Default (None)":
            QApplication.instance().setStyleSheet("")  # clear stylesheet
            self.main_app.current_theme = None
            self.main_app.statusBar.showMessage("Theme cleared")
            return
        qss_path = os.path.join(self.project_root, "QSS", sel)
        if os.path.exists(qss_path):
            try:
                with open(qss_path, "r", encoding="utf-8") as fh:
                    qss = fh.read()
                QApplication.instance().setStyleSheet(qss)
                self.main_app.current_theme = sel
                self.main_app.statusBar.showMessage(f"Applied theme: {sel}")
            except Exception as e:
                QMessageBox.warning(self, "Theme Error", f"Failed to apply theme: {e}")
        else:
            QMessageBox.warning(self, "Theme Error", "Selected theme file not found.")

    def _update_field_visibility(self):
        online = self.online_radio.isChecked()
        # online key visible if online, local inputs visible if local
        self.online_key_input.setVisible(online)
        self.local_url_input.setVisible(not online)
        self.local_model_input.setVisible(not online)

    def load_config_to_ui(self):
        # Load config if it exists and populate fields
        cfg = {}
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
        except Exception:
            cfg = {}
        mode = cfg.get("api_mode", "online")
        if mode == "local":
            self.local_radio.setChecked(True)
        else:
            self.online_radio.setChecked(True)

        self.online_key_input.setText(cfg.get("openai_key", ""))
        self.local_url_input.setText(cfg.get("host_url", ""))
        self.local_model_input.setText(cfg.get("model", ""))

        theme = cfg.get("theme")
        if theme:
            # refresh list and set current index
            self._refresh_theme_list()
            idx = self.theme_dropdown.findText(theme)
            if idx >= 0:
                self.theme_dropdown.setCurrentIndex(idx)

        self._update_field_visibility()

    def save_and_test(self):
        # Gather config
        api_mode = "local" if self.local_radio.isChecked() else "online"
        cfg = {
            "api_mode": api_mode,
            "openai_key": self.online_key_input.text().strip(),
            "host_url": self.local_url_input.text().strip(),
            "model": self.local_model_input.text().strip(),
            "theme": self.theme_dropdown.currentText() if self.theme_dropdown.currentText() != "Default (None)" else None
        }
        # Save to file
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
        except Exception as e:
            QMessageBox.warning(self, "Save Error", f"Failed to save config: {e}")
            return

        # Apply theme immediately if chosen
        if cfg.get("theme"):
            qss_path = os.path.join(self.project_root, "QSS", cfg["theme"])
            if os.path.exists(qss_path):
                try:
                    with open(qss_path, "r", encoding="utf-8") as fh:
                        QApplication.instance().setStyleSheet(fh.read())
                except Exception:
                    pass

        # Try a connection test (instantiation of ChatOpenAI)
        ok, message = self.main_app.test_connection(cfg)
        if ok:
            # success: green label, update status bar, auto-refresh chat sidebar
            self.status_label.setText(f"<font color='green'>{message}</font>")
            self.main_app.statusBar.showMessage(message)
            self.main_app.on_config_updated(True, cfg)
        else:
            self.status_label.setText(f"<font color='red'>{message}</font>")
            self.main_app.statusBar.showMessage("API Connection Failed")
            self.main_app.on_config_updated(False, cfg)

    def on_back(self):
        # navigate back to main preview UI
        self.main_app.hide_settings_page()


# ---------------------------- Main application ----------------------------
class DroidForen(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DroidForen")
        self.setGeometry(200, 100, 1100, 650)
        self.setFont(QFont("Segoe UI", 10))
        self.device = None
        self.devices_map = {}
        self._chat_open = False  # <-- track chat state

        self.ext_map = {
            "Photos": {".jpg", ".jpeg", ".png"},
            "Documents": {".pdf", ".doc", ".docx", ".txt", ".xls", ".xlsx", ".ppt", ".pptx"},
            "Videos": {".mp4", ".avi", ".mov", ".mkv", ".flv", ".wmv"},
            "Audio": {".mp3", ".wav", ".aac", ".flac", ".ogg", ".m4a"},
            "Archives": {
                ".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz",
                ".lzh", ".lha", ".ace", ".alz", ".cab", ".arj", ".cfs", ".dmg",
                ".xar", ".zst", ".wim", ".iso", ".shar", ".uue", ".b1", ".kgb", ".afa"
            }
        }
        # Added "Usage Stats" section
        self.SectionList = [
            "Call Logs", "SMS", "WhatsApp", "Photos", "Videos",
            "Audio", "Documents", "Contacts", "Archives", "Usage Stats"
        ]

        # file paths & project dirs
        self.project_root = os.path.dirname(os.path.abspath(__file__))
        self.temp_dir = os.path.join(self.project_root, "TempData")
        os.makedirs(self.temp_dir, exist_ok=True)

        # config path
        self.config_path = os.path.join(self.project_root, "config.json")
        self.current_theme = None

        # central content area container
        self.central_container = QWidget()
        self.setCentralWidget(self.central_container)
        self.main_layout = QVBoxLayout(self.central_container)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout = QHBoxLayout()
        self.main_layout.addLayout(self.content_layout)

        # Top controls (device dropdown + connect)
        self.device_dropdown = QComboBox()
        self.device_dropdown.setPlaceholderText("Select a device")
        self.main_layout.addWidget(self.device_dropdown, alignment=Qt.AlignCenter)
        self.device_dropdown.setVisible(False)

        self.connect_button = QPushButton("Connect")
        self.connect_button.clicked.connect(self.connect_device)
        self.main_layout.addWidget(self.connect_button, alignment=Qt.AlignCenter)

        # Left tree (sidebar)
        self.sidebarTree = QTreeWidget()
        self.sidebarTree.setHeaderHidden(True)
        self.sidebarTree.setMaximumWidth(220)
        self.sidebarTree.itemClicked.connect(self.open_or_focus_tab)
        self.content_layout.addWidget(self.sidebarTree)

        # Main preview tabs (center)
        self.previewTabs = QTabWidget()
        self.previewTabs.setTabsClosable(True)
        self.previewTabs.tabCloseRequested.connect(self.previewTabs.removeTab)
        self.previewTabs.setMovable(True)
        self.previewTabs.setTabBar(FixedWidthTabBar())
        self.content_layout.addWidget(self.previewTabs, 1)

        # Settings page widget (hidden by default). It will replace previewTabs when shown.
        self.settings_page = SettingsPage(self)
        self.settings_page.setVisible(False)
        # We'll add settings_page to the layout but keep it hidden; toggling visibility will display it
        self.content_layout.addWidget(self.settings_page, 1)

        # Toolbar
        self.toolbar = QToolBar("Main Toolbar")
        self.toolbar.setMovable(False)
        self.addToolBar(Qt.TopToolBarArea, self.toolbar)

        open_act = QAction("Open File", self)
        open_act.triggered.connect(self.open_file)
        export_act = QAction("Export", self)
        export_act.triggered.connect(self.export_data)
        # Settings action should show the settings page in-app (not a dialog)
        settings_act = QAction("Settings", self)
        settings_act.triggered.connect(self.show_settings_page)
        disconnect_act = QAction("Disconnect", self)
        disconnect_act.triggered.connect(self.disconnect_device)

        self.toolbar.addAction(open_act)
        self.toolbar.addAction(export_act)
        self.toolbar.addAction(settings_act)
        self.toolbar.addAction(disconnect_act)

        # Status bar
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Not Connected")

        # start hidden UI (until device connected)
        self.sidebarTree.setVisible(False)
        self.previewTabs.setVisible(False)
        self.toolbar.setVisible(False)
        self.statusBar.setVisible(False)

        self.populate_list()

        # Chat sidebar (right)
        self.chatSidebar = ChatSidebar(self)
        self.chat_width = 320
        self.chatSidebar.setFixedHeight(self.height())
        self.chatSidebar.move(self.width(), 0)

        self.chatToggleBtn = QPushButton("<", self)
        self.chatToggleBtn.setFixedWidth(28)
        self.chatToggleBtn.setFixedHeight(80)
        self.chatToggleBtn.clicked.connect(self.toggle_chat_sidebar)
        self.chatToggleBtn.raise_()
        self.position_chat_controls()

        self.chatAnim = QPropertyAnimation(self.chatSidebar, b"geometry")
        self.chatAnim.setDuration(250)
        self.chatAnim.setEasingCurve(QEasingCurve.InOutCubic)

        # Load config at startup and test
        self.loaded_config = {}
        self.load_config_and_test_on_startup()

    # ----------------- UI layout helpers -----------------
    def position_chat_controls(self):
        self.chatToggleBtn.move(self.width() - self.chatToggleBtn.width(), int((self.height() - self.chatToggleBtn.height()) / 2))
        if not self._chat_open:
            self.chatSidebar.move(self.width(), 0)
        else:
            self.chatSidebar.move(self.width() - self.chat_width, 0)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.chatSidebar.setFixedHeight(self.height())
        self.position_chat_controls()

    def toggle_chat_sidebar(self):
        self.chatSidebar.setFixedWidth(self.chat_width)
        self.chatAnim.stop()
        if self._chat_open:
            start = QRect(self.width() - self.chat_width, 0, self.chat_width, self.height())
            end = QRect(self.width(), 0, self.chat_width, self.height())
            self.chatAnim.setStartValue(start)
            self.chatAnim.setEndValue(end)
            self.chatAnim.start()
            self._chat_open = False
            self.chatToggleBtn.setText("<")
        else:
            start = QRect(self.width(), 0, self.chat_width, self.height())
            end = QRect(self.width() - self.chat_width, 0, self.chat_width, self.height())
            self.chatAnim.setStartValue(start)
            self.chatAnim.setEndValue(end)
            self.chatAnim.start()
            self._chat_open = True
            self.chatToggleBtn.setText(">")

    # ----------------- ADB/device logic -----------------
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
            self.statusBar.setVisible(True)
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

    def Extract(self, title):
        try:
            temp_sub_dir = os.path.join(self.temp_dir, title)
            os.makedirs(temp_sub_dir, exist_ok=True)
            for f in glob.glob(os.path.join(temp_sub_dir, "*")):
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
                    if any(part.lower().endswith(ext) for ext in self.ext_map[title]):
                        file_paths.append(f"{current_dir}/{part}")

            download_files = list()
            for path in file_paths:
                filename = os.path.basename(path)
                dest = os.path.join(temp_sub_dir, filename)
                try:
                    self.device.pull(path, dest)
                    download_files.append(dest)
                except:
                    pass
            for i in range(self.sidebarTree.topLevelItemCount()):
                item = self.sidebarTree.topLevelItem(i)
                if item.text(0) == title:
                    item.takeChildren()
                    for file_path in download_files:
                        child = QTreeWidgetItem([os.path.basename(file_path)])
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
        elif title == "Usage Stats":
            # Embedded Usage Stats tab (auto-pulls & shows table).
            usage_widget = UsageStatsWidget(self.device, self.temp_dir)  # <<< pass TempData
            idx = self.previewTabs.addTab(usage_widget, "Usage Stats")
            self.previewTabs.setCurrentIndex(idx)
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
        current_tab = None
        if self.previewTabs.count() > 0 and self.previewTabs.currentIndex() >= 0:
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
            elif current_tab == "Usage Stats":
                src = os.path.join(self.temp_dir, "UsageStats", "usage_dump.txt")  # <<< TempData/UsageStats
                if os.path.exists(src):
                    shutil.copy2(src, os.path.join(folder, "usage_dump.txt"))
                else:
                    self.statusBar.showMessage("No local usage_dump.txt to export.")
                    return
            else:
                self.statusBar.showMessage("No export action for this tab.")
                return

            self.statusBar.showMessage(f"Exported {current_tab} to {folder}")
        except Exception as e:
            self.statusBar.showMessage(f"Export failed: {str(e)}")

    # ----------------- Settings page & config integration -----------------
    def show_settings_page(self):
        # hide preview tabs and show settings page (replace center area)
        self.previewTabs.setVisible(False)
        self.settings_page.setVisible(True)
        # ensure left tree remains visible
        self.sidebarTree.setVisible(True)

    def hide_settings_page(self):
        # show back the normal preview UI
        self.settings_page.setVisible(False)
        self.previewTabs.setVisible(True)

    def load_config_and_test_on_startup(self):
        cfg = {}
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
        except Exception:
            cfg = {}

        self.loaded_config = cfg
        if cfg:
            ok, message = self.test_connection(cfg)
            if ok:
                self.statusBar.setVisible(True)
                self.statusBar.showMessage(message)
                self.chatSidebar.show_chat_ui()
                theme = cfg.get("theme")
                if theme:
                    qss_path = os.path.join(self.project_root, "QSS", theme)
                    if os.path.exists(qss_path):
                        try:
                            with open(qss_path, "r", encoding="utf-8") as fh:
                                QApplication.instance().setStyleSheet(fh.read())
                        except Exception:
                            pass
                self.current_theme = theme
                return
        # if we get here, not connected / no config
        self.chatSidebar.show_not_configured()
        self.statusBar.setVisible(True)
        if not cfg:
            self.statusBar.showMessage("Agent not configured")
        else:
            self.statusBar.showMessage("Agent connection failed")

    def test_connection(self, cfg):
        """
        Try to instantiate ChatOpenAI with the provided config.
        Returns (ok: bool, message: str)
        """
        if ChatOpenAI is None:
            return False, "ChatOpenAI import not available (install appropriate langchain package)"

        try:
            mode = cfg.get("api_mode", "online")
            model = cfg.get("model") or "gpt-4o-mini"
            if mode == "local":
                url = cfg.get("host_url", "").strip()
                if not url:
                    return False, "Local URL is empty"
                # instantiate ChatOpenAI with base_url
                # NOTE: depending on your langchain version the parameter names may differ.
                ChatOpenAI(api_key=cfg.get("openai_key", "") or os.getenv("API_KEY"), base_url=url, model=model)
                return True, f"Connected to {url}"
            else:
                key = cfg.get("openai_key", "").strip()
                if not key:
                    return False, "OpenAI API key is empty"
                ChatOpenAI(api_key=key, model=model)
                return True, "API Connected"
        except Exception as e:
            # return the error message for diagnostics
            return False, f"Connection failed: {e}"

    def on_config_updated(self, success, cfg):
        """
        Called by SettingsPage after save & test.
        If success == True -> update chat UI, statusbar, and remember config.
        """
        self.loaded_config = cfg
        if success:
            theme = cfg.get("theme")
            if theme:
                qss_path = os.path.join(self.project_root, "QSS", theme)
                if os.path.exists(qss_path):
                    try:
                        with open(qss_path, "r", encoding="utf-8") as fh:
                            QApplication.instance().setStyleSheet(fh.read())
                    except Exception:
                        pass
                self.current_theme = theme
            # update status bar and chat sidebar
            if cfg.get("api_mode") == "local":
                url = cfg.get("host_url", "")
                self.statusBar.showMessage(f"Connected to {url}")
            else:
                self.statusBar.showMessage("API Connected")
            self.chatSidebar.show_chat_ui()
        else:
            self.chatSidebar.show_not_configured()

    # ----------------- Application entry helpers -----------------


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # QSS Disabled for testing by default. The settings page can apply them.
    # app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())

    window = DroidForen()
    window.show()
    sys.exit(app.exec_())
