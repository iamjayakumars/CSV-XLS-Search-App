import sys
import os
import logging
from datetime import datetime

def check_dependencies():
    required_packages = {
        'pandas': 'pandas',
        'polars': 'polars',
        'PyQt6': 'PyQt6',
        'matplotlib': 'matplotlib',
        'seaborn': 'seaborn',
        'numpy': 'numpy',
        'openpyxl': 'openpyxl',
        'xlsxwriter': 'xlsxwriter',
        'reportlab': 'reportlab',
        'psutil': 'psutil'
    }
    
    missing_packages = []
    for package, import_name in required_packages.items():
        try:
            __import__(import_name)
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        error_msg = "Missing required packages. Please install:\n"
        error_msg += "\n".join(f"- {package}" for package in missing_packages)
        error_msg += "\n\nRun: pip install -r requirements.txt"
        
        if sys.platform == 'win32':
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, error_msg, "Missing Dependencies", 0x10)
        else:
            print(error_msg)
        sys.exit(1)

# Check dependencies before importing other modules
check_dependencies()

# Now import other modules
import pandas as pd
import polars as pl
import glob
import re
import json
from sqlalchemy import create_engine
from PyQt6.QtWidgets import QAbstractItemView, QProgressBar, QMenu, QSizePolicy
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import numpy as np
import traceback

# Try to import psutil, but don't fail if it's not available
try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False
    logging.warning("psutil module not available. Memory information will not be logged.")

# Configure logging with more detailed format
def setup_logging():
    # Create logs directory if it doesn't exist
    if not os.path.exists('logs'):
        os.makedirs('logs')
    
    # Create a unique log file name with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = f'logs/app_{timestamp}.log'
    
    # Configure logging format with more details
    logging.basicConfig(
        level=logging.DEBUG,  # Set to DEBUG for maximum logging
        format='%(asctime)s - %(levelname)s - [%(name)s] - %(message)s\n%(pathname)s:%(lineno)d\n',
        handlers=[
            logging.FileHandler(log_file, mode='a', encoding='utf-8'),
            logging.StreamHandler()  # Also print to console
        ]
    )
    
    # Create a logger instance
    logger = logging.getLogger(__name__)
    
    # Log system information in the background
    def log_system_info():
        try:
            logger.info("Application started")
            logger.info(f"Python version: {sys.version}")
            logger.info(f"Operating System: {sys.platform}")
            logger.info(f"Working Directory: {os.getcwd()}")
            
            if PSUTIL_AVAILABLE:
                try:
                    memory = psutil.virtual_memory()
                    logger.info(f"System Memory: {memory.total / (1024**3):.2f} GB")
                    logger.info(f"Available Memory: {memory.available / (1024**3):.2f} GB")
                    logger.info(f"Memory Usage: {memory.percent}%")
                except Exception as e:
                    logger.warning(f"Failed to log memory information: {str(e)}")
            
            # Log package versions
            logger.info(f"Pandas version: {pd.__version__}")
            logger.info(f"Polars version: {pl.__version__}")
            logger.info(f"PyQt6 version: {sys.modules['PyQt6'].__version__}")
            logger.info(f"Matplotlib version: {plt.__version__}")
            logger.info(f"Seaborn version: {sns.__version__}")
            logger.info(f"Numpy version: {np.__version__}")
        except Exception as e:
            logger.error(f"Failed to log system information: {str(e)}", exc_info=True)
    
    # Start logging in a background thread
    import threading
    thread = threading.Thread(target=log_system_info)
    thread.daemon = True
    thread.start()
    
    return logger

# Initialize logger
logger = setup_logging()

# Log system information
logger.info(f"Python version: {sys.version}")
logger.info(f"Operating System: {sys.platform}")
logger.info(f"Working Directory: {os.getcwd()}")

# Log memory information if psutil is available
if PSUTIL_AVAILABLE:
    try:
        memory = psutil.virtual_memory()
        logger.info(f"System Memory: {memory.total / (1024**3):.2f} GB")
        logger.info(f"Available Memory: {memory.available / (1024**3):.2f} GB")
        logger.info(f"Memory Usage: {memory.percent}%")
    except Exception as e:
        logger.warning(f"Failed to log memory information: {str(e)}")

# Log imported package versions
try:
    logger.info(f"Pandas version: {pd.__version__}")
    logger.info(f"Polars version: {pl.__version__}")
    logger.info(f"PyQt6 version: {sys.modules['PyQt6'].__version__}")
    logger.info(f"Matplotlib version: {plt.__version__}")
    logger.info(f"Seaborn version: {sns.__version__}")
    logger.info(f"Numpy version: {np.__version__}")
except Exception as e:
    logger.error(f"Failed to log package versions: {str(e)}", exc_info=True)

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLineEdit, QFileDialog,
    QTableView, QLabel, QComboBox, QCheckBox, QHBoxLayout, QMessageBox, QAbstractScrollArea, QHeaderView, QDialog, QScrollArea, QWidget as QScrollWidget
)
from PyQt6.QtCore import QAbstractTableModel, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QColor, QAction, QPalette

# -------------------- Model for Fast Table View Rendering -------------------- #
class PandasTableModel(QAbstractTableModel):
    def __init__(self, df):
        super().__init__()
        self.logger = logging.getLogger(__name__)
        self.logger.info("Initializing PandasTableModel")
        self._df = df
        self.highlight_patterns = []
        self.highlight_color = QColor(255, 255, 0, 100)
        self.duplicate_color = QColor(255, 200, 200, 100)
        self._sort_column = -1
        self._sort_order = Qt.SortOrder.AscendingOrder
        self.duplicate_columns = []
        self._cache = {}
        self._row_cache = {}
        self._column_cache = {}
        self.logger.info(f"PandasTableModel initialized with DataFrame shape: {df.shape}")

    def rowCount(self, parent=None):
        return self._df.shape[0]

    def columnCount(self, parent=None):
        return self._df.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        try:
            if not index.isValid():
                return None
            
            row, col = index.row(), index.column()
            
            if role == Qt.ItemDataRole.DisplayRole:
                    cache_key = (row, col)
                    if cache_key not in self._cache:
                        self._cache[cache_key] = str(self._df.iloc[row, col])
                    return self._cache[cache_key]
                    
            elif role == Qt.ItemDataRole.BackgroundRole:
                    value = str(self._df.iloc[row, col])
                    
                    for pattern in self.highlight_patterns:
                        if pattern in value:
                            return self.highlight_color
                    
                    if col in self.duplicate_columns:
                        return self.duplicate_color
                        
                    return None
                    
            return None
        except Exception as e:
            self.logger.error(f"Error in data method: {str(e)}", exc_info=True)
            return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                if section not in self._column_cache:
                    self._column_cache[section] = self._df.columns[section]
                return self._column_cache[section]
            else:
                if section not in self._row_cache:
                    self._row_cache[section] = str(section)
                return self._row_cache[section]
        return None

    def set_highlight_patterns(self, patterns):
        self.highlight_patterns = patterns
        self._cache.clear()  # Clear cache when highlighting changes
        self.layoutChanged.emit()

    def set_duplicate_columns(self, columns):
        self.duplicate_columns = columns
        self._cache.clear()  # Clear cache when duplicates change
        self.layoutChanged.emit()

    def sort(self, column, order=Qt.SortOrder.AscendingOrder):
        try:
            self.logger.debug(f"Sorting column {column} in {'ascending' if order == Qt.SortOrder.AscendingOrder else 'descending'} order")
            self._sort_column = column
            self._sort_order = order
            
            if column >= 0 and column < self._df.shape[1]:
                ascending = order == Qt.SortOrder.AscendingOrder
                self._df = self._df.sort_values(
                    by=self._df.columns[column],
                    ascending=ascending,
                    na_position='last'
                )
                self._cache.clear()
                self.layoutChanged.emit()
                self.logger.debug("Sort completed successfully")
        except Exception as e:
            self.logger.error(f"Error in sort method: {str(e)}", exc_info=True)

    def get_dataframe(self):
        return self._df.copy()

    def clear_cache(self):
        """Clear all caches when data changes"""
        self._cache.clear()
        self._row_cache.clear()
        self._column_cache.clear()

# -------------------- File Loader Thread (if needed) --------------------
class FileLoaderThread(QThread):
    data_loaded = pyqtSignal(pd.DataFrame, str)
    progress_updated = pyqtSignal(int)
    error_occurred = pyqtSignal(str)
    sheet_names_loaded = pyqtSignal(list)

    def __init__(self, file_path):
        super().__init__()
        self.logger = logging.getLogger(__name__)
        self.logger.info(f"Initializing FileLoaderThread for file: {file_path}")
        self.file_path = file_path
        self.chunk_size = 50000
        self.is_running = True

    def stop(self):
        self.is_running = False
        self.logger.info("FileLoaderThread stopped")

    def run(self):
        try:
            self.logger.info(f"Starting file load: {self.file_path}")
            self.logger.debug(f"File size: {os.path.getsize(self.file_path) / (1024**2):.2f} MB")
            
            if self.file_path.endswith(".csv"):
                self.logger.info("Processing CSV file")
                # Get total rows for progress calculation
                total_rows = sum(1 for _ in open(self.file_path, 'r')) - 1  # Subtract header row
                self.logger.info(f"Total rows in CSV: {total_rows}")
                
                # Read first chunk to get column names
                df = pd.read_csv(self.file_path, nrows=1, dtype=str)
                columns = df.columns
                self.logger.info(f"CSV columns: {', '.join(columns)}")
                
                # Initialize empty DataFrame with correct columns
                result_df = pd.DataFrame(columns=columns)
                
                # Read chunks and concatenate
                for i in range(0, total_rows, self.chunk_size):
                    if not self.is_running:
                        self.logger.info("File loading interrupted")
                        return
                        
                    chunk = pd.read_csv(
                        self.file_path,
                        skiprows=i + 1,  # Skip header row
                        nrows=self.chunk_size,
                        dtype=str,
                        low_memory=False
                    )
                    
                    result_df = pd.concat([result_df, chunk], ignore_index=True)
                    
                    # Update progress
                    progress = min(100, int((i + self.chunk_size) / total_rows * 100))
                    self.progress_updated.emit(progress)
                    self.logger.debug(f"CSV loading progress: {progress}%")
                    
            else:
                self.logger.info("Processing Excel file")
                # For Excel files, first get sheet names
                excel_file = pd.ExcelFile(self.file_path, engine="openpyxl")
                sheet_names = excel_file.sheet_names
                self.logger.info(f"Excel sheets found: {', '.join(sheet_names)}")
                
                # Emit sheet names first
                self.sheet_names_loaded.emit(sheet_names)
                
                # Read the first sheet by default
                result_df = pd.read_excel(
                    self.file_path,
                    sheet_name=0,
                    engine="openpyxl",
                    dtype=str,
                    na_filter=False,
                    keep_default_na=False
                )
                
                # Update progress for Excel
                self.progress_updated.emit(100)
                self.logger.info(f"Excel file loaded successfully: {self.file_path}")

            if not self.is_running:
                self.logger.info("File loading interrupted before completion")
                return

            # Validate DataFrame
            if result_df.empty:
                self.logger.error("Loaded DataFrame is empty")
                self.error_occurred.emit("The file is empty")
                return

            # Check for duplicate column names
            if len(result_df.columns) != len(set(result_df.columns)):
                self.logger.error("File contains duplicate column names")
                self.error_occurred.emit("File contains duplicate column names")
                return

            # Log success
            self.logger.info(f"File loaded successfully: {self.file_path}")
            self.logger.info(f"DataFrame shape: {result_df.shape}")
            self.data_loaded.emit(result_df, self.file_path)

        except pd.errors.EmptyDataError:
            self.logger.error("Empty file encountered")
            self.error_occurred.emit("The file is empty")
        except pd.errors.ParserError as e:
            self.logger.error(f"Parser error: {str(e)}")
            self.error_occurred.emit("Failed to parse the file. It may be corrupted or in an invalid format.")
        except PermissionError:
            self.logger.error("Permission denied when accessing file")
            self.error_occurred.emit("Permission denied. Please check file permissions.")
        except Exception as e:
            self.logger.error(f"Unexpected error during file loading: {str(e)}", exc_info=True)
            self.error_occurred.emit(f"Error loading file: {str(e)}")

# -------------------- Main Application Class --------------------
class CSVSearchApp(QWidget):
    def __init__(self):
        super().__init__()
        self.logger = logging.getLogger(__name__)
        self.logger.info("Initializing CSVSearchApp")
        self.df = None
        self.last_loaded_path = ""
        self.loader_thread = None
        self.is_loading = False
        self.initUI()
        self.logger.info("CSVSearchApp initialization completed")

    def initUI(self):
        # Define base styles
        self.base_style = """
            QWidget {
                background-color: #f4f6f7;
                color: #2d3436;
            }
            QLabel {
                color: #2d3436;
            }
            QLineEdit, QComboBox {
                background-color: #ffffff;
                color: #2d3436;
                border: 2px solid #dfe6e9;
                border-radius: 8px;
                padding: 8px;
                font-size: 14px;
            }
            QLineEdit:focus, QComboBox:focus {
                border: 2px solid #007AFF;
            }
            QTableView {
                background-color: #ffffff;
                color: #2d3436;
                border: 2px solid #dfe6e9;
                border-radius: 8px;
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #007AFF;
                color: #ffffff;
                padding: 8px;
                border: none;
                border-radius: 8px;
                font-weight: bold;
            }
            QCheckBox {
                color: #2d3436;
                font-size: 14px;
                padding: 5px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 4px;
                border: 2px solid #dfe6e9;
            }
            QCheckBox::indicator:checked {
                background-color: #007AFF;
                border: 2px solid #007AFF;
            }
            QPushButton {
                background-color: #007AFF;
                color: #ffffff;
                border: 2px solid #007AFF;
                border-radius: 8px;
                padding: 10px 20px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #005ecb;
                border: 2px solid #005ecb;
            }
            QPushButton:pressed {
                background-color: #004BA0;
                border: 2px solid #004BA0;
            }
            QScrollArea {
                background-color: #ffffff;
                border: 2px solid #dfe6e9;
                border-radius: 8px;
            }
            QScrollBar:vertical {
                border: none;
                background-color: #f4f6f7;
                width: 10px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background-color: #007AFF;
                border-radius: 5px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #005ecb;
            }
        """

        button_style = """
            QPushButton {
                font-size: 14px;
                font-weight: bold;
                padding: 12px 24px;
                border-radius: 8px;
                border: 2px solid #007AFF;
                background-color: #007AFF;
                color: #ffffff;
            }
            QPushButton:hover {
                background-color: #005ecb;
                border: 2px solid #005ecb;
            }
            QPushButton:pressed {
                background-color: #004BA0;
                border: 2px solid #004BA0;
            }
        """
        self.setWindowTitle("CSV/Excel Search App")
        self.setGeometry(100, 100, 900, 600)
        self.setAcceptDrops(True)

        self.setStyleSheet("""
            background-color: #f4f6f7;
            border-radius: 12px;
        """)

        layout = QVBoxLayout()
        layout.setSpacing(10)

        # File name and status in a horizontal layout
        file_status_layout = QHBoxLayout()
        file_status_layout.setSpacing(15)
        
        # Add load/unload button to the left with adjusted size
        self.load_button = QPushButton("📂 Load", self)
        self.load_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.load_button.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                font-weight: bold;
                padding: 15px 35px;
                border-radius: 8px;
                border: 2px solid #007AFF;
                background-color: #007AFF;
                color: #ffffff;
                min-width: 80px;
                max-width: 120px;
            }
            QPushButton:hover {
                background-color: #005ecb;
                border: 2px solid #005ecb;
            }
            QPushButton:pressed {
                background-color: #004BA0;
                border: 2px solid #004BA0;
            }
        """)
        self.load_button.clicked.connect(self.toggle_load_unload)
        file_status_layout.addWidget(self.load_button)
        
        # File name label with improved styling and word wrap
        self.file_name_label = QLabel("Please load a file to begin", self)
        self.file_name_label.setWordWrap(True)
        self.file_name_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.file_name_label.setStyleSheet("""
            font-size: 16px;
            color: #2d3436;
            padding: 12px 20px;
            font-weight: bold;
        """)
        file_status_layout.addWidget(self.file_name_label)
        
        # Status label with improved styling and word wrap
        self.status_label = QLabel("", self)
        self.status_label.setWordWrap(True)
        self.status_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.status_label.setStyleSheet("""
            font-size: 14px;
            color: #2d3436;
            font-weight: bold;
            padding: 12px 20px;
        """)
        file_status_layout.addWidget(self.status_label)

        layout.addLayout(file_status_layout)

        # Sheet and Column selection in a single line
        controls_layout = QHBoxLayout()
        controls_layout.setSpacing(15)
        
        # Sheet selection
        sheet_container = QWidget()
        sheet_container.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        sheet_layout = QHBoxLayout(sheet_container)
        sheet_layout.setSpacing(5)
        
        sheet_label = QLabel("Sheet:", self)
        sheet_label.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        sheet_label.setStyleSheet("font: bold 14px Arial;")
        sheet_layout.addWidget(sheet_label)
        
        self.sheet_selector = QComboBox(self)
        self.sheet_selector.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.sheet_selector.setStyleSheet("""
            QComboBox {
                background-color: #ffffff;
                font-size: 14px;
                padding: 8px;
                border: 2px solid #dfe6e9;
                border-radius: 8px;
                min-width: 200px;
                selection-color: #007AFF;
            }
            QComboBox:hover {
                border: 2px solid #007AFF;
            }
            QComboBox QAbstractItemView {
                background-color: #ffffff;
                selection-background-color: #007AFF;
                selection-color: #ffffff;
                border: 1px solid #dfe6e9;
                border-radius: 4px;
            }
            QComboBox QAbstractItemView::item {
                padding: 5px;
                min-height: 25px;
            }
        """)
        self.sheet_selector.setEnabled(False)
        self.sheet_selector.currentIndexChanged.connect(self.load_selected_sheet)
        sheet_layout.addWidget(self.sheet_selector)
        controls_layout.addWidget(sheet_container)

        # Column selection
        column_container = QWidget()
        column_container.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        column_layout = QHBoxLayout(column_container)
        column_layout.setSpacing(5)
        
        column_label = QLabel("Select Column:", self)
        column_label.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        column_label.setStyleSheet("font: bold 14px Arial;")
        column_layout.addWidget(column_label)

        self.column_selector = QComboBox(self)
        self.column_selector.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.column_selector.setEnabled(False)  # Initially disabled
        self.column_selector.setStyleSheet("""
            QComboBox {
            background-color: #ffffff;
                font-size: 14px;
            padding: 8px;
                border: 2px solid #dfe6e9;
                border-radius: 8px;
                min-width: 200px;
                selection-color: #007AFF;
            }
            QComboBox:hover {
                border: 2px solid #007AFF;
            }
            QComboBox QAbstractItemView {
                background-color: #ffffff;
                selection-background-color: #007AFF;
                selection-color: #ffffff;
            border: 1px solid #dfe6e9;
                border-radius: 4px;
            }
            QComboBox QAbstractItemView::item {
                padding: 5px;
                min-height: 25px;
            }
        """)
        self.column_selector.currentIndexChanged.connect(self.update_selected_columns)
        column_layout.addWidget(self.column_selector)
        controls_layout.addWidget(column_container)
        
        layout.addLayout(controls_layout)

        # Add progress bar with improved styling
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #dfe6e9;
                border-radius: 8px;
                text-align: center;
                background-color: #f4f6f7;
                height: 20px;
            }
            QProgressBar::chunk {
                background-color: #007AFF;
                border-radius: 6px;
            }
        """)
        self.progress_bar.hide()
        layout.addWidget(self.progress_bar)

        # Search section with improved layout and word wrap
        search_container = QWidget()
        search_container.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        search_container.setStyleSheet("""
            QWidget {
            background-color: #ffffff;
                border: 2px solid #dfe6e9;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        search_layout = QVBoxLayout(search_container)
        search_layout.setSpacing(15)

        # Search inputs row
        search_inputs = QHBoxLayout()
        search_inputs.setSpacing(15)
        
        # First search box
        search_box1_container = QWidget()
        search_box1_container.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        search_box1_layout = QHBoxLayout(search_box1_container)
        search_box1_label = QLabel("First Query:", self)
        search_box1_label.setWordWrap(True)
        search_box1_label.setStyleSheet("font: bold 14px Arial;")
        search_box1_layout.addWidget(search_box1_label)
        
        self.search_box1 = QLineEdit(self)
        self.search_box1.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.search_box1.setPlaceholderText("Enter first query")
        self.search_box1.setStyleSheet("""
            QLineEdit {
            background-color: #ffffff;
            font-size: 14px;
                padding: 10px;
                border: 2px solid #dfe6e9;
                border-radius: 8px;
                min-width: 200px;
            }
            QLineEdit:focus {
                border: 2px solid #007AFF;
            }
        """)
        search_box1_layout.addWidget(self.search_box1)
        search_inputs.addWidget(search_box1_container)

        # Logic selector
        logic_container = QWidget()
        logic_container.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        logic_layout = QHBoxLayout(logic_container)
        logic_label = QLabel("Logic:", self)
        logic_label.setWordWrap(True)
        logic_label.setStyleSheet("font: bold 14px Arial;")
        logic_layout.addWidget(logic_label)

        self.logic_selector = QComboBox(self)
        self.logic_selector.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.logic_selector.addItems(["AND", "OR", "NOT"])
        self.logic_selector.setStyleSheet("""
            QComboBox {
            background-color: #ffffff;
            font-size: 14px;
                padding: 10px;
                border: 2px solid #dfe6e9;
                border-radius: 8px;
                min-width: 100px;
                selection-color: #007AFF;
            }
            QComboBox:hover {
                border: 2px solid #007AFF;
            }
        """)
        logic_layout.addWidget(self.logic_selector)
        search_inputs.addWidget(logic_container)

        # Second search box
        search_box2_container = QWidget()
        search_box2_container.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        search_box2_layout = QHBoxLayout(search_box2_container)
        search_box2_label = QLabel("Second Query:", self)
        search_box2_label.setWordWrap(True)
        search_box2_label.setStyleSheet("font: bold 14px Arial;")
        search_box2_layout.addWidget(search_box2_label)

        self.search_box2 = QLineEdit(self)
        self.search_box2.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.search_box2.setPlaceholderText("Enter second query")
        self.search_box2.setStyleSheet("""
            QLineEdit {
            background-color: #ffffff;
            font-size: 14px;
                padding: 10px;
                border: 2px solid #dfe6e9;
                border-radius: 8px;
                min-width: 200px;
            }
            QLineEdit:focus {
                border: 2px solid #007AFF;
            }
        """)
        search_box2_layout.addWidget(self.search_box2)
        search_inputs.addWidget(search_box2_container)

        search_layout.addLayout(search_inputs)

        # Search options row
        search_options = QHBoxLayout()
        search_options.setSpacing(20)
        
        # Create a container for checkboxes
        checkbox_container = QWidget()
        checkbox_container.setStyleSheet("""
            QWidget {
                background-color: #f8f9fa;
                border: 2px solid #dfe6e9;
                border-radius: 8px;
                padding: 10px;
            }
        """)
        checkbox_layout = QHBoxLayout(checkbox_container)
        checkbox_layout.setSpacing(20)
        
        self.match_case_checkbox = QCheckBox("Match Case", self)
        self.match_case_checkbox.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.match_case_checkbox.setStyleSheet("""
            QCheckBox {
                font-size: 14px;
                color: #2d3436;
                padding: 8px;
                spacing: 12px;
                background-color: transparent;
            }
            QCheckBox::indicator {
                width: 22px;
                height: 22px;
                border-radius: 4px;
                border: 2px solid #dfe6e9;
            }
            QCheckBox::indicator:hover {
                border: 2px solid #007AFF;
            }
            QCheckBox::indicator:checked {
                background-color: #007AFF;
                border: 2px solid #007AFF;
            }
            QCheckBox::indicator:checked:hover {
                background-color: #005ecb;
            }
        """)
        checkbox_layout.addWidget(self.match_case_checkbox)
        
        self.entire_field_checkbox = QCheckBox("Entire Field", self)
        self.entire_field_checkbox.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.entire_field_checkbox.setStyleSheet("""
            QCheckBox {
                font-size: 14px;
                color: #2d3436;
                padding: 8px;
                spacing: 12px;
                background-color: transparent;
            }
            QCheckBox::indicator {
                width: 22px;
                height: 22px;
                border-radius: 4px;
                border: 2px solid #dfe6e9;
            }
            QCheckBox::indicator:hover {
                border: 2px solid #007AFF;
            }
            QCheckBox::indicator:checked {
                background-color: #007AFF;
                border: 2px solid #007AFF;
            }
            QCheckBox::indicator:checked:hover {
                background-color: #005ecb;
            }
        """)
        checkbox_layout.addWidget(self.entire_field_checkbox)
        
        search_options.addWidget(checkbox_container)
        
        # Add search buttons container
        search_buttons_container = QWidget()
        search_buttons_container.setStyleSheet("""
            QWidget {
                background-color: #f8f9fa;
                border: 2px solid #dfe6e9;
                border-radius: 8px;
                padding: 10px;
            }
        """)
        search_buttons_layout = QHBoxLayout(search_buttons_container)
        search_buttons_layout.setSpacing(10)
        
        # Add Find Matches button
        self.search_button = QPushButton("🔍 Find Matches", self)
        self.search_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.search_button.setStyleSheet(button_style)
        self.search_button.clicked.connect(self.search_data)
        search_buttons_layout.addWidget(self.search_button)
        
        # Add Clear Search button
        self.reset_search_button = QPushButton("🔄 Clear Search", self)
        self.reset_search_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.reset_search_button.setStyleSheet(button_style)
        self.reset_search_button.clicked.connect(self.reset_search)
        search_buttons_layout.addWidget(self.reset_search_button)
        
        search_options.addWidget(search_buttons_container)
        search_layout.addLayout(search_options)
        layout.addWidget(search_container)

        # Button Layout with improved organization
        button_container = QWidget()
        button_container.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        button_container.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                border: 2px solid #dfe6e9;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        button_layout = QVBoxLayout(button_container)
        button_layout.setSpacing(15)

        # Row 1: Search operations
        row1 = QHBoxLayout()
        row1.setSpacing(15)
        
        self.highlight_duplicates_button = QPushButton("🔍 Find Duplicates", self)
        self.highlight_duplicates_button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.highlight_duplicates_button.setStyleSheet(button_style)
        self.highlight_duplicates_button.clicked.connect(self.highlight_duplicates)
        row1.addWidget(self.highlight_duplicates_button)
        
        self.refresh_button = QPushButton("🔄 Reload File", self)
        self.refresh_button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.refresh_button.setStyleSheet(button_style)
        self.refresh_button.clicked.connect(self.refresh_data)
        row1.addWidget(self.refresh_button)
        
        button_layout.addLayout(row1)

        # Row 2: Export and visualization
        row2 = QHBoxLayout()
        row2.setSpacing(15)
        
        self.export_button = QPushButton("💾 Save As...", self)
        self.export_button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.export_button.setStyleSheet(button_style)
        self.export_button.clicked.connect(self.export_data)
        row2.addWidget(self.export_button)
        
        self.stats_button = QPushButton("📊 Show Statistics", self)
        self.stats_button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.stats_button.setStyleSheet(button_style)
        self.stats_button.clicked.connect(self.show_statistics)
        row2.addWidget(self.stats_button)
        
        self.visualize_button = QPushButton("📈 Create Chart", self)
        self.visualize_button.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.visualize_button.setStyleSheet(button_style)
        self.visualize_button.clicked.connect(self.show_visualization)
        row2.addWidget(self.visualize_button)
        button_layout.addLayout(row2)
        
        layout.addWidget(button_container)

        # Stats panel with improved styling
        self.stats_panel = QLabel("📊 Stats Panel", self)
        self.stats_panel.setWordWrap(False)  # Disable word wrap
        self.stats_panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.stats_panel.setStyleSheet("""
            font-size: 14px;
            color: #2d3436;
            font-weight: bold;
            padding: 12px 20px;
            background-color: #ffffff;
            border: 2px solid #dfe6e9;
            border-radius: 8px;
        """)
        layout.addWidget(self.stats_panel, alignment=Qt.AlignmentFlag.AlignRight)

        # Table view with improved styling
        self.table = QTableView(self)
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table.setSizeAdjustPolicy(QAbstractScrollArea.SizeAdjustPolicy.AdjustIgnored)
        self.table.setStyleSheet("""
            QTableView {
                font-size: 14px;
                background-color: #ffffff;
                border: 2px solid #dfe6e9;
                border-radius: 8px;
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #007AFF;
                color: #ffffff;
                font-size: 14px;
                font-weight: bold;
                padding: 10px;
                border-radius: 8px;
            }
        """)
        layout.addWidget(self.table)

        # Footer with improved styling and word wrap
        self.footer_label = QLabel("Developed by Jayakumar Sadhasivam - jayakumars.in\niamjayakumars@gmail.com", self)
        self.footer_label.setWordWrap(True)
        self.footer_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.footer_label.setStyleSheet("""
            font-size: 12px;
            color: #2d3436;
            padding: 10px;
            text-align: center;
            font-weight: bold;
        """)
        self.footer_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.footer_label)
 
        self.setLayout(layout)

    # -------------------- File Loading with Improved Performance -------------------- #
    def load_file(self, file_path=None):
        # Check if already loading
        if self.is_loading:
            self.logger.warning("Attempted to load file while another file is being loaded")
            QMessageBox.warning(self, "Warning", "A file is already being loaded. Please wait.")
            return
            
        if not file_path:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Open File", self.last_loaded_path, "All Supported Files (*.csv *.xls *.xlsx *.xlsm *.xlsb)"
            )
            
        if file_path:
            try:
                # Validate file extension
                if not any(file_path.lower().endswith(ext) for ext in ['.csv', '.xls', '.xlsx', '.xlsm', '.xlsb']):
                    self.logger.warning(f"Invalid file extension: {file_path}")
                    QMessageBox.warning(self, "Invalid File", "Please select a valid CSV or Excel file.")
                    return
                
                # Check file size (100MB limit)
                file_size = os.path.getsize(file_path)
                if file_size > 100 * 1024 * 1024:  # 100MB in bytes
                    self.logger.warning(f"File too large: {file_size / (1024*1024):.2f}MB")
                    QMessageBox.warning(self, "File Too Large", "File size exceeds 100MB limit.")
                    return
                
                # Check file permissions
                if not os.access(file_path, os.R_OK):
                    self.logger.error(f"Permission denied for file: {file_path}")
                    QMessageBox.warning(self, "Access Denied", "You don't have permission to read this file.")
                    return
                
                self.logger.info(f"Starting file load process: {file_path}")
                self.is_loading = True
                self.last_loaded_path = file_path
                self.status_label.setText("Loading file... Please wait.")
                self.df = None  # Clear previous data
                
                # Show progress bar
                self.progress_bar.setValue(0)
                self.progress_bar.show()

            # Clear previous UI selections
                self.selected_columns = []  # Reset selected columns
                self.sheet_selector.clear()
                self.sheet_selector.setEnabled(False)
                self.column_selector.clear()
                self.column_selector.setEnabled(False)
                
                # Reset search boxes
                self.search_box1.clear()
                self.search_box2.clear()
                
                # Create and start loader thread
                self.loader_thread = FileLoaderThread(file_path)
                self.loader_thread.data_loaded.connect(self.on_file_loaded)
                self.loader_thread.progress_updated.connect(self.update_progress)
                self.loader_thread.error_occurred.connect(self.handle_load_error)
                self.loader_thread.sheet_names_loaded.connect(self.update_sheet_selector)
                self.loader_thread.start()
                
            except Exception as e:
                self.logger.error(f"Failed to start file loading: {str(e)}", exc_info=True)
                self.is_loading = False
                self.progress_bar.hide()
                QMessageBox.critical(self, "Error", f"Failed to start file loading:\n{str(e)}")
                self.status_label.setText("❌ Failed to start file loading")

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def handle_load_error(self, error_message):
        self.logger.error(f"File loading error: {error_message}")
        self.is_loading = False
        self.progress_bar.hide()
        
        # Determine the type of error and provide specific feedback
        if "No such file or directory" in error_message:
            QMessageBox.warning(self, "File Not Found", 
                              "The selected file could not be found. It may have been moved or deleted.")
            self.status_label.setText("❌ File not found")
            
        elif "Permission denied" in error_message:
            QMessageBox.warning(self, "Access Denied", 
                              "You don't have permission to access this file. Please check file permissions.")
            self.status_label.setText("❌ Access denied")
            
        elif "File is corrupted" in error_message or "Invalid file format" in error_message:
            QMessageBox.warning(self, "Invalid File", 
                              "The file appears to be corrupted or in an invalid format.")
            self.status_label.setText("❌ Invalid file format")
            
        else:
            QMessageBox.warning(self, "Error", 
                              f"Failed to load file:\n{error_message}")
            self.status_label.setText("❌ File loading failed")
            
        # Reset UI state
        self.df = None
        self.last_loaded_path = ""
        self.file_name_label.setText("No file loaded")
        self.selected_columns = []
        self.sheet_selector.clear()
        self.sheet_selector.setEnabled(False)
        self.column_selector.clear()
        self.column_selector.setEnabled(False)
        self.display_data(pd.DataFrame())

    def update_sheet_selector(self, sheet_names):
        """Update the sheet selector with available sheet names"""
        if not sheet_names:
            self.sheet_selector.clear()
            self.sheet_selector.setEnabled(False)
            self.status_label.setText("⚠️ No sheets found in the file")
            return
            
        # Disconnect the signal temporarily to prevent triggering load_selected_sheet
        try:
            self.sheet_selector.currentIndexChanged.disconnect(self.load_selected_sheet)
        except:
            pass  # Signal might not be connected yet
        
        # Update the sheet selector
        self.sheet_selector.clear()
        
        # Validate and clean sheet names
        valid_sheets = []
        for sheet in sheet_names:
            if isinstance(sheet, str) and sheet.strip():
                valid_sheets.append(sheet.strip())
            else:
                self.status_label.setText("⚠️ Some sheet names were invalid and were skipped")
        
        if valid_sheets:
            self.sheet_selector.addItems(valid_sheets)
            self.sheet_selector.setEnabled(True)
            self.sheet_selector.setCurrentIndex(0)
            self.status_label.setText(f"✅ Found {len(valid_sheets)} valid sheets")
            
            # Load the first sheet automatically
            self.load_selected_sheet()
        else:
            self.sheet_selector.setEnabled(False)
            self.status_label.setText("❌ No valid sheets found")
        
        # Reconnect the signal
        self.sheet_selector.currentIndexChanged.connect(self.load_selected_sheet)

    def load_selected_sheet(self):
        if not self.last_loaded_path:
            return
            
        if self.sheet_selector.currentIndex() < 0:
            return
            
            sheet_name = str(self.sheet_selector.currentText()).strip()
        if not sheet_name:
                return
        
        try:
            # Read the selected sheet
            df = pd.read_excel(
                self.last_loaded_path,
                sheet_name=sheet_name,
                engine="openpyxl",
                dtype=str,
                na_filter=False,
                keep_default_na=False
            )
            
            if df.empty:
                self.status_label.setText(f"⚠️ Sheet '{sheet_name}' is empty")
                return
            
            # Update the data and UI
            self.df = df
            self.file_name_label.setText(f"📂 {os.path.basename(self.last_loaded_path)} - {sheet_name}")
            
            # Update column selector
            self.column_selector.clear()
            self.column_selector.addItems(self.df.columns)
            self.column_selector.setCurrentIndex(0)
            
            self.selected_columns = [self.df.columns[0]]
            self.status_label.setText(f"✅ Sheet loaded successfully: {sheet_name}")
            
            # Clear any existing highlights or filters
            if isinstance(self.table.model(), PandasTableModel):
                self.table.model().set_highlight_patterns([])
                self.table.model().set_duplicate_columns([])
            
            # Display the new data
            self.display_data(self.df)
            
            # Update statistics
            self.show_statistics()
            
        except Exception as e:
            self.logger.error(f"Error loading sheet '{sheet_name}': {str(e)}", exc_info=True)
            self.status_label.setText(f"❌ Failed to load sheet '{sheet_name}'")

    def on_file_loaded(self, df, file_path):
        self.progress_bar.hide()
        if df is None or df.empty:
            QMessageBox.warning(self, "Error", "File is empty or failed to load.")
            return
            
        # Update the data and UI
        self.df = df
        self.file_name_label.setText(f"📂 {file_path.split('/')[-1]}")
        
        # Update column selector
        self.column_selector.clear()
        self.column_selector.addItems(self.df.columns)
        self.column_selector.setCurrentIndex(0)  # Select first column by default
        self.column_selector.setEnabled(True)  # Enable the selector
        
        self.selected_columns = [self.df.columns[0]]  # Start with first column selected
        self.status_label.setText(f"✅ File loaded successfully: {file_path}")
        
        # Update button text to show Unload
        self.load_button.setText("⏏ Unload")
        
        self.display_data(self.df)

    def update_selected_columns(self):
        """Update selected columns based on dropdown selection"""
        if self.column_selector.currentText():
            self.selected_columns = [self.column_selector.currentText()]
            self.column_selector.setEnabled(True)

    def refresh_data(self):
        """Refresh the data by reloading the current file"""
        if self.last_loaded_path:
            current_sheet = self.sheet_selector.currentText() if self.sheet_selector.isEnabled() else None
            self.load_file(self.last_loaded_path)
            if current_sheet:
                index = self.sheet_selector.findText(current_sheet)
                if index >= 0:
                    self.sheet_selector.setCurrentIndex(index)
            self.status_label.setText("✅ Data refreshed successfully")
        else:
            self.status_label.setText("⚠️ No file loaded to refresh")

    def toggle_load_unload(self):
        if self.df is not None:
            # If file is loaded, unload it
            self.df = None
            self.last_loaded_path = ""
            self.is_loading = False
            
            # Reset UI elements
            self.file_name_label.setText("Please load a file to begin")
            self.status_label.setText("No file loaded")
            self.stats_panel.setText("📊 Stats Panel")
            
            # Reset selectors
            self.selected_columns = []
            self.sheet_selector.clear()
            self.sheet_selector.setEnabled(False)
            self.column_selector.clear()
            self.column_selector.setEnabled(False)
            
            # Reset search boxes
            self.search_box1.clear()
            self.search_box2.clear()
            
            # Hide progress bar
            self.progress_bar.hide()
            
            # Clear the table view
            self.display_data(pd.DataFrame())
            
            # Update footer
            self.footer_label.setText("Developed by Jayakumar Sadhasivam - jayakumars.in\niamjayakumars@gmail.com")
            
            # Update button text
            self.load_button.setText("📂 Load")
        else:
            # If no file is loaded, load a new file
            self.load_file()

    def reset_search(self):
        """Clear search results and show all data"""
        if self.df is not None:
            # Clear search boxes
            self.search_box1.clear()
            self.search_box2.clear()
            
            # Reset checkboxes
            self.match_case_checkbox.setChecked(False)
            self.entire_field_checkbox.setChecked(False)
            
            # Clear any highlights
            if isinstance(self.table.model(), PandasTableModel):
                self.table.model().set_highlight_patterns([])
                self.table.model().set_duplicate_columns([])
            
            # Show all data
            self.display_data(self.df)
            self.status_label.setText("✅ Search cleared, showing all data")
        else:
            self.status_label.setText("⚠️ No data available")

    def highlight_duplicates(self):
        if self.df is None:
            QMessageBox.warning(self, "Error", "No data available.")
            return
            
        if not self.selected_columns:
            QMessageBox.warning(self, "Error", "Please select at least one column to highlight duplicates.")
            return

        # Find duplicates across selected columns
        duplicates = self.df[self.df.duplicated(subset=self.selected_columns, keep=False)]
        if duplicates.empty:
            QMessageBox.information(self, "Info", "No duplicates found in the selected columns.")
            return

        # Get the column indices for highlighting
        column_indices = [self.df.columns.get_loc(col) for col in self.selected_columns]
        
        # Update display with all data but highlight duplicates
        self.display_data(self.df)
        
        # Set duplicate highlighting
        if isinstance(self.table.model(), PandasTableModel):
            self.table.model().set_duplicate_columns(column_indices)
            
        # Update status with duplicate count
        duplicate_count = len(duplicates)
        unique_count = len(self.df[self.selected_columns].drop_duplicates())
        
        # Set status label with custom styling for duplicate count
        self.status_label.setStyleSheet("""
            font-size: 18px;
            color: #FF3B30;
            font-weight: bold;
            padding: 12px 20px;
            background-color: #ffffff;
            border: 2px solid #FF3B30;
            border-radius: 8px;
        """)
        self.status_label.setText(f"Found {duplicate_count} duplicate entries in selected columns")
        
        # Update statistics
        self.show_statistics()

    def display_data(self, df):
        if df is None or df.empty:
            self.table.setModel(None)
            self.status_label.setText("No data to display.")
            return

        # Update status first with detailed stats
        total_rows = df.shape[0]
        total_columns = df.shape[1]
        stats_text = f"📊 Total Rows: {total_rows} | Total Columns: {total_columns}"
        self.status_label.setText(stats_text)
        
        # Update stats panel with the same information
        self.stats_panel.setText(stats_text)

        # Create and set the model
        model = PandasTableModel(df)
        self.table.setModel(model)
        
        # Enable interactive adjustment and sorting
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setSortingEnabled(True)
        
        # Auto-resize columns
        self.table.resizeColumnsToContents()

    def show_statistics(self):
        if self.df is not None:
            # Get search statistics with optimized vectorization
            query1 = self.search_box1.text().strip()
            query2 = self.search_box2.text().strip()
            
            if query1 or query2:
                search_df = self.df.astype(str)
                if not self.match_case_checkbox.isChecked():
                    search_df = search_df.apply(lambda x: x.str.lower())
                    query1 = query1.lower()
                    query2 = query2.lower()
                
                yellow_highlights = search_df.apply(lambda x: x.str.contains(query1, na=False)).sum().sum()
                green_highlights = search_df.apply(lambda x: x.str.contains(query2, na=False)).sum().sum()
                total_matches = yellow_highlights + green_highlights
            else:
                total_matches = 0

            # Get duplicate statistics if columns are selected
            duplicate_stats = ""
            if self.selected_columns:
                duplicates = self.df[self.df.duplicated(subset=self.selected_columns, keep=False)]
                if not duplicates.empty:
                    duplicate_count = len(duplicates)
                    unique_count = len(self.df[self.selected_columns].drop_duplicates())
                    duplicate_stats = f" | Duplicates: {duplicate_count} | Unique Values: {unique_count}"

            # Update stats panel with all information
            stats_text = f"📊 Total Rows: {self.df.shape[0]} | Total Columns: {self.df.shape[1]}"
            if total_matches > 0:
                stats_text += f" | Search Matches: {total_matches}"
            stats_text += duplicate_stats
            self.stats_panel.setText(stats_text)

    def export_data(self):
        if self.df is None:
            QMessageBox.warning(self, "Error", "No data available to export.")
            return
            
        try:
            self.logger.info("Starting export operation")
            file_name, _ = QFileDialog.getSaveFileName(
                self, 
                "Save File", 
                "", 
                "Excel Files (*.xlsx);;CSV Files (*.csv);;JSON Files (*.json);;PDF Files (*.pdf)"
            )
            
            if not file_name:
                return
                
            # Get selected rows or all rows
            selected_rows = self.table.selectionModel().selectedRows()
            selected_indexes = [index.row() for index in selected_rows] if selected_rows else range(self.df.shape[0])
            export_df = self.df.iloc[selected_indexes].copy()  # Create a copy to avoid modifying original
            
            # Create column selection dialog with improved UI
            dialog = QDialog(self)
            dialog.setWindowTitle("Select Columns to Export")
            dialog.setMinimumWidth(400)
            dialog.setMinimumHeight(500)
            dialog.setStyleSheet("""
                QDialog {
                    background-color: #f4f6f7;
                    border-radius: 10px;
                }
                QLabel {
                    font-size: 14px;
                    color: #2d3436;
                    padding: 5px;
                }
                QPushButton {
                    background-color: #007AFF;
                    color: white;
                    border: none;
                    padding: 8px 15px;
                    border-radius: 5px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #005ecb;
                }
                QPushButton:pressed {
                    background-color: #004BA0;
                }
                QCheckBox {
                    font-size: 13px;
                    padding: 5px;
                    color: #2d3436;
                }
                QCheckBox:hover {
                    background-color: #e9ecef;
                }
                QScrollArea {
                    border: 1px solid #dfe6e9;
                    border-radius: 5px;
                    background-color: white;
                }
            """)
            
            layout = QVBoxLayout()
            layout.setSpacing(10)
            layout.setContentsMargins(15, 15, 15, 15)
            
            # Add header with instructions
            header_label = QLabel("Select columns to export:")
            header_label.setStyleSheet("""
                font-size: 16px;
                font-weight: bold;
                color: #2d3436;
                padding: 10px;
                background-color: #e9ecef;
                border-radius: 5px;
            """)
            layout.addWidget(header_label)
            
            # Create scroll area for checkboxes
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            scroll_widget = QScrollWidget()
            scroll_layout = QVBoxLayout(scroll_widget)
            scroll_layout.setSpacing(5)
            scroll_layout.setContentsMargins(10, 10, 10, 10)
            
            # Add checkboxes for each column
            checkboxes = {}
            for column in export_df.columns:
                cb = QCheckBox(column)
                cb.setChecked(True)  # Default to checked
                checkboxes[column] = cb
                scroll_layout.addWidget(cb)
            
            scroll.setWidget(scroll_widget)
            layout.addWidget(scroll)
            
            # Add button container with improved styling
            button_container = QWidget()
            button_container.setStyleSheet("""
                QWidget {
                    background-color: #e9ecef;
                    border-radius: 5px;
                    padding: 10px;
                }
            """)
            button_layout = QHBoxLayout(button_container)
            button_layout.setSpacing(10)
            
            # Create buttons with icons
            select_all = QPushButton("✓ Select All")
            deselect_all = QPushButton("✗ Deselect All")
            export_button = QPushButton("💾 Export")
            cancel_button = QPushButton("❌ Cancel")
            
            # Add buttons to layout
            button_layout.addWidget(select_all)
            button_layout.addWidget(deselect_all)
            button_layout.addWidget(export_button)
            button_layout.addWidget(cancel_button)
            
            layout.addWidget(button_container)
            dialog.setLayout(layout)
            
            # Connect button signals
            def on_select_all():
                for cb in checkboxes.values():
                    cb.setChecked(True)
                    
            def on_deselect_all():
                for cb in checkboxes.values():
                    cb.setChecked(False)
                    
            def on_export():
                progress_dialog = None  # Initialize progress_dialog variable
                try:
                    # Get selected columns
                    selected_columns = [col for col, cb in checkboxes.items() if cb.isChecked()]
                    if not selected_columns:
                        QMessageBox.warning(dialog, "Warning", "Please select at least one column to export.")
                        return
                        
                    # Get selected rows or all rows
                    selected_rows = self.table.selectionModel().selectedRows()
                    selected_indexes = [index.row() for index in selected_rows] if selected_rows else range(self.df.shape[0])
                    export_df = self.df.iloc[selected_indexes].copy()  # Create a copy to avoid modifying original
                    
                    # Filter DataFrame to selected columns
                    export_df = export_df[selected_columns]
                    
                    # Show progress dialog with improved UI
                    progress_dialog = QDialog(self)
                    progress_dialog.setWindowTitle("Exporting...")
                    progress_dialog.setMinimumWidth(300)
                    progress_dialog.setStyleSheet("""
                        QDialog {
                            background-color: #f4f6f7;
                            border-radius: 10px;
                        }
                        QLabel {
                            font-size: 14px;
                            color: #2d3436;
                            padding: 10px;
                        }
                        QProgressBar {
                            border: 1px solid #dfe6e9;
                            border-radius: 5px;
                            text-align: center;
                            background-color: #f4f6f7;
                        }
                        QProgressBar::chunk {
                            background-color: #007AFF;
                            border-radius: 4px;
                        }
                    """)
                    
                    progress_layout = QVBoxLayout()
                    progress_label = QLabel("Exporting data, please wait...")
                    progress_bar = QProgressBar()
                    progress_bar.setRange(0, 0)  # Indeterminate progress
                    progress_layout.addWidget(progress_label)
                    progress_layout.addWidget(progress_bar)
                    progress_dialog.setLayout(progress_layout)
                    progress_dialog.show()
                    
                    # Process export based on file type
                    if file_name.endswith(".csv"):
                        export_df.to_csv(file_name, index=False, encoding='utf-8')
                        self.status_label.setText("✅ CSV file exported successfully!")
                        
                    elif file_name.endswith(".json"):
                        export_df.to_json(file_name, orient="records", indent=4, force_ascii=False)
                        self.status_label.setText("✅ JSON file exported successfully!")
                        
                    elif file_name.endswith(".pdf"):
                        from reportlab.lib import colors
                        from reportlab.lib.pagesizes import letter, landscape
                        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
                        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
                        
                        # Create PDF document
                        doc = SimpleDocTemplate(
                            file_name,
                            pagesize=landscape(letter),
                            rightMargin=30,
                            leftMargin=30,
                            topMargin=30,
                            bottomMargin=30
                        )
                        
                        # Container for the 'Flowable' objects
                        elements = []
                        
                        # Add title
                        title_style = ParagraphStyle(
                            'CustomTitle',
                            parent=getSampleStyleSheet()['Heading1'],
                            fontSize=16,
                            spaceAfter=30
                        )
                        elements.append(Paragraph("Data Export Report", title_style))
                        
                        # Convert DataFrame to list of lists for table
                        data = [export_df.columns.tolist()]  # Headers
                        data.extend(export_df.values.tolist())  # Data
                        
                        # Create table
                        table = Table(data)
                        
                        # Add style to table
                        style = TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('FONTSIZE', (0, 0), (-1, 0), 12),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                            ('FONTSIZE', (0, 1), (-1, -1), 10),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ])
                        
                        # Add style to table
                        table.setStyle(style)
                        
                        # Add table to elements
                        elements.append(table)
                        
                        # Build PDF
                        doc.build(elements)
                        self.status_label.setText("✅ PDF file exported successfully!")
                        
                    else:  # Excel files
                        with pd.ExcelWriter(file_name, engine="xlsxwriter") as writer:
                            export_df.to_excel(writer, index=False, sheet_name="Filtered Results")
                            worksheet = writer.sheets["Filtered Results"]
                            
                            # Auto-adjust column widths
                        for col_num, value in enumerate(export_df.columns.values):
                                max_length = max(
                                    len(str(value)),
                                    export_df[value].astype(str).str.len().max()
                                )
                                worksheet.set_column(col_num, col_num, max_length + 2)
                            
                            # Add a header format
                                header_format = writer.book.add_format({
                                'bold': True,
                                'bg_color': '#4B5563',
                                'font_color': 'white',
                                'border': 1
                                })
                            
                            # Apply header format
                        for col_num, value in enumerate(export_df.columns.values):
                            worksheet.write(0, col_num, value, header_format)
                        
                        self.status_label.setText("✅ Excel file exported successfully!")
                    
                    if progress_dialog:
                        progress_dialog.close()
                    dialog.accept()
                    
                except Exception as e:
                    if progress_dialog:
                        progress_dialog.close()
                    QMessageBox.critical(dialog, "Export Error", f"Failed to export file:\n{str(e)}")
                    
            def on_cancel():
                dialog.reject()
            
            select_all.clicked.connect(on_select_all)
            deselect_all.clicked.connect(on_deselect_all)
            export_button.clicked.connect(on_export)
            cancel_button.clicked.connect(on_cancel)
            
            # Show dialog
            dialog.exec()
                
            self.logger.info(f"Export completed successfully to {file_name}")
                
        except Exception as e:
            self.logger.error(f"Error in export_data: {str(e)}", exc_info=True)
            QMessageBox.critical(self, "Error", f"Failed to export file:\n{str(e)}")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
        if file_path.endswith((".csv", ".xls", ".xlsx", ".xlsm", ".xlsb")):
            self.load_file(file_path)

    def show_visualization(self):
        try:
            self.logger.info("Starting visualization")
            if self.df is None or self.df.empty:
                QMessageBox.warning(self, "Error", "No data available to visualize.")
                return
            dialog = VisualizationDialog(self.df, self)
            dialog.exec()
            self.logger.info("Visualization completed successfully")
            
        except Exception as e:
            self.logger.error(f"Error in show_visualization: {str(e)}", exc_info=True)
            QMessageBox.critical(self, "Error", f"Failed to create visualization:\n{str(e)}")

    def search_data(self):
        try:
            self.logger.info("Starting search operation")
            query1 = self.search_box1.text().strip()
            query2 = self.search_box2.text().strip()
            logic = self.logic_selector.currentText()
            
            self.logger.debug(f"Search parameters - Query1: {query1}, Query2: {query2}, Logic: {logic}")
            
            if not query1 and not query2:
                self.logger.warning("Search attempted with no queries")
                QMessageBox.warning(self, "Warning", "Please enter at least one search query.")
                return

            # Convert DataFrame to string type for searching
            search_df = self.df.astype(str)

            # Apply case sensitivity if specified
            if not self.match_case_checkbox.isChecked():
                search_df = search_df.apply(lambda x: x.str.lower())
                query1 = query1.lower()
                query2 = query2.lower()

            # Create masks for each query with error handling
            try:
                if self.entire_field_checkbox.isChecked():
                    # For entire field match, use exact match with word boundaries
                    mask1 = search_df.apply(lambda x: x.str.match(f'^{re.escape(query1)}$', na=False))
                    mask2 = search_df.apply(lambda x: x.str.match(f'^{re.escape(query2)}$', na=False))
                else:
                    # For partial match, use contains with escaped pattern
                    mask1 = search_df.apply(lambda x: x.str.contains(re.escape(query1), na=False))
                    mask2 = search_df.apply(lambda x: x.str.contains(re.escape(query2), na=False))
            except Exception as e:
                QMessageBox.warning(self, "Search Error", f"Invalid search pattern: {str(e)}")
                return

            # Combine masks based on logic
            if logic == "AND":
                final_mask = mask1 & mask2
            elif logic == "OR":
                final_mask = mask1 | mask2
            else:  # NOT
                final_mask = mask1 & ~mask2

            # Filter DataFrame with error handling
            try:
                filtered_df = self.df[final_mask.any(axis=1)]
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to filter data: {str(e)}")
                return

            # Update display
            self.display_data(filtered_df)

            # Update highlights in the model
            if isinstance(self.table.model(), PandasTableModel):
                highlight_patterns = []
                if query1:
                    if self.entire_field_checkbox.isChecked():
                        highlight_patterns.append(f'^{re.escape(query1)}$')
                else:
                        highlight_patterns.append(re.escape(query1))
                if query2:
                    if self.entire_field_checkbox.isChecked():
                        highlight_patterns.append(f'^{re.escape(query2)}$')
                    else:
                        highlight_patterns.append(re.escape(query2))
                self.table.model().set_highlight_patterns(highlight_patterns)

            # Update status with detailed information
            total_matches = len(filtered_df)
            total_rows = len(self.df)
            match_percentage = (total_matches / total_rows * 100) if total_rows > 0 else 0
            self.status_label.setText(f"Found {total_matches} matches ({match_percentage:.1f}% of total rows)")
            
            # Update statistics
            self.show_statistics()

            self.logger.info(f"Search completed. Found {len(filtered_df)} matches")
            
        except Exception as e:
            self.logger.error(f"Error in search_data: {str(e)}", exc_info=True)
            QMessageBox.critical(self, "Error", f"Search failed:\n{str(e)}")

    def closeEvent(self, event):
        try:
            self.logger.info("Application closing...")
            
            # Stop any running loader thread
            if hasattr(self, 'loader_thread') and self.loader_thread and self.loader_thread.isRunning():
                self.logger.info("Stopping loader thread")
                self.loader_thread.stop()
                self.loader_thread.wait()
            
            # Clear data
            self.df = None
            
            # Reset UI state
            self.file_name_label.setText("No file loaded")
            self.status_label.setText("Ready")
            self.stats_panel.setText("📊 Stats Panel")
            
            # Clear selectors
            self.sheet_selector.clear()
            self.sheet_selector.setEnabled(False)
            self.column_selector.clear()
            self.column_selector.setEnabled(False)
            
            # Clear search boxes
            self.search_box1.clear()
            self.search_box2.clear()
            
            # Hide progress bar
            self.progress_bar.hide()
            
            # Clear table
            self.table.setModel(None)
            
            self.logger.info("Application closed successfully")
            event.accept()
            
        except Exception as e:
            self.logger.error(f"Error during application closure: {str(e)}", exc_info=True)
            event.accept()

class VisualizationDialog(QDialog):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.logger = logging.getLogger(__name__)
        self.logger.info("Initializing VisualizationDialog")
        self.df = df
        self.initUI()
        self.logger.info("VisualizationDialog initialization completed")

    def initUI(self):
        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(15, 15, 15, 15)

        # Chart type selection
        chart_type_layout = QHBoxLayout()
        chart_type_label = QLabel("Chart Type:")
        self.chart_type = QComboBox()
        self.chart_type.addItems([
            "Bar Chart",
            "Line Chart",
            "Pie Chart",
            "Scatter Plot"
        ])
        chart_type_layout.addWidget(chart_type_label)
        chart_type_layout.addWidget(self.chart_type)
        layout.addLayout(chart_type_layout)

        # X-axis selection
        x_axis_layout = QHBoxLayout()
        x_axis_label = QLabel("X-Axis:")
        self.x_axis = QComboBox()
        self.x_axis.addItems(self.df.columns)
        x_axis_layout.addWidget(x_axis_label)
        x_axis_layout.addWidget(self.x_axis)
        layout.addLayout(x_axis_layout)

        # Y-axis selection
        y_axis_layout = QHBoxLayout()
        y_axis_label = QLabel("Y-Axis:")
        self.y_axis = QComboBox()
        self.y_axis.addItems(self.df.columns)
        y_axis_layout.addWidget(y_axis_label)
        y_axis_layout.addWidget(self.y_axis)
        layout.addLayout(y_axis_layout)

        # Create matplotlib figure
        self.figure = Figure(figsize=(8, 6))
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)

        # Button container
        button_container = QWidget()
        button_container.setStyleSheet("""
            QWidget {
                background-color: #e9ecef;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        button_layout = QHBoxLayout(button_container)
        button_layout.setSpacing(10)

        # Create buttons
        plot_button = QPushButton("📊 Plot")
        save_button = QPushButton("💾 Save")
        close_button = QPushButton("❌ Close")

        # Add buttons to layout
        button_layout.addWidget(plot_button)
        button_layout.addWidget(save_button)
        button_layout.addWidget(close_button)

        layout.addWidget(button_container)
        self.setLayout(layout)

        # Connect signals
        plot_button.clicked.connect(self.plot_chart)
        save_button.clicked.connect(self.save_chart)
        close_button.clicked.connect(self.close)
        self.chart_type.currentTextChanged.connect(self.update_axis_options)

    def update_axis_options(self, chart_type):
        # Update available options based on chart type
        if chart_type in ["Bar Chart", "Line Chart", "Scatter Plot"]:
            self.x_axis.setEnabled(True)
            self.y_axis.setEnabled(True)
        elif chart_type == "Pie Chart":
            self.x_axis.setEnabled(True)
            self.y_axis.setEnabled(False)

    def plot_chart(self):
        try:
            self.logger.info("Starting chart plotting")
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            
            chart_type = self.chart_type.currentText()
            x_col = self.x_axis.currentText()
            y_col = self.y_axis.currentText()
            
            # Get the selected text from the parent window
            selected_text = self.parent().search_box1.text().strip()
            
            if chart_type == "Bar Chart":
                plot = sns.barplot(data=self.df, x=x_col, y=y_col, ax=ax)
                # Highlight bars containing selected text
                if selected_text:
                    for i, label in enumerate(plot.get_xticklabels()):
                        if selected_text.lower() in label.get_text().lower():
                            plot.patches[i].set_facecolor('blue')
                            plot.patches[i].set_alpha(0.5)
                
            elif chart_type == "Line Chart":
                plot = sns.lineplot(data=self.df, x=x_col, y=y_col, ax=ax)
                # Highlight points containing selected text
                if selected_text:
                    for i, label in enumerate(plot.get_xticklabels()):
                        if selected_text.lower() in label.get_text().lower():
                            plot.lines[0].get_xdata()[i] = i
                            plot.lines[0].get_ydata()[i] = plot.lines[0].get_ydata()[i]
                            ax.plot(i, plot.lines[0].get_ydata()[i], 'bo', markersize=10)
                
            elif chart_type == "Scatter Plot":
                plot = sns.scatterplot(data=self.df, x=x_col, y=y_col, ax=ax)
                # Highlight points containing selected text
                if selected_text:
                    for i, label in enumerate(plot.get_xticklabels()):
                        if selected_text.lower() in label.get_text().lower():
                            ax.scatter(i, plot.get_ydata()[i], color='blue', s=100)
                
            elif chart_type == "Pie Chart":
                # Create pie chart with highlighted segments
                data = self.df[x_col].value_counts()
                colors = ['#ff9999', '#66b2ff', '#99ff99', '#ffcc99', '#ff99cc', '#99ccff']
                if selected_text:
                    # Highlight segments containing selected text
                    for i, label in enumerate(data.index):
                        if selected_text.lower() in str(label).lower():
                            colors[i] = 'blue'
                
                wedges, texts, autotexts = ax.pie(data, labels=data.index, autopct='%1.1f%%',
                                                colors=colors, startangle=90)
                
                # Make highlighted text blue
                for i, text in enumerate(texts):
                    if selected_text and selected_text.lower() in text.get_text().lower():
                        text.set_color('blue')
                        text.set_fontweight('bold')
                        if i < len(autotexts):
                            autotexts[i].set_color('blue')
                            autotexts[i].set_fontweight('bold')

            # Customize the plot
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            
            # Add title with selected text if any
            if selected_text:
                ax.set_title(f'Chart with highlighted matches for "{selected_text}"', 
                           pad=20, color='blue')
            
            self.canvas.draw()
            self.logger.info("Chart plotted successfully")
            
        except Exception as e:
            self.logger.error(f"Error in plot_chart: {str(e)}", exc_info=True)
            QMessageBox.critical(self, "Error", f"Failed to create chart:\n{str(e)}")

    def save_chart(self):
        try:
            file_name, _ = QFileDialog.getSaveFileName(
                self, "Save Chart", "", "PNG Files (*.png);;PDF Files (*.pdf);;SVG Files (*.svg)"
            )
            if file_name:
                self.figure.savefig(file_name, bbox_inches='tight', dpi=300)
                QMessageBox.information(self, "Success", "Chart saved successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save chart:\n{str(e)}")

if __name__ == "__main__":
    try:
        logger.info("Starting application")
        app = QApplication(sys.argv)
        window = CSVSearchApp()
        window.show()
        logger.info("Application window displayed")
        sys.exit(app.exec())
    except Exception as e:
        logger.critical(f"Application failed to start: {str(e)}", exc_info=True)
        sys.exit(1)