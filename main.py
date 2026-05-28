# File: main.py
import sys
import os
# Dynamic sys.path expansion to allow embedded python to locate adjacent scripts
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import time
import json
import base64
import re
import ast
import traceback
import tempfile
from typing import Optional, List, Dict, Any

# PyQt6 Imports
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QLineEdit, QTextEdit, QPushButton, QComboBox,
    QListWidget, QListWidgetItem, QGroupBox, QFileDialog, QMessageBox,
    QProgressBar, QDialog, QCheckBox, QFrame, QSplitter, QSpinBox,
    QGraphicsDropShadowEffect, QTableWidget, QTableWidgetItem, QHeaderView,
    QAbstractItemView
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt6.QtGui import QFont, QColor, QIcon, QLinearGradient, QBrush, QPalette

import pandas as pd
import openpyxl
import requests
import fitz  # PyMuPDF

# Local Imports
import settings_manager

# Create a clean folder for temp files inside the project
TEMP_DIR = os.path.join(os.path.dirname(__file__), "temps")
os.makedirs(TEMP_DIR, exist_ok=True)


# --- Helper: Robust JSON Parsing & Extraction ---
def robust_parse_json(text: str) -> Any:
    """
    Cleans smart curly quotes and parses JSON/Python dict literal strings.
    Strips markdown code blocks and extracts the outermost {} or [].
    """
    if not text:
        return {}
    
    clean_text = text.strip()
    if not clean_text:
        return {}

    # Strip markdown fences
    for fence in ["```json", "```JSON", "```"]:
        if fence in clean_text:
            try:
                parts = clean_text.split(fence)
                if len(parts) > 1:
                    clean_text = parts[1].split("```")[0].strip()
            except Exception:
                pass

    # Extract outermost {} or []
    start_idx = -1
    for i, char in enumerate(clean_text):
        if char in ('{', '['):
            start_idx = i
            break
            
    end_idx = -1
    for i in range(len(clean_text) - 1, -1, -1):
        if clean_text[i] in ('}', ']'):
            end_idx = i
            break
            
    if start_idx != -1 and end_idx != -1 and end_idx > start_idx:
        clean_text = clean_text[start_idx : end_idx + 1]

    # Replace smart/curly quotes
    clean_text = (clean_text
                  .replace('“', '"')
                  .replace('”', '"')
                  .replace('‘', "'")
                  .replace('’', "'"))
                  
    try:
        return json.loads(clean_text)
    except Exception:
        # Try evaluating as python literal
        try:
            return ast.literal_eval(clean_text)
        except Exception:
            # Normalize common JSON-vs-Python terms
            try:
                normalized = re.sub(r'\bnull\b', 'None', clean_text)
                normalized = re.sub(r'\btrue\b', 'True', normalized)
                normalized = re.sub(r'\bfalse\b', 'False', normalized)
                return ast.literal_eval(normalized)
            except Exception as final_e:
                raise ValueError(f"JSON Parse Error: {final_e}. Raw: {text[:300]}")


# --- Helper: Merge lists of page dictionaries ---
def combine_page_jsons(page_dicts: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Combines page-by-page JSON results into a single unified JSON structure.
    Concatenates 'items' arrays and merges 'header' and 'other infor' keys.
    """
    if not page_dicts:
        return {}
        
    combined = {
        "header": {},
        "items": [],
        "other infor": {}
    }
    
    for pd in page_dicts:
        if not isinstance(pd, dict):
            continue
            
        # 1. Merge Header Fields
        header = pd.get("header")
        if isinstance(header, dict):
            for k, v in header.items():
                # Take first non-empty value
                if v is not None and str(v).strip() != "" and k not in combined["header"]:
                    combined["header"][k] = v
                    
        # 2. Concatenate Item Fields
        items = pd.get("items")
        if isinstance(items, list):
            for item in items:
                if isinstance(item, dict):
                    combined["items"].append(item)
                    
        # 3. Merge Other Info Fields
        other = pd.get("other infor") or pd.get("other_info") or pd.get("summary")
        if isinstance(other, dict):
            for k, v in other.items():
                if v is not None and str(v).strip() != "":
                    # Overwrite or append based on numeric vs text
                    combined["other infor"][k] = v
                    
    return combined


# --- Helper: Discover JSON Keys ---
def discover_json_keys(data: Any, prefix: str = "") -> List[str]:
    """Recursively lists all unique dotted paths representing fields in the JSON."""
    paths = []
    if isinstance(data, dict):
        for k, v in data.items():
            curr_path = f"{prefix}.{k}" if prefix else k
            if isinstance(v, (dict, list)):
                paths.extend(discover_json_keys(v, curr_path))
            else:
                paths.append(curr_path)
    elif isinstance(data, list):
        if data:
            # Analyze first item in list for key structure
            paths.extend(discover_json_keys(data[0], prefix))
    else:
        if prefix:
            paths.append(prefix)
    return sorted(list(set(paths)))


# --- UI: Schema Column Selector Dialog ---
class ColumnSelectorDialog(QDialog):
    def __init__(self, raw_json: dict, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Schema Column Discovery & Mapper")
        self.setMinimumSize(650, 500)
        self.raw_json = raw_json
        
        self.header_fields = []
        self.item_fields = []
        self.other_fields = []
        
        self.selected_columns = []
        
        self._parse_schema_groups()
        self._init_ui()
        
    def _parse_schema_groups(self):
        """Categorizes raw JSON keys into Header, Items, and Other Info fields."""
        # 1. Header category
        header = self.raw_json.get("header", {})
        if isinstance(header, dict):
            for k in header.keys():
                self.header_fields.append(f"header.{k}")
        else:
            # Fallback
            for k in self.raw_json.keys():
                if k not in ["items", "other infor", "other_info", "summary"]:
                    self.header_fields.append(k)

        # 2. Items category
        items = self.raw_json.get("items", [])
        if isinstance(items, list) and len(items) > 0:
            first_item = items[0]
            if isinstance(first_item, dict):
                for k in first_item.keys():
                    self.item_fields.append(f"items.{k}")
                    
        # 3. Other category
        other = self.raw_json.get("other infor") or self.raw_json.get("other_info") or self.raw_json.get("summary", {})
        if isinstance(other, dict):
            for k in other.keys():
                self.other_fields.append(f"other infor.{k}")

    def _init_ui(self):
        layout = QVBoxLayout(self)
        
        title = QLabel("Select Discovered Schema Fields to Add to Excel:")
        title.setFont(QFont("Inter", 11, QFont.Weight.Bold))
        title.setStyleSheet("color: #F3F4F6; margin-bottom: 8px;")
        layout.addWidget(title)
        
        # Grid containing three checkable lists
        grid_widget = QWidget()
        grid_layout = QHBoxLayout(grid_widget)
        grid_layout.setContentsMargins(0, 0, 0, 0)
        grid_layout.setSpacing(12)
        
        # 1. Header List Card
        h_box = QGroupBox("Header Fields")
        h_box.setStyleSheet("QGroupBox { font-weight: bold; color: #8B5CF6; border: 1px solid #333333; border-radius: 8px; margin-top: 12px; } QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px; }")
        h_layout = QVBoxLayout(h_box)
        self.h_list = QListWidget()
        h_layout.addWidget(self.h_list)
        grid_layout.addWidget(h_box)
        
        # 2. Items List Card
        i_box = QGroupBox("Item Fields (Tabular)")
        i_box.setStyleSheet("QGroupBox { font-weight: bold; color: #10B981; border: 1px solid #333333; border-radius: 8px; margin-top: 12px; } QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px; }")
        i_layout = QVBoxLayout(i_box)
        self.i_list = QListWidget()
        i_layout.addWidget(self.i_list)
        grid_layout.addWidget(i_box)
        
        # 3. Other Info Card
        o_box = QGroupBox("Summary / Other Fields")
        o_box.setStyleSheet("QGroupBox { font-weight: bold; color: #F59E0B; border: 1px solid #333333; border-radius: 8px; margin-top: 12px; } QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px; }")
        o_layout = QVBoxLayout(o_box)
        self.o_list = QListWidget()
        o_layout.addWidget(self.o_list)
        grid_layout.addWidget(o_box)
        
        layout.addWidget(grid_widget)
        
        # Populate widgets
        for f in self.header_fields:
            item = QListWidgetItem(f.replace("header.", ""))
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked) # Checked by default
            self.h_list.addItem(item)
            
        for f in self.item_fields:
            item = QListWidgetItem(f.replace("items.", ""))
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            self.i_list.addItem(item)
            
        for f in self.other_fields:
            item = QListWidgetItem(f.replace("other infor.", ""))
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            self.o_list.addItem(item)
            
        # Select Buttons
        btn_layout = QHBoxLayout()
        select_all = QPushButton("Select All Fields")
        select_none = QPushButton("Clear Selection")
        btn_layout.addWidget(select_all)
        btn_layout.addWidget(select_none)
        layout.addLayout(btn_layout)
        
        # OK / Cancel Buttons
        self.btn_confirm = QPushButton("Generate Reference Prompt & Apply Schema")
        self.btn_confirm.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #10B981, stop:1 #059669); color: white; padding: 10px; font-weight: bold;")
        layout.addWidget(self.btn_confirm)
        
        # Connections
        select_all.clicked.connect(self._select_all)
        select_none.clicked.connect(self._clear_all)
        self.btn_confirm.clicked.connect(self.accept)
        
    def _select_all(self):
        for lst in [self.h_list, self.i_list, self.o_list]:
            for i in range(lst.count()):
                lst.item(i).setCheckState(Qt.CheckState.Checked)
                
    def _clear_all(self):
        for lst in [self.h_list, self.i_list, self.o_list]:
            for i in range(lst.count()):
                lst.item(i).setCheckState(Qt.CheckState.Unchecked)

    def get_selected_schema(self) -> Dict[str, List[str]]:
        """Returns the dictionary mapping categories to user-selected fields."""
        selected = {
            "header": [],
            "items": [],
            "other infor": []
        }
        
        for i in range(self.h_list.count()):
            item = self.h_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected["header"].append(item.text())
                
        for i in range(self.i_list.count()):
            item = self.i_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected["items"].append(item.text())
                
        for i in range(self.o_list.count()):
            item = self.o_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                selected["other infor"].append(item.text())
                
        return selected


# --- UI: PDF Viewer Dialog ---
class PDFViewerDialog(QDialog):
    def __init__(self, filepath: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle(f"PDF Viewer - {os.path.basename(filepath)}")
        self.setMinimumSize(800, 600)
        self.filepath = filepath
        self.doc = None
        
        self._init_ui()
        
    def _init_ui(self):
        layout = QVBoxLayout(self)
        
        # Load PDF using fitz
        try:
            self.doc = fitz.open(self.filepath)
            num_pages = len(self.doc)
        except Exception as e:
            layout.addWidget(QLabel(f"Failed to open PDF file: {e}"))
            return

        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("Select Page:"))
        
        self.page_selector = QComboBox()
        for i in range(num_pages):
            self.page_selector.addItem(f"Page {i + 1} of {num_pages}")
        top_layout.addWidget(self.page_selector)
        top_layout.addStretch()
        layout.addLayout(top_layout)
        
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setStyleSheet("background-color: #181818; color: #F3F4F6; font-family: monospace; font-size: 12px;")
        layout.addWidget(self.text_edit)
        
        info_label = QLabel(f"Total Pages: {num_pages} | Path: {self.filepath}")
        info_label.setStyleSheet("color: #9CA3AF; font-style: italic;")
        layout.addWidget(info_label)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        close_btn.setStyleSheet("background-color: #2D2D30; color: #F3F4F6; padding: 6px 12px;")
        layout.addWidget(close_btn)
        
        self.page_selector.currentIndexChanged.connect(self.load_page_text)
        
        if num_pages > 0:
            self.load_page_text(0)
            
    def load_page_text(self, index: int):
        if not self.doc:
            return
        try:
            page = self.doc[index]
            text = page.get_text()
            if not text.strip():
                text = "(No text content could be extracted from this page. Note: This could be a scanned document or image PDF.)"
            self.text_edit.setPlainText(text)
        except Exception as e:
            self.text_edit.setPlainText(f"Error loading page text: {e}")
            
    def closeEvent(self, event):
        if self.doc:
            try:
                self.doc.close()
            except Exception:
                pass
        super().closeEvent(event)


# --- UI: Excel Viewer Dialog ---
class ExcelViewerDialog(QDialog):
    def __init__(self, filepath: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle(f"Excel Viewer - {os.path.basename(filepath)}")
        self.setMinimumSize(950, 650)
        self.filepath = filepath
        
        self._init_ui()
        
    def _init_ui(self):
        layout = QVBoxLayout(self)
        
        # Load the excel sheets
        try:
            xls = pd.ExcelFile(self.filepath)
            sheet_names = xls.sheet_names
        except Exception as e:
            layout.addWidget(QLabel(f"Failed to load Excel file: {e}"))
            return

        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("Select Sheet:"))
        
        self.sheet_selector = QComboBox()
        self.sheet_selector.addItems(sheet_names)
        top_layout.addWidget(self.sheet_selector)
        top_layout.addStretch()
        layout.addLayout(top_layout)
        
        self.table = QTableWidget()
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: #181818;
                border: 1px solid #333333;
                color: #F3F4F6;
                gridline-color: #2D2D2D;
            }
            QHeaderView::section {
                background-color: #1E1E1E;
                color: #A855F7;
                font-weight: bold;
                padding: 6px;
                border: 1px solid #333333;
            }
        """)
        self.table.setShowGrid(True)
        layout.addWidget(self.table)
        
        self.info_label = QLabel(f"Path: {self.filepath}")
        self.info_label.setStyleSheet("color: #9CA3AF; font-style: italic;")
        layout.addWidget(self.info_label)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        close_btn.setStyleSheet("background-color: #2D2D30; color: #F3F4F6; padding: 6px 12px;")
        layout.addWidget(close_btn)
        
        self.sheet_selector.currentTextChanged.connect(self.load_sheet_data)
        
        if sheet_names:
            self.load_sheet_data(sheet_names[0])
            
    def load_sheet_data(self, sheet_name: str):
        try:
            df = pd.read_excel(self.filepath, sheet_name=sheet_name)
            self.table.clear()
            self.table.setRowCount(df.shape[0])
            self.table.setColumnCount(df.shape[1])
            self.table.setHorizontalHeaderLabels([str(c) for c in df.columns])
            
            for r_idx in range(df.shape[0]):
                for c_idx in range(df.shape[1]):
                    val = df.iloc[r_idx, c_idx]
                    val_str = "" if pd.isna(val) else str(val)
                    item = QTableWidgetItem(val_str)
                    self.table.setItem(r_idx, c_idx, item)
            
            self.table.resizeColumnsToContents()
            self.info_label.setText(f"Rows: {df.shape[0]} | Columns: {df.shape[1]} | Path: {self.filepath}")
        except Exception as e:
            self.info_label.setText(f"Error loading sheet '{sheet_name}': {e}")


# --- Thread Worker for Async AI Batch Operations ---
class ConversionWorker(QThread):
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, str)  # percentage, text
    file_finished_signal = pyqtSignal(str, str)  # input_path, output_path
    batch_finished_signal = pyqtSignal(bool, str)  # success, summary
    schema_discovered_signal = pyqtSignal(dict)  # raw json discovered

    def __init__(self, 
                 files: List[str],
                 ai_engine: str,
                 api_key: str,
                 model_name: str,
                 prompt_text: str,
                 output_dir: str,
                 delay: int,
                 is_schema_discovery: bool = False,
                 selected_schema: Optional[Dict[str, List[str]]] = None,
                 fail_attempts: int = 3,
                 fail_delay: int = 5,
                 discover_all_pages: bool = False):
        super().__init__()
        self.files = files
        self.ai_engine = ai_engine
        self.api_key = api_key
        self.model_name = model_name
        self.prompt_text = prompt_text
        self.output_dir = output_dir
        self.delay = delay
        self.is_schema_discovery = is_schema_discovery
        self.selected_schema = selected_schema
        self.fail_attempts = fail_attempts
        self.fail_delay = fail_delay
        self.discover_all_pages = discover_all_pages
        
        self.is_running = True

    def run(self):
        try:
            if self.is_schema_discovery:
                self.run_schema_discovery()
            else:
                self.run_batch_conversion()
        except Exception as e:
            self.log_signal.emit(f"Critical Worker Exception: {e}")
            self.batch_finished_signal.emit(False, str(e))

    def run_schema_discovery(self):
        """Processes 1 selected file and returns its raw JSON structure for schema creation."""
        if not self.files:
            self.batch_finished_signal.emit(False, "No file selected for schema discovery.")
            return

        file_path = self.files[0]
        self.log_signal.emit(f"🚀 COMMENCING SCHEMA DISCOVERY ON: {os.path.basename(file_path)}")
        self.progress_signal.emit(10, "Extracting page schema...")

        temp_split_path = None
        actual_pdf_path = file_path
        if not self.discover_all_pages:
            try:
                doc = fitz.open(file_path)
                if len(doc) > 1:
                    new_doc = fitz.open()
                    new_doc.insert_pdf(doc, from_page=0, to_page=0)
                    temp_split_path = os.path.join(TEMP_DIR, f"discover_page1_{os.path.basename(file_path)}")
                    new_doc.save(temp_split_path)
                    new_doc.close()
                    actual_pdf_path = temp_split_path
                    self.log_signal.emit("📄 Multi-page PDF detected; running Schema Discovery on the first page only.")
                doc.close()
            except Exception as e:
                self.log_signal.emit(f"⚠️ Failed to extract first page, falling back to all pages. Error: {e}")

        try:
            # Execute conversion on single file
            json_response = self._process_single_pdf_to_json(actual_pdf_path)
            if isinstance(json_response, list):
                if len(json_response) == 1 and isinstance(json_response[0], dict) and ("items" in json_response[0] or "header" in json_response[0]):
                    json_response = json_response[0]
                else:
                    json_response = {
                        "header": {},
                        "items": json_response,
                        "other infor": {}
                    }
            elif not isinstance(json_response, dict):
                json_response = {}
            self.schema_discovered_signal.emit(json_response)
            self.batch_finished_signal.emit(True, "Schema discovery completed successfully.")
        except Exception as e:
            self.log_signal.emit(f"❌ Schema Discovery Failed: {e}")
            self.batch_finished_signal.emit(False, f"Schema Discovery Failed: {str(e)}")
        finally:
            if temp_split_path and os.path.exists(temp_split_path):
                try:
                    os.remove(temp_split_path)
                    self.log_signal.emit("🧹 Cleaned up temporary first-page PDF.")
                except Exception as cleanup_err:
                    self.log_signal.emit(f"⚠️ Failed to delete temporary file {temp_split_path}: {cleanup_err}")

    def run_batch_conversion(self):
        """Executes full batch translation of files."""
        total_files = len(self.files)
        if total_files == 0:
            self.batch_finished_signal.emit(False, "No PDF files selected to convert.")
            return

        self.log_signal.emit(f"🚀 STARTING BATCH CONVERSION: {total_files} files with {self.ai_engine}")
        
        success_count = 0
        failed_files = []

        for idx, pdf_path in enumerate(self.files):
            if not self.is_running:
                self.log_signal.emit("⏹️ Batch execution aborted by user.")
                break

            filename = os.path.basename(pdf_path)
            
            # Emit progress immediately at start of step
            percent = int((idx / total_files) * 100)
            self.progress_signal.emit(percent, f"Checking/Processing {filename}...")

            # Check if this PDF file has already been converted (Excel exists in output folder)
            xlsx_filename = os.path.splitext(filename)[0] + ".xlsx"
            xlsx_path = os.path.join(self.output_dir, xlsx_filename)
            if os.path.exists(xlsx_path):
                self.log_signal.emit(f"⏭️ Skipping (Excel already exists): {xlsx_filename}")
                self.file_finished_signal.emit(pdf_path, xlsx_path)
                success_count += 1
                continue

            self.log_signal.emit(f"\n--- Processing File {idx + 1}/{total_files}: {filename} ---")


            # Apply Inter-File Delay (except first file)
            if idx > 0 and self.delay > 0:
                self.log_signal.emit(f"⏳ Sleeping for {self.delay}s to safeguard API limits...")
                for i in range(self.delay):
                    if not self.is_running:
                        break
                    time.sleep(1)
                if not self.is_running:
                    break

            try:
                # 1. Get structured JSON from AI
                extracted_json = self._process_single_pdf_to_json(pdf_path)

                # 2. Normalize and construct Pandas DataFrame
                df = self._convert_json_to_df(extracted_json)

                # 3. Export to Output XLSX Folder
                xlsx_filename = os.path.splitext(filename)[0] + ".xlsx"
                xlsx_path = os.path.join(self.output_dir, xlsx_filename)
                
                # Ensure output dir exists
                os.makedirs(self.output_dir, exist_ok=True)
                
                # Write to Excel
                df.to_excel(xlsx_path, index=False)
                self.log_signal.emit(f"✅ Excel sheet created successfully at: {xlsx_path}")
                
                self.file_finished_signal.emit(pdf_path, xlsx_path)
                success_count += 1
            except Exception as e:
                self.log_signal.emit(f"❌ Failed to process '{filename}': {e}")
                traceback.print_exc()
                failed_files.append(filename)

        self.progress_signal.emit(100, "Finished batch!")
        summary_msg = f"Completed {success_count}/{total_files} conversions."
        if failed_files:
            summary_msg += f"\nFailed files: {', '.join(failed_files)}"
            
        self.batch_finished_signal.emit(success_count > 0, summary_msg)

    # --- Router: Single PDF to JSON Conversion ---
    def _process_single_pdf_to_json(self, pdf_path: str) -> dict:
        """Helper to invoke correct AI engine with Rate-Limit/Fail Retry Mechanism."""
        retries = self.fail_attempts
        backoff = self.fail_delay

        for attempt in range(1, retries + 1):
            try:
                if self.ai_engine == "Gemini":
                    return self._call_gemini_api(pdf_path)
                elif self.ai_engine == "GPT (OpenAI)":
                    return self._call_gpt_api(pdf_path)
                else:
                    return self._call_local_ai_api(pdf_path)
            except Exception as e:
                if attempt < retries:
                    self.log_signal.emit(f"⚠️ [Attempt {attempt}/{retries}] AI model call failed. Retrying in {backoff}s... Error: {e}")
                    time.sleep(backoff)
                else:
                    raise e

    # --- Engine: Gemini Multimodal SDK API ---
    def _call_gemini_api(self, pdf_path: str) -> dict:
        if not self.api_key:
            raise ValueError("Gemini API key is required. Input directly in the Middle panel.")

        self.log_signal.emit(f"Connecting to Gemini API using model: {self.model_name}...")
        
        # Lazy import of google genai to avoid start lags
        try:
            import google.genai
            from google.genai import types
        except ImportError:
            raise RuntimeError("The 'google-genai' library is missing from your python environment. Install it via requirements.txt.")

        # Initialize Client
        client = google.genai.Client(api_key=self.api_key)
        uploaded_file = None
        
        try:
            self.log_signal.emit(f"Uploading PDF to Gemini Files Storage: {os.path.basename(pdf_path)}...")
            with open(pdf_path, "rb") as f:
                uploaded_file = client.files.upload(
                    file=f,
                    config=types.UploadFileConfig(
                        display_name=os.path.basename(pdf_path),
                        mime_type="application/pdf"
                    )
                )
            self.log_signal.emit(f"Upload complete. File URI: {uploaded_file.uri}")

            # Prepare content payload
            contents = [uploaded_file, self.prompt_text]

            # Request generation
            self.log_signal.emit("Sending content generation request to Gemini...")
            response = client.models.generate_content(
                model=self.model_name,
                contents=contents
            )
            
            raw_text = response.text
            self.log_signal.emit(f"Response received. Content length: {len(raw_text)}")
            
            # Robustly parse structured JSON
            parsed = robust_parse_json(raw_text)
            return parsed
        finally:
            # Always clean up files in storage
            if uploaded_file:
                try:
                    client.files.delete(name=uploaded_file.name)
                    self.log_signal.emit("Cleaned up Gemini storage upload.")
                except Exception as ce:
                    self.log_signal.emit(f"Warning: Failed to clean up Gemini file: {ce}")

    # --- Engine: GPT Vision API via Base64 Pages ---
    def _call_gpt_api(self, pdf_path: str) -> dict:
        if not self.api_key:
            raise ValueError("GPT API Key is required. Input directly in the Middle panel.")

        self.log_signal.emit("Converting PDF pages to images in-memory via PyMuPDF...")
        
        # Convert all PDF pages to Base64 in-memory images
        images_base64 = []
        doc = fitz.open(pdf_path)
        for page_idx, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
            img_bytes = pix.tobytes("png")
            b64_str = base64.b64encode(img_bytes).decode('utf-8')
            images_base64.append(b64_str)
        doc.close()
        
        self.log_signal.emit(f"Successfully converted {len(images_base64)} pages to Base64 images.")

        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }

        # Prepare multimodal request message body
        content_payload = [{"type": "text", "text": self.prompt_text}]
        for b64 in images_base64:
            content_payload.append({
                "type": "image_url",
                "image_url": {
                    "url": f"data:image/png;base64,{b64}"
                }
            })

        payload = {
            "model": self.model_name,
            "messages": [
                {
                    "role": "user",
                    "content": content_payload
                }
            ],
            "response_format": {"type": "json_object"}, # Enforce JSON schema
            "temperature": 0.1
        }

        self.log_signal.emit(f"Calling OpenAI Chat Completions Vision API with {self.model_name}...")
        response = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers=headers,
            json=payload,
            timeout=300
        )
        
        if response.status_code != 200:
            raise ValueError(f"OpenAI API Error (Status {response.status_code}): {response.text}")

        resp_json = response.json()
        raw_text = resp_json["choices"][0]["message"]["content"]
        
        self.log_signal.emit("Received response from OpenAI GPT. Parsing JSON...")
        parsed = robust_parse_json(raw_text)
        return parsed

    # --- Engine: Local AI Split Page-by-Page API ---
    def _call_local_ai_api(self, pdf_path: str) -> dict:
        self.log_signal.emit(f"Checking Local AI PDF path: {pdf_path}")
        
        base_url = self.api_key.strip() if self.api_key else "http://api-localai.germantest.net"
        base_url = base_url.rstrip("/")
        endpoint = "/invoice-custom"
        full_url = f"{base_url}{endpoint}"
        
        # Local AI Page-by-Page Execution
        self.log_signal.emit("Splitting multi-page PDF into single-page documents for client-side Local AI execution...")
        
        split_paths = []
        doc = fitz.open(pdf_path)
        num_pages = len(doc)
        
        for page_num in range(num_pages):
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
            out_name = f"split_{os.path.basename(pdf_path)}_page_{page_num + 1}.pdf"
            out_path = os.path.join(TEMP_DIR, out_name)
            new_doc.save(out_path)
            new_doc.close()
            split_paths.append(out_path)
        doc.close()
        
        self.log_signal.emit(f"Created {num_pages} temporary single-page PDF files.")
        
        page_jsons = []
        
        try:
            for page_idx, page_path in enumerate(split_paths):
                if not self.is_running:
                    break
                    
                self.log_signal.emit(f"Sending Page {page_idx + 1}/{num_pages} to Local AI Vision Server...")
                
                with open(page_path, "rb") as f:
                    files = {"file": (os.path.basename(page_path), f, "application/pdf")}
                    data = {
                        "prompt": self.prompt_text,
                        "model": self.model_name
                    }
                    
                    response = requests.post(full_url, data=data, files=files, timeout=600)
                    
                if response.status_code != 200:
                    raise RuntimeError(f"Local AI Server Error (Status {response.status_code}): {response.text}")
                    
                resp_data = response.json()
                self.log_signal.emit(f"Page {page_idx + 1} extraction successfully completed.")
                
                # Check response type (dict or string)
                if isinstance(resp_data, dict):
                    page_jsons.append(resp_data)
                else:
                    page_jsons.append(robust_parse_json(str(resp_data)))
                    
        finally:
            # Clean up page files
            for p in split_paths:
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except Exception:
                        pass

        # 3. Combine page JSON files
        self.log_signal.emit("Aggregating all page JSON elements into a combined schema...")
        combined = combine_page_jsons(page_jsons)
        return combined

    # --- Excel Normalizer ---
    def _convert_json_to_df(self, raw_json: dict) -> pd.DataFrame:
        """
        Custom normalizes extracted JSON to keep only user-selected columns.
        Supports standard header-items split tables.
        """
        # If no selected schema is active, export everything at root or normalized
        if not self.selected_schema:
            try:
                # Direct JSON normalizer
                if "items" in raw_json and isinstance(raw_json["items"], list):
                    # Standard json normalize
                    meta_paths = []
                    header = raw_json.get("header")
                    if isinstance(header, dict):
                        meta_paths = [[ "header", k] for k in header.keys()]
                    other = raw_json.get("other infor") or raw_json.get("other_info")
                    if isinstance(other, dict):
                        meta_paths.extend([[ "other infor", k] for k in other.keys()])
                        
                    df = pd.json_normalize(raw_json, record_path="items", meta=meta_paths, errors="ignore")
                else:
                    df = pd.json_normalize(raw_json)
                return df
            except Exception as e:
                self.log_signal.emit(f"Normalization warning: {e}. Outputting plain text rows.")
                return pd.DataFrame([str(raw_json)], columns=["Raw JSON Output"])

        # We have a strict user-defined schema limit!
        # Re-map exactly to columns selected by user
        selected_header = self.selected_schema.get("header", [])
        selected_items = self.selected_schema.get("items", [])
        selected_other = self.selected_schema.get("other infor", [])
        
        rows = []
        
        # Extract items array
        items = raw_json.get("items", [])
        if not isinstance(items, list) or len(items) == 0:
            # Create a single row fallback if no list items exist
            items = [{}]

        header_data = raw_json.get("header", {})
        if not isinstance(header_data, dict):
            header_data = {}
            
        other_data = raw_json.get("other infor") or raw_json.get("other_info") or raw_json.get("summary", {})
        if not isinstance(other_data, dict):
            other_data = {}

        # Loop tabular list
        for item in items:
            row = {}
            # 1. Fill Headers
            for h in selected_header:
                row[h] = header_data.get(h, "")
                
            # 2. Fill List items
            for i in selected_items:
                row[i] = item.get(i, "")
                
            # 3. Fill Other info
            for o in selected_other:
                row[o] = other_data.get(o, "")
                
            rows.append(row)
            
        return pd.DataFrame(rows)


# --- Core UI: PyQt6 MainWindow Layout ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AI PDF-to-Excel Vision Extraction Engine")
        self.setMinimumSize(1400, 850)
        self.resize(1400, 850)
        
        # Properties
        self.settings = settings_manager.load_settings()
        self.active_worker: Optional[ConversionWorker] = None
        self.selected_schema: Optional[Dict[str, List[str]]] = None
        
        self._setup_style()
        self._init_ui()
        self._apply_saved_settings()

    def _setup_style(self):
        """Sets up custom sleek dark-mode stylesheet (wow aesthetics)."""
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(18, 18, 18))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(229, 231, 235))
        self.setPalette(palette)

        qss = """
        QMainWindow, QDialog, QMessageBox {
            background-color: #121212;
        }
        
        QWidget {
            color: #E5E7EB;
            font-family: 'Inter', 'Segoe UI', Arial, sans-serif;
            font-size: 13px;
        }
        
        /* QMessageBox Label Colors */
        QMessageBox QLabel {
            color: #F3F4F6;
        }
        
        /* QComboBox Dropdown Popup View */
        QComboBox QAbstractItemView {
            background-color: #181818;
            border: 1px solid #333333;
            color: #F3F4F6;
            selection-background-color: #6366F1;
            selection-color: white;
            outline: 0px;
        }
        
        QGroupBox {
            font-weight: bold;
            font-size: 14px;
            color: #A855F7; /* Glowing purple headers */
            border: 1px solid #2B2D31;
            border-radius: 8px;
            margin-top: 16px;
            padding-top: 12px;
            background-color: #1E1E1E;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 14px;
            padding: 0 5px;
        }
        
        QListWidget, QTableWidget {
            background-color: #181818;
            border: 1px solid #2B2D31;
            border-radius: 6px;
            padding: 5px;
            color: #D1D5DB;
            gridline-color: transparent;
        }
        QListWidget::item, QTableWidget::item {
            padding: 6px 10px;
            border-bottom: 1px solid #232428;
        }
        QListWidget::item:hover, QTableWidget::item:hover {
            background-color: #2D2D2D;
            border-radius: 4px;
        }
        QListWidget::item:selected, QTableWidget::item:selected {
            background-color: #6366F1;
            color: white;
            border-radius: 4px;
        }
        QHeaderView::section {
            background-color: #1E1E1E;
            color: #A855F7;
            font-weight: bold;
            padding: 6px;
            border: 1px solid #2B2D31;
            border-left: none;
            border-top: none;
        }
        
        QLineEdit, QTextEdit, QSpinBox, QComboBox {
            background-color: #181818;
            border: 1px solid #333333;
            border-radius: 6px;
            padding: 6px 10px;
            color: #F3F4F6;
        }
        QLineEdit:focus, QTextEdit:focus, QSpinBox:focus, QComboBox:focus {
            border: 1px solid #8B5CF6; /* Violet glow on focus */
        }
        
        QPushButton {
            background-color: #2D2D30;
            border: 1px solid #444446;
            border-radius: 6px;
            padding: 8px 16px;
            font-weight: bold;
            color: #F3F4F6;
        }
        QPushButton:hover {
            background-color: #3D3D42;
            border-color: #55555A;
        }
        QPushButton:pressed {
            background-color: #1E1E20;
        }
        
        QProgressBar {
            background-color: #181818;
            border: 1px solid #333333;
            border-radius: 8px;
            text-align: center;
            font-weight: bold;
            color: white;
        }
        QProgressBar::chunk {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #6366F1, stop:1 #A855F7);
            border-radius: 7px;
        }
        
        QScrollBar:vertical {
            border: none;
            background: #121212;
            width: 8px;
            margin: 0;
        }
        QScrollBar::handle:vertical {
            background: #333333;
            min-height: 20px;
            border-radius: 4px;
        }
        QScrollBar::handle:vertical:hover {
            background: #555555;
        }
        """
        self.setStyleSheet(qss)

    def _init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(12)
        
        # Title Banner
        banner_layout = QHBoxLayout()
        title_label = QLabel("🤖 AI PDF-to-Excel Vision Extraction Engine")
        title_label.setFont(QFont("Inter", 16, QFont.Weight.Bold))
        title_label.setStyleSheet("color: #FFFFFF; padding-bottom: 5px;")
        banner_layout.addWidget(title_label)
        
        # Reset schema state label
        self.schema_state_lbl = QLabel("No active filter (outputting all fields)")
        self.schema_state_lbl.setStyleSheet("color: #9CA3AF; font-style: italic;")
        banner_layout.addStretch()
        banner_layout.addWidget(self.schema_state_lbl)
        main_layout.addLayout(banner_layout)

        # Three Column Splitter
        columns_splitter = QSplitter(Qt.Orientation.Horizontal)
        columns_splitter.setHandleWidth(8)
        
        # ==================== LEFT COLUMN: INPUT PANEL ====================
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        input_group = QGroupBox("1. Left Column: Input Folder")
        input_vlayout = QVBoxLayout(input_group)
        input_vlayout.setSpacing(10)
        
        # Input folder picker buttons
        in_folder_btn_layout = QHBoxLayout()
        self.in_folder_le = QLineEdit()
        self.in_folder_le.setReadOnly(True)
        self.in_folder_le.setPlaceholderText("Select Input Folder...")
        self.in_folder_btn = QPushButton("Select Folder")
        self.in_folder_btn.setStyleSheet("background-color: #2E2A47; border-color: #4C3F75; color: #C0B7E5;")
        in_folder_btn_layout.addWidget(self.in_folder_le)
        in_folder_btn_layout.addWidget(self.in_folder_btn)
        input_vlayout.addLayout(in_folder_btn_layout)
        
        # List of PDF Files
        input_vlayout.addWidget(QLabel("Discovered PDF Files:"))
        self.pdf_list_widget = QTableWidget()
        self.pdf_list_widget.setColumnCount(2)
        self.pdf_list_widget.setHorizontalHeaderLabels(["File Name", "Status"])
        self.pdf_list_widget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.pdf_list_widget.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.pdf_list_widget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.pdf_list_widget.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.pdf_list_widget.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.pdf_list_widget.verticalHeader().setVisible(False)
        self.pdf_list_widget.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.pdf_list_widget.setShowGrid(False)
        input_vlayout.addWidget(self.pdf_list_widget)
        
        left_layout.addWidget(input_group)
        columns_splitter.addWidget(left_widget)
        
        # ==================== MIDDLE COLUMN: CONTROLS PANEL ====================
        middle_widget = QWidget()
        middle_layout = QVBoxLayout(middle_widget)
        middle_layout.setContentsMargins(0, 0, 0, 0)
        
        config_group = QGroupBox("2. Config AI Model & Batch Options")
        config_vlayout = QVBoxLayout(config_group)
        config_vlayout.setSpacing(10)
        
        # Grid layout for inputs
        grid_config = QGridLayout()
        grid_config.setSpacing(8)
        
        # 1. AI Selector
        grid_config.addWidget(QLabel("AI Engine:"), 0, 0)
        self.ai_selector = QComboBox()
        self.ai_selector.addItems(["Local AI", "Gemini", "GPT (OpenAI)"])
        grid_config.addWidget(self.ai_selector, 0, 1)
        
        # 2. API Key / URL Box
        self.api_label = QLabel("API Key:")
        grid_config.addWidget(self.api_label, 1, 0)
        self.api_key_le = QLineEdit()
        self.api_key_le.setEchoMode(QLineEdit.EchoMode.Password)
        self.api_key_le.setPlaceholderText("API Key not needed for Local AI")
        grid_config.addWidget(self.api_key_le, 1, 1)
        
        # 3. Model selector
        grid_config.addWidget(QLabel("AI Model:"), 2, 0)
        self.model_selector = QComboBox()
        grid_config.addWidget(self.model_selector, 2, 1)
        
        # 4. Batch limits
        grid_config.addWidget(QLabel("Batch Limit:"), 3, 0)
        self.batch_selector = QComboBox()
        self.batch_selector.addItems(["All Files", "1 File", "5 Files", "10 Files"])
        grid_config.addWidget(self.batch_selector, 3, 1)
        
        # 5. Inter-file delay
        grid_config.addWidget(QLabel("File Delay (sec):"), 4, 0)
        self.delay_spin = QSpinBox()
        self.delay_spin.setRange(0, 60)
        self.delay_spin.setValue(2)
        grid_config.addWidget(self.delay_spin, 4, 1)
        
        # 6. Fail Retry Attempts
        grid_config.addWidget(QLabel("Fail Retry Attempts:"), 5, 0)
        self.fail_attempts_spin = QSpinBox()
        self.fail_attempts_spin.setRange(1, 20)
        self.fail_attempts_spin.setValue(3)
        grid_config.addWidget(self.fail_attempts_spin, 5, 1)
        
        # 7. Fail Retry Delay (sec):
        grid_config.addWidget(QLabel("Fail Retry Delay (sec):"), 6, 0)
        self.fail_delay_spin = QSpinBox()
        self.fail_delay_spin.setRange(0, 300)
        self.fail_delay_spin.setValue(5)
        grid_config.addWidget(self.fail_delay_spin, 6, 1)
        
        config_vlayout.addLayout(grid_config)
        
        # Multiline Prompt Customization Textbox
        config_vlayout.addWidget(QLabel("Data Extraction Prompt Template:"))
        self.prompt_text_edit = QTextEdit()
        self.prompt_text_edit.setMinimumHeight(150)
        self.prompt_text_edit.setAcceptRichText(False)
        config_vlayout.addWidget(self.prompt_text_edit)
        
        # Checkbox for Schema Discovery Page Range
        self.cb_discover_all_pages = QCheckBox("Discover all pages (unchecked = first page only)")
        self.cb_discover_all_pages.setStyleSheet("color: #D1D5DB; font-size: 11px; margin-bottom: 4px;")
        config_vlayout.addWidget(self.cb_discover_all_pages)
        
        # Action Buttons
        btn_action_layout = QHBoxLayout()
        self.btn_discovery = QPushButton("🔍 Discovery Schema (1 File)")
        self.btn_discovery.setStyleSheet("background-color: #1A3E2A; border-color: #2F6B4C; color: #8BECAE; font-weight: bold; padding: 10px;")
        
        self.btn_convert = QPushButton("⚡ START CONVERSION")
        self.btn_convert.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #6366F1, stop:1 #A855F7); color: white; border: none; font-weight: bold; padding: 10px;")
        
        btn_action_layout.addWidget(self.btn_discovery)
        btn_action_layout.addWidget(self.btn_convert)
        config_vlayout.addLayout(btn_action_layout)
        
        # Config Save/Load Buttons
        config_file_layout = QHBoxLayout()
        self.btn_save_config = QPushButton("💾 Save Config")
        self.btn_load_config = QPushButton("📂 Load Config")
        self.btn_save_config.setStyleSheet("background-color: #2E2A47; border-color: #4C3F75; color: #C0B7E5; font-weight: bold;")
        self.btn_load_config.setStyleSheet("background-color: #2E2A47; border-color: #4C3F75; color: #C0B7E5; font-weight: bold;")
        config_file_layout.addWidget(self.btn_save_config)
        config_file_layout.addWidget(self.btn_load_config)
        config_vlayout.addLayout(config_file_layout)
        
        # Progress Bar and Status State
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(18)
        config_vlayout.addWidget(self.progress_bar)
        
        # Active Running Status Label
        status_h = QHBoxLayout()
        self.status_icon = QFrame()
        self.status_icon.setFixedSize(10, 10)
        self.status_icon.setStyleSheet("background-color: #9CA3AF; border-radius: 5px;")
        self.status_text_lbl = QLabel("System Idle")
        self.status_text_lbl.setStyleSheet("color: #9CA3AF; font-weight: bold;")
        status_h.addWidget(self.status_icon)
        status_h.addWidget(self.status_text_lbl)
        status_h.addStretch()
        config_vlayout.addLayout(status_h)
        
        middle_layout.addWidget(config_group)
        columns_splitter.addWidget(middle_widget)
        
        # ==================== RIGHT COLUMN: OUTPUT PANEL ====================
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        output_group = QGroupBox("3. Right Column: Output Folder")
        output_vlayout = QVBoxLayout(output_group)
        output_vlayout.setSpacing(10)
        
        # Output folder picker
        out_folder_btn_layout = QHBoxLayout()
        self.out_folder_le = QLineEdit()
        self.out_folder_le.setReadOnly(True)
        self.out_folder_le.setPlaceholderText("Select Output Folder...")
        self.out_folder_btn = QPushButton("Select Folder")
        self.out_folder_btn.setStyleSheet("background-color: #2E2A47; border-color: #4C3F75; color: #C0B7E5;")
        out_folder_btn_layout.addWidget(self.out_folder_le)
        out_folder_btn_layout.addWidget(self.out_folder_btn)
        output_vlayout.addLayout(out_folder_btn_layout)
        
        # Excel sheet List Box
        output_vlayout.addWidget(QLabel("Finished Excel Files:"))
        self.xlsx_list_widget = QTableWidget()
        self.xlsx_list_widget.setColumnCount(2)
        self.xlsx_list_widget.setHorizontalHeaderLabels(["File Name", "Status"])
        self.xlsx_list_widget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.xlsx_list_widget.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.xlsx_list_widget.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.xlsx_list_widget.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.xlsx_list_widget.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.xlsx_list_widget.verticalHeader().setVisible(False)
        self.xlsx_list_widget.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.xlsx_list_widget.setShowGrid(False)
        output_vlayout.addWidget(self.xlsx_list_widget)
        
        # Output folder action buttons
        out_buttons_layout = QGridLayout()
        out_buttons_layout.setSpacing(6)
        self.btn_refresh_xlsx = QPushButton("🔄 Refresh")
        self.btn_delete_xlsx = QPushButton("🗑️ Delete")
        self.btn_check_xlsx = QPushButton("🔍 Check Form")
        self.btn_combine_xlsx = QPushButton("📊 Combine All")
        
        self.btn_refresh_xlsx.setStyleSheet("background-color: #242526; border-color: #3C3D3E; color: #E5E7EB;")
        self.btn_delete_xlsx.setStyleSheet("background-color: #3E1E1E; border-color: #6B2D2D; color: #FCA5A5;")
        self.btn_check_xlsx.setStyleSheet("background-color: #2E2A47; border-color: #4C3F75; color: #C0B7E5;")
        self.btn_combine_xlsx.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #10B981, stop:1 #059669); color: white; border: none; font-weight: bold;")
        
        out_buttons_layout.addWidget(self.btn_refresh_xlsx, 0, 0)
        out_buttons_layout.addWidget(self.btn_delete_xlsx, 0, 1)
        out_buttons_layout.addWidget(self.btn_check_xlsx, 1, 0)
        out_buttons_layout.addWidget(self.btn_combine_xlsx, 1, 1)
        output_vlayout.addLayout(out_buttons_layout)
        
        right_layout.addWidget(output_group)
        columns_splitter.addWidget(right_widget)
        
        columns_splitter.setSizes([350, 500, 350])
        main_layout.addWidget(columns_splitter)
        
        # ==================== BOTTOM STATUS LOG BOX ====================
        log_group = QGroupBox("Status Execution Logs")
        log_group.setMaximumHeight(150)
        log_v = QVBoxLayout(log_group)
        log_v.setContentsMargins(8, 8, 8, 8)
        self.log_widget = QListWidget()
        self.log_widget.setStyleSheet("background-color: #0E0E10; border: none; font-family: monospace; font-size: 11px;")
        log_v.addWidget(self.log_widget)
        main_layout.addWidget(log_group)
        
        # Connections
        self.in_folder_btn.clicked.connect(self._select_input_folder)
        self.out_folder_btn.clicked.connect(self._select_output_folder)
        self.ai_selector.currentIndexChanged.connect(self._on_ai_engine_changed)
        self.btn_discovery.clicked.connect(self._run_schema_discovery)
        self.btn_convert.clicked.connect(self._run_batch_conversion)
        self.btn_save_config.clicked.connect(self._on_save_config_clicked)
        self.btn_load_config.clicked.connect(self._on_load_config_clicked)
        self.btn_refresh_xlsx.clicked.connect(self._scan_output_folder)
        self.btn_delete_xlsx.clicked.connect(self._delete_xlsx_file)
        self.btn_check_xlsx.clicked.connect(self._check_xlsx_schema)
        self.btn_combine_xlsx.clicked.connect(self._combine_all_excel)
        
        self.pdf_list_widget.cellDoubleClicked.connect(self._on_pdf_double_clicked)
        self.xlsx_list_widget.cellDoubleClicked.connect(self._on_xlsx_double_clicked)
        
        # Shadow effect on groups
        for g in [input_group, config_group, output_group, log_group]:
            shadow = QGraphicsDropShadowEffect(g)
            shadow.setBlurRadius(15)
            shadow.setColor(QColor(0, 0, 0, 80))
            shadow.setOffset(0, 4)
            g.setGraphicsEffect(shadow)

    def _apply_saved_settings(self):
        """Loads and pre-populates GUI fields with last saved state values."""
        self.in_folder_le.setText(self.settings.get("input_folder", ""))
        self.out_folder_le.setText(self.settings.get("output_folder", ""))
        
        self.ai_selector.setCurrentText(self.settings.get("ai_engine", "Local AI"))
        self._on_ai_engine_changed() # Triggers key boxes & models list update
        
        # Load API keys / URL
        engine = self.ai_selector.currentText()
        if engine == "Gemini":
            self.api_key_le.setText(self.settings.get("gemini_api_key", ""))
        elif engine == "GPT (OpenAI)":
            self.api_key_le.setText(self.settings.get("gpt_api_key", ""))
        else:
            self.api_key_le.setText(self.settings.get("local_api_url", "http://api-localai.germantest.net"))
        
        # Load models
        engine = self.ai_selector.currentText()
        if engine == "Gemini":
            self.model_selector.setCurrentText(self.settings.get("gemini_model", "gemini-2.5-flash"))
        elif engine == "GPT (OpenAI)":
            self.model_selector.setCurrentText(self.settings.get("gpt_model", "gpt-4o-mini"))
        else:
            self.model_selector.setCurrentText(self.settings.get("local_model", "qwen2.5vl:7b"))
            
        self.batch_selector.setCurrentText(self.settings.get("batch_limit", "All Files"))
        self.delay_spin.setValue(self.settings.get("inter_file_delay", 2))
        self.fail_attempts_spin.setValue(self.settings.get("model_fail_attempts", 3))
        self.fail_delay_spin.setValue(self.settings.get("model_fail_delay", 5))
        self.prompt_text_edit.setPlainText(self.settings.get("custom_prompt", ""))
        self.cb_discover_all_pages.setChecked(self.settings.get("discover_all_pages", False))
        
        # Load schema
        self.selected_schema = self.settings.get("selected_schema", None)
        if self.selected_schema:
            selected = self.selected_schema
            self.schema_state_lbl.setText(f"Filter active ({len(selected.get('header', []))}H / {len(selected.get('items', []))}I / {len(selected.get('other infor', []))}O)")
            self.schema_state_lbl.setStyleSheet("color: #10B981; font-weight: bold; font-style: normal;")
        else:
            self.schema_state_lbl.setText("No active filter (outputting all fields)")
            self.schema_state_lbl.setStyleSheet("color: #9CA3AF; font-style: italic;")
            
        # Scan folders if already saved
        self._scan_input_folder()
        self._scan_output_folder()

    def _save_active_settings(self):
        """Extracts values from GUI and saves to app_settings.json."""
        self.settings["input_folder"] = self.in_folder_le.text()
        self.settings["output_folder"] = self.out_folder_le.text()
        
        engine = self.ai_selector.currentText()
        self.settings["ai_engine"] = engine
        
        if engine == "Gemini":
            self.settings["gemini_api_key"] = self.api_key_le.text()
            self.settings["gemini_model"] = self.model_selector.currentText()
        elif engine == "GPT (OpenAI)":
            self.settings["gpt_api_key"] = self.api_key_le.text()
            self.settings["gpt_model"] = self.model_selector.currentText()
        else:
            self.settings["local_api_url"] = self.api_key_le.text()
            self.settings["local_model"] = self.model_selector.currentText()
            
        self.settings["batch_limit"] = self.batch_selector.currentText()
        self.settings["inter_file_delay"] = self.delay_spin.value()
        self.settings["model_fail_attempts"] = self.fail_attempts_spin.value()
        self.settings["model_fail_delay"] = self.fail_delay_spin.value()
        self.settings["custom_prompt"] = self.prompt_text_edit.toPlainText()
        self.settings["selected_schema"] = self.selected_schema
        self.settings["discover_all_pages"] = self.cb_discover_all_pages.isChecked()
        
        settings_manager.save_settings(self.settings)

    # --- Folder Scanning Routines ---
    def _select_input_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select PDF Input Folder", self.in_folder_le.text())
        if folder:
            self.in_folder_le.setText(folder)
            self._scan_input_folder()
            self._save_active_settings()

    def _select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Excel Output Folder", self.out_folder_le.text())
        if folder:
            self.out_folder_le.setText(folder)
            self._scan_output_folder()
            self._save_active_settings()

    def _scan_input_folder(self):
        """Scans input folder path and populates left table widget with PDFs and their conversion status."""
        self.pdf_list_widget.setRowCount(0)
        folder = self.in_folder_le.text()
        output_folder = self.out_folder_le.text()
        if not folder or not os.path.exists(folder):
            return
            
        for f in sorted(os.listdir(folder)):
            if f.lower().endswith(".pdf"):
                status_text = "Pending"
                status_color = "#F59E0B" # Amber
                
                if output_folder and os.path.exists(output_folder):
                    xlsx_filename = os.path.splitext(f)[0] + ".xlsx"
                    xlsx_path = os.path.join(output_folder, xlsx_filename)
                    if os.path.exists(xlsx_path):
                        status_text = "Done"
                        status_color = "#10B981" # Green
                
                row_idx = self.pdf_list_widget.rowCount()
                self.pdf_list_widget.insertRow(row_idx)
                
                name_item = QTableWidgetItem(f)
                status_item = QTableWidgetItem(status_text)
                status_item.setForeground(QBrush(QColor(status_color)))
                
                self.pdf_list_widget.setItem(row_idx, 0, name_item)
                self.pdf_list_widget.setItem(row_idx, 1, status_item)

    def _get_excel_columns(self, filepath: str) -> List[str]:
        """Reads only the first row of an Excel file to quickly extract column names, ensuring workbook is closed."""
        wb = None
        cols = []
        try:
            wb = openpyxl.load_workbook(filepath, read_only=True)
            sheet = wb.active
            for row in sheet.iter_rows(max_row=1, values_only=True):
                cols = [str(cell) for cell in row if cell is not None]
                break
        except Exception as e:
            self.add_log(f"Error reading columns from {os.path.basename(filepath)}: {e}")
        finally:
            if wb is not None:
                wb.close()
        return cols


    def _scan_output_folder(self):
        """Scans output folder path and populates right list widget with Excel sheets, highlighting column mismatches."""
        self.xlsx_list_widget.setRowCount(0)
        folder = self.out_folder_le.text()
        if not folder or not os.path.exists(folder):
            return
            
        expected_cols = []
        if self.selected_schema:
            expected_cols = (self.selected_schema.get("header", []) + 
                             self.selected_schema.get("items", []) + 
                             self.selected_schema.get("other infor", []))

        for f in sorted(os.listdir(folder)):
            if f.lower().endswith(".xlsx") and not f.startswith("~$"):
                filepath = os.path.join(folder, f)
                
                status_text = "Checking..."
                status_color = "#9CA3AF"
                tooltip = ""
                
                if expected_cols:
                    excel_cols = self._get_excel_columns(filepath)
                    expected_set = set(expected_cols)
                    excel_set = set(excel_cols)
                    
                    missing = expected_set - excel_set
                    extra = excel_set - expected_set
                    
                    if not missing and not extra:
                        status_text = "Matched"
                        status_color = "#10B981"
                        tooltip = "Columns match the active schema perfectly!"
                    else:
                        status_text = "Not Matched"
                        status_color = "#EF4444"
                        tooltip = "Schema column mismatch!\n"
                        if missing:
                            tooltip += f"Missing: {list(missing)}\n"
                        if extra:
                            tooltip += f"Extra: {list(extra)}"
                else:
                    status_text = "No Schema"
                    status_color = "#9CA3AF"
                    tooltip = "No active schema to check against."
                
                row_idx = self.xlsx_list_widget.rowCount()
                self.xlsx_list_widget.insertRow(row_idx)
                
                name_item = QTableWidgetItem(f)
                name_item.setToolTip(tooltip)
                
                status_item = QTableWidgetItem(status_text)
                status_item.setForeground(QBrush(QColor(status_color)))
                status_item.setToolTip(tooltip)
                
                self.xlsx_list_widget.setItem(row_idx, 0, name_item)
                self.xlsx_list_widget.setItem(row_idx, 1, status_item)
                
        # Automatically sync the left-side PDF conversion status columns
        self._scan_input_folder()

    def _on_pdf_double_clicked(self, row, column):
        name_item = self.pdf_list_widget.item(row, 0)
        if not name_item:
            return
        folder = self.in_folder_le.text()
        if not folder or not os.path.exists(folder):
            QMessageBox.warning(self, "Folder Missing", "Please select a valid PDF Input Folder first.")
            return
            
        filepath = os.path.join(folder, name_item.text())
        if not os.path.exists(filepath):
            QMessageBox.warning(self, "File Not Found", f"PDF file does not exist: {filepath}")
            return
            
        dialog = PDFViewerDialog(filepath, self)
        dialog.exec()

    def _on_xlsx_double_clicked(self, row, column):
        name_item = self.xlsx_list_widget.item(row, 0)
        if not name_item:
            return
            
        folder = self.out_folder_le.text()
        if not folder or not os.path.exists(folder):
            QMessageBox.warning(self, "Folder Missing", "Please select a valid Excel Output Folder first.")
            return
            
        filepath = os.path.join(folder, name_item.text())
        if not os.path.exists(filepath):
            QMessageBox.warning(self, "File Not Found", f"Excel file does not exist: {filepath}")
            return
            
        dialog = ExcelViewerDialog(filepath, self)
        dialog.exec()

    # --- AI State Toggle Handler ---
    def _on_ai_engine_changed(self):
        engine = self.ai_selector.currentText()
        self.model_selector.clear()
        
        if engine == "Local AI":
            self.api_label.setText("API URL:")
            self.api_key_le.setEnabled(True)
            self.api_key_le.setEchoMode(QLineEdit.EchoMode.Normal)
            self.api_key_le.setPlaceholderText("Enter Local AI Server URL (e.g. http://api-localai.germantest.net)...")
            self.api_key_le.setText(self.settings.get("local_api_url", "http://api-localai.germantest.net"))
            
            # Local models list matching qwen
            self.model_selector.addItems(["qwen2.5vl:7b", "qwen2.5vl:3b", "qwen2.5vl:32b"])
        elif engine == "Gemini":
            self.api_label.setText("API Key:")
            self.api_key_le.setEnabled(True)
            self.api_key_le.setEchoMode(QLineEdit.EchoMode.Password)
            self.api_key_le.setPlaceholderText("Enter Google Gemini API Key...")
            self.api_key_le.setText(self.settings.get("gemini_api_key", ""))
            
            # Gemini models list
            self.model_selector.addItems([
                "gemini-2.5-flash",
                "gemini-2.5-pro",
                "gemini-1.5-flash",
                "gemini-1.5-pro"
            ])
        else: # GPT (OpenAI)
            self.api_label.setText("API Key:")
            self.api_key_le.setEnabled(True)
            self.api_key_le.setEchoMode(QLineEdit.EchoMode.Password)
            self.api_key_le.setPlaceholderText("Enter OpenAI API Key (sk-...)...")
            self.api_key_le.setText(self.settings.get("gpt_api_key", ""))
            
            # GPT models
            self.model_selector.addItems([
                "gpt-4o-mini",
                "gpt-4o"
            ])

    # --- Log Appender ---
    def add_log(self, text: str):
        item = QListWidgetItem(f"[{time.strftime('%H:%M:%S')}] {text}")
        self.log_widget.addItem(item)
        self.log_widget.scrollToBottom()

    # --- UI Active Status Toggle ---
    def set_gui_active(self, active: bool, text: str = "System Idle"):
        self.in_folder_btn.setEnabled(active)
        self.out_folder_btn.setEnabled(active)
        self.ai_selector.setEnabled(active)
        self.api_key_le.setEnabled(active if self.ai_selector.currentText() != "Local AI" else False)
        self.model_selector.setEnabled(active)
        self.batch_selector.setEnabled(active)
        self.delay_spin.setEnabled(active)
        self.fail_attempts_spin.setEnabled(active)
        self.fail_delay_spin.setEnabled(active)
        self.prompt_text_edit.setEnabled(active)
        self.btn_discovery.setEnabled(active)
        
        if not active:
            self.status_icon.setStyleSheet("background-color: #3B82F6; border-radius: 5px;") # Active indicator
            self.status_text_lbl.setText(text)
            self.btn_convert.setText("🛑 CANCEL / ABORT")
            self.btn_convert.setStyleSheet("background-color: #EF4444; color: white; border: none; font-weight: bold; padding: 10px;")
        else:
            self.status_icon.setStyleSheet("background-color: #10B981; border-radius: 5px;") # Success indicator
            self.status_text_lbl.setText(text)
            self.btn_convert.setText("⚡ START CONVERSION")
            self.btn_convert.setStyleSheet("background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #6366F1, stop:1 #A855F7); color: white; border: none; font-weight: bold; padding: 10px;")

    # --- Action: Run Schema Column Discovery ---
    def _run_schema_discovery(self):
        # 1. Validation checks
        selected_indexes = self.pdf_list_widget.selectedIndexes()
        if not selected_indexes:
            QMessageBox.warning(self, "Selection Required", "Please select one PDF file from the left list to discover schema.")
            return

        row = selected_indexes[0].row()
        selected_pdf = self.pdf_list_widget.item(row, 0).text()
        input_folder = self.in_folder_le.text()
        full_pdf_path = os.path.join(input_folder, selected_pdf)

        if not os.path.exists(full_pdf_path):
            QMessageBox.critical(self, "Error", f"Selected PDF file does not exist: {full_pdf_path}")
            return
            
        # Initialize key/URL validation
        engine = self.ai_selector.currentText()
        api_val = self.api_key_le.text().strip()
        if not api_val:
            QMessageBox.warning(self, "Input Required", f"Please enter your {engine} {'API URL' if engine == 'Local AI' else 'API Key'} in the middle panel first.")
            return

        self._save_active_settings()
        self.set_gui_active(False, "Schema Discovery Running...")
        self.progress_bar.setValue(20)
        self.add_log(f"Starting schema extraction on '{selected_pdf}'...")

        # Spawn Thread Worker (Use general prompt for discovery)
        self.active_worker = ConversionWorker(
            files=[full_pdf_path],
            ai_engine=engine,
            api_key=api_val,
            model_name=self.model_selector.currentText(),
            prompt_text=(
                "please extract header and item infromation from this document.\n"
                "Return a structured JSON object with the following keys:\n"
                "- \"header\": A dictionary of fields that apply to the entire document (e.g., invoice/annex number, contract/document date, supplier/customer details, etc.).\n"
                "- \"items\": A list of dictionaries representing the rows of any tabular items (e.g., product name, style number, quantity, unit price, total price, color, delivery date, etc.).\n"
                "- \"other infor\": A dictionary of summary information (e.g., total quantity, total amount, general notes).\n\n"
                "Return only the raw JSON object, without markdown formatting or code blocks."
            ),
            output_dir=self.out_folder_le.text(),
            delay=0,
            is_schema_discovery=True,
            fail_attempts=self.fail_attempts_spin.value(),
            fail_delay=self.fail_delay_spin.value(),
            discover_all_pages=self.cb_discover_all_pages.isChecked()
        )

        self.active_worker.log_signal.connect(self.add_log)
        self.active_worker.progress_signal.connect(lambda p, t: self.progress_bar.setValue(p))
        self.active_worker.schema_discovered_signal.connect(self._on_schema_discovered)
        self.active_worker.batch_finished_signal.connect(self._on_schema_discovery_finished)
        self.active_worker.start()

    def _on_schema_discovered(self, raw_json: dict):
        """Called when thread successfully extracts JSON schema structure."""
        self.add_log("Extracting keys structure... launching Column Selection Dialog.")
        
        # Pop Column selector dialog
        dialog = ColumnSelectorDialog(raw_json, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            # User accepted selected schema
            selected = dialog.get_selected_schema()
            self.selected_schema = selected
            self.add_log(f"Selected columns schema: Header={len(selected['header'])}, Items={len(selected['items'])}, Other={len(selected['other infor'])}")
            
            # Format custom reference prompt
            self._update_prompt_with_schema(selected)
            self.schema_state_lbl.setText(f"Filter active ({len(selected['header'])}H / {len(selected['items'])}I / {len(selected['other infor'])}O)")
            self.schema_state_lbl.setStyleSheet("color: #10B981; font-weight: bold; font-style: normal;")
        else:
            self.add_log("Schema columns selection aborted.")

    def _on_schema_discovery_finished(self, success: bool, msg: str):
        self.progress_bar.setValue(100 if success else 0)
        self.set_gui_active(True, "System Idle")
        if success:
            self.add_log("Schema discovery fully completed.")
        else:
            QMessageBox.critical(self, "Discovery Failed", f"Failed to extract schema structure:\n{msg}")

    def _update_prompt_with_schema(self, schema: Dict[str, List[str]]):
        """Constructs a strict customized JSON prompt using the chosen columns list."""
        header_template = {}
        for h in schema["header"]:
            header_template[h] = f"content of {h}"
            
        items_template = {}
        for i in schema["items"]:
            items_template[i] = f"content of item1 {i}"
            
        other_template = {}
        for o in schema["other infor"]:
            other_template[o] = f"content of {o}"
            
        custom_json = {
            "header": header_template,
            "items": [items_template],
            "other infor": other_template
        }
        
        json_formatted = json.dumps(custom_json, indent=2, ensure_ascii=False)
        
        custom_prompt = f"""please extract header and item infromation from this file and put in json format. return only json without any additional text
this json template is your reference only. json should content actual extracted data

{json_formatted}"""
        
        self.prompt_text_edit.setPlainText(custom_prompt)
        self._save_active_settings()
        self.add_log("Prompt template and default JSON schema updated successfully.")

    # --- Action: Run Full Batch Conversion ---
    def _run_batch_conversion(self):
        # Cancel if thread active
        if self.active_worker and self.active_worker.isRunning():
            self.add_log("Cancelling conversions batch...")
            self.active_worker.is_running = False
            self.btn_convert.setEnabled(False)
            return

        # 1. Validation Checks
        input_folder = self.in_folder_le.text()
        output_folder = self.out_folder_le.text()
        
        if not input_folder or not os.path.exists(input_folder):
            QMessageBox.warning(self, "Paths Missing", "Please select a valid Input Folder on the Left side first.")
            return
        if not output_folder:
            QMessageBox.warning(self, "Paths Missing", "Please select a valid Output Folder on the Right side first.")
            return

        # Gather list of PDFs
        all_pdfs = [os.path.join(input_folder, self.pdf_list_widget.item(i, 0).text()) 
                    for i in range(self.pdf_list_widget.rowCount())]
        
        selected_indexes = self.pdf_list_widget.selectedIndexes()
        if selected_indexes:
            start_row = selected_indexes[0].row()
            start_pdf_name = self.pdf_list_widget.item(start_row, 0).text()
            self.add_log(f"Starting batch from selected file: {start_pdf_name} (index {start_row})")
        else:
            start_row = 0
            self.add_log("No specific file selected. Starting batch from the first file.")

        # Parse Batch Limit
        limit_str = self.batch_selector.currentText()
        limit = None
        if limit_str == "1 File":
            limit = 1
        elif limit_str == "5 Files":
            limit = 5
        elif limit_str == "10 Files":
            limit = 10

        # Collect only files not yet done starting from start_row up to the batch limit
        files_to_convert = []
        for i in range(start_row, self.pdf_list_widget.rowCount()):
            pdf_path = all_pdfs[i]
            filename = os.path.basename(pdf_path)
            xlsx_filename = os.path.splitext(filename)[0] + ".xlsx"
            xlsx_path = os.path.join(output_folder, xlsx_filename)
            
            if os.path.exists(xlsx_path):
                continue
                
            files_to_convert.append(pdf_path)
            if limit is not None and len(files_to_convert) >= limit:
                break

        if not files_to_convert:
            if selected_indexes:
                msg = f"All PDF files starting from '{start_pdf_name}' to the end of the list are already converted."
            else:
                msg = "All PDF files in the input directory are already converted."
            QMessageBox.information(self, "No Files to Convert", msg)
            return

        # Key/URL validation
        engine = self.ai_selector.currentText()
        api_val = self.api_key_le.text().strip()
        if not api_val:
            QMessageBox.warning(self, "Input Required", f"Please enter your {engine} {'API URL' if engine == 'Local AI' else 'API Key'} first.")
            return

        # Save settings
        self._save_active_settings()
        self.set_gui_active(False, "Converting Batch...")
        self.progress_bar.setValue(0)
        
        # Spawn thread worker
        self.active_worker = ConversionWorker(
            files=files_to_convert,
            ai_engine=engine,
            api_key=api_val,
            model_name=self.model_selector.currentText(),
            prompt_text=self.prompt_text_edit.toPlainText(),
            output_dir=output_folder,
            delay=self.delay_spin.value(),
            selected_schema=self.selected_schema,
            fail_attempts=self.fail_attempts_spin.value(),
            fail_delay=self.fail_delay_spin.value()
        )
        
        self.active_worker.log_signal.connect(self.add_log)
        self.active_worker.progress_signal.connect(self._on_batch_progress)
        self.active_worker.file_finished_signal.connect(self._on_file_converted)
        self.active_worker.batch_finished_signal.connect(self._on_batch_finished)
        self.active_worker.start()

    def _on_batch_progress(self, percent: int, text: str):
        self.progress_bar.setValue(percent)
        self.status_text_lbl.setText(text)

    def _on_file_converted(self, pdf_path: str, xlsx_path: str):
        self.add_log(f"Successfully finished file conversion: {os.path.basename(xlsx_path)}")
        self._scan_output_folder() # Refresh output folders list widget

    def _on_batch_finished(self, success: bool, msg: str):
        self.set_gui_active(True, "System Idle")
        self.btn_convert.setEnabled(True)
        self.progress_bar.setValue(100 if success else 0)
        self._scan_output_folder()
        
        if success:
            QMessageBox.information(self, "Batch Done", f"Batch conversion fully finished:\n{msg}")
            self.add_log("Batch execution completed successfully.")
        else:
            QMessageBox.critical(self, "Batch Error", f"Batch conversion failed:\n{msg}")
            self.add_log("Batch execution finished with error.")
            
        self.active_worker = None

    def _on_save_config_clicked(self):
        self._save_active_settings()
        filepath, _ = QFileDialog.getSaveFileName(self, "Save Configuration File", "", "JSON Files (*.json)")
        if filepath:
            try:
                with open(filepath, "w", encoding="utf-8") as f:
                    json.dump(self.settings, f, indent=4, ensure_ascii=False)
                self.add_log(f"Configuration saved manually to: {os.path.basename(filepath)}")
                QMessageBox.information(self, "Config Saved", f"Configuration successfully saved to:\n{filepath}")
            except Exception as e:
                self.add_log(f"Failed to save config: {e}")
                QMessageBox.critical(self, "Save Error", f"Failed to save configuration:\n{e}")

    def _on_load_config_clicked(self):
        filepath, _ = QFileDialog.getOpenFileName(self, "Load Configuration File", "", "JSON Files (*.json)")
        if filepath:
            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                
                loaded_settings = settings_manager.DEFAULT_SETTINGS.copy()
                loaded_settings.update(data)
                
                self.settings = loaded_settings
                self._apply_saved_settings()
                
                self.selected_schema = self.settings.get("selected_schema")
                if self.selected_schema:
                    selected = self.selected_schema
                    self.schema_state_lbl.setText(f"Filter active ({len(selected.get('header', []))}H / {len(selected.get('items', []))}I / {len(selected.get('other infor', []))}O)")
                    self.schema_state_lbl.setStyleSheet("color: #10B981; font-weight: bold; font-style: normal;")
                else:
                    self.schema_state_lbl.setText("No active filter (outputting all fields)")
                    self.schema_state_lbl.setStyleSheet("color: #9CA3AF; font-style: italic;")
                
                self._save_active_settings()
                self.add_log(f"Configuration manually loaded from: {os.path.basename(filepath)}")
                QMessageBox.information(self, "Config Loaded", f"Configuration successfully loaded from:\n{filepath}")
            except Exception as e:
                self.add_log(f"Failed to load config: {e}")
                QMessageBox.critical(self, "Load Error", f"Failed to load configuration:\n{e}")

    def _delete_xlsx_file(self):
        selected_indexes = self.xlsx_list_widget.selectedIndexes()
        if not selected_indexes:
            QMessageBox.warning(self, "Selection Required", "Please select an Excel file to delete.")
            return

        row = selected_indexes[0].row()
        filename = self.xlsx_list_widget.item(row, 0).text()
        output_folder = self.out_folder_le.text()
        filepath = os.path.join(output_folder, filename)
        
        if not os.path.exists(filepath):
            QMessageBox.critical(self, "Error", f"File does not exist: {filepath}")
            return
            
        reply = QMessageBox.question(
            self, 
            "Confirm Delete", 
            f"Are you sure you want to permanently delete the file '{filename}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            try:
                os.remove(filepath)
                self.add_log(f"Deleted file: {filename}")
                self._scan_output_folder()
            except Exception as e:
                self.add_log(f"Failed to delete file {filename}: {e}")
                QMessageBox.critical(self, "Delete Error", f"Failed to delete file:\n{e}")

    def _check_xlsx_schema(self):
        selected_indexes = self.xlsx_list_widget.selectedIndexes()
        if not selected_indexes:
            QMessageBox.warning(self, "Selection Required", "Please select an Excel file to check.")
            return

        row = selected_indexes[0].row()
        filename = self.xlsx_list_widget.item(row, 0).text()
        output_folder = self.out_folder_le.text()
        filepath = os.path.join(output_folder, filename)
        
        if not os.path.exists(filepath):
            QMessageBox.critical(self, "Error", f"Selected file does not exist: {filepath}")
            return
            
        if not self.selected_schema:
            QMessageBox.information(self, "No Active Schema", "No active schema selected. Please run Schema Discovery first.")
            return

        expected_cols = (self.selected_schema.get("header", []) + 
                         self.selected_schema.get("items", []) + 
                         self.selected_schema.get("other infor", []))
                         
        if not expected_cols:
            QMessageBox.information(self, "Empty Schema", "The active schema does not define any columns.")
            return
            
        excel_cols = self._get_excel_columns(filepath)
        
        expected_set = set(expected_cols)
        excel_set = set(excel_cols)
        
        missing = expected_set - excel_set
        extra = excel_set - expected_set
        
        if not missing and not extra:
            QMessageBox.information(self, "Schema Check", f"✅ Match!\n\n'{filename}' matches the active schema perfectly.")
        else:
            msg = f"❌ Schema Mismatch for '{filename}'!\n\n"
            if missing:
                msg += f"Missing Columns ({len(missing)}):\n- " + "\n- ".join(sorted(missing)) + "\n\n"
            if extra:
                msg += f"Extra Columns ({len(extra)}):\n- " + "\n- ".join(sorted(extra)) + "\n\n"
            msg += f"Expected columns: {expected_cols}\n\nActual columns: {excel_cols}"
            QMessageBox.warning(self, "Schema Check Failed", msg)

    def _combine_all_excel(self):
        output_folder = self.out_folder_le.text()
        if not output_folder or not os.path.exists(output_folder):
            QMessageBox.warning(self, "Folder Missing", "Please select a valid Output Folder first.")
            return
            
        xlsx_files = [f for f in os.listdir(output_folder) if f.lower().endswith(".xlsx") and not f.startswith("~$") and f != "combined_output.xlsx"]
        if not xlsx_files:
            QMessageBox.warning(self, "No Files", "No Excel files found in output folder to combine.")
            return
            
        save_path, _ = QFileDialog.getSaveFileName(
            self, 
            "Save Combined Excel File", 
            os.path.join(output_folder, "combined_output.xlsx"), 
            "Excel Files (*.xlsx)"
        )
        if not save_path:
            return
            
        self.add_log(f"Combining {len(xlsx_files)} Excel files...")
        
        dfs = []
        for f in xlsx_files:
            filepath = os.path.join(output_folder, f)
            if os.path.abspath(filepath) == os.path.abspath(save_path):
                continue
            try:
                df = pd.read_excel(filepath)
                df.insert(0, "Source File", f)
                dfs.append(df)
            except Exception as e:
                self.add_log(f"Failed to read file {f}: {e}")
                
        if not dfs:
            QMessageBox.warning(self, "Combine Failed", "No valid data could be combined.")
            return
            
        try:
            combined_df = pd.concat(dfs, ignore_index=True)
            combined_df.to_excel(save_path, index=False)
            self.add_log(f"✅ Combined file created: {os.path.basename(save_path)}")
            QMessageBox.information(self, "Combine Success", f"Successfully combined {len(dfs)} files into:\n{save_path}")
            self._scan_output_folder()
        except Exception as e:
            self.add_log(f"Error combining files: {e}")
            QMessageBox.critical(self, "Combine Error", f"Failed to combine files:\n{e}")


# --- Entry Point ---
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Modern dark window theme configuration
    app.setStyle("Fusion")
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
