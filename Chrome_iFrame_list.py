import sys
import time
from typing import Optional, List, Dict, Any

# --- External Library Imports ---
from selenium import webdriver

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QDialogButtonBox,
    QComboBox, QWidget, QGroupBox, QMessageBox, QRadioButton, QHBoxLayout
)
from PyQt6.QtCore import Qt

# --- Main App Imports (with fallback for standalone testing) ---
try:
    from my_lib.shared_context import ExecutionContext
except ImportError:
    print("Warning: Could not import main app libraries. Using fallback for ExecutionContext.")
    class ExecutionContext:
        def add_log(self, message: str): print(f"LOG: {message}")
        def get_variable(self, name: str, default=None): return None
        def set_variable(self, name: str, value: Any): print(f"Set @{name} = {value}")

# --- JavaScript payload to detect all iFrames and Shadow Hosts recursively ---
JS_LIST_CONTEXTS = """
const results = { iframes: [], shadow_hosts: [] };
function getXPath(el) {
    let path = [];
    while (el && el.nodeType === 1) {
        let index = 0, sibling = el.previousSibling;
        while (sibling) {
            if (sibling.nodeType === 1 && sibling.tagName === el.tagName) index++;
            sibling = sibling.previousSibling;
        }
        const tagName = (el.tagName || '').toLowerCase();
        const pathIndex = (index > 0) ? `[${index + 1}]` : '';
        path.unshift(tagName + pathIndex);
        el = el.parentNode;
    }
    return path.join('/');
}
function findElements(contextNode, docPath) {
    const iframes = contextNode.querySelectorAll('iframe');
    for (const iframe of iframes) {
        const framePath = getXPath(iframe);
        const fullPath = docPath ? `${docPath} -> IFRAME` : `IFRAME`;
        let info = { xpath: framePath, full_path: `${fullPath}(${framePath})`, id: iframe.id || '', name: iframe.name || '' };
        results.iframes.push(info);
        try {
            if (iframe.contentDocument) findElements(iframe.contentDocument, fullPath);
        } catch (e) { info.error = 'Cross-origin iframe'; }
    }
    const allElements = contextNode.querySelectorAll('*');
    for (const el of allElements) {
        if (el.shadowRoot) {
            const hostPath = getXPath(el);
            const fullPath = docPath ? `${docPath} -> SHADOW_HOST` : `SHADOW_HOST`;
            results.shadow_hosts.push({ host_xpath: hostPath, full_path: `${fullPath}(${hostPath})`, host_id: el.id || '', host_tag: el.tagName.toLowerCase(), mode: el.shadowRoot.mode });
            findElements(el.shadowRoot, fullPath);
        }
    }
}
findElements(document, 'DOCUMENT');
return results;
"""


class _ScanContextsConfigDialog(QDialog):
    def __init__(self, global_variables: List[str], parent: Optional[QWidget] = None,
                 initial_config: Optional[Dict[str, Any]] = None,
                 initial_variable: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("Configure Get iFrames & Shadow DOMs")
        self.setMinimumWidth(500)
        self.global_variables = global_variables

        main_layout = QVBoxLayout(self)

        # 1. Select Input WebDriver Variable Group
        driver_group = QGroupBox("Select WebDriver Source")
        driver_layout = QFormLayout(driver_group)
        self.driver_var_combo = QComboBox()
        self.driver_var_combo.addItems(["-- Select Variable --"] + self.global_variables)
        driver_layout.addRow("WebDriver Variable:", self.driver_var_combo)
        main_layout.addWidget(driver_group)

        # 2. Assign Result Group
        assign_group = QGroupBox("Assign Comma-Separated List to Variable")
        assign_layout = QFormLayout(assign_group)
        self.new_var_radio = QRadioButton("New Variable Name:")
        self.new_var_input = QLineEdit("context_list")
        self.existing_var_radio = QRadioButton("Existing Variable:")
        self.existing_var_combo = QComboBox()
        self.existing_var_combo.addItems(["-- Select --"] + self.global_variables)
        self.new_var_radio.setChecked(True)

        assign_layout.addRow(self.new_var_radio, self.new_var_input)
        assign_layout.addRow(self.existing_var_radio, self.existing_var_combo)
        main_layout.addWidget(assign_group)

        # Dialog Buttons
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        main_layout.addWidget(self.button_box)

        if initial_config:
            self._populate_from_initial_config(initial_config, initial_variable)

    def _populate_from_initial_config(self, config: Dict[str, Any], variable: Optional[str]):
        driver_var = config.get("driver_var", "")
        if driver_var in self.global_variables:
            self.driver_var_combo.setCurrentText(driver_var)

        if variable:
            if variable in self.global_variables:
                self.existing_var_radio.setChecked(True)
                self.existing_var_combo.setCurrentText(variable)
            else:
                self.new_var_radio.setChecked(True)
                self.new_var_input.setText(variable)

    def get_executor_method_name(self) -> str:
        return "_execute_scan_contexts"

    def get_assignment_variable(self) -> Optional[str]:
        if self.new_var_radio.isChecked():
            var_name = self.new_var_input.text().strip()
            if not var_name:
                QMessageBox.warning(self, "Input Error", "New variable name cannot be empty.")
                return None
            return var_name
        else:
            var_name = self.existing_var_combo.currentText()
            if var_name == "-- Select --":
                QMessageBox.warning(self, "Input Error", "Please select an existing variable.")
                return None
            return var_name

    def get_config_data(self) -> Optional[Dict[str, Any]]:
        driver_var = self.driver_var_combo.currentText()
        if driver_var == "-- Select Variable --":
            QMessageBox.warning(self, "Input Error", "Please select a WebDriver variable.")
            return None

        return {
            "driver_var": driver_var
        }


class ScanContexts_API:
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context

    def _log(self, message: str, level: str = "INFO"):
        if self.context:
            self.context.add_log(f"ScanContexts_API [{level}]: {message}")
        else:
            print(f"ScanContexts_API [{level}]: {message}")

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening Scan Contexts configuration dialog.")
        return _ScanContextsConfigDialog(
            global_variables=global_variables,
            parent=parent_window,
            **kwargs
        )

    def _execute_scan_contexts(self, context: ExecutionContext, config_data: dict) -> str:
        self.context = context
        driver_var = config_data.get("driver_var")
        
        self._log(f"Retrieving WebDriver from global variable '@{driver_var}'...")
        driver = context.get_variable(driver_var)

        if not driver or not hasattr(driver, "execute_script"):
            raise TypeError(f"Global variable '@{driver_var}' does not contain a active Selenium WebDriver instance.")

        found_paths = []
        try:
            # Ensure we scan from the root document
            driver.switch_to.default_content()
            self._log("Scanning active browser page for iFrames and Shadow DOMs...")
            
            raw_results = driver.execute_script(JS_LIST_CONTEXTS)
            
            # Format IFrames
            for item in raw_results.get("iframes", []):
                identifier = item.get("id") or item.get("name") or "unknown"
                xpath = item.get("xpath", "")
                found_paths.append(f"IFRAME:#{identifier}({xpath})")

            # Format Shadow Hosts
            for item in raw_results.get("shadow_hosts", []):
                tag = item.get("host_tag", "")
                host_id = item.get("host_id") or "unknown"
                xpath = item.get("host_xpath", "")
                found_paths.append(f"SHADOW:<{tag}>#{host_id}({xpath})")

            result_str = ", ".join(found_paths)
            self._log(f"Scan complete. Found {len(found_paths)} context(s): {result_str}")
            return result_str

        except Exception as e:
            self._log(f"Failed to scan page contexts: {e}", "ERROR")
            raise e
        finally:
            try:
                driver.switch_to.default_content()
            except Exception:
                pass