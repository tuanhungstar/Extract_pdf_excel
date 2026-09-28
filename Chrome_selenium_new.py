import sys
import time
import re
from typing import Optional, List, Dict, Any

# --- External Library Imports ---
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException

# --- PyQt6 Imports ---
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout, QLineEdit, QPushButton, QDialogButtonBox,
    QComboBox, QWidget, QGroupBox, QMessageBox, QLabel, QRadioButton, QHBoxLayout,
    QFileDialog, QStackedWidget
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

# --- Javascript Engines (from chrome_driver_tester) ---
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
        try { if (iframe.contentDocument) findElements(iframe.contentDocument, fullPath); } 
        catch (e) { info.error = 'Cross-origin iframe'; }
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

JS_SMART_SEARCH_ENGINE = """
const searchText = (arguments[0] || '').toLowerCase();
const mode = arguments[1] || 'smart';
const searchContext = arguments.length > 2 ? arguments[2] : document;
if (!searchText) return [];
function findElements(contextNode) {
    let matches = [];
    const checkNode = (el) => {
        try {
            if (el.nodeType !== 1) return false;
            const style = window.getComputedStyle(el);
            if (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') return false;
            const text = (el.innerText || el.textContent || el.value || el.placeholder || '').trim().toLowerCase();
            if (mode === 'exact') return text === searchText;
            return text.includes(searchText);
        } catch (e) { return false; }
    };
    const nodes = contextNode.querySelectorAll('*');
    for (const node of nodes) {
        if (checkNode(node)) matches.push(node);
        if (node.shadowRoot) matches = matches.concat(findElements(node.shadowRoot));
    }
    if (contextNode === document) {
        const iframes = contextNode.querySelectorAll('iframe');
        for (const iframe of iframes) {
            try { if (iframe.contentDocument) matches = matches.concat(findElements(iframe.contentDocument)); } catch (e) {}
        }
    }
    return matches;
}
let allMatches = findElements(searchContext);
if (mode === 'smart' && allMatches.length > 0) {
    const exactMatches = allMatches.filter(el => (el.innerText || el.textContent || el.value || '').trim().toLowerCase() === searchText);
    if (exactMatches.length > 0) allMatches = exactMatches;
}
return allMatches.filter(el => !allMatches.some(other => el !== other && el.contains(other)));
"""

class _SmartSearchConfigDialog(QDialog):
    def __init__(self, global_variables: List[str], execution_context: Optional[ExecutionContext], 
                 parent: Optional[QWidget] = None, initial_config: Optional[Dict[str, Any]] = None,
                 initial_variable: Optional[str] = None):
        super().__init__(parent)
        self.setWindowTitle("Configure Smart Search Web Automation")
        self.setMinimumWidth(650)
        self.global_variables = global_variables
        self.execution_context = execution_context

        main_layout = QVBoxLayout(self)

        # 1. WebDriver Configuration
        driver_group = QGroupBox("WebDriver & Target")
        driver_layout = QVBoxLayout(driver_group)
        
        self.driver_source_layout = QHBoxLayout()
        self.driver_var_radio = QRadioButton("Use Existing WebDriver Variable")
        self.driver_new_radio = QRadioButton("Launch New Browser")
        self.driver_var_radio.setChecked(True)
        self.driver_source_layout.addWidget(self.driver_var_radio)
        self.driver_source_layout.addWidget(self.driver_new_radio)
        driver_layout.addLayout(self.driver_source_layout)

        self.driver_stack = QStackedWidget()
        
        # Page 1: Existing WebDriver
        self.existing_driver_widget = QWidget()
        existing_layout = QHBoxLayout(self.existing_driver_widget)
        existing_layout.setContentsMargins(0, 0, 0, 0)
        self.driver_var_combo = QComboBox()
        self.driver_var_combo.addItems(["-- Select Variable --"] + self.global_variables)
        self.btn_scan_contexts = QPushButton("投 Scan Contexts (Iframes/Shadow)")
        self.btn_scan_contexts.clicked.connect(self._scan_contexts)
        existing_layout.addWidget(QLabel("WebDriver Variable:"))
        existing_layout.addWidget(self.driver_var_combo, 1)
        existing_layout.addWidget(self.btn_scan_contexts)
        self.driver_stack.addWidget(self.existing_driver_widget)
        
        # Page 2: Launch New Browser
        self.new_driver_widget = QWidget()
        new_layout = QFormLayout(self.new_driver_widget)
        new_layout.setContentsMargins(0, 0, 0, 0)
        self.url_edit = QLineEdit("https://")
        self.driver_path_edit = QLineEdit()
        self.driver_path_edit.setPlaceholderText("Leave blank if ChromeDriver is in PATH")
        browse_button = QPushButton("Browse...")
        browse_button.clicked.connect(self._browse_driver)
        driver_path_layout = QHBoxLayout()
        driver_path_layout.setContentsMargins(0,0,0,0)
        driver_path_layout.addWidget(self.driver_path_edit)
        driver_path_layout.addWidget(browse_button)
        new_layout.addRow("Target URL:", self.url_edit)
        new_layout.addRow("ChromeDriver Path:", driver_path_layout)
        self.driver_stack.addWidget(self.new_driver_widget)
        
        driver_layout.addWidget(self.driver_stack)
        main_layout.addWidget(driver_group)
        
        self.driver_var_radio.toggled.connect(self._toggle_driver_source)

        # 2. Context Configuration
        context_group = QGroupBox("Search Context")
        context_layout = QFormLayout(context_group)
        self.context_type_combo = QComboBox()
        # Default options + manual overrides if scan is not used
        self.context_type_combo.addItem("Default Document", {"mode": "default", "xpath": ""})
        self.context_type_combo.addItem("Manual IFrame (by XPath)", {"mode": "manual_iframe", "xpath": ""})
        self.context_type_combo.addItem("Manual Shadow Host (by XPath)", {"mode": "manual_shadow", "xpath": ""})
        self.context_type_combo.addItem("From Global Variable", {"mode": "variable", "xpath": ""})
        
        self.context_xpath_edit = QLineEdit()
        self.context_xpath_edit.setPlaceholderText("e.g., //iframe[@id='main-frame'] (Used for manual entry)")
        self.context_xpath_edit.setEnabled(False)

        self.context_var_combo = QComboBox()
        self.context_var_combo.addItems(["-- Select Variable --"] + self.global_variables)
        self.context_var_combo.setVisible(False)
        
        self.context_type_combo.currentIndexChanged.connect(self._on_context_changed)
        
        context_layout.addRow("Context Selection:", self.context_type_combo)
        
        self.context_input_layout = QHBoxLayout()
        self.context_input_layout.setContentsMargins(0, 0, 0, 0)
        self.context_input_layout.addWidget(self.context_xpath_edit)
        self.context_input_layout.addWidget(self.context_var_combo)
        
        context_layout.addRow("Context Value:", self.context_input_layout)
        main_layout.addWidget(context_group)

        # 3. Search Text Configuration
        search_group = QGroupBox("Smart Search Text")
        search_layout = QVBoxLayout(search_group)
        self.search_source_layout = QHBoxLayout()
        self.text_search_radio = QRadioButton("Hardcoded Text")
        self.var_search_radio = QRadioButton("Use Global Variable")
        self.text_search_radio.setChecked(True)
        self.search_source_layout.addWidget(self.text_search_radio)
        self.search_source_layout.addWidget(self.var_search_radio)
        
        self.search_text_edit = QLineEdit()
        self.search_text_edit.setPlaceholderText("Enter the exact text or partial text to find...")
        self.search_var_combo = QComboBox()
        self.search_var_combo.addItems(["-- Select Variable --"] + self.global_variables)
        self.search_var_combo.setVisible(False)
        self.text_search_radio.toggled.connect(self._toggle_search_input)
        
        search_layout.addLayout(self.search_source_layout)
        search_layout.addWidget(self.search_text_edit)
        search_layout.addWidget(self.search_var_combo)
        main_layout.addWidget(search_group)

        # 4. Action Configuration
        action_group = QGroupBox("Action on Found Element")
        action_layout = QFormLayout(action_group)
        self.action_combo = QComboBox()
        self.action_combo.addItems(["Click Element", "Get Text", "Send Keys (Type)"])
        
        self.send_keys_widget = QWidget()
        send_keys_layout = QVBoxLayout(self.send_keys_widget)
        send_keys_layout.setContentsMargins(0,0,0,0)
        
        self.sk_source_layout = QHBoxLayout()
        self.sk_text_radio = QRadioButton("Hardcoded Value")
        self.sk_var_radio = QRadioButton("Use Global Variable")
        self.sk_text_radio.setChecked(True)
        self.sk_source_layout.addWidget(self.sk_text_radio)
        self.sk_source_layout.addWidget(self.sk_var_radio)
        
        self.sk_text_edit = QLineEdit()
        self.sk_text_edit.setPlaceholderText("Text to type into the element...")
        self.sk_var_combo = QComboBox()
        self.sk_var_combo.addItems(["-- Select Variable --"] + self.global_variables)
        self.sk_var_combo.setVisible(False)
        self.sk_text_radio.toggled.connect(self._toggle_sk_input)
        
        send_keys_layout.addLayout(self.sk_source_layout)
        send_keys_layout.addWidget(self.sk_text_edit)
        send_keys_layout.addWidget(self.sk_var_combo)
        self.send_keys_widget.setVisible(False)
        
        self.action_combo.currentIndexChanged.connect(lambda idx: self.send_keys_widget.setVisible(idx == 2))
        
        action_layout.addRow("Action:", self.action_combo)
        action_layout.addRow("Input Data:", self.send_keys_widget)
        main_layout.addWidget(action_group)

        # 5. Assign Result Group
        assign_group = QGroupBox("Assign Result to Variable")
        assign_layout = QFormLayout(assign_group)
        self.new_var_radio = QRadioButton("New Variable Name:")
        self.new_var_input = QLineEdit("automation_result")
        self.existing_var_radio = QRadioButton("Existing Variable:")
        self.existing_var_combo = QComboBox()
        self.existing_var_combo.addItems(["-- Select --"] + self.global_variables)
        self.new_var_radio.setChecked(True)
        
        assign_layout.addRow(self.new_var_radio, self.new_var_input)
        assign_layout.addRow(self.existing_var_radio, self.existing_var_combo)
        main_layout.addWidget(assign_group)

        # Dialog Buttons
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        main_layout.addWidget(self.button_box)

        if initial_config:
            self._populate_from_initial_config(initial_config, initial_variable)

    def _browse_driver(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select ChromeDriver", "", "Executables (*.exe);;All Files (*)")
        if file_path:
            self.driver_path_edit.setText(file_path)

    def _toggle_driver_source(self):
        if self.driver_var_radio.isChecked():
            self.driver_stack.setCurrentIndex(0)
        else:
            self.driver_stack.setCurrentIndex(1)

    def _on_context_changed(self, idx: int):
        data = self.context_type_combo.itemData(idx)
        mode = data.get("mode") if data else "default"
        
        if mode in ["manual_iframe", "manual_shadow"]:
            self.context_xpath_edit.setVisible(True)
            self.context_var_combo.setVisible(False)
            self.context_xpath_edit.setEnabled(True)
        elif mode == "variable":
            self.context_xpath_edit.setVisible(False)
            self.context_var_combo.setVisible(True)
        else:
            self.context_xpath_edit.setVisible(True)
            self.context_var_combo.setVisible(False)
            self.context_xpath_edit.setEnabled(False)

    def _toggle_search_input(self):
        is_text = self.text_search_radio.isChecked()
        self.search_text_edit.setVisible(is_text)
        self.search_var_combo.setVisible(not is_text)
        
    def _toggle_sk_input(self):
        is_text = self.sk_text_radio.isChecked()
        self.sk_text_edit.setVisible(is_text)
        self.sk_var_combo.setVisible(not is_text)

    def _scan_contexts(self):
        if not self.execution_context:
            QMessageBox.warning(self, "Context Error", "Execution context is not available to retrieve live variables.")
            return

        var_name = self.driver_var_combo.currentText()
        if var_name == "-- Select Variable --":
            QMessageBox.warning(self, "Input Error", "Please select a WebDriver variable first.")
            return

        driver = self.execution_context.get_variable(var_name)
        if not hasattr(driver, 'execute_script'):
            QMessageBox.warning(self, "Type Error", f"Variable '@{var_name}' does not appear to contain an active WebDriver.")
            return

        try:
            driver.switch_to.default_content()
            results = driver.execute_script(JS_LIST_CONTEXTS)
            
            # Reset combo box to base options
            self.context_type_combo.clear()
            self.context_type_combo.addItem("Default Document", {"mode": "default", "xpath": ""})
            self.context_type_combo.addItem("Manual IFrame (by XPath)", {"mode": "manual_iframe", "xpath": ""})
            self.context_type_combo.addItem("Manual Shadow Host (by XPath)", {"mode": "manual_shadow", "xpath": ""})
            self.context_type_combo.addItem("From Global Variable", {"mode": "variable", "xpath": ""})
            
            # Append scanned contexts
            count = 0
            for item in results.get('iframes', []):
                self.context_type_combo.addItem(f"剥 Scanned IFRAME: #{item['id'] or item['name'] or 'unknown'} ({item['xpath']})", 
                                                {"mode": "scanned_iframe", "xpath": item['xpath']})
                count += 1
                
            for item in results.get('shadow_hosts', []):
                self.context_type_combo.addItem(f"剥 Scanned SHADOW: <{item['host_tag']}> #{item['host_id'] or 'unknown'}", 
                                                {"mode": "scanned_shadow", "xpath": item['host_xpath']})
                count += 1
                
            QMessageBox.information(self, "Scan Complete", f"Successfully found {count} contexts on the current page.")
            self.context_type_combo.setCurrentIndex(0)
            
        except Exception as e:
            QMessageBox.critical(self, "Scan Failed", f"An error occurred while scanning the page:\n{str(e)}")

    def _populate_from_initial_config(self, config, variable):
        if config.get("driver_source") == "variable":
            self.driver_var_radio.setChecked(True)
            self.driver_var_combo.setCurrentText(config.get("driver_var", ""))
        else:
            self.driver_new_radio.setChecked(True)
            self.url_edit.setText(config.get("url", ""))
            self.driver_path_edit.setText(config.get("driver_path", ""))
            
        # Re-populate selected context from config
        saved_mode = config.get("context_mode", "default")
        saved_xpath = config.get("context_xpath", "")
        saved_var = config.get("context_var", "")
        
        # Determine if it was a manual mode we can restore easily
        match_idx = -1
        for i in range(self.context_type_combo.count()):
            if self.context_type_combo.itemData(i).get("mode") == saved_mode:
                if saved_mode in ["default", "manual_iframe", "manual_shadow", "variable"]:
                    match_idx = i; break
        
        if match_idx >= 0:
            self.context_type_combo.setCurrentIndex(match_idx)
            if saved_mode in ["manual_iframe", "manual_shadow"]:
                self.context_xpath_edit.setText(saved_xpath)
            elif saved_mode == "variable":
                self.context_var_combo.setCurrentText(saved_var)
        elif saved_mode.startswith("scanned_"):
            # If it was a scanned item, we append it back as a cached option
            label = f"沈 Cached {saved_mode.split('_')[1].upper()} Context"
            self.context_type_combo.addItem(label, {"mode": saved_mode, "xpath": saved_xpath})
            self.context_type_combo.setCurrentIndex(self.context_type_combo.count() - 1)
        
        # Search Text Setup
        if config.get("search_type") == "variable":
            self.var_search_radio.setChecked(True)
            self.search_var_combo.setCurrentText(config.get("search_value", ""))
        else:
            self.text_search_radio.setChecked(True)
            self.search_text_edit.setText(config.get("search_value", ""))
            
        # Action Setup
        self.action_combo.setCurrentIndex(config.get("action_idx", 0))
        if config.get("sk_type") == "variable":
            self.sk_var_radio.setChecked(True)
            self.sk_var_combo.setCurrentText(config.get("sk_value", ""))
        else:
            self.sk_text_radio.setChecked(True)
            self.sk_text_edit.setText(config.get("sk_value", ""))

        if variable:
            if variable in self.global_variables:
                self.existing_var_radio.setChecked(True)
                self.existing_var_combo.setCurrentText(variable)
            else:
                self.new_var_radio.setChecked(True)
                self.new_var_input.setText(variable)

    def get_executor_method_name(self) -> str:
        return "_execute_smart_search"

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
        config = {}
        
        # Driver Config
        if self.driver_var_radio.isChecked():
            config["driver_source"] = "variable"
            config["driver_var"] = self.driver_var_combo.currentText()
            if config["driver_var"] == "-- Select Variable --":
                QMessageBox.warning(self, "Input Error", "Please select a WebDriver variable."); return None
        else:
            config["driver_source"] = "new"
            config["url"] = self.url_edit.text().strip()
            config["driver_path"] = self.driver_path_edit.text().strip()
            if not config["url"]:
                QMessageBox.warning(self, "Input Error", "Target URL is required for a new browser."); return None
                
        # Context Config
        context_data = self.context_type_combo.currentData()
        config["context_mode"] = context_data.get("mode")
        if config["context_mode"] in ["manual_iframe", "manual_shadow"]:
            config["context_xpath"] = self.context_xpath_edit.text().strip()
            if not config["context_xpath"]:
                QMessageBox.warning(self, "Input Error", "XPath is required for manual context entry."); return None
        elif config["context_mode"] == "variable":
            config["context_var"] = self.context_var_combo.currentText()
            if config["context_var"] == "-- Select Variable --":
                QMessageBox.warning(self, "Input Error", "Please select a variable for the context."); return None
        else:
            config["context_xpath"] = context_data.get("xpath")

        # Search Config
        config["search_type"] = "text" if self.text_search_radio.isChecked() else "variable"
        config["search_value"] = self.search_text_edit.text() if config["search_type"] == "text" else self.search_var_combo.currentText()
        if config["search_type"] == "variable" and config["search_value"] == "-- Select Variable --":
            QMessageBox.warning(self, "Input Error", "Please select a variable for the search text."); return None
            
        # Action Config
        config["action_idx"] = self.action_combo.currentIndex()
        config["sk_type"] = "text" if self.sk_text_radio.isChecked() else "variable"
        config["sk_value"] = self.sk_text_edit.text() if config["sk_type"] == "text" else self.sk_var_combo.currentText()
        if config["action_idx"] == 2 and config["sk_type"] == "variable" and config["sk_value"] == "-- Select Variable --":
            QMessageBox.warning(self, "Input Error", "Please select a variable for the send keys text."); return None

        return config


class SmartSearch_API:
    def __init__(self, context: Optional[ExecutionContext] = None):
        self.context = context

    def _log(self, message: str, level: str = "INFO"):
        if self.context:
            self.context.add_log(f"SmartSearch_API [{level}]: {message}")
        else:
            print(f"SmartSearch_API [{level}]: {message}")

    def configure_data_hub(self, parent_window: QWidget, global_variables: List[str], **kwargs) -> QDialog:
        self._log("Opening Smart Search API configuration dialog.")
        return _SmartSearchConfigDialog(
            global_variables=global_variables,
            execution_context=self.context, # Pass live context to allow scanning
            parent=parent_window,
            **kwargs
        )

    def _execute_smart_search(self, context: ExecutionContext, config_data: dict) -> Any:
        self.context = context
        driver = None
        owns_driver = False
        result = "Failed"

        try:
            # Resolve Dynamic Variables
            search_text = config_data['search_value']
            if config_data['search_type'] == 'variable':
                search_text = context.get_variable(search_text)
                
            sk_text = config_data['sk_value']
            if config_data['action_idx'] == 2 and config_data['sk_type'] == 'variable':
                sk_text = context.get_variable(sk_text)

            # Initialize / Retrieve WebDriver
            if config_data.get('driver_source') == 'variable':
                var_name = config_data['driver_var']
                self._log(f"Retrieving WebDriver from variable '@{var_name}'...")
                driver = context.get_variable(var_name)
                owns_driver = False
                if not hasattr(driver, 'execute_script'):
                    raise TypeError(f"Variable '@{var_name}' does not contain a valid WebDriver instance.")
            else:
                owns_driver = True
                url = config_data.get('url')
                driver_path = config_data.get('driver_path')
                service = ChromeService(executable_path=driver_path) if driver_path else ChromeService()
                
                self._log("Launching new Chrome Driver...")
                driver = webdriver.Chrome(service=service)
                self._log(f"Navigating to {url}")
                driver.get(url)
                time.sleep(2) # Brief pause to allow initial render

            # Switch to Target Context
            driver.switch_to.default_content()
            context_mode = config_data.get('context_mode', 'default')
            context_xpath = config_data.get('context_xpath', '')
            search_args = [search_text, 'smart']

            # --- Extract XPath from Variable ---
            if context_mode == 'variable':
                context_var = config_data.get('context_var')
                raw_context = str(context.get_variable(context_var) or "").strip()
                
                # Extract XPath inside the parentheses e.g. IFRAME:#unknown(...)
                match = re.search(r'\(([^)]+)\)$', raw_context)
                
                if match:
                    context_xpath = match.group(1)
                    # Automatically determine if it's a Shadow Host or IFrame from the prefix
                    context_mode = 'manual_shadow' if 'SHADOW' in raw_context.upper() else 'manual_iframe'
                else:
                    # Fallback: Assume the string is a raw XPath if no parentheses formatting is found
                    context_xpath = raw_context
                    context_mode = 'manual_iframe'

            if context_mode in ['manual_iframe', 'scanned_iframe']:
                self._log(f"Switching to IFrame via XPath: {context_xpath}")
                frame_el = driver.find_element(By.XPATH, context_xpath)
                driver.switch_to.frame(frame_el)
            elif context_mode in ['manual_shadow', 'scanned_shadow']:
                self._log(f"Accessing Shadow Host via XPath: {context_xpath}")
                host_el = driver.find_element(By.XPATH, context_xpath)
                search_args.append(host_el.shadow_root)

            # Execute Smart Search
            self._log(f"Executing Smart Search for text: '{search_text}'")
            elements = driver.execute_script(JS_SMART_SEARCH_ENGINE, *search_args)

            if not elements:
                self._log("No elements found matching the text.", "WARN")
                return "Element Not Found"

            target_el = elements[0]
            self._log(f"Found element: <{target_el.tag_name}>")

            # Perform Configured Action
            action_idx = config_data.get('action_idx', 0)
            
            if action_idx == 0: # Click
                try:
                    target_el.click()
                    self._log("Native click was successful.")
                except Exception as e:
                    self._log(f"Native click failed: {e}. Trying JS fallback...", "WARN")
                    driver.execute_script("arguments[0].click();", target_el)
                result = "Clicked Successfully"
                
            elif action_idx == 1: # Get Text
                ext_text = target_el.text or target_el.get_attribute('value') or ""
                self._log(f"Extracted Text: {ext_text}")
                result = ext_text
                
            elif action_idx == 2: # Send Keys
                target_el.send_keys(str(sk_text))
                self._log(f"Sent keys: '{sk_text}'")
                result = f"Sent Keys: {sk_text}"

        except Exception as e:
            self._log(f"Execution Error: {e}", "ERROR")
            result = f"Error: {str(e)}"
        
        finally:
            if driver:
                # Ensure we reset to default context for subsequent commands
                try: driver.switch_to.default_content() 
                except: pass
                
                # Only quit the browser if this module launched it
                if owns_driver:
                    self._log("Closing new browser session.")
                    try: driver.quit()
                    except: pass

        return result