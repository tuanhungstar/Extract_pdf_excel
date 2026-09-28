import sys
import time
from typing import Optional, List, Dict, Any

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QComboBox,
    QTextEdit,
    QFileDialog,
    QGroupBox,
    QSplitter,
)

from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By

# --- JAVASCRIPT LIBRARIES ---

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
    // Don't search iframes if we are already in a specific context
    if (contextNode === document) {
        const iframes = contextNode.querySelectorAll('iframe');
        for (const iframe of iframes) {
            try {
                if (iframe.contentDocument) matches = matches.concat(findElements(iframe.contentDocument));
            } catch (e) {}
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

JS_HIGHLIGHT_ELEMENT = "arguments[0].style.outline = `3px solid ${arguments[1] || 'red'}`;"
JS_CLEAR_HIGHLIGHT = "if(arguments[0]) arguments[0].style.outline = '';"
JS_GET_XPATH = "let e=arguments[0],p;for(p=[];e&&e.nodeType==1;e=e.parentNode){let i=1,s=e.previousSibling;for(;s;s=s.previousSibling)1==s.nodeType&&s.tagName==e.tagName&&i++;p.unshift(e.tagName.toLowerCase()+(i>1?`[${i}]`:''))}return p.join('/')"


class ChromeDriverTesterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ChromeDriver Interactive Tester v8 (Context Switching)")
        self.resize(1200, 900)
        self.driver: Optional[webdriver.Chrome] = None
        self.found_elements_cache: List[WebElement] = []
        self.context_cache: List[Dict[str, Any]] = []
        self.highlighted_element: Optional[WebElement] = None
        self._init_ui()

    def _init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # --- Group 1: Config ---
        config_group = QGroupBox("1. ChromeDriver & URL Setup")
        config_layout = QGridLayout(config_group)
        config_layout.addWidget(QLabel("ChromeDriver Path:"), 0, 0)
        self.driver_path_input = QLineEdit()
        config_layout.addWidget(self.driver_path_input, 0, 1)
        self.btn_browse = QPushButton("Browse...")
        config_layout.addWidget(self.btn_browse, 0, 2)
        config_layout.addWidget(QLabel("Target URL:"), 1, 0)
        self.url_input = QLineEdit("https://www.google.com")
        config_layout.addWidget(self.url_input, 1, 1)
        btn_box = QHBoxLayout()
        self.btn_open = QPushButton("Open Browser / Navigate")
        btn_box.addWidget(self.btn_open)
        self.btn_close = QPushButton("Close Browser")
        btn_box.addWidget(self.btn_close)
        config_layout.addLayout(btn_box, 1, 2)
        main_layout.addWidget(config_group)

        splitter = QSplitter(Qt.Orientation.Vertical)
        
        top_pane_widget = QWidget()
        top_pane_layout = QVBoxLayout(top_pane_widget)

        # --- Group 2: Context ---
        context_group = QGroupBox("2. Set Page Context for Searching")
        context_layout = QGridLayout(context_group)
        self.btn_scan_contexts = QPushButton("📊 Scan for Contexts (Iframes & Shadow DOMs)")
        context_layout.addWidget(self.btn_scan_contexts, 0, 0, 1, 2)
        context_layout.addWidget(QLabel("Search In:"), 1, 0)
        self.context_combo = QComboBox()
        context_layout.addWidget(self.context_combo, 1, 1)
        top_pane_layout.addWidget(context_group)

        # --- Group 3: Find ---
        locator_group = QGroupBox("3. Find Elements (within selected context)")
        locator_layout = QGridLayout(locator_group)
        locator_layout.addWidget(QLabel("Locator By:"), 0, 0)
        self.combo_by = QComboBox()
        self.combo_by.addItems(["Smart Search (Text)", "XPATH", "CSS_SELECTOR", "ID", "NAME"])
        locator_layout.addWidget(self.combo_by, 0, 1)
        locator_layout.addWidget(QLabel("Selector / Text:"), 1, 0)
        self.selector_input = QLineEdit()
        locator_layout.addWidget(self.selector_input, 1, 1)
        self.btn_find = QPushButton("🔍 Find Matching Elements")
        locator_layout.addWidget(self.btn_find, 2, 0, 1, 2)
        top_pane_layout.addWidget(locator_group)

        # --- Group 4: Inspect ---
        inspector_group = QGroupBox("4. Inspect & Act on Selected Element")
        inspector_layout = QGridLayout(inspector_group)
        inspector_layout.addWidget(QLabel("Found Elements:"), 0, 0)
        self.results_combo = QComboBox()
        inspector_layout.addWidget(self.results_combo, 0, 1)
        inspector_layout.addWidget(QLabel("Selected's XPath:"), 1, 0)
        self.selected_locator_input = QLineEdit()
        self.selected_locator_input.setReadOnly(True)
        inspector_layout.addWidget(self.selected_locator_input, 1, 1)
        action_btn_layout = QHBoxLayout()
        
        self.btn_click = QPushButton("👆 Click")
        action_btn_layout.addWidget(self.btn_click)
        
        self.btn_get_text = QPushButton("📝 Get Text")
        action_btn_layout.addWidget(self.btn_get_text)
        
        self.btn_send_keys = QPushButton("⌨️ Send Keys")
        action_btn_layout.addWidget(self.btn_send_keys)
        
        self.value_input = QLineEdit()
        self.value_input.setPlaceholderText("Text for Send Keys")
        action_btn_layout.addWidget(self.value_input)
        
        inspector_layout.addLayout(action_btn_layout, 2, 0, 1, 2)
        top_pane_layout.addWidget(inspector_group)
        
        splitter.addWidget(top_pane_widget)

        # --- Group 5: Log ---
        console_group = QGroupBox("5. Activity Log")
        console_layout = QVBoxLayout(console_group)
        self.log_console = QTextEdit()
        self.log_console.setReadOnly(True)
        console_layout.addWidget(self.log_console)
        splitter.addWidget(console_group)

        main_layout.addWidget(splitter)
        splitter.setSizes([450, 450])

        self._reset_context_combo()

        # --- Connect Signals ---
        self.btn_browse.clicked.connect(self._browse_chromedriver)
        self.btn_open.clicked.connect(self._launch_or_navigate)
        self.btn_close.clicked.connect(self._close_browser)
        self.btn_scan_contexts.clicked.connect(self._action_scan_contexts)
        self.btn_find.clicked.connect(self._action_find_elements)
        self.results_combo.activated.connect(self._on_element_selected)
        self.btn_click.clicked.connect(self._action_click)
        self.btn_get_text.clicked.connect(self._action_get_text)
        self.btn_send_keys.clicked.connect(self._action_send_keys)

    def _reset_context_combo(self):
        self.context_combo.clear()
        self.context_cache.clear()
        self.context_combo.addItem("Default Page Content", -1)

    def log(self, text: str, level: str = "INFO"):
        stamp = time.strftime("%H:%M:%S")
        color = {"ERROR": "red", "SUCCESS": "green", "WARN": "orange", "HEADING": "blue"}.get(level, "black")
        self.log_console.append(f'<span style="color: gray;">[{stamp}]</span> <b style="color: {color};">[{level}]</b> {text}')

    def _browse_chromedriver(self):
        fname, _ = QFileDialog.getOpenFileName(self, "Select chromedriver.exe", "", "Executables (*.exe)")
        if fname: self.driver_path_input.setText(fname)

    def _launch_or_navigate(self):
        url = self.url_input.text().strip()
        if not url: return self.log("URL cannot be empty.", "WARN")
        try:
            if not self.driver:
                path = self.driver_path_input.text().strip()
                service = ChromeService(executable_path=path) if path else ChromeService()
                self.driver = webdriver.Chrome(service=service)
            self.driver.get(url)
            self.log(f"Loaded: {self.driver.title}", "SUCCESS")
            self._reset_context_combo()
        except WebDriverException as e:
            self.log(f"Launch/Navigate failed: {e.msg}", "ERROR")

    def _close_browser(self):
        self._clear_highlight()
        if self.driver:
            try: self.driver.quit()
            except Exception: pass
            self.driver = None
        self.found_elements_cache = []
        self.results_combo.clear()
        self.selected_locator_input.clear()
        self._reset_context_combo()
        self.log("Browser closed.", "INFO")

    def _action_scan_contexts(self):
        if not self.driver: return self.log("Browser not open.", "ERROR")
        self.log("Scanning page for iframes and Shadow DOMs...", "INFO")
        try:
            self.driver.switch_to.default_content()
            results = self.driver.execute_script(JS_LIST_CONTEXTS)
            self._reset_context_combo()
            
            for item in results.get('iframes', []):
                item['type'] = 'iframe'
                self.context_cache.append(item)
                display_text = f"IFRAME: #{item['id'] or item['name'] or 'unknown'}"
                self.context_combo.addItem(display_text, len(self.context_cache) - 1)

            for item in results.get('shadow_hosts', []):
                item['type'] = 'shadow_host'
                self.context_cache.append(item)
                display_text = f"SHADOW: <{item['host_tag']}> #{item['host_id'] or 'unknown'}"
                self.context_combo.addItem(display_text, len(self.context_cache) - 1)
            
            self.log(f"Scan complete. Found {len(self.context_cache)} contexts.", "SUCCESS")
        except Exception as e: self.log(f"Failed to scan page: {e}", "ERROR")

    def _action_find_elements(self):
        if not self.driver: return self.log("Browser not open.", "ERROR")
        
        strategy = self.combo_by.currentText()
        selector = self.selector_input.text().strip()
        if not selector: return self.log("Selector/Text cannot be empty.", "WARN")

        self.results_combo.clear(); self.found_elements_cache = []
        self._clear_highlight()

        context_idx = self.context_combo.currentData()
        context_info = self.context_cache[context_idx] if context_idx != -1 else {"type": "default"}
        
        self.log(f"Finding elements by {strategy} in context: {self.context_combo.currentText()}", "HEADING")
        
        try:
            self.driver.switch_to.default_content()
            elements = []
            search_args = [selector, 'smart' if strategy == "Smart Search (Text)" else 'contains']

            if context_info['type'] == 'iframe':
                frame_element = self.driver.find_element(By.XPATH, context_info['xpath'])
                self.driver.switch_to.frame(frame_element)
                if strategy == "Smart Search (Text)":
                    elements = self.driver.execute_script(JS_SMART_SEARCH_ENGINE, *search_args)
                else:
                    elements = self.driver.find_elements(getattr(By, strategy), selector)

            elif context_info['type'] == 'shadow_host':
                host_element = self.driver.find_element(By.XPATH, context_info['host_xpath'])
                search_args.append(host_element.shadow_root)
                if strategy == "Smart Search (Text)":
                    elements = self.driver.execute_script(JS_SMART_SEARCH_ENGINE, *search_args)
                else: # Fallback for non-JS search in shadow root
                    elements = host_element.shadow_root.find_elements(getattr(By, strategy.replace("_", " ")), selector)

            else: # Default content
                if strategy == "Smart Search (Text)":
                    elements = self.driver.execute_script(JS_SMART_SEARCH_ENGINE, *search_args)
                else:
                    elements = self.driver.find_elements(getattr(By, strategy), selector)

            if not elements: return self.log("No elements found in this context.", "WARN")

            self.found_elements_cache = elements
            self.log(f"Found {len(elements)} element(s). Populating dropdown.", "SUCCESS")
            for i, el in enumerate(elements):
                try:
                    text = (el.text or el.get_attribute('value') or f"<{el.tag_name}>").strip()[:60]
                    self.results_combo.addItem(f"{i}: {text}", i)
                except Exception: self.results_combo.addItem(f"{i}: [Stale Element]", i)
            self.results_combo.setCurrentIndex(0)
            self._on_element_selected()

        except Exception as e: self.log(f"Find failed: {e}", "ERROR")
        finally:
            if self.driver: self.driver.switch_to.default_content()

    def _switch_to_correct_context(self) -> None:
        """
        Switches the driver to the context selected in the dropdown before performing an action.
        """
        if not self.driver:
            return
            
        self.driver.switch_to.default_content() # Start from a clean slate
        context_idx = self.context_combo.currentData()
        if context_idx is None or context_idx == -1:
            return # In default content, do nothing

        context_info = self.context_cache[context_idx]
        if context_info['type'] == 'iframe':
            try:
                frame_element = self.driver.find_element(By.XPATH, context_info['xpath'])
                self.driver.switch_to.frame(frame_element)
                self.log(f"Switched to context: {self.context_combo.currentText()}", "INFO")
            except Exception as e:
                self.log(f"Could not switch back to iframe context: {e}", "ERROR")

    def _on_element_selected(self):
        if not self.driver or not self.found_elements_cache: return
        index = self.results_combo.currentData()
        if index is None: return

        element = self.found_elements_cache[index]
        self._clear_highlight() 

        try:
            self._switch_to_correct_context()
            
            self.driver.execute_script(JS_HIGHLIGHT_ELEMENT, element, 'blue')
            self.highlighted_element = element
            
            is_shadow = self.driver.execute_script("return arguments[0].getRootNode() instanceof ShadowRoot", element)
            prefix = "(in Shadow DOM) > " if is_shadow else ""
            xpath = self.driver.execute_script(JS_GET_XPATH, element)
            self.selected_locator_input.setText(prefix + xpath)

        except Exception as e:
            self.log(f"Failed to inspect element: {e}", "WARN")
        finally:
            if self.driver:
                self.driver.switch_to.default_content()

    def _get_selected_element(self) -> Optional[WebElement]:
        if not self.highlighted_element:
            self.log("No element selected.", "WARN")
            return None
        return self.highlighted_element

    def _action_click(self):
        el = self._get_selected_element()
        if el:
            try:
                self._switch_to_correct_context()
                self.log("Attempting native Selenium click...", "INFO")
                el.click()
                self.log("Native click was successful.", "SUCCESS")
                
            except Exception as e:
                self.log(f"Native click failed: {type(e).__name__}. Trying JavaScript click as a fallback.", "WARN")
                try:
                    self.driver.execute_script("arguments[0].click();", el)
                    self.log("JavaScript fallback click was successful.", "SUCCESS")
                except Exception as e2:
                    self.log(f"JavaScript fallback click also failed: {e2}", "ERROR")
            finally:
                if self.driver:
                    self.driver.switch_to.default_content()

    def _action_get_text(self):
        el = self._get_selected_element()
        if el:
            try:
                # Switch to the context right before the action
                self._switch_to_correct_context()
                
                # Try getting the visible text first
                extracted_text = el.text
                
                # If text is empty (common for input/textarea fields), try retrieving the value
                if not extracted_text:
                    extracted_text = el.get_attribute('value') or el.get_attribute('textContent') or ""
                
                # Log the output to the activity log console
                self.log(f"Extracted Text: '{extracted_text.strip()}'", "SUCCESS")
                
            except Exception as e:
                self.log(f"Get Text failed: {e}", "ERROR")
            finally:
                # Always switch back after the action is complete
                if self.driver:
                    self.driver.switch_to.default_content()

    def _action_send_keys(self):
        el = self._get_selected_element()
        text_to_send = self.value_input.text()
        if el and text_to_send:
            try:
                self._switch_to_correct_context()
                el.send_keys(text_to_send)
                self.log(f"Sent keys: '{text_to_send}'", "SUCCESS")
            except Exception as e:
                self.log(f"Send Keys failed: {e}", "ERROR")
            finally:
                if self.driver:
                    self.driver.switch_to.default_content()

    def _clear_highlight(self):
        if self.highlighted_element and self.driver:
            try: 
                self.driver.switch_to.default_content()
                self.driver.execute_script(JS_CLEAR_HIGHLIGHT, self.highlighted_element)
            except Exception: 
                pass 
        self.highlighted_element = None
        
    def closeEvent(self, event):
        self._close_browser()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ChromeDriverTesterApp()
    window.show()
    sys.exit(app.exec())