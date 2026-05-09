from burp import IBurpExtender, ITab, IMessageEditorController, IContextMenuFactory, IContextMenuHandler
from javax.swing import JPanel, JTextField, JButton, JLabel, JTextArea, JScrollPane, JComboBox, BoxLayout, SwingUtilities, JMenuItem
from java.awt import BorderLayout, Dimension
from java.awt.event import ActionListener, ItemListener
import requests
import json
import re
import threading
from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
import time

class QwenBurpExtender(IBurpExtender, ITab, IContextMenuFactory, IContextMenuHandler):
    def registerExtenderCallbacks(self, callbacks):
        self._callbacks = callbacks
        self._helpers = callbacks.getHelpers()
        callbacks.setExtensionName("Qwen AI Pentest")
        
        # Configuration
        self.api_key = ""
        self.model_name = "qwen-max"
        self.browser_type = "chrome"
        self.driver = None
        
        # UI Components
        self._splitPane = None
        self._log_output = None
        self._request_viewer = None
        self._response_viewer = None
        
        # Initialize UI
        SwingUtilities.invokeLater(self.init_ui)
        
        # Register Context Menu
        callbacks.registerContextMenuFactory(self)
        
        self._callbacks.printOutput("Qwen AI Pentest Extension loaded successfully.")

    def init_ui(self):
        # Main Panel
        main_panel = JPanel(BorderLayout())
        
        # Top Config Panel
        config_panel = JPanel()
        config_panel.setLayout(BoxLayout(config_panel, BoxLayout.Y_AXIS))
        
        # API Key
        key_panel = JPanel()
        key_panel.add(JLabel("Qwen API Key:"))
        self._api_key_field = JTextField(30)
        key_panel.add(self._api_key_field)
        
        # Model Selection
        model_panel = JPanel()
        model_panel.add(JLabel("Model:"))
        models = ["qwen-max", "qwen-plus", "qwen-turbo", "qwen-long"]
        self._model_combo = JComboBox(models)
        self._model_combo.addItemListener(lambda e: self.update_model(e.getItem()))
        model_panel.add(self._model_combo)
        
        # Browser Selection
        browser_panel = JPanel()
        browser_panel.add(JLabel("Headless Browser:"))
        browsers = ["chrome", "firefox"]
        self._browser_combo = JComboBox(browsers)
        self._browser_combo.addItemListener(lambda e: self.update_browser(e.getItem()))
        browser_panel.add(self._browser_combo)
        
        # Test Browser Button
        self._test_browser_btn = JButton("Test Headless Browser", actionPerformed=self.test_browser_action)
        browser_panel.add(self._test_browser_btn)
        
        config_panel.add(key_panel)
        config_panel.add(model_panel)
        config_panel.add(browser_panel)
        
        # Log Output
        self._log_output = JTextArea(20, 50)
        self._log_output.setEditable(False)
        log_scroll = JScrollPane(self._log_output)
        
        # Request/Response Viewers (Placeholders)
        self._request_viewer = self._callbacks.createMessageEditor(None, False)
        self._response_viewer = self._callbacks.createMessageEditor(None, False)
        
        # Split Pane for Log and Viewers
        split_log_view = self._callbacks.createSplitPane(True, log_scroll, self._request_viewer.getComponent())
        
        main_panel.add(config_panel, BorderLayout.NORTH)
        main_panel.add(split_log_view, BorderLayout.CENTER)
        
        self._callbacks.customizeUiComponent(main_panel)
        self._callbacks.addSuiteTab(self)
        self._splitPane = main_panel

    def getTabCaption(self):
        return "Qwen AI Pentest"

    def getUiComponent(self):
        return self._splitPane

    def update_model(self, model):
        self.model_name = model
        self.log("Model updated to: " + model)

    def update_browser(self, browser):
        self.browser_type = browser
        self.log("Browser type updated to: " + browser)
        if self.driver:
            self.driver.quit()
            self.driver = None

    def test_browser_action(self, event):
        def run_test():
            try:
                self.log("Initializing headless browser test...")
                driver = self._init_driver()
                driver.get("data:text/html,<h1>Browser Test Successful</h1><script>console.log('XSS Test Payload Executed');</script>")
                time.sleep(2)
                logs = driver.get_log('browser')
                if any('XSS Test Payload Executed' in str(log) for log in logs):
                    self.log("SUCCESS: Browser launched and executed JavaScript correctly.")
                else:
                    self.log("WARNING: Browser launched but JS execution log not found.")
                driver.quit()
            except Exception as e:
                self.log("ERROR: Failed to launch browser: " + str(e))
        
        threading.Thread(target=run_test).start()

    def _init_driver(self):
        if self.driver:
            return self.driver
            
        if self.browser_type == "chrome":
            options = ChromeOptions()
            options.add_argument("--headless=new")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-gpu")
            # Suppress logs unless needed for debugging
            options.set_capability('goog:loggingPrefs', {'browser': 'ALL'})
            try:
                from selenium.webdriver.chrome.service import Service
                self.driver = webdriver.Chrome(options=options)
            except Exception:
                # Fallback if webdriver-manager isn't installed or path issues
                self.driver = webdriver.Chrome(options=options)
                
        elif self.browser_type == "firefox":
            options = FirefoxOptions()
            options.add_argument("--headless")
            options.set_preference('browser.console.stdoutEnabled', True)
            try:
                from selenium.webdriver.firefox.service import Service
                self.driver = webdriver.Firefox(options=options)
            except Exception:
                self.driver = webdriver.Firefox(options=options)
        
        return self.driver

    def createMenuItems(self, invocation):
        menu_items = []
        menu_items.append(JMenuItem("Send to Qwen AI", actionPerformed=lambda e: self.send_to_qwen(invocation)))
        menu_items.append(JMenuItem("Send to Qwen AI (with Browser Verify)", actionPerformed=lambda e: self.send_to_qwen(invocation, use_browser=True)))
        return menu_items

    def send_to_qwen(self, invocation, use_browser=False):
        def process_request():
            selected_messages = invocation.getSelectedMessages()
            if not selected_messages:
                return
            
            message = selected_messages[0]
            request_info = self._helpers.analyzeRequest(message)
            
            # Reconstruct Request
            request_bytes = message.getRequest()
            request_str = self._helpers.bytesToString(request_bytes)
            
            # Get Response if available
            response_str = ""
            if message.getResponse():
                response_bytes = message.getResponse()
                response_str = self._helpers.bytesToString(response_bytes)
            
            url = str(request_info.getUrl())
            
            analysis_prompt = f"""
You are an expert Penetration Tester. Analyze the following HTTP transaction for security vulnerabilities.
URL: {url}

REQUEST:
{request_str}

RESPONSE:
{response_str}

Identify potential vulnerabilities (SQLi, XSS, CSRF, IDOR, Misconfigurations, etc.).
Provide a severity rating (Critical, High, Medium, Low) and specific remediation steps.
"""

            # Browser Verification Logic for XSS
            if use_browser:
                self.log(f"Launching headless browser to verify XSS for: {url}")
                browser_findings = self._verify_with_browser(url, request_info.getMethod())
                if browser_findings:
                    analysis_prompt += f"\n\n[BROWSER VERIFICATION RESULTS]:\n{browser_findings}\n\nBased on the browser execution results above, confirm if XSS is exploitable."
                else:
                    analysis_prompt += "\n\n[BROWSER VERIFICATION]: No obvious XSS payload execution detected or browser test skipped for non-GET/simple requests."

            # Call Qwen API
            api_key = self._api_key_field.getText()
            if not api_key:
                self.log("ERROR: API Key missing!")
                return

            self.log("Sending request to Qwen AI...")
            try:
                headers = {
                    "Authorization": f"Bearer {api_key}",
                    "Content-Type": "application/json"
                }
                payload = {
                    "model": self.model_name,
                    "messages": [
                        {"role": "system", "content": "You are a helpful security assistant integrated into Burp Suite."},
                        {"role": "user", "content": analysis_prompt}
                    ]
                }
                
                response = requests.post("https://dashscope.aliyuncs.com/api/v1/services/aigc/text-generation/generation", 
                                         headers=headers, json=payload, timeout=60)
                
                if response.status_code == 200:
                    result = response.json()
                    content = result['output']['text']
                    self.log("\n--- Qwen AI Analysis ---\n" + content + "\n------------------------")
                    
                    # Update UI viewers
                    SwingUtilities.invokeLater(lambda: self._request_viewer.setMessage(message.getRequest(), True))
                    if message.getResponse():
                        SwingUtilities.invokeLater(lambda: self._response_viewer.setMessage(message.getResponse(), False))
                else:
                    self.log(f"API Error: {response.status_code} - {response.text}")
                    
            except Exception as e:
                self.log("Error calling Qwen API: " + str(e))

        threading.Thread(target=process_request).start()

    def _verify_with_browser(self, url, method):
        if method != "GET":
            # For simplicity, this demo only auto-visits GET requests. 
            # POST would require form reconstruction which is complex for a snippet.
            return None

        try:
            driver = self._init_driver()
            driver.get(url)
            time.sleep(2) # Wait for DOM load
            
            findings = []
            
            # Check for standard XSS alerts (if we injected one, but here we just observe)
            # In a real scenario, you'd append ?param=<script>alert(1)</script> to the URL
            # Here we just check console logs for errors that might indicate broken JS or existing XSS
            
            logs = driver.get_log('browser')
            for log in logs:
                msg = str(log)
                if "Uncaught" in msg or "SyntaxError" in msg:
                    findings.append(f"Console Error: {msg}")
            
            # Simple heuristic: Check if the page title contains script tags (reflected)
            title = driver.title
            if "<" in title or ">" in title:
                findings.append(f"Suspicious Title Reflection: {title}")
                
            driver.quit()
            return "\n".join(findings) if findings else "No immediate anomalies detected in console."
            
        except Exception as e:
            return f"Browser verification failed: {str(e)}"

    def log(self, message):
        self._log_output.append(message + "\n")
        self._log_output.setCaretPosition(self._log_output.getDocument().getLength())