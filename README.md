How to Install & Run
Install Dependencies:
Open your terminal/command prompt in the folder where you saved the files and run:
bash
1
(Note: Ensure you have Chrome or Firefox installed on your machine for the headless browser to work).
Configure Burp Suite:
Open Burp Suite.
Go to Extensions > Extender > Options.
Under Python Environment, ensure Jython is loaded (download jython-installer.jar from jython.org if not already set).
Go to the Extensions tab > Add.
Select Extension Type: Python.
Select Extension File: Choose the QwenBurpExtension.py you just saved.
Click Next.
Use the Extension:
A new tab named Qwen AI Pentest will appear.
Enter your Alibaba Cloud DashScope API Key.
Select your preferred Model and Browser.
Click Test Headless Browser to ensure it launches correctly.
Go to the Proxy or Repeater tab, right-click a request, and choose:
Send to Qwen AI: Standard text analysis.
Send to Qwen AI (with Browser Verify): Launches a headless browser to visit the URL and checks for client-side execution before sending data to the LLM.
