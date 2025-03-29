import sys
from PyQt5.QtCore import Qt

import os
import time
import pandas as pd

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget,
    QFileDialog, QMessageBox, QLabel, QLineEdit
)

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from extractor import HoldingsProcessor


class HoldingsExtractor(QMainWindow):
    def __init__(self):
        super().__init__()

        self.init_ui()
        self.driver = None
        self.file_path = None
        self.client_codes = []
        self.successful_downloads = 0
        self.max_parallel_clients = 3  

        self.failed_clients = [] 
        
        if getattr(sys, 'frozen', False):  
            base_path = os.path.dirname(sys.executable)  
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))  

        self.download_folder = os.path.join(base_path, "Holding")

    
        
        if not os.path.exists(self.download_folder):
            os.makedirs(self.download_folder)  

    def log(self, message):
        print(message)
        self.status_label.setText(message)  


    def init_ui(self):
        self.setWindowTitle("PORTFOLIO REVIEW APP")
        self.setGeometry(100, 100, 500, 450)
        
        self.setStyleSheet("""
            background-color: #f8f9fa;  /* Light grey background */
            font-family: Arial;
            font-size: 12px;
        """)

        layout = QVBoxLayout()

        self.title_label = QLabel("PORTFOLIO REVIEW APP")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)  
        self.title_label.setStyleSheet("""
            font-size: 20px;
            font-weight: bold;
            color: #1e3a8a;  /* Deep blue */
            padding: 10px;
        """)
        layout.addWidget(self.title_label)

        self.url_label = QLabel("🔗 Enter URL:")
        self.url_label.setStyleSheet("font-weight: bold; color: #2c3e50;")
        self.url_input = QLineEdit()
        self.url_input.setPlaceholderText("https://example.com")
        layout.addWidget(self.url_label)
        layout.addWidget(self.url_input)

        self.username_label = QLabel("👤 Enter Username:")
        self.username_label.setStyleSheet("font-weight: bold; color: #2c3e50;")
        self.username_input = QLineEdit()
        layout.addWidget(self.username_label)
        layout.addWidget(self.username_input)

        self.password_label = QLabel("🔒 Enter Password:")
        self.password_label.setStyleSheet("font-weight: bold; color: #2c3e50;")
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.password_label)
        layout.addWidget(self.password_input)

        self.login_button = QPushButton("🚀 LOGIN")
        self.login_button.setStyleSheet("""
            background-color: #2563eb; /* Blue */
            color: white;
            font-weight: bold;
            padding:  3px 10px;  /* Reduced padding */
            border-radius: 5px;
        """)
        self.login_button.clicked.connect(self.login)
        layout.addWidget(self.login_button)

        self.status_label = QLabel("Status: Not logged in")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("""
            font-weight: bold;
            color: #34495e; /* Dark grey */
            padding: 5px;
        """)
        layout.addWidget(self.status_label) 

        self.load_excel_button = QPushButton("📂 LOAD CLIENT EXCEL FILE")
        self.load_excel_button.setStyleSheet("""
            background-color: #22c55e; /* Green */
            color: white;
            font-weight: bold;
            padding:  3px 10px;  /* Reduced padding */
            border-radius: 5px;
        """)
        self.load_excel_button.clicked.connect(self.open_excel_file)
        self.load_excel_button.setEnabled(False)
        layout.addWidget(self.load_excel_button)
        
        self.process_button = QPushButton("📊 PROCESS HOLDINGS")
        self.process_button.setStyleSheet("""
            background-color: #f59e0b; /* Yellow */
            color: white;
            font-weight: bold;
            padding:  3px 10px;  /* Reduced padding */
            border-radius: 5px;
        """)
        self.process_button.clicked.connect(self.process_holdings)
        self.process_button.setEnabled(True)
        layout.addWidget(self.process_button)

        self.summary_label = QLabel("📌 Summary: No processing yet.")
        self.summary_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.summary_label.setStyleSheet("""
            font-weight: bold;
            color: #475569; /* Soft blue-grey */
            padding: 8px;
        """)
        layout.addWidget(self.summary_label)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def login(self):
        url = self.url_input.text().strip()
        username = self.username_input.text().strip()
        password = self.password_input.text().strip()

        if not url or not username or not password:
            QMessageBox.warning(self, "Error", "Please fill in all fields!")
            return

        print(f"🔍 DEBUG: URL={url}, Username={username}, Password=******")

        try:
            print("🚀 Starting Selenium WebDriver...")
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-popup-blocking")

            prefs = {
                "download.default_directory": self.download_folder,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True
            }
            options.add_experimental_option("prefs", prefs)
            
            self.driver = webdriver.Chrome(options=options)
            self.driver.get(url)

            wait = WebDriverWait(self.driver, 15)  

            print("🔵 Waiting for username field...")
            username_field = wait.until(EC.presence_of_element_located((By.ID, "Input_Username")))  # ✅ Updated ID
            username_field.clear()
            username_field.send_keys(username)

            print("🔵 Waiting for password field...")
            password_field = wait.until(EC.presence_of_element_located((By.ID, "Input_Password")))  # ✅ Correct ID
            password_field.clear()
            password_field.send_keys(password)

            print("🔵 Clicking login button...")
            login_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "btn-primary-dark")))  # ✅ Updated selector
            login_button.click()

            time.sleep(5)

            tabs = self.driver.window_handles
            print(f"📝 Open Tabs: {tabs}")

            if len(tabs) > 1:
                self.driver.switch_to.window(tabs[-1])

            print("✅ Login Successful")
            self.status_label.setText("Login Successful")

            self.load_excel_button.setEnabled(True)

        except Exception as e:
            print(f"❌ Login Failed: {e}")
            self.status_label.setText("Login Failed")
            QMessageBox.critical(self, "Login Error", f"Failed to log in: {e}")

    def open_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls);;All Files (*)")
        if file_path:
            self.file_path = file_path
            print(f"📚 Selected file: {self.file_path}")

            df = pd.read_excel(self.file_path)
            self.client_codes = df['Client Code'].dropna().tolist()  

            if not self.client_codes:
                QMessageBox.warning(self, "Error", "No client codes found in the Excel file!")
                return

            QMessageBox.information(self, "Success", f"Loaded {len(self.client_codes)} client codes.")
            self.process_clients()

    def process_clients(self):
        for i in range(0, len(self.client_codes), self.max_parallel_clients):
            batch = self.client_codes[i:i+self.max_parallel_clients]
            print(f"🚀 Processing batch: {batch}")
            for client_code in batch:
                self.search_client(client_code)
            time.sleep(5)  

        QMessageBox.information(self, "Success", "All holdings extracted successfully!")


        
    def search_client(self, client_code):
        print(f"🔎 Processing client: {client_code}")
        wait = WebDriverWait(self.driver, 10)

        initial_tabs = self.driver.window_handles
        if len(initial_tabs) > 1:
            self.driver.switch_to.window(initial_tabs[1])  

        try:
            search_box = wait.until(EC.presence_of_element_located((By.ID, "UCBanner_txtSearch")))
            search_box.clear()
            search_box.send_keys(client_code)
            print(f"⌨️ Entering client code: {client_code}")

            first_suggestion = self.driver.find_element(By.CLASS_NAME, "ui-menu-item")
            first_suggestion.click()

            time.sleep(3)
            new_tabs = self.driver.window_handles
            if len(new_tabs) > len(initial_tabs):
                client_profile_tab = new_tabs[-1]
                self.driver.switch_to.window(client_profile_tab)
                print("🆕 Switched to the Client Profile tab.")

                capital_gain_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Capital Gain Report')]")))
                capital_gain_btn.click()

                time.sleep(3)
                dashboard_tabs = self.driver.window_handles
                if len(dashboard_tabs) > len(new_tabs):
                    client_dashboard_tab = dashboard_tabs[-1]
                    self.driver.switch_to.window(client_dashboard_tab)
                    print("📊 Switched to the Client Dashboard tab.")

                    self.driver.switch_to.window(client_profile_tab)
                    self.driver.close()
                    print("❌ Closed Client Profile tab.")

                    self.driver.switch_to.window(client_dashboard_tab)

                    self.download_holdings(client_code)

                    self.driver.close()
                    self.driver.switch_to.window(initial_tabs[1])
                    print("🔄 Closed client tab and returned to search tab.")
                    return True

            print(f"🚨 No new Client Dashboard tab opened for {client_code}")
            return False

        except Exception as e:
            print(f"❌ Error processing {client_code}: {str(e)}")
            return False

    def download_holdings(self, client_code):
        try:
            print("📊 Navigating to Holdings for {client_code}")
            wait = WebDriverWait(self.driver, 10)

            holding_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'Holding')]")))
            holding_menu.click()

            time.sleep(2)  
            as_on_date_holding = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "As on date holding")))
            as_on_date_holding.click()

            time.sleep(3)  

            print("💾 Downloading Holdings")
            excel_button = wait.until(EC.element_to_be_clickable((By.ID, "MainContent_imgExcel")))
            excel_button.click()

            time.sleep(4)
            
            downloaded_files = sorted(
                [f for f in os.listdir(self.download_folder) if f.endswith(".xls") or f.endswith(".xlsx")],
                key=lambda x: os.path.getmtime(os.path.join(self.download_folder, x)),
                reverse=True
            )

            if downloaded_files:
                latest_file = os.path.join(self.download_folder, downloaded_files[0])
                new_file_name = os.path.join(self.download_folder, f"{client_code}.xlsx")
                os.rename(latest_file, new_file_name)
                print(f"✅ Holdings saved as: {new_file_name}")

                self.successful_downloads += 1
                self.update_summary_label()
                return True
                
            else:
                print(f"⚠️ No file found for {client_code}")
                self.failed_clients.append(client_code)  
                self.update_summary_label()
                return False

        except Exception as e:
            print(f"❌ Error processing {client_code}: {e}")
            self.failed_clients.append(client_code)  
            self.update_summary_label()
            return False

    def update_summary_label(self):
        success_count = len(self.client_codes) - len(self.failed_clients)
        failed_count = len(self.failed_clients)

        failed_text = "\nFailed Clients: " + ", ".join(self.failed_clients) if self.failed_clients else ""

        self.summary_label.setText(
            f"✅ Downloaded Holdings: {success_count}/{len(self.client_codes)} Clients\n"
            f"❌ Failed Downloads: {failed_count}{failed_text}"
        )

    
    def process_holdings(self):
        try:
            folder_path = QFileDialog.getExistingDirectory(self, "Select Holdings Folder")
            
            if not folder_path:  
                QMessageBox.warning(self, "Error", "No folder selected. Please select a valid folder.")
                return  

            self.log(f"📂 Processing holdings from: {folder_path}")

            processor = HoldingsProcessor(folder_path)  
            output_file = processor.process_holdings()

            if output_file:
                df = pd.read_excel(output_file)
                client_count = df.shape[0]

                self.summary_label.setText(
                    f"✅ Extracted Holdings for {client_count} Clients.\n"
                    f"📂 Report Saved: {output_file}"
                )
                QMessageBox.information(self, "Success", 
                    f"✅ Holdings processing completed!\n\n"
                    f"📊 Clients Processed: {client_count}\n"
                    f"📂 Report saved: {output_file}")

                print(f"✅ Consolidated report saved at {output_file}")

            else:
                self.summary_label.setText("⚠️ No valid holdings files found.")
                QMessageBox.warning(self, "Error", "No valid holdings files found for processing.")

        except Exception as e:
            error_msg = f"❌ Error: {str(e)}"
            self.log(error_msg)
            QMessageBox.critical(self, "Critical Error", error_msg)

    def closeEvent(self, event):
        if self.driver:
            self.driver.quit()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = HoldingsExtractor()
    window.show()
    sys.exit(app.exec())
