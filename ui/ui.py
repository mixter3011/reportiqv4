import os
import glob
import time
from datetime import date
import pandas as pd
import traceback
import shutil
from generator.excel import excel_generator
from generator.report import report_gen
from generator.xirr import conv, parse_float, proc
from web.web import Scraper
from PyQt5.QtCore import Qt, QDate
from utils.processor import Processor
from PyQt5.QtWidgets import (
    QMainWindow, QPushButton, QVBoxLayout, QWidget,
    QFileDialog, QMessageBox, QLabel, QLineEdit, QHBoxLayout, QDateEdit, QCheckBox, QScrollArea
)

class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.scraper = None
        self.file_path = None
        self.processor = None
        self.dl_folder = self._get_dl_path()
        self.mf_folder = self._get_mf_path()
        
        self.required_files = {
            "Ledger": None,
            "MF Transactions": None,
            "SIP": None
        }
        
        self.init_ui()
        
    def _get_dl_path(self):
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        path = os.path.join(desktop, "Holding")
        if not os.path.exists(path):
            os.makedirs(path)
        return path

    def _get_mf_path(self):
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        path = os.path.join(desktop, "MF Transactions")
        if not os.path.exists(path):
            os.makedirs(path)
        return path
    
    def _convert_mf_transactions(self):
        try:
            self.log("Converting all MF transaction files to CSV...")
            mf_files = [f for f in os.listdir(self.mf_folder) if f.endswith((".xlsx", ".xls")) and "MFTrans" in f]
            converted = 0
            for file in mf_files:
                try:
                    self._convert_mf_transaction_file(file)
                    converted += 1
                except Exception as e:
                    self.log(f"Error converting {file}: {str(e)}")
            self.log(f"Converted {converted} out of {len(mf_files)} MF transaction files to CSV")
        except Exception as e:
            self.log(f"Error in batch conversion: {str(e)}")
    
    def _convert_ledger_files(self):
        ledger_dir = os.path.join(os.path.expanduser('~'), 'Desktop', 'Ledger')
        if not os.path.exists(ledger_dir):
            self.log("⚠️ Ledger directory not found!")
            return False

        conversion_needed = False
        conversion_done = False
    
        for fname in os.listdir(ledger_dir):
            if fname.lower().endswith(('.xlsx', '.xls')):
                conversion_needed = True
                try:
                    full_path = os.path.join(ledger_dir, fname)
                    csv_path = os.path.join(ledger_dir, fname.rsplit('.', 1)[0] + '.csv')
                
                    if not os.path.exists(csv_path):
                        df = pd.read_excel(full_path)
                        df.to_csv(csv_path, index=False)
                        self.log(f"Converted ledger: {fname} → {csv_path}")
                        conversion_done = True
                    else:
                        self.log(f"CSV already exists for: {fname}")
                except Exception as e:
                    self.log(f"🚨 Ledger conversion failed for {fname}: {str(e)}")
                    return False

        if conversion_needed:
            return conversion_done  
        else:
            if any(fname.endswith('.csv') for fname in os.listdir(ledger_dir)):
                self.log("✓ Ledger files already in CSV format")
                return True
            self.log("⚠️ No ledger files found (neither Excel nor CSV)")
            return False
    
    def _get_client_codes(self):
        manual_code = self.xirr_code_input.text().strip()
        if manual_code:
            return [manual_code], None
            
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Client Codes File", 
                                                 "", "Excel Files (*.xlsx *.xls)")
        if not file_path:
            return None, "No file selected"

        try:
            df = pd.read_excel(file_path)
            column_names = [col.lower().strip() for col in df.columns]
            client_code_variations = ['client code', 'clientcode', 'client_code', 
                                    'code', 'client id', 'clientid', 'client_id']
            code_column = None
            for variant in client_code_variations:
                if variant in column_names:
                    code_column = df.columns[column_names.index(variant)]
                    break
            if not code_column:
                return None, "No client code column found"
            return df[code_column].dropna().astype(str).str.strip().tolist(), None
        except Exception as e:
            return None, f"Error reading Excel: {str(e)}"
    
    def log(self, msg):
        print(msg)
        self.status_lbl.setText(msg)
    
    def fetch_manual_code(self):
        code = self.manual_code_input.text().strip()
        if not code:
            QMessageBox.warning(self, "Error", "Please enter a client code")
            return
        if not self.scraper:
            QMessageBox.warning(self, "Error", "Please login first")
            return
        
        self.log(f"Processing manual client code: {code}")
    
        reply = QMessageBox.question(
            self,
            "Download Holdings",
            f"Download holdings for client {code}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes
        )
    
        if reply == QMessageBox.StandardButton.Yes:
            holdings_success, holdings_fails = self.scraper.process_all_clients([code], self.update_sum)
            if holdings_success > 0:
                self.process_hdng(auto_continue=True, single_client=code)
        else:
            self.generate_excel(auto_mode=True, single_client=code)

    def init_ui(self):
        self.setWindowTitle("PORTFOLIO RETURNS TRACKER")
        self.setGeometry(100, 100, 500, 600)
        self.setStyleSheet("""
        background-color: #f5f5f5;
        font-family: Arial;
        font-size: 12px;
        """)

        layout = QVBoxLayout()
        
        title = QLabel("PORTFOLIO RETURNS TRACKER")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("""
        font-size: 20px;
        font-weight: bold;
        color: black;
        padding: 10px;
        """)
        layout.addWidget(title)

        url_lbl = QLabel("ENTER URL:")
        url_lbl.setStyleSheet("font-weight: bold; color: black;")
        self.url_in = QLineEdit()
        self.url_in.setPlaceholderText("https://mofirst.motilaloswal.com")
        self.url_in.setStyleSheet("color: black;")
        layout.addWidget(url_lbl)
        layout.addWidget(self.url_in)

        user_lbl = QLabel("ENTER USERNAME:")
        user_lbl.setStyleSheet("font-weight: bold; color: black;")
        self.user_in = QLineEdit()
        self.user_in.setPlaceholderText("ROHIT ABHAY")
        self.user_in.setStyleSheet("color: black;")
        layout.addWidget(user_lbl)
        layout.addWidget(self.user_in)

        pass_lbl = QLabel("ENTER PASSWORD:")
        pass_lbl.setStyleSheet("font-weight: bold; color: black;")
        self.pass_in = QLineEdit()
        self.pass_in.setPlaceholderText("******")
        self.pass_in.setEchoMode(QLineEdit.EchoMode.Password)
        self.pass_in.setStyleSheet("color: black;")
        layout.addWidget(pass_lbl)
        layout.addWidget(self.pass_in)

        login_btn = QPushButton("LOGIN")
        login_btn.setStyleSheet("""
        background-color: black;
        color: white;  
        font-weight: bold;
        padding: 3px 10px;
        border-radius: 5px;
        """)
        login_btn.clicked.connect(self.login)
        layout.addWidget(login_btn)

        self.status_lbl = QLabel("STATUS: NOT LOGGED IN")
        self.status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_lbl.setStyleSheet("""
            font-weight: bold;
            color: black;
            padding: 5px;
        """)
        layout.addWidget(self.status_lbl)
        
        manual_code_layout = QHBoxLayout()

        manual_input_title = QLabel("CLIENT CODE")
        manual_input_title.setStyleSheet("""
            font-size: 14px;
            font-weight: bold;
            color: black;
            padding: 5px;
        """)

        self.manual_code_input = QLineEdit()
        self.manual_code_input.setPlaceholderText("ROMO####")
        self.manual_code_input.setStyleSheet("color: black;")

        manual_fetch_btn = QPushButton("GENERATE")
        manual_fetch_btn.setStyleSheet("""
            background-color: black;
            color: white;
            font-weight: bold;
            padding: 3px 10px;
            border-radius: 5px;
        """)
        manual_fetch_btn.clicked.connect(self.fetch_manual_code)

        manual_code_layout.addWidget(manual_input_title)
        manual_code_layout.addWidget(self.manual_code_input)
        manual_code_layout.addWidget(manual_fetch_btn)

        layout.addLayout(manual_code_layout)

        self.excel_btn = QPushButton("GENERATE INTERNAL REVIEW SHEET")
        self.excel_btn.setStyleSheet("""
            background-color: black;
            color: white;
            font-weight: bold;
            padding: 3px 10px;
            border-radius: 5px;
        """)
        self.excel_btn.clicked.connect(self.open_excel)
        self.excel_btn.setEnabled(False)
        layout.addWidget(self.excel_btn)
        
        mf_date_title = QLabel("DATE RANGE FOR MF TRANSACTIONS")
        mf_date_title.setStyleSheet("""
            font-size: 14px;
            font-weight: bold;
            color: black;
            padding: 5px;
        """)
        layout.addWidget(mf_date_title)
        
        date_layout = QHBoxLayout()
        
        from_date_label = QLabel("FROM:")
        from_date_label.setStyleSheet("color: black;")
        date_layout.addWidget(from_date_label)
        
        self.from_date = QDateEdit(calendarPopup=True)
        self.from_date.setDate(QDate.currentDate().addMonths(-1))
        self.from_date.setStyleSheet("color: black;")
        date_layout.addWidget(self.from_date)
        
        to_date_label = QLabel("TO:")
        to_date_label.setStyleSheet("color: black;")
        date_layout.addWidget(to_date_label)
        
        self.to_date = QDateEdit(calendarPopup=True)
        self.to_date.setDate(QDate.currentDate())
        self.to_date.setStyleSheet("color: black;")
        date_layout.addWidget(self.to_date)
        
        self.use_date_range = QCheckBox("USE RANGE")
        self.use_date_range.setStyleSheet("color: black;")
        date_layout.addWidget(self.use_date_range)
        
        layout.addLayout(date_layout)
        
        start_date_layout = QHBoxLayout()
        
        start_date_label = QLabel("ENTER START DATE")
        start_date_label.setStyleSheet("color: black; font-size: 14px; font-weight: bold; padding: 5px;")
        start_date_layout.addWidget(start_date_label)
        
        self.start_date = QDateEdit(calendarPopup=True)
        self.start_date.setDate(QDate.currentDate())
        self.start_date.setStyleSheet("color: black;")
        start_date_layout.addWidget(self.start_date)
        
        layout.addLayout(start_date_layout)
        
        xirr_code_layout = QHBoxLayout()

        xirr_input_title = QLabel("ENTER CLIENT CODE FOR XIRR")
        xirr_input_title.setStyleSheet("""
            font-size: 14px;
            font-weight: bold;
            color: black;
            padding: 5px;
        """)

        self.xirr_code_input = QLineEdit()
        self.xirr_code_input.setPlaceholderText("ROMO####")
        self.xirr_code_input.setStyleSheet("color: black;")

        xirr_code_layout.addWidget(xirr_input_title)
        xirr_code_layout.addWidget(self.xirr_code_input)

        layout.addLayout(xirr_code_layout)
        
        init_portfolio_val_layout = QHBoxLayout()

        init_portfolio_val_title = QLabel("ENTER INITIAL PORTFOLIO VALUE")
        init_portfolio_val_title.setStyleSheet("""
            font-size: 14px;
            font-weight: bold;
            color: black;
            padding: 5px;
        """)

        self.init_portfolio_val_input = QLineEdit()
        self.init_portfolio_val_input.setPlaceholderText("₹1,00,000")
        self.init_portfolio_val_input.setStyleSheet("color: black;")

        init_portfolio_val_layout.addWidget(init_portfolio_val_title)
        init_portfolio_val_layout.addWidget(self.init_portfolio_val_input)

        layout.addLayout(init_portfolio_val_layout)
        
        gen_xirr_btn = QPushButton("GENERATE XIRR")
        gen_xirr_btn.setStyleSheet("""
            background-color: black;
            color: white;
            font-weight: bold;
            padding: 3px 10px;
            border-radius: 5px;                           
        """)
        gen_xirr_btn.clicked.connect(self.gen_xirr)
        layout.addWidget(gen_xirr_btn)
        
        gen_tracker_btn = QPushButton("GENERATE RETURN TRACKER SHEET")
        gen_tracker_btn.setStyleSheet("""
            background-color: black;
            color: white;
            font-weight: bold;
            padding: 3px 10px;
            border-radius: 5px;
        """)
        gen_tracker_btn.clicked.connect(self.generate_excel)
        layout.addWidget(gen_tracker_btn)
        
        generate_report_btn = QPushButton("GENERATE CLIENT PORTFOLIO REPORT")
        generate_report_btn.setStyleSheet("""
            background-color: black;
            color: white;
            font-weight: bold;
            padding: 3px 10px;
            border-radius: 5px;
        """)
        generate_report_btn.clicked.connect(self.generate_report)
        layout.addWidget(generate_report_btn)
        
        self.sum_lbl = QLabel("Summary: No processing yet.")
        self.sum_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.sum_lbl.setStyleSheet("""
            font-weight: bold;
            color: black;
            padding: 8px;
        """)
        layout.addWidget(self.sum_lbl)
        
        container = QWidget()
        container.setLayout(layout)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(container)

        self.setCentralWidget(scroll_area)
    
    def categorize_file(self, file_path):
        filename = os.path.basename(file_path).lower()
        
        if "ledger" in filename:
            return "Ledger"
        elif "mf" in filename or "transaction" in filename or "trans" in filename:
            return "MF Transactions"
        elif "sip" in filename:
            return "SIP"
        
        try:
            try:
                df = pd.read_excel(file_path)
            except:
                try:
                    df = pd.read_excel(file_path, engine='openpyxl')
                except:
                    try:
                        df = pd.read_excel(file_path, engine='xlrd')
                    except:
                        try:
                            df = pd.read_csv(file_path)
                        except:
                            return None
            
            headers = [col.lower() for col in df.columns]
            
            if any(h in headers for h in ['ledger', 'general ledger']):
                return "Ledger"
            elif any(h in headers for h in ['transaction', 'transactions', 'mf transaction']):
                return "MF Transactions"
            elif any(h in headers for h in ['sip', 'systematic']):
                return "SIP"
        except:
            pass
        
        msg_box = QMessageBox()
        msg_box.setWindowTitle("Select File Type")
        msg_box.setText(f"Please select the type for: {os.path.basename(file_path)}")
        
        ledger_btn = msg_box.addButton("Ledger", QMessageBox.ButtonRole.ActionRole)
        mf_btn = msg_box.addButton("MF Transactions", QMessageBox.ButtonRole.ActionRole)
        sip_btn = msg_box.addButton("SIP", QMessageBox.ButtonRole.ActionRole)
        cancel_btn = msg_box.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
        
        msg_box.exec()
        
        if msg_box.clickedButton() == ledger_btn:
            return "Ledger"
        elif msg_box.clickedButton() == mf_btn:
            return "MF Transactions"
        elif msg_box.clickedButton() == sip_btn:
            return "SIP"
        else:
            return None
    
    def process_uploaded_files(self, file_paths):
        for file_path in file_paths:
            file_type = self.categorize_file(file_path)
            if file_type:
                self.required_files[file_type] = file_path
                self.uploaded_files_display.update_file(file_type, file_path)
        
        files_received = sum(1 for path in self.required_files.values() if path is not None)
        
        self.log(f"Uploaded files: {files_received}/3 required files uploaded.")
        
        if files_received == 3:
            self.sum_lbl.setText("All required files uploaded. Ready for processing.")
            self.sum_lbl.setStyleSheet("""
                font-weight: bold;
                color: black;
                padding: 8px;
            """)

    def login(self):
        url = self.url_in.text().strip()
        user = self.user_in.text().strip()
        pwd = self.pass_in.text().strip()

        if not url or not user or not pwd:
            QMessageBox.warning(self, "Error", "Please fill in all fields")
            return

        try:
            self.scraper = Scraper(self.dl_folder, self.mf_folder)
            if self.scraper.login(url, user, pwd):
                self.status_lbl.setText("Login successful")
                self.excel_btn.setEnabled(True)
            else:
                self.status_lbl.setText("Login failed")
                QMessageBox.critical(self, "Login Error", "Failed to log in")
        except Exception as e:
            self.status_lbl.setText("Login failed")
            QMessageBox.critical(self, "Login Error", f"Failed to log in: {e}")

    def open_excel(self):
        reply = QMessageBox.question(
            self,
            "Download Holdings",
            "Do you want to download holdings before generating reports?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            file_path, _ = QFileDialog.getOpenFileName(self, "Select Client Codes File", 
                                                     "", "Excel Files (*.xlsx *.xls)")
            if not file_path:
                return

            try:
                df = pd.read_excel(file_path)
                column_names = [col.lower().strip() for col in df.columns]
                client_code_variations = ['client code', 'clientcode', 'client_code', 'code', 'client id', 'clientid', 'client_id']
                code_column = None
                for variant in client_code_variations:
                    if variant in column_names:
                        code_column = df.columns[column_names.index(variant)]
                        break

                if code_column is None:
                    QMessageBox.warning(self, "Error", "No client code column found!")
                    return

                codes = df[code_column].dropna().astype(str).tolist()
                codes = [code.strip() for code in codes]

                if not codes:
                    QMessageBox.warning(self, "Error", "No client codes found in the Excel file!")
                    return

                self.status_lbl.setText(f"Downloading holdings for {len(codes)} clients...")
                holdings_success, holdings_fails = self.scraper.process_all_clients(codes, self.update_sum)

                if holdings_success > 0:
                    self.process_hdng(auto_continue=True)
                else:
                    QMessageBox.critical(self, "Error", "Failed to download any holdings!")

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error processing file: {str(e)}")
        else:
            self.generate_excel(auto_mode=False)
    
    def update_sum(self, success, total, fails):
        fail_txt = "\nFailed clients: " + ", ".join(fails) if fails else ""
        self.sum_lbl.setText(
            f"Downloaded holdings: {success}/{total} clients\n"
            f"Failed downloads: {len(fails)}{fail_txt}"
        )
    
    def process_hdng(self, auto_continue=False, single_client=None):
        folder = self.dl_folder
        self.log(f"Processing holdings from: {folder}")
        try:
            excel_files = [os.path.join(folder, f) for f in os.listdir(folder)
                          if f.endswith(('.xlsx', '.xls'))]
        
            if not excel_files:
                self.sum_lbl.setText("No Excel files found in Holdings folder.")
                QMessageBox.warning(self, "Error", "No Excel files found in Holdings folder.")
                return
            
            if single_client:
                excel_files = [f for f in excel_files if single_client in os.path.basename(f)]
                self.log(f"Filtered to {len(excel_files)} files for client {single_client}")
            
            converted_count = 0
            for excel_file in excel_files:
                try:
                    try:
                        df = pd.read_excel(excel_file)
                    except Exception as e1:
                        try:
                            df = pd.read_excel(excel_file, engine='openpyxl')
                        except Exception as e2:
                            df = pd.read_excel(excel_file, engine='xlrd')
                    csv_file = os.path.splitext(excel_file)[0] + '.csv'
                    df.to_csv(csv_file, index=False)
                    converted_count += 1
                except Exception as e:
                    print(f"Error converting {excel_file}: {str(e)}")
                
            self.sum_lbl.setText(
                f"Converted {converted_count}/{len(excel_files)} Excel files to CSV in {folder}"
            )
        
            self.processor = Processor(folder)
            if hasattr(self.processor, 'set_required_files') and "Ledger" in self.required_files and self.required_files["Ledger"] is not None:
                self.processor.set_required_files(
                    ledger=self.required_files["Ledger"],
                    mf_transactions=None,
                    sip=None
                )
            
            out_file = self.processor.process_holdings()
            if out_file:
                df = pd.read_excel(out_file)
                count = df.shape[0]
                self.sum_lbl.setText(
                    f"Converted {converted_count} files to CSV.\n"
                    f"Extracted holdings for {count} clients.\n"
                    f"Report saved: {out_file}"
                )
            
            if auto_continue:
                self.status_lbl.setText("Generating internal review sheets...")
                self.generate_excel(auto_mode=True, single_client=single_client)
            
        except Exception as e:
            error_details = traceback.format_exc()
            print(f"Processing error: {error_details}")
            err_msg = f"Error: {str(e)}"
            self.log(err_msg)
            QMessageBox.critical(self, "Critical Error", err_msg)
    
    def process_mf_trans(self):
        missing_files = []
    
        if "MF Transactions" not in self.required_files or self.required_files["MF Transactions"] is None:
            missing_files.append("MF Transactions")
    
        if "SIP" not in self.required_files or self.required_files["SIP"] is None:
            missing_files.append("SIP")
    
        if missing_files:
            QMessageBox.warning(self, "Missing Files", 
                            f"Please upload the following required files: {', '.join(missing_files)}")
            return
    
        folder = QFileDialog.getExistingDirectory(self, "Select MF transactions folder", self.dl_folder)
    
        if not folder:
            QMessageBox.warning(self, "Error", "No folder selected.")
            return

        self.log(f"Processing MF transactions from: {folder}")

        try:
            date_range_info = ""
            if hasattr(self, 'use_date_range') and self.use_date_range.isChecked():
                from_date = self.from_date.date().toString("dd/MM/yyyy")
                to_date = self.to_date.date().toString("dd/MM/yyyy")
                date_range_info = f" with date range: {from_date} to {to_date}"
                self.log(f"Using date range: {from_date} to {to_date}")

            self.processor = Processor(folder)
        
            if hasattr(self.processor, 'set_required_files'):
                self.processor.set_required_files(
                    ledger=self.required_files.get("Ledger"),
                    mf_transactions=self.required_files["MF Transactions"],
                    sip=self.required_files["SIP"]
                )
        
            out_file = self.processor.run_mf_transactions()

            if out_file:
                df = pd.read_excel(out_file)
                count = df.shape[0]

                self.sum_lbl.setText(
                    f"Processed MF transactions{date_range_info} for {count} clients.\n"
                    f"Report saved: {out_file}"
                )
                QMessageBox.information(self, "Success", 
                    f"MF transactions processing completed{date_range_info}!\n\n"
                    f"Clients processed: {count}\n"
                    f"Report saved: {out_file}")
            else:
                self.sum_lbl.setText("No valid MF transactions files found.")
                QMessageBox.warning(self, "Error", "No valid MF transactions files found.")
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Processing error: {error_details}")
            err_msg = f"Error: {str(e)}"
            self.log(err_msg)
            QMessageBox.critical(self, "Critical Error", err_msg)
    
    def generate_report(self):
        try:
            self.log("Generating report...")

            desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
            holding_folder = os.path.join(desktop, "Holding")
            ledger_folder = os.path.join(desktop, "Ledger")
            client_reports_folder = os.path.join(desktop, "client_reports")
            client_xirr_folder = os.path.join(desktop, "xirr_reports")

            for folder in [holding_folder, ledger_folder, client_reports_folder, client_xirr_folder]:
                if not os.path.exists(folder):
                    os.makedirs(folder)

            holding_csv_files = [f for f in os.listdir(holding_folder) if f.endswith('.csv')]

            if not holding_csv_files:
                QMessageBox.warning(self, "Error", "No CSV files found in Holding folder.")
                return

            processed_count = 0
            skipped_count = 0

            for holding_csv in holding_csv_files:
                try:
                    base_filename = os.path.splitext(holding_csv)[0]

                    ledger_file = os.path.join(ledger_folder, f"{base_filename}_Ledger.csv")
    
                    if not os.path.exists(ledger_file):
                        for ext in ['.csv', '.xlsx', '.xls']:
                            potential_file = os.path.join(ledger_folder, f"{base_filename}_Ledger{ext}")
                            if os.path.exists(potential_file):
                                ledger_file = potential_file
                                break
        
                        if not os.path.exists(ledger_file):
                            for ext in ['.csv', '.xlsx', '.xls']:
                                potential_file = os.path.join(ledger_folder, base_filename + ext)
                                if os.path.exists(potential_file):
                                    ledger_file = potential_file
                                    break

                    if not os.path.exists(ledger_file):
                        self.log(f"Skipped {holding_csv}: No matching file in Ledger folder")
                        skipped_count += 1
                        continue

                    if not ledger_file.endswith('.csv'):
                        try:
                            try:
                                ledger_df = pd.read_excel(ledger_file)
                            except Exception as e1:
                                try:
                                    ledger_df = pd.read_excel(ledger_file, engine='openpyxl')
                                except Exception as e2:
                                    ledger_df = pd.read_excel(ledger_file, engine='xlrd')
        
                            ledger_csv_file = os.path.join(ledger_folder, f"{base_filename}_Ledger.csv")
                            ledger_df.to_csv(ledger_csv_file, index=False)
                            ledger_file = ledger_csv_file
        
                        except Exception as e:
                            self.log(f"Failed to convert {ledger_file} to CSV: {str(e)}")
                            skipped_count += 1
                            continue

                    holding_df = pd.read_csv(os.path.join(holding_folder, holding_csv))
                    ledger_df = pd.read_csv(ledger_file)
        
                    xirr_df = None
                    try:
                        info_row = holding_df[holding_df['Unnamed: 0'] == 'Client Equity Code/UCID/Name'].index[0]
                        c_info = str(holding_df.iloc[info_row, 1]).strip()
                        client_code = c_info.split('/')[0].strip()
            
                        xirr_files = [f for f in os.listdir(client_xirr_folder) if f.endswith(('.csv', '.xlsx', '.xls'))]
                        xirr_file = None
            
                        for file in xirr_files:
                            if client_code in file:
                                xirr_file = os.path.join(client_xirr_folder, file)
                                break
            
                        if xirr_file:
                            if xirr_file.endswith('.csv'):
                                xirr_df = pd.read_csv(xirr_file)
                            else:
                                try:
                                    xirr_df = pd.read_excel(xirr_file)
                                except Exception as e1:
                                    try:
                                        xirr_df = pd.read_excel(xirr_file, engine='openpyxl')
                                    except Exception as e2:
                                        xirr_df = pd.read_excel(xirr_file, engine='xlrd')
                
                            self.log(f"Found XIRR file for {client_code}: {os.path.basename(xirr_file)}")
                        else:
                            self.log(f"No XIRR file found for client code: {client_code}")
                
                    except Exception as e:
                        self.log(f"Error processing XIRR for {holding_csv}: {str(e)}")

                    output_file = os.path.join(client_reports_folder, f"{base_filename}_report.pdf")
                
                    try:
                        pdf_path = report_gen(holding_df, ledger_df, xirr_df, output_path=output_file)
                    
                        generated_path = os.path.abspath(pdf_path)
                    
                        if os.path.exists(generated_path) and os.path.getsize(generated_path) > 0:
                            processed_count += 1
                            self.log(f"Generated report for {base_filename} at {generated_path}")
                        else:
                            self.log(f"Failed to generate report for {base_filename} or file is empty. Expected at: {generated_path}")
                            skipped_count += 1
                    except Exception as e:
                        self.log(f"Error in report_gen for {base_filename}: {str(e)}")
                        skipped_count += 1

                except Exception as e:
                    self.log(f"Error processing {holding_csv}: {str(e)}")
                    skipped_count += 1

            if processed_count > 0:
                self.sum_lbl.setText(
                    f"Generated {processed_count} client reports.\n"
                    f"Skipped {skipped_count} files.\n"
                    f"Reports saved to: {client_reports_folder}"
                )
                QMessageBox.information(self, "Success", 
                    f"Client reports successfully generated!\n\n"
                    f"Files processed: {processed_count}\n"
                    f"Files skipped: {skipped_count}\n"
                    f"Reports location: {client_reports_folder}")
            else:
                self.log("Failed to generate any client reports")
                QMessageBox.warning(self, "Error", "Failed to generate any client reports")

        except Exception as e:
            error_details = traceback.format_exc()
            print(f"Report generation error: {error_details}")
            self.log(f"Report generation error: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to generate reports: {str(e)}")

    def generate_excel(self, auto_mode=False, single_client=None):
        try:
            self.log("Generating Excel files from CSV files...")
            folder = self.dl_folder  
            ledger_folder = os.path.join(os.path.dirname(folder), "Ledger")  

            desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
            excel_reports_folder = os.path.join(desktop, "excel_reports")
            if not os.path.exists(excel_reports_folder):
                os.makedirs(excel_reports_folder)

            holdings_files = [os.path.join(folder, f) for f in os.listdir(folder) 
                        if f.endswith('.csv')]

            if single_client:
                holdings_files = [f for f in holdings_files if single_client in os.path.basename(f)]
                self.log(f"Filtered to {len(holdings_files)} holdings files for client {single_client}")

            if not holdings_files:
                self.log("No matching CSV files found in Holdings folder.")
                QMessageBox.warning(self, "Error", "No matching CSV files found in Holdings folder.")
                return
        
            ledger_files = {}
            if os.path.exists(ledger_folder):
                for f in os.listdir(ledger_folder):
                    if f.endswith('_Ledger.csv'):
                        client_code = f.replace('_Ledger.csv', '')
                        if not single_client or single_client == client_code:
                            ledger_files[client_code] = os.path.join(ledger_folder, f)
                
                if single_client and single_client not in ledger_files:
                    self.log(f"Warning: No ledger file found for client {single_client}")
            else:
                self.log("Ledger folder not found. Only Holdings data will be processed.")

            processed_count = 0
            for holdings_csv in holdings_files:
                try:
                    base_filename = os.path.splitext(os.path.basename(holdings_csv))[0]
                    client_code = base_filename  
            
                    df_holdings = pd.read_csv(holdings_csv)
                    if df_holdings.empty:
                        print(f"Skipping empty holdings file: {holdings_csv}")
                        continue
            
                    df_ledger = None
                    if client_code in ledger_files:
                        ledger_csv = ledger_files[client_code]
                        try:
                            df_ledger = pd.read_csv(ledger_csv)
                            print(f"Found matching ledger file for {client_code}: {ledger_csv}")
                        except Exception as le:
                            print(f"Error reading ledger file {ledger_csv}: {str(le)}")
                    else:
                        print(f"No matching ledger file found for client code: {client_code}")
            
                    output_file = excel_generator(df_holdings, df_ledger)  
        
                    if output_file:
                        if os.path.exists(output_file):
                            dest_file = os.path.join(excel_reports_folder, f"{base_filename}_report.xlsx")
                            shutil.move(output_file, dest_file)
                            processed_count += 1
                        else:
                            print(f"Output file not found: {output_file}")
                    else:
                        try:
                            potential_dirs = [os.getcwd(), os.path.dirname(holdings_csv), desktop]
                            newest_file = None
                            newest_time = 0
                            for check_dir in potential_dirs:
                                excel_files = glob.glob(os.path.join(check_dir, "*.xlsx"))
                                for file in excel_files:
                                    file_time = os.path.getmtime(file)
                        
                                    if time.time() - file_time < 10 and file_time > newest_time:  
                                        newest_file = file
                                        newest_time = file_time
                
                            if newest_file:
                                dest_file = os.path.join(excel_reports_folder, f"{base_filename}_report.xlsx")
                                shutil.copy2(newest_file, dest_file)
                                processed_count += 1
                                print(f"Processed: {holdings_csv} → {dest_file} (found recent file)")
                            else:
                                print(f"Could not locate output file for {holdings_csv}")
                        except Exception as inner_e:
                            print(f"Error locating output for {holdings_csv}: {str(inner_e)}")
            
                except Exception as e:
                    print(f"Error processing {holdings_csv}: {str(e)}")

            if processed_count > 0:
                self.log(f"Generated {processed_count}/{len(holdings_files)} Excel reports in {excel_reports_folder}")
            
                if single_client:
                    success_message = (f"Internal Report generated for client {single_client}!\n\n"
                                f"Report location: {excel_reports_folder}")
                else:
                    success_message = (f"Complete workflow executed successfully!\n\n"
                                f"Files processed: {processed_count}/{len(holdings_files)}\n"
                                f"Reports location: {excel_reports_folder}")
            
                if auto_mode:
                    self.status_lbl.setText("Completed full workflow")
                    if single_client:
                        self.sum_lbl.setText(f"Generated Internal Sheet for client {single_client}.")
                    else:
                        self.sum_lbl.setText(f"Automated workflow complete. Generated {processed_count} internal review sheets.")
            
                QMessageBox.information(self, "Success", success_message)
            else:
                if single_client:
                    self.log(f"Failed to generate Excel report for client {single_client}")
                    QMessageBox.warning(self, "Error", f"Failed to generate Excel report for client {single_client}")
                else:
                    self.log("Failed to generate any Excel reports")
                    QMessageBox.warning(self, "Error", "Failed to generate any Excel reports")

        except Exception as e:
            error_details = traceback.format_exc()
            print(f"Excel generation error: {error_details}")
            QMessageBox.critical(self, "Error", f"Failed to generate Excel files: {str(e)}")

    def xirr_workflow(self):
        from_date = self.from_date.date().toString("dd/MM/yyyy") if self.use_date_range.isChecked() else None
        to_date = self.to_date.date().toString("dd/MM/yyyy") if self.use_date_range.isChecked() else None

        codes, error = self._get_client_codes()
        if error:
            QMessageBox.critical(self, "Error", error)
            return

        self.log(f"Downloading MF transactions for {len(codes)} clients...")
        success, fails = self.scraper.process_all_clients_mf_trans(codes, self.update_sum, 
                                                                 from_date=from_date, to_date=to_date)
    
        if success == 0:
            QMessageBox.warning(self, "Warning", "No MF transactions downloaded")
            return

        mf_folder = self.mf_folder
        for fname in os.listdir(mf_folder):
            if fname.endswith(('.xlsx', '.xls')):
                csv_path = conv(os.path.join(mf_folder, fname))
                if csv_path:
                    print(f"Converted {fname} to CSV")
    
        try:
            init_val = float(self.init_portfolio_val_input.text().replace('₹','').replace(',',''))
            curr_val = float(self.cur_portfolio_val_input.text().replace('₹','').replace(',',''))
        except:
            QMessageBox.warning(self, "Error", "Invalid portfolio values")
            return

        try:
            results = proc(input_dir=mf_folder, 
                          init_val=init_val,
                          curr_val=curr_val)
        
            if results:
                QMessageBox.information(self, "Success", 
                    f"XIRR reports generated:\n{'\n'.join(results)}")
            else:
                QMessageBox.warning(self, "Warning", "XIRR calculation completed with no results")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"XIRR failed: {str(e)}")

    def gen_xirr(self):
        try:
            manual_code = self.xirr_code_input.text().strip()
            
            from_date = None
            to_date = None
            if self.use_date_range.isChecked():
                from_date = self.from_date.date().toString("dd/MM/yyyy")
                to_date = self.to_date.date().toString("dd/MM/yyyy")
                self.log(f"Using date range: {from_date} to {to_date}")
        
            if manual_code:
                self.log(f"Manual client code detected: {manual_code}")
                
                try:
                    init_val = float(self.init_portfolio_val_input.text().replace('₹','').replace(',',''))
                except:
                    init_val = 100000  
                    self.log(f"Using default initial value: {init_val}")
                
                start_date = self.start_date.date().toPyDate()
                self.log(f"Using manual start date: {start_date}")
                
                try:
                    curr_val = float(self.cur_portfolio_val_input.text().replace('₹','').replace(',',''))
                except:
                    curr_val = init_val 
                    self.log(f"Using initial value as current value: {curr_val}")
            
                reply = QMessageBox.question(
                    self,
                    "Process Manual Client",
                    f"Download MF transactions for {manual_code}?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes
                )
            
                if reply == QMessageBox.StandardButton.Yes:
                    self.status_lbl.setText(f"Downloading MF transactions for {manual_code}...")
                    success = self.scraper.search_client_mf_trans(manual_code, from_date, to_date)
                
                    if success:
                        self.log(f"✅ Successfully downloaded MF transactions for {manual_code}")
                        self._convert_mf_transaction_file(f"{manual_code}_MFTrans.xlsx")
                    else:
                        self.log(f"❌ Failed to download MF transactions for {manual_code}")
                        QMessageBox.warning(self, "Warning", f"Failed to download MF transactions for {manual_code}")
                
                self.status_lbl.setText(f"Generating XIRR for {manual_code}...")
                out_file = proc(code=manual_code, init_val=init_val, curr_val=curr_val, start_date=start_date)
                
                if out_file:
                    self.log(f"✅ Successfully generated XIRR report for {manual_code}")
                    self.status_lbl.setText(f"XIRR report generated for {manual_code}")
                    QMessageBox.information(
                        self, 
                        "XIRR Report Generated", 
                        f"XIRR report has been generated for {manual_code}.\n\nSaved to: {out_file}"
                    )
                else:
                    self.log(f"❌ Failed to generate XIRR report for {manual_code}")
                    QMessageBox.warning(
                        self, 
                        "Warning", 
                        f"Failed to generate XIRR report for {manual_code}. Check if required files exist."
                    )
                return


            self.log("Prompting user to select client codes file")
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Select Client Codes File", 
                "", "Excel Files (*.xlsx *.xls)"
            )
            if not file_path:
                self.log("User cancelled file selection.")
                return

            try:
                self.log(f"Loading client codes from: {file_path}")
                df = pd.read_excel(file_path)
            
                code_col = None
                init_val_col = None
                start_date_col = None
            
                for col in df.columns:
                    col_lower = col.lower()
                    if 'client' in col_lower and 'code' in col_lower:
                        code_col = col
                    elif 'initial' in col_lower and 'value' in col_lower:
                        init_val_col = col
                    elif 'start' in col_lower and 'date' in col_lower:
                        start_date_col = col
            
                if not code_col:
                    self.log("Could not find client code column in Excel file")
                    QMessageBox.critical(self, "Error", "Could not find client code column in Excel file")
                    return
            
                if not init_val_col:
                    self.log("Warning: Initial value column not found, will use default values")
                    QMessageBox.warning(self, "Warning", "Initial value column not found. Will use default values.")
            
                if not start_date_col:
                    self.log("Warning: Start date column not found, will use default dates")
                    QMessageBox.warning(self, "Warning", "Start date column not found. Will use default dates.")
            
                client_data = {}
                for _, row in df.iterrows():
                    code = str(row[code_col]).strip()
                    if not code or pd.isna(code):
                        continue
                
                    client_data[code] = {
                        'code': code,
                        'init_val': float(row[init_val_col]) if init_val_col and not pd.isna(row[init_val_col]) else None,
                        'start_date': None
                    }
                
                    if start_date_col and not pd.isna(row[start_date_col]):
                        start_date_raw = row[start_date_col]
                        if isinstance(start_date_raw, str):
                            try:
                                for fmt in ['%d/%m/%y', '%d/%m/%Y', '%m/%d/%y', '%m/%d/%Y']:
                                    try:
                                        client_data[code]['start_date'] = pd.to_datetime(start_date_raw, format=fmt).date()
                                        break
                                    except:
                                        continue
                                else:
                                    client_data[code]['start_date'] = pd.to_datetime(start_date_raw).date()
                            except:
                                self.log(f"Could not parse start date for {code}: {start_date_raw}")
                        else:
                            client_data[code]['start_date'] = pd.to_datetime(start_date_raw).date()
            
                codes = list(client_data.keys())
                self.log(f"Found {len(codes)} client codes in file")
            
                if not codes:
                    self.log("No valid client codes found in file")
                    QMessageBox.warning(self, "Warning", "No valid client codes found in file")
                    return
                
            except Exception as e:
                self.log(f"Error reading client codes file: {str(e)}")
                QMessageBox.critical(self, "Error", f"Error reading client file: {str(e)}")
                return

            reply = QMessageBox.question(
                self,
                "Download MF Transactions",
                f"Download MF transactions for {len(codes)} clients?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes
            )

            if reply == QMessageBox.StandardButton.Yes:
                self.status_lbl.setText(f"Downloading MF transactions for {len(codes)} clients...")
                success, fails = self.scraper.process_all_clients_mf_trans(codes, self.update_sum, from_date, to_date)
                self.log(f"Download complete: {success} succeeded, {len(fails)} failed")
            
                if fails:
                    self.log(f"Failed clients: {', '.join(fails)}")
                    QMessageBox.warning(self, "Warning", f"Failed to download MF transactions for {len(fails)} clients")
            
                self._convert_mf_transactions()
            
                self.status_lbl.setText("Generating XIRR reports...")
            
                try:
                    out_files = []
                
                    default_init_val = None
                    try:
                        default_init_val = float(self.init_portfolio_val_input.text().replace('₹','').replace(',',''))
                    except:
                        default_init_val = 100000
                
                    default_start_date = self.start_date.date().toPyDate()
                
                    for code in codes:
                        if code in fails:
                            self.log(f"Skipping failed client: {code}")
                            continue
                        
                        client = client_data[code]
                        init_val = client['init_val'] if client['init_val'] is not None else default_init_val
                        start_date = client['start_date'] if client['start_date'] is not None else default_start_date
                    
                        self.log(f"Processing XIRR for {code}: init_val={init_val}, start_date={start_date}")
                    
                        try:
                            report_file = proc(code=code, init_val=init_val, start_date=start_date)
                            if report_file:
                                out_files.append(report_file)
                                self.log(f"Generated XIRR report for {code}: {report_file}")
                            else:
                                self.log(f"Failed to generate XIRR report for {code}")
                        except Exception as client_error:
                            self.log(f"Error processing XIRR for {code}: {str(client_error)}")
                
                    if out_files and len(out_files) > 0:
                        self.log(f"✅ Successfully generated {len(out_files)} XIRR reports")
                        self.status_lbl.setText(f"Generated {len(out_files)} XIRR reports")
                    
                        report_list = "\n".join(out_files[:5])
                        if len(out_files) > 5:
                            report_list += f"\n... and {len(out_files) - 5} more"
                    
                        QMessageBox.information(
                            self, 
                            "XIRR Reports Generated", 
                            f"Generated {len(out_files)} XIRR reports.\n\nReports saved to Desktop/xirr_reports\n\nSample reports:\n{report_list}"
                        )
                    else:
                        self.log("⚠️ No XIRR reports were generated")
                        QMessageBox.warning(
                            self, 
                            "Warning", 
                            "No XIRR reports were generated. Please check if required files exist."
                        )
                except Exception as e:
                    self.log(f"❌ Error generating XIRR reports: {str(e)}")
                    QMessageBox.critical(
                        self, 
                        "Error", 
                        f"Failed to generate XIRR reports: {str(e)}"
                    )
            
        except Exception as e:
            self.log(f"Error in gen_xirr: {str(e)}")
            QMessageBox.critical(
                self, 
                "Error", 
                f"XIRR generation failed: {str(e)}\n\n"
                f"Common issues:\n"
                f"1. Check file paths and format\n"
                f"2. Validate client codes format\n"
                f"3. Verify date formats"
            )
     
    def process_single_xirr(self, client_code):
        try:
            file_path, _ = QFileDialog.getOpenFileName(self, "Select Client Data File", 
                                                       "", "Excel Files (*.xlsx *.xls)")
            if not file_path:
                return

            df = pd.read_excel(file_path)
            code_col = next(col for col in df.columns if 'client' in col.lower() or 'code' in col.lower())
            init_val_col = next(col for col in df.columns if 'initial' in col.lower() and 'value' in col.lower())
            start_date_col = next(col for col in df.columns if 'start date' in col.lower())
            
            client_data = df[df[code_col] == client_code].iloc[0]
            init_val = client_data[init_val_col]
            start_date = pd.to_datetime(client_data[start_date_col]).date()
            
            consolidated_path = os.path.join(self.dl_folder, "Consolidated_Holdings.xlsx")
            if not os.path.exists(consolidated_path):
                QMessageBox.critical(self, "Error", "Consolidated_Holdings.xlsx not found!")
                return
                
            consolidated_df = pd.read_excel(consolidated_path)
            curr_val_col = next(col for col in consolidated_df.columns if 'portfolio value' in col.lower())
            curr_val = consolidated_df.loc[consolidated_df[code_col] == client_code, curr_val_col].values[0]
            
            reply = QMessageBox.question(
                self,
                "Download MF Transactions",
                f"Do you want to download MF transactions for {client_code}?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                self.status_lbl.setText(f"Downloading MF transactions for {client_code}...")
                success, _ = self.scraper.process_all_clients_mf_trans([client_code], self.update_sum)
                self._convert_mf_excel_to_csv()
                self._convert_ledger_files()
            
            result = proc(cl_code=client_code, init_val=init_val, curr_val=curr_val, start_date=start_date)
            QMessageBox.information(self, "Success", f"XIRR report generated:\n{result}")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Single client XIRR failed: {str(e)}") 
     
    def closeEvent(self, event):
        if self.scraper:
            self.scraper.quit()
        event.accept()