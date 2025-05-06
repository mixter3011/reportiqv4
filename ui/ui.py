import os
import glob
import time
from datetime import date
import pandas as pd
from generator.excel import excel_generator
from generator.report import report_gen
from generator.xirr import proc
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
        
        holdings_success, holdings_fails = self.scraper.process_all_clients([code], self.update_sum)
        
        if holdings_success > 0:
            self.status_lbl.setText(f"Successfully downloaded holdings for {code}")
            
            choice = QMessageBox.question(
                self,
                "Download MF Transactions",
                f"Holdings downloaded for {code}. Do you want to download MF transactions too?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes
            )
            
            if choice == QMessageBox.StandardButton.Yes:
                self.status_lbl.setText(f"Downloading MF transactions for {code}...")
                mf_success, mf_fails = self.scraper.process_all_clients_mf_trans([code], self.update_sum)
                
                if mf_success > 0:
                    self.status_lbl.setText(f"Successfully downloaded holdings and MF transactions for {code}")
                    QMessageBox.information(self, "Success", 
                        f"Successfully downloaded holdings and MF transactions for {code}")
                else:
                    self.status_lbl.setText(f"Downloaded holdings but failed to download MF transactions for {code}")
                    QMessageBox.warning(self, "Partial Success", 
                        f"Downloaded holdings but failed to download MF transactions for {code}")
            else:
                QMessageBox.information(self, "Success", f"Downloaded holdings for {code}")
        else:
            self.status_lbl.setText(f"Failed to download holdings for {code}")
            QMessageBox.critical(self, "Error", f"Failed to download holdings for {code}")

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

        manual_fetch_btn = QPushButton("LOAD")
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

        self.excel_btn = QPushButton("LOAD CLIENT CODES EXCEL FILE")
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
        
        mf_date_title = QLabel("DATE RANGE")
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
        
        proc_mf_btn = QPushButton("DOWNLOAD MF TRANSACTIONS")
        proc_mf_btn.setStyleSheet("""
            background-color: black;
            color: white;
            font-weight: bold;
            padding: 3px 10px;
            border-radius: 5px;
        """)
        proc_mf_btn.clicked.connect(self.process_mf_trans)
        layout.addWidget(proc_mf_btn)

        proc_btn = QPushButton("PROCESS HOLDINGS")
        proc_btn.setStyleSheet("""
            background-color: black;
            color: white;
            font-weight: bold;
            padding: 3px 10px;
            border-radius: 5px;
        """)
        proc_btn.clicked.connect(self.process_hdng)
        layout.addWidget(proc_btn)        
        
        generate_excel_btn = QPushButton("GENERATE INTERNAL REVIEW SHEET")
        generate_excel_btn.setStyleSheet("""
            background-color: black;
            color: white;
            font-weight: bold;
            padding: 3px 10px;
            border-radius: 5px;
        """)
        generate_excel_btn.clicked.connect(self.generate_excel)
        layout.addWidget(generate_excel_btn)
        
        start_date_layout = QHBoxLayout()
        
        start_date_label = QLabel("ENTER START DATE")
        start_date_label.setStyleSheet("color: black;")
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
        
        cur_portfolio_val_layout = QHBoxLayout()

        cur_portfolio_val_title = QLabel("ENTER CURRENT PORTFOLIO VALUE")
        cur_portfolio_val_title.setStyleSheet("""
            font-size: 14px;
            font-weight: bold;
            color: black;
            padding: 5px;
        """)

        self.cur_portfolio_val_input = QLineEdit()
        self.cur_portfolio_val_input.setPlaceholderText("₹1,00,000")
        self.cur_portfolio_val_input.setStyleSheet("color: black;")

        cur_portfolio_val_layout.addWidget(cur_portfolio_val_title)
        cur_portfolio_val_layout.addWidget(self.cur_portfolio_val_input)

        layout.addLayout(cur_portfolio_val_layout)
        
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
        file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls);;All Files (*)")
        if not file_path:
            return
        
        self.file_path = file_path    
    
        try:
            try:
                df = pd.read_excel(file_path)
            except Exception as e1:
                try:
                    df = pd.read_excel(file_path, engine='openpyxl')
                except Exception as e2:
                    df = pd.read_excel(file_path, engine='xlrd')
        
            column_names = [col.lower().strip() for col in df.columns]
            client_code_variations = ['client code', 'clientcode', 'client_code', 'code', 'client id', 'clientid', 'client_id']
        
            code_column = None
            for variant in client_code_variations:
                if variant in column_names:
                    code_column = df.columns[column_names.index(variant)]
                    break
        
            if code_column is None:
                QMessageBox.warning(self, "Error", f"No client code column found. Available columns: {', '.join(df.columns)}")
                return
            
            codes = df[code_column].dropna().astype(str).tolist()

            if not codes:
                QMessageBox.warning(self, "Error", "No client codes found in the Excel file!")
                return

            QMessageBox.information(self, "Success", f"Loaded {len(codes)} client codes.")

            if not self.scraper:
                QMessageBox.warning(self, "Error", "Please login first before processing client codes.")
                return
        
            choice = QMessageBox.question(
                self, 
                "Download Holdings", 
                f"Do you want to download holdings for {len(codes)} clients?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes
            )
            
            if choice == QMessageBox.StandardButton.No:
                return
                
            self.status_lbl.setText("Downloading holdings...")
            holdings_success, holdings_fails = self.scraper.process_all_clients(codes, self.update_sum)
            
            mf_choice = QMessageBox.question(
                self, 
                "Download MF Transactions", 
                f"Holdings download completed ({holdings_success}/{len(codes)} successful).\n\nDo you want to download MF transactions?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes
            )
            
            mf_success = 0
            mf_fails = []
            
            if mf_choice == QMessageBox.StandardButton.Yes:
                from_date = None
                to_date = None
                
                if self.use_date_range and self.use_date_range.isChecked():
                    from_date = self.from_date.date().toString("dd/MM/yyyy")
                    to_date = self.to_date.date().toString("dd/MM/yyyy")
                    self.status_lbl.setText(f"Downloading MF transactions with date range: {from_date} to {to_date}...")
                else:
                    self.status_lbl.setText("Downloading MF transactions...")

                mf_success, mf_fails = self.scraper.process_all_clients_mf_trans(codes, self.update_sum)
            
            summary = []
            summary.append(f"Downloaded holdings: {holdings_success}/{len(codes)} clients")
            summary.append(f"Failed holdings: {len(holdings_fails)}")
            if holdings_fails:
                summary.append(f"Failed holdings clients: {', '.join(holdings_fails[:5])}" + 
                           ("..." if len(holdings_fails) > 5 else ""))
            
            if mf_choice == QMessageBox.StandardButton.Yes:
                summary.append(f"Downloaded MF transactions: {mf_success}/{len(codes)} clients")
                summary.append(f"Failed MF transactions: {len(mf_fails)}")
                if mf_fails:
                    summary.append(f"Failed MF transactions clients: {', '.join(mf_fails[:5])}" + 
                               ("..." if len(mf_fails) > 5 else ""))
        
            self.sum_lbl.setText("\n".join(summary))
        
            QMessageBox.information(self, "Download Complete", 
                f"Process completed with the following results:\n\n" + "\n".join(summary))
        
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Excel processing error: {error_details}")
            QMessageBox.critical(self, "Error", f"Failed to process file: {str(e)}\n\nCheck console for full error details.")
    
    def update_sum(self, success, total, fails):
        fail_txt = "\nFailed clients: " + ", ".join(fails) if fails else ""
        self.sum_lbl.setText(
            f"Downloaded holdings: {success}/{total} clients\n"
            f"Failed downloads: {len(fails)}{fail_txt}"
        )
    
    def process_hdng(self):
        folder = self.dl_folder  
        self.log(f"Processing holdings from: {folder}")

        try:
            excel_files = [os.path.join(folder, f) for f in os.listdir(folder) 
                          if f.endswith(('.xlsx', '.xls'))]
        
            if not excel_files:
                self.sum_lbl.setText("No Excel files found in Holdings folder.")
                QMessageBox.warning(self, "Error", "No Excel files found in Holdings folder.")
                return
            
        
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
        
            QMessageBox.information(self, "Success", 
                f"Holdings conversion completed!\n\n"
                f"Files converted to CSV: {converted_count}/{len(excel_files)}\n"
                f"Location: {folder}")
        
        
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
        
        except Exception as e:
            import traceback
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
        
            for folder in [holding_folder, ledger_folder, client_reports_folder]:
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
                
                    ledger_file = None
                    for ext in ['.csv', '.xlsx', '.xls']:
                        potential_file = os.path.join(ledger_folder, base_filename + ext)
                        if os.path.exists(potential_file):
                            ledger_file = potential_file
                            break
                
                    if not ledger_file:
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
                        
                            ledger_csv_file = os.path.join(ledger_folder, base_filename + '.csv')
                            ledger_df.to_csv(ledger_csv_file, index=False)
                            ledger_file = ledger_csv_file
                        
                        except Exception as e:
                            self.log(f"Failed to convert {ledger_file} to CSV: {str(e)}")
                            skipped_count += 1
                            continue
                
                    holding_df = pd.read_csv(os.path.join(holding_folder, holding_csv))
                    ledger_df = pd.read_csv(ledger_file)
                
                    report_content = report_gen(holding_df, ledger_df)
                
                    output_file = os.path.join(client_reports_folder, f"{base_filename}_report.pdf")
                
                    if isinstance(report_content, pd.DataFrame):
                        temp_excel = os.path.join(client_reports_folder, f"{base_filename}_temp.xlsx")
                        report_content.to_excel(temp_excel, index=False)
                    
                        output_file = os.path.join(client_reports_folder, f"{base_filename}_report.xlsx")
                        import shutil
                        shutil.move(temp_excel, output_file)
                    
                    elif isinstance(report_content, str) and os.path.exists(report_content):
                        import shutil
                        extension = os.path.splitext(report_content)[1]
                        output_file = os.path.join(client_reports_folder, f"{base_filename}_report{extension}")
                        shutil.copy2(report_content, output_file)
                
                    else:
                        with open(output_file, 'wb') as f:
                            if isinstance(report_content, bytes):
                                f.write(report_content)
                            else:
                                f.write(str(report_content).encode('utf-8'))
                
                    if os.path.exists(output_file):
                        processed_count += 1
                        self.log(f"Generated report for {base_filename}")
                    else:
                        self.log(f"Failed to generate report for {base_filename}")
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
            import traceback
            error_details = traceback.format_exc()
            print(f"Report generation error: {error_details}")
            QMessageBox.critical(self, "Error", f"Failed to generate reports: {str(e)}")

    def generate_excel(self):
        try:
            self.log("Generating Excel files from CSV files...")
            folder = self.dl_folder  
            desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
            excel_reports_folder = os.path.join(desktop, "excel_reports")
            if not os.path.exists(excel_reports_folder):
                os.makedirs(excel_reports_folder)
        
            csv_files = [os.path.join(folder, f) for f in os.listdir(folder) 
                        if f.endswith('.csv')]
        
            if not csv_files:
                self.log("No CSV files found in Holdings folder.")
                QMessageBox.warning(self, "Error", "No CSV files found in Holdings folder.")
                return
        
            processed_count = 0
            for csv_file in csv_files:
                try:
                    df = pd.read_csv(csv_file)
                    if df.empty:
                        print(f"Skipping empty file: {csv_file}")
                        continue
                
                    base_filename = os.path.splitext(os.path.basename(csv_file))[0]
                    
                    output_file = excel_generator(df)
                
                    if output_file:
                        if os.path.exists(output_file):
                            dest_file = os.path.join(excel_reports_folder, f"{base_filename}_report.xlsx")
                            import shutil
                            shutil.copy2(output_file, dest_file)
                            processed_count += 1
                            print(f"Processed: {csv_file} → {dest_file}")
                        else:
                            print(f"Output file not found: {output_file}")
                    else:
                        try:
                            potential_dirs = [os.getcwd(), os.path.dirname(csv_file), desktop]
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
                                import shutil
                                shutil.copy2(newest_file, dest_file)
                                processed_count += 1
                                print(f"Processed: {csv_file} → {dest_file} (found recent file)")
                            else:
                                print(f"Could not locate output file for {csv_file}")
                        except Exception as inner_e:
                            print(f"Error locating output for {csv_file}: {str(inner_e)}")
                    
                except Exception as e:
                    print(f"Error processing {csv_file}: {str(e)}")
        
            if processed_count > 0:
                self.log(f"Generated {processed_count}/{len(csv_files)} Excel reports in {excel_reports_folder}")
                QMessageBox.information(self, "Success", 
                    f"Excel reports successfully generated!\n\n"
                    f"Files processed: {processed_count}/{len(csv_files)}\n"
                    f"Reports location: {excel_reports_folder}")
            else:
                self.log("Failed to generate any Excel reports")
                QMessageBox.warning(self, "Error", "Failed to generate any Excel reports")
        
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Excel generation error: {error_details}")
            QMessageBox.critical(self, "Error", f"Failed to generate Excel files: {str(e)}")
     
    def gen_xirr(self):
        client_code = self.xirr_code_input.text().strip()
    
        try:
            init_val_text = self.init_portfolio_val_input.text().strip()
            init_val_text = init_val_text.replace('₹', '').replace(',', '')
            initial_value = float(init_val_text) if init_val_text else 100000
    
        except ValueError:
            print("Invalid initial portfolio value. Using default of 100000.")
            initial_value = 100000
    
        try:
            curr_val_text = self.cur_portfolio_val_input.text().strip()
            curr_val_text = curr_val_text.replace('₹', '').replace(',', '')
            current_value = float(curr_val_text) if curr_val_text else initial_value
    
        except ValueError:
            print("Invalid current portfolio value. Using same as initial value.")
            current_value = initial_value
    
        q_date = self.start_date.date()
        start_date = date(q_date.year(), q_date.month(), q_date.day())
    
        print("\nProcessing with the following parameters:")
        print(f"Client Code: {client_code if client_code else 'All clients'}")
        print(f"Initial Value: {initial_value}")
        print(f"Current Value: {current_value}")
        print(f"Start Date: {start_date.strftime('%d/%m/%Y')}")
    
        processing_msg = QMessageBox()
        processing_msg.setWindowTitle("Processing")
        processing_msg.setText("Generating XIRR calculation...")
        processing_msg.setStandardButtons(QMessageBox.NoButton)
        processing_msg.show()
    
        try:
            if client_code:
                result = proc(cl_code=client_code, init_val=initial_value, curr_val=current_value, start_date=start_date)
                processing_msg.close()
            
                if result:
                    QMessageBox.information(self, "Success", f"XIRR calculation complete.\nResults saved to: {result}")
                else:
                    QMessageBox.warning(self, "Error", "Failed to generate XIRR. Please check the files and inputs.")
            else:
                results = proc(init_val=initial_value, curr_val=current_value, start_date=start_date)
                processing_msg.close()
            
                if results:
                    success_text = "XIRR calculations complete. Results saved to:\n"
                    for res in results:
                        success_text += f"  - {res}\n"
                    QMessageBox.information(self, "Success", success_text)
                else:
                    QMessageBox.warning(self, "Error", "Failed to generate XIRR. Please check the files and inputs.")
        
            return result if client_code else results
        except Exception as e:
            processing_msg.close()
            error_message = f"Error generating XIRR: {e}"
            print(error_message)
            QMessageBox.critical(self, "Error", error_message)
            return None
     
    def closeEvent(self, event):
        if self.scraper:
            self.scraper.quit()
        event.accept()