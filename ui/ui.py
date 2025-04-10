import os
import pandas as pd
from generator.excel import excel_generator
from web.web import Scraper
from PyQt5.QtCore import Qt, QMimeData
from utils.processor import Processor
from PyQt5.QtWidgets import (
    QMainWindow, QPushButton, QVBoxLayout, QWidget,
    QFileDialog, QMessageBox, QLabel, QLineEdit, QHBoxLayout, QFrame, QGridLayout
)
from PyQt5.QtGui import QDragEnterEvent, QDropEvent


class FileDropZone(QFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.setAcceptDrops(True)
        self.setFrameShape(QFrame.Box)
        self.setFrameShadow(QFrame.Sunken)
        self.setStyleSheet("""
            background-color: #f0f0f0;
            border: 2px dashed #aaaaaa;
            border-radius: 5px;
            padding: 5px;
            min-height: 100px;
        """)
        
        layout = QVBoxLayout()
        
        self.title_label = QLabel("Upload Required Files")
        self.title_label.setStyleSheet("font-weight: bold; color: black;")
        layout.addWidget(self.title_label)
        
        self.status_label = QLabel("Drag & drop files here or click to select")
        self.status_label.setStyleSheet("color: black;")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.status_label)
        
        file_types_label = QLabel("Required: Ledger, MF Transactions, SIP")
        file_types_label.setStyleSheet("color: black;")
        file_types_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(file_types_label)
        
        self.setLayout(layout)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet("""
                background-color: #e0f7fa;
                border: 2px dashed #00acc1;
                border-radius: 5px;
                padding: 5px;
                min-height: 100px;
            """)
    
    def dragLeaveEvent(self, event):
        self.setStyleSheet("""
            background-color: #f0f0f0;
            border: 2px dashed #aaaaaa;
            border-radius: 5px;
            padding: 5px;
            min-height: 100px;
        """)
    
    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            file_paths = [url.toLocalFile() for url in event.mimeData().urls()]
            self.process_files(file_paths)
        
        self.setStyleSheet("""
            background-color: #f0f0f0;
            border: 2px dashed #aaaaaa;
            border-radius: 5px;
            padding: 5px;
            min-height: 100px;
        """)
    
    def mousePressEvent(self, event):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "Select Required Files", "", 
            "All Files (*);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv)"
        )
        if file_paths:
            self.process_files(file_paths)
    
    def process_files(self, file_paths):
        try:
            if hasattr(self.parent, 'process_uploaded_files'):
                self.parent.process_uploaded_files(file_paths)
            
            count = len(file_paths)
            if count > 0:
                self.status_label.setText(f"{count} file(s) uploaded")
                self.status_label.setStyleSheet("color: black; font-weight: bold;")
                
        except Exception as e:
            self.status_label.setText(f"Error: {str(e)}")
            self.status_label.setStyleSheet("color: black; font-weight: bold;")


class UploadedFilesDisplay(QFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.setFrameShape(QFrame.Box)
        self.setFrameShadow(QFrame.Sunken)
        self.setStyleSheet("""
            background-color: white;
            border: 1px solid #cccccc;
            border-radius: 5px;
            padding: 5px;
        """)
        
        self.layout = QGridLayout()
        self.layout.setColumnStretch(1, 1)  
        
        header_type = QLabel("File Type")
        header_type.setStyleSheet("font-weight: bold; color: black;")
        self.layout.addWidget(header_type, 0, 0)
        
        header_name = QLabel("File Name")
        header_name.setStyleSheet("font-weight: bold; color: black;")
        self.layout.addWidget(header_name, 0, 1)
        
        self.file_type_labels = {}
        self.file_name_labels = {}
        
        row = 1
        for file_type in ["Ledger", "MF Transactions", "SIP"]:
            type_label = QLabel(f"{file_type}:")
            type_label.setStyleSheet("color: black;")
            self.layout.addWidget(type_label, row, 0)
            
            name_label = QLabel("Not uploaded")
            name_label.setStyleSheet("color: black;")
            self.layout.addWidget(name_label, row, 1)
            
            self.file_type_labels[file_type] = type_label
            self.file_name_labels[file_type] = name_label
            
            row += 1
        
        self.setLayout(self.layout)
    
    def update_file(self, file_type, file_path):
        if file_type in self.file_name_labels:
            file_name = os.path.basename(file_path)
            self.file_name_labels[file_type].setText(file_name)
            self.file_name_labels[file_type].setStyleSheet("color: black; font-weight: bold;")


class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.scraper = None
        self.file_path = None
        self.processor = None
        self.dl_folder = self._get_dl_path()
        
        # File tracking
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

    def log(self, msg):
        print(msg)
        self.status_lbl.setText(msg)

    def init_ui(self):
        self.setWindowTitle("REPORT IQ")
        self.setGeometry(100, 100, 500, 600)
        self.setStyleSheet("""
            background-color: #f8f9fa;
            font-family: Arial;
            font-size: 12px;
        """)

        layout = QVBoxLayout()
        
        title = QLabel("REPORT IQ")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("""
            font-size: 20px;
            font-weight: bold;
            color: #1e3a8a;
            padding: 10px;
        """)
        layout.addWidget(title)

        url_lbl = QLabel("Enter URL:")
        url_lbl.setStyleSheet("font-weight: bold; color: black;")
        self.url_in = QLineEdit()
        self.url_in.setPlaceholderText("https://example.com")
        self.url_in.setStyleSheet("color: black;")
        layout.addWidget(url_lbl)
        layout.addWidget(self.url_in)

        user_lbl = QLabel("Enter Username:")
        user_lbl.setStyleSheet("font-weight: bold; color: black;")
        self.user_in = QLineEdit()
        self.user_in.setStyleSheet("color: black;")
        layout.addWidget(user_lbl)
        layout.addWidget(self.user_in)

        pass_lbl = QLabel("Enter Password:")
        pass_lbl.setStyleSheet("font-weight: bold; color: black;")
        self.pass_in = QLineEdit()
        self.pass_in.setEchoMode(QLineEdit.EchoMode.Password)
        self.pass_in.setStyleSheet("color: black;")
        layout.addWidget(pass_lbl)
        layout.addWidget(self.pass_in)

        login_btn = QPushButton("LOGIN")
        login_btn.setStyleSheet("""
            background-color: #2563eb;
            color: white;  
            font-weight: bold;
            padding: 3px 10px;
            border-radius: 5px;
        """)
        login_btn.clicked.connect(self.login)
        layout.addWidget(login_btn)

        self.status_lbl = QLabel("Status: Not logged in")
        self.status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_lbl.setStyleSheet("""
            font-weight: bold;
            color: black;
            padding: 5px;
        """)
        layout.addWidget(self.status_lbl)
        
        upload_title = QLabel("Required Files")
        upload_title.setStyleSheet("""
            font-size: 16px;
            font-weight: bold;
            color: #1e3a8a;
            padding: 5px;
        """)
        upload_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(upload_title)
        
        self.file_drop_zone = FileDropZone(self)
        layout.addWidget(self.file_drop_zone)
        
        uploaded_files_label = QLabel("Uploaded Files")
        uploaded_files_label.setStyleSheet("""
            font-size: 14px;
            font-weight: bold;
            color: black;
            padding: 5px;
        """)
        uploaded_files_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(uploaded_files_label)
        
        self.uploaded_files_display = UploadedFilesDisplay(self)
        layout.addWidget(self.uploaded_files_display)

        self.excel_btn = QPushButton("LOAD CLIENT EXCEL FILE")
        self.excel_btn.setStyleSheet("""
            background-color: #22c55e;
            color: white;
            font-weight: bold;
            padding: 3px 10px;
            border-radius: 5px;
        """)
        self.excel_btn.clicked.connect(self.open_excel)
        self.excel_btn.setEnabled(False)
        layout.addWidget(self.excel_btn)
        
        proc_btn = QPushButton("PROCESS HOLDINGS")
        proc_btn.setStyleSheet("""
            background-color: #f59e0b;
            color: white;
            font-weight: bold;
            padding: 3px 10px;
            border-radius: 5px;
        """)
        proc_btn.clicked.connect(self.process_hdng)
        layout.addWidget(proc_btn)

        proc_mf_btn = QPushButton("PROCESS MF TRANSACTIONS")
        proc_mf_btn.setStyleSheet("""
            background-color: #f59e0b;
            color: white;
            font-weight: bold;
            padding: 3px 10px;
            border-radius: 5px;
        """)
        proc_mf_btn.clicked.connect(self.process_mf_trans)
        layout.addWidget(proc_mf_btn)
        
        
        generate_report_btn = QPushButton("GENERATE REPORT")
        generate_report_btn.setStyleSheet("""
            background-color: #8b5cf6;
            color: white;
            font-weight: bold;
            padding: 3px 10px;
            border-radius: 5px;
            """)
        generate_report_btn.clicked.connect(self.generate_report)
        layout.addWidget(generate_report_btn)

        generate_excel_btn = QPushButton("GENERATE EXCEL")
        generate_excel_btn.setStyleSheet("""
            background-color: #ef4444;
            color: white;
            font-weight: bold;
            padding: 3px 10px;
            border-radius: 5px;
            """)
        generate_excel_btn.clicked.connect(self.generate_excel)
        layout.addWidget(generate_excel_btn)
        
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
        self.setCentralWidget(container)
    
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
            self.scraper = Scraper(self.dl_folder)
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
                
            success, fails = self.scraper.process_all_clients(codes, self.update_sum)
            
            QMessageBox.information(self, "Success", f"Extracted {success} of {len(codes)} holdings.")
            
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
        folder = QFileDialog.getExistingDirectory(self, "Select holdings folder", self.dl_folder)
    
        if not folder:
            QMessageBox.warning(self, "Error", "No folder selected.")
            return

        self.log(f"Processing holdings from: {folder}")

        try:
            self.processor = Processor(folder)
        
            if hasattr(self.processor, 'set_required_files') and "Ledger" in self.required_files and self.required_files["Ledger"] is not None:
                self.processor.set_required_files(
                    ledger=self.required_files["Ledger"],
                    mf_transactions=None,
                    sip=None
                )
        
            out_file = self.processor.run()

            if out_file:
                df = pd.read_excel(out_file)
                count = df.shape[0]

                self.sum_lbl.setText(
                    f"Extracted holdings for {count} clients.\n"
                    f"Report saved: {out_file}"
                )
                QMessageBox.information(self, "Success", 
                    f"Holdings processing completed!\n\n"
                    f"Clients processed: {count}\n"
                    f"Report saved: {out_file}")
            else:
                self.sum_lbl.setText("No valid holdings files found.")
                QMessageBox.warning(self, "Error", "No valid holdings files found.")
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
                    f"Processed MF transactions for {count} clients.\n"
                    f"Report saved: {out_file}"
                )
                QMessageBox.information(self, "Success", 
                    f"MF transactions processing completed!\n\n"
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
        
            if not hasattr(self, 'processor') or not self.processor:
                QMessageBox.warning(self, "Error", "Please process holdings first.")
                return
            
            report_path = QFileDialog.getSaveFileName(self, "Save Report As", "", "PDF Files (*.pdf);;All Files (*)")[0]
            if not report_path:
                return
            
            self.log(f"Report generated and saved to: {report_path}")
            QMessageBox.information(self, "Success", f"Report successfully generated and saved to:\n{report_path}")
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Report generation error: {error_details}")
            QMessageBox.critical(self, "Error", f"Failed to generate report: {str(e)}")

    def generate_excel(self):
        try:
            self.log("Generating Excel file...")
    
            if not hasattr(self, 'processor') or not self.processor:
                QMessageBox.warning(self, "Error", "Please process holdings first.")
                return
        
            excel_path = QFileDialog.getSaveFileName(self, "Save Excel As", "", "Excel Files (*.xlsx);;All Files (*)")[0]
            if not excel_path:
                return
        
            result = excel_generator()
        
            if result:
                self.log(f"Excel file generated and saved to: {excel_path}")
                QMessageBox.information(self, "Success", f"Excel file successfully generated and saved to:\n{excel_path}")
            else:
                self.log("Failed to generate Excel file")
                QMessageBox.warning(self, "Error", "Failed to generate Excel file")
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"Excel generation error: {error_details}")
            QMessageBox.critical(self, "Error", f"Failed to generate Excel file: {str(e)}")
     
    def closeEvent(self, event):
        if self.scraper:
            self.scraper.quit()
        event.accept()