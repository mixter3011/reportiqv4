import os
from web.web import Scraper
from PyQt5.QtCore import Qt
from utils.processor import Processor
from PyQt5.QtWidgets import (
    QMainWindow, QPushButton, QVBoxLayout, QWidget,
    QFileDialog, QMessageBox, QLabel, QLineEdit
)



class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.scraper = None
        self.file_path = None
        self.init_ui()
        self.dl_folder = self._get_dl_path()
        
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
        self.setGeometry(100, 100, 500, 450)
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
        url_lbl.setStyleSheet("font-weight: bold; color: #2c3e50;")
        self.url_in = QLineEdit()
        self.url_in.setPlaceholderText("https://example.com")
        layout.addWidget(url_lbl)
        layout.addWidget(self.url_in)

        user_lbl = QLabel("Enter Username:")
        user_lbl.setStyleSheet("font-weight: bold; color: #2c3e50;")
        self.user_in = QLineEdit()
        layout.addWidget(user_lbl)
        layout.addWidget(self.user_in)

        pass_lbl = QLabel("Enter Password:")
        pass_lbl.setStyleSheet("font-weight: bold; color: #2c3e50;")
        self.pass_in = QLineEdit()
        self.pass_in.setEchoMode(QLineEdit.EchoMode.Password)
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
            color: #34495e;
            padding: 5px;
        """)
        layout.addWidget(self.status_lbl)

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

        self.sum_lbl = QLabel("Summary: No processing yet.")
        self.sum_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.sum_lbl.setStyleSheet("""
            font-weight: bold;
            color: #475569;
            padding: 8px;
        """)
        layout.addWidget(self.sum_lbl)
        
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def login(self):
        url = self.url_in.text().strip()
        user = self.user_in.text().strip()
        pwd = self.pass_in.text().strip()

        if not url or not user or not pwd:
            QMessageBox.warning(self, "error", "please fill in all fields")
            return

        try:
            self.scraper = Scraper(self.dl_folder)
            if self.scraper.login(url, user, pwd):
                self.status_lbl.setText("login successful")
                self.excel_btn.setEnabled(True)
            else:
                self.status_lbl.setText("login failed")
                QMessageBox.critical(self, "login error", "failed to log in")
        except Exception as e:
            self.status_lbl.setText("login failed")
            QMessageBox.critical(self, "login error", f"failed to log in: {e}")

    def open_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "open excel file", "", "excel files (*.xlsx *.xls);;all files (*)")
        if not file_path:
            return
            
        import pandas as pd
        try:
            df = pd.read_excel(file_path)
            codes = df['client code'].dropna().tolist()

            if not codes:
                QMessageBox.warning(self, "error", "no client codes found in the excel file!")
                return

            QMessageBox.information(self, "success", f"loaded {len(codes)} client codes.")
            
            success, fails = self.scraper.process_all_clients(codes, self.update_sum)
            
            QMessageBox.information(self, "success", f"extracted {success} of {len(codes)} holdings.")
            
        except Exception as e:
            QMessageBox.critical(self, "error", f"failed to process file: {e}")
    
    def update_sum(self, success, total, fails):
        fail_txt = "\nfailed clients: " + ", ".join(fails) if fails else ""
        self.sum_lbl.setText(
            f"downloaded holdings: {success}/{total} clients\n"
            f"failed downloads: {len(fails)}{fail_txt}"
        )
    
    def process_hdng(self):
        default_path = self.dl_folder
        folder = QFileDialog.getExistingDirectory(self, "select holdings folder", default_path)
        
        if not folder:
            QMessageBox.warning(self, "error", "no folder selected.")
            return

        self.log(f"processing holdings from: {folder}")

        try:
            proc = Processor(folder)
            out_file = proc.run()

            if out_file:
                import pandas as pd
                df = pd.read_excel(out_file)
                count = df.shape[0]

                self.sum_lbl.setText(
                    f"extracted holdings for {count} clients.\n"
                    f"report saved: {out_file}"
                )
                QMessageBox.information(self, "success", 
                    f"holdings processing completed!\n\n"
                    f"clients processed: {count}\n"
                    f"report saved: {out_file}")
            else:
                self.sum_lbl.setText("no valid holdings files found.")
                QMessageBox.warning(self, "error", "no valid holdings files found.")
        except Exception as e:
            err_msg = f"Error: {str(e)}"
            self.log(err_msg)
            QMessageBox.critical(self, "critical error", err_msg)

    def closeEvent(self, event):
        if self.scraper:
            self.scraper.quit()
        event.accept()