import sys
from ui.ui import Main
from PyQt5.QtWidgets import QApplication


def main():
    app = QApplication(sys.argv)
    window = Main()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()