import sys
import os

# Add src directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

import database as db
from app_qt import MainWindow

os.makedirs(db.get_userdata_dir(), exist_ok=True)
db.init_db()

QApplication.setHighDpiScaleFactorRoundingPolicy(
    Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("备案配方表助手")
    app.setOrganizationName("CosmeticHelper")
    app.setFont(QFont("Microsoft YaHei", 12))

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == '__main__':
    main()
