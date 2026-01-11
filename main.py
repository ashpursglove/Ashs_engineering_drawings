"""
main.py

Entry point for Ash's Engineering Drawing Maker.
"""

import sys
from PyQt5 import QtWidgets

from gui import EngineeringDrawingMaker


def main() -> None:
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName("Engineering Drawing Maker")
    app.setStyle("Fusion")

    w = EngineeringDrawingMaker()
    w.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
