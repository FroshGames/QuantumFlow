from PyQt5.QtWidgets import (
    QMainWindow, QTableWidget, QTableWidgetItem, QPushButton, 
    QVBoxLayout, QWidget, QHBoxLayout, QMessageBox
)
from PyQt5.QtGui import QFont
from excel_handler import ExcelHandler

class ExcelApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Mini Excel con PyQt5 y Pandas")
        self.setGeometry(100, 100, 900, 700)

        self.excel_handler = ExcelHandler()
        self.init_ui()

    def init_ui(self):
        self.table = QTableWidget()
        self.table.setFont(QFont("Arial", 10))
        self.table.setStyleSheet("background-color: #f5f5f5; gridline-color: #ccc;")
        self.table.verticalHeader().setVisible(False)

        # Botones
        self.btn_open = QPushButton("📂 Abrir Excel")
        self.btn_open.clicked.connect(self.open_excel)

        self.btn_save = QPushButton("💾 Guardar Excel")
        self.btn_save.clicked.connect(self.save_excel)

        self.btn_new = QPushButton("🆕 Nuevo Archivo")
        self.btn_new.clicked.connect(self.new_excel)

        # Estilizar botones
        buttons = [self.btn_open, self.btn_save, self.btn_new]
        for btn in buttons:
            btn.setStyleSheet("padding: 10px; font-size: 14px; background-color: #4CAF50; color: white; border-radius: 5px;")
            btn.setFont(QFont("Arial", 12))

        # Layouts
        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.btn_open)
        btn_layout.addWidget(self.btn_save)
        btn_layout.addWidget(self.btn_new)

        main_layout = QVBoxLayout()
        main_layout.addLayout(btn_layout)
        main_layout.addWidget(self.table)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def open_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Abrir Archivo Excel", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.excel_handler.load_excel(file_path)
            self.load_data()

    def save_excel(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Archivo", "", "Excel Files (*.xlsx)")
        if file_path:
            self.get_table_data()
            self.excel_handler.save_excel(file_path)
            QMessageBox.information(self, "Guardado", "El archivo se ha guardado correctamente.")

    def new_excel(self):
        file_path = self.excel_handler.new_excel()
        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        QMessageBox.information(self, "Nuevo Archivo", f"Se ha creado un nuevo archivo en: {file_path}")

    def load_data(self):
        data = self.excel_handler.get_data()
        if not data.empty:
            self.table.setRowCount(data.shape[0])
            self.table.setColumnCount(data.shape[1])
            self.table.setHorizontalHeaderLabels(data.columns)

            for row in range(data.shape[0]):
                for col in range(data.shape[1]):
                    self.table.setItem(row, col, QTableWidgetItem(str(data.iloc[row, col])))

    def get_table_data(self):
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                self.excel_handler.data.iloc[row, col] = item.text() if item else ""
