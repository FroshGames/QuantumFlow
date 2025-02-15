from PyQt5.QtWidgets import (
    QMainWindow, QFileDialog, QTableWidget, QTableWidgetItem, QPushButton, 
    QVBoxLayout, QWidget, QHBoxLayout, QMessageBox
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt
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
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: #f9f9f9;
                gridline-color: #bbb;
                border: 1px solid #ccc;
            }
            QHeaderView::section {
                background-color: #4CAF50;
                color: white;
                font-size: 14px;
                font-weight: bold;
                padding: 5px;
            }
            QTableWidget::item {
                padding: 8px;
            }
        """)

        # Ocultar numeración predeterminada en el índice vertical
        self.table.verticalHeader().setVisible(False)

        # Ajuste automático del tamaño de las columnas
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setDefaultSectionSize(100)
        self.table.verticalHeader().setDefaultSectionSize(30)

        # Alternar colores de filas para mejor legibilidad
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("alternate-background-color: #E6F7FF; background-color: white;")

        # Centrar el texto en las celdas
        self.table.setStyleSheet("""
            QTableWidget::item {
                text-align: center;
            }
        """)

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
            btn.setStyleSheet("""
                padding: 10px; 
                font-size: 14px; 
                background-color: #4CAF50; 
                color: white; 
                border-radius: 5px;
            """)
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
        """Abrir archivo Excel con manejo de errores"""
        file_path, _ = QFileDialog.getOpenFileName(self, "Abrir Archivo Excel", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.excel_handler.load_excel(file_path)
            self.load_data()

    def save_excel(self):
        """Guardar los datos en un archivo Excel"""
        file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Archivo", "", "Excel Files (*.xlsx)")
        if file_path:
            self.get_table_data()  # Capturar los datos antes de guardarlos
            self.excel_handler.save_excel(file_path)
            QMessageBox.information(self, "Guardado", "El archivo se ha guardado correctamente.")

    def new_excel(self):
        """Crear un nuevo archivo Excel con datos aleatorios y abrirlo automáticamente"""
        file_path = self.excel_handler.new_excel()
        self.load_data()  # Asegurar que se carguen los datos en la tabla
        QMessageBox.information(self, "Nuevo Archivo", f"Se ha creado y abierto un nuevo archivo con datos aleatorios: {file_path}")

    def load_data(self):
        """Cargar los datos del Excel en la interfaz con numeración lateral pero sin contenido visible"""
        data = self.excel_handler.get_data()

        if data is None or data.empty:
            self.table.setRowCount(1)  # Al menos una fila para evitar errores
            self.table.setColumnCount(2)  # Incluye columna de numeración
            self.table.setHorizontalHeaderLabels(["#", "Columna1"])
            self.table.setItem(0, 1, QTableWidgetItem(""))  # Celda vacía
            return  

        self.table.setRowCount(data.shape[0])
        self.table.setColumnCount(data.shape[1] + 1)  # +1 para la columna de numeración
        self.table.setHorizontalHeaderLabels(["#"] + list(data.columns))

        # Agregar numeración lateral
        for row in range(data.shape[0]):
            num_item = QTableWidgetItem(str(row + 1))  # Número de fila
            num_item.setBackground(Qt.gray)  # Hacer que tenga otro color
            num_item.setTextAlignment(Qt.AlignCenter)  # Centrar texto
            self.table.setItem(row, 0, num_item)  # Agregar número de fila

            for col in range(data.shape[1]):
                empty_item = QTableWidgetItem("")  # Celdas vacías
                self.table.setItem(row, col + 1, empty_item)

        # Ajuste de la primera columna de numeración
        self.table.setColumnWidth(0, 50)  # Espacio para la numeración
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)  # Bloquear edición en numeración

    def get_table_data(self):
        """Actualizar el DataFrame con la información de la tabla sin afectar la numeración lateral"""
        for row in range(self.table.rowCount()):
            for col in range(1, self.table.columnCount()):  # Empezamos en 1 para ignorar la numeración
                item = self.table.item(row, col)
                self.excel_handler.data.iloc[row, col - 1] = item.text() if item else ""
