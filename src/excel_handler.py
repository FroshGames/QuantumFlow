import pandas as pd
import os
import random
import numpy as np

class ExcelHandler:
    def __init__(self):
        self.data = pd.DataFrame()
        self.data_folder = "data/"
        if not os.path.exists(self.data_folder):
            os.makedirs(self.data_folder)

    def load_excel(self, file_path):
        """Carga un archivo Excel y maneja errores si está vacío o tiene problemas"""
        try:
            self.data = pd.read_excel(file_path)
            if self.data.empty:
                self.data = pd.DataFrame(columns=["Columna1"])  # Agregar una columna por defecto
        except Exception as e:
            print(f"Error al cargar el archivo: {e}")
            self.data = pd.DataFrame(columns=["Columna1"])  # Crear un DataFrame por defecto

    def save_excel(self, file_path):
        """Guarda el archivo Excel sin índice"""
        self.data.to_excel(file_path, index=False)

    def new_excel(self):
        """Crea un nuevo archivo Excel con datos aleatorios y actualiza self.data"""
        if not os.path.exists(self.data_folder):
            os.makedirs(self.data_folder)

        # Generar nombre único
        file_number = 1
        while True:
            file_path = os.path.join(self.data_folder, f"nuevo_archivo_{file_number:03d}.xlsx")
            if not os.path.exists(file_path):
                break
            file_number += 1

        # Generar tamaño aleatorio entre 10 y 60
        num_rows = random.randint(10, 60)
        num_cols = random.randint(10, 60)

        # Generar nombres de columnas (Ej: "Columna1", "Columna2", ..., "ColumnaN")
        column_names = [f"Columna{i+1}" for i in range(num_cols)]

        # Generar datos aleatorios (números entre 1 y 1000)
        new_data = pd.DataFrame(np.random.randint(1, 1000, size=(num_rows, num_cols)), columns=column_names)

        # Guardar archivo con datos aleatorios
        new_data.to_excel(file_path, index=False)

        # Actualizar self.data
        self.data = new_data  

        return file_path  # Retorna la ruta del archivo creado

    def get_data(self):
        """Retorna el DataFrame actual"""
        return self.data
