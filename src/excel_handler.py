import pandas as pd
import os

class ExcelHandler:
    def __init__(self):
        self.data = pd.DataFrame()
        self.data_folder = "data/"
        if not os.path.exists(self.data_folder):
            os.makedirs(self.data_folder)

    def load_excel(self, file_path):
        self.data = pd.read_excel(file_path)

    def save_excel(self, file_path):
        self.data.to_excel(file_path, index=False)

    def new_excel(self):
        """Crea y guarda una nueva hoja de cálculo vacía en la carpeta 'data/'."""
        if not os.path.exists(self.data_folder):
            os.makedirs(self.data_folder)

        # Generar nombre único
        file_number = 1
        while True:
            file_path = os.path.join(self.data_folder, f"nuevo_archivo_{file_number:03d}.xlsx")
            if not os.path.exists(file_path):
                break
            file_number += 1

        self.data = pd.DataFrame()
        self.data.to_excel(file_path, index=False)
        return file_path  # Retornar el nombre del archivo creado
