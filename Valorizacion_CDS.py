#VALORIZACION DE UNA CDS 

# Lectura de Librerías
import pandas as pd
from openpyxl import load_workbook
import os

# Entrada de base de datos
base_dir = os.getcwd()

archivo_bonos = os.path.join(base_dir, "Entrada", "Data TF Derivados.xlsx")