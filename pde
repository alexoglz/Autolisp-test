import pandas as pd
import os
from tkinter import Tk, filedialog

def seleccionar_archivo():
    """Abre un cuadro de diálogo para seleccionar un archivo Excel"""
    root = Tk()
    root.withdraw()
    archivo = filedialog.askopenfilename(title="Selecciona un archivo Excel", filetypes=[("Excel Files", "*.xlsx")])
    return archivo

def extraer_codigos_faltantes(database_file, aero_file):
    # Leer database.xlsx
    df_db = pd.read_excel(database_file, engine="openpyxl")
    
    # Convertir fechas de la primera columna a formato real de fecha
    df_db.iloc[:, 0] = pd.to_datetime(df_db.iloc[:, 0], format="%b %d, %Y", errors='coerce')

    # Filtrar fechas dentro de las próximas dos semanas
    fecha_actual = pd.to_datetime("today")
    fecha_limite = fecha_actual + pd.Timedelta(days=14)
    df_db = df_db[(df_db.iloc[:, 0] >= fecha_actual) & (df_db.iloc[:, 0] <= fecha_limite)]
    
    # Extraer códigos antes de los dos puntos en la columna E
    df_db["Codigo"] = df_db.iloc[:, 4].astype(str).str.split(":").str[0].str.strip()
    
    # Lista de códigos únicos
    codigos_database = df_db["Codigo"].unique()

    # Leer Aero.xlsx
    xls_aero = pd.ExcelFile(aero_file, engine="openpyxl")

    # Leer las hojas "2D activities" y "3D activities"
    df_2d = pd.read_excel(xls_aero, sheet_name="2D activities", usecols=["L"], engine="openpyxl")
    df_3d = pd.read_excel(xls_aero, sheet_name="3D activities", usecols=["H"], engine="openpyxl")

    # Combinar todas las columnas de Aero en una lista de códigos
    codigos_aero = set(df_2d.iloc[:, 0].dropna().astype(str)) | set(df_3d.iloc[:, 0].dropna().astype(str))

    # Encontrar códigos que NO están en Aero
    codigos_faltantes = [codigo for codigo in codigos_database if codigo not in codigos_aero]

    # Guardar en un nuevo archivo Excel
    output_file = "Missing_Codes.xlsx"
    pd.DataFrame({"Codigos Faltantes": codigos_faltantes}).to_excel(output_file, index=False)

    print(f"Proceso completado. Se generó el archivo: {output_file}")

if __name__ == "__main__":
    print("Selecciona el archivo database.xlsx")
    database_file = seleccionar_archivo()

    print("Selecciona el archivo Aero.xlsx")
    aero_file = seleccionar_archivo()

    extraer_codigos_faltantes(database_file, aero_file)
