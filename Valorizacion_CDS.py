#VALORIZACION DE UNA CDS 

# Lectura de Librerías
import pandas as pd
import numpy as np
from scipy.optimize import curve_fit
import os

# Entrada de base de datos
base_dir = os.getcwd()

archivo_bonos = os.path.join(base_dir, "Entrada", "Data TF Derivados.xlsx")  # Lectura del archivo

# --- Aquí mantienes tu lectura EXACTA, sin modificar ---
df = pd.read_excel(archivo_bonos, sheet_name='Tasas US Trea %', skiprows=1)

Plazos = {
    "1M" : 1,
    "3M" : 3,
    "6M" : 6,
    "12M" : 12,
    "2YR" : 24,
    "3YR" : 36,
    "5YR" : 60,
    "7YR" : 84,
    "10YR" : 120,
    "20YR" : 240
}

# Funciones auxiliares

def extraer_plazo(col, dicc_plazos):
    for clave in dicc_plazos:
        if clave in col:
            return dicc_plazos[clave]
    return None

def nss(t, beta0, beta1, beta2, beta3, tau1, tau2):
    term1 = (1 - np.exp(-t / tau1)) / (t / tau1)
    term2 = term1 - np.exp(-t / tau1)
    term3 = (1 - np.exp(-t / tau2)) / (t / tau2) - np.exp(-t / tau2)
    return beta0 + beta1 * term1 + beta2 * term2 + beta3 * term3

def preparar_columnas_plazos(columnas_usgg, dicc_plazos):
    col_plazo = [(col, extraer_plazo(col, dicc_plazos)) for col in columnas_usgg]
    col_plazo = [(c, p) for c, p in col_plazo if p is not None]
    if not col_plazo:
        raise ValueError("No se encontraron columnas válidas para ajustar.")
    cols_validas = [c for c, _ in col_plazo]
    plazos_meses = np.array([p for _, p in col_plazo])
    plazos_anos = plazos_meses / 12
    return cols_validas, plazos_meses, plazos_anos, col_plazo

def ajustar_e_interpolar(df, cols_validas, plazos_meses, plazos_anos, col_plazo):
    plazos_objetivo_meses = np.arange(1, 241)
    plazos_objetivo_anos = plazos_objetivo_meses / 12
    resultados = []

    for idx, fila in df.iterrows():
        tasas_observadas = fila[cols_validas].values.astype(float)

        if np.isnan(tasas_observadas).any():
            print(f"Fila {idx} contiene datos faltantes. Resultado con NaN.")
            resultados.append(np.full(len(plazos_objetivo_anos), np.nan))
            continue

        p0 = [0.03, -0.02, 0.02, 0.01, 1.0, 3.0]

        try:
            params_opt, _ = curve_fit(nss, plazos_anos, tasas_observadas, p0=p0, maxfev=20000)
            tasas_estimadas = nss(plazos_objetivo_anos, *params_opt)
        except Exception as e:
            print(f"Error en fila {idx}: {e}")
            tasas_estimadas = np.full(len(plazos_objetivo_anos), np.nan)

        # Mantener valores originales en sus meses correspondientes
        for col, mes in col_plazo:
            idx_mes = mes - 1
            tasas_estimadas[idx_mes] = fila[col]

        resultados.append(tasas_estimadas)

    nombres_columnas = [f"{m}M_NSS" for m in plazos_objetivo_meses]
    df_interpolado = pd.DataFrame(resultados, columns=nombres_columnas, index=df.index)

    # Insertar columna de fechas si existe
    if "Dates" in df.columns:
        df_interpolado.insert(0, "Dates", df["Dates"].values)
    elif "Date" in df.columns:
        df_interpolado.insert(0, "Date", df["Date"].values)

    return df_interpolado

def exportar_excel(df_original, df_interpolado, ruta_salida):
    with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
        df_original.to_excel(writer, sheet_name="Datos_Originales", index=False)
        df_interpolado.to_excel(writer, sheet_name="Interpolados_NSS", index=False)
    print(f"Archivo exportado a: {ruta_salida}")

# --- EJECUCIÓN ---

columnas_usgg = [col for col in df.columns if "USGG" in col]
cols_validas, plazos_meses, plazos_anos, col_plazo = preparar_columnas_plazos(columnas_usgg, Plazos)

df_nss = ajustar_e_interpolar(df, cols_validas, plazos_meses, plazos_anos, col_plazo)

ruta_salida = os.path.join(base_dir, "Salida", "Resultados_NSS.xlsx")
exportar_excel(df, df_nss, ruta_salida)
