import pandas as pd
import sys

# Recibe un DataFrame y devuelve (altura, base)

def obtener_dimensiones(df_recibido: pd.DataFrame) -> tuple[int, int]:
    
    altura, base = df_recibido.shape # la dimencion desde 1
    return altura , base

# Buscamos los valores limpios en tabla auxiliar 

def buscar_aux_balance(df_recibido: pd.DataFrame, item: str) -> list[str]:
    # declar variables
    indices: list[int] = [] # índice donde el valor es True
    fila : int = ""
    nombre_lim: str = ""
    clasificacion: str = ""
    tipo: str = ""
    orden_clasi: str = ""
    valores : list[str] = []

    
    # buscamo el item                                  
    indices = df_recibido.index[df_recibido.iloc[:, 0] == item] # [:, 0] fila y columna
    
    if indices.empty:
        print(f"❌ No se encontró el item '{item}'")
        sys.exit()
    else:
        fila = indices[0]
        nombre_lim = str(df_recibido.iloc[fila,1])
        clasificacion = str(df_recibido.iloc[fila,2])
        tipo = str(df_recibido.iloc[fila,3])
        orden_clasi = str(df_recibido.iloc[fila,4])

        valores = [nombre_lim, clasificacion, tipo, orden_clasi]

    return valores

# limpiamos y unimos dia + mes + anio

def limpiar_fecha(year: str, dia_mes: str) -> str:
    partes: list[str] = []
    dia: str = ""
    mes_letr: str = ""
    mes: str = ""
    anio: str = ""
    tot_fecha: str = ""

    if "-" not in dia_mes:
        print("❌ Error: el texto no contiene el separador '-'")
        sys.exit()  # Finaliza el programa inmediatamente

    partes = dia_mes.split("-") # Dic - 14

    mes_letr = partes[0].strip()

    if  mes_letr == "Dic":
        mes = "12"
    else:
        print("❌ Error: el texto es distinto de Dic")
        sys.exit()  # Finaliza el programa inmediatamente
    
    dia = partes[1].strip() # limpia los espacios
    anio = year.strip()

    tot_fecha = dia + "/" + mes + "/" + anio
 
    return tot_fecha

# Buscamos los ID en las tablas

def buscar_id(df_recibido: pd.DataFrame, item: str) -> int:
    # declar variables
    indices: list[int] = [] # índice donde el valor es True
    fila: int = 0
    id: int = 0

    
    # buscamo el item                                  
    indices = df_recibido.index[df_recibido.iloc[:, 1] == item] # [:, 0] fila y columna
    
    if indices.empty:
        print(f"❌ No se encontró el item '{item}'")
        sys.exit()
    else:
        fila = indices[0]
        id = (df_recibido.iloc[fila,0])

    return id