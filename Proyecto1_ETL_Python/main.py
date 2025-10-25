from openpyxl import load_workbook                 # Para trabajar con Excel
from openpyxl.workbook.workbook import Workbook    # Tipo del libro
from extraer import capturar_ruta
from extraer import cargar_datos_extra
from extraer import cargar_auxiliar_balance
from extraer import cargar_clasi
from extraer import cargar_tipo
from extraer import cargar_detalle
from transformar import obtener_dimensiones
from transformar import buscar_aux_balance
from transformar import limpiar_fecha
from cargar import registrar
from cargar import formato_tabla
from cargar import guardar_libro
import pandas as pd

def main() -> None:
 ruta : str = ""
 nombre_valor : str = ""
 clasi : str = ""
 fecha_stri : str = ""
 valores_limp: list[str] = []
 valor_num : float = 0
 i : int = 0
 confirmar: bool
 df_extraido: pd.DataFrame = None
 df_auxiliar: pd.DataFrame = None
 df_clasi:pd.DataFrame = None
 df_tipo:pd.DataFrame = None
 df_detalle:pd.DataFrame = None

 # recibimos la ruta del archivo de se va extraer los datos
 ruta_extraer = capturar_ruta(0)

 # recibimos la ruta del archivo de se va cargar los datos
 ruta_cargar = capturar_ruta(1)
 libro_cargar : Workbook = load_workbook(ruta_cargar,data_only=True)

 #cargamos los datos al dataframe
 df_extraido = cargar_datos_extra(ruta_extraer)
 df_auxiliar = cargar_auxiliar_balance (ruta_cargar)

 df_clasi = cargar_clasi(ruta_cargar)
 df_tipo = cargar_tipo(ruta_cargar)
 df_detalle = cargar_detalle(ruta_cargar)


 print(df_auxiliar)
 print(df_extraido)

 #calculamos la dimenciones del dataframe
 altura, base = obtener_dimensiones(df_extraido)
 

 # inicio recorrer columnas ; range empieza desde 0 y no incluye al 79 y el 12
 for colm in range(1,base):
    fecha_stri = limpiar_fecha(str(df_extraido.iloc[0,0]),str(df_extraido.iloc[0,colm])) # año, fecha columna
   # inicio recorrer fila
    for fila in range(1, altura):
       
       print(f"fila {fila} -- columna {colm}")      

       nombre_valor = df_extraido.iloc[fila,0]
       valores_limp = buscar_aux_balance(df_auxiliar,nombre_valor)
       clasi = valores_limp[3]

       match clasi:
          case "total":
             continue
          case "detalle":
             valor_num = df_extraido.iloc[fila,colm]
             registrar(libro_cargar, fecha_stri, valores_limp, valor_num, df_clasi, df_tipo, df_detalle)
             i+= 1
          case _: 
             print(f"Valor inesperado: {nombre_valor}") # captura cualquier valor inesperado
          
      # fin recorrer fila
 # fin recorrer columnas
 formato_tabla(libro_cargar)
 confirmar = guardar_libro(ruta_cargar,libro_cargar, i)

 if confirmar == False:
    print("Se omitió el guardado final por error.")
 else:
    print("Se guardo exitosamente.")
       
 return
 
# El condicional siempre va afuera
if __name__ == "__main__":
    main()