Attribute VB_Name = "Main"
Option Explicit
Sub proceso_ETL()
Dim Libro As Workbook
Dim hoja_DB As Worksheet
Dim iHoja As Worksheet
Dim hoja_MD As Worksheet
Dim obtener_ruta As String, rango_MD As String, year As String
Dim dia_mes_lim As String, unir As String
Dim ultima_fila_MD As Long, altura As Long, base As Long, fila As Long, colm As Long
Dim ultima_fila_DB As Long
Dim array_MD As Variant, datos_limpios As Variant
Dim CLASI As Long, TIPO As Long, DETALLE As Long



Set Libro = ThisWorkbook
Set hoja_DB = Libro.Worksheets("BASE DATOS")
Application.ScreenUpdating = False

'Capturamos la ruta del archivo
obtener_ruta = ruta_archivo()
If obtener_ruta = "0" Then
   GoTo FIN
End If

'Abrimos el Archivo
Set Libro = Workbooks.Open(obtener_ruta, 0)

'Recorremos las Hojas del Libro A6:L84
For Each iHoja In Libro.Sheets
    If iHoja.Name = "Original" Then
       Set hoja_MD = Libro.Worksheets(iHoja.Name)
           hoja_MD.Activate
           
           ultima_fila_MD = hoja_MD.Cells(Rows.Count, 2).End(xlUp).Row
           rango_MD = "A6" & ":" & "L" & ultima_fila_MD
           array_MD = hoja_MD.Range(rango_MD).Value
           
           altura = UBound(array_MD, 1) - LBound(array_MD, 1) + 1
           base = UBound(array_MD, 2) - LBound(array_MD, 2) + 1
           
           year = array_MD(1, 1)
           
           For colm = 2 To base

               dia_mes_lim = limpiar_fecha(array_MD(1, colm))
               unir = dia_mes_lim & "/" & year
               
           
               For fila = 2 To altura
               
                   ultima_fila_DB = hoja_DB.Cells(Rows.Count, 1).End(xlUp).Row
                   
                   datos_limpios = orden_balance(array_MD(fila, 1))
                   
    
                   
                   Select Case datos_limpios(3)
                   
                          Case "total"
                                GoTo TOTAL
                                
                          Case "detalle"
                                hoja_DB.Cells(ultima_fila_DB + 1, 1) = Format(CDate(unir), "dd/mm/yyyy")       'fecha
                                hoja_DB.Cells(ultima_fila_DB + 1, 2) = datos_limpios(1)       'clasificacion
                                hoja_DB.Cells(ultima_fila_DB + 1, 3) = datos_limpios(2)       'tipo
                                hoja_DB.Cells(ultima_fila_DB + 1, 4) = datos_limpios(0)       'detalle
                                hoja_DB.Cells(ultima_fila_DB + 1, 5) = array_MD(fila, colm)   'valor
                                
                                CLASI = id_clasificacion(datos_limpios(1))
                                hoja_DB.Cells(ultima_fila_DB + 1, 6) = CLASI    'ID_CLASIFICACION
                                
                                TIPO = id_tipo(datos_limpios(2))
                                hoja_DB.Cells(ultima_fila_DB + 1, 7) = TIPO      'ID_TIPO
                                
                                DETALLE = id_detalle(datos_limpios(0))
                                hoja_DB.Cells(ultima_fila_DB + 1, 8) = DETALLE   'ID_DETALLE
                                
                                Case Else
                                MsgBox "Valor inesperado: " & datos_limpios(3)
                                GoTo FIN
                          End Select
TOTAL:
               Next fila
               
           Next colm
           
    End If
           
Next iHoja
FIN:
End Sub
