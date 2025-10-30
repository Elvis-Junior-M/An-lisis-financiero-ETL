Attribute VB_Name = "Funciones"
'======================================================================
'Busca clasificaciones,tipo,subtipo y detalle
'Author: Elvis Junior Mamani Mendez
'Email: Elvisjuniorm21@gmail.com
'======================================================================
Function orden_balance(ByRef item) As Variant
Dim Libro As Workbook
Dim HojaAux As Worksheet
Dim valor As Range
Dim ruta, rango, nombre_lim As String, clasificacion As String, orden_clasi As String, TIPO As String
Dim ultimoValorAux, Nro As Long

Set Libro = ThisWorkbook
Set HojaAux = Libro.Worksheets("Auxiliar Balance")

ultimoValorAux = HojaAux.Cells(Rows.Count, 1).End(xlUp).Row
rango = "A1:A" & ultimoValorAux
    With HojaAux.Range(rango)
        Set valor = .Find(item, LookIn:=xlValues, LookAt:=xlWhole)
        If Not valor Is Nothing Then
            ruta = valor.Address
            Nro = Right(ruta, Len(ruta) - 3)
            
            nombre_lim = HojaAux.Cells(Nro, 2).Value
            clasificacion = HojaAux.Cells(Nro, 3).Value
            TIPO = HojaAux.Cells(Nro, 4).Value
            orden_clasi = HojaAux.Cells(Nro, 5).Value
            
           orden_balance = Array(nombre_lim, clasificacion, TIPO, orden_clasi)
           
        End If
    End With
End Function
'======================================================================
'Busca Solo Primera Mayuscula
'Author: Elvis Junior Mamani Mendez
'Email: Elvisjuniorm21@gmail.com
'======================================================================
Sub SoloPrimeraMayuscula()
    Dim celda As Range
    Dim texto As String
    
    For Each celda In Range("C1:C78")
        If Not IsEmpty(celda.Value) Then
            texto = LCase(Trim(celda.Value)) ' todo en minúscula y quita espacios extra
            ' primera letra en mayúscula
            texto = UCase(Left(texto, 1)) & Mid(texto, 2)
            celda.Value = texto
        End If
    Next celda
End Sub
'======================================================================
'limpiamos la fecha
'Author: Elvis Junior Mamani Mendez
'Email: Elvisjuniorm21@gmail.com
'======================================================================
Function limpiar_fecha(ByRef item) As String
Dim fecha As String
Dim partes() As String
Dim mes As String
Dim ano As String
    
    ' Separar usando el guion como separador
    partes = Split(item, "-")
    
    ' Asignar variables con Trim para quitar espacios
    mes = Trim(partes(0))  ' "Dic"
    dia = Trim(partes(1))  ' "14"
    
    fecha = dia & "/" & mes
    
    limpiar_fecha = fecha
    
End Function
'======================================================================
'objectivo
'1 Recibe como datos clasificacion
'2 Retorna el id de la clasificacion
'Author: Elvis Junior Mamani Mendez
'Email: Elvisjuniorm21@gmail.com
'======================================================================
Function id_clasificacion(ByRef nombre) As String
Dim Libro As Workbook
Dim HojaAux As Worksheet
Dim valor As Range
Dim ruta, rango, codigo As String
Dim ultimoValorAux, Nro As Long

Set Libro = ThisWorkbook
Set HojaAux = Libro.Worksheets("Tablas")

ultimoValorAux = HojaAux.Cells(Rows.Count, 1).End(xlUp).Row
rango = "B3:B" & ultimoValorAux
    With HojaAux.Range(rango)
        Set valor = .Find(nombre, LookIn:=xlValues, LookAt:=xlWhole)
        If Not valor Is Nothing Then
            ruta = valor.Address
            Nro = Right(ruta, Len(ruta) - 3)
            codigo = HojaAux.Cells(Nro, 1).Value
            id_clasificacion = codigo
        End If
    End With
End Function
'======================================================================
'objectivo
'1 Recibe como datos tipo
'2 Retorna el id de la tipo
'Author: Elvis Junior Mamani Mendez
'Email: Elvisjuniorm21@gmail.com
'======================================================================
Function id_tipo(ByRef nombre) As String
Dim Libro As Workbook
Dim HojaAux As Worksheet
Dim valor As Range
Dim ruta, rango, codigo As String
Dim ultimoValorAux, Nro As Long

Set Libro = ThisWorkbook
Set HojaAux = Libro.Worksheets("Tablas")

ultimoValorAux = HojaAux.Cells(Rows.Count, 4).End(xlUp).Row
rango = "E3:E" & ultimoValorAux
    With HojaAux.Range(rango)
        Set valor = .Find(nombre, LookIn:=xlValues, LookAt:=xlWhole)
        If Not valor Is Nothing Then
            ruta = valor.Address
            Nro = Right(ruta, Len(ruta) - 3)
            codigo = HojaAux.Cells(Nro, 4).Value
            id_tipo = codigo
        End If
    End With
End Function
'======================================================================
'objectivo
'1 Recibe como datos activo
'2 Retorna el id del activo
'Author: Elvis Junior Mamani Mendez
'Email: Elvisjuniorm21@gmail.com
'======================================================================
Function id_detalle(ByRef nombre) As String
Dim Libro As Workbook
Dim HojaAux As Worksheet
Dim valor As Range
Dim ruta, rango, codigo As String
Dim ultimoValorAux, Nro As Long

Set Libro = ThisWorkbook
Set HojaAux = Libro.Worksheets("Tablas")

ultimoValorAux = HojaAux.Cells(Rows.Count, 7).End(xlUp).Row
rango = "H3:H" & ultimoValorAux
    With HojaAux.Range(rango)
        Set valor = .Find(nombre, LookIn:=xlValues, LookAt:=xlWhole)
        If Not valor Is Nothing Then
            ruta = valor.Address
            Nro = Right(ruta, Len(ruta) - 3)
            codigo = HojaAux.Cells(Nro, 7).Value
            id_detalle = codigo
        End If
    End With
End Function
'======================================================================
'objectivo
'1 Recibe como datos pasivo
'2 Retorna el id del pasivo
'Author: Elvis Junior Mamani Mendez
'Email: Elvisjuniorm21@gmail.com
'======================================================================
Function id_pasivo(ByRef nombre) As String
Dim Libro As Workbook
Dim HojaAux As Worksheet
Dim valor As Range
Dim ruta, rango, codigo As String
Dim ultimoValorAux, Nro As Long

Set Libro = ThisWorkbook
Set HojaAux = Libro.Worksheets("Tablas")

ultimoValorAux = HojaAux.Cells(Rows.Count, 10).End(xlUp).Row
rango = "K3:K" & ultimoValorAux
    With HojaAux.Range(rango)
        Set valor = .Find(nombre, LookIn:=xlValues, LookAt:=xlWhole)
        If Not valor Is Nothing Then
            ruta = valor.Address
            Nro = Right(ruta, Len(ruta) - 3)
            codigo = HojaAux.Cells(Nro, 10).Value
            id_pasivo = codigo
        End If
    End With
End Function
'======================================================================
'objectivo
'1 Recibe como datos patrimonio
'2 Retorna el id del patrimonio
'Author: Elvis Junior Mamani Mendez
'Email: Elvisjuniorm21@gmail.com
'======================================================================
Function id_patrimonio(ByRef nombre) As String
Dim Libro As Workbook
Dim HojaAux As Worksheet
Dim valor As Range
Dim ruta, rango, codigo As String
Dim ultimoValorAux, Nro As Long

Set Libro = ThisWorkbook
Set HojaAux = Libro.Worksheets("Tablas")

ultimoValorAux = HojaAux.Cells(Rows.Count, 13).End(xlUp).Row
rango = "N3:N" & ultimoValorAux
    With HojaAux.Range(rango)
        Set valor = .Find(nombre, LookIn:=xlValues, LookAt:=xlWhole)
        If Not valor Is Nothing Then
            ruta = valor.Address
            Nro = Right(ruta, Len(ruta) - 3)
            codigo = HojaAux.Cells(Nro, 13).Value
            id_patrimonio = codigo
        End If
    End With
End Function
'======================================================================
'objectivo
'1 Recibe como datos cuenta corriente
'2 Retorna el id del cuenta corriente
'Author: Elvis Junior Mamani Mendez
'Email: Elvisjuniorm21@gmail.com
'======================================================================
Function id_cuenta_corr(ByRef nombre) As String
Dim Libro As Workbook
Dim HojaAux As Worksheet
Dim valor As Range
Dim ruta, rango, codigo As String
Dim ultimoValorAux, Nro As Long

Set Libro = ThisWorkbook
Set HojaAux = Libro.Worksheets("Tablas")

ultimoValorAux = HojaAux.Cells(Rows.Count, 16).End(xlUp).Row
rango = "Q3:Q" & ultimoValorAux
    With HojaAux.Range(rango)
        Set valor = .Find(nombre, LookIn:=xlValues, LookAt:=xlWhole)
        If Not valor Is Nothing Then
            ruta = valor.Address
            Nro = Right(ruta, Len(ruta) - 3)
            codigo = HojaAux.Cells(Nro, 16).Value
            id_cuenta_corr = codigo
        End If
    End With
End Function
'======================================================================
'objectivo
'1 Recibe como datos cuenta orden
'2 Retorna el id del cuenta orden
'Author: Elvis Junior Mamani Mendez
'Email: Elvisjuniorm21@gmail.com
'======================================================================
Function id_cuenta_orden(ByRef nombre) As String
Dim Libro As Workbook
Dim HojaAux As Worksheet
Dim valor As Range
Dim ruta, rango, codigo As String
Dim ultimoValorAux, Nro As Long

Set Libro = ThisWorkbook
Set HojaAux = Libro.Worksheets("Tablas")

ultimoValorAux = HojaAux.Cells(Rows.Count, 19).End(xlUp).Row
rango = "T3:T" & ultimoValorAux
    With HojaAux.Range(rango)
        Set valor = .Find(nombre, LookIn:=xlValues, LookAt:=xlWhole)
        If Not valor Is Nothing Then
            ruta = valor.Address
            Nro = Right(ruta, Len(ruta) - 3)
            codigo = HojaAux.Cells(Nro, 19).Value
            id_cuenta_orden = codigo
        End If
    End With
End Function
'======================================================================
'objectivo
'1 Recibe como datos estado de resultados
'2 Retorna el id del estado de resultados
'Author: Elvis Junior Mamani Mendez
'Email: Elvisjuniorm21@gmail.com
'======================================================================
Function id_estado_resu(ByRef nombre) As String
Dim Libro As Workbook
Dim HojaAux As Worksheet
Dim valor As Range
Dim ruta, rango, codigo As String
Dim ultimoValorAux, Nro As Long

Set Libro = ThisWorkbook
Set HojaAux = Libro.Worksheets("Tablas")

ultimoValorAux = HojaAux.Cells(Rows.Count, 22).End(xlUp).Row
rango = "W3:W" & ultimoValorAux
    With HojaAux.Range(rango)
        Set valor = .Find(nombre, LookIn:=xlValues, LookAt:=xlWhole)
        If Not valor Is Nothing Then
            ruta = valor.Address
            Nro = Right(ruta, Len(ruta) - 3)
            codigo = HojaAux.Cells(Nro, 22).Value
            id_estado_resu = codigo
        End If
    End With
End Function
