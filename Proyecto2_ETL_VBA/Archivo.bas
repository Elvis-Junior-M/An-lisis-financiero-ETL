Attribute VB_Name = "Archivo"
Option Explicit
 ' ------------------------------------------------------
 ' Nombre: archivo()
 ' Descripcion: Esta Macro abre un cuadro de dialogo donde se selecciona el archivo *.xlsx"? ?
 ' Autor: Elvis Junior Mamani Mendez
 ' Fecha: 27/09/2025
 ' ------------------------------------------------------
Function ruta_archivo() As String
Dim ruta As String
    With Application.FileDialog(msoFileDialogFilePicker)
            .InitialFileName = ThisWorkbook.Path & "\"
            .Title = "Seleccionar archivo"
            .Filters.Clear
            .Filters.Add "Excel Files", "*.xls?", 1
            .AllowMultiSelect = False
            .Show
            If .SelectedItems.Count = 0 Then
                MsgBox "No se selecciono un archivo"
                ruta_archivo = "0"
                ruta_archivo = ruta
                GoTo SaltoFin
            Else
                ruta = .SelectedItems(1)
                ruta_archivo = ruta
            End If
    End With
SaltoFin:
End Function
