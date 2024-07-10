Attribute VB_Name = "ValidacionRangos"
Sub AsignarNombresDeRango()
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim lastRow As Long
    Dim col As Long
    Dim rng As Range
    Dim nombreRango As String

    Set ws = ThisWorkbook.Sheets("Variantes") ' Cambia "Sheet1" por el nombre de tu hoja

    ' Obtener la última columna con datos en la fila 1
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Recorrer cada columna desde la 1 hasta la última con datos en la fila 1
    For col = 1 To lastCol
        nombreRango = ws.Cells(1, col).Value ' Obtener el contenido de la primera celda de la columna como nombre de rango

        ' Saltar si el nombre del rango está vacío o no es un nombre válido
        If nombreRango <> "" And Not IsNumeric(nombreRango) Then
            ' Reemplazar espacios y caracteres no permitidos en los nombres de rango
            nombreRango = Application.Substitute(nombreRango, " ", "_")
            nombreRango = Application.Substitute(nombreRango, "-", "_")
            nombreRango = Application.Substitute(nombreRango, ".", "_")
            nombreRango = Application.Clean(nombreRango)

            ' Obtener la última fila con datos en la columna actual
            lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

            ' Si la última fila es mayor a 1 (hay datos debajo del encabezado)
            If lastRow > 1 Then
                ' Seleccionar el rango desde la fila 2 hasta la última fila con datos en la columna actual
                Set rng = ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col))

                ' Asignar el nombre de rango a nivel de libro
                ThisWorkbook.Names.Add Name:=nombreRango, RefersTo:=rng
            End If
        End If
    Next col

    Debug.Print "Nombres de rango asignados exitosamente."

End Sub


