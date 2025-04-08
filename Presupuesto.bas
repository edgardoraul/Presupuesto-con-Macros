Attribute VB_Name = "Presupuesto"
Option Explicit
Public Ruta As String
Public ultimaConDatos As Integer
Public grupo As String
Public url As String
Public codigo As String
Public carpetaPrincipal As String
Public carpetaCodigo As String
Public imagenUrl As String
Public imagenDestino As String
Public imgPath As String

' Guarda una copia en pdf y abre el archivo
Sub guardaPdf()
    Dim Rutita As String
    Dim nombre As String
    Rutita = ActiveWorkbook.Path
    nombre = Year(Date) & "-" & Month(Date) & "-" & Day(Date) & ". PRESUPUESTO - " & Range("B4").Value

    Dim fecha As String
    fecha = Year(Date) & "-" & Month(Date) & "-" & Day(Date)
    
    ' Sistema de control
    If Worksheets(1).Range("B4").Value = "" Then
        MsgBox ("Te faltó el nombre o razón social.")
        Worksheets(1).Range("B4").Activate
        Exit Sub
    End If

    nombre = ThisWorkbook.Path & "\" & fecha & ". PRESUPUESTO - " & Worksheets(1).Range("B4").Value & ".xlsm"
    ThisWorkbook.SaveCopyAs nombre
    
    ActiveWorkbook.Save
    nombre = fecha & ". PRESUPUESTO - " & Worksheets(1).Range("B4").Value
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        Rutita & "\" & nombre & ".pdf", _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True
    
    ' Muestra el archivo en carpeta para enviar por mail o imprimir.
    Shell "explorer " & Rutita, vbNormalFocus
    
End Sub
Function actualizarSuma(ultimaConDatos)
    Cells(ultimaConDatos + 2, 7).Value = "=SUM(G9:G" & ultimaConDatos & ")"
End Function

Sub insertarFila()
    ' Inserta una fila arriba, copiando el formato de la de abajo.
    Rows(9).Insert Shift:=xlShiftUp, CopyOrigin:=xlFormatFromRightOrBelow
    Cells(9, 1).Activate
    'ultimaConDatos = Cells(Rows.Count, 1).End(xlUp).Row
    Call ultima
    Debug.Print "La última fila es " & ultimaConDatos
    
    ' Actualiza la autosuma
    Call actualizarSuma(ultimaConDatos)
End Sub

Sub borrarFila()
    Dim i As Byte
    ' Debe haber la cantidad
    Call ultima
    
    Debug.Print "La última fila es " & ultimaConDatos
    'If Cells(9, 1).Value = "" And Cells(9, 3).Value = "" And Cells(9, 4).Value = "" And Cells(9, 5).Value = "" And Cells(9, 6).Value = "" Then
    If ultimaConDatos > 9 Then
        Rows(9).Delete
        ultimaConDatos = Cells(8, 1).End(xlDown).Row
    ElseIf ultimaConDatos <= 9 Then
        MsgBox ("¡Ojo!" & vbNewLine & vbNewLine & "¡No se puede borrar esta última fila!")
        ultimaConDatos = 9
        Exit Sub
    Else
        ' Recorre la fila para que el usuario borre el contenido. Da tiempo de arrepentirse.
        For i = 1 To 8
            If Cells(9, i) <> "" Then
                MsgBox "Primero borrá el contenido."
                Cells(9, i).Activate
                Exit Sub
            End If
        Next i
    End If
    Cells(9, 1).Activate
    'ultimaConDatos = Cells(8, 1).End(xlDown).Row
    'ultimaConDatos = Cells(Rows.Count, 1).End(xlUp).Row
    Call ultima
    Debug.Print "La última fila es " & ultimaConDatos
    actualizarSuma (ultimaConDatos)
End Sub

Sub darFormato()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim pageBreakRow As Long
    Dim totalRows As Long
    Dim minRowsOnLastPage As Long
    
    Set ws = ActiveSheet
    minRowsOnLastPage = 9 ' Cantidad mínima de filas que deben estar completas en la última página
    
    ' Última fila con datos
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row + 9
    
    ' Configuración de impresión
    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .LeftMargin = Application.CentimetersToPoints(0.64)
        .RightMargin = Application.CentimetersToPoints(0.64)
        .TopMargin = Application.CentimetersToPoints(2.5)
        .BottomMargin = Application.CentimetersToPoints(1.91)
        .HeaderMargin = Application.CentimetersToPoints(0.76)
        .FooterMargin = Application.CentimetersToPoints(0.76)
        .CenterHorizontally = True
        .CenterVertically = False
        .Zoom = False
        .FitToPagesWide = 1
    End With
    
    ' Restablecer saltos de página y cambiar a vista previa de saltos
    ws.ResetAllPageBreaks
    ActiveWindow.View = xlPageBreakPreview
    
    ' Verificar si hay al menos un salto de página horizontal
    If ws.HPageBreaks.Count > 0 Then
        pageBreakRow = ws.HPageBreaks(1).Location.Row - 1 + 9
    Else
        pageBreakRow = lastRow ' Si no hay saltos de página, tomar la última fila
    End If
    
    ' Si el salto de página corta las últimas 9 filas, reajustar el área de impresión
    If (lastRow - pageBreakRow) < minRowsOnLastPage Then
        totalRows = lastRow - minRowsOnLastPage + 9
        ws.PageSetup.PrintArea = ws.Range("A1:H" & totalRows).Address
    Else
        ws.PageSetup.PrintArea = ws.Range("A1:H" & lastRow).Address
    End If
    
    ActiveWindow.View = xlNormalView ' Volver a la vista normal
End Sub


Function creandoRuta()
    Dim hojita As Worksheet
    Dim ws As Object
    Set ws = CreateObject("WScript.network")
    
    Application.EnableEvents = True ' Re-activar eventos
    
    ' Asignando algunos valores de acuerdo en qué equipo de la red esté
    If ws.ComputerName = "EDGAR" Then
        Ruta = "D:\Web\imagenes_rerda\"
    Else
        Ruta = "\\EDGAR\Web\imagenes_rerda\"
    End If
    imgPath = Ruta & codigo & "\"
End Function

Function mostrarErrorRed()
    MsgBox ("Hay que prender la computadora EDGAR")
End Function

Function ultima()
    ultimaConDatos = Cells(Rows.Count, 1).Offset(0, 6).End(xlUp).Row - 3
    If ultimaConDatos >= 9 Then
        ultimaConDatos = Cells(8, 1).Offset(0, 6).End(xlDown).Row - 3
    Else
        Exit Function
    End If
End Function

Function EstaEnGrupoDeTrabajo() As Boolean
    Dim objWMI As Object
    Dim objItem As Object
    Dim colItems As Object
    Dim strGrupoTrabajo As String
    
    grupo = "RERDA"
    
    ' Obtener información del sistema
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery("Select * from Win32_ComputerSystem")
    
    ' Extraer el grupo de trabajo
    For Each objItem In colItems
        strGrupoTrabajo = objItem.Workgroup
        Exit For
    Next
    
    Call VerificarRed(strGrupoTrabajo)
End Function

Sub VerificarRed(red As String)
    url = "https://raw.githubusercontent.com/edgardoraul/imagenes_rerda/main/"
    If red = grupo Then
        Call creandoRuta
    Else
        Call DescargarImagen(url, codigo)
    End If
End Sub



Function DescargarImagen(url As String, codigo As String)
    Dim fso As Object
    
    Dim http As Object
    Dim stream As Object
    
    ' Obtener la ruta del workbook actual
    carpetaPrincipal = ThisWorkbook.Path & "\imagenes_rerda"
    carpetaCodigo = carpetaPrincipal & "\" & codigo
    imagenUrl = url & codigo & "/1.jpg"
    imagenDestino = carpetaCodigo & "\1.jpg"

    
    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Crear carpeta principal si no existe
    If Not fso.FolderExists(carpetaPrincipal) Then
        fso.CreateFolder carpetaPrincipal
    End If
    
    ' Crear carpeta de código si no existe
    If Not fso.FolderExists(carpetaCodigo) Then
        fso.CreateFolder carpetaCodigo
    End If
    
    ' Actualizar la dirección de las carpetas
    Ruta = carpetaCodigo
    
    ' Descargar la imagen
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", imagenUrl, False
    http.send
    
    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1 ' Binario
        stream.Open
        stream.Write http.responseBody
        stream.SaveToFile imagenDestino, 2 ' Guardar archivo
        stream.Close
        Set stream = Nothing
    End If
    
    ' Limpiar objetos
    Set http = Nothing
    Set fso = Nothing
    
    imgPath = carpetaCodigo & "\"
End Function

