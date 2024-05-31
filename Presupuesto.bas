Attribute VB_Name = "Presupuesto"
Option Explicit
Public Ruta As String
Public ultimaConDatos As Integer
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
    ultimaConDatos = Cells(Rows.Count, 1).End(xlUp).Row
    Debug.Print "La última fila es " & ultimaConDatos
    
    ' Actualiza la autosuma
    Call actualizarSuma(ultimaConDatos)
End Sub

Sub borrarFila()
    Dim i As Byte
    
    
    Debug.Print "La última fila es " & ultimaConDatos
    If Cells(9, 1).Value = "" And Cells(9, 3).Value = "" And Cells(9, 4).Value = "" And Cells(9, 5).Value = "" And Cells(9, 6).Value = "" Then
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
    ultimaConDatos = Cells(8, 1).End(xlDown).Row
    ultimaConDatos = Cells(Rows.Count, 1).End(xlUp).Row
    Debug.Print "La última fila es " & ultimaConDatos
    actualizarSuma (ultimaConDatos)
End Sub

Sub darFormato()
' Formato de impresión
    With ActiveSheet.PageSetup
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
        '.PrintArea = ActiveSheet.Range("A1:H21")
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With
    
End Sub
Sub creandoRuta()
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
End Sub

Function mostrarErrorRed()
    MsgBox ("Hay que prender la computadora EDGAR")
End Function
