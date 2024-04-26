' ====================================
'		MODULO PRESUPUESTO
' ====================================

Option Explicit
Public Ruta As String
' Guarda una copia en pdf y abre el archivo
Sub guardaPdf()
    Dim Rutita As String
    Dim nombre As String
    Rutita = ActiveWorkbook.Path
    nombre = Year(Date) & "-" & Month(Date) & "-" & Day(Date) & ". PRESUPUESTO - " & Range("B4").Value

    Dim fecha As String
    fecha = Year(Date) & "-" & Month(Date) & "-" & Day(Date)

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

Sub insertarFila()
    Dim ultimaConDatos As Integer
    
    ultimaConDatos = Cells(8, 1).End(xlDown).Row
    
    ' Inserta una fila arriba, copiando el formato de la de abajo.
    Rows(9).Insert Shift:=xlShiftUp, CopyOrigin:=xlFormatFromRightOrBelow
    Cells(9, 1).Activate
End Sub

Sub borrarFila()
    Dim i As Byte
    
    If Cells(9, 1).Value = "" And Cells(9, 3).Value = "" And Cells(9, 4).Value = "" And Cells(9, 5).Value = "" And Cells(9, 6).Value = "" Then
        Rows(9).Delete
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
End Sub


' ====================================
'		THISWORKBOOK (Eventos)
' ====================================
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Desactivo el cuadro de diálogo.
    Cancel = False
    ActiveWorkbook.Save
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    SaveAsUI = False
    Cancel = False
    Dim nombre As String
    Dim fecha As String
    fecha = Year(Date) & "-" & Month(Date) & "-" & Day(Date)
    nombre = ThisWorkbook.Path & "\" & fecha & ". PRESUPUESTO - " & Worksheets(1).Range("B4").Value & ".xlsm"
    ActiveWorkbook.SaveCopyAs nombre
End Sub

Private Sub Workbook_Open()
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



' ====================================
'		HOJA1: PRESUPUESTO
' ====================================
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rng As Range
    Dim cel As Range
    Dim codigo As String
    Dim imagePath As String
    Dim imgPath As String
    Dim existingPic As Picture
    Dim picToDelete As Picture
    Dim originalRowHeight As Double
    Dim rowHasImage As Boolean
    Dim ultimaConDatos As Integer
    
    originalRowHeight = 18
    ' Define el rango de celdas que activará el evento (por ejemplo, columna A)
    ultimaConDatos = Cells(8, 1).End(xlDown).Row
    Set rng = Intersect(Target, Range(Cells(9, 1), Cells(ultimaConDatos, 1)))

    If Not rng Is Nothing Then
        Application.EnableEvents = False ' Desactivar eventos para evitar bucles infinitos

        For Each cel In rng
            codigo = cel.Value
            If codigo <> "" Then
                               
                ' Comprobar si la carpeta correspondiente al código del producto existe
                imgPath = Ruta & codigo & "\"
                
                Debug.Print imgPath
                
                If Dir(imgPath, vbDirectory) <> "" Then
                    ' Obtener la primera imagen en la carpeta
                    imagePath = Dir(imgPath & "*.*")

                    ' Comprobar si se encontró alguna imagen en la carpeta
                    If imagePath <> "" Then
                    
                        ' Guardar la altura original de la fila
                        ' originalRowHeight = cel.EntireRow.RowHeight
                        
                        rowHasImage = True

                        ' Eliminar imagen existente en la misma fila
                        For Each existingPic In cel.Offset(0, 7).Parent.Pictures
                            If Not Intersect(existingPic.TopLeftCell.EntireRow, cel.EntireRow) Is Nothing Then
                                existingPic.Delete
                                rowHasImage = False
                            End If
                        Next existingPic

                        ' Insertar la nueva imagen en la celda adyacente
                        With cel.Offset(0, 7)
                            .ColumnWidth = 20 ' Ajustar el ancho de la columna para la imagen
                            .RowHeight = 108 ' Ajustar la altura de la fila para la imagen
                            .Activate
                            Set picToDelete = ActiveSheet.Pictures.Insert(imgPath & imagePath)


                            With picToDelete
                                .Top = .TopLeftCell.Top + 4
                                .Left = .TopLeftCell.Left + 4
                                .ShapeRange.LockAspectRatio = msoTrue
                                .ShapeRange.Height = 100
                            End With

                        End With
                        
                        
                        ' Agregar borde superior a las celdas de la columna 1 a la 8
                        If cel.Column <= 7 And cel.Row >= 8 Then
                            Me.Cells(cel.Row, 1).Resize(1, 8).Borders(xlEdgeTop).LineStyle = xlContinuous
                        End If

                    Else
                    
                        ' Restaurar la altura original de la fila, sólo si no hay imagen, de lo contrario se autoajusta al contenido
                        If rowHasImage Then
                            cel.EntireRow.RowHeight = originalRowHeight
                        End If
                        rowHasImage = False
                        
                        
                    End If
                
                Else
                    ' Acomoda el alto de la fila de acuerdo al contenido
                    cel.EntireRow.AutoFit
                    
                End If
            Else
                ' Restaurar la altura original de la fila si se borra el contenido de la celda
                cel.EntireRow.RowHeight = originalRowHeight
                
            End If
            
            ' Introducir la fórmula en la celda de la columna 2 si se edita una celda en la columna 1
            If cel.Column = 1 Then
                Dim formulaCel As Range
                Set formulaCel = cel.Offset(0, 1)
                formulaCel.Formula = "=IFERROR(VLOOKUP(" & cel.Address & ",Resultados!A$1:E$10000,2,FALSE),"""")"
            End If
            
            ' Introducir la fórmula en la celda de la columna 7 si se edita una celda en la columna 1
            If cel.Column = 1 Then
                Dim formulaCel2 As Range
                Set formulaCel2 = cel.Offset(0, 6)
                formulaCel2.Formula = "=IF(E" & cel.Row & "*F" & cel.Row & "=0,"""",E" & cel.Row & "*F" & cel.Row & ")"
            End If
            ' Para parase en la fila siguiente
            cel.Offset(1, 0).Activate
        Next cel
        
    End If
    Application.EnableEvents = True ' Re-activar eventos
End Sub

