VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
    'Dim ultimaConDatos As Integer
    Dim ws As Worksheet
    Dim pass As String

    pass = "Rerda2024" ' Contraseña para proteger y desproteger
    Set ws = Me ' Refierete a la hoja actual
    
    Call creandoRuta
    
    originalRowHeight = 18
    
    ' Control para evitar errores de desbordamiento.
    Call ultima
    
    ' Define el rango de celdas que activará el evento (por ejemplo, columna A)
    Set rng = Intersect(Target, Range(Cells(9, 1), Cells(ultimaConDatos, 1)))

    If Not rng Is Nothing Then
        Application.EnableEvents = False ' Desactivar eventos para evitar bucles infinitos

        ' Desproteger la hoja antes de hacer cambios
        ws.Unprotect password:=pass

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
                            
                            Set picToDelete = ws.Pictures.Insert(imgPath & imagePath)

                            With picToDelete
                                .Top = cel.Offset(0, 7).Top + 4
                                .Left = cel.Offset(0, 7).Left + 4
                                .ShapeRange.LockAspectRatio = msoTrue
                                .ShapeRange.Height = 100
                            End With
                        End With

                        ' Agregar borde superior a las celdas de la columna 1 a la 8
                        If cel.Column <= 7 And cel.Row >= 8 Then
                            ws.Cells(cel.Row, 1).Resize(1, 8).Borders(xlEdgeTop).LineStyle = xlContinuous
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
        Next cel
        
        ' Proteger la hoja después de hacer cambios
        ws.Protect password:=pass, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
        Application.EnableEvents = True ' Aactivar eventos
    End If

    Application.EnableEvents = True ' Re-activar eventos
End Sub

