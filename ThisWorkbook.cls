VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
    
    ' Sistema de control
    If Worksheets(1).Range("B4").Value = "" Then
        MsgBox ("Te faltó el nombre o razón social.")
        Worksheets(1).Range("B4").Activate
        Exit Sub
    End If
    
    nombre = ThisWorkbook.Path & "\" & fecha & ". PRESUPUESTO - " & Worksheets(1).Range("B4").Value & ".xlsm"
    ActiveWorkbook.Save
End Sub

Private Sub Workbook_Open()

    ' Actualizando la ruta
    Call creandoRuta
    
    ' Dar formato de impresión
    Call darFormato
End Sub






