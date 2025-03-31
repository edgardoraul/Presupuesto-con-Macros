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
    Dim pass As String
    pass = "Rerda2024" ' Contraseña para proteger y desproteger
    ' Desactivo el cuadro de diálogo.
    Cancel = False
    
    ActiveWorkbook.Protect password:=pass
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
    Dim hojita As Worksheet
    Dim pass As String
    pass = "Rerda2024"
      
    ActiveWorkbook.Unprotect password:=pass
    
    ' Ocultando algunas
    Sheets(1).Visible = True
    Sheets(2).Visible = False
    Sheets(3).Visible = False

   
    ' Actualizando la ruta
    'Call creandoRuta
    Call EstaEnGrupoDeTrabajo
    
    ' Dar formato de impresión
    Call darFormato
    
    ' Protegiendo las hojas
    For Each hojita In Application.Worksheets
        hojita.Protect password:=pass, DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
    Next hojita
    
    ' Posicionando
    Sheets(1).Activate
    Sheets(1).Range("A9").Activate
    
    ActiveWorkbook.Protect password:=pass
End Sub






