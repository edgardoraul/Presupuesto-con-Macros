Option Explicit
Dim ruta As String
Dim nombre As String
Dim archivo As String

Sub Presupuesto()
	'Se asignan valores a las variables globales
	ruta = ActiveWorkbook.Path
	nombre = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
	archivo = ActiveWorkbook.FullName
	Debug.Print archivo

	'## Crea un presupuesto. ¿Está cargado este cliente de antes?
	'if then
		'Sí. Abre un cuadro de diálogo para buscar el presupuesto donde tomar los datos
		'¿Mismas condiciones que antes? (iva, plazo, etc...)
		'Cuadro de diálogo que pregunta
		'if then NO
			' Nuevas condiciones
		'else SI
			' Deja las que estaban
		'end if

		'Renombra el nuevo presupuesto
		'Call saveFile(nombre)
	'else
		'Genera un archivo nuevo
		'Nuevas condiciones
		'Call saveFile(ruta, nombre)
	'end if

	'## Actualiza en segundo plano los precios

	'## ¿Se usará SKU?
	'Cuadro de diálogo que pregunta
	' if then SI
		' Se coloca fórmula a medida que va escribiendo
		' para mostrar leyenda junto con el respectivo precio
	' else NO
		'Se carga todo manual
	' end if

	'## El usuario va cargando talles, colores y cantidades

	'## ¿Tiene foto? Preguntar con cuadro de diálogo
	'If SI Then
		' Buscar archivo automático de foto y elige la primera
		' Call cargaFoto(sku)
	'else NO
		'Seguir sin foto, nomás
	'end if

	'## Guarda cambios, publica en pdf y cierra el archivo
	'Call guardaPdf(ruta)
End Sub
Function cargaFoto(sku)
' ¿Hay foto?
'If YES Then
	' carga foto
'Else
	' buscar manual
'End If
End Function
Function saveFile(ruta, nombre)
	' GUARDA EL ARCHIVO CON UN CRITERIO
	' Número incremental
	' Palabra Presupuesto
	' Nombre o razón social del cliente
End Function

Function guardaPdf(ruta, nombre)
	' Guarda una copia en pdf y abre el archivo

	ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
		ruta & "\" & nombre & ".pdf", _
		Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
		:=False, OpenAfterPublish:=True
	ActiveWorkbook.Save
	ActiveWorkbook.Close
End Function
Function exceljson()
	'Obtiene la información vía rest api
	Dim http As Object, JSON As Object, i As Integer, item As Variant
	Set http = CreateObject("MSXML2.XMLHTTP")
	http.Open "GET", "http://jsonplaceholder.typicode.com/users", False
	http.send
	Set JSON = ParseJson(http.responseText)
	i = 2
	For Each item In JSON
		With Worksheets("Resultados")
			.Cells(i, 1).Value = item("id")
			.Cells(i, 2).Value = item("name")
			.Cells(i, 3).Value = item("username")
			.Cells(i, 4).Value = item("email")
			.Cells(i, 5).Value = item("address")("city")
			.Cells(i, 6).Value = item("phone")
			.Cells(i, 7).Value = item("website")
			.Cells(i, 8).Value = item("company")("name")
		End With
		i = i + 1
	Next
	MsgBox ("complete")
End Function
Private Sub Workbook_BeforeClose(Cancel As Boolean)
	ThisWorkbook.Save
End Sub


