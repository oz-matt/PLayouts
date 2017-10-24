' Sample 22: Excel Drill Drawing with Testpoints.BAS
' 
' This sample demonstrates how to generate PADS Layout reports in Excel.
'
' For more details, please refer to the PADS Layout Basic Editor Help File.
'

' Main reference to Excel
Dim excelApp As Object
' Drill counter
Dim DrillCount As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Switch to a given Excel Sheet. Create the sheet if it doesn't exist.
Sub switchToSheet(sheetName$)
	On Error GoTo noSheet
	excelApp.Sheets(sheetName).Select
	Exit Sub
noSheet:
	excelApp.Sheets.Add
	excelApp.ActiveSheet.Name = sheetName
	excelApp.Range("A1").Select
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Output drill coordinate information in Excel. One Excel sheet per drill size
Sub outputDrillInformation()
	DrillCount = 0
	' Output all PIN informations (one sheet per drill size).
	For Each nextPin In ActiveDocument.Pins
		If nextPin.IsSMD = False Then 
			If nextPin.Plated = False Then 
				If (nextPin.TestPoint <> 0) Then
					sheetName = "Drill " & nextPin.DrillSize & "TP"
				Else
					sheetName = "Drill " & nextPin.DrillSize
				End If
			Else
				If (nextPin.TestPoint <> 0) Then
					sheetName = "Drill " & nextPin.DrillSize & "+TP"
				Else
					sheetName = "Drill " & nextPin.DrillSize & "+"
				End If
			End If
			switchToSheet(sheetName)
			excelApp.Cells(excelApp.ActiveCell.Row, 1).FormulaR1C1 = nextPin.PositionX
			excelApp.Cells(excelApp.ActiveCell.Row, 2).FormulaR1C1 = nextPin.PositionY
			If (nextPin.TestPoint = 1) Then
				excelApp.Cells(excelApp.ActiveCell.Row, 3).FormulaR1C1 = "Testpoint Top Side"
			ElseIf  (nextPin.TestPoint = 2) Then
				excelApp.Cells(excelApp.ActiveCell.Row, 3).FormulaR1C1 = "Testpoint Bottom Side"
			End If
			excelApp.Cells(excelApp.ActiveCell.Row+1, 1).Select
			DrillCount = DrillCount + 1
		End If	
	Next nextPin

	' Output all VIA informations (one sheet per drill size).
	For Each nextVia In ActiveDocument.Vias
		If nextVia.Plated = False Then 
			If (nextVia.TestPoint <> 0) Then
				sheetName = "Drill " & nextVia.DrillSize & "TP"
			Else
				sheetName = "Drill " & nextVia.DrillSize
			End If
		Else
			If (nextVia.TestPoint <> 0) Then
				sheetName = "Drill " & nextVia.DrillSize & "+TP"
			Else
				sheetName = "Drill " & nextVia.DrillSize & "+"
			End If
		End If
		switchToSheet(sheetName)
		excelApp.Cells(excelApp.ActiveCell.Row, 1).FormulaR1C1 = nextVia.PositionX
		excelApp.Cells(excelApp.ActiveCell.Row, 2).FormulaR1C1 = nextVia.PositionY
		If (nextVia.TestPoint = 1) Then
			excelApp.Cells(excelApp.ActiveCell.Row, 3).FormulaR1C1 = "Testpoint Top Side"
		ElseIf  (nextVia.TestPoint = 2) Then
			excelApp.Cells(excelApp.ActiveCell.Row, 3).FormulaR1C1 = "Testpoint Bottom Side"
		End If
		excelApp.Cells(excelApp.ActiveCell.Row+1, 1).Select
		DrillCount = DrillCount + 1
	Next nextVia
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Main Routine
Sub Main
	' Start Excel (exit if Excel is not installed)
	On Error GoTo noExcel
	Set excelApp = CreateObject("Excel.Application")
	On Error GoTo 0
	excelApp.Workbooks.Add
	
	MsgBox "This script will generate an Excel Drill Drawing, which may take a minute or so. Please be patient..."

	' Output infos in Excel
	LockServer
	outputDrillInformation
	UnlockServer

	If DrillCount = 0 Then
		MsgBox "The design has no drills"
		Set excelApp = Nothing
		Exit Sub
	End If
	
	' Create Excel chart (XY Scatter type) from data
	excelApp.Charts.Add
	excelApp.ActiveChart.ChartType = -4169 ' xlXYScatter
	' For Each drill size (i.e. Each Excel sheet)...
	For Each nextSheet In excelApp.Sheets
		If Left$(nextSheet.Name, 5) = "Drill" Then
			' ... add a new serie for that drill size
			Set newSerie = excelApp.ActiveChart.SeriesCollection.NewSeries
			newSerie.XValues = nextSheet.Range("A1:A" & nextSheet.Range("A1").End(-4121).Row) ' xlDown
			newSerie.Values = nextSheet.Range("B1:B" & nextSheet.Range("B1").End(-4121).Row) ' xlDown
			' Serie name is the drill size
			newSerie.Name = nextSheet.Name & " (x" & nextSheet.Range("A1").End(-4121).Row & ")"
			' Use filled round markers (but we can change that)
			newSerie.MarkerBackgroundColorIndex = -4105' xlAutomatic
			newSerie.MarkerForegroundColorIndex = -4105' xlAutomatic
			newSerie.MarkerStyle = 8 ' xlCircle
			' The marker size is proportional to the drill size (how cool...!)
			drillsizename = Mid$(nextSheet.Name, 7)
			whereIsThePlus = InStr(drillsizename, "+")
			If  whereIsThePlus > 0 Then 
				drillsizename = Left$(drillsizename, whereIsThePlus-1)
			End If
			drillsizesize = CInt(drillsizename)
			If drillsizesize/10 < 2 Then drillsizesize = 20
			newSerie.MarkerSize = drillsizesize/10
		End If
	Next nextSheet
	' Set various chert prefs (no axes, no scales, black background...)
	excelApp.ActiveChart.Location Where:=1 ' xlLocationAsNewSheet
	excelApp.ActiveChart.HasTitle = True
	excelApp.ActiveChart.ChartTitle.Characters.Text = ActiveDocument.FullName & " Drill Drawing on " & Now
	excelApp.ActiveChart.PlotArea.Interior.ColorIndex = 1
	excelApp.ActiveChart.HasAxis(1, 1) = False ' xlCategory, xlPrimary
	excelApp.ActiveChart.HasAxis(2, 1) = False ' xlValue, xlPrimary
	excelApp.ActiveChart.Axes(1, 1).CategoryType = -4105 ' xlCategory, xlPrimary, xlAutomatic
	With excelApp.ActiveChart.Axes(1) ' xlCategory
		.HasMajorGridlines = False
		.HasMinorGridlines = False
	End With
	With excelApp.ActiveChart.Axes(2) ' xlValue
		.HasMajorGridlines = False
		.HasMinorGridlines = False
	End With
	excelApp.ActiveChart.ApplyDataLabels Type:=-4142, LegendKey:=False ' xlDataLabelsShowNone
	excelApp.Visible = True
	Set excelApp = Nothing
	Exit Sub
	
noExcel:
	MsgBox "You need Microsoft Excel 97 to run this Basic script. Sorry..."

End Sub
