' Sample 20: Excel Drill Report.BAS
' 
' This sample demonstrates how to generate PADS Layout reports in Excel.
'
' For more details, please refer to the PADS Layout Basic Editor Help File.
'

' Drill counter
Dim DrillCount As Long

Sub Main
	' Open temporarly text file
	Randomize
	filename = DefaultFilePath & "\tmp"  & CInt(Rnd()*10000) & ".txt"
	Open filename For Output As #1

	' Output Headers
	Print #1, "Type";		Space(9); 
	Print #1, "Name";		Space(26); 
	Print #1, "Net";		Space(28); 
	Print #1, "Drill";		Space(25); 
	Print #1, "Plated";		Space(2); 
	Print #1, "X";			Space(29); 
	Print #1, "Y";			Space(0)

	DrillCount = 0
	
	' Lock server to speed up process
	LockServer

	' Go through each pin in the design and output values
	For Each nextPin In ActiveDocument.Pins
		If nextPin.IsSMD = False Then 
			' Output Pin Type Property
			a = "Pin"
			Print #1, a; Space$(13-Len(a)); 

			' Output Pin Name Property
			a = nextPin.Name
			Print #1, a; Space$(30-Len(a)); 

			' Output Pin Attached Net Property
			Set nnet = nextPin.Net
			If (nnet Is Nothing) Then a = "N/A" Else a = nnet.Name 
			Print #1, a; Space$(30-Len(a)); 
		
			' Output Pin Drill Size Property
			a = nextPin.DrillSize
			Print #1, a; Space$(30-Len(a)); 
		
			' Output Pin Plated Property
			a = nextPin.Plated
			Print #1, a; Space$(8-Len(a)); 

			' Output Pin PositionX Property
			a = Format$(nextPin.PositionX, "#.00")
			Print #1, a; Space$(30-Len(a)); 

			' Output Pin PositionY Property
			a = Format$(nextPin.PositionY, "#.00")
			Print #1, a
			DrillCount = DrillCount + 1
		End If	
	Next nextPin

	' Go through each via in the design and output values
	For Each nextVia In ActiveDocument.Vias

		' Output Via Type Property
		a = "Via"
		Print #1, a; Space$(13-Len(a)); 

		' Output Via Name Property
		a = nextVia.Type
		Print #1, a; Space$(30-Len(a)); 

		' Output Via Attached Net Property
		Set nnet = nextVia.Net
		If (nnet Is Nothing) Then a = "N/A" Else a = nnet.Name 
		Print #1, a; Space$(30-Len(a)); 
		
		' Output Via Drill Size Property
		a = nextVia.DrillSize
		Print #1, a; Space$(30-Len(a)); 
		
		' Output Via Plated Property
		a = nextVia.Plated
		Print #1, a; Space$(8-Len(a)); 

		' Output Via PositionX Property
		a = Format$(nextVia.PositionX, "#.00")
		Print #1, a; Space$(30-Len(a)); 

		' Output Via PositionY Property
		a = Format$(nextVia.PositionY, "#.00")
		Print #1, a
		
		DrillCount = DrillCount + 1
	Next nextVia

	' Unlock the server
	UnlockServer

	' Close the text file
	Close #1

	If DrillCount = 0 Then
		MsgBox "The design has no drills"
		Exit Sub
	End If

	' Start Excel and loads the text file
	On Error GoTo noExcel
	Dim excelApp As Object
	Set excelApp = CreateObject("Excel.Application")
	On Error GoTo 0
	excelApp.Visible = True
	excelApp.Workbooks.OpenText 	FileName:= filename
	excelApp.Rows("1:1").Select
	With excelApp.Selection
		.Font.Bold = True
		.Font.Italic = True
	End With
	excelApp.Cells.Select
	excelApp.Selection.Sort Key1:=excelApp.ActiveSheet.Columns("D"), Order1:=1, Header:=0
	excelApp.Selection.Subtotal GroupBy:=4, Function:=-4112, TotalList:=Array(5), Replace:=True
	excelApp.ActiveSheet.Outline.ShowLevels RowLevels:=2
	excelApp.Range("A1").Select
	Set excelApp = Nothing
	End
	
noExcel:
	' Display the text file
	Shell "Notepad " & filename, 3

End Sub
