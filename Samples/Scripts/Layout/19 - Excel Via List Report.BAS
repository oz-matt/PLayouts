' Sample 19: Excel Via List Report.BAS
' 
' This sample demonstrates how to generate PADS Layout reports in Excel.
'
' For more details, please refer to the PADS Layout Basic Editor Help File.
'

Sub Main
	' Open temporarly text file
	Randomize
	filename = DefaultFilePath & "\tmp"  & CInt(Rnd()*10000) & ".txt"
	Open filename For Output As #1

	' Output Headers
	Print #1, "Name";		Space(10); 
	Print #1, "Type";		Space(27); 
	Print #1, "Start";		Space(2); 
	Print #1, "End";		Space(3); 
	Print #1, "Net";		Space(28); 
	Print #1, "Drill";		Space(25); 
	Print #1, "X";			Space(29); 
	Print #1, "Y";			Space(29); 
	Print #1, "Plated";		Space(2); 
	Print #1, "TestPoint";	Space(1);
	Print #1, "Thermal";	Space(1);
	Print #1, "Glued";		Space(2) 
	
	' Lock server to speed up process
	LockServer

	' Go through each pin in the design and output values
	For Each nextVia In ActiveDocument.Vias
		' Output Via Name Property
		a = nextVia.Name
		Print #1, a; Space$(14-Len(a)); 
		
		' Output Via Type Property
		a = nextVia.Type
		Print #1, a; Space$(30-Len(a)); 

		' Output Via StartLayer Property
		a = nextVia.StartLayer
		Print #1, a; Space$(6-Len(a)); 
		
		' Output Via EndLayer Property
		a = nextVia.EndLayer
		Print #1, a; Space$(6-Len(a)); 
		
		' Output Via Attached Net Property
		a = nextVia.Net.Name
		Print #1, a; Space$(30-Len(a)); 
		
		' Output Via Drill Size Property
		a = nextVia.DrillSize
		Print #1, a; Space$(30-Len(a)); 
		
		' Output Via PositionX Property
		a = Format$(nextVia.PositionX, "#.00")
		Print #1, a; Space$(30-Len(a)); 

		' Output Via PositionY Property
		a = Format$(nextVia.PositionY, "#.00")
		Print #1, a; Space$(30-Len(a)); 

		' Output Via Plated Property
		a = nextVia.Plated
		Print #1, a; Space$(8-Len(a)); 
		
		' Output ViaTestPoint Property
		a = "Error"
		b = nextVia.TestPoint
		If (b=0) Then a = "No" 
		If (b=1) Then a = "Top" 
		If (b=2) Then a = "Bottom"
		Print #1, a; Space$(10-Len(a)); 

		' Output Via PlaneThermal Property
		a = nextVia.PlaneThermal
		Print #1, a; Space$(8-Len(a)); 

		' Output Via Glued Property
		a = nextVia.Glued
		Print #1, a; Space$(8-Len(a)); 

		Print #1, ""
	Next nextVia

	' Unlock the server
	UnlockServer

	' Close the text file
	Close #1

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
	excelApp.Range("A1").Select
	Set excelApp = Nothing
	End

noExcel:
	' Display the text file
	Shell "Notepad " & filename, 3
	
End Sub
