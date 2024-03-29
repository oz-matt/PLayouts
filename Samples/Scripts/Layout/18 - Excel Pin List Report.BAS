' Sample 18: Excel Pin List Report.BAS
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
	Print #1, "Name";		Space(26); 
	Print #1, "Component";	Space(21); 
	Print #1, "Net";		Space(28); 
	Print #1, "Drill";		Space(25); 
	Print #1, "X";			Space(29); 
	Print #1, "Y";			Space(29); 
	Print #1, "Plated";		Space(2); 
	Print #1, "TestPoint";	Space(1); 
	Print #1, "Thermal";	Space(2); 
	Print #1, "Glued";		Space(2); 
	Print #1, "SMD";		Space(0)

	' Lock server to speed up process
	LockServer

	' Go through each pin in the design and output values
	For Each nextPin In ActiveDocument.Pins

		' Output Pin Name Property
		a = nextPin.Name
		Print #1, a; Space$(30-Len(a)); 

		' Output Pin Owning Component Property
		a = nextPin.Component.Name
		Print #1, a; Space$(30-Len(a)); 

		' Output Pin Attached Net Property
		Set nnet = nextPin.Net
		If (nnet Is Nothing) Then a = "N/A" Else a = nnet.Name 
		Print #1, a; Space$(30-Len(a)); 
		
		' Output Pin Drill Size Property
		a = nextPin.DrillSize
		Print #1, a; Space$(30-Len(a)); 
		
		' Output Pin PositionX Property
		a = Format$(nextPin.PositionX, "#.00")
		Print #1, a; Space$(30-Len(a)); 

		' Output Pin PositionY Property
		a = Format$(nextPin.PositionY, "#.00")
		Print #1, a; Space$(30-Len(a)); 
		
		' Output Pin Plated Property
		a = nextPin.Plated
		Print #1, a; Space$(8-Len(a)); 

		' Output Pin TestPoint Property
		a = "Error"
		b = nextPin.TestPoint
		If (b=0) Then a = "No" 
		If (b=1) Then a = "Top" 
		If (b=2) Then a = "Bottom"
		Print #1, a; Space$(10-Len(a)); 

		' Output Pin PlaneThermal Property
		a = nextPin.PlaneThermal
		Print #1, a ; Space$(9-Len(a)); 

		' Output Pin Glued Property
		a = nextPin.Glued
		Print #1, a; Space$(7-Len(a)); 
		
		' Output Pin IsSMD Property
		a = nextPin.IsSMD
		Print #1, a
		
	Next nextPin

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

