' This is old version of Sample 17: Excel Part List Report.BAS
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
	Print #1, "PartType";	Space(32); 
	Print #1, "RefDes";		Space(24); 
	Print #1, "PartDecal";	Space(32); 
	Print #1, "Pins";		Space(6); 
	Print #1, "Layer";		Space(26); 
	Print #1, "Orient.";	Space(24); 
	Print #1, "X";			Space(30); 
	Print #1, "Y";			Space(29); 
	Print #1, "SMD";		Space(7); 
	Print #1, "Glued";		Space(0)
	
	' Lock server to speed up process
	LockServer

	' Go through each component in the design and output values
	For Each nextComp In ActiveDocument.Components
		Print #1, nextComp.PartType;	Space$(40-Len(nextComp.PartType)); 
		Print #1, nextComp.Name;		Space$(30-Len(nextComp.Name)); 
		Print #1, nextComp.Decal;		Space$(40-Len(nextComp.Decal)); 
		Print #1, nextComp.Pins.Count;	Space$(10-Len(nextComp.Pins.Count)); 
		Print #1, ActiveDocument.LayerName(nextComp.layer);	Space$(30-Len(ActiveDocument.LayerName(nextComp.layer))); 
		Print #1, nextComp.Orientation;	Space$(30-Len(nextComp.Orientation)); 
		Print #1, nextComp.PositionX;	Space$(30-Len(nextComp.PositionX)); 
		Print #1, nextComp.PositionY;	Space$(30-Len(nextComp.PositionY)); 
		Print #1, nextComp.IsSMD;		Space$(10-Len(nextComp.IsSMD)); 
		Print #1, nextComp.Glued;		Space$(10-Len(nextComp.Glued))
	Next nextComp

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

