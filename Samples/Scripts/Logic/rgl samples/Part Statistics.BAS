'#Uses "RGL.bas"
Sub Main
	Dim outFile  As Integer
	
	outFile = OpenReport (DefaultFilePath & "\Part Statistics.txt")
	Print #outFile, "PART STATUS REPORT -- ";
	Print #outFile, ActiveDocument.Name; " -- ";
	Print #outFile, GetTime
	Print #outFile

	'	for each component in document
	Set comps = ActiveDocument.Components
	cnt = 0
	For Each nextComp In comps
		' output component's name and part's name of component with VALUE & TOLERANCE
		Columns 0, 13
			Out nextComp.Name
			
			outString = nextComp.PartType
			If Not (nextComp.Attributes.Item("VALUE")  Is Nothing) Then
				value = nextComp.Attributes.Item("VALUE") 
				If value <> "" Then outString = outString + "," + value
			End If
			If Not(nextComp.Attributes.Item("TOLERANCE") Is Nothing) Then
				tol = nextComp.Attributes.Item("TOLERANCE")
				If tol <> "" Then outString = outString + "," +  tol
			End If			
			Out outString	' output resulted string
		End_Columns
		
		' output each pin's name with type and number of sheet
		Columns 0, 26, 52
			Set pinsSort = nextComp.Pins
			pinsSort.Sort
			For Each nextPin In pinsSort
				If Not(nextPin.Net Is Nothing) Then
					outString = nextPin.Name + "["
					outString = outString + PinTyp(nextPin)  + "-"
					If nextPin.Gate Is Nothing Then
						Set minGate = nextPin.Component.Gates(1)
						For Each nextGate In nextPin.Component.Gates
							If minGate.Name > nextGate.Name Then Set minGate = nextGate
						Next nextGate
						outString = outString + SheetNumOrName(minGate.Sheet) + "]"
					Else
						outString = outString + SheetNumOrName(nextPin.Gate.Sheet) + "]"
					End If					
					outString = outString + " " + nextPin.Net.Name
					
					Out outString	' output resulted string
				End If
			Next nextPin
		End_Columns
		Print #outFile
	Next nextComp
	
	CloseReport	
End Sub
