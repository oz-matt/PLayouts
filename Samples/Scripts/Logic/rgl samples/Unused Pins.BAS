'#Uses "RGL.bas"
Sub Main
	Dim outFile  As Integer
	
	outFile = OpenReport (DefaultFilePath & "\Unused pins list.txt")
	Print #outFile, "UNUSED ITEMS REPORT -- ";
	Print #outFile, ActiveDocument.Name; " -- ";
	Print #outFile, GetTime
	Print #outFile
	Print #outFile
	Print #outFile, "UNUSED PINS LIST"
	Print #outFile
	
	Dim unGatesInPkg() As UnusedInfoInPkg	
	GetUnusedInfoInPkg ActiveDocument, unGatesInPkg
	
	' for each package with unused gates
	For i = 1 To UBound(unGatesInPkg)
		' outputing Package name with Attributes VALUE & TOLERANCE
		outString =  unGatesInPkg(i).Name
		Set attr = unGatesInPkg(i).unusedGatesLst(1).Component.Attributes
		If Not (attr.Item("VALUE")  Is Nothing) Then
			value = attr.Item("VALUE") 
			If value <> "" Then outString = outString + "," + value
		End If
		If Not(attr.Item("TOLERANCE") Is Nothing) Then
			tol = attr.Item("TOLERANCE")
			If tol <> "" Then outString = outString + "," +  tol
		End If
		Out outString	
		 
		' outputing the current gate's name
		Columns 5, 20, 35, 50, 65
			prevComp = ""
			For j = 1 To UBound(unGatesInPkg(i).unusedGatesLst)	' for each unused gate
				Set curGate = unGatesInPkg(i).unusedGatesLst(j)

				' Output new line for new component
				If prevComp = "" Then 
					prevComp = curGate.Component
				ElseIf prevComp <>  curGate.Component Then ' new line for new component
					prevComp = curGate.Component
					End_Columns
					Columns 5, 20, 35, 50, 65
				End If
				
				If unGatesInPkg(i).bHaveUnusedPins(j) Then	' Output all pins not belong any gate and not a sigmnal or not a power
					For Each nextPin In curGate.Component.Pins
						If (nextPin.Gate Is Nothing) And (nextPin.Net Is Nothing) Then
							Out nextPin.Name
						End If
					Next nextPin				
				Else	' output pins of used and unused gates
					For Each nextPin In curGate.Pins
						'	If current pin is not connected
						If nextPin.Net Is Nothing Then
							Out nextPin.Name
						End If
					Next nextPin
				End If
				
			Next j
		End_Columns
		Out
	Next i
	
	CloseReport
End Sub
