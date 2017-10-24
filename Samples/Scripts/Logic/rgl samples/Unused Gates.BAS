'#Uses "RGL.bas"
Sub Main
	Dim outFile  As Integer
	
	outFile = OpenReport (DefaultFilePath & "\Unused gate list.txt")
	Print #outFile, "UNUSED ITEMS REPORT -- ";
	Print #outFile, ActiveDocument.Name; " -- ";
	Print #outFile, GetTime
	Print #outFile
	Print #outFile
	Print #outFile, "UNUSED GATE LIST"
	Print #outFile
	
	Dim unGatesInPkg() As UnusedInfoInPkg
	GetUnusedInfoInPkg ActiveDocument, unGatesInPkg
	
	'	outputing all unused gates list
	For i = 1 To UBound(unGatesInPkg)
		Set curGate = unGatesInPkg(i).unusedGatesLst(1)
		
		'	Outputing current gate. If it is a connector, then omit it.
		If (curGate.IsConnector = False) And (unGatesInPkg(i).bHaveUnusedPins(1) = False) Then
			'	outputing the current gate's name
			countOut = 0
			Columns 5, 20, 35, 50, 65
				prevComp = ""
				For j = 1 To UBound(unGatesInPkg(i).unusedGatesLst)
					Set curGate = unGatesInPkg(i).unusedGatesLst(j)

					'	Only for gates with all not connected pins					
					If unGatesInPkg(i).bAllPinsDisconnected(j) Then
						If countOut = 0 Then
							' outputing Package name with Attributes VALUE & TOLERANCE
							outString =  unGatesInPkg(i).Name
							Set attr = curGate.Component.Attributes
							If Not (attr.Item("VALUE")  Is Nothing) Then
								value = attr.Item("VALUE") 
								If value <> "" Then outString = outString + "," + value
							End If
							If Not(attr.Item("TOLERANCE") Is Nothing) Then
								tol = attr.Item("TOLERANCE")
								If tol <> "" Then outString = outString + "," +  tol
							End If
							Print #outFile, outString
						End If
						
						If prevComp = "" Then 
							prevComp = curGate.Component
						ElseIf prevComp <>  curGate.Component Then ' new line for new component
							prevComp = curGate.Component
							End_Columns
							Columns 5, 20, 35, 50, 65
						End If
						Out curGate.Name
						countOut = countOut + 1
					End If
				Next j
			End_Columns
			If countOut <> 0 Then Out
		End If
	Next i
	
	CloseReport
End Sub
