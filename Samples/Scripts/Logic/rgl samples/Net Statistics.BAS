'#Uses "RGL.bas"
Sub Main
	Dim outFile  As Integer
	
	outFile = OpenReport (DefaultFilePath & "\Net Statistics.txt")
	Print #outFile
	Print #outFile, "NET STATUS REPORT -- ";
	Print #outFile, ActiveDocument.Name; " -- ";
	Print #outFile, GetTime
	Print #outFile

	' for each net in document
	Set nets = ActiveDocument.Nets		
	For Each nextNet In  nets
		'	output count of Sources, Loads & Undefined pins in the net
		GetSLU nextNet.Pins, sources% , loads%, undefined%
		Columns 0, 13, 28, 41
			Out nextNet.Name
			Out "SOURCES = " + Str(sources)
			Out "LOADS = " + Str(loads)
			Out "UNDEFINED = " + Str(undefined)
		End_Columns
		
		' output each pin's name with type and number of sheet
		Columns 0, 19, 38, 57
			For Each nextPin In nextNet.Pins
				If Not(nextPin.Net Is Nothing) Then
					outString = nextPin.GatePinName + "["
					outString = outString + PinTyp(nextPin)
					If nextPin.Gate Is Nothing Then
						outString = outString + "]"
					Else
						outString = outString + "-" + SheetNumOrName(nextPin.Gate.Sheet) + "]"
					End If
					
					Out outString
				End If
			Next nextPin
		End_Columns
		
		'	output additional informations
		If sources = 0 Then Print #outFile, "--- Net has no defined source.";
		If loads = 0 Then
			Print #outFile, "--- Net has no defined loads."
		Else 
			Print #outFile
		End If
		Print #outFile
	Next nextNet
	
	CloseReport	
End Sub

' Calculate count of Sources, Loads & Undefined pins in the net
Sub GetSLU(netsPins As objects, ByRef sources  As Integer, ByRef loads  As Integer, ByRef undefined  As Integer)
	sources = 0
	loads = 0
	undefined = 0
	For Each nextPin In netsPins
		Select Case nextPin.ElectricalType
			Case plogElectricalTypeUnknown 				
					undefined = undefined + 1
			Case plogElectricalTypeSource
					sources = sources + 1
			Case plogElectricalTypeBidirectional
					sources = sources + 1
					loads = loads + 1
			Case plogElectricalTypeOpenCollector
					loads = loads + 1
			Case plogElectricalTypeOrTieableSource
					sources = sources + 1
			Case plogElectricalTypeTristate
					sources = sources + 1
					loads = loads + 1
			Case plogElectricalTypeLoad
					loads = loads + 1
			Case plogElectricalTypeTerminator
					loads = loads + 1
			Case plogElectricalTypePower
					undefined = undefined + 1
			Case plogElectricalTypeGround
					undefined = undefined + 1
			Case Else 
					undefined = undefined + 1
		End Select		
	Next nextPin
End Sub

