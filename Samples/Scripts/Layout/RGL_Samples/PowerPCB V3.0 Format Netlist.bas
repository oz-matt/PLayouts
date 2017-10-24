'PowerPCB V3.0 Format Netlist

'#Uses "RGL.bas"
Sub Main
	With ActiveDocument
		outFile = OpenReport (DefaultFilePath & "\PowerPCB V3.0 Format Netlist.rep")
		
		Out 
		Out "!PADS-POWERPCB-V3.0-MILS! DESIGN DATABASE ASCII FILE 2.0"
		Out
		Out "*PART*       ITEMS"
		
		Columns 0, 17
			For Each nextComp In .Components
				Out nextComp, nextComp.PartType
			Next nextComp
		End_Columns
		
		Out "*NET*"
		Out
		
		For Each sig In .Nets
			Out "*SIG* " & sig

			Between 22
				MaxCols 4
				For Each nextPin In sig.Pins
					Out nextPin
				Next nextPin
			End_Between
		Next sig
		
		Out
		Print #outFile, "*END*     OF ASCII OUTPUT FILE"
		
		CloseReport		
	End With
End Sub

