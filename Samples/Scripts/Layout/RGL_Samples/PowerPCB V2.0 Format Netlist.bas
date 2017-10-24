'PowerPCB V2.0 Format Netlist

'#Uses "RGL.bas"
Sub Main
	With ActiveDocument
		outFile = OpenReport (DefaultFilePath & "\PowerPCB V2.0 Format Netlist.rep")
		
		Out "!PADS-POWERPCB-V2.0-MILS! DESIGN DATABASE ASCII FILE 1.0"

		Out "*PART*       ITEMS"
		
		Columns 0, 9
			For Each nextComp In .Components
				Out nextComp, nextComp.PartType
			Next nextComp
		End_Columns
		
		Out "*NET*"
		
		For Each sig In .Nets
			Out "*SIG* " & sig

			Between 10
				MaxCols 5
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

