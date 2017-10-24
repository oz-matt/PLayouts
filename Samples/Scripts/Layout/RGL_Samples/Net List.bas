'Net List 

'#Uses "RGL.bas"
Sub Main
	With ActiveDocument
		outFile = OpenReport (DefaultFilePath & "\Net List.rep")
		
		Print #outFile, "NET LIST REPORT -- ";
		Print #outFile, .Name; " -- ";
		Print #outFile, GetTime
		Print #outFile

		For Each sig In .Nets
			Out "*SIG " & sig

			Between 10
				MaxCols 4
				Set pins = sig.Pins
				For i = 1 To pins.Count
					Out pins(i) & "." & PinTyp(pins(i))
				Next i
			End_Between

			Out
		Next sig
		
		CloseReport		
	End With
End Sub

