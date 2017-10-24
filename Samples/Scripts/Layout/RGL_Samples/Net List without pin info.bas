'Net List w/o pin info

'#Uses "RGL.bas"
Sub Main
	With ActiveDocument
		outFile = OpenReport (DefaultFilePath & "\Net list without pin info.rep")
		
		Print #outFile, "NET LIST REPORT -- ";
		Print #outFile, .Name; " -- ";
		Print #outFile, GetTime
		Print #outFile

		For Each sig In .Nets
			Out "*SIG " & sig
			
			Between 24
				MaxCols 4
				For Each nextPin In sig.Pins
					Out nextPin
				Next nextPin	
			End_Between
			
			Out
		Next sig
		
		CloseReport		
	End With
End Sub

