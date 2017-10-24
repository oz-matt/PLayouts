'Parts List 1

'#Uses "RGL.bas"
Sub Main
	With ActiveDocument
		outFile = OpenReport (DefaultFilePath & "\Part List 1.rep")
		
		Print #outFile, "Parts List Report -- ";
		Print #outFile, .Name; " -- ";
		Print #outFile, GetTime
		Print #outFile, "Total components = "; .Components.Count; "  component(s)"
		Print #outFile

		Out "Reference Designation      PartType                     Logic Type"
		Out "------------------------------------------------------------------"
	
		Columns 10, 27, 60
			For Each nextComp In .Components
				Out nextComp, nextComp.PartType, nextComp.PartTypeLogic
			Next nextComp
		End_Columns
	
		CloseReport
	End With
End Sub

