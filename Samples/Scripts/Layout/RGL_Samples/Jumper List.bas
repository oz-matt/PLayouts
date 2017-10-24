' Jumper List

'#Uses "RGL.bas"
Sub Main
	With ActiveDocument
		outFile = OpenReport (DefaultFilePath & "\Jumper List.rep")
		
		Print #outFile, "Jumper List Report -- ";
		Print #outFile, .Name; " -- ";
		Print #outFile, GetTime
		
		Out "Total jumpers = " & .Jumpers.Count & "  Jumper(s)"
		Out
		
		Out "Ref.Nm              Angle     Length    X1        Y1        X2        Y2        Signal"
		Out "--------------------------------------------------------------------------------------"

		Columns 0, 20, 30, 40, 50, 60, 70, 80
			For Each jmp In .Jumpers
				nameJmp = jmp
				angle = Format(jmp.Orientation, "###.0")
				lengthJmp = Format(jmp.Length, "0.#####")
								
				Set vias = jmp.Points
				X1 = Format(vias(1).PositionX, "0.#####")
				Y1 = Format(vias(1).PositionY, "0.#####")
				X2 = Format(vias(2).PositionX, "0.#####")
				Y2 = Format(vias(2).PositionY, "0.#####")
				netJmp = jmp.Net
				Out nameJmp, angle, lengthJmp, X1, Y1, X2, Y2, netJmp
			Next jmp
		End_Columns
		
		CloseReport
	End With
End Sub

