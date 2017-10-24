' DIE report

'#Uses "RGL_Samples\\RGL.bas"

Public Function DesignUnits () As String
	Select Case ActiveDocument.unit
		Case ppcbUnitMils 				
				DesignUnits = "Mill"
		Case ppcbUnitInch 				
				DesignUnits = "Inch"
		Case ppcbUnitMetric 				
				DesignUnits = "Metric"
		Case Else 
				DesignUnits = "Unknown"
	End Select	
End Function
Public Function PadShape (pad As Object) As String
	Select Case pad.Shape
		Case ppcbBondPadShapeUnknown 				
				PadShape = "Unknown"
		Case ppcbBondPadShapeRectangle 				
				PadShape = "Rect"
		Case ppcbBondPadShapeOval 				
				PadShape = "Oval"
		Case Else 
				PadShape = "Unknown"
	End Select	
End Function
Public Function PadEdge (pad As CBP) As String
	Select Case pad.Edge
		Case ppcbBondPadEdgeUnknown 				
				PadEdge = "Unknown"
		Case ppcbBondPadEdgeLeft
				PadEdge = "Left"
		Case ppcbBondPadEdgeTop 				
				PadEdge = "Top"
		Case ppcbBondPadEdgeRight 				
				PadEdge = "Right"
		Case ppcbBondPadEdgeBottom 				
				PadEdge = "Bottom"
		Case Else 
				PadEdge = "Unknown"
	End Select	
End Function
Public Function WBLength (wb As Wirebond) As Double
	WBLength = Sqr ((wb.StartX - wb.EndX)*(wb.StartX - wb.EndX) + (wb.StartY - wb.EndY)*(wb.StartY - wb.EndY))
End Function

Sub Main
	With ActiveDocument
		outFile = OpenReport (DefaultFilePath & "\die_report.rep")
		
		Print #outFile, "Die Report -- ";
		Print #outFile, .Name; " - ";
		Print #outFile, GetTime
		
		Print #outFile
		Print #outFile, "Units: "; DesignUnits
		
		For Each cmp In .Components
			If (cmp.IsDiePart) Then
				Print #outFile
				Print #outFile
				Print #outFile, "Die Part "; cmp.Name; ", "; cmp.PartType
				Print #outFile, "-------------------"
				Print #outFile
				
				Print #outFile, "DieLength:          "; Format (cmp.DieLength, "0.#####")
				Print #outFile, "DieWidth:           "; Format (cmp.DieWidth, "0.#####")
				Print #outFile, "DieHeight:          "; Format (cmp.DieHeight, "0.#####")
				Print #outFile
				
				Print #outFile, "Bond Wire Rules"
				Print #outFile, "       MinLength:   "; Format (cmp.WireBondRulesLengthMinimum, "0.#####")
				Print #outFile, "       MaxLength:   "; Format (cmp.WireBondRulesLengthMaximum, "0.#####")
				Print #outFile, "       MaxAngle:    "; Format (cmp.WireBondRulesAngleMaximum, "0.#####")
				Print #outFile, "       WireToWire:  "; Format (cmp.WireBondRulesClearanceWireToWire, "0.#####")
				Print #outFile, "       WireToPad:   "; Format (cmp.WireBondRulesClearanceWireToPad, "0.#####")
				Print #outFile
				
				Print #outFile, "Component Bond Pads"
				Print #outFile
				Print #outFile, " Name          Xcoord         Ycoord         Shape    Length     Width    Function   Edge      Layer"
				Print #outFile
				
				Columns 0, 15, 30, 45, 55, 65, 75, 85, 95
				Dim nextCBP As CBP
				For Each nextCBP In cmp.CBPs
					nameCBP = nextCBP
					X = Format (nextCBP.PositionX, "0.#####")
					Y = Format (nextCBP.PositionY, "0.#####")
					shape = PadShape (nextCBP)
					length = Format (nextCBP.Length, "0.#####")
					width = Format (nextCBP.Width, "0.#####")
					cbpFunction = nextCBP.Function
					edge = PadEdge (nextCBP)
					layer = nextCBP.layer
					Out nameCBP, X, Y, shape, length, width, cbpFunction, edge, layer
				Next nextCBP
				End_Columns
				
				Print #outFile
				Print #outFile, "Substrate Bond Pads"
				Print #outFile
				Print #outFile, " Name          Xcoord         Ycoord         Shape    Length     Width  Orientation Function  Tier       Layer"
				Print #outFile
				
				Columns 0, 15, 30, 45, 55, 65, 75, 85, 95, 105
				Dim nextSBP As Object
				For Each nextSBP In cmp.SBPs
					nameSBP = nextSBP
					X = Format (nextSBP.PositionX, "0.#####")
					Y = Format (nextSBP.PositionY, "0.#####")
					shape = PadShape (nextSBP)
					length = Format (nextSBP.Length, "0.#####")
					width = Format (nextSBP.Width, "0.#####")
					angle = Format(nextSBP.Orientation, "#####.0")
					sbpFunction = nextSBP.Function
					tier = nextSBP.Tier
					layer = nextSBP.layer
					Out nameSBP, X, Y, shape, length, width, angle, cbpFunction, tier, layer
				Next nextSBP
				End_Columns

				
				Print #outFile
				Print #outFile, "Wirebonds"
				Print #outFile
				Print #outFile, "Start Pad      End Pad              Start Point                    End Point              Start Offset         End Offset         Angle          Length"
				Print #outFile
				
				Columns 0, 15, 30, 45, 60, 75, 90, 100, 110, 120, 130, 145
				Dim nextWB As Wirebond
				For Each nextWB In cmp.Wirebonds
					pad1 = nextWB.StartPad
					pad2 = nextWB.EndPad
					X1 = Format (nextWB.StartX, "0.#####")
					Y1 = Format (nextWB.StartY, "0.#####")
					X2 = Format (nextWB.EndX, "0.#####")
					Y2 = Format (nextWB.EndY, "0.#####")
					X3 = Format (nextWB.StartOffsetX, "0.#####")
					Y3 = Format (nextWB.StartOffsetY, "0.#####")
					X4 = Format (nextWB.EndOffsetX, "0.#####")
					Y4 = Format (nextWB.EndOffsetY, "0.#####")
					angle = Format(nextWB.angle, "0.#####")
					length = Format(WBLength (nextWB), "0.####")
					Out pad1, pad2, X1, Y1, X2, Y2, X3, Y3, X4, Y4, angle, length
				Next nextWB
				End_Columns
			End If
		Next cmp
		
		CloseReport
	End With
End Sub

