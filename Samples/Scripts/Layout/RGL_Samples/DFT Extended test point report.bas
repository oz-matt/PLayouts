'DFT Extended test point report

'#Uses "RGL.bas"
Sub Main
	Set doc = ActiveDocument
	outFile = OpenReport (DefaultFilePath & "\DFT Extended test point.rep")
	
	Print #outFile, "=============================================================================="
	Print #outFile, "=  List All test points by Net -- ";
	Print #outFile, doc.Name; " -- ";
	Print #outFile, GetTime
	Print #outFile, "=============================================================================="	
	Columns 3, 20, 25, 30, 40, 45, 55
		For Each nextNet In doc.Nets
			If TPCnt(nextNet) <> 0 Then
				Print #outFile, "   -------------------------------------------------------"
				totalCnt = Str(TPCnt(nextNet))
				pinCnt = Str(TPPinCnt(nextNet))
				viaCnt = Str(TPViaCnt(nextNet))
				
				Out nextNet, "#TP= ", totalCnt, "#TP_PINS= ", pinCnt, "#TP_VIAS= ", viaCnt
				Print #outFile, "   -------------------------------------------------------"
				
				Columns 16, 29, 40, 50
					For Each nextTP In TestPoints(doc)
						If nextTP.Net Is Nothing Then Exit For 'The end of array of Test points is place of unused component
						
						If nextTP.Net = nextNet Then
							'Test point name.
							'outString(1) = IIf(nextTP.ObjectType = ppcbObjectTypePin, nextTP, nextTP.type)
							If nextTP.ObjectType = ppcbObjectTypePin Then
								nameTP = nextTP
							Else
								nameTP = nextTP.type
							End If
							
							'X coordinate for test point.
							XTP = Format2(nextTP.PositionX)
							
							'Y coordinate for test point.
							YTP = Format2(nextTP.PositionY)
					
							'Testing side for the test point: TOP or BOTTOM
							side = IIf(nextTP.TestPoint = ppcbTestPointTopLayer, "TOP", "BOTTOM")
							
							Out nameTP, XTP, YTP, side
						End If
					Next nextTP					
				End_Columns
				Print #outFile
			End If
		Next nextNet
	End_Columns
	Print #outFile

	Print #outFile, "=============================================================================="
	Print #outFile, "=  List All nets without test points -- ";
	Print #outFile, doc.Name; " -- ";
	Print #outFile, GetTime
	Print #outFile, "=============================================================================="	
	
	For Each nextNet In doc.Nets
		If TPCnt(nextNet) = 0 Then
			Columns 3
				Out nextNet
			End_Columns
		End If
	Next nextNet
	Print #outFile
	

	Print #outFile, "=============================================================================="
	Print #outFile, "=  List All nets And whether they have testpoints -- ";
	Print #outFile, doc.Name; " -- ";
	Print #outFile, GetTime
	Print #outFile, "=============================================================================="	
	Print #outFile
	Print #outFile, "   Signal           TP In Net?  #TP"
	Print #outFile, "   ------           ----------  ---"

	Columns 3, 20, 32
		For Each nextNet In doc.Nets
			Out nextNet, IIf(TPCnt(nextNet) = 0, "NO", "YES"), Str(TPCnt(nextNet))
		Next nextNet
	End_Columns
	Print #outFile
	
	Print #outFile, "=============================================================================="
	Print #outFile, "=  All TP On board, by Net -- ";
	Print #outFile, doc.Name; " -- ";
	Print #outFile, GetTime
	Print #outFile, "=============================================================================="	

	Print #outFile, "------------------------------------------------------"
	Print #outFile, "Signal Name    TP Name       X          Y         Side"
	Print #outFile, "------------------------------------------------------"

	'Search within test points
	Columns 0, 16, 29, 40, 50
		For Each nextTP In TestPoints(doc)
			'Signal (net) name. If the test point is on an unused component pin, the name *NONE* is output.			
			netName = IIf(nextTP.Net Is Nothing, "*NONE*", nextTP.Net)
			
			'Test point name.
			'outString(1) = IIf(nextTP.ObjectType = ppcbObjectTypePin, nextTP, nextTP.type)
			If nextTP.ObjectType = ppcbObjectTypePin Then
				nameTP = nextTP
			Else
				nameTP = nextTP.type
			End If
	
			'X coordinate for test point.
			XTP = Format2(nextTP.PositionX)
			
			'Y coordinate for test point.
			YTP = Format2(nextTP.PositionY)
	
			'Testing side for the test point: TOP or BOTTOM
			side = IIf(nextTP.TestPoint = ppcbTestPointTopLayer, "TOP", "BOTTOM")		
			
			Out netName, nameTP, XTP, YTP, side
		Next nextTP
	End_Columns
	
	CloseReport
End Sub


