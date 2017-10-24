'Test points report

'#Uses "RGL.bas"
Sub Main
	Set doc = ActiveDocument
	outFile = OpenReport (DefaultFilePath & "\Test points report.rep")
	
	Print #outFile, "Test points List Report -- ";
	Print #outFile, doc.Name; " -- ";
	Print #outFile, GetTime
	Print #outFile, "Total test points = "; TestPoints(doc).Count; "  test point(s)"
	Print #outFile
	
	Out "Test point Name Signal Name  X          Y          Layer"
	Out "--------------------------------------------------------"

	'Search within test points
	Columns 0, 16, 29, 40, 51
		For Each nextTP In TestPoints(doc)
			'Test point name.
			'outStrings(1) = IIf(nextTP.ObjectType = ppcbObjectTypePin, nextTP, nextTP.type)
			If nextTP.ObjectType = ppcbObjectTypePin Then
				nameTP = nextTP
			Else
				nameTP = nextTP.type
			End If
			
			'Signal (net) name. If the test point is on an unused component pin, the name *NONE* is output.			
			netName = IIf(nextTP.Net Is Nothing, "*NONE*", nextTP.Net)
			
			'X coordinate for test point.
			XTP = Format2(nextTP.PositionX)
			
			'Y coordinate for test point.
			YTP = Format2(nextTP.PositionY)
	
			'Testing side for the test point: TOP or BOTTOM
			side = IIf(nextTP.TestPoint = ppcbTestPointTopLayer, "TOP", "BOTTOM")
			
			Out nameTP, netName, XTP, YTP, side
		Next nextTP
	End_Columns
	
	CloseReport
End Sub


