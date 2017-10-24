' Sample 15: Select All Test Points.BAS
' 
' This sample demonstrates the Test Point feature of PADS Layout.
' It selects all test points.
'
' For more details, please refer to the PADS Layout Basic Editor Help File.
'

Sub Main
	'Unselect all objects
	ActiveDocument.SelectObjects(ppcbObjectTypeAll, "", False)
	For Each nextPin In ActiveDocument.Pins
		If nextPin.TestPoint <> 0 Then  nextPin.Selected = True
	Next nextPin
	For Each nextVia In ActiveDocument.Vias
		If nextVia.TestPoint <> 0 Then  nextVia.Selected = True
	Next nextVia
End Sub
