' Sample 12: Select All Test Points.BAS
' 
' This sample demonstrates the Test Point feature of PADS Router.
' It selects all test points.
'
' For more details, please refer to the PADS Router Visual Basic Editor Help File.
'

Sub Main
	'Unselect all objects
	ActiveDocument.SelectObjects(blazeObjectTypeAll, "", False)
	For Each nextPin In ActiveDocument.Pins
		If nextPin.TestPoint <> 0 Then  nextPin.Selected = True
	Next nextPin
	For Each nextVia In ActiveDocument.Vias
		If nextVia.TestPoint <> 0 Then  nextVia.Selected = True
	Next nextVia
End Sub
