' Sample 11: Select by Pin Count.BAS
' 
' This sample demonstrates how to select PADS Logic objects based on custom criterias.
' It prompts the user for a pin count and select all PADS Logic components with that pin count.
'
' For more details, please refer to the PADS Logic Basic Editor Help File.
'

Sub Main
	Begin Dialog UserDialog 230,63,"Select by Pin Number",.CallbackFunc
		TextBox 20,7,190,21,.TextBox1
		OKButton 50,35,130,21,.OkButton1
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub
Function CallbackFunc%(DlgItem$, Action%, SuppValue%)
	Select Case Action%
		Case 2 ' Value changing or button pressed
			If DlgItem$ = "OkButton1" Then
				If DlgText("TextBox1") <> "" Then
					ActiveDocument.SelectObjects(plogObjectTypeAll, "", False)
					Dim nextComp As Object
					For Each nextComp In ActiveDocument.Components
						If nextComp.Pins.Count = CLng(DlgText("TextBox1")) Then
							nextComp.Selected = True
						End If
					Next NextComp
				End If
			End If
			LocateSelection
	End Select
End Function

Sub LocateSelection
	'look for selection in current sheet: if it has selection then do nothing
	If ActiveDocument.ActiveSheet.GetObjects(plogObjectTypeComponent,,True).Count > 0 Then Exit Sub
	For Each Sheet In ActiveDocument.Sheets
		If Sheet.GetObjects(plogObjectTypeComponent,,True).Count > 0 Then 
			Sheet.Activate
			Exit Sub
		End If
	Next
End Sub
