' Sample 12: Move by Pin Count.BAS
' 
' This sample demonstrates how to use the ProcessCommand function.
' It prompts the user for a pin count, selects all PADS Layout components 
' with that pin count, and starts a dynamic move of the selected group.
'
' For more details, please refer to the PADS Layout Basic Editor Help File.
'
Sub Main
	Begin Dialog UserDialog 230,63,"Move by Pin Number",.CallbackFunc
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
					ActiveDocument.SelectObjects(ppcbObjectTypeAll, "", False)
					Dim nextComp As Object
					For Each nextComp In ActiveDocument.Components
						If nextComp.Pins.Count = CLng(DlgText("TextBox1")) Then
							nextComp.Selected = True
						End If
					Next nextComp
				End If
				ProcessCommand(58) ' MOVE COMMAND
			End If
	End Select
End Function
