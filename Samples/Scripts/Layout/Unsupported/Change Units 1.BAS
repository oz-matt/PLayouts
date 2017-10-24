' Change Units 1.BAS
' 
' This sample demontrates how to create some simple push buttons to 
' take care of some normally mundane tasks such as switching units.
' This is handy, for example when creating a metric part and then jumping
' back to you board and needing to switch back to another unit.
'
'

Sub Main
	currentUnit = ActiveDocument.unit
	Begin Dialog UserDialog 280,49,"Units Toolbar",.CallbackFunc
		PushButton 190,0,80,21,"Metric",.ButtonMetric 'Button for Selecting Metric Units
		PushButton 10,0,80,21,"Mils.",.ButtonMils 'Button for Selecting Mils Units
		PushButton 100,0,80,21,"Inches",.ButtonInches 'Button for Selecting Inches Units
		Text 110,35,70,14,"",.Text1
		OKButton 190,28,80,21,.OK
		Text 10,35,90,14,"Current Units:",.Text2 'Just a label for the dialog box
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub

Rem See DialogFunc help topic for more information.
Private Function CallbackFunc(DlgItem$, Action%, SuppValue%) As Boolean
	'Dim unitName$(100)
	Select Case Action%
	Case 1 ' Dialog box initialization
		currentUnit = ActiveDocument.unit
		If currentUnit = 2 Then
			unitName$ = "Mils"
		ElseIf currentUnit = 3 Then
			unitName$ = "Inches"
		ElseIf urrentUnit = 4 Then
			unitName$ = "Metric"
		End If
		DlgText "Text1", unitName
	Case 2 ' Value changing or button pressed
		If DlgItem$ <> "OK" Then
			CallbackFunc = True ' Prevent button press from closing the dialog box
				currentUnit = ActiveDocument.unit
			DlgText "Text1", "Current Unit " & currentUnit
			Select Case DlgItem$
				Case "ButtonMetric"
					DlgText "Text1","Metric"
					ActiveDocument.unit = 4
				Case "ButtonMils"
					DlgText "Text1","Mils"
					ActiveDocument.unit = 2
				Case "ButtonInches"
					DlgText "Text1","Inches"
					ActiveDocument.unit = 3
			End Select
		End If
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem CallbackFunc = True ' Continue getting idle actions
	End Select
End Function
