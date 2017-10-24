' Change Units 2.BAS
' 
' This simple prgram implements a unit change using radio buttons to 
' take care of some normally mundane tasks such as switching units.
' This is handy, for example when creating a metric part and then jumping
' back to you board and needing to switch back to another unit.
'
'

Sub Main
	currentUnit = ActiveDocument.unit
	Begin Dialog UserDialog 160,98,"Units Toggle",.CallbackFunc
		OKButton 20,70,110,21,.OK
		GroupBox 10,0,140,63,"Current Unit",.GroupBox1 'Display/Set Units Group
		OptionGroup .options
			OptionButton 30,14,60,14,"&Mils",.ButtonMils
			OptionButton 30,28,70,14,"&Inches",.ButtonInches
			OptionButton 30,42,70,14,"M&etric",.ButtonMetric
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub

Rem See DialogFunc help topic for more information.
Private Function CallbackFunc(DlgItem$, Action%, SuppValue%) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		currentUnit = ActiveDocument.unit
		Select Case currentUnit
			Case 2
				DlgValue "options",0
				ActiveDocument.unit = 2
				Debug.Print "Changed to Mils..."
			Case 3
				DlgValue "options",1
				Debug.Print "Changed to Inches"
			Case 4
				DlgValue "options",2
				Debug.Print "Changed to Metric"
			End Select
			Debug.Print currentUnit
			Debug.Print "New ballgame, eh!"
	Case 2 ' Value changing or button pressed
			If DlgItem$ <> "OK" Then
			CallbackFunc = True ' Prevent button press from closing the dialog box
			'currentUnit = ActiveDocument.unit
			Select Case SuppValue%
				Case 0
					ActiveDocument.unit = 2
					Debug.Print "Changed to Mils..."
				Case 1
					ActiveDocument.unit = 3
					Debug.Print "Changed to Inches"
				Case 2
					ActiveDocument.unit = 4
					Debug.Print "Changed to Metric"
			End Select
		End If
		Debug.Print "Case 2"
		Debug.Print SuppValue%
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem CallbackFunc = True ' Continue getting idle actions
	End Select
End Function
