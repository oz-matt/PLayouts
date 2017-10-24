' Sample 08: Using a Custom Dialog3.BAS
' 
' This sample demonstrates using dialog callbacks.
'
' A dialog callback funtion is a function which receives events occuring in the
' dialog. You implement the reaction to these events in the dialog callback
' function that we called CallbackFunc.
'
' For more details, please refer to the PADS Logic Basic Editor Help File.
'
Sub Main
	Begin Dialog UserDialog 260,98,"Sample8",.CallbackFunc ' %GRID:10,7,1,1
		Text 10,7,140,14,"Enter a Value Here:",.Text1
		TextBox 10,21,140,21,.TextBox1
		CheckBox 10,49,90,14,"Check1",.CheckBox1
		CheckBox 10,63,90,14,"Check2",.CheckBox2
		CheckBox 10,77,90,14,"Check3",.CheckBox3
		OKButton 160,14,90,21
		PushButton 160,42,90,21,"Apply",.PushButton1
		PushButton 160,70,90,21,"Close",.PushButton2
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
	
End Sub

Rem See DialogFunc help topic for more information.
Private Function CallbackFunc(DlgItem$, Action%, SuppValue%) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText "TextBox1", "123456789"
		
	Case 2 ' Value changing or button pressed
		Rem CallbackFunc = True ' Prevent button press from closing the dialog box
		If DlgItem$ = "OK" Or DlgItem$ = "PushButton1" Then
			MsgBox "You entered " & DlgText$("TextBox1") & ". Checks are (" & DlgValue("CheckBox1") & ", " & DlgValue("CheckBox2") & ", " & DlgValue("CheckBox3") & "). Thank you."
			If DlgItem$ = "PushButton1" Then  CallbackFunc = True
		End If
		
		
	Case 3 ' TextBox or ComboBox text changed
	
	Case 4 ' Focus changed
	
	Case 5 ' Idle
		Rem CallbackFunc = True ' Continue getting idle actions
	End Select
End Function
