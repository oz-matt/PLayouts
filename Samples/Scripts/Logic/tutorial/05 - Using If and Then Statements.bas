' Sample 05: Using If and Then Statements.BAS
' 
' This sample demonstrates how to use the If, Then Else Basic keywords.
'
' For conditional coding, the If, Then, Else keywords are used. "If" defines
' the condition, "Then" defines what happens when the condition is True, 
' "Else" defines what happens when the condition is False. Note that this
' sample also shows an extended version of the MsgBox Basic function.
'
' For more details, please refer to the PADS Logic Basic Editor Help File.
'
Sub Main
	message = "Hi, how are you? Do you like this Basic scripting stuff?"
	answer = MsgBox(message, vbYesNoCancel)
	If answer = vbYes Then
		MsgBox "We are glad you do!"
	End If
	If answer = vbNo Then
		MsgBox "Sorry..."
	End If
	If answer = vbCancel Then
		MsgBox "That was a Yes/No question... :-)"
	End If
End Sub
