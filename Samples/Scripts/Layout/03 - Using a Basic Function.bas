' Sample 03: Using a VB Function.BAS
' 
' This sample demonstrates how to use concatenated strings.
' 
' The "message" variable is assigned a string which is calculated
' as a static string concatenated with the current date (using the 
' Visual Basic Date function). The concatenation of strings is done 
' using the Visual Basic & operator.
'
' For more details, please refer to the PADS Layout Basic Editor Help File.
'
Sub Main
	message = "Hi, how are you? Today's date is " & Date
	MsgBox message
End Sub
