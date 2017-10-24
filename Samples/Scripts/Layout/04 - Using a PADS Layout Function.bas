' Sample 04: Using a PADS Layout Function.BAS
' 
' This sample demonstrates how to use a simple PADS Layout OLE function.
' 
' PADS Layout is an OLE Automation server and defines a set of functions,
' which can be used directly in Basic. One of this function is
' the ActiveDocument (which returns the ... active document object) and
' one function available from the active document is its Name.
'
' For more details, please refer to the PADS Layout Basic Editor Help File.
'
Sub Main
	message = "Hi, how are you? You are now working with " & ActiveDocument.Name
	MsgBox message
End Sub
