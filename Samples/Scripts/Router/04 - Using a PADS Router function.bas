' Sample 04: Using a PADS Router Function.BAS
' 
' This sample demonstrates how to use a simple PADS Router OLE function.
' 
' PADS Router is an OLE Automation server and defines a set of functions,
' which can be used directly in Visual Basic. One of this function is
' the ActiveDocument (which returns the ... active document object) and
' one function available from the active document is its Name.
'
' For more details, please refer to the PADS Router Visual Basic Editor Help File.
'
Sub Main
	message = "Hi, how are you? You are now working with " & ActiveDocument.Name
	MsgBox message
End Sub
