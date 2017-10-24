' Sample 13: Width Table.BAS
' 
' This sample demonstrates how to use the ProcessParameter function.
' It displays a list of possible width values. Click on one value will change
' PADS Layout current width.
'
' For more details, please refer to the PADS Layout Basic Editor Help File.
'

Dim ListboxList$(10 )

Sub Main
	ListboxList$(0) = CStr(8)
	ListboxList$(1) = CStr(10)
	ListboxList$(2) = CStr(16)
	ListboxList$(3) = CStr(24)
	ListboxList$(4) = CStr(32)
	ListboxList$(5) = CStr(48)
	ListboxList$(6) = CStr(64)
	ListboxList$(7) = CStr(128)
	Begin Dialog UserDialog 200,119,"Width",.CallbackFunc
		ListBox 10,7,180,77,ListboxList(),.ListBox1
		OKButton 50,91,90,21
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub

Function CallbackFunc%(DlgItem$, Action%, SuppValue%)
	Select Case Action%
		Case 2 ' Value changing or button pressed
			If DlgItem$ = "ListBox1" Then
				ProcessParameter(45 , "w" & ListboxList(SuppValue%))
			End If
	End Select
End Function
