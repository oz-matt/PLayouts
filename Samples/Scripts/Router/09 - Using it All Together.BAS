' Sample 09: Using it All Together.BAS
' 
' This sample demonstrates a real life usage of PADS Layout Visual Basic.
'
' This little script displays a dialog listing all .PCB files that are in the PADS Layout
' FileDir directory. Upon selecting one of the files in the list box, the scripts opens
' that PCB file in PADS Layout. The OK button just closes the dialog.
'
' For more details, please refer to the PADS Layout Visual Basic Editor Help File.
'
Dim ListboxList$(1000 )
Sub Main
	F$ = Dir$(DefaultFilePath & "\*.pcb")
	index = 0
	While F$ <> ""
		ListboxList$(index) = F$
		index = index + 1
		F$ = Dir$()
	Wend

	Begin Dialog UserDialog 230,329,"Available Designs",.CallbackFunc ' %GRID:10,7,1,1
		ListBox 10,7,210,280,ListboxList(),.ListBox1
		OKButton 70,301,100,21
		Text 10,287,210,14,"Click to open file in PADS Layout",.Text1
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub

Function CallbackFunc%(DlgItem$, Action%, SuppValue%)
	Select Case Action%
		Case 2 ' Value changing or button pressed
			If DlgItem$ = "ListBox1" Then
				OpenDocument(DefaultFilePath & "\" & ListboxList(SuppValue%))
			End If
	End Select
End Function
