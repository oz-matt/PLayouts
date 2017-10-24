

Sub Main
	Begin Dialog UserDialog 310,28,"Select Component Pins by Prefix",.DialogFunc
		TextBox 130,4,60,21,.TestPointPrefix
		OKButton 200,4,100,21
		Text 10,7,120,14,"Component Prefix:",.Text1
	End Dialog
    	Dim dlg As UserDialog
    	Dialog dlg
    	
	PrefixLen = Len(dlg.TestPointPrefix$) 'Set PrefixLen to the number of characters in the test point component prefix
    	Dim NextPin As Object
		For Each NextPin In ActiveDocument.GetObjects(3, "", False) ' Search through all Pins in design
			If StrComp(Mid$(NextPin.Name,1,PrefixLen),dlg.TestPointPrefix$,1) = 0 And IsNumeric(Mid$(NextPin.Name,(PrefixLen+1),1)) Then 'Make sure that next pin is that of a component with the specified prefix
				NextPin.Selected = True ' Select next pin in design
			End If
		Next NextPin
End Sub

Function DialogFunc%(DlgItem$, Action%, SuppValue%)
   	Select Case Action%
    	Case 1 ' Dialog box initialization
       	DlgText "TestPointPrefix", "TP" ' Set the TestPointPrefix Dialog Text Box to initially read TP as default
    	End Select
End Function
