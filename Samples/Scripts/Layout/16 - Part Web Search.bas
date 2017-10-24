' Sample 16: Part Web Search.BAS
' 
' This sample demonstrates links with Internet Explorer.
' It searches major semicon manufacturers for the currently PADS Layout 
' selected part by connecting to these manufacturers Web sites.
' Additionally, the sample demonstrates using of SelectionChange event.
' The event handler enables or disables OK button depending on PADS Layout selection
'
' For more details, please refer to the PADS Layout Basic Editor Help File.
'

Sub Main
	Begin Dialog UserDialog 260,140,"Part Web Search",.CallbackFunc ' %GRID:10,7,1,1
		Text 50,84,170,14,"",.Text1
		OKButton 40,112,90,21,.OK
		CancelButton 140,112,90,21
		OptionGroup .Manufacturer
			OptionButton 60,14,140,14,"Texas Instrument",.OptionButton1
			OptionButton 60,35,90,14,"National",.OptionButton2
			OptionButton 60,56,90,14,"FairChild",.OptionButton3
	End Dialog
	Dim dlg As UserDialog
	If Dialog(dlg) = -1 Then 'OK was pressed
		Set objs = ActiveDocument.GetObjects(ppcbObjectTypeComponent, "", True)
		compName = objs.Item(1).PartType	
		Set ie = CreateObject("InternetExplorer.Application")
		ie.Visible = True
		Select Case dlg.Manufacturer
			Case 0
				ie.Navigate("http://www-s.ti.com/sc/docs/psheets/pids2.htm")
				SendKeys compName & "~"
			Case 1
				ie.Navigate("http://www.national.com/search/search.cgi/company?keywords=" & compName)
			Case 2
				ie.Navigate("http://www.fairchildsemi.com/search/search.cgi/design?keywords=" & compName & "&x=76&y=11")
			End Select
	End If
End Sub

Rem See DialogFunc help topic for more information.
Private Function CallbackFunc(DlgItem$, Action%, SuppValue%) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		Document_SelectionChange
	Case 2 ' Value changing or button pressed
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem CallbackFunc = True ' Continue getting idle actions
	End Select
End Function

Public Sub Document_SelectionChange()
	Set objs = ActiveDocument.GetObjects(ppcbObjectTypeComponent, "", True)
	DlgEnable "OK", False
	If objs.Count = 1 Then
		compName = objs.Item(1).PartType	
		DlgText "Text1", "Selected Part: " & compName
		DlgEnable "OK", True
	ElseIf objs.Count > 1 Then
		DlgText "Text1", "Multiple Selection"
	Else
		DlgText "Text1", "Select a part"
	End If
End Sub
