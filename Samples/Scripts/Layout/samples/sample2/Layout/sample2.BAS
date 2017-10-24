'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SAMPLE2.BAS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This is a part of the PADS-PowerPCB OLE Automation server SAMPLE2 sample.
' Copyright (C) 2003 Mentor Graphics Corp.
' All rights reserved.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This source code is only intended as a supplement to the PADS-PowerPCB  
' Automation Server API Help file.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Main
	Begin Dialog UserDialog 440,182,"Sample2 (PowerPCB Basic Version)",.CallbackFunc ' %GRID:10,7,1,1
		Text 10,7,420,14,"PowerPCB Document Name",.Text1
		TextBox 10,21,420,21,.EdtDocname
		TextBox 10,63,420,21,.EdtCompCount
		TextBox 10,105,420,21,.EdtSelCompCount
		PushButton 220,133,100,21,"Refresh",.BtnRefresh
		PushButton 330,133,100,21,"Close",.BtnClose
		Text 10,161,410,14,"Connected to PowerPCB.",.Status
		Text 10,49,420,14,"PowerPCB Component Count",.Text2
		Text 10,91,420,14,"PowerPCB Selected Component Count",.Text3
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub

Rem See DialogFunc help topic for more information.
Private Function CallbackFunc(DlgItem$, Action%, SuppValue%) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		CallbackFunc = True ' Prevent button press from closing the dialog box
		If DlgItem$ = "BtnClose" Then CallbackFunc = False
		If DlgItem$ = "BtnRefresh" Then 
			DlgText "EdtDocName",ActiveDocument.Name
			DlgText "EdtCompCount",ActiveDocument.Components.Count & " components object(s)"
			DlgText "EdtSelCompCount",ActiveDocument.GetObjects(ppcbObjectTypeComponent, "", True).Count & " selected component object(s)"
		End If
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem CallbackFunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Public Sub Application_OpenDocument(Doc As PowerPCB.Document)
	CallbackFunc("BtnRefresh", 2, 0) 
End Sub

Public Sub Document_SelectionChange()
	CallbackFunc("BtnRefresh", 2, 0) 
End Sub




