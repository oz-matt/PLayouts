Dim ListNets$(10000)
Dim index%

Sub Main
	Dim nextNet As Object
	
	'Lock Serever to increase performance
	LockServer
	
	'find what nets are assigned to planes,add to list
	For Each nextNet In ActiveDocument.Nets
		If nextNet.Power = True Then  
			ListNets$(index) =  nextNet.Name 
			index = index +1
		End If
	Next nextNet
	
	'abort if no plane nets exist
	If index = 0 Then 
		MsgBox "Sorry, no plane nets detected",vbExclamation,"Plane Net Thermal Tool"
		'Do not forget to unlock the server upon end of program!!
		UnlockServer
		End
	End If
	
	'Clear Current Selections in PowerPCB
	ActiveDocument.SelectObjects(ppcbObjectTypeAll,"",False)
	
	'my dialog
	Begin Dialog UserDialog 350,126,"Turn On Plane Thermals",.Dialog_Function ' %GRID:10,7,1,1
		GroupBox 10,7,320,63,"Choose Plane Net",.ChooseNet
		DropListBox 30,35,280,70,ListNets(),.SelectPlaneNet
		CheckBox 150,21,170,14,"Select All Plane Nets",.SelectAllNets
		PushButton 20,98,90,21,"OK",.Ok
		PushButton 130,98,90,21,"App&ly",.Apply
		PushButton 240,98,90,21,"Cancel",.Cancel
		Text 20,77,310,14,"Text1",.StatusField
	End Dialog
	
	'define dialog as object
	Dim dlg As UserDialog

	'invoke dialog
	Dialog dlg
	
	'Clear Current Selections in PowerPCB - leave program cleanly
	ActiveDocument.SelectObjects(ppcbObjectTypeAll,"",False)

	'Do not forget to unlock the server upon end of program!!
	UnlockServer
	
End Sub

Rem See DialogFunc help topic for more information.
Private Function Dialog_Function(DlgItem$, Action%, SuppValue&) As Boolean
	Dim nextPin As Object
	Dim nextVia As Object
	
	Select Case Action%
	Case 1'dialog initialize
	
			'set dialog initial default behaviour
			DlgValue "SelectAllNets", 1
			DlgEnable "SelectPlaneNet", False
			DlgFocus "Ok"
			DlgText "StatusField", ""
			
	Case 2 ' Value changing or button pressed
				
			If DlgItem$ = "Apply"  Or Dlgitem$ = "Ok" Then
				
				'Clear Current Selections in PowerPCB
				ActiveDocument.SelectObjects(9999,"",False)

				DlgText "StatusField", "Selecting pins......."

				'sub to decide what pins to select
				SelectThePins

				DlgText "StatusField", "Updating thermals on pins......."

				'sub to set all pins to have thermals
				TurnOnPinThermals

				'Clear Current Selections in PowerPCB
				ActiveDocument.SelectObjects(9999,"",False)

				DlgText "StatusField", "Selecting vias......."
				
				'sub to decide what vias to select
				SelectTheVias

				DlgText "StatusField", "Updating thermals on vias......."

				'sub to set all vias to have thermals
				TurnOnViaThermals

				'Clear Current Selections in PowerPCB
				ActiveDocument.SelectObjects(9999,"",False)

				DlgText "StatusField", "Operation complete"

				If dlgitem$ = "Apply" Then
					'keep dialog open
					Dialog_Function = True
				End If
			End If
									
			'enable/disable net selection combo box depending on "select all" checkbox
			If DlgValue ("SelectAllNets") = 1 Then
				DlgEnable "SelectPlaneNet", False
			Else
				DlgEnable "SelectPlaneNet", True
			End If
			
	End Select
	
End Function

Sub SelectThePins

	If DlgValue ("SelectAllNets") = 1 Then
		'If all plane net pins are to be selected do the following
		For EachNet = 0 To (index-1)
				For Each nextPin In ActiveDocument.Nets(ListNets(EachNet)).Pins
					'do not select pins that already have a thermal
					If nextPin.PlaneThermal = False Then
						ActiveDocument.SelectObjects(ppcbObjectTypePin,nextPin,True)
					End If
				Next nextPin
		Next EachNet
	Else
		'if only a specific plane net is selected do the following
		For Each nextPin In ActiveDocument.Nets(ListNets(DlgValue ("SelectPlaneNet"))).Pins
			'do not select pins that already have a thermal
			If nextPin.PlaneThermal = False Then
				ActiveDocument.SelectObjects(ppcbObjectTypePin,nextPin,True)
			End If
		Next nextPin
	End If

End Sub

Sub SelectTheVias

	If DlgValue ("SelectAllNets") = 1 Then
		'If all plane net vias are to be selected do the following
		For EachNet = 0 To (index-1)
				For Each nextVia In ActiveDocument.Nets(ListNets(EachNet)).Vias
					'do not select vias that already have a thermal
					If nextVia.PlaneThermal = False Then
						ActiveDocument.SelectObjects(ppcbObjectTypeVia,nextVia,True)
					End If
				Next nextVia
		Next EachNet
	Else
		'if only a specific plane net is selected do the following
		For Each nextVia In ActiveDocument.Nets(ListNets(DlgValue ("SelectPlaneNet"))).Vias
			'do not select vias that already have a thermal
			If nextVia.PlaneThermal = False Then
				ActiveDocument.SelectObjects(ppcbObjectTypeVia,nextVia,True)
			End If
		Next nextVia
	End If
	
End Sub

Sub TurnOnPinThermals
	'The following are undocumented commands that allow the passing of 
	'PPCB macro commands and parameters through VB
	'The commented lines are the lines recorded in a PPCB macro

	'Set dialog from coming up on screen - undocumented command
	Preference("InteractiveMode") = False

	'PARAMETER Integer 3701 1 # UndoRedoAvailableParameterID
	ProcessParameter(3701,1)
	'Command 63 # CLI_EDIT_QUERY_MODIFY
	 ProcessCommand(63)
	'PARAMETER Integer 6054 1 # AssOptSetLevel
	ProcessParameter(6054,1)
	'PARAMETER Integer 3701 0 # UndoRedoAvailableParameterID   
	ProcessParameter(3701,0)   
    'Command 220 # CLI_INQUERY
	ProcessCommand(220)
	'PARAMETER Integer 40 15 # QueryDlgParameterID
	ProcessParameter(40,15)
	'PARAMETER Integer 6050 1 # AssOptEnforceCDWI
	ProcessParameter(6050,1)
   'PARAMETER VOID 6042  # PerformGrpParameterID
	ProcessParameter(6042)
	'PARAMETER INTEGER 3701 1 # UndoRedoAvailableParameterID
	ProcessParameter(3701,1)
	'PARAMETER String 41 "6 5" # QueryDlgParameterModifyID
	ProcessParameter(41,"6 5")
	'PARAMETER String 41 "5 Sthermal=1" # QueryDlgParameterModifyID
	ProcessParameter(41,"5 Sthermal=1" )
	'PARAMETER String 39 "5 Sthermal=1" # 
	ProcessParameter(39,"5 Sthermal=1" )
	'PARAMETER VOID 1752  # UpdateChangesForStatusParameterID
	ProcessParameter(1752)
	'PARAMETER Integer 40 35 # QueryDlgParameterID
	ProcessParameter(40,35)

	'Restore normal dialog interaction
	Preference("InteractiveMode") = True

End Sub

Sub TurnOnViaThermals
	'The following are undocumented commands that allow the passing of 
	'PPCB macro commands and parameters through VB
	'The commented lines are the lines recorded in a PPCB macro

	'Set dialog from coming up on screen - undocumented command
	Preference("InteractiveMode") = False
	
	'PARAMETER Integer 3701 1 # UndoRedoAvailableParameterID
	ProcessParameter(3701,1)
	'Command 63 # CLI_EDIT_QUERY_MODIFY
	 ProcessCommand(63)
	'PARAMETER Integer 6054 1 # AssOptSetLevel
	ProcessParameter(6054,1)
	'PARAMETER Integer 3701 0 # UndoRedoAvailableParameterID   
	ProcessParameter(3701,0)   
    'Command 220 # CLI_INQUERY
	ProcessCommand(220)
	'PARAMETER Integer 40 14 # QueryDlgParameterID
	ProcessParameter(40,14)
   'Command 222 # CLI_INQUERY_VIA
   	ProcessCommand(222)
   'PARAMETER Integer 6050 1 # AssOptEnforceCDWI	
	ProcessParameter(6050,1)   
	'PARAMETER VOID 6042  # PerformGrpParameterID
	ProcessParameter(6042)   
	'PARAMETER INTEGER 3701 1 # UndoRedoAvailableParameterID
	ProcessParameter(3701,1)   
	'PARAMETER String 41 "6 4" # QueryDlgParameterModifyID
	ProcessParameter(41,"6 4")
	'PARAMETER String 41 "4 Sthermal=1" # QueryDlgParameterModifyID
	ProcessParameter(41,"4 Sthermal=1" )
	'PARAMETER VOID 1752  # UpdateChangesForStatusParameterID	
	ProcessParameter(1752)   
	'PARAMETER String 39 "4 Sthermal=1" # 
	ProcessParameter(39,"4 Sthermal=1" )
	'PARAMETER VOID 1752  # UpdateChangesForStatusParameterID
	ProcessParameter(1752)   
	'PARAMETER Integer 40 34 # QueryDlgParameterID
	ProcessParameter(40,34)

	'Restore normal dialog interaction
	Preference("InteractiveMode") = True

End Sub
