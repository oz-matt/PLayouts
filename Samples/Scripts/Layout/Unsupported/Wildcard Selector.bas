' This Visual Basic script allows you to use full wildcarding to select/unselect parts, components
' pins, nets, and vias.  Zero or more wildcard characters (*) may appear anywhere in the search
' string.  You may add or remove names from the select list by combining multiple selections and
' unselections.  The names of the selected objects will appear alphabetically in a listbox as part
' of the dialog box.
'
' Parts are displayed as:			<part_type>:<comp_name>		Example: 12-34567-89:R12
' Components are displayed as:	<comp_name>					Example: R12
' Pins are displayed as:			<comp_name>.<pin_name>		Example: R12.1
' Nets are display as:				<net_name>						Example: GND
' Vias are displayed as:			<internal_via_name>				Example: Via3C000123
'
' Some useful examples might be:
'	In part mode, select all 12 class capacitors:					12*:c*  [press Select]
'	In comp mode, select all diodes and resistors:					d*  [press Select]  r*  [press Select]
'	In pin mode, select pin one on all of the capacitors:			c*.1  [press Select]
'	In net mode, select all non-negative asserted clock nets:		*clk*  [press Select]  -*  [press Unselect]
'	In via mode, select all vias									* [press Select]
'

Const t_part% = 0, t_comp% = 1, t_net% = 2, t_pin% = 3, t_via% = 4, t_all% = 9999

Private SelList$(50000)
Private SelCnt%, SelType%

Sub Main
	Begin Dialog UserDialog 320,231,"Wildcard",.CallbackFunc
		Text 10,7,100,14,"Search String:",.t_search
		TextBox 10,21,200,21,.tb_search
		PushButton 230,7,70,21,"Select",.pb_select
		PushButton 230,35,70,21,"Unselect",.pb_unselect
		PushButton 230,189,70,21,"Close",.pb_close
		Text 10,49,150,14,"Selected: 0",.t_selected
		ListBox 10,63,200,161,SelList(),.lb_selects
		GroupBox 220,70,90,98,"Types",.gb_types
		OptionGroup .g_types
			OptionButton 230,91,70,14,"Parts",.b_parts
			OptionButton 230,105,70,14,"Comps",.b_comps
			OptionButton 230,119,70,14,"Nets",.b_nets
			OptionButton 230,133,70,14,"Pins",.b_pins
			OptionButton 230,147,70,14,"Vias",.b_vias
	End Dialog
	Dim dlg As UserDialog

	ActiveDocument.SelectObjects(t_all, "", False)			' Unselect all objects
	SelCnt = 0												' No objects are selected
	dlg.g_types = t_comp									' Start in Comp mode
	SelType =  dlg.g_types									' Remember what mode we are in

	Dialog dlg
End Sub

Private Function CallbackFunc%(DlgItem$, Action%, SuppValue%)
	Select Case Action
	Case 2													' Value changed or button pressed
		If (DlgItem = "g_types") Then							' Changing mode clears select list
			ActiveDocument.SelectObjects(t_all, "", False)
			ClearList()
			UpdateMenu()
			SelType = SuppValue
		ElseIf (DlgItem = "pb_select") Then					' Add objects to select list
			SelectObjects(SelType)
			CallbackFunc = True	
		ElseIf (DlgItem = "pb_unselect") Then					' Remove objects from select list
			UnselectObjects(SelType)
			CallbackFunc = True	
		End If
	End Select
End Function

Private Function SelectObjects%(ObjType%)
	Dim NextObj As Object
	Dim SearchType%
	Dim ObjName$
	
	ClearList()

	SearchType = IIf(ObjType = t_part, t_comp, ObjType)
	
	For Each NextObj In ActiveDocument.GetObjects(SearchType, "", False)
		ObjName = GetObjName(ObjType, NextObj)
		If (NextObj.Selected) Then
			SelList(SelCnt) = ObjName
			SelCnt = SelCnt + 1
		ElseIf (StringMatch(DlgText("tb_search"), ObjName)) Then
			NextObj.Selected = True
			SelList(SelCnt) = ObjName
			SelCnt = SelCnt + 1
		End If
	Next NextObj

	UpdateMenu()

	SelectObjects = True
End Function

Private Function UnselectObjects%(ObjType%)
	Dim NextObj As Object
	Dim SearchType%
	Dim ObjName$
	
	ClearList()

	SearchType = IIf(ObjType = t_part, t_comp, ObjType)

	For Each NextObj In ActiveDocument.GetObjects(SearchType, "", True)
		ObjName = GetObjName(ObjType, NextObj)
		If (StringMatch(DlgText("tb_search"), ObjName)) Then
			NextObj.Selected = False
		Else
			SelList(SelCnt) = ObjName
			SelCnt = SelCnt + 1
		End If
	Next NextObj

	UpdateMenu()

	UnselectObjects = True
End Function

Private Function GetObjName$(ObjType%, NextObj As Object)
	If (ObjType <> t_part) Then
		GetObjName = NextObj.Name
	Else
		' For parts, I append the component name to the part type.  I also remove any value and
		' tolerance values from the part type, but that could be commented out as shown below.

		i = InStr(NextObj.PartType, ",")										' comment out to include values/tolerances
		If (i = 0) Then														' comment out to include values/tolerances
			GetObjName = NextObj.PartType & ":" & NextObj.Name
		Else																' comment out to include values/tolerances
			GetObjName = Left(NextObj.PartType, i - 1) & ":" & NextObj.Name	' comment out to include values/tolerances
		End If																' comment out to include values/tolerances
	End If
End Function

Private Function ClearList%()
	For i = 0 To (SelCnt - 1)
		SelList(i) = Empty
	Next i
	SelCnt = 0
	
	ClearList = True
End Function

Private Function UpdateMenu%()
	TempStr$ = ""

	For i1 = 0 To (SelCnt - 2)									' Sort Selected items
		For i2 = (i1 + 1) To (SelCnt - 1)
			If (StrComp(SelList(i2), SelList(i1), 1) < 0) Then
				TempStr = SelList(i1)
				SelList(i1) = SelList(i2)
				SelList(i2) = TempStr
			End If
		Next i2
	Next i1

	DlgListBoxArray "lb_selects", SelList()						' Update the selected listbox
	DlgText 5, "Selected: " & SelCnt							' Update selected count string
	DlgText 1, ""											' Clear the search string
	DlgFocus 1												' Set focus to the text box

	UpdateMenu = True
End Function

Private Function StringMatch%(OrigWildStr$, OrigTestStr$)

	StringMatch = True
	PrevStar% = False

	WildStr = UCase(OrigWildStr)
	WildLen% = Len(WildStr)
	TestStr = UCase(OrigTestStr)
	TestLen% = Len(TestStr)

	While (WildLen > 0)

		i1 = InStr(WildStr, "*")
		If (i1 = 0) Then WildSubStr = WildStr Else WildSubStr = Left(WildStr, i1 - 1)
		WildSubLen = Len(WildSubStr)
		
		i2 = InStr(TestStr, WildSubStr)

		If ((i2 = 0) Or ((PrevStar = False) And (i2 <> 1))) Then
			StringMatch = False
			WildLen = 0
		Else
			TestStr = Right(TestStr, TestLen - (i2 + WildSubLen - 1))
			TestLen = Len(TestStr)
			If (i1 = 0) Then
				WildLen = 0
				PrevStar = False
			Else
				WildStr = Right(WildStr, WildLen - (WildSubLen + 1))
				WildLen = Len(WildStr)
				PrevStar = True
			End If
		End If
			
	Wend

	If ((StringMatch = True) And (PrevStar = False) And (TestLen > 0)) Then
		StringMatch = False
	End If

End Function

