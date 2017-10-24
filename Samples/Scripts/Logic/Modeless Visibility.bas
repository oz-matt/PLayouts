Dim NameArray$()
Dim parts As objects
Dim gates As objects
Dim attrsEqual As Boolean
Dim prevItem As String
Dim visStates(8) As Integer

Sub Main
	Begin Dialog UserDialog 0,0,470,210,"Visibility",.dialogfunc ' %GRID:10,7,1,1
		OKButton 350,182,110,21
		GroupBox 10,7,450,77,"Attributes for part: ",.PartGroupBox
		Text 20,28,50,14,"&Name:",.Text1
		DropListBox 70,28,260,175,NameArray(),.DropListBox
		Text 20,56,40,14,"&Value:",.Text2
		TextBox 70,56,260,21,.TextBox
		PushButton 340,28,50,21,"&Add",.AddBtn
		PushButton 400,28,50,21,"&Del",.DelBtn
		PushButton 340,56,110,21,"&Set",.SetBtn
		GroupBox 10,91,450,84,"Visibility  for gate: ",.GateGroupBox
		PushButton 360,112,90,21,"App&ly",.ApplyBtn
		CheckBox 20,112,140,14,"Current Attribute",.CheckBox1,2
		CheckBox 20,126,200,14,"Current Attribute Name",.CheckBox2,2
		CheckBox 20,140,180,14,"Reference Designator",.CheckBox3,2
		CheckBox 20,154,90,14,"Part Type",.CheckBox4,2
		CheckBox 220,112,110,14,"Pin Numbers",.CheckBox5,2
		CheckBox 220,126,90,14,"Pin Name",.CheckBox6,2
		CheckBox 220,140,90,14,"PCBDecal",.CheckBox7,2
		CheckBox 220,154,130,14,"PCBDecalName",.CheckBox8,2
	End Dialog
	
	Dim dlg As UserDialog
	Dialog dlg
	
End Sub

Function dialogfunc(DlgItem$, Action%, SuppValue%)  As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		Document_SelectionChange	
	Case 2 ' Value changing or button pressed
		If DlgItem$ = "DropListBox" Then 
			UpdateValue
			UpdateVisibilityButtons 1
			UpdateButtons
		ElseIf DlgItem$ = "AddBtn" Then 
			AddNewAttribute
    	   	dialogfunc = True
		ElseIf DlgItem$ = "DelBtn" Then 
			DelAttribute
    	   	dialogfunc = True
		ElseIf DlgItem$ = "SetBtn" Then 
			SetNewValue
    	   	dialogfunc = True
		ElseIf DlgItem$ = "ApplyBtn" Then 
			SetVisibility
    	   	dialogfunc = True
		End If
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
		If DlgItem$ = "TextBox" Then
			attrsEqual = True 'to enable Set button
			UpdateButtons
		End If
	Case 5 ' Idle
	End Select
End Function

' Update editBox
Sub UpdateValue
	attrsEqual = False
	DlgText "TextBox", ""
	If parts.Count = 0 Then Exit Sub
	prevItem = DlgText("DropListBox")
	Set attr1 = parts(1).Attributes(prevItem)
	If attr1 Is Nothing Then Exit Sub
	value = attr1.Value
	For i = 2 To parts.Count
		Set attr = parts(i).Attributes(prevItem)
		If attr Is Nothing Then Exit Sub
		If value <> attr.Value Then  Exit Sub
	Next i			
	attrsEqual = True
	DlgText "TextBox", value
End Sub

Sub	SetVisibility
	nm = DlgText("DropListBox")
	prevEnable = DlgEnable (-1)
	DlgEnable -1, parts.Count < 10
	UpdateProgress gates.Count, "Setting visibility for selected gates..."
	For Each gt In gates
		firstItem = IIf(gt.Component.Attributes(nm) Is Nothing, 2, 0)
		For i = firstItem To 7
			item = DlgNumber("CheckBox1")+i
			If DlgValue(item) <> 2 Then
					gt.Visibility(i, nm) = DlgValue(item) = 1
			End If
		Next i
		UpdateProgress 1
	Next gt
	UpdateProgress 0
	UpdateVisibilityButtons 7
	DlgEnable -1, prevEnable
	ActiveDocument.ActiveView.Refresh
End Sub

Sub GetGateVisibility(gt As Object, nm As String, lastItem As Integer)
	firstItem = 0
	If gt.Component.Attributes(nm) Is Nothing Then
		firstItem = 2
		If visStates(0) = -1 Then 
			visStates(0) = 0
		ElseIf visStates(0) = 1 Then 
			visStates(0) = 2
		End If
		If visStates(1) = -1 Then 
			visStates(1) = 0
		ElseIf visStates(1) = 1 Then
			visStates(1) = 2
		End If
	End If
	For i = firstItem To lastItem
		state = IIf(gt.Visibility(visItem + i, nm), 1, 0)
		If visStates(i) = -1 Then
			visStates(i) = state
		ElseIf visStates(i) <> 2 And visStates(i) <> state Then
			visStates(i) = 2
		End If
	Next
End Sub

Sub UpdateVisibilityButtons(lastItem As Integer)
	nm = DlgText("DropListBox")
	For i = 0 To lastItem
		 visStates(i) = -1
	Next
	prevEnable = DlgEnable (-1)
	DlgEnable -1, parts.Count < 10
	UpdateProgress gates.Count, "Scanning selected gates for visibility Info..."
	For Each gt In gates
		GetGateVisibility gt, nm, lastItem
		UpdateProgress 1
	Next
	UpdateProgress 0
	DlgEnable -1, prevEnable
	For i = 0 To lastItem
		item = DlgNumber("CheckBox1")+i
		DlgEnable item, gates.Count <> 0 And visStates(i) <> -1
		If visStates(i) <> -1 Then DlgValue item, visStates(i)
	Next
	If gates.Count = 0 Then
		gtName = ""
	ElseIf gates.Count = 1 Then
		gtName = gates(1).Name
	Else
		gtName = "Multiple"
	End If
	DlgText "GateGroupBox", "Visibility of gate: " & gtName & "          "
End Sub

Sub RefreshList(Optional CurItemName As String = "")
	prevEnable = DlgEnable (-1)
	DlgEnable -1, parts.Count < 10
	ReDim NameArray(0)
	CurItem = 0
	PartName = ""
	If parts.Count  = 1 Then
		Set attrs = parts(1).Attributes
		ReDim NameArray(attrs.Count)
		PartName = parts(1).Name
		For i=1 To attrs.Count
			NameArray(i-1) = Attrs(i).Name
			If NameArray(i-1) = CurItemName Then  CurItem = i-1
		Next
	ElseIf parts.Count  > 1 Then
		PartName = "Multiple"

		UpdateProgress parts.Count, "Scanning selected parts for attribute info..."
		' calculate total count of attributes
		totalCount = 0
		For Each part In parts
			totalCount = totalCount + part.Attributes.Count
		Next

		' get all attributes and add them temporarily to the first object
		Dim tmpArray$()
		ReDim tmpArray(totalCount)
		realCount = 0
		Set attrs1 = parts(1).Attributes
		For i = 2 To parts.Count
			Set attrs = parts(i).Attributes
			For j = 1 To attrs.Count
				If attrs1(attrs(j).Name) Is Nothing Then
					tmpArray(realCount) = attrs(j).Name
					attrs1.Add(tmpArray(realCount), "")
					realCount = realCount + 1
				End If
			Next j
			UpdateProgress 1
		Next i
		UpdateProgress 0
		
		' fill associated array
		If attrs1.Count Then ReDim NameArray(attrs1.Count)
		For i = 1 To attrs1.Count
			NameArray(i-1) = attrs1(i).Name
			If NameArray(i-1) = CurItemName Then  CurItem = i-1
		Next i
		
		' remove added temp attributes
		For i = 1 To realCount
			attrs1.Delete(tmpArray(i-1))
		Next i		
	End If
    DlgListBoxArray "DropListBox",NameArray$()
    DlgValue "DropListBox", CurItem
    DlgText "PartGroupBox", "Attributes of part: " & PartName & "               "
   	UpdateValue
	DlgEnable -1, prevEnable
   	UpdateVisibilityButtons 7
   	UpdateButtons
End Sub

Sub UpdateButtons
	nm = DlgText("DropListBox")
	count = parts.Count
	DlgEnable "AddBtn", count <> 0
	DlgEnable "DelBtn", count <> 0 And nm <> ""
	DlgEnable "SetBtn", count <> 0  And nm <> "" And attrsEqual
	DlgEnable "DropListBox", count <> 0  
	DlgEnable "TextBox", count <> 0 And nm <> ""
	DlgEnable "ApplyBtn", count <> 0
End Sub

' Set new value for current attribute in each part 
Sub SetNewValue
	prevEnable = DlgEnable (-1)
	DlgEnable -1, parts.Count < 10
	nm = DlgText("DropListBox")
	newVal = DlgText("TextBox")
	Dim attrs As Attributes
	UpdateProgress parts.Count, "Setting attribute value for selected parts..."
	For Each part In parts
		Set attrs = part.Attributes
		If attrs(nm) Is Nothing Then 
			attrs.Add(nm, newVal)
		Else
			attrs(nm) = newVal
		End If
		UpdateProgress 1
	Next
	UpdateProgress 0
	DlgEnable -1, prevEnable
End Sub

' Add new attribute in each part
Sub AddNewAttribute
	ItemName = DlgText("DropListBox")
	caption = "Add Attribute for Selected Parts"
	NewName = InputBox("Enter the name:", caption)
	If NewName <> ""  Then
		prevEnable = DlgEnable (-1)
		DlgEnable -1, parts.Count < 10
		UpdateProgress parts.Count, "Adding attribute to selected parts..."
		For Each part In parts
			Set attrs = part.Attributes
			If attrs(NewName) Is Nothing Then attrs.Add(NewName,"")
			UpdateProgress 1
		Next
		UpdateProgress 0
		RefreshList NewName
		DlgFocus "TextBox"
		DlgEnable -1, prevEnable
	End If
End Sub

' Delete attribute from each part
Sub DelAttribute
	prevEnable = DlgEnable (-1)
	DlgEnable -1, parts.Count < 10
	ItemName = DlgText("DropListBox")
	UpdateProgress parts.Count, "Deleting attribute from selected parts..."
	For Each part In parts
		Set attrs = part.Attributes
		If Not attrs(ItemName) Is Nothing Then
			attrs.Delete(ItemName)
		End If
		UpdateProgress 1
	Next
	UpdateProgress 0
	RefreshList
	DlgEnable -1, prevEnable
End Sub

Public Sub Document_SelectionChange()
	Set parts = ActiveDocument.GetObjects(plogObjectTypeComponent,"",True)
	Set gates = ActiveDocument.GetObjects(plogObjectTypeGate,"",True)
	RefreshList prevItem
End Sub

Sub UpdateProgress(Optional count As Long, Optional msg As String = "")
	Static  sCurrent, sPercent, sTotal As Long
	If count > 0 And msg <> "" Then
		StatusBarText = msg
		sCurrent = 0
		sPercent = 0
		sTotal = count
	ElseIf count > 0 Then
		sCurrent = 	sCurrent + count
	Else
		ProgressBar = -1
		StatusBarText = ""
		Exit Sub
	End If
	NewProg = sCurrent * 100 / sTotal
	If sPercent <> NewProg Then
		sPercent = NewProg
		ProgressBar = sPercent
	End If
End Sub


