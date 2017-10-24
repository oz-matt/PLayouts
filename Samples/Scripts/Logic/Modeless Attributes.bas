Dim NameArray$()
Dim parts As objects
Dim attrsEqual As Boolean
Dim prevItem As String

Sub Main
	Begin Dialog UserDialog 0,0,460,91,"Attributes",.dialogfunc ' %GRID:10,7,1,1
		OKButton 340,63,110,21
		Text 10,7,50,14,"&Name:",.Text1
		DropListBox 70,7,250,84,NameArray(),.DropListBox
		Text 10,35,40,14,"&Value:",.Text2
		TextBox 70,35,250,21,.TextBox
		PushButton 340,7,50,21,"&Add",.AddBtn
		PushButton 400,7,50,21,"&Del",.DelBtn
		PushButton 340,35,110,21,"&Set",.SetBtn
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

Sub RefreshList(Optional CurItemName As String = "")
	prevEnable = DlgEnable (-1)
	DlgEnable -1, parts.Count < 10
	ReDim NameArray(0)
	CurItem = 0
	DlgText -1, "Attributes"
	If parts.Count  = 1 Then
		Set attrs = parts(1).Attributes
		ReDim NameArray(attrs.Count)
		DlgText -1, "Attributes of part " & parts(1).Name
		For i=1 To attrs.Count
			NameArray(i-1) = Attrs(i).Name
			If NameArray(i-1) = CurItemName Then  CurItem = i-1
		Next
	ElseIf parts.Count  > 1 Then
		DlgText -1, "Attributes of selected parts"

		UpdateProgress parts.Count, "Scanning selected parts for attribute Info..."
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
   	UpdateValue
	DlgEnable -1, prevEnable
   	UpdateButtons
End Sub

Public Sub UpdateButtons
	nm = DlgText("DropListBox")
	count = parts.Count
	DlgEnable "AddBtn", count <> 0
	DlgEnable "DelBtn", count <> 0 And nm <> ""
	DlgEnable "SetBtn", count <> 0  And nm <> "" And attrsEqual
	DlgEnable "DropListBox", count <> 0  
	DlgEnable "TextBox", count <> 0 And nm <> ""
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
	If parts.Count > 1 Then
		caption = "Add Attribute to Selected Parts"
	Else 
		caption = "Add Attribute to Parts " & parts(1).Name
	End If
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


