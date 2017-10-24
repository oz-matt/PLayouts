Dim parts As objects
Dim gates As objects
Dim curSht As String
Dim changing As Boolean

Sub Main
	Begin Dialog UserDialog 430,259,"Part/Gate Properties",.dialogfunc ' %GRID:10,7,1,1
		Text 70,14,130,14,"",.RefDes
		OKButton 330,231,90,21
		GroupBox 10,126,220,98,"Position",.GroupBox1
		PushButton 40,196,90,21,"&Move",.MoveTo
		Text 10,14,60,14,"Ref-Des:",.Text1
		Text 110,56,100,14,"",.PartTypeNm
		PushButton 320,7,100,21,"Rename &Part",.partRename
		PushButton 210,7,100,21,"Rename &Gate",.gateRename
		Text 20,147,20,14,"&X:",.Text5
		TextBox 40,140,90,21,.posX
		Text 20,175,20,14,"&Y:",.Text6
		TextBox 40,168,90,21,.posY
		GroupBox 240,126,180,98,"Orientation",.GroupBox2
		CheckBox 260,147,90,14,"X &Mirror",.chkXMir
		CheckBox 260,168,90,14,"Y M&irror",.chkYMir
		CheckBox 260,189,90,14,"Rotated&90",.ckhRot
		GroupBox 10,35,410,84,"Part Information",.GroupBox3
		Text 20,98,90,14,"Logic Family:",.Text3
		Text 20,77,90,14,"PCB Decal:",.Text4
		Text 220,56,80,14,"Gate Count:",.Text7
		Text 330,56,60,14,"Unused:",.Text8
		Text 110,77,100,14,"",.PCBDecalNm
		Text 110,98,100,14,"",.LogicFam
		Text 20,56,80,14,"Part Type:",.Text2
		Text 300,56,30,14,"",.GateCount
		Text 390,56,20,14,"",.UnusedGateCnt
		Text 220,77,90,14,"Pin Count:",.Text9
		Text 300,77,30,14,"",.PinCount
		PushButton 140,168,40,21,"&Lf",.LeftBtn
		PushButton 180,168,40,21,"&Rt",.RightBtn
		PushButton 160,189,40,21,"&Dn",.DownBtn
		PushButton 160,147,40,21,"&Up",.UpBtn
		Text 330,77,60,14,"Unused:",.Text10
		Text 390,77,20,14,"",.UnusedPinCnt
		Text 220,98,120,14,"Signal Pin Count:",.Text11
		Text 340,98,60,14,"",.SigPinCnt
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
	
End Sub

Rem See DialogFunc help topic for more information.
Private Function dialogfunc(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		View_Change
		Document_SelectionChange
	Case 2 ' Value changing or button pressed
		changing = True
		dialogfunc = True
		If DlgItem$ = "partRename" Then 
			RenamePart
		ElseIf DlgItem$ = "gateRename" Then 
			RenameGate
		ElseIf DlgItem$ = "chkXMir" Then 
			gates(1).ReflectedX = DlgValue("chkXMir") = 1
		ElseIf DlgItem$ = "chkYMir" Then 
			gates(1).ReflectedY = DlgValue("chkYMir") = 1
		ElseIf DlgItem$ = "ckhRot" Then 
			gates(1).Rotated90 = DlgValue("ckhRot") = 1
		ElseIf DlgItem$ = "MoveTo" Then
			MoveTo
		ElseIf DlgItem$ = "UpBtn" Then
			MoveBy 0, ActiveDocument.GridY(plogGridDesign)
		ElseIf DlgItem$ = "DownBtn" Then
			MoveBy 0, -ActiveDocument.GridY(plogGridDesign)
		ElseIf DlgItem$ = "LeftBtn" Then
			MoveBy -ActiveDocument.GridX(plogGridDesign), 0
		ElseIf DlgItem$ = "RightBtn" Then
			MoveBy ActiveDocument.GridX(plogGridDesign), 0
		ElseIf DlgItem$ = "OK" Then
			dialogfunc = False
		End If
		changing = False
		Document_SelectionChange
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem dialogfunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Sub MoveBy(dx As Double, dy As Double)
	For Each gt In gates
		gt.Move gt.PositionX + dx, gt.PositionY + dy
	Next
End Sub

Sub MoveTo
	On Error GoTo CheckError    ' Enable error trapping.
	posX = CDbl(DlgText("posX"))
	posY = CDbl(DlgText("posY"))
	gates(1).Move(posX, posY)
	On Error GoTo 0 ' Disable error trapping.  
	Exit Sub
	
CheckError:
    MsgBox Err.Description, vbExclamation, "Error"
    UpdateStrings
    On Error GoTo 0 ' Disable error trapping.    
    Exit Sub		
End Sub

Sub RenamePart
	newName = InputBox("Enter the new name:", "Rename Part " & parts(1).Name, parts(1).Name)
	If newName <> ""  Then 
		On Error GoTo CheckError    ' Enable error trapping.
		parts(1).Name = newName
		On Error GoTo 0 ' Disable error trapping.    
	End If
	Exit Sub
	
CheckError:
    MsgBox Err.Description, vbExclamation, "Error"
    On Error GoTo 0 ' Disable error trapping.    
    Exit Sub	
End Sub

Sub RenameGate
	newName = InputBox("Enter the new name:", "Rename Gate " & gates(1).Name, gates(1).Name)
	If newName <> ""  Then 
		On Error GoTo CheckError    ' Enable error trapping.
		gates(1).Name = newName
		On Error GoTo 0 ' Disable error trapping.
	End If
	Exit Sub
	
CheckError:
	MsgBox Err.Description, vbExclamation, "Error"
    changing = False
    On Error GoTo 0 ' Disable error trapping.    
    Exit Sub		
End Sub

Sub UpdateStrings
	If parts.Count <> 1 Then 
		DlgText "RefDes", ""
	ElseIf gates.Count > 1 Then
		DlgText "RefDes", parts(1).Name
	Else 
		DlgText "RefDes", gates(1).Name
	End If

	If gates.Count <> 1 Then
		DlgText "posX", ""
		DlgText "posY", ""
		DlgValue "chkXMir",  0
		DlgValue "chkYMir", 0
		DlgValue "ckhRot", 0
	Else
		DlgText "posX", gates(1).PositionX
		DlgText "posY", gates(1).PositionY
		DlgValue "chkXMir",  gates(1).ReflectedX 
		DlgValue "chkYMir", gates(1).ReflectedY
		DlgValue "ckhRot", gates(1).Rotated90
	End If

	If parts.Count = 1 Then
		DlgText "PartTypeNm", parts(1).PartType
		DlgText "PCBDecalNm", parts(1).PCBDecal
		DlgText "LogicFam", parts(1).PartTypeLogic
		DlgText "GateCount", parts(1).Gates.Count + parts(1).UnusedGates.Count
		DlgText "UnusedGateCnt", parts(1).UnusedGates.Count
		DlgText "PinCount", parts(1).Pins.Count
		unusedPins = 0
		sigPins = 0
		For Each aPin In parts(1).Pins
			If aPin.Gate Is Nothing And Not aPin.Net Is Nothing Then sigPins = sigPins + 1
			If aPin.Net Is Nothing Then unusedPins = unusedPins + 1
		Next
		DlgText "UnusedPinCnt", unusedPins
		DlgText "SigPinCnt", sigPins
	Else		
		DlgText "PartTypeNm", ""
		DlgText "PCBDecalNm", ""
		DlgText "LogicFam", ""
		DlgText "GateCount", ""
		DlgText "UnusedGateCnt", ""
		DlgText "PinCount", ""
		DlgText "UnusedPinCnt", ""
		DlgText "SigPinCnt", ""
	End If
End Sub

Public Sub UpdateButtons
	DlgEnable "partRename", parts.Count = 1
	DlgEnable "gateRename", gates.Count = 1
	DlgEnable "chkXMir", gates.Count = 1
	DlgEnable "chkYMir", gates.Count = 1
	DlgEnable "ckhRot", gates.Count = 1
	DlgEnable "posX", gates.Count = 1
	DlgEnable "posY", gates.Count = 1
	DlgEnable "MoveTo", gates.Count = 1
	DlgEnable "UpBtn", gates.Count > 0
	DlgEnable "DownBtn", gates.Count > 0
	DlgEnable "LeftBtn", gates.Count > 0
	DlgEnable "RightBtn", gates.Count > 0
	
	UpdateStrings
End Sub

Public Sub Document_SelectionChange()
	If Not changing Then 
		Set parts = ActiveDocument.ActiveSheet.GetObjects(plogObjectTypeComponent,"",True)
		Set gates = ActiveDocument.ActiveSheet.GetObjects(plogObjectTypeGate,"",True)
		UpdateButtons
	End If
End Sub

Public Sub View_Change()
	If ActiveDocument.ActiveSheet <> curSht Then
		curSht = ActiveDocument.ActiveSheet
		Document_SelectionChange
	End If
End Sub
