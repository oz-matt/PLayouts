'''''''''#Uses "E:\Elec\Pads\ole\vbscripts\Module\Standard.BAS"
Option Explicit
Const RegString$="PadsScript"
Const RegAppString$="PowerLogic Compare PCB"

Dim CheckFile$

Global schErrs() As String
Global pcbErrs() As String
Dim schErrNr%, pcbErrNr%

Global EqualName() As String
Global EqualPins() As String
Global sDivPins() As String
Global pDivPins() As String
Dim eqNr%
Dim Report$, sp$
Dim ECOName$
Dim RepName$
'********************************************************************************
'********************************************************************************
'********************************************************************************
Sub main
Dim i%,j%,t$
sp = Chr(13) & Chr(10)

ECOName = DefaultFilePath & "\eco.eco"
RepName = DefaultFilePath & "\Netrep.txt"

CheckFile = GetSetting (RegString,RegAppString,"CheckFile")
If Not exists(DefaultFilePath & "\NetnameCompare.mcr") Then Call InitMac
If Not exists(CheckFile) Then
	If MsgBox("Cannot find PowerLogic Checkfile, search ?",17,"Error.") = vbOK Then
		t = GetFilePath("Checkasc","lst",,"Look For Checkasc.lst",0)
		If exists(t) Then
		CheckFile = t
		SaveSetting RegString,RegAppString,"CheckFile",CheckFile
		Else
			Beep
			End
		End If
	Else
		End
	End If 
End If 

	Begin Dialog UserDialog 480,42,"PowerLogic Net Rename",.BackFunc
		GroupBox 10,0,280,35,"Status",.GroupBox2
		Text 20,14,260,14,"Text1",.Meldung
		PushButton 300,7,80,28,"Cancel",.PushButton1
		PushButton 390,7,80,28,"Accept",.PushButton2
	End Dialog
Dim dlg As UserDialog
Dialog dlg


End Sub
'********************************************************************************
'********************************************************************************
'********************************************************************************
Private Function BackFunc(DlgItem$, Action%, SuppValue%) As Boolean
Debug.Print "Action " & action & "   DLGItem  " & dlgitem & "   SuppValue " & suppvalue

Dim i%,j%,x&
	Select Case Action%
	Case 1 
	Case 2
		If dlgitem = "PushButton2" Then
			RunMacro(DefaultFilePath & "\NetnameCompare.mcr","ECOin")
		End If
	Case 3
	Case 4
	Case 5
		DlgEnable "PushButton1",False
		DlgEnable "PushButton2",False
		Call AnalyseReport
		DlgText "Meldung","Read Report"
		Call GetEqualNets(schErrs(), schErrNr, pcbErrs(), pcbErrNr, 0)
		Call GetEqualNets(pcbErrs(), pcbErrNr, schErrs(), schErrNr, 1)
		eqNr = eqNr - 1
		Call MakeDebugRep
		DlgText "Meldung","Make Report"
		Call MakeReport(i,j)
		If i = True Then DlgEnable "Pushbutton2",True
		DlgEnable "Pushbutton1", True
		If i = True Or j = True Then x = Shell("Notepad.exe " & Repname,1)
		DlgText "Meldung","Ready"
		If i = False And j = False Then  DlgText "Meldung","Nothing to do."
		Beep
	End Select
End Function

'********************************************************************************
'********************************************************************************
'********************************************************************************
Sub AnalyseReport
Dim i%,t$,j%,u$
Dim inline$
Dim schMPins%
Dim pcbMPins%
Open CheckFile For Input As #1
While Not EOF(1)
Line Input #1, inline
	If Left(inline,3) = "Net" Then
		If getel(inline," ",5) = "input" Then
			Call GetPins(schErrs(), schErrNr, inline)
		ElseIf getel(inline," ",5) = "data" Then
			Call GetPins(pcbErrs(), pcbErrNr, inline)
		End If
	End If
Wend	
Close #1
schErrNr = schErrNr - 1
pcbErrNr = pcbErrNr - 1
End Sub
'********************************************************************************
'********************************************************************************
'********************************************************************************
'********************************************************************************
'********************************************************************************
'********************************************************************************
Sub CheckNetToNet(s%,p%)
Dim i%,j%,k%,t$,x%,y%
Dim sch$,pcb$
Dim CheckFound%, PinFound%, ArrayNr%

sch = schErrs(s)
pcb = pcbErrs(p)
x = countel(sch," ")
y = countel(pcb," ")

t = getel(schErrs(s)," ",1) & " " & getel(pcbErrs(p)," ",1)
CheckFound = False
For i = 0 To eqNr-1
	If EqualName(i) = t Then 
		CheckFound = True
		ArrayNr = i
		Exit For
	End If	
Next i

If Checkfound = False Then
	ReDim Preserve EqualName(eqNr) As String
	ReDim Preserve EqualPins(eqNr) As String
	ReDim Preserve sDivPins(eqNr) As String
	ReDim Preserve pDivPins(eqNr) As String
	EqualName(eqNr) = t
	ArrayNr = eqNr
	eqNr =eqNr + 1	
End If

For i = 2 To x
	For j = 2 To y
		t = getel(sch," ",i)
		If t = getel(pcb," ",j) And Left(t,1) <> "#" Then
			PinFound = False
			For k = 1 To countel(EqualPins(ArrayNr)," ")
				If t = getel(EqualPins(ArrayNr)," ",k) Then
					PinFound = True
					Mid(sch,InStr(1,sch,t),Len(t)) = String(Len(t),"#")
					Mid(pcb,InStr(1,pcb,t),Len(t)) = String(Len(t),"#")
					Exit For
				End If
			Next k
			If PinFound = False Then
				EqualPins(ArrayNr) = EqualPins(ArrayNr) & " " & t
				Mid(sch,InStr(1,sch,t),Len(t)) = String(Len(t),"#")
				Mid(pcb,InStr(1,pcb,t),Len(t)) = String(Len(t),"#")
			End If
		End If
	Next j
Next i

Call CheckErrPins(sch,sDivPins(),ArrayNr) 
Call CheckErrPins(pcb,pDivPins(),ArrayNr) 

End Sub
'********************************************************************************
'********************************************************************************
'********************************************************************************
Sub CheckErrPins(t$,arr() As String,Nr%) 
Dim i%,u$,j%,found%
For i = 2 To countel(t," ")
	u = getel(t," ",i)
	If Left(u,1) <> "#" Then
		found = False
		For j = 1 To countel(arr(nr)," ")
			If getel(arr(nr)," ",j) = u Then 
				found = True
				Exit For
			End If	
		Next j
		If found = False Then arr(Nr) = arr(Nr) & " " & u	
	End If
Next i
End Sub
'********************************************************************************
'********************************************************************************
'********************************************************************************
Sub GetPins(Errs()As String, ErrNr%, il$)
Dim inline$
inline = " "
ReDim Preserve errs(errnr) As String
errs(errnr) = getel(il," ",2)
While inline <> ""
	If Not EOF(1) Then Line Input #1,inline
	If inline <> "" Then errs(errnr) = errs(errnr) & " " & Trim(inline)
Wend
errnr = errnr + 1
End Sub
'********************************************************************************
'********************************************************************************
'********************************************************************************
Sub GetEqualNets(sa()As String, saNr%, pa() As String, paNr%, mode%)
Dim i%,j%,k%,l%,t$
For i = 0 To saNr
	DlgText "Meldung", "Check " & getel(sa(i)," ",1)
	For j = 0 To paNr
		For k = 2 To countel(sa(i)," ")
			For l = 2 To countel(pa(j)," ")
				If getel(sa(i)," ",k) = getel(pa(j)," ",l) Then 
					If mode = 0 Then Call CheckNetToNet(i,j) Else Call CheckNetToNet(j,i)
					Exit For
				End If
			Next l
		Next k
	Next j
Next i
End Sub
'********************************************************************************
'********************************************************************************
'********************************************************************************
Sub MakeReport(ren%,show%)
Dim i%,j%,k%,x&,Nname$, found%,l%
Dim t$

Open repname For Output As #1
Open econame For Output As #2
Print #2,"*PADS-ECO*" 
Print #2,"*RENNET*"

Print #1,""
Print #1,"Identical Nets with different Names : "
Print #1,"======================================"
Print #1,"Schematic                        PCB"
Print #1,"--------------------------------------"


For i = 0 To eqNr
	If sDivPins(i) = "" And pDivPins(i) = "" Then
		ren = True
		Print #1,getel(EqualName(i)," ",1) & String(1,Chr(9)) & " <-> "& String(2,Chr(9)) & getel(EqualName(i)," ",2)
		Print #2,getel(EqualName(i)," ",2) & "  " & getel(EqualName(i)," ",1)
		EqualName(i) = "#" & EqualName(i) 
	End If
Next i

Print #1,""
Print #1,"Different Pincount but no Shortcuts : "
Print #1,"======================================"
Print #1,"Schematic                        PCB"
Print #1,"--------------------------------------"


For i = 0 To eqNr
	If Left(EqualName(i),1) <> "#" Then
		found = False
		For j = 1 To countel(sDivPins(i)," ")
			If ChkNPin(pDivPins(), getel(sDivPins(i)," ",j)) Or ChkNPin(EqualPins(), getel(sDivPins(i)," ",j)) = True Then 
				found = True
				Exit For
			End If
		Next j
		If found = False Then
			For j = 1 To countel(pDivPins(i)," ")
				If ChkNPin(sDivPins(), getel(pDivPins(i)," ",j)) Or ChkNPin(EqualPins(), getel(pDivPins(i)," ",j)) = True Then 
					found = True
					Exit For
				End If
			Next j
		End If
		If found = False Then
			ren = True
			Print #1,getel(EqualName(i)," ",1) & String(1,Chr(9)) & " <-> "& String(2,Chr(9)) & getel(EqualName(i)," ",2)
			Print #2,getel(EqualName(i)," ",2) & "  " & getel(EqualName(i)," ",1)
			EqualName(i) = "#" & EqualName(i) 
		End If
	End If	
Next i

Print #1,""
Print #1,"ShortCuts : "
Print #1,"======================================"
Print #1,"Schematic                        PCB"
Print #1,"--------------------------------------"

For i = 0 To eqNr
	If Left(EqualName(i),1) <> "#" Then
		show = True
		Print #1,getel(EqualName(i)," ",1) & String(1,Chr(9)) & " <-> "& String(2,Chr(9)) & getel(EqualName(i)," ",2)
	End If
Next i

Print #2,"*END*"
Close #1,#2


End Sub
'********************************************************************************
'********************************************************************************
'********************************************************************************
Function ChkNPin(arr()As String, t$) As Integer
Dim i%,j%,k%,l%

For i = 0 To eqNr
	For j = 1 To countel(arr(i)," ")
		If t = getel(arr(i)," ",j) Then
			ChkNPin = True
			Exit Function
		End If
	Next j	
Next i

End Function
'********************************************************************************
'********************************************************************************
'********************************************************************************
Sub MakeDebugRep
Dim i%
Open "f:\t.txt" For Output As #1
For i = 0 To eqnr
Print #1,"Name  " & equalname(i)
Print #1,"Pins  " & equalPins(i)
Print #1,"SDiv  " & sdivpins(i)
Print #1,"PDiv  " & pdivpins(i)
Print #1,""
Next i
Close 
End Sub
'********************************************************************************
'********************************************************************************
'********************************************************************************
Sub InitMac
On Error Resume Next
Open DefaultFilePath & "\NetnameCompare.mcr" For Output As #1

Print #1,"MACRO ECOin"
Print #1,"VERSION 2.0"
Print #1,"KEY"
Print #1,"DATE"
Print #1,"DESCRIPTION "
Print #1,"BEGIN"
Print #1,"   COMMAND 33 # CLI_FILE_IMPORT"
Print #1,"   PARAMETER INTEGER 6054 1 # AssOptSetLevel"
Print #1,"   PARAMETER INTEGER 3701 0 # UndoRedoAvailableParameterID"
Print #1,"   PROMPT YES_NO_CANCEL " & Chr(34) & "Save old file before loading?" & Chr(34) &  "   NO"
Print #1,"   PARAMETER STRING 1205 " & Chr(34) & DefaultFilePath & "\eco.eco" & Chr(34) & " # InputAsciiFileParameterID"
Print #1,"   COMMAND 82 # CLI_VIEW_REDRAW"
Print #1,"END"
	
Close #1
On Error GoTo 0
End Sub
'********************************************************************************
'********************************************************************************
'********************************************************************************
'							Common Functions
'********************************************************************************
'********************************************************************************
'********************************************************************************
Function Exists(f$) As Integer
	Dim X&
	On Error Resume Next
	X = FileLen(f$)
	If X Then Exists = True
End Function
'********************************************************************************
'********************************************************************************
'********************************************************************************
Function GetEl(tt$, ch$, el%) As String
Dim i%, a%, l%, b%, u$, t$
t = tt
While Left(t, 1) = ch
	 t = Right(t, Len(t) - 1)
 Wend          
b = 1
a = InStr(1, t, ch)
While a > 0
   If b = el Then GetEl = Left(t, a - 1): Exit Function
   t = Right(t, Len(t) - a)
   While Left(t, 1) = ch
   	t = Right(t, Len(t) - 1)
   Wend
   b = b + 1
   a = InStr(1, t, ch)
Wend
GetEl = t
End Function

'********************************************************************************
'********************************************************************************
'********************************************************************************
Function CountEl(tt$,u$)
Dim t$,a%
t = tt
While Left(t,1) = u
	t = Right(t, Len(t)-1)
Wend				
While Right(t,1) = u
	t = Left(t, Len(t)-1)
Wend
If t = "" Then Exit Function
While t <>""
	a = InStr(1,t,u)
	If a > 0 Then
		t = Right(t,Len(t)-a)
		While Left(t,1) = u
			t = Right(t, Len(t)-1)
		Wend				
	CountEl = CountEl + 1
	Else
		CountEl = CountEl + 1
		Exit Function
	End If
Wend
End Function

'********************************************************************************
'********************************************************************************
'********************************************************************************

'********************************************************************************
'********************************************************************************
'********************************************************************************


























