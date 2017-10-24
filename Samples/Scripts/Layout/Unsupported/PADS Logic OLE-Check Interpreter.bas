''''''''#Uses "E:\Elec\Pads\ole\vbscripts\Module\Standard.BAS"
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

'********************************************************************************
'********************************************************************************
'********************************************************************************
Sub main
Dim i%,t$

CheckFile = GetSetting (RegString,RegAppString,"CheckFile")

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

Call AnalyseReport

If pcbErrNr < 0 Then 
	MsgBox"     Checkfile is empty.",32,"PowerLogic OLE-Check Interpreter"
	End
End If

ReDim pcbnets(pcbErrNr) As String
ReDim pcbPins(0) As String
ReDim schnets(pcbErrNr) As String
ReDim schPins(0) As String

ReDim EqualName(0) As String

For i = 0 To pcbErrNr
	pcbnets(i) = getel(pcbErrs(i)," ",1) & Space(50) & Str(i)
Next i
For i = 0 To schErrNr
	schnets(i) = getel(schErrs(i)," ",1) & Space(50) & Str(i)
Next i

Call Sortitems(pcbErrNr, pcbnets())
Call Sortitems(schErrNr, schnets())

	Begin Dialog UserDialog 730,175,"PowerLogic  OLE ""Check Design"" Interpreter",.CheckCallBack
		GroupBox 160,119,450,49,"Selected",.GroupBox4
		GroupBox 10,0,350,119,"Netlist items in PCB",.GroupBox1
		ListBox 20,14,210,98,pcbnets(),.pcbNetListBox
		ListBox 240,14,110,91,pcbPins(),.pcbPinListBox
		GroupBox 10,119,130,49,"",.GroupBox2
		OptionGroup .Group1
			OptionButton 20,133,110,14,"Select Pins",.OptionButton1
			OptionButton 20,147,110,14,"Select Nets",.OptionButton2
		GroupBox 370,0,350,119,"Netlist items in Schematic",.GroupBox3
		ListBox 380,14,210,91,schnets(),.schNetListBox
		ListBox 600,14,110,91,schPins(),.schPinListBox
		Text 180,140,370,21,"",.SelText
		PushButton 630,147,90,21,"Close",.CloseButton
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub
'********************************************************************************
'********************************************************************************
'********************************************************************************
Private Function CheckCallBack(DlgItem$, Action%, SuppValue%) As Boolean
Dim i%,t$,x%,p%
Dim pin As Object

Debug.Print "Action : " & action & "   DLGItem : " & dlgitem & "   Suppvalue : " & suppvalue
	Select Case Action%
		Case 1
			DlgValue "schNetListBox", -1
		Case 2
			'***********************************
			If dlgitem = "Group1" Then
				If suppvalue = 1 Then
					DlgEnable "pcbNetListBox",1
					DlgEnable "pcbPinListBox",0
					DlgEnable "schNetListBox",0
					DlgEnable "schPinListBox",0
				Else
					DlgEnable "pcbNetListBox",1
					DlgEnable "pcbPinListBox",1
					DlgEnable "schNetListBox",1
					DlgEnable "schPinListBox",1
				End If
				DlgText "SelText", ""
				DlgValue "schNetListBox", -1
				DlgValue "schPinListBox", -1
				DlgValue "pcbNetListBox", -1
				DlgValue "pcbPinListBox", -1
			'***********************************
			ElseIf dlgitem = "pcbNetListBox" Then
				DlgValue "schNetListBox", -1
				DlgValue "schPinListBox", -1
				DlgValue "pcbPinListBox", -1
				Call processNetListBox(pcbnets(), pcbpins(), pcberrs(), suppvalue, 0)
			'***********************************
			ElseIf dlgitem = "pcbPinListBox" Then
				DlgValue "schNetListBox", -1
				DlgValue "schPinListBox", -1
				Call SelectPins(pcbpins(SuppValue%))
				DlgText "SelText", "Pin " & pcbpins(SuppValue%)			
			'***********************************
			ElseIf dlgitem = "schNetListBox" Then
				DlgValue "schPinListBox", -1
				DlgValue "pcbNetListBox", -1
				DlgValue "pcbPinListBox", -1
				Call processNetListBox(schnets(), schpins(), scherrs(), suppvalue, 1)
			'***********************************
			ElseIf dlgitem = "schPinListBox" Then
				DlgValue "pcbNetListBox", -1
				DlgValue "pcbPinListBox", -1
				Call SelectPins(schpins(SuppValue%))
				DlgText "SelText", "Pin " & schpins(SuppValue%)			
			'***********************************
			End If
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
Open checkfile For Input As #1
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
Sub ProcessNetListBox(nets() As String, pins() As String, errs() As String, index%, boxnr%)
Dim t$,i%,p%,x%
If DlgValue("Group1") = 1 Then
	ActiveDocument.SelectObjects(9999, "", False)
	t = getel(Nets(index%)," ",1)
	ActiveDocument.SelectObjects(2, t, True)
	DlgText "SelText", "Net : " & t			
End If
x = Val(getel(Nets(index)," ",2))
For i = 2  To countel(errs(x)," ")
		ReDim Preserve Pins(p)
		pins(p) = getel(errs(x)," ",i)
		p = p + 1
Next i
p=p-1
Call sortitems(UBound(pins),pins())
If boxnr = 0 Then DlgListBoxArray("pcbPinListBox", pins())
If boxnr = 1 Then DlgListBoxArray("schPinListBox", pins())
If DlgValue("Group1") = 0 Then
	ActiveDocument.SelectObjects(9999, "", False)
	t = "Pins(s) : "
	For i = 0 To p
		t = t & pins(i) & " "
		ActiveDocument.SelectObjects(3, pins(i), True)
	Next i	
	DlgText "SelText", t			
End If
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
Sub SelectPins(t$)
Dim Pin As Object
ActiveDocument.SelectObjects(9999, "", False)
ActiveDocument.SelectObjects(3, t, True)
For Each pin In	ActiveDocument.GetObjects(3, t, True)
	ActiveDocument.ActiveView.Pan(0,pin.PositionX(0),pin.PositionY(0))
Next pin	
End Sub
'********************************************************************************
'********************************************************************************
'********************************************************************************
'						Common Functions
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
Sub SortItems(index%,comps() As String)
Dim chg%, i%
Dim hstring$, lstring$, tmp$
	chg=True
	While chg=True
		chg=False
		For i = 0 To index-1
				Lstring=UCase(Comps(i))
				Hstring=UCase(Comps(i+1))
			If Lstring > Hstring Then
				tmp=Comps(i)
				Comps(i)=Comps(i+1)
				Comps(i+1)=tmp
				chg=True
			End If
		Next i
	Wend
End Sub
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































