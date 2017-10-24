''''#Uses "E:\Elec\Pads\ole\vbscripts\Module\Standard.BAS"
Option Explicit
Const RegString$="PadsScript"
Const RegAppString$="Copy Components"

Global comps() As String
Global compX() As Double
Global compY() As Double

Global refs() As String
Global vals() As String
Dim ErrString$
Dim compnr%, refnr%
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
Sub Main

Call GetCopyData
Call sortitems(refnr,refs())	

	Begin Dialog UserDialog 370,147,"Component Group Copy",.IncCallBack
		GroupBox 10,0,350,140,"",.GroupBox1
		Text 20,14,130,14,"Refs          Inc",.Text3
		PushButton 200,112,70,21,"Close",.CloseButton
		PushButton 280,112,70,21,"Apply",.ApplyButton
		TextBox 20,28,140,70,.Liste,1
		Text 170,31,50,14,"Repeat",.Text1
		TextBox 230,28,120,21,.RepeatBox
		TextBox 230,56,120,21,.XOffsetBox
		Text 170,60,50,14,"XOffset",.Text2
		Text 170,80,50,14,"YOffset",.Text4
		TextBox 230,77,120,21,.YOffsetBox
		TextBox 100,112,60,21,.SetAllBox
		PushButton 20,112,70,21,"Set All",.SetAllButton
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
Private Function IncCallBack(DlgItem$, Action%, SuppValue%) As Boolean
Static focus$
Dim i%,t$,j%
'Debug.Print "Action "& action & "     DlgString "& dlgitem & "     SuppValue "& SuppValue & "    Focus " & focus
	Select Case Action%
	Case 1 
	Case 2 
		If dlgitem = "CloseButton" Then
			Call SaveReg	
		ElseIf dlgitem = "ApplyButton" Then
			errstring=""
			Call SaveReg
			If GetVals(DlgText("Liste")) = True Then
				IncCallBack = True
				i = Val(DlgText("RepeatBox"))
				If i = 0 Then i = 1
				Call CheckNewRefs(i)
				If errstring <> "" Then Exit Function
				For j = 1 To i
				Call ProcessCopy
				Next j
				ProcessCommand(8)
				ProcessCommand(8)
			Else
				MsgBox "Do not edit the left column.",16,"I am confused !!"		
				Call main
			End If
		ElseIf dlgitem = "SetAllButton" Then
				IncCallBack = True
				Call SetAllIncs
		End If
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
		focus = dlgitem	
	Case 5 ' Idle
		DlgText "xoffsetbox", GetSetting(RegString, RegAppString,"XOffset")
		DlgText "yoffsetbox", GetSetting(RegString, RegAppString,"yOffset")
		DlgText "RepeatBox", GetSetting(RegString, RegAppString,"Repeat")
		DlgText "SetallBox", GetSetting(RegString, RegAppString,"SetAll")
		For i = 0 To refnr
			t = t & refs(i) & Chr(9) & vals(i) & Chr(13) & Chr(10)
		Next i
		t = Left(t,Len(t)-2)
		DlgText "Liste",t
	End Select
End Function
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
Function Getvals(t$) As Integer

Dim Lines%,i%,u$,v$

lines = countel(t,Chr(13))
For i = 1 To lines
	u = getel(t,Chr(13),i)
	If Left(u,1) = Chr(10) Then u = Mid(u,2,Len(u)-1)
	v = getel(u,Chr(9),1)
	If v <> refs(i-1) Then 
		Exit Function
	Else 
		vals(i-1)	= getel(u,Chr(9),2)
	End If
Next i
getvals = True
End Function
'*****************************************************************************
'*****************************************************************************
'*****************************************************************************
'*****************************************************************************
Sub ProcessCopy
Dim nextObj As Object
Dim i%, j%, k%, t$,u$,inc%
Dim found%,xoff%,yoff%
Dim ncomp As Object

xoff = Val(DlgText("xoffsetbox"))
yoff = Val(DlgText("yoffsetbox"))

Static CopyCount%
If CopyCount = 0 Then CopyCount = 1

If DlgText ("XOffsetBox") = "" Or DlgText ("YOffsetBox") = "" Then
	 MsgBox "Set X and Y Offset and try again.",33,"Set Offset" 
	Exit Sub
End If

ActiveDocument.SelectObjects(9999, "", False)
For i = 0 To compnr
	ActiveDocument.SelectObjects(1,comps(I) , True)
    t =  GetNewRef(i,CopyCount)
	ProcessCommand(239)
 	SendKeys t & "{ENTER}"
	ProcessCommand(882)
	ProcessPointer(40,0,compX(i) + (xoff * CopyCount),compY(i) + (yoff * CopyCount))
	ProcessPointer(4,0,compX(i) + (xoff * CopyCount),compY(i) + (yoff * CopyCount))
Next i
CopyCount = CopyCount + 1
End Sub
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
Sub GetCopyData
Dim i%, j%, t$
Dim found%, k%, inctext$

inctext =GetSetting(RegString, RegAppString,"SetAll")


	Dim nextComp As Object
	For Each nextComp In ActiveDocument.GetObjects(1, "", True)
		ReDim Preserve comps(i) As String
		ReDim Preserve compX(i) As Double
		ReDim Preserve compY(i) As Double
		comps(i) = nextComp.Name
		compX(i) = nextComp.PositionX(0)
		compY(i) = nextComp.PositionY(0)
		i = i + 1
	Next nextComp
	compnr = i -1
	If i = 0 Then
		Call MsgBox ("You must select some components.",16,"Group Copy Error") 
		End
	End If
	
	For i = 0 To compnr
		For j = 1 To Len(comps(i))
		If Asc(Mid(comps(i),j,1)) >= 48 And Asc(Mid(comps(i),j,1)) <= 57 Then
			t = Left(comps(i),j-1)
			found = False
			For k = 0 To refnr - 1
				If refs(k) = t Then 
					found = True
					Exit For
				End If
			Next k
			If found = False Then
				ReDim Preserve refs(refnr) As String
				ReDim Preserve vals(refnr) As String
				vals(refnr) = IncText
				refs(refnr) = t
				refnr = refnr + 1
			End If
     		Exit For
		End If	
		Next j 
	Next i
refnr = refnr - 1
End Sub
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
Sub SetAllIncs
Dim Lines%,i%,u$,v$,t$

For i = 0 To refnr
	vals(i) = DlgText("SetallBox")
Next i
For i = 0 To refnr
	t = t & refs(i) & Chr(9) & vals(i) & Chr(13) & Chr(10)
Next i
t = Left(t,Len(t)-2)
DlgText "Liste",t

End Sub
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************

Sub SaveReg
SaveSetting RegString, RegAppString,"XOffset",DlgText("Xoffsetbox")
SaveSetting RegString, RegAppString,"YOffset",DlgText("Yoffsetbox")
SaveSetting RegString, RegAppString,"Repeat",DlgText("RepeatBox")
SaveSetting RegString, RegAppString,"Setall",DlgText("SetAllBox")
End Sub
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
Sub CheckNewRefs(Ct%)
Dim i%,j%,cpnr%,x%

ReDim chk((compnr+1) * ct) As String
ReDim cp(0) As String
Dim nextComp As Object

For Each nextComp In ActiveDocument.GetObjects(1, "", False)
	ReDim Preserve cp(i) As String
	cp(i) = nextComp.Name
	i = i + 1
Next nextComp
cpnr = i - 1

For i = 1 To ct
	For j = 0 To compnr
		chk(x) = GetNewRef(j,i)
		If Len(chk(x)) > 6 Then 
			ErrString =  ErrString & Chr(13) & Chr(10) & "Reference " & cp(i) & "  Name is too long."
		End If
		x = x + 1
	Next j
Next i

For i = 0 To cpnr
	For j = 0 To (compnr+1) * ct
		If cp(i) = chk(j) Then 
			ErrString =  ErrString & Chr(13) & Chr(10) & "Reference " & cp(i) & "  already used."
		End If
	Next j
Next i
If ErrString <> "" Then
	Call ErrForm
End If
End Sub
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
Function GetNewRef(nr%,ct%) As String
Dim j%,k%,inc%,found%
Dim u$,t$
	For j = 1 To Len(comps(nr))
		found = False
		If Asc(Mid(comps(nr),j,1)) < 65 Then 
			t = Left(comps(nr),j-1)
			u = Right(comps(nr),Len(comps(nr)) - j+1)
			For k = 1 To Len(u)
				If Asc(Mid(u,k,1)) > 64 Then
					MsgBox "Cannot deal with References like " & t & u & "." & Chr(13) & Chr(10) & "Rename before copy.",16,"Error"
				    End	
				End If 
			Next k
			Exit For
		End If
	Next j	
		For k = 0 To refnr
			If t = refs(k) Then
				inc = Val(Vals(k))
				found = True
				Exit For
			End If
			If found = True Then Exit For
		Next k
	inc = (inc * Ct) + Val(u)
	GetNewRef = t & LTrim(Str(inc))
End Function
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
Sub ErrForm
	Begin Dialog UserDialog 400,252,"Used References",.ErrCallBack
		GroupBox 10,0,380,245,"",.GroupBox1
		TextBox 20,14,360,196,.ErrText,1
		PushButton 260,217,120,21,"Close",.PushButton1
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
Private Function ErrCallBack(DlgItem$, Action%, SuppValue%) As Boolean
	Select Case Action%
	Case 5 ' Idle
		DlgText "ErrText","There are some Errors :" & Chr(13) & Chr(10) & errstring
	End Select
End Function
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
'            Common Funtions
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************

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
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
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
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
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
'****************************************************************************************************
'****************************************************************************************************
'****************************************************************************************************
