'''''''#Uses "E:\Elec\Pads\ole\vbscripts\Module\Standard.BAS"
Option Explicit

Const RegString$="PadsScript"
Dim index%
Dim WildCard$
Dim CompName$
Dim FName$
Dim ActTemplate$
Dim ActTemp%
Dim UserDir$
Dim OnlyNets%, FormatID1%, FormatID2%

'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Sub Main

	ReDim comps(0)As String
	ReDim pins(0) As String
	ReDim nets(0) As String
	ReDim Templates(0) As String

	WildCard = GetSetting (RegString,"ListChipPins","WildCard")
	FName= GetSetting (RegString,"ListChipPins","FName")
	UserDir= GetSetting (RegString,"ListChipPins","UserDir")

	Call GetTemplateNames
	Call compdlg
	index = GetPins(compname, Pins(), nets())
    If MsgBox("Template : " & ActTemplate & Chr(13)&Chr(10)&"Component : "& compname,33,"Make Report")= vbOK Then Call MakeReport
	SaveSetting RegString,"ListChipPins","WildCard",WildCard
	SaveSetting RegString,"ListChipPins","UserDir",UserDir
	SaveSetting RegString,"ListChipPins","FName",FName
End Sub
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'**********************************************  ----  Component Form  ----  ************************************************
Private Sub CompDLG
Dim index%
ReDim comps(0) As String
index =  GetComps(comps())
Call SortItemNrL(index,comps())
 

	Begin Dialog UserDialog 520,182,"Component Pin Report",.CompCallBack
		GroupBox 10,49,250,98,"Component",.GroupBox1
		ListBox 20,63,150,77,Comps(),.CompListBox
		OKButton 420,154,90,21
		TextBox 180,84,70,21,.WildCardText
		PushButton 180,119,70,21,"Refresh",.RefreshButton
		PushButton 320,154,90,21,"Cancel",.CancelButton
		GroupBox 10,0,500,42,"Report File",.GroupBox2
		TextBox 20,14,410,21,.FileTextBox
		PushButton 440,14,60,21,"...",.FileButton
		GroupBox 270,49,240,98,"Templates",.GroupBox3
		ListBox 280,63,220,63,Templates(),.TemplateBox
		Text 180,63,60,14,"WildCard",.Text1
		PushButton 280,119,220,21,"Templates",.TemplateSetButton
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Private Function CompCallBack(DlgItem$, Action%, SuppValue%) As Boolean
Dim t$,index%
Static ActFocus$
Debug.Print "Action "& action & "     DlgString "& dlgitem & "     SuppValue "&SuppValue & "   Focus  " & actfocus
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText "WildCardText", WildCard
		compname=comps(0)
		DlgText "FileTextBox", UserDir & Fname
	Case 2 ' Value changing or button pressed
		If dlgitem$ ="RefreshButton" Then
			ReDim comps(0) As String
			index =  GetComps(comps())
			Call SortItemNrL(index,comps())
			DlgListBoxArray "CompListBox",Comps()
			CompCallBack = True
		ElseIf dlgitem$ ="CompListBox" Then
			compname=comps(SuppValue%)
		ElseIf dlgitem$ ="CancelButton" Then
			End
		ElseIf dlgitem$ ="FileButton" Then
			CompCallBack = True
			t = GetFilePath(,, UserDir,"Select Reportfile",2)
			If t <> "" Then
				Call SplitPath(t, UserDir, Fname)
				DlgText "FileTextBox", UserDir & Fname
			End If
		ElseIf dlgitem$ ="TemplateSetButton"  Then
				CompCallBack = True
				Call StartSetup
		ElseIf dlgitem$ ="TemplateBox"	Then
				CompCallBack = True
		ElseIf dlgitem$ ="OK" Then
			If actfocus <> "OK" Then CompCallBack = True
			If actfocus = "TemplateBox" Then Call StartSetup
			ActTemplate = Templates(DlgValue("TemplateBox")+1)
		End If

	Case 3 ' TextBox or ComboBox text changed
		WildCard=UCase(DlgText$("WildCardText") )
		
	Case 4 ' Focus changed
		ActFocus = DlgFocus
	Case 5 ' Idle
		Rem CompCallBack = True ' Continue getting idle actions
	End Select
End Function
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Private Function GetComps(comps() As String)
Dim index%
Dim nextComp As Object
	For Each nextComp In ActiveDocument.GetObjects(1, "", False)
	If WildCard="" Then
		ReDim Preserve comps(index) As String
		Comps$(index) = nextComp.Name
		index = index + 1
	Else
		If Left(nextComp.Name,Len(WildCard))=Wildcard Then
			ReDim Preserve comps(index) As String
			Comps$(index) = nextComp.Name
			index = index + 1
		End If	
	End If
	Next nextComp
	getcomps = index-1
End Function
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Private Function GetPins(compname$, Pins() As String, Nets() As String)
Dim i%
Dim nextPin As Object

	For Each nextPin In ActiveDocument.GetObjects(3, "", False)
		i = InStr(1,nextPin.Name, ".")
		If Left(nextPin.Name,i-1)=compname Then
			ReDim Preserve Pins(index) As String
			ReDim Preserve Nets(index) As String
			On Error GoTo EmptyField
				Pins(index) = getel(nextPin.Name,".",2)
				Nets(index) = nextPin.Net
			On Error GoTo 0
			index = index + 1
		End If
	Next nextPin	
	GetPins=index-1
	
Exit Function
EmptyField:
	nets(index)=""
Resume Next	
End Function
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Private Sub MakeReport

Dim i%, x&, j%
Dim string1$
Dim string2$
Dim string3$
Dim string4$
Dim string5$
Dim t$,u$,v$
Dim MaxlenID%
Dim Sort%

'If exists(UserDir & fname) And MsgBox ("File  " & UserDir & fname & "  exists.",49,"Warning") =vbCancel Then Exit Sub
If fname = ""  Then Call MsgBox ("No Valid Filename",16,"Error") : Exit Sub

string1 = GetSetting (RegString,"ListChipPins\Templates\"&ActTemplate,"String1")
string2 = GetSetting (RegString,"ListChipPins\Templates\"&ActTemplate,"String2")
string3 = GetSetting (RegString,"ListChipPins\Templates\"&ActTemplate,"String3")
string4 = GetSetting (RegString,"ListChipPins\Templates\"&ActTemplate,"String4")
string5 = GetSetting (RegString,"ListChipPins\Templates\"&ActTemplate,"String5")
OnlyNets = Val(GetSetting (RegString,"ListChipPins\Templates\"&ActTemplate,"OnlyNets"))
FormatID1 = Val(GetSetting (RegString,"ListChipPins\Templates\"&ActTemplate,"FormatID1"))
FormatID2 = Val(GetSetting (RegString,"ListChipPins\Templates\"&ActTemplate,"FormatID2"))
Sort = Val(GetSetting (RegString,"ListChipPins\Templates\"&ActTemplate,"Sort"))

If Sort = 0 Then 
	If string2 = "Pin" Then
		Call SortVals(index%, pins(), nets())
	Else
		Call SortItemNrLSpec(index%, nets(), pins())
	End If	
ElseIf Sort = 1 Then 
	If string4 = "Pin" Then
		Call SortVals(index%, pins(), nets())
	Else
		Call SortItemNrLSpec(index%, nets(), pins())
	End If	
End If

ReDim o(0) As String

Open Userdir & Fname For Output As #1

For i = 0 To UBound(pins())
	If OnlyNets = 0 Or (Onlynets = 1 And nets(i) <> "") Then
		ReDim Preserve o(j) As String
		If string2 = "Pin" Then o(j) = string1 & Pins(i) Else o(j) = string1 & Nets(i)
		j = j + 1
	End If
Next i

If FormatID1 = 1 Then
	For i = 0 To UBound(o())
		If MaxlenID < Len(o(i)) Then MaxlenID = Len(o(i))
	Next i
	MaxlenID = MaxlenID + 1
	For i = 0 To UBound(o())
		o(i) = o(i) + Space(maxlenID -Len(o(i)))
	Next i
End If

For i = 0 To UBound(o())
	o(i) = o(i) + string3
Next i

j = 0
For i = 0 To UBound(pins())
	If OnlyNets = 0 Or (Onlynets = 1 And nets(i) <> "") Then
		If string4 = "Pin" Then o(j) = o(j) & Pins(i) Else o(j) = o(j) & Nets(i)
		j = j + 1
	End If
Next i

MaxlenID = 0
If FormatID2 = 1 Then
	For i = 0 To UBound(o())
		If MaxlenID < Len(o(i)) Then MaxlenID = Len(o(i))
	Next i
	MaxlenID = MaxlenID + 1
	For i = 0 To UBound(o())
		o(i) = o(i) + Space(maxlenID -Len(o(i)))
	Next i
End If
For i = 0 To UBound(o())
	o(i) = o(i) + string5
	Print #1, o(i)
Next i

Close #1
x=Shell("notepad " & UserDir & fname,1)

End Sub
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Sub SetupForm

ReDim ComboText(1) As String
Combotext(0)="PinNumber"
Combotext(1)="NetName"


	ReDim listarray(1) As String
	Begin Dialog UserDialog 630,168,"Setup",.SetupDlgFunc
		GroupBox 270,56,200,49,"List",.GroupBox3
		GroupBox 480,56,140,49,"Sort",.GroupBox2
		GroupBox 10,7,610,42,"Line ",.GroupBox1
		TextBox 20,21,110,21,.String1Box
		TextBox 260,21,110,21,.String3Box
		TextBox 500,21,110,21,.String5Box
		OKButton 480,140,140,21
		DropListBox 140,21,110,70,ComboText(),.String2Box
		DropListBox 380,21,110,70,ComboText(),.String4Box
		OptionGroup .Group1
			OptionButton 490,70,100,14,"Sort first ID",.Sort1
			OptionButton 490,84,120,14,"Sort second ID",.Sort2
		OptionGroup .Group2
			OptionButton 280,70,100,14,"List All Pins",.OptionAll
			OptionButton 280,84,180,14,"List only connected Pins",.OptionOnly
		GroupBox 10,56,250,105,"Templates",.GroupBox5
		PushButton 20,133,70,21,"Add",.AddButton
		GroupBox 270,112,200,49,"Format",.GroupBox4
		CheckBox 280,126,150,14,"Format after first ID",.FormatID1Box
		CheckBox 280,140,170,14,"Format after second ID",.FormatID2Box
		ListBox 20,70,230,56,Templates(),.TemplateBox
		PushButton 100,133,70,21,"Delete",.DeleteButton
		PushButton 180,133,70,21,"Save",.SaveButton
	
	End Dialog
	Dim dlg As UserDialog

	Dialog dlg


End Sub
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Private Function SetupDlgFunc(DlgItem$, Action%, SuppValue%) As Boolean
Dim t$, i%, j%, u$
Dim Found%
Static focus$
Debug.Print "Action "& action & "     DlgString "& dlgitem & "     SuppValue "&SuppValue & "   Focus " & focus

	Select Case Action%
	Case 1 ' Dialog box initialization

	Case 2 ' Value changing or button pressed

		If dlgitem$ ="TemplateBox" Then
			Call GetRegStrings(suppvalue+1)
			acttemp = suppvalue
			DlgValue "TemplateBox", acttemp

 		ElseIf dlgitem$ ="AddButton" Then
			t = InputBox("Enter Model Name")
			If t = "" Then 
				SetupDlgFunc = True
				Exit Function
			End If
			found = False
			For i = 0 To Val(GetSetting (RegString,"ListChipPins","Templates"))
				If UCase(t) = UCase(Templates(i)) Then
					found = True
					If MsgBox (t & "  exists.",49,"Warning") <> vbOK Then
						SetupDlgFunc = True
						Exit Function
					End If 
				Exit For
				End If
			Next i
			If found = False Then
				ReDim Preserve Templates(UBound(Templates()) + 1) As String
				Templates(UBound(Templates())) = t
				Call sortitems(UBound(Templates()), Templates())
				j = Val(GetSetting (RegString,"ListChipPins","Templates")) +1
				SaveSetting(RegString,"ListChipPins", "Templates",j)
				SaveSetting(RegString,"ListChipPins\Templates", Trim(Str(j)), t)
				Call SaveRegStrings(t)
				DlgListBoxArray "TemplateBox",Templates()
			End If 	
			SetupDlgFunc = True
 		ElseIf dlgitem$ ="SaveButton" Then
			Call SaveRegStrings(Templates(DlgValue("TemplateBox") + 1))
			SetupDlgFunc = True
 		ElseIf dlgitem$ = "DeleteButton" Then
				SetupDlgFunc = True
				i = MsgBox("Delete  " & Templates(DlgValue("TemplateBox")+1) & " ?",33)
				If i <>vbOK Then Exit Function
				Call DeleteTemplate(DlgValue("TemplateBox")+1)
				DlgListBoxArray "TemplateBox",Templates()
		ElseIf dlgitem$ = "FormatID1Box" Then
				FormatID1 = suppvalue
		ElseIf dlgitem$ = "FormatID2Box" Then
				FormatID2 = suppvalue
		ElseIf dlgitem$ = "OK" Then
			If focus <> "OK" Then 
				SetupDlgFunc = True
			End If	
		End If	


		Rem SetupDlgFunc = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 
		focus = dlgitem
	Case 5 ' Idle
			Call GetRegStrings(acttemp+1)
			DlgValue "TemplateBox", acttemp
	End Select
End Function
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Sub GetTemplateNames
Dim m%,i%,t$
		m = Val(GetSetting (RegString,"ListChipPins","Templates"))
		ReDim Templates(m) As String
		For i = 1 To m 		
			t = "ListChipPins\Templates"
			Templates(i)=GetSetting(RegString,t, LTrim(Str(i)))
		Next i
		Call Sortitems(m, Templates())
End Sub
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Sub DeleteTemplate(x%)
Dim i%, t$, f%,m%
t = Templates(x)
Templates(x)=""
m = GetSetting (RegString,"ListChipPins","Templates")
For i = 1 To  m - 1
	If Templates(i) = "" Then f=True
	If f=True Then
		Templates(i)=Templates(i+1)
	End If
Next i
ReDim Preserve Templates(m-1) As String

For i = 1 To m
	DeleteSetting(RegString,"ListChipPins\Templates", Trim(Str(i)))
Next i
For i = 1 To m-1
	SaveSetting (RegString, "ListChipPins\Templates", Trim(Str(i)), Templates(i))
Next i
SaveSetting (RegString,"ListChipPins","Templates", m-1)

On Error Resume Next
DeleteSetting(RegString,"ListChipPins\Templates\" & t, "String1")
DeleteSetting(RegString,"ListChipPins\Templates\" & t, "String2")
DeleteSetting(RegString,"ListChipPins\Templates\" & t, "String3")
DeleteSetting(RegString,"ListChipPins\Templates\" & t, "String4")
DeleteSetting(RegString,"ListChipPins\Templates\" & t, "String5")
DeleteSetting(RegString,"ListChipPins\Templates\" & t, "Sort")
On Error GoTo 0
End Sub
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Private Sub SortItemNrLSpec(index%, comps()As String, nets() As String)
Dim chg%, i%, j%
Dim hstring$, lstring$, tmp$
Dim HNr%, LNr%
	chg=True
	While chg=True
		DoEvents
		chg=False
		For i = 0 To index-1
			For j = Len(comps(i)) To 1 Step -1 
				If Asc(Mid(comps(i),j,1))>= 65 Or Mid(comps(i),j,1) = "$" Then
					lstring = Left(comps(i),j)
					LNr = Val(Right(comps(i),Len(comps(i))-j))
					Exit For
				End If
			Next j				
			For j = Len(comps(i + 1)) To 1 Step -1 
				If Asc(Mid(comps(i + 1),j,1))>= 65 Or Mid(comps(i + 1),j,1) = "$" Then
					hstring = Left(comps(i + 1),j)
					HNr = Val(Right(comps(i + 1),Len(comps(i + 1))-j))
					Exit For
				End If
			Next j				
			If (Lstring > Hstring) Or (Lstring = Hstring And LNr > HNr) Then
				tmp=Comps(i)
				Comps(i)=Comps(i+1)
				Comps(i+1)=tmp
				chg=True
				tmp=nets(i)
				nets(i)=nets(i+1)
				nets(i+1)=tmp
			End If
		Next i
	Wend
End Sub
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Sub SortVals(index%, comps$(), nets$())
Dim chg%, i%, j%
Dim hstring$, lstring$, tmp$
Dim HNr%, LNr%
	chg=True
	While chg=True
		DoEvents
		chg=False
		For i = 0 To index-1
			If Val(comps(i)) > Val(comps(i+1)) Then
				tmp=Comps(i)
				Comps(i)=Comps(i+1)
				Comps(i+1)=tmp
				chg=True
				tmp=nets(i)
				nets(i)=nets(i+1)
				nets(i+1)=tmp
			End If
		Next i
	Wend
End Sub
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Sub GetRegStrings(nr%)
Dim t$
DlgText "String1Box", GetSetting (RegString,"ListChipPins\Templates\"&Templates(nr),"String1")
DlgText "String3Box", GetSetting (RegString,"ListChipPins\Templates\"&Templates(nr),"String3")
DlgText "String5Box", GetSetting (RegString,"ListChipPins\Templates\"&Templates(nr),"String5")
t = GetSetting (RegString,"ListChipPins\Templates\"&Templates(nr),"String2")
If t = "Pin" Then DlgValue "String2Box",0 Else DlgValue "String2Box",1
t = GetSetting (RegString,"ListChipPins\Templates\"&Templates(nr),"String4")
If t = "Pin" Then DlgValue "String4Box",0 Else DlgValue "String4Box",1
DlgValue "Group1", Val(GetSetting (RegString,"ListChipPins\Templates\"&Templates(nr),"Sort"))
DlgValue "Group2", Val(GetSetting (RegString,"ListChipPins\Templates\"&Templates(nr),"OnlyNets"))
DlgValue "FormatID1Box", Val(GetSetting (RegString,"ListChipPins\Templates\"&Templates(nr),"FormatID1"))
DlgValue "FormatID2Box", Val(GetSetting (RegString,"ListChipPins\Templates\"&Templates(nr),"FormatID2"))
DlgValue "TemplateBox", nr
End Sub
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Sub SaveRegStrings(t$)
Dim u$
SaveSetting(RegString,"ListChipPins\Templates\" & t,"String1",DlgText("String1Box"))
SaveSetting(RegString,"ListChipPins\Templates\" & t,"String3",DlgText("String3Box"))
SaveSetting(RegString,"ListChipPins\Templates\" & t,"String5",DlgText("String5Box"))
If DlgValue("String2Box") = 0 Then u = "Pin" Else u = "Net"
	SaveSetting(RegString,"ListChipPins\Templates\" & t,"String2", u)
If DlgValue("String4Box") = 0 Then u = "Pin" Else u = "Net"
	SaveSetting(RegString,"ListChipPins\Templates\" & t,"String4", u)
SaveSetting(RegString,"ListChipPins\Templates\" & t,"Sort",DlgValue("Group1"))
SaveSetting(RegString,"ListChipPins\Templates\" & t,"OnlyNets",DlgValue("Group2"))
SaveSetting(RegString,"ListChipPins\Templates\" & t,"FormatID1", FormatID1)
SaveSetting(RegString,"ListChipPins\Templates\" & t,"FormatID2", FormatID2)
End Sub
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Sub StartSetup
	ActTemp = DlgValue("TemplateBox")
	If ActTemp < 0 Then ActTemp = 0
	Call SetupForm
	DlgListBoxArray "TemplateBox",Templates()
	DlgValue"TemplateBox",0
End Sub
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'										Common Functions
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Sub SortItemNrL(index%,comps() As String)
Dim chg%, i%, j%
Dim hstring$, lstring$, tmp$
Dim HNr%, LNr%
	chg=True
	While chg=True
		DoEvents
		chg=False
		For i = 0 To index-1
			For j = Len(comps(i)) To 1 Step -1 
				If Asc(Mid(comps(i),j,1))>= 65 Then
					lstring = Left(comps(i),j)
					LNr = Val(Right(comps(i),Len(comps(i))-j))
				End If
			Next j				
			For j = Len(comps(i + 1)) To 1 Step -1 
				If Asc(Mid(comps(i + 1),j,1))>= 65 Then
					hstring = Left(comps(i + 1),j)
					HNr = Val(Right(comps(i + 1),Len(comps(i + 1))-j))
				End If
			Next j				
		
			If Lstring > Hstring Then
				tmp=Comps(i)
				Comps(i)=Comps(i+1)
				Comps(i+1)=tmp
				chg=True
			ElseIf Lstring = Hstring And LNr > HNr Then
				tmp=Comps(i)
				Comps(i)=Comps(i+1)
				Comps(i+1)=tmp
				chg=True
			End If
		Next i
	Wend
End Sub
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
Sub SplitPath(InString$, PathString$, FileString$)
Dim x%
For x = Len(InString) To 1 Step -1
If Mid(InString, x, 1)="\" Then
	PathString = Left(InString, x)
	FileString = Right(InString, Len(InString) - x)
	Exit For
End If
Next x
End Sub
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
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
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
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
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************
'******************************************************************************************************************************

