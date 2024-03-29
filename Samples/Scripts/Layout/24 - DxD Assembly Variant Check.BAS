' Sample 24: DxDesigner Assembly Variant Check.BAS
' 
' This sample allow to check if the assembly variants generated by DxDesigner program
' (*_variants.asc file) and assembly variants from current pcb design are synchronized.
' 
'
Option Explicit

Type partNew				'part type declaration
	refDes As String		'reference designator of a part
	varChange() As String	'indicate what change is made in a variant (S - substituted, I - installed, U - uninstalled)
	substituteName() As String 'name of substituted component
End Type

Dim partsNew() As partNew	'dynamic array of parts from variant ASCII file
Dim pcbParts() As partNew	'dynamic array of parts currently in the pcb design
Dim pcbPartsCount As Long	'number of parts in pcb design
Dim partsCount As Long		'number of parts in varianr ASCII file
Dim dxdProjectDir$			'DxDesigner project path
Dim previousAsciiName$		'name of previous variant ASCII file
Dim dxdVariantFile$ 		'ASCII file name(*_variants.asc)
Dim VariantList$()			'List of variants from variant file and pcb design
Dim dxVariantsNum As Integer	'number of variants in variant ASCII file
Dim extraPCBVarNum As Integer	'number of variants that are in pcb design and doesn't exist in variant ASCII file
Dim outputFile$ 			'output file name
Dim mismatches As Long		'number of mismatches, including incorrect variants and incorrect components status in variants

Sub Main
	Begin Dialog UserDialog 610,222,"Dx Assembly Variant Checker ",.CallbackFunc
		TextBox 20,24,570,21,.EdtProjectName
		TextBox 20,69,470,21,.EdtASCIIName
		TextBox 20,165,470,21,.EdtOutFileName
		DropListBox 20,117,310,100,VariantList(),.variantList
		PushButton 1,1,0,0,"",.BtnUnused
		PushButton 500,195,90,21,"Close",.BtnClose
		PushButton 380,195,90,21,"Check",.BtnCheck
		Text 20,7,300,14,"PADS Layout pcb file",.Text1
		Text 20,52,300,14,"Dx Variant ASCII file",.Text2
		Text 20,99,310,15,"Variant:",.Text3
		Text 20,147,310,15,"Output File Name",.Text4
		Text 370,120,120,15,"Check All Variants",.Text5
		CheckBox 340,120,20,15,"Check",.Check
		PushButton 500,69,90,21,"Browse...",.BtnOpenAscii
		Text 20,198,330,15,"status",.status
		PushButton 500,165,90,21,"Browse...",.BtnOpenOutput
	End Dialog
	
	Dim dlg As UserDialog
	Dialog dlg
	
End Sub

Private Function CallbackFunc(DlgItem$, Action%, Supp3Value%) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText "status", "Working"
		DlgText "EdtProjectName", ActiveDocument.Name
	    DlgEnable "EdtProjectName", False
	    DlgText "EdtOutFileName", ActiveDocument.path & "\DxVariantCheck.rep"
	    clearVariantList
	    Initialize(True)
	    
	Case 2 ' Value changing or button pressed
		CallbackFunc = True ' Prevent button press from closing the dialog box
		If DlgItem$ = "BtnClose" Then
			closeBtnPressed
			CallbackFunc = False
		ElseIf (DlgItem$ = "BtnCheck") Then 
			checkProcess
		ElseIf (DlgItem$ = "BtnOpenAscii")Then
			OpenAsciiPressed
		ElseIf (DlgItem$ = "BtnOpenOutput")Then
			OpenOutputPressed
		ElseIf (DlgItem$ = "variantList")Then
			'do nothing
		Else
			If ( Supp3Value% = 0 )Then
				DlgEnable "variantList", True
			Else
				DlgEnable "variantList", False
			End If
		End If
	
	Case 3 ' TextBox or ComboBox text changed
		If DlgItem$ = "EdtASCIIName" Then
			previousAsciiName$ = DlgText("EdtASCIIName")
			dxdProjectDir$ = ""
			dxdVariantFile$ = previousAsciiName$'DlgText("EdtASCIIName")
			'If( setCorrectNames() )Then
				Initialize(False)
			'End If
		End If
	
	Case 4 ' Focus changed
	Case 5 ' Idle
	Case 6 ' Function key
	End Select
End Function

Public Sub Application_OpenDocument( Doc As PowerPCB.Document )
	DlgText "EdtProjectName",ActiveDocument.Name
	Initialize(False)
End Sub

Private Sub closeBtnPressed
	Erase VariantList$()
End Sub

Private Sub OpenAsciiPressed
	Dim tmpStr$ 'As String
	tmpStr$ = DlgText("EdtASCIIName")

	If( dxdProjectDir$ = "" )Then
		dxdVariantFile$ = GetFilePath$( , "asc", ActiveDocument.path, "Choose DxDesigner Variant ASCII file", 0 )
	Else
		dxdVariantFile$ = GetFilePath$( , "asc", dxdProjectDir$, "Choose DxDesigner Variant ASCII file", 0 )
	End If
	
	If ( dxdVariantFile$ = "" )Then
		dxdVariantFile$ = tmpStr$
		setCorrectNames
		DlgText "EdtASCIIName", tmpStr$
	Else
		If( setCorrectNames = False )Then
			Dim message$
			message$ = "It's not a variant file!"& vbCrLf &"DxD variant ASCII files' names ends with '_variants.asc'."
			MsgBox message$, 48,"Incorrect file."
			previousAsciiName$ = tmpStr$
			Exit Sub
		End If
		previousAsciiName$ = dxdVariantFile$
		Initialize(False)
	End If
End Sub

Public Sub OpenOutputPressed
	Dim tmpStr$ 'As String
	Dim fName$ 'As String	'file name
	Dim dname$ 'As String	'file path
	Dim pos As Integer
	
	tmpStr$ = DlgText( "EdtOutFileName" )
	pos = InStrRev( tmpStr$, "\" )
	
	If( pos <> 0 )Then
		fName$ = Mid$( tmpStr$, pos+1, Len( tmpStr$ ) )
		dname$ = Mid$( tmpStr$, 1, pos-1 ) 
	Else
		fName$ = tmpStr$
		dname$ = ""
	End If
	
	If( dname$ <> "" )Then
		dname$ = GetFilePath$( fName$, "rep;txt", dname$, "Choose output report file", 2 )
	Else
		dname$ = GetFilePath$( fName$, "rep;txt", ActiveDocument.path, "Choose output report file", 2 )
	End If
	
	If( dname$ = "" )Then
		DlgText "EdtOutFileName", tmpStr$
	Else
		DlgText "EdtOutFileName", dname$
	End If
		
End Sub

Private Function setCorrectNames As Boolean
	setCorrectNames = True
	
	If( dxdVariantFile$ = "" )Then
		setCorrectNames = False
		Exit Function
	End If
	
	Dim pos
	
	pos = InStrRev( dxdVariantFile$, "\" )
	If( pos > 0 )Then
		dxdProjectDir$ = Mid$( dxdVariantFile$, 1, pos )
		dxdVariantFile$ = Mid$( dxdVariantFile$, pos+1, Len( dxdVariantFile$ ) )
	End If
	
End Function

Private Sub Initialize (firstInit As Boolean)
	
	'check if DxDesigner is working currently
	If ( dxdVariantFile$ = "" And firstInit ) Then
		Dim DxDApp As Object
		Err = 0
		On Error Resume Next
		Set DxDApp = GetObject(,"Viewdraw.Application")
		
		If ( DxDApp Is Nothing ) Then
			'do nothing
		Else
			Dim DxDProj
			Set DxDProj = DxDApp.Project
			dxdProjectDir$ = DxDProj.path
		End If
		If ( Err <> 0 )Then
			'DlgText "status", "DxDesigner is not running."
			dxdProjectDir$ = ""
			Err = 0
		End If
		On Error GoTo 0
	End If
	
	'try to get *.asc files' directory from last used project list (project.lst)
	If ( dxdProjectDir$ = "" And dxdVariantFile$ = "" And firstInit ) Then
		Dim env$
		Dim lstFile$
		env$ = Environ("WDIR")		'get environment variable
		If( env$ <> "" ) Then
			Dim pos			
			Dim lstDir$
			Do
				pos = InStr( env$ , ";")
				If( pos = 0 )Then
					pos = Len( env$ )+1
				End If
				lstDir$ = Mid$( env$, 1, pos-1 )
				
				env$ = Mid$( env$, pos+1, Len( env$ ) )
				
				If ( (Right$ ( lstDir$,1 )) <> "\" ) Then
					lstDir$ = lstDir$ & "\"
				End If
				
				lstFile = lstDir$ & "vf\project.lst"
				
				On Error Resume Next
				lstFile$ = Dir( lstFile$ )
				If ( Err <> 0 )Then
					Err = 0
					'DlgText "status", "project.lst file not found."
					lstFile$ = ""
				End If
				On Error GoTo 0
				
				If( lstFile <> "" ) Then
					lstFile = lstDir$ & "vf\project.lst"
					'DlgText "status", "project.lst file found."
					Exit Do
				End If
			Loop While ( env$ <> "" )
		End If
		
		While ( lstFile$ <> "" )
			
			On Error Resume Next
				If ( FileLen( lstFile$ ) < 1 ) Then
					'DlgText "status", "project.lst file is empty."
					Exit While
				End If
				If ( Err <> 0 )Then
					Err = 0
					'DlgText "status", "project.lst file not found."
					Exit While
				End If

				Open lstFile$ For Input As #1
				If ( Err <> 0 )Then
					Err = 0
					'DlgText "status", "Error: can't open project.lst file."
					Exit While
				End If
			On Error GoTo 0
						
			Dim readString$ 'As String
			Dim readLineNo As Integer
			Line Input #1,readString$
			readLineNo = CInt( readString$ )
			
			While( Not EOF(1) And readLineNo >= 0 )
				Line Input #1,readString$
				readLineNo = readLineNo - 1
			Wend
			
			readLineNo = InStrRev( readString$, " " )
			If( readLineNo > 0 ) Then
				dxdProjectDir$ = Mid$(readString$, 1, readLineNo-1)
			ElseIf ( readLineNo = 0 ) Then
				dxdProjectDir$ = readString$
			End If
			
			Close #1
			lstFile = ""
		Wend
		
	End If

	setCorrectNames
	
	'if still no variant file directory set, take current pcb design directory
	If( dxdProjectDir$ = "" )Then
		dxdProjectDir$ = ActiveDocument.path & "\"
	End If
	
	'search variant file
	Err = 0
	On Error Resume Next
		If( dxdProjectDir$ = "")Then
			'do nothing
		ElseIf( firstInit Or dxdVariantFile$ = "" )Then
			dxdVariantFile$ = Dir$ (dxdProjectDir$ & "*_variants.asc")
			If( dxdVariantFile$ = "" )Then
				dxdVariantFile$ = Dir$ (dxdProjectDir$ & "*.asc")
			End If
		Else
			dxdVariantFile$ = Dir$ (dxdProjectDir$ & dxdVariantFile$)
		End If
		If ( Err <> 0 Or dxdVariantFile$ = "")Then
			Err = 0
			'DlgText "status", "DxDesigner netlist file not found!"
			dxdVariantFile$ = ""
		End If
	On Error GoTo 0
	
	Dim foundVariants As Integer	'number of found variants in pcb design and variant ASCII file
	
	setCorrectNames
	
	foundVariants = ReadVariantsNames()
	DlgEnable "variantList", False
	
	If ( foundVariants < 1 ) Then					'no variants found
		DlgText "status", "No variants found!"
		DlgEnable "check", False
		DlgValue "check", False
		DlgListBoxArray "variantList", VariantList$()
		'DlgText "variantList", "No variants"
		DlgText "variantList", VariantList$(0)
	Else											'there are variants
		DlgText "status", foundVariants & " variants found."
		DlgEnable "check", True
		DlgValue "check", True
		VariantList$(0) = ""
		DlgListBoxArray "variantList", VariantList$()
		DlgText "variantList", VariantList$(1)
	End If
	
	If( dxdVariantFile$ = "" And firstInit = False ) Then 'variant ASCII file not found
		Dim msg$
		msg$ = "File "& previousAsciiName$ &" was not found."& vbCrLf &"Check file path and name."
		MsgBox msg$, 48,"Incorrect file"
		DlgText "EdtASCIIName", previousAsciiName$
	Else
		DlgText "EdtASCIIName", dxdProjectDir & dxdVariantFile$
	End If
	
	moveCursorInASCIIToEnd
	
End Sub

Private Function ReadVariantsNames As Integer
	
	ReadVariantsNames = 0

	Dim endCode_
	Dim fileWithPath$
	Dim tmp$
	
	endCode_ = 0
	
	'clear variant list
	clearVariantList
	
	' file's name variable can't be empty !
	If ( dxdVariantFile$ = "" ) Then
'		DlgText "status", "DxDesigner variant file not found!"
		endCode_ = 1
	End If
	
	'check variant ASCII file and open it
	fileWithPath$ = dxdProjectDir & dxdVariantFile$
	On Error Resume Next
		If ( endCode_ = 0 ) Then	
			If ( FileLen( fileWithPath$ ) < 1 ) Then
'				DlgText "status", "Variant file is empty!"
				endCode_ = 1
			End If
			If ( Err <> 0 )Then
'				DlgText "status", "Variant file not found!"
				endCode_ = 1
				Err = 0
			End If 
		End If
		
		If ( endCode_ = 0 ) Then
			Open fileWithPath$ For Input As #1
			If ( Err <> 0 )Then
'				DlgText "status", "Error opening variant file!"
				endCode_ = 2
				Err = 0
			End If 
		End If
	On Error GoTo 0
	
	'read variants names from variant ASCII file
	If( endCode_ = 0 )Then
		While ( Not EOF(1) And endCode_ = 0 )
			Line Input #1,tmp$
			Dim pos
			pos = InStr( tmp$ , "ASSEMBLY_OPTIONS" )
			
			If ( pos <> 0 ) Then
				pos = Len( "ASSEMBLY_OPTIONS " )
				tmp$ = Mid$( tmp$, pos+1, Len(tmp$) )
				
				Dim commaPos
				
				While ( Len( tmp$ ) > 0 )
					commaPos = InStr( tmp$ , ",")
					If ( commaPos = 0 ) Then
						endCode_ = 3
						Exit While
					End If
					
					dxVariantsNum = dxVariantsNum + 1
					
					ReDim Preserve variantList$( dxVariantsNum )
					variantList$( dxVariantsNum ) = Mid$( tmp$, 1, commaPos-1 )
										
					pos = InStr( tmp$, ";")
					If ( pos = 0 ) Then
						endCode_ = 3
						Exit While
					End If
					tmp$ = Mid$( tmp$, pos+1, Len( tmp$ ))				
				Wend
			End If		
		Wend
	End If
	
	'close file if it was open
	If ( endCode_ >	1 ) Then
		Close #1
	End If
	
	Dim var As Integer
	
	'from pcb design get assembly variants that doesn't occur in ASCII variant file
	pos = 1
	While( pos <= ActiveDocument.AssemblyOptions.Count )
		var = 1
		tmp$ = ActiveDocument.AssemblyOptions.Item(pos).Name
		tmp$ = Mid$( tmp$, 1, InStr( tmp$, ":" ) - 1 )
		If StrComp(tmp$, "BaseOption") <> 0 Then
			While( var <= dxVariantsNum + extraPCBVarNum )
				If( tmp$ = variantList$(var) )Then
					Exit While
				End If
				var = var + 1
			Wend
			If( var > dxVariantsNum + extraPCBVarNum )Then
				extraPCBVarNum = extraPCBVarNum + 1
			
				ReDim Preserve variantList$( dxVariantsNum + extraPCBVarNum )
				variantList$( dxVariantsNum + extraPCBVarNum ) = tmp$

			End If
		End If
		pos = pos +1
	Wend
	
	'return number of found variants or 0 if no variants found
	If( dxVariantsNum + extraPCBVarNum > 0 )Then
		ReadVariantsNames = dxVariantsNum + extraPCBVarNum
	Else
		ReadVariantsNames = 0
	End If
		
End Function

Private Sub clearVariantList		'clear variant list and associated variables
	Erase VariantList$()
	ReDim VariantList$(1)
	VariantList$(0) = "No variants"
	dxVariantsNum = 0
	extraPCBVarNum = 0
	On Error Resume Next
		DlgListBoxArray "variantList", VariantList$()
		Err = 0
	On Error GoTo 0
End Sub

Private Sub checkProcess
	
	frozeGui(True)
	
	Dim prevStat$ 'as string
	prevStat$ = DlgText("status")
	
	DlgText "status", "reading variant file..."
	readVariantAsciiFile
	
	DlgText "status", "getting " & ActiveDocument.Name & " components..."
	getPcbParts
	DlgText "status", "sorting components..."
	sort3
	
	mismatches = 0

	Dim checkCorrect As Integer

	DlgText "status", "checking components..."
	checkCorrect = checkParts()			'check sybchronization
	
	If( DlgText("variantList") = "No variants" And DlgValue( "Check" ) = False )Then		'no variants to check
		DlgText "status", "No variant(s) to check!"
	ElseIf( checkCorrect = 1 )Then															'synchronization check detects no missmatches
		MsgBox "No differences in project's variant(s) were detected", 64,"Synchronization result"
	ElseIf( checkCorrect = 0 )Then															'mismatches detected
		DlgText "status", "synchronization check finished"
		DlgText "status", "number of mismatches = "& mismatches
		Dim msg As String
		msg = "Project's variant(s) are not synchronized!" & vbCrLf  & vbCrLf & " Would you like to see the output file now?"
		If MsgBox( msg, 48 + 4,"Synchronization result") = vbYes Then
			outputFile$ = DlgText( "EdtOutFileName" )
		
			'dialog that displays results of synchronization check
			Begin Dialog UserDialog 820,469,"Result window",.resultWindowFunc
				TextBox 10,7,800,413,.result,1
				PushButton 710,434,90,21,"Close",.CloseResult
				Text 20,448,580,14,"Text1",.Text1
				CheckBox 20,427,20,14,"CheckBox1",.CheckBox1
				Text 50,427,450,14,"show only errors",.Text2
			End Dialog
			
			Dim dlg2 As UserDialog
			On Error Resume Next
				Dim txtFile As String
				txtFile = "Notepad " & outputFile$
				Shell(txtFile,1)
				If Err <> 0 Then
					Dim sysDir As String
					sysDir = Environ("windir")
					If sysdir <> "" Then
						Err = 0
						txtFile = sysDir & "\Notepad.exe " &  outputFile$
						Shell(txtFile,1)
					End If
				End If
				If Err <> 0 Then
					msg = "Notepad.exe was not found!" & vbCrLf
					msg = msg & "Results will be displayed in standard window." ' & vbCrLf
					MsgBox msg, 48,"Application not found"
					Dialog dlg2
				End If
			On Error GoTo 0
			Err = 0
		End If
		DlgText "status", prevStat$
	End If
	
	'clear variables and arrays after check process
	Erase partsNew()
	Erase pcbParts()
	pcbPartsCount = 0
	partsCount = 0
	
	frozeGui(False)
	
End Sub

Private Sub readVariantAsciiFile	'reads components and their variant status from variant ASCII file
	
	Dim fileWithPath$
	fileWithPath$ = dxdProjectDir & dxdVariantFile$
	
	On Error Resume Next	
		If ( FileLen( fileWithPath$ ) < 1 ) Then
			'DlgText "status", "DxD variant ASCII file is empty!"
			Exit Sub
		End If
		
		Err = 0
		Open fileWithPath$ For Input As #1
		If ( Err <> 0 )Then
			Err = 0
			Exit Sub
		End If
	On Error GoTo 0
	
	Dim readLine$
	Dim partName$
	Dim state As Integer
	Dim currentPart As Integer
	Dim pos,pos2,lines As Integer
	Dim currentVariant As Integer
		
	lines = 0
	state = 0
	currentPart = 0
	
	While ( Not EOF(1) )
		Line Input #1, readLine$
		
		If( isEmptyLine ( readLine$ ) )Then
			'do nothing
			
		ElseIf( InStr( readLine$, "ATTRIBUTE VALUES" ) <> 0 )Then
			state = 1
			
		ElseIf( InStr( readLine$, "PART" ) <> 0 And state = 1 )Then
			partName$ = Mid$( readLine$, 6, Len( readLine$ ) )
			ReDim Preserve partsNew( currentPart )'+ 1 )
			partsNew(currentPart).refDes = partName$
			state = 2
		
		ElseIf( InStr( readLine$, "PCB DEFAULT") <> 0 And state = 1 )Then
			ReDim Preserve partsNew( currentPart )'+ 1 )
			partsNew(currentPart).refDes = "default"
			state = 2
		
		ElseIf( InStr( readLine$, "ASSEMBLY_OPTIONS" ) <> 0 And state = 2 )Then
			ReDim partsNew( currentPart ).varChange( dxVariantsNum )
			ReDim partsNew( currentPart ).substituteName( dxVariantsNum )
			currentVariant = 1
			While( currentVariant <= dxVariantsNum )
				pos = InStr( readLine$, "," )
				partsNew( currentPart ).varChange( currentVariant ) = Mid$( readLine$, pos+1, 1 )
				
				If( partsNew( currentPart ).varChange( currentVariant ) = "S" )Then
					pos2 = InStr( readLine$, ";" )
					If( pos2 = 0 )Then
						partsNew( currentPart ).substituteName( currentVariant ) = Mid$( readLine$, pos+3, Len( readLine$ ) )
					Else
						partsNew( currentPart ).substituteName( currentVariant ) = Mid$( readLine$, pos+3, pos2-( pos+3 ) )
					End If
				End If
				
				pos = InStr( readLine$, ";" )
				readLine$ = Mid$( readLine$, pos+1, Len( readLine$ ) )
				currentVariant = currentVariant + 1
			Wend
			state = 1
			currentPart = currentPart + 1		
		ElseIf( InStr( readLine$, "*END*" ) <> 0 )Then
			Exit While
		End If
		lines = lines + 1
		Dim sign As String
		If lines < 8 Then
			sign = "|"
		ElseIf lines < 16 Then
			sign = "||"
		ElseIf lines < 24 Then
			sign = "|||"
		ElseIf lines < 32 Then
			sign = "||||"
		ElseIf lines < 40 Then
			sign = "|||||"
		Else
			sign = "|"
			lines = 1
		End If
		DlgText "status", "Reading components from variant file- " & sign
	Wend
	
	partsCount = currentPart - 1
	
	Close #1
End Sub
Private Sub getPcbParts			'gets parts and their variant status from current design 
	
	Erase pcbParts()
	pcbPartsCount = ActiveDocument.Components.Count
	
	If( pcbPartsCount < 1 )Then
		pcbPartsCount = 0
		Exit Sub
	End If
	
	ReDim pcbParts( pcbPartsCount )
	
	Dim pcbDoc() As Object
	If( dxVariantsNum > 0 )Then
		ReDim pcbDoc(dxVariantsNum)
	Else
		Exit Sub
	End If
	
	Dim tmpStr As String
	Dim variantCount,assemblyCount As Integer
	Dim assemblyVar
	Dim progres As Integer
	progres = 0
		
	variantCount = 1
	Set assemblyVar = ActiveDocument.AssemblyOptions
	
	While( variantCount <= dxVariantsNum )
		assemblyCount = 2
		While( assemblyCount <= assemblyVar.Count )
			tmpStr = assemblyVar.Item(assemblyCount).Name
			tmpStr = Mid$( tmpStr, 1, InStr( tmpStr, ":" ) - 1 )
			If( tmpStr = VariantList$( variantCount ) )Then
				Set pcbDoc( variantCount ) = assemblyVar.Item(assemblyCount)
				Exit While
			End If
			assemblyCount = assemblyCount + 1
		Wend
		If( assemblyCount > assemblyVar.Count )Then
			Set pcbDoc( variantCount ) = Nothing
		End If
		variantCount = variantCount + 1
	Wend
	
	Dim pcbComponent
	Dim state As String
	Dim partCount As Integer
	partCount = 0
	
	While( partCount < pcbPartsCount )
		ReDim pcbParts( partCount ).varChange( dxVariantsNum )
		ReDim pcbParts( partCount ).substituteName( dxVariantsNum )
		variantCount = 0
		While( variantCount <= dxVariantsNum )
			
			If( pcbDoc( variantCount ) Is Nothing )Then
				pcbParts( partCount ).varChange( variantCount ) = "X"
			Else
				Set pcbComponent = pcbDoc( variantCount ).Components.Item( partCount + 1 )
			
				If( Not pcbComponent.Installed )Then
					state = "U"
				ElseIf( pcbComponent.Substituted )Then
					state = "S"
				Else
					state = "I"
				End If
				
				pcbParts( partCount ).substituteName( variantCount ) = pcbComponent.PartType
				pcbParts( partCount ).varChange( variantCount ) = state
				pcbParts( partCount ).refDes = pcbComponent.Name
				
			End If
			
			variantCount = variantCount + 1
			progres = 100*(dxVariantsNum * partCount + variantCount)/(dxVariantsNum*pcbPartsCount)
			DlgText "status", "Reading pcb components - " & progres & "% done."
		Wend
		partCount = partCount + 1
	Wend
	
	Erase pcbDoc()
End Sub

Private Function checkParts As Integer		'checks synchronization 
	checkParts = 1
	
	Dim tmpStr As String
	Dim variantCount As Integer
	Dim assemblyCount As Integer
	Dim partsChecked() As Boolean
	
	variantCount = 1
	
	Dim checkAll As Boolean
	checkAll = DlgValue( "Check" )
	
	If( DlgValue ( "variantList" ) = -1 )Then
		 DlgValue "variantList", 0
	End If
	
	On Error Resume Next
		Err = 0
		Open DlgText( "EdtOutFileName" ) For Output As #1
		If ( Err <> 0 )Then
			Err = 0
			MsgBox "Unable to create/open output file. Try different file's name (directory) or check file's properties.", 48,"Output file error"
			checkParts = -1
			Exit Function
		End If
	On Error GoTo 0
	
	Print #1,"         Result file generated by Dx Variant Checker "
	Print #1," "
	Print #1,"    Date: ";Date
	Print #1,"    Time: ";Time
	Print #1," "
	Print #1,"    PADS Layout pcb file:  ";ActiveDocument.Name
	Print #1,"    DxDesigner ASCII file: ";dxdVariantFile$
	Print #1," "
	Print #1,"    Synchronization check performed for ";
	If( checkAll )Then
		Print #1,"all variants."
	Else
		Print #1,"'"& DlgText( "variantList" );"' variant."
	End If
	
	Print #1," "
	
	If( pcbPartsCount < 1 ) Then
		Print #1," "
		Print #1,"!   No parts to check found in "; ActiveDocument.Name ;"."
		checkParts = 0
		Close #1
		Exit Function
	End If
	
	If( extraPCBVarNum > 0 And checkAll )Then
		Print #1," "
		assemblyCount = 1
		While( assemblyCount <= extraPCBVarNum )
			Print #1, ""
			Print #1,"!   Variant '"; VariantList$( assemblyCount + dxVariantsNum ) ;"' not found in "; dxdVariantFile$ ;" file."
			assemblyCount = assemblyCount + 1
			mismatches = mismatches + 1
			checkParts = 0
		Wend		
	End If
	
	If( checkAll = False And ( DlgValue("variantList") + 1 ) > dxVariantsNum )Then
		Print #1,"!   Variant '"; DlgText("variantList") ;"' not found in "; dxdVariantFile$ ;" file."
		mismatches = mismatches + 1
		checkParts = 0
		Close #1
		Exit Function
	End If
	
	Dim compCount As Integer
	Dim pcbCompCount As Integer
	Dim compInfo As String 's- substituted, u - uninstalled, i - installed
	Dim pcbCompInfo As String 's- substituted, u - uninstalled, i - installed
	Dim pcbCompPartType As String
	Dim progres As Integer
	
	variantCount = 1
	pcbCompInfo = "X"
	CompInfo = "X"
	progres = 0
	
	If( partsCount > 0 )Then
		ReDim partsChecked(partsCount)
		compCount = 0
		While( compCount < partsCount )
			partsChecked( compCount ) = False
			compCount = compCount + 1
		Wend
	End If
	
	While( variantCount <= dxVariantsNum )
		While( checkAll Or ( DlgValue( "variantList" ) + 1 ) = variantCount )
			
			If( pcbParts( 0 ).varChange( variantCount ) = "X" )Then
				mismatches = mismatches + 1
				checkParts = 0
				Print #1, ""
				Print #1, "!   Variant '"; VariantList$( variantCount ); "' not found in "; ActiveDocument.Name ;"."
				Exit While
			End If
			
			Print #1," "
			Print #1," Variant '"; VariantList$( variantCount );"':"
			
			pcbCompCount = 0
			
			While ( pcbCompCount <= pcbPartsCount-1 )
				
				pcbCompInfo = pcbParts(pcbCompCount).varChange( variantCount )
				pcbCompPartType = pcbParts(pcbCompCount).substituteName( variantCount )
				compCount = 0
				
				While ( compCount < partsCount )
					
					If ( pcbParts(pcbCompCount).refDes = partsNew(compCount).refDes )Then
						
						compInfo = partsNew(compCount).varChange( variantCount )
						partsChecked( compCount ) = True
						
						If( pcbCompInfo = compInfo )Then
							If( pcbCompInfo <> "S" )Then
								Print #1,"    "; pcbParts(pcbCompCount).refDes;" : ";retAttrib( CompInfo )
							ElseIf( partsNew(compCount).substituteName( variantCount ) = pcbCompPartType )Then
								Print #1,"    "; pcbParts(pcbCompCount).refDes;" : ";retAttrib( CompInfo )
							Else
								Print #1,"!   "; pcbParts(pcbCompCount).refDes;" Substituted Part Tye mismatch: ";
								Print #1,"(DxD: "; partsNew(compCount).substituteName( variantCount ) ;", Layout: "; pcbCompPartType;")"
								mismatches = mismatches + 1
								checkParts = 0
							End If
						Else
							Print #1,"!   "; pcbParts(pcbCompCount).refDes;" Status mismatch: ";
							Print #1,"(DxD: "; retAttrib( CompInfo ) ;", Layout: "; retAttrib( pcbCompInfo ) ;")"
							mismatches = mismatches + 1
							checkParts = 0
						End If
						
						Exit While
					End If

					compCount = compCount +1
				Wend

				If ( compCount >= partsCount )Then
					
					compInfo = partsNew(compCount).varChange( variantCount )
										
					If( pcbCompInfo = compInfo )Then
						'Print #1,"    Component "; pcbParts(pcbCompCount).refDes;" matched (is ";retAttrib( CompInfo );")"
						If( pcbCompInfo <> "S" )Then
							Print #1,"    "; pcbParts(pcbCompCount).refDes;" : ";retAttrib( CompInfo )
						ElseIf( partsNew(compCount).substituteName( variantCount ) = pcbCompPartType )Then
							Print #1,"    "; pcbParts(pcbCompCount).refDes;" : ";retAttrib( CompInfo )
						Else
							Print #1,"!   "; pcbParts(pcbCompCount).refDes;" Substituted Part Tye mismatch: ";
							Print #1,"(DxD: "; partsNew(compCount).substituteName( variantCount ) ;", Layout: "; pcbCompPartType;")"
							mismatches = mismatches + 1
							checkParts = 0
						End If
					Else
						Print #1,"!   "; pcbParts(pcbCompCount).refDes;" Status mismatch: ";
						Print #1,"(DxD: "; retAttrib( CompInfo ) ;", Layout: "; retAttrib( pcbCompInfo ) ;")"
						mismatches = mismatches + 1
						checkParts = 0
					End If
				End If
				
				pcbCompInfo = "X"
				CompInfo = "X"
				pcbCompCount = pcbCompCount + 1
			
				progres = 100*((VariantCount-1) * pcbPartsCount + pcbCompCount)/(dxVariantsNum*pcbPartsCount)
				DlgText "status", "Checking variants - " & progres & "% done."
				
			Wend
			
			Exit While
		Wend
		variantCount = variantCount + 1
	Wend
	
	compCount = 0
	While( compCount < partsCount )
		If( partsChecked( compCount ) = False )Then
			Print #1," "
			Print #1,"!   Component "; partsNew(compCount).refDes;" not found in " & ActiveDocument.Name & "."
			mismatches = mismatches + 1
			checkParts = 0
		End If
		compCount = compCount + 1
	Wend
	
	Erase partsChecked()
	
	Close #1
	
End Function

Private Function isEmptyLine( lline$ ) As Boolean		'returns true when lline$ is empty line (only white signs occurs)
	If( Len( lline$ ) = 0 )Then
		isEmptyLine = True
		Exit Function
	End If
	
	While( Len( lline$ ) > 0)
		If( Left$( lline$, 1 ) =" " Or Left( lline$, 1 ) ="\t" )Then
			lline$ = Right$( lline$, Len( lline$ ) - 1 )
		Else
			isEmptyLine = False
			Exit Function
		End If
	Wend
	isEmptyLine = True
End Function

Private Function retAttrib( lline$ ) As String	'returns full name of variant status
	Select Case UCase( lline$ )
	Case "I"
		retAttrib = "installed"
	Case "U"
		retAttrib = "uninstalled"
	Case "S"
		retAttrib = "substituted"
	Case Else
		retAttrib = "unknown"
	End Select
End Function

Dim tmpCloseResultWindow As Boolean

Private Function resultWindowFunc(DlgItem$, Action%, Supp2Value&) As Boolean	'function that cooperates with result dialog
	
	tmpCloseResultWindow = True
	
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText "Text1", "You may also see the results later in output file."
		tmpCloseResultWindow = fillResultWindow( True )
		DlgValue "Checkbox1", True
		
	Case 2 ' Value changing or button pressed
		If DlgItem$ = "CloseResult" Then
			resultWindowFunc = False
		ElseIf( Supp2Value& = 0 ) Then
			resultWindowFunc = True
			fillResultWindow( False )
		Else 
			resultWindowFunc = True
			fillResultWindow( True )
		End If
		Rem resultWindowFunc = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		If( tmpCloseResultWindow = False )Then
			resultWindowFunc = False
		End If
	Case 6 ' Function key
	End Select
End Function

Function fillResultWindow( errors As Boolean ) As Boolean		'puts synchronization results from file to result dialog
	fillResultWindow = True
	Err = 0
	On Error Resume Next
	Open outputFile$ For Input As #1
	If ( Err <> 0 )Then
		Err = 0
		MsgBox "Unable to open output file to show results",48,"Output file error"
		fillResultWindow = False
		Exit Function
	End If
	On Error GoTo 0
	Dim readLine As String
	Dim tmpStr As String
	Dim whole As String 
	Dim lineCount As Integer
	Dim checkline As String
	
	lineCount = 1
	While ( Not EOF(1) )
		Line Input #1, readLine
		tmpStr = readLine
		If( errors = False )Then
			whole = whole & vbCrLf & readLine
		ElseIf( lineCount < 11 ) Then
			whole = whole & vbCrLf & tmpStr
		ElseIf( isEmptyLine( readLine ) )Then
			'do nothing
		ElseIf( InStr( readLine, "!   Variant" ) = 1 )Then
			whole = whole & vbCrLf & vbCrLf & tmpStr
		ElseIf ( InStr( readLine, "Variant '") <> 0 )Then
			checkLine = tmpStr
		ElseIf( InStr( readLine, "!" ) = 1 )Then
			If( checkLine <> "" )Then
				whole = whole & vbCrLf & vbCrLf & checkLine
			End If
			whole = whole & vbCrLf & tmpStr
			checkLine = ""
		End If
		lineCount = lineCount + 1
	Wend
	DlgText "result", whole
	
	Close #1
	
End Function

Private Sub sort3 ' fast quick sort
	
	If( pcbPartsCount < 2 )Then
		Exit Sub
	End If
	
	quickSort( 0, pcbPartsCount-1 )
	insertionSort( 0, pcbPartsCount-1 )
	
End Sub

Private Function quickSort(l As Integer, r As Integer)
	Dim M As Integer
	Dim i As Integer
	Dim j As Integer
	Dim v As partNew
	M = 4
	
	If( (r-l) > M )Then
		i = (r+l) / 2
		
		If( strnumcmp( pcbParts( l ).refDes, pcbParts( i ).refDes ) > 0 )Then
			swap( l, i )
		End If
		If( strnumcmp( pcbParts( l ).refDes, pcbParts( r ).refDes ) > 0 )Then
			swap( l, r )
		End If
		If( strnumcmp( pcbParts( i ).refDes, pcbParts( r ).refDes ) > 0 )Then
			swap( i, r )
		End If
		
		j = r - 1
		swap( i, j )
		i = l
		v = pcbParts( j )
		While(True)
			Do
				i = i + 1
			Loop While( strnumcmp( pcbParts( i ).refDes, v.refDes ) < 0 )
			Do
				j = j - 1
			Loop While( strnumcmp( pcbParts( j ).refDes, v.refDes ) > 0 )
			
			If( j < i )Then
				Exit While
			End If
			
			swap( i, j )
		Wend
		
		swap( i, r-1 )
		quickSort( l, j )
		quickSort( i+1, r )
		
	End If
End Function

Private Function swap (a As Integer,b As Integer)	'swap two parts in array
	Dim tmp As partNew
	tmp = pcbParts( a )
	pcbParts( a ) = pcbParts( b )
	pcbParts( b ) = tmp
End Function

Private Function insertionSort(lo0 As Integer, hi0 As Integer)
	Dim i,j As Integer
	Dim v As partNew
	
	i = lo0 + 1
	While( i <= hi0 )
		v = pcbParts( i )
		j = i
		While( (j>lo0) And strnumcmp( pcbParts( j-1 ).refDes, v.refDes ) > 0 )
			pcbParts( j ) = pcbParts( j-1 )
			j = j - 1
		Wend
		pcbParts( j ) = v
		i = i + 1
	Wend
	
End Function

Private Function strnumcmp(string1 As String, string2 As String) As Integer	'compare two reference designators
	'examples
	'string1 = "U1", string2 = "U2", return value > 0
	'string1 = "U2", string2 = "U10", return value < 0
	'string1 = "R11", string2 = "R11", return value = 0	
	'string1 = "u01", string2 = "U1", return value = 0
	
	string1 = UCase( string1 )
	string2 = UCase( string2 )
	
	If( string1 = string2 )Then
		strnumcmp = 0
		Exit Function
	End If
	
	Dim c1,c2,sign As Integer
	
	c1 = Asc( Mid$( string1, 1, 1 ) )
	c2 = Asc( Mid$( string2, 1, 1 ) )
	
	sign = Asc("$")
	
	If( c1 = sign And c2 <> sign )Then
		strnumcmp = 1
		Exit Function
	ElseIf( c1 <> sign And c2 = sign )Then
		strnumcmp = -1
		Exit Function
	End If
	
	Dim count As Integer
	count = 1
	
	On Error Resume Next
	While( count <= Len (string1) And count <= Len (string2) )
		
		If( c1 <> c2 )Then
			Exit While
		End If
		
		Err = 0
		c1 = CInt( Mid$( string1, count, Len(string1) ) )
		c2 = CInt( Mid$( string2, count, Len(string2) ) )	
		
		If( Err = 0 )Then
			Exit While
		Else
			c1 = Asc( Mid$( string1, count, 1 ) )
			c2 = Asc( Mid$( string2, count, 1 ) )
		End If
		count = count + 1
 	Wend
 	Err = 0
 	On Error GoTo 0
	
	strnumcmp = c1 - c2
	
End Function

Private Sub frozeGui(ByVal state As Boolean)
		DlgEnable "EdtASCIIName", Not state
		DlgEnable "EdtOutFileName", Not state
		DlgEnable "BtnClose", Not state
		DlgEnable "BtnCheck", Not state
		DlgEnable "BtnOpenAscii", Not state
		DlgEnable "BtnOpenOutput", Not state
		DlgEnable "Check", Not state
		If state = False Then
			DlgEnable "variantList", DlgValue ("Check")
		Else
			DlgEnable "variantList", False
		End If
		
End Sub

Private Sub moveCursorInASCIIToEnd
	'move cursor to the end in the ASCII file name, so that name is visible
	On Error Resume Next
		DlgFocus "EdtASCIIName"
		SendKeys "{End}"
		Err=0
	On Error GoTo 0
End Sub
