' PowerBGA Automation Utility 1.BAS
' 
' This Sax Basic Script enables the user to create two different reports with two different formats:
'
' BGA Pin to Die Function Name (with or without Die Pin Numbers)
' Die Function Name to Pin Report (with or without Die Pin Numbers)
'


Dim FileList$(1000)
Dim ListSameFunction$(50)
Dim value As Integer

Sub Main

	On Error GoTo ByeBye

	Begin Dialog UserDialog 380,224,"PADS-PowerBGA - Automation Utilities",.CallbackFunc
		OKButton 40,168,110,21,.OK
		GroupBox 20,7,340,77,"Utility",.GroupBox1 'Utility Group
		OptionGroup .options
			OptionButton 40,21,290,14,"&BGA Pin to Die Function Name Report",.ButtonBGA
			OptionButton 40,42,270,14,"&Die Function Name to BGA Pin Report",.ButtonDie
		GroupBox 20,98,340,63,"Result",.GroupBox2
		CheckBox 80,63,210,14,"Include Die Pin Numbers",.CheckBox1
		Text 40,112,310,42,"",.Text1
		Text 126,196,140,14,"Mentor Graphics Corp.",.Text2
		Text 74,210,250,14,"Your Interconnect Technology Partner",.Text3
		CancelButton 210,168,110,21
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

	' Display a message and let the user know that you are off and running...
	MsgBox "Depending on the complexity of your BGA design, this may take a moment or two... "

	index = 0
	
	' Lock server to speed up process
	LockServer
	
	' Open temporary text file for each component
	Randomize
	dummyfile = DefaultFilePath & "\tmp"  & CInt(Rnd()*10000) & ".txt"
	Open dummyfile For Output As #1



	'  Go through each pin in the design and output values to a temporary file for
	'  the component to which that pin belongs.  This will enable you to compare
	'  pin mappings on a component-by-component basis.
	
	Dim nextPin As Object
	For Each nextPin In ActiveDocument.GetObjects(3, "", False)


		' Attempt to speed up the processing by skipping this first section if the IPowerPCB::Component 
		' automation method returns the same Name as the previous component.  In other words, don't 
		' assume that you will always be handed all the pins for a particular component before moving to
		' the next component.
		a = nextPin.Component.Name
		'Debug.Print a
		'Debug.Print currentcomponent
		'Debug.Print "-----------"
		If (a <> currentcomponent) Then
			currentcomponent = a
			If ((FileAttr(1,1) = 2) Or ( FileAttr(1,1) = 8)) Then Close #1   ' If a file is already open for output or append status, close it.
			filename = DefaultFilePath & "\tmp"  & a & ".txt"
			filefound = False
			inc = 0
			While (inc < index)
				If (filename = FileList$(inc)) Then ' If the file already exists, open it for appending
					Open filename For Append As #1
					filefound = True
					inc = index ' Dump out of the loop a little quicker...
				End If
				inc = inc + 1
			Wend
			If (Not filefound) Then ' If the file wasn't found, open it for output
				Open filename For Output As #1
				FileList$(index) = filename
				'Debug.Print FileList$(index) 
				index = index + 1
			End If
		End If
	       
		' Output Pin Name Property
		a = nextPin.Name
		Print #1, a; Space$(30-Len(a)); 

		' Output Pin Owning FunctionName Property
		a = nextPin.FunctionName
		Print #1, a; Space$(30-Len(a)); 
				
		' Output Pin Attached Net Property
		Set nnet = nextPin.Net
		If (nnet Is Nothing) Then a = "N/A" Else a = nnet.Name 
		Print #1, a; Space$(30-Len(a))
		
	
	Next nextPin	

	' Unlock the server
	UnlockServer

	' Close the last file left open
	Close #1

	On Error GoTo fileinuse
	'  Create the file that will be used to store the cross reference information.
	filename = DefaultFilePath & "\Connection Report.txt"
	Open filename For Output As #1

	'Debugfilename = DefaultFilePath & "\dbg"  & ".txt"
	'Open Debugfilename For Output As #4
	
	outerindex = 0
	innerindex = 0
	'Debug.Print index
	'Debug.Print dlg.CheckBox1
	'Debug.Print value
	If (value = 0 ) Then
		If (dlg.CheckBox1) Then
			Print #1 ,"BGA Pin	Die Function (Pin #)"
		Else
			Print #1 ,"BGA Pin	Die Function"
		End If
		While (outerindex < index)
			'Debug.Print FileList$(outerindex)
			Open FileList$(outerindex) For Input As #2
			While Not EOF (2)
				Line Input #2, inputline$
				OuterName = Left$(inputline,30)
				'Print #4,  "OuterName =";OuterName; "    ";
				OuterFunctionName = Trim$(Mid$(inputline,31,30))
				'Print #4,  "OuterFunctionName= ";OuterFunctionName; "    ";
				OuterPinNetName = Trim$(Right$(inputline,30))
				'Print #4,  "OuterPinNetName =";OuterPinNetName
				If (Len(OuterFunctionName) = 0) Then
					Print #1,  Trim$(Right$(OuterName,(30 - InStr(OuterName,"."))));"	";
					If (StrComp(OuterPinNetName ,"N/A") = 0) Then 
						Print#1, "Not Connected"
						innerindex = index ' skip the next loop
						'Print #4,  "****OuterPinNetName =N/A"
					End If
					first = True
				      	While (innerindex < index)
						Open FileList$(innerindex) For Input As #3
						While Not EOF (3)
							Line Input #3, inputline$
							InnerName = Left$(inputline,30)
							'Print #4,  "InnerName =";InnerName; "    ";
							InnerFunctionName = Trim$(Mid$(inputline,31,30))
							'Print #4,  "InnerFunctionName= ";InnerFunctionName; "    ";
							InnerPinNetName = Trim$(Right$(inputline,30))
							'Print #4,  "InnerPinNetName =";InnerPinNetName
							If (Len(InnerFunctionName) <> 0) Then
								If (OuterPinNetName = InnerPinNetName) Then 
									'Print #4,  "*****InnerPinNetName =";InnerPinNetName;"   OuterPinNetName =";OuterPinNetName
									If first Then
										If (dlg.CheckBox1) Then
											Print #1, Trim$(InnerFunctionName);"(";Trim$(Right$(InnerName,(30 - InStr(InnerName,"."))));")";
										Else
											Print #1, Trim$(InnerFunctionName);
										End If
										first = False
									Else
										If (dlg.CheckBox1) Then
											Print #1, ",";Trim$(InnerFunctionName);"(";Trim$(Right$(InnerName,(30 - InStr(InnerName,"."))));")";
										Else
											Print #1, ",";Trim$(InnerFunctionName);
										End If
									End If
								End If
							End If
						Wend
						Close #3
						innerindex = innerindex + 1
					Wend
					Print #1
					innerindex = 0
				End If
			Wend
			Close #2
			outerindex = outerindex + 1
		Wend
	Else
		If (dlg.CheckBox1) Then
			Print #1 ,"Die Function (Pin #)	BGA Pin(s)"
		Else
			Print #1 ,"Die Function	BGA Pin(s)"
		End If
		While (outerindex < index)
			'Debug.Print FileList$(outerindex)
			Open FileList$(outerindex) For Input As #2
			While Not EOF (2)
				Line Input #2, inputline$
				OuterName = Left$(inputline,30)
				'Print #4,  "OuterName =";OuterName; "    ";
				OuterFunctionName = Trim$(Mid$(inputline,31,30))
				'Print #4,  "OuterFunctionName= ";OuterFunctionName; "    ";
				OuterPinNetName = Trim$(Right$(inputline,30))
				'Print #4,  "OuterPinNetName =";OuterPinNetName
				If (Len(OuterFunctionName) <> 0) Then
					If (dlg.CheckBox1) Then
						Print #1,  Trim$(OuterFunctionName);" (";Trim$(Right$(OuterName,(30 - InStr(OuterName,"."))));")";"	";
					Else
						Print #1,  Trim$(OuterFunctionName);"	";
					End If
					If (StrComp(OuterPinNetName ,"N/A") = 0) Then 
						Print#1, "Not Connected"
						innerindex = index ' skip the next loop
						'Print #4,  "****OuterPinNetName =N/A"
					End If
					first = True
				      	While (innerindex < index)
						Open FileList$(innerindex) For Input As #3
						While Not EOF (3)
							Line Input #3, inputline$
							InnerName = Left$(inputline,30)
							'Print #4,  "InnerName =";InnerName; "    ";
							InnerFunctionName = Trim$(Mid$(inputline,31,30))
							'Print #4,  "InnerFunctionName= ";InnerFunctionName; "    ";
							InnerPinNetName = Trim$(Right$(inputline,30))
							'Print #4,  "InnerPinNetName =";InnerPinNetName
							If (Len(InnerFunctionName) = 0) Then
								If (OuterPinNetName = InnerPinNetName) Then 
									If first Then
										Print #1,  Trim$(Right$(InnerName,(30 - InStr(InnerName,"."))));
										first = False
									Else
										Print #1,  ",";Trim$(Right$(InnerName,(30 - InStr(InnerName,"."))));
									End If
								End If
							End If
						Wend
						Close #3
						innerindex = innerindex + 1
					Wend
					Print #1
					innerindex = 0
				End If
			Wend
			Close #2
			outerindex = outerindex + 1
		Wend
	End If
	
	
	Close #1
	'Close #4
	filename = DefaultFilePath & "\Connection Report.txt"
	

	' Do some house cleaning...  Get rid of the evidence before somebody finds it...
	Kill dummyfile	
	inc = 0
	While (inc < index)
		Kill FileList$(inc)
		inc = inc + 1
	Wend

	' Start Excel and load the text file
	
	On Error GoTo noExcel
	Dim excelApp As Object
	Set excelApp = CreateObject("Excel.Application")
	On Error GoTo 0
	excelApp.Visible = True
	excelApp.Workbooks.OpenText 	FileName:= filename
	excelApp.Rows("1:1").Select
	With excelApp.Selection
		.Font.Bold = True
		.Font.Italic = True
		.RowHeight = 18
	End With
	excelApp.Columns("A:A").Select
	With excelApp.Selection
		.ColumnWidth = 25 ' Make the first column a little wider
	End With
	excelApp.Columns("B:B").Select
	With excelApp.Selection
		.ColumnWidth = 90 ' Make the second column a little wider
	End With
	excelApp.Range("A2:B999").Select  ' Select everything except the header row
	excelApp.Selection.Sort(excelApp.Range("A1"), 1 ) ' Call the sort function from Excel
	excelApp.Range("A1").Select
	Set excelApp = Nothing
	End
	
noExcel:
	' Display the text file
	Shell "Notepad " & filename, 3

fileinuse:
	MsgBox "Unable to create '" & filename & "'"
	Exit All
	ByeBye:
		Exit All
End Sub


Rem See DialogFunc help topic for more information.
Private Function CallbackFunc(DlgItem$, Action%, SuppValue%) As Boolean
	Select Case Action%
		Case 1 ' Dialog box initialization
			DlgValue "options",0
			result = "Generates a MS Excel Spreadsheet that lists the BGA Pin Number (solder ball) along with the Die Function Name(s) associated with that pin."
			DlgText "Text1", result
			DlgValue "CheckBox1",1 ' The default is to include die pin numbers
		Case 2 ' Value changing or button pressed
			If ((DlgItem$ <> "OK") And (DlgItem$ <> "Cancel")) Then
				CallbackFunc = True ' Prevent button press from closing the dialog box
				If (DlgItem$ = "options") Then
					Select Case SuppValue%
						Case 0
							result = "Generates a MS Excel Spreadsheet that lists the BGA Pin Number (solder ball) along with the Die Function Name(s) associated with that pin."
							DlgText "Text1", result
							value = 0
						Case 1
							result = "Generates a MS Excel Spreadsheet that lists the Die Function Name and the BGA Pin(s) {solder ball} associated with that function."
							DlgText "Text1", result
							value = 1
					End Select
				End If
			 ElseIf DlgItem$ = "Cancel" Then
			 	CallbackFunc = False ' closing the dialog box
			End If
		Case 3 ' TextBox or ComboBox text changed
		Case 4 ' Focus changed
		Case 5 ' Idle
			Rem CallbackFunc = True ' Continue getting idle actions
	End Select
End Function
