'V1 for PADS Logic released 10/17/06 as an unsupported script
'Based on V2 script for PADS Layout
'For PADS Logic 2007 or future releases only - this will not work on PADS Logic 2005 Spac3 or lower.
'ECO Registration Finder.BAS originally created 4/12/99 by Tom Woundy, Manager, Technical Services, PADS Software, Inc.
'Modified by PADS User, Jon DeGenova, to present part types in the display rather than by reference designator and replaced
'the pan disable/enable with a turn off confirmation control.  This revision was not shipped with releases V3.5.1 or earlier
'Users can modify and use this script as they see fit, PADS Software, Inc. does not support this script, use at you own risk
'This script is designed to find what parts are currently registered for ECO comarisions and which are not
'It will allow users to probe to PowerPCB to locate these parts as well as change the registration status in the database
'A report feature is also included that will create a text file showing the registration status as well.

' ******************** Important Facts - Please read carefully ******************************
'Be aware that ECO Registration is a Part Type level change, all parts of a given type will assume the same registration status
'Also be aware that there is currently no other way to modify the registation status other than updating the part type in the library
'and then using Change Part command to effect a change.
'Knowing that, be aware that using this script WILL put parts in this database OUT OF SYNC with your library
'If you save parts from your design, back into the library, after running this script, the library WILL assume the ECO Registration status
'that was last modofied on your design.

Dim ListRegisteredParts$() ' this is used to manage registered part types
Dim ListNonRegisteredParts$() ' this is used to manage non-registered part types

Sub Main

	'get list of parts and registration status
	GetList

	'define dialog to post results to
	Begin Dialog UserDialog 430,343,"ECO Part Registration Finder",.dialog_function ' %GRID:10,7,1,1
		GroupBox 20,7,390,231,"Part Types",.GroupBox1
		ListBox 30,35,180,196,ListRegisteredParts(),.RegPartList
		ListBox 220,35,180,196,ListNonRegisteredParts(),.NonRegPartList
		OKButton 60,308,110,28
		Text 30,21,90,14,"Registered",.Text1
		Text 220,21,170,14,"Non-Registered",.Text2
		PushButton 250,308,110,28,"Report",.ReportButton
		PushButton 70,245,100,21,"Unregister >>",.UnregisterButton
		PushButton 260,245,100,21,"<< Register",.RegisterButton
		CheckBox 30,280,190,14,"Turn off confirmation",.Confirm_CheckBox
	End Dialog
	
	'declare dialog
	Dim dlg As UserDialog
	
	'invoke dialog
	Dialog dlg
	
	'clear all selections
	ActiveDocument.SelectObjects(9999,,False)
	
End Sub

Rem See DialogFunc help topic for more information.
Private Function dialog_function(DlgItem$, Action%, SuppValue&) As Boolean
Dim subject$

	Select Case Action%
	Case 1 ' Dialog box initialization
		'clear selections in Logic so crossprobing starts fresh
		ActiveDocument.SelectObjects(9999,,False)
		
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
		Case "ReportButton"
			GetList
			MakeReport
			dialog_function = True
			
		Case "UnregisterButton"
			'get selected part type from Registered list
			state = DlgValue ("RegPartList")
			If state >= 0 And state <= UBound(ListRegisteredParts) - 1 Then
				subject = ListRegisteredParts(state)
				'warn user about impending registration change
				If DlgValue ("Confirm_CheckBox") = 0 Then
					answer = MsgBox ("All parts on this of the Part Type " & subject & _
									 " will be ignored by ECO, Continue?", 52, "PADS Logic" )
				Else
					answer = vbYes
				End If
			Else
				answer = vbNo
			End If
			If answer = vbYes Then
				'remove eco registration
				For Each NextPartType In ActiveDocument.GetObjects(plogObjectTypePartType,,False)
					If NextPartType = subject Then
						NextPartType.ECORegistered = False
					End If
				Next NextPartType				
				'update script with changes made in database
				GetList
				'update dialog with updated list
				DlgListBoxArray ("RegPartList", ListRegisteredParts())
				DlgListBoxArray ("NonRegPartList", ListNonRegisteredParts())
			End If
			'allow dialog to remain on screen
			dialog_function = True
			
		Case "RegisterButton"
			'get selected part type from Non-registered list
			state = DlgValue ("NonRegPartList")
			If state >= 0 And state <= UBound(ListNonRegisteredParts) - 1 Then
				subject = ListNonRegisteredParts(state)
				'warn user about impending registration change
				If DlgValue ("Confirm_CheckBox") = 0 Then
					answer = MsgBox ("All parts on this of the Part Type " & subject & _
									 " will be included by ECO, Continue?", 52, "PADS Logic")
				Else
					answer = vbYes
				End If
			Else
				answer = vbNo
			End If
			If answer = vbYes Then
				'add eco registration
				For Each NextPartType In ActiveDocument.GetObjects(plogObjectTypePartType,,False)
					If NextPartType = subject Then
						NextPartType.ECORegistered = True
					End If
				Next NextPartType
				'update script with changes made in database
				GetList
				'update dialog with updated list
				DlgListBoxArray ("RegPartList", ListRegisteredParts())
				DlgListBoxArray ("NonRegPartList", ListNonRegisteredParts())
			End If
			'allow dialog to remain on screen
			dialog_function = True
			
		Case "RegPartList" 'crossprobe to PowerPCB
			'clear all selections
			ActiveDocument.SelectObjects(9999,,False)
			state = DlgValue ("RegPartList")
			If state >= 0 And state <= UBound(ListRegisteredParts) - 1 Then
				subject = ListRegisteredParts(state)
				CrossProbeToApplication (subject)
			End If
				
		Case "NonRegPartList" 'crossprobe to PowerPCB
			'clear all selections
			ActiveDocument.SelectObjects(9999,,False)
			state = DlgValue ("NonRegPartList")
			If state >= 0 And state <= UBound(ListNonRegisteredParts) - 1 Then
				subject = ListNonRegisteredParts(state)
				CrossProbeToApplication (subject)
			End If
					
		End Select
	
	Case 4 ' Focus changed
		
		'grey out remove button if focus is on registered part types
		If DlgFocus = "RegPartList" Then
			DlgEnable "UnregisterButton" ,True
			DlgEnable "RegisterButton" ,False
		End If
		
		'grey out add button if focus is on non-registered part types
		If DlgFocus = "NonRegPartList" Then
			DlgEnable "UnregisterButton" ,False
			DlgEnable "RegisterButton" ,True
		End If
	
	End Select
End Function

Sub MakeReport()
	'this routine will create a text file that will pop into notepad with results seen in dialog
	filename = DefaultFilePath & "\ECO Registration Report.txt"
	Open filename For Output As #1
		
	'make header for report
	Print #1, ActiveDocument.Name & " " & Date & " " & Time
	Print #1, ""
		
	'print list of registered part types in file
	Print #1, "***** ECO Registered Part Types ******"
	For print_loop = 0 To UBound(ListRegisteredParts) - 1
		Print #1, ListRegisteredParts(print_loop)
	Next print_loop
	Print #1, ""

	'print list of non-registered part types in file
	Print #1, "***** Non-ECO Registered Part Types ******"
	For print_loop = 0 To UBound(ListNonRegisteredParts) - 1
		Print #1, ListNonRegisteredParts(print_loop)
	Next print_loop

	'close Text file
	Close #1

	'spawn notepad with file just created
	Shell "notepad " & """" & filename & """", 1
End Sub

Sub GetList ()
'get list of parts that are registered and non-registered and put into global arrays

Dim NextPartType As Object
Dim RegIndex%
Dim UnRegIndex%

'since these were declared as dynamic arrays, they must be 
'sized before being used - 5000 indicates that design has no more than 5000 parts registered 
'or non-registered if design is larger, user will need to increase these values
ReDim ListRegisteredParts$(5000)
ReDim ListNonRegisteredParts$(5000)

'create lists of ECO Registered and Non Registerd Part Types
RegIndex = 0
UnRegIndex = 0
For Each NextPartType In ActiveDocument.GetObjects(plogObjectTypePartType,,False)
	If NextPartType.ECORegistered Then
		ListRegisteredParts$(RegIndex) = NextPartType.Name
		RegIndex=RegIndex+1
	Else
		ListNonRegisteredParts$(UnRegIndex) = NextPartType.Name
		UnRegIndex = UnRegIndex+1
	End If
Next NextPartType

'Can't have arrays with zero length, so add empty name on end of each list
ListRegisteredParts$(RegIndex) = ""
ListNonRegisteredParts$(UnRegIndex) = ""

'redefine arrays, but preserve the values in them, so that they are now 
'the size they really need to be
ReDim Preserve ListRegisteredParts$(RegIndex)
ReDim Preserve ListNonRegisteredParts$(UnRegIndex)

End Sub

Private Function CrossProbeToApplication (subject) As Boolean

	'select part type user chose in dialog
	ActiveDocument.SelectObjects(plogObjectTypePartType,subject,True)

End Function
