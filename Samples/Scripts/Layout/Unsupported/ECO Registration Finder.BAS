'V2 Released 9/21/00 to the PADS OLE Exchange Web Site
'For PowerPCB V3.0 Or future only - this will Not work On PowerPCB V2.1 Or lower.
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
'and then using ECO/Change Part command (in PowerPCB) to effect a change.
'Knowing that, be aware that using this script WILL put parts in this database OUT OF SYNC with your library
'If you save parts from your design, back into the library, after running this script, the library WILL assume the ECO Registration status
'that was last modofied on your design.

Dim ListRegisteredParts$() ' this is used to manage ref-des only - internally
Dim ListNonRegisteredParts$() ' this is used to manage ref-des only - internally
Dim DialogListRegisteredParts$() ' this is used to manage ref-des and Part type to outside world
Dim DialogListNonRegisteredParts$() ' this is used to manage ref-des and Part type to outside world

Sub Main

	'get list of parts and registration status
	GetList

	'define dialog to post results to
	Begin Dialog UserDialog 430,343,"ECO Part Registration Finder",.dialog_function ' %GRID:10,7,1,1
		GroupBox 20,7,390,231,"Part Types",.GroupBox1
		ListBox 30,35,180,196,DialogListRegisteredParts(),.RegPartList
		ListBox 220,35,180,196,DialogListNonRegisteredParts(),.NonRegPartList
		OKButton 60,308,110,28
		Text 30,21,90,14,"Registered",.Text1
		Text 220,21,170,14,"Non-Registered",.Text2
		PushButton 250,308,110,28,"Report",.ReportButton
		PushButton 70,245,90,21,"Add >>",.AddButton
		PushButton 260,245,90,21,"<< Remove",.RemoveButton
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
		'clear selections in PPCB so crossprobing starts fresh
		ActiveDocument.SelectObjects(9999,,False)
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
		Case "ReportButton"
				GetList
				MakeReport
				dialog_function = True
		Case "AddButton"
			 ' here we cross link the array of "ref-des only" array to what is is in the dialog as "ref des - part type"
				subject = ListRegisteredParts(DlgValue ("RegPartList"))
								
				'warn user about impending registration change
				If DlgValue ("Confirm_CheckBox") = 0 Then
				  answer = MsgBox ("All parts on this of the Part Type " & subject & _
				  " will be ignored by ECO, Continue?", 52, "PowerPCB" )
				Else
				  answer = vbYes
				End If  
				If answer = vbYes Then
					'remove eco registration
					'ActiveDocument.Components.subject.ECORegistered = False
	           For Each NextPart In ActiveDocument.GetObjects(ppcbObjectTypeComponent,,False)
    	           If NextPart.PartType = subject Then
    	             NextPart.PartTypeECORegistered = False
               	   End If  
	           Next NextPart				
					'update script with changes made in database
					GetList
					
					'update dialog with updated list
					DlgListBoxArray ("RegPartList", DialogListRegisteredParts())
					DlgListBoxArray ("NonRegPartList", DialogListNonRegisteredParts())
				End If
				
				'allow dialog to remain on screen
				dialog_function = True
		Case "RemoveButton"
			' here we cross link the array of "ref-des only" array to what is is in the dialog as "ref des - part type"
				subject = ListNonRegisteredParts(DlgValue ("NonRegPartList"))
				
				'warn user about impending registration change
				If DlgValue ("Confirm_CheckBox") = 0 Then
				  answer = MsgBox ("All parts on this of the Part Type " & subject & _
				  " will be included by ECO, Continue?", 52, "PowerPCB")
				Else
				  answer = vbYes
				End If   
				If answer = vbYes Then
					'add eco registration
					'ActiveDocument.Components(subject).PartTypeECORegistered = True
				For Each NextPart In ActiveDocument.GetObjects(ppcbObjectTypeComponent,,False)
    	           If NextPart.PartType = subject Then
    	             NextPart.PartTypeECORegistered = True
               	   End If  
	           Next NextPart
				
					
					'update script with changes made in database
					GetList
					
					'update dialog with updated list
					DlgListBoxArray ("RegPartList", DialogListRegisteredParts())
					DlgListBoxArray ("NonRegPartList", DialogListNonRegisteredParts())
				End If
				
				'allow dialog to remain on screen
				dialog_function = True
			Case "RegPartList" 'crossprobe to PowerPCB
				'clear all selections
				ActiveDocument.SelectObjects(9999,,False)
				
				subject = ListRegisteredParts(DlgValue ("RegPartList"))
				
				'pan = DlgValue ("Pan_CheckBox")
				pan = 0
				
				CrossProbeToApplication (subject,Pan)

				
			Case "NonRegPartList" 'crossprobe to PowerPCB
				'clear all selections
				ActiveDocument.SelectObjects(9999,,False)
				
				subject = ListNonRegisteredParts(DlgValue ("NonRegPartList"))
				
				'pan = DlgValue ("Pan_CheckBox")
				pan = 0
				
				CrossProbeToApplication (subject,Pan)
					
		End Select
	
	Case 4 ' Focus changed
		
		'grey out remove button if focus is on registered part types
		If DlgFocus = "RegPartList" Then
			DlgEnable "AddButton" ,True
			DlgEnable "RemoveButton" ,False
		End If
		
		'grey out add button if focus is on non-registered part types
		If DlgFocus = "NonRegPartList" Then
			DlgEnable "AddButton" ,False
			DlgEnable "RemoveButton" ,True
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
		For print_loop = 0 To (UBound(ListRegisteredParts) -1)
			Print #1, DialogListRegisteredParts(print_loop)
		Next print_loop
		Print #1, ""

		'print list of non-registered part types in file
		Print #1, "***** Non-ECO Registered Part Types ******"
		For print_loop = 0 To (UBound (ListNonRegisteredParts) - 1)
			Print #1, DialogListNonRegisteredParts(print_loop)
		Next print_loop

		'close Text file
		Close #1

		'spawn notepad with file just created
 		Shell "notepad " & """" & filename & """", 1
End Sub

Sub GetList ()
'get list of parts that are registered and non-registered and put into global arrays

Dim NextPart As Object
Dim RegIndex%
Dim UnRegIndex%
Dim I As Integer  'Add a counter
Dim ShouldDo As Integer 'Add a logical yes/no

'since these were declared as dynamic arrays, they must be 
'sized before being used - 5000 indicates that design has no more than 5000 parts registered 
'or non-registered if design is larger, user will need to increase these values
ReDim ListRegisteredParts$(5000)
ReDim ListNonRegisteredParts$(5000)
ReDim DialogListRegisteredParts$(5000)
ReDim DialogListNonRegisteredParts$(5000)

'create lists of ECO Registered and Non Registerd Part Types
For Each NextPart In ActiveDocument.GetObjects(ppcbObjectTypeComponent,,False)
	If NextPart.PartTypeECORegistered Then
		ShouldDo = 1
		For I = 0 To RegIndex
		  If NextPart.PartType = ListRegisteredParts$(I) Then
             ShouldDo = 0
             Exit For
          End If
        Next I
        If ShouldDo = 1 Then
		  ListRegisteredParts$(RegIndex) = NextPart.PartType 'for internal use
		  DialogListRegisteredParts$(RegIndex) = (NextPart.PartType) ' for external use
		  RegIndex=RegIndex+1
		End If
	Else
	  
	   ShouldDo = 1
	   For I = 0 To UnRegIndex + 1
	    
	     If NextPart.PartType = ListNonRegisteredParts$(I) Then
	       ShouldDo = 0
	       
		   Exit For
		 End If
		Next I
		If ShouldDo = 1 Then
		   ListNonRegisteredParts$(UnRegIndex) = NextPart.PartType ' for internal use
		   DialogListNonRegisteredParts$(UnRegIndex) = (NextPart.PartType) 'for external use
		   UnRegIndex = UnRegIndex+1
		End If  
	End If
Next NextPart

'redefine arrays, but preserve the values in them, so that they are now 
'the size they really need to be
ReDim Preserve ListRegisteredParts$(RegIndex)
ReDim Preserve ListNonRegisteredParts$(UnRegIndex)
ReDim Preserve DialogListRegisteredParts$(RegIndex)
ReDim Preserve DialogListNonRegisteredParts$(UnRegIndex)

End Sub

Private Function CrossProbeToApplication (subject,Pan) As Boolean
Dim x As Double
Dim y As Double

'select part user chose in dialog

   ActiveDocument.SelectObjects(ppcbObjectTypePartType,subject,True)
 

If pan = 1 Then
 
End If    

End Function
