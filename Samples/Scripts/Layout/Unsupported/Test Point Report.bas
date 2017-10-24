'Test Point Locations Report
'
'Written 11/98 by Kevin Adams, Xerox Engineering Systems
'
'This file reports details on all testpoints in a design. It also lists what nets are missing testpoints.
'It also flags which testpoints are not on .100 inch centers.  It scans component pins and vias for testpoints.
'
'Modified
'12/2/98 - Added capability to list nets that are missing test points
'
Option Explicit

Const conSortEvent = 20
Const conEvent = 200
Const conNameSpace = 20 
Const conNetSpace = 19 
Const conPosXSpace = 10
Const conPosYSpace = 12

Sub Main

	On Error GoTo 0

				
	Dim nextObject As Object
	Dim PCB As Object
	Dim testpt (0 To 4,0 To 4000) As String
	Dim nets(0 To 3000) As String
	Dim varNumTestpts As Integer
	Dim varNumNets As Integer
	Dim varCount As Integer
	Dim varCount2 As Integer
	Dim varCount3 As Integer
	Dim varFound As Integer
	Dim varSort1 As String
	Dim varSort2 As String
	Dim varSort3 As String
	Dim varSort4 As String
	Dim varSort5 As String
	Dim filename As String
	
	' Go through each component pin in the design and find  testpoints

	ActiveDocument.SelectObjects(9999, "", False)
	Set PCB = ActiveDocument.GetObjects(1, "", False)
	If PCB.Count = 0 Then
		MsgBox "You have not loaded a PCB file yet!"
		GoTo EndProg:
	End If
	
	MsgBox "Will Begin Testpoint Report.  It may take a minute to locate and sort all Testpoints and Nets."

	' Lock server to speed up process
	LockServer
	
	varNumTestpts = 0

	'Get all components in database.  Then pull out pins by component.
	On Error Resume Next
	For varCount = 1 To PCB.Count
		Set NextObject = PCB(varCount).Pins
		For varCount2 = 1 To  NextObject.Count
			With nextObject(varCount2)
				If .TestPoint <> 0 Then  
					testpt(0,varNumTestpts) = .Name
					testpt(1,varNumTestpts) = .Net
					If Err.Number > 0 Then
						testpt(1,varNumTestpts) = "UNUSED PIN"
						Err.Clear
					End If
					testpt(2,varNumTestpts) = Str(.PositionX(0))
					testpt(3,varNumTestpts) = Str(.PositionY(0))
					If .DrillSize(0) = 0 Then 
						testpt(4,varNumTestpts) = "Yes"
					Else
						testpt(4,varNumTestpts) = "No"
					End If
					varNumTestpts = varNumTestpts + 1
				End If
			End With
		Next varCount2
		If (varcount Mod conEvent) = 0 Then DoEvents
		Set NextObject = Nothing
	Next varCount
	
	' Go through each via in the design and find testpoints
	Set PCB = ActiveDocument.GetObjects(4, "", False)
	On Error Resume Next
	For varCount = 1 To PCB.Count
		With PCB(varCount)
			If .TestPoint <> 0 Then  
				testpt(0,varNumTestpts) = .Name
				testpt(1,varNumTestpts) = .Net
				If Err.Number > 0 Then
					testpt(1,varNumTestpts) = "UNUSED PINS"
					Err.Clear
				End If
				testpt(2,varNumTestpts) = Str(.PositionX(0))
				testpt(3,varNumTestpts) = Str(.PositionY(0))
				If .DrillSize(0) = 0 Then 
					testpt(4,varNumTestpts) = "Yes"
				Else
					testpt(4,varNumTestpts) = "No"
				End If
				varNumTestpts = varNumTestpts + 1
			End If
		End With
		If (varcount Mod conEvent) = 0 Then DoEvents
	Next varCount

	On Error GoTo 0	

	' Unlock the server
	UnlockServer
	
	'Sort by net to make net cross-reference faster
	QuickSortMultiStringArray testpt, 0, varNumTestpts-1, 5, 2

	' Lock server to speed up process
	LockServer
	
	varNumNets = 0
	
	' Go through each net in the design and see if it has a testpoint
	Set PCB = ActiveDocument.GetObjects(2, "", False)
	For varCount = 1 To PCB.Count
		With PCB(varCount)
			varSort1 = .Name
			varFound = 0
			For varCount2 = 0 To varNumTestpts - 1
				If testpt(1,varCount2) > varSort1 Then Exit For
				If varSort1 = testpt(1,varCount2)  Then  
					varFound = 1
				End If
			Next varCount2
			If varFound = 0  Then  
				Nets(varNumNets) = varSort1
				varNumNets = varNumNets + 1
			End If
		End With
		If (varNumNets Mod conEvent) = 0 Then DoEvents
	Next varCount

	' Unlock the server
	UnlockServer

	'Sort by name of via or component to display
	QuickSortMultiStringArray testpt, 0, varNumTestpts-1, 5, 1

	'Printout
	Randomize
	filename = DefaultFilePath & "\tmp"  & CInt(Rnd()*10000) & ".txt"
	Open filename For Output As #1

	Print #1, "Testpoint Report"
	Print #1
   	Print #1, "Report For PCB File: "; ActiveDocument.Name
	Print #1, "Date: "; Format(Date,"short date");"   Time: ";Format(Time,"HH:MM");" EST"
	Print #1
	Print #1, "Sorted by Pin Name"
		
	' Output Headers
	Print #1, "Pin/Via Name";	Space(8); 
	Print #1, "Assigned Net";	Space(10); 
	Print #1, "X";	Space(8); 
	Print #1, "Y" ;	Space(6);
	Print #1, "SMD Pad Only?"

	'Output testpoints	
	For varCount = 0 To varNumTestpts - 1
			Print #1, testpt(0,varCount) ;		Space$(conNameSpace - Len(testpt(0,varCount))); 
			Print #1, testpt(1,varCount) ; 		Space$(conNetSpace - Len(testpt(1,varCount))); 
			Print #1, testpt(2,varCount) ;		Space$(conPosXSpace - Len(testpt(2,varCount))); 
			Print #1, testpt(3,varCount) ; 		Space$(conPosYSpace - Len(testpt(3,varCount))) ; 
			Print #1, testpt(4, varCount);
			varcount2 = Val(testpt(2,varCount))
			varcount3 = Val(testpt(3,varCount))
			If (varcount2 Mod 100) > 0 Or (varcount3 Mod 100) > 0 Then
				Print #1, "   *"
			Else
				Print #1, ""
			End If
			DoEvents
	Next varCount

	Print #1, ""
	Print #1, "Total Number of Testpoints: "; varNumTestpts
	Print #1, ""
	Print #1, "(Testpoints marked with an * are not on standard .100 inch centers.)"
	Print #1

	'Output any nets that are missing testpoints
	If varNumNets > 0 Then
		If varNumNets > 1 Then
			Print #1, "THE FOLLOWING";varNumNets;" NETS ARE MISSING TESTPOINTS!"
		Else
			Print #1, "THE FOLLOWING NET IS MISSING A TESTPOINT!"
		End If 
		For varCount = 0 To varNumNets - 1
			Print #1, Space$(5);nets(varCount)
		Next varCount
	Else
		Print #1, "ALL NETS HAVE TESTPOINTS!"
	End If

	' Close the text file
	Close #1
		
	' Display the text file
	Shell "Notepad " & filename, 3
	Kill filename

EndProg:

	End Sub

Private Sub QuickSortMultiStringArray( _
  avarIn() As String, _
  ByVal intLowBound As Integer, _
  ByVal intHighBound As Integer, _
  ByVal intArray As Integer, _
  ByVal intSort As Integer)

  Dim intX As Integer
  Dim intY As Integer
  Dim intCount As Integer
  Dim varMidBound As String
  Dim varTmp As String
  
  On Error GoTo PROC_ERR

  If (intHighBound - intLowBound) >= conSortEvent Then DoEvents
  If intHighBound > intLowBound Then
    If intArray = 1 Then
	varMidBound = avarIn((intLowBound + intHighBound) \ 2)
    Else
    	varMidBound = avarIn(intSort-1,(intLowBound + intHighBound) \ 2)
    End If
    intX = intLowBound
    intY = intHighBound

    If intArray = 1 Then
        Do While intX <= intY
		If avarIn(intX) >= varMidBound And avarIn(intY) <= varMidBound Then
			varTmp = avarIn(intX)
	    		avarIn(intX) = avarIn(intY)
		       avarIn(intY) = varTmp
	       	intX = intX + 1
       		intY = intY - 1
		Else
       		If avarIn(intX) < varMidBound Then  intX = intX + 1
	       	If avarIn(intY) > varMidBound Then  intY = intY - 1
      		End If
      	  Loop
    Else
         Do While intX <= intY
	      	If avarIn(intSort-1,intX) >= varMidBound And avarIn(intSort-1,intY) <= varMidBound Then
			For intCount = 0 To intArray-1
		        	varTmp = avarIn(intCount,intX)
    				avarIn(intCount,intX) = avarIn(intCount,intY)
		        	avarIn(intCount,intY) = varTmp
			Next intCount
	       	intX = intX + 1
       	 	intY = intY - 1
		Else
       		If avarIn(intSort-1,intX) < varMidBound Then  intX = intX + 1
	       	If avarIn(intSort-1,intY) > varMidBound Then  intY = intY - 1
      		End If
	  Loop
    End If
    QuickSortMultiStringArray avarIn(), intLowBound, intY, intArray, intSort
    QuickSortMultiStringArray avarIn(), intX, intHighBound, intArray, intSort
  End If
    
PROC_EXIT:
Exit Sub

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "QuickSortMultiStringArray"
  Resume PROC_EXIT
End Sub
