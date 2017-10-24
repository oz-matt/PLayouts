'Net List Report
'
'Written 12/98 by Kevin Adams, Xerox Engineering Systems
'
'This program outputs the netlist for a PCB with spaces or periods as delimiter.  Uses Quicksort Algorythym to speed up process.  Takes about 30 seconds for 1400+ nets and ~7800 pins
'
'
'PADS Comment: QuickSortMultiStringArray() is not needed for V3.0 
'since there is an internal sort function Objects.Sort

Option Explicit

Sub Main

	On Error GoTo 0
	Const conNameSpace = 24 
	Const conRefDesSpace = 9
	Const ConNumCols = 6
	Const conEvent = 300
	Const conUnused = "UNUSED_NET"
	
	Dim varNumNets As Integer
	Dim varNumPins As Integer
	Dim varNumActivePins As Integer
	Dim varNumUnusedPins As Integer
	Dim varCount As Integer
	Dim varCount2 As Integer
	Dim varCount3 As Integer
	Dim varFound As Integer
	Dim varNewLine As Integer
	Dim varStart As Integer
	Dim varResp As Variant
	Dim varTemp As String
	Dim filename As String
	Dim nets(0 To 4000) As String 
	Dim pins(0 To 1, 0 To 12000) As String
	Dim NextObject As Object
	Dim PCB As Object

	ActiveDocument.SelectObjects(9999, "", False)
	Set PCB = ActiveDocument.GetObjects(2, "", False)

	If PCB.Count = 0 Then
		MsgBox "You have not loaded a PCB file yet!"
		GoTo EndReport:
	End If
	
	varResp = MsgBox ("Will begin Net List Report.  It may take a minute to locate and sort all nets and pins." & Chr(13) & Chr(13) _
					& "Do you want to use PERIODS INSTEAD OF SPACES for delimiting Refdes.Pin information (Default = Yes)?", vbYesNo)

	' Lock server to speed up process
	LockServer
	
	varNumNets = 0
	varNumPins = 0
	varNumActivePins = 0
	varNumUnusedPins = 0
	
	' Go through each net in the design and store in array
	For Each nextObject In PCB
		Nets(varNumNets) = nextObject.Name
		varNumNets = varNumNets + 1
		If (varNumNets Mod conEvent) = 0 Then DoEvents
	Next nextObject
	
	If varNumNets <> PCB.Count Then 
		MsgBox "The number of nets found does not match number reported by PADS!"
		GoTo EndReport
	End If
	
	Set PCB = Nothing

	'Get all components in database.  Then pull out pins by component.
	On Error Resume Next
	Set PCB = ActiveDocument.GetObjects(1, "", False)
	For varCount = 1 To PCB.Count
		Set NextObject = PCB(varCount).Pins
		For varCount2 = 1 To  NextObject.Count
			With nextObject(varCount2)
				pins(0,varNumPins) = .Name
				'For each pin, if user said substitute period (check Msgbox response), 
				'take period out of Refdes.Pin name if user wants spaces instead of periods
				If varResp = vbNo Then
					vartemp = .Name
					varcount3 = InStr(1,varTemp,".")
					pins(0,varNumPins) = Left$(vartemp,varcount3 - 1) & " " & Right$(vartemp,Len(vartemp) - varcount3)
				End If
				Pins(1,varNumPins) = .Net
				'If no net (error), give unused name
				If Err > 0 Then
					Pins(1, varNumPins) = conUnused
					varNumUnusedPins = varNumUnusedPins + 1
					Err.Clear
				End If
			End With
			varNumPins = varNumPins + 1
			If (varNumPins Mod conEvent) = 0 Then DoEvents
		Next varCount2
		Set NextObject = Nothing
	Next varCount
	
	Set PCB = Nothing

	' Unlock the server
	UnlockServer

	'Sort the Nets in ascending order
	QuickSortMultiStringArray nets, 0,varnumnets - 1,1,1

	DoEvents	

	'Sort Pins in ascending order by net
	QuickSortMultiStringArray pins, 0,varnumpins - 1,2,2

	On Error GoTo 0

	'Printout
	MsgBox "Nets and pins are identified and sorted. Now will begin cross-referencing pins per net. This may take a few minutes."

	Randomize
	filename = DefaultFilePath & "\tmp"  & CInt(Rnd()*10000) & ".txt"
	Open filename For Output As #1

	Print #1, "Net List Report"
   	Print #1
  	Print #1, "Report For PCB File: "; ActiveDocument.Name
	Print #1, "Date: "; Format(Date,"short date");"   Time: ";Format(Time,"HH:MM");" EST"
	Print #1
	Print #1, "Sorted by Net Name."
	Print #1
		
	' Output Headers
	Print #1, "Net Name";	Space(conNameSpace - Len("Net Name")); 
	If varResp = vbYes Then
		Print #1, "RefDes.Pin#"
	Else
		Print #1, "RefDes <sp> Pin#"
	End If

	'Output pins for each net
	varStart = 0
	For varCount = 0 To varNumNets - 1
		varcount3 = 0
		varNewline = 0
		varFound = 0
		Print #1, nets(varCount) ;	Space$(conNameSpace - Len(nets(varCount))); 
		For varCount2 = varStart To varNumPins - 1
			If pins(1,varCount2) > nets(varcount) And varFound = 1 Then 
				varStart = varCount2
				Exit For
			End If
			If pins(1,varCount2) = nets(varcount) Then
				varFound = 1
				If varnewline = 1 Then
					Print #1,Space$(conNameSpace);
					varNewLine = 0
				End If
				Print #1, pins(0,varCount2) ; Space$(conRefDesSpace - Len(pins(0,varCount2))); 
				varNumActivePins = varNumActivePins + 1
				varcount3 = varcount3 + 1
				If varcount3 = conNumCols Then
					Print #1
					varCount3 = 0
					varNewLine = 1
				End If
			End If
		Next varCount2
		If varFound = 0 Then Print #1, "UNUSED";
		If varNewLine = 0 Then Print #1
		If (varCount Mod conEvent) = 0 Then DoEvents
	Next varCount

	
	Print #1 
	Print #1, "Total Number of Nets: "; varNumNets
	Print #1, "Total Number of Pins: ";varNumPins
	Print #1, "Total Number of Active Pins: "; varNumActivePins
	Print #1, "Total Number of Unused Pins: "; varNumUnusedPins
	Print #1 
	If varNumUnusedPins > 0 Then
		'Sort Pins in ascending order by RefDes
		QuickSortMultiStringArray pins, 0,varnumpins - 1,2,1
		Print #1, "The pins that are UNUSED are as follows (sorted by RefDes):"
		Print #1
		varcount2 = 0
		For varcount = 0 To varNumPins - 1
			If pins(1,varCount) = conUnused Then
			Print #1, pins(0,varcount) ; 	Space$(conRefDesSpace - Len(pins(0,varcount))); 
			varcount2 = varcount2 + 1
			If varcount2 = (conNumCols + Int(conNameSpace/conRefDesSpace)) Then
				Print #1
				varcount2 = 0
			End If
			End If
			If (varCount Mod conEvent) = 0 Then DoEvents
		Next varcount
	End If
	
	Print #1
	Print #1
	Print #1, "End of Report."
	Print #1
	' Close the text file
	Close #1
		
	' Display the text file
	Shell "Notepad " & filename, 3
	Kill filename

EndReport:
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
  
  Const conSortEvent = 25

  On Error GoTo PROC_ERR

  If (intHighBound - intLowBound) >= conSortEvent Then 	DoEvents
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
