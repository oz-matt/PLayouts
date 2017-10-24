'Part List Report
'
'Written 11/98 by Kevin Adams, Xerox Engineering Systems
'
'This program outputs all parts in alphabetical order by RefDes and identifies which are SMD parts.
'Note, I check the drill size of the first pin, because the SMD flag is not always right!
'It also suppresses our global fiducials and mounting holes and other special items.
'
'PADS Comment: QuickSortMultiStringArray() is not needed for V3.0 
'since there is an internal sort function Objects.Sort

Sub Main

	Const conEvents = 200
	Const conSortEvent = 20
	Dim varnumcomp As Integer
	Dim varSort1 As String
	Dim varSort2 As String
	Dim varSort3 As String
	Dim varSort4 As String
	Dim varSort5 As String
	Dim varSort6 As String
	Dim varSort7 As String
	Dim varSort8 As String
	Dim varTemp As String
	Dim part(0 To 7, 0 To 2000) As String
	Dim nextComp As Object
	Dim PCB As Object

	ActiveDocument.SelectObjects(9999, "", False)	
	Set PCB = ActiveDocument.GetObjects(1, "", False)

	If PCB.Count = 0 Then
		MsgBox "You have not loaded a PCB file yet!"
		GoTo EndProg:
	End If

	' Lock server to speed up process
	LockServer

	MsgBox "Will Begin Part List Report.  It may take a minute to locate and sort all the components."

	' Go through each component in the design and output values
	varnumcomp = 0

	For Each nextComp In PCB
		With nextComp
			varTemp = .Name
			If Left$(varTemp,1) <> "A" And Left$(varTemp,1) <> "M" And Left$(varTemp,2) <> "TH"  And Left$(varTemp,2) <> "TS"  And Left$(varTemp,3) <> "TPD"  And Left$(varTemp,2) <> "HH" Then
				part(0,varNumComp) =  varTemp
				part(1,varNumComp) = .PartType
				part(2,varNumComp) = ActiveDocument.LayerName(.layer)
				part(3,varNumComp) = .Orientation
				part(4,varNumComp) = Str(.PositionX(0))
				part(5,varNumComp) = Str(.PositionY(0))
				If .Pins(1).DrillSize(0) = 0 Then
					part(6,varNumComp) = "Yes"
				Else
					part(6,varNumComp) = "No"
				End If
				part(7,varNumComp) = .Decal
			varNumComp = varNumComp + 1
			End If
		End With
		If (varNumComp Mod conEvents) = 0 Then DoEvents
	Next nextComp

	' Unlock the server
	UnlockServer

	QuickSortMultiStringArray part,0,varNumComp-1,8,1

	' Printout to Text file
	Randomize
	
	filename = DefaultFilePath & "\tmp"  & CInt(Rnd()*10000) & ".txt"
	Open filename For Output As #1

	Print #1, "Part List Report"
   	Print #1
	Print #1, "Report For PCB File: "; ActiveDocument.Name
	Print #1, "Date: "; Format(Date,"short date");"   Time: ";Format(Time,"HH:MM");" EST"
	Print #1
	Print #1, "Sorted by Reference Designator."
	Print #1, "Mounting holes, Global Fiducials, and TestPads/TestHoles have been removed from this list.
	Print #1, "Orientation is based on original library decal orientation and 0 deg is UP (facing top view of decal)."
	Print #1, "X and Y locations are centroid for SMD parts and pin 1 for through-hole parts."
	Print #1
		
	' Output Headers
	Print #1, "RefDes";		Space(5); 
	Print #1, "PartType";	Space(22); 
	Print #1, "Layer";		Space(25); 
	Print #1, "Orient.";		Space(7); 
	Print #1, "X";			Space(11); 
	Print #1, "Y";			Space(11); 
	Print #1, "SMD";		Space(7);
	Print #1, "Decal Name"
	
	For varCount = 0 To varNumComp - 1
		Print #1, part(0,varCount);		Space$(11-Len(part(0,varCount))); 
		Print #1, part(1,varCount);	Space$(30-Len(part(1,varCount))); 
		Print #1, part(2,varCount);	Space$(30-Len(part(2,varCount))); 
		Print #1, part(3,varCount);	Space$(13-Len(part(3,varCount))); 
		Print #1, part(4,varCount);	Space$(12-Len(part(4,varCount))); 
		Print #1, part(5,varCount);	Space$(13-Len(part(5,varCount))); 
		Print #1, part(6,varCount);		Space$(10-Len(part(6,varCount))) ;
		Print #1, part(7,varCount)
		Next varCount

	Print #1
	Print #1, "Total Number of Components: "; varnumcomp

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

