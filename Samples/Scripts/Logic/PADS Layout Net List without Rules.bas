Dim nTotal As Long
Dim nCurrent As Long
Dim nPercent As Integer

Sub Main
	'set this flag to False if you do not need attributes info
	includePartAttributes = True
    'show progress bar
	StatusBarText = "Generating NetList..."
	ProgressBar = 0
	
	nTotal = ActiveDocument.Components.Count + ActiveDocument.Nets.Count
	If includePartAttributes Then nTotal = nTotal + ActiveDocument.Components.Count
	
	fname = ActiveDocument
	If fname = "" Then 
		fname = "Untitled"
		report = "default.asc"
	Else
		nm = Left(fname, Len( fname) - 4)
		report = DefaultFilePath & "\" & nm & "1.asc"
	End If
	Open report For Output As #1

	Print #1, "!PADS-POWERPCB-V4.0-MILS! NETLIST FILE FROM POWERLOGIC V4.0"
	Print #1, "*REMARK* "; fname; " --  "; Now
	Print #1, "*REMARK*  "
	Print #1
	
	Print #1, "*PART*          ITEMS"
	For Each part In ActiveDocument.Components
		Print #1, part; Space(16 - Len(part));
		Print #1, part.PartType;
		Print #1, "@"; part.PCBDecal
		IncProgress
	Next

	Print #1, "*NET*"
	For Each Net In ActiveDocument.Nets
		Print #1, "*SIGNAL* "; Net
		i = 0
		For Each pin In net.Pins
			Print #1, pin; " ";
			i = i +1
			If i = 5 Then Print #1: i = 0	'allow 5 pins on each row
		Next
		If i > 0 Then Print #1
		IncProgress
	Next
	
	If includePartAttributes Then
		Print #1
		Print #1, "*MISC*      MISCELLANEOUS PARAMETERS"
		Print #1
		Print #1, "ATTRIBUTE VALUES"
		Print #1, "{
		For Each part In ActiveDocument.Components
			If part.Attributes.Count > 0 Then
				Print #1, "PART "; part.Name
				Print #1, "{"
				For Each attr In part.Attributes
					Print #1, """"; attr.Name; """ "; attr.value
				Next
				Print #1, "}"
			End If
			IncProgress
		Next
		Print #1, "}"
	End If
	Print #1
	Print #1, "*END*     OF ASCII OUTPUT FILE"
	Close #1
	'show text file
	Shell "notepad " & report, 1
    'hide progress bar
	StatusBarText = ""
	ProgressBar = -1
End Sub

Sub IncProgress
	nCurrent = 	nCurrent + 1
	nNewProg = nCurrent * 100 / nTotal
	If nPercent <> nNewProg Then
		nPercent = nNewProg
		ProgressBar = nPercent
	End If
End Sub


