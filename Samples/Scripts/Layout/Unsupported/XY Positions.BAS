' XYpositions.BAS
' Based on Sample 17: Excel Part List Report.BAS
' 
' SR Telecom Position File
'
'

Sub Main
	' Open  text file
	Randomize
	filename = DefaultFilePath & "\Positions"  & CInt(Rnd()*10000) & ".txt"
	Open filename For Output As #1

	' Output Headers
	Print #1, "RefDes";		Space(5); 
	Print #1, "X";			Space(10); 
	Print #1, "Y";			Space(10); 
	Print #1, "Rot.";	    Space(7); 
	Print #1, "Exact";      Space(5);
	Print #1, "(FREE)";		Space(4);
	Print #1, "Side";		Space(0)
		
	' Lock server to speed up process
	LockServer

	' Go through each component in the design and output values
	Dim nextComp As Object
	For Each nextComp In ActiveDocument.GetObjects(1, "", False)
		Print #1, nextComp.Name;		Space$(11-Len(nextComp.Name)); 
            a = Format$(nextComp.PositionX(0), "00.0000")
		Print #1, a; Space$(11-Len(a)); 
            a = Format$(nextComp.PositionY(0), "00.0000")
		Print #1, a; Space$(11-Len(a)); 
            a = Format$(nextComp.Orientation, "000")
		Print #1, a; Space$(11-Len(a)); 
		Print #1, "EXACT";     	Space(5); 
		Print #1, "(FREE)";		Space(4);
		Print #1, ActiveDocument.LayerName(nextComp.layer);	Space$(30-Len(ActiveDocument.LayerName(nextComp.layer)))
	Next nextComp

	' Unlock the server
	UnlockServer

	' Close the text file
	Close #1
	
		' Display the text file
	Shell "Notepad " & filename, 3

End Sub
