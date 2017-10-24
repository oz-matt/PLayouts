 '
' This will check each component to see which side it is mounted on.  Then output a bom file for everything mounted on the top
'and everything mounted on the bottom
'mb 11-13-98

Dim ListComps$(10000)

Sub Main
	filenamebot = DefaultFilePath + "\" +"bombot.txt"
	filenametop = DefaultFilePath + "\" +"bomtop.txt"
	Open filenametop For Output As #1
	Open filenamebot For Output As #2


	index = 0
	Dim nextComp As Object
	For Each nextComp In ActiveDocument.GetObjects(1, "", False)
	
		ListComps$(index) = nextComp.Name + "	" + nextComp.PartType + "	" + nextComp.Decal + "	"+"Layer"+Str(nextcomp.layer)
		If nextcomp.layer =1 Then Print #1, ListComps$(index) 	Else Print #2, ListComps$(index)
		index = index + 1
	Next nextComp
	
Close 
	'open notepad to view the files
	notepadtop = "notepad " + filenametop
	opentop=Shell(notepadtop,1)
	AppActivate opentop
	notepadbot = "notepad " + filenamebot
	openbot=Shell(notepadbot,1)
	AppActivate openbot
	
End Sub