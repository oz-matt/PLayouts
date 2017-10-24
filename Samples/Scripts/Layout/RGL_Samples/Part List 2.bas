'Parts List 2

'#Uses "RGL.bas"
Sub Main
	Set doc = ActiveDocument
	outFile = OpenReport (DefaultFilePath & "\Part List 2.rep")
	
	Print #outFile, "Part List Report -- ";
	Print #outFile, doc.Name; " -- ";
	Print #outFile, GetTime

	'Returns the total number of packages (part types).
	Out "Total packages = " & PkgCnt(doc)

	Out "Part Type             Description"
	Out "-------------------------------------------------------------------------------"

	Dim pkgs() As Package
	GetPackageList doc, pkgs, True 'get sorted list of packages
	
	'Search within packages
	Columns 1, 22
		For i = 1 To UBound(pkgs)			
			Out pkgs(i).Name, pkgs(i).Description
	
			Columns 6, 23, 40, 57, 74
				For Each nextComp In doc.Components
					If pkgs(i).Name = nextComp.PartType Then
						Out nextComp
					End If
				Next nextComp
			End_Columns
			
			Out
		Next i	
	End_Columns
	
	CloseReport
End Sub


