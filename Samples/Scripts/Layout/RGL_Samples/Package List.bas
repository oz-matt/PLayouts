'Package List

'#Uses "RGL.bas"
Dim doc  As Document
Sub Main
	Set doc = ActiveDocument

	outFile = OpenReport (DefaultFilePath & "\Package List.rep")
	Print #outFile, "Package List Report -- ";
	Print #outFile, doc.Name; " -- ";
	Print #outFile, GetTime
	Print #outFile

	'Returns the total number of packages (part types).
	Out "Total packages = " & PkgCnt(doc)
	Out

	Out "Part Type        Reference Designation               Description"
	Out "-------------------------------------------------------------------------------"

	Dim pkgs() As Package
	GetPackageList doc, pkgs, True 'get sorted list of packages
	
	'Search within packages
	Columns 0, 50
		For i = 1 To UBound(pkgs)	
			Out pkgs(i).Name, pkgs(i).Description
	
			Columns 20, 30, 40
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


