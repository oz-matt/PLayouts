Dim part As Object
Sub Main
	report = DefaultFilePath & "\report.rep"
	Open report For Output As #1
	Print #1, "Bill Of Materials for "; ActiveDocument; " on  "; Date
	sep = Chr(9)
	item = 1
	For Each comp In ActiveDocument.Components
		Print #1,	item; 
		Print #1,	sep; comp.Name; 
		Print #1,	sep; comp.PartType;
		Print #1,	sep; AttrValue(comp, "PART DESC");
		Print #1,	sep; AttrValue(comp, "MFG. #1");
		Print #1,	sep; AttrValue(comp, "VALUE");
		Print #1,	sep; AttrValue(comp, "TOLERANCE");
		Print #1
		item = item + 1
	Next
	Close #1
	Shell "notepad " & report, 1
End Sub

Function AttrValue (comp As Object, atrName As String) As String
	If comp.Attributes(atrName) Is Nothing Then
		AttrValue = ""
	Else
		AttrValue = comp.Attributes(atrName).value
	End If
End Function
