Dim xl As Object
Dim row As Integer
Dim part As Object
Sub Main
	Set xl = CreateObject("Excel.Application")
	xl.Workbooks.Add
	xl.Visible = True
	Set cell = xl.ActiveCell
	row = 2
	For Each comp In ActiveDocument.Components
		cell.Item(row, 1) = comp.Name
		cell.Item(row, 2) = comp.PartType
		cell.Item(row, 3) = AttrValue(comp, "PART DESC")
		cell.Item(row, 4) = AttrValue(comp, "MFG. #1")
		cell.Item(row, 5) = AttrValue(comp, "VALUE")
		cell.Item(row, 6) = AttrValue(comp, "TOLERANCE")
		row = row + 1
	Next
	On Error GoTo 0
	xl.Columns("A:F").NumberFormat = "@"
	xl.Range("A1:F1") = Array("RefDes","Part Type","Description","Manufacturer","Value","Tolerance")
	xl.Range("A1:F1").Font.Bold = True
	xl.Range("A1:F1").AutoFilter
	xl.Range("A1").Subtotal GroupBy:=2, Function:=2, TotalList:=Array(2)
	xl.Columns("A:F").AutoFit
End Sub

Function AttrValue (comp As Object, atrName As String) As String
	If comp.Attributes(atrName) Is Nothing Then
		AttrValue = ""
	Else
		AttrValue = comp.Attributes(atrName).value
	End If
End Function

