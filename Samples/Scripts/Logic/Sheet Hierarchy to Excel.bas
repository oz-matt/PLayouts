Dim row As Integer
Dim xl As Object
Sub Main
	row = 1
	Set xl = CreateObject("Excel.Application")
	xl.Workbooks.Add
	xl.Visible = True
	ShowHierLevel ActiveDocument.Sheets, True, 0
End Sub

Sub ShowHierLevel (collection, subSheets, level)
	Dim Val As String
	col = Chr (Asc("A") + level)
	For Each sheet In collection
		Val = sheet.Name
		xl.Range(col&row).Select
		xl.ActiveCell = Val
		row = row + 1
		If subSheets Then
			ShowHierLevel Sheet.ChildSheets, True, level+1
		End If
	Next
End Sub
