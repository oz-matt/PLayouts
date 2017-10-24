'This VB Script is suitable for both PowerLogic and PowerPCB 
'The sample demonstrates the following PowerLogic/PowerPCB VB capabilities:
' - Part Attribute Collection
' - User Dialog with dynamic list boxes
' - Text File Reports
' - Simple Sort function
' - Clipboard using
' - Microsoft Excel connection and formatting.
' 
' For more details, please refer to the PowerLogic/PowerPCB Visual Basic Editor Help File.

Dim AttrsAll() As String
Dim AttrsSelected() As String
Dim parts As objects

Sub Main
	Begin Dialog UserDialog 580,280,"Attributes to Excel",.dialogfunc ' %GRID:10,7,1,1
		ListBox 10,21,210,259,AttrsAll(),.List1
		ListBox 350,21,220,259,AttrsSelected(),.List2
		Text 10,7,120,14,"Part Attributes:",.Text1
		Text 350,7,200,14,"Selected Attributes:",.Text2
		PushButton 230,21,110,21,"Add >>",.btnAdd
		PushButton 230,49,110,21,"Add All >>",.btnAddAll
		PushButton 230,91,110,21,"<< Remove",.btnRemove
		PushButton 230,119,110,21,"<< Remove All",.btnRemoveAll
		OKButton 230,224,110,21
		PushButton 230,196,110,21,"Apply",.btnExportToExcel
		CancelButton 230,252,110,21
	End Dialog
	Dim dlg As UserDialog

	rc = Dialog(dlg)
End Sub

Rem See DialogFunc help topic for more information.
Private Function dialogfunc(DlgItem$, Action%, SuppValue&) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		FillListOfAttributes
	Case 2 ' Value changing or button pressed
		dialogfunc = True ' Prevent button press from closing the dialog box
		If DlgItem$ = "btnAdd" Then 
			AddOne
		ElseIf DlgItem$ = "btnAddAll" Then 
			AddAll
		ElseIf DlgItem$ = "btnRemove" Then 
			RemoveOne
		ElseIf DlgItem$ = "btnRemoveAll" Then 
			RemoveAll
		ElseIf DlgItem$ = "btnExportToExcel" Then 
			ExportToExcel
		ElseIf DlgItem$ = "OK" Then
			ExportToExcel
			dialogfunc = False
		ElseIf DlgItem$ = "Cancel" Then
			dialogfunc = False
		Else
			UpdateButtons
		End If
		
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem dialogfunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Sub FillListOfAttributes
	Set parts = ActiveDocument.Components
	GetAttributeListWithPartType
	ReDim AttrsSelected(0)
	DlgListBoxArray "List1",AttrsAll()
	UpdateButtons
End Sub

Sub UpdateButtons
	AllCount = UBound(AttrsAll)
	SelCount = UBound(AttrsSelected)
	CurItem1 = DlgValue("List1") 
	CurItem2 = DlgValue("List2") 
	DlgEnable "btnAdd", allCount <> 0 And CurItem1 >= 0
	DlgEnable "btnAddAll", allCount <> 0
	DlgEnable "btnRemove", selCount <> 0 And CurItem2 >= 0
	DlgEnable "btnRemoveAll", selCount <> 0
	DlgEnable "btnExportToExcel", selCount <> 0
	DlgEnable "OK", selCount <> 0
End Sub

Sub AddOne
	CurItem = DlgValue("List1") 
	Count = UBound(AttrsSelected)
	ReDim Preserve AttrsSelected(Count + 1)
	AttrsSelected(Count + 1) = AttrsAll(CurItem + 1)
	DlgListBoxArray "List2",AttrsSelected()
	DlgValue "List1", CurItem + 1
	UpdateButtons
End Sub

Sub AddAll
	prevCount = UBound(AttrsSelected)
	ReDim Preserve AttrsSelected(0 To prevCount + UBound(AttrsAll))
	For i = 1 To UBound(AttrsAll)
		AttrsSelected(prevCount + i) = AttrsAll(i)
	Next i
	DlgListBoxArray "List2",AttrsSelected()
	UpdateButtons
End Sub

Sub RemoveOne
	CurItem = DlgValue("List2")
	Count = UBound(AttrsSelected)
	For i = CurItem + 2 To Count
		AttrsSelected(i - 1) = AttrsSelected(i)
	Next i
	ReDim Preserve AttrsSelected(Count - 1)
	DlgListBoxArray "List2",AttrsSelected()
	DlgValue "List2", CurItem
	UpdateButtons
End Sub

Sub RemoveAll
	ReDim Preserve AttrsSelected(0)
	DlgListBoxArray "List2",AttrsSelected()
	UpdateButtons
End Sub

Sub ExportToExcel
    FillClipboard
    
    On Error Resume Next
    Dim excelApp As Object
	Set excelApp =  GetObject(,"Excel.Application")
	On Error GoTo ExcelError	' Enable error trapping.
	If excelApp Is Nothing Then 
		Set excelApp =  CreateObject("Excel.Application")
	End If
	excelApp.Workbooks.Add
    excelApp.Visible = True
    excelApp.ActiveSheet.Paste
    
	Set header = excelApp.Range(excelApp.Cells(1, 1), excelApp.Cells(1, UBound(AttrsSelected) + 1))
	Set allCells = excelApp.Range(excelApp.Cells(1, 1), excelApp.Cells(parts.Count + 1, UBound(AttrsSelected) + 1))
	' Format title row
	header.Font.Bold = True
	header.Font.Italic = True
	allCells.Columns(1).Font.Bold = True
	' Format array names
	header.AutoFilter
	allCells.Columns.AutoFit
    excelApp.Range("A2").Select
    excelApp.ActiveWindow.FreezePanes = True
	On Error GoTo 0 ' Disable error trapping. 
	StatusBarText = "Report completed"
	Exit Sub    

ExcelError:
    MsgBox Err.Description, vbExclamation, "Error Running Excel"
    On Error GoTo 0 ' Disable error trapping.    
    Exit Sub
End Sub

Sub FillClipboard
	StatusBarText = "Export Data to Excel..."
    ' Create temporarly tab-delimited text file
    tempFile = DefaultFilePath & "\attr2xl.tmp"
    Open tempFile For Output As #1

    ' Output data to temporarly text file
    Print #1, "Name"; vbTab;
    For i = 1 To UBound(AttrsSelected)
    	Print #1, AttrsSelected(i); vbTab;
    Next i
    Print #1
    For Each aPart In parts
    	Print #1, aPart; vbTab; 
    	Set attrs = aPart.Attributes
    	For i = 1 To UBound(AttrsSelected)
    		If AttrsSelected(i) = "Part Type" Then
    			Print #1, aPart.PartType;
    		Else
	    		Set attr = attrs(AttrsSelected(i))
    			If Not attr Is Nothing Then Print #1, attr.Value;
    		End If
    		
   			If i <> UBound(AttrsSelected) Then Print #1, vbTab; 
    	Next i
	    Print #1
    Next aPart
    ' Close the temporarly text file
    Close #1
	' Load whole file to string variable    
    Open tempFile  For Input As #1
    L = LOF(1)
    AllData$ = Input$(L,1)
    Close #1
    'Copy whole data to clipboard
    Clipboard AllData$ 
    Kill tempFile
	StatusBarText = ""
End Sub

Public Sub GetAttributeListWithPartType()
	StatusBarText = "Scanning all parts for attribute Info..."
	' calculate total count of attributes
	totalCount = 0
	For Each obj In parts
		totalCount = totalCount + obj.Attributes.Count
	Next
	' fill string array of attribute names avoiding duplications	
	MyCount = 1
	ReDim AttrsAll(0 To totalCount + 1)
	AttrsAll(1) = "Part Type"
	For Each nextObj In parts
		Set attrs = nextObj.Attributes
		For Each nextAttr In attrs
			bFound = False
			'avoid duplications
			For i=1 To myCount
				If AttrsAll(i) = nextAttr.Name Then
					bFound=True
					Exit For
				End If
			Next i
			If bFound=False Then
				myCount = myCount+1
				AttrsAll(myCount) = nextAttr.Name
			End If
		Next nextAttr
	Next nextObj
	If (MyCount = 0) Then
		ReDim  AttrsAll(0)
	Else
		ReDim Preserve AttrsAll(MyCount)
	End If
	'sort list
	SortStringArray AttrsAll()
	StatusBarText = ""
End Sub

'Sort String Array
Public Sub SortStringArray(ByRef strList$())
	lo = 1
	hi = UBound(strList)
    While hi > lo
        max = lo
        For p = lo+1 To hi
			If StrComp(strList(p), strList(max), 1) > 0 Then
                max = p
            End If
		Next
		'swap max with hi
		temp = strList(max)
		strList(max) = strList(hi)
		strList(hi) = temp
		hi = hi - 1
    Wend
End Sub
