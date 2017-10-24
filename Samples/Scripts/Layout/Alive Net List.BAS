'#Reference {00020813-0000-0000-C000-000000000046}#1.2#0#D:\Program Files\Microsoft Office\Office\EXCEL8.OLB#Microsoft Excel 8.0 Object Library
' PADS Layout VB Script: Alive Net List.BAS
' 
' This sample demonstrates how to use new features of PADS Layout V3.0 automation server:
'  - how to handle events of external automation servers (Excel in our case)
'  - how to handle events of PADS Layout itself
'  - how to organize bidirectional cross probing between PADS Layout and external automation server
'  - how to connect to already running extarnal automation server
'  - how to use PADS Layout progress bar
'
' ! You need to have Microsoft Excel installed to run this script

Dim WithEvents xl As Excel.Application
Dim procSel As Boolean
Dim Title As String
Dim nTotal As Long
Dim nCurrent As Long
Dim nPercent As Integer

Sub Main
	Title = "!PADS-POWERPCB-V2.0-MILS! DESIGN DATABASE ASCII FILE 1.0"
	On Error Resume Next
	Set xl =  GetObject(,"Excel.Application")
	On Error GoTo 0
	If xl Is Nothing Then
		Set xl =  New Excel.Application
		FillCells
	Else
		If xl.Range("A1") = Title And xl.Range("B3") = ActiveDocument.Name Then
			ans = MsgBox ("Connect to existing Net List?", vbYesNoCancel)
			If ans  = vbNo Then
				FillCells
			ElseIf ans = vbCancel Then
				Exit Sub
			End If
		Else
			FillCells
		End If
	End If
   	xl.Visible= True
    procSel = True
   
    MsgBox "WHILE THIS MESSAGE ON THE SCREEN:" & vbCr & _
                     "Either select non-empty cell(s) in Excel" & vbCr & _
                     "or select Part(s), Net(s), or/and Pin(s)" &  vbCr & _
                     "in PADS Layout design" & vbCr & _
                     "Press Ok to finish cross probing", _
                      0, "Cross Probing"
End Sub

Public Sub Document_SelectionChange()
	If Not procSel Then Exit Sub	'avoid recursion
    procSel = False
    'show progress bar
	nCurrent = 0
	nTotal = ActiveDocument.GetObjects(ppcbObjectTypeComponent,"", True).Count
	nTotal = nTotal + ActiveDocument.GetObjects(ppcbObjectTypeNet,"", True).Count
	nTotal = nTotal + ActiveDocument.GetObjects(ppcbObjectTypePin,"", True).Count
	If nTotal > 1 Then
		StatusBarText = "Sending selection to Excel..."
		ProgressBar = 0
	End If
    Dim aRange As Object
    SelectCells ppcbObjectTypeComponent, aRange
    SelectCells ppcbObjectTypeNet, aRange
    SelectCells ppcbObjectTypePin, aRange
    If Not aRange Is Nothing Then
		aRange.Select
		xl.Visible= True	'bring Excel to the top
	Else
		'clear selection in Excel
		xl.ActiveCell.Select
    End If
    'hide Progress
	ProgressBar = -1
	StatusBarText = ""
    procSel = True
End Sub

Private Sub xl_SheetSelectionChange(Sh As Object, Target As Range)
	If Not procSel Then Exit Sub	'avoid recursion
    procSel = False
	ActiveDocument.SelectObjects ( ,, False)
	Set aRange = xl.Intersect(xl.ActiveSheet.UsedRange, Target.Cells)
	If Not aRange Is Nothing Then
	    'show progress bar if needed
		nCurrent = 0
		nTotal = aRange.Count
		If nTotal > 1 Then
			nCurrent = 0
			StatusBarText = "Receiving selection from Excel..."
			ProgressBar = 0
		End If
		For Each cell In aRange
			objName = Str(cell)
			If objName <> "" Then
				If objName = "*SIG*" Then objName = Str(cell.Next)
				With ActiveDocument
					Do
						Set aComp = .Components(objName)
						If Not aComp Is Nothing Then aComp.Selected = True:	Exit Do
						Set aNet = .Nets(objName)
						If Not aNet Is Nothing Then	aNet.Selected = True: Exit Do
						'skip non-pins 
						If InStr(objName, ".") = 0 Then Exit Do
						Set aPin = .Pins(objName)
						If Not aPin Is Nothing Then	aPin.Selected = True: Exit Do
						Exit Do
					Loop
				End With
			End If
			IncProgress
		Next
	End If		
    'hide Progress
	ProgressBar = -1
	StatusBarText = ""
    procSel = True
End Sub

Sub SelectCells (aType, ByRef aRange As Object)
	For Each obj In ActiveDocument.GetObjects(aType,"", True)
		Set found = xl.Cells.Find(obj, ,xlValues, xlWhole)
		If Not found Is Nothing Then
			If aRange Is Nothing Then
				Set aRange = found
			Else 
				Set aRange = xl.Union(aRange, found)
			End If
		End If
		IncProgress
    Next
End Sub

Sub FillCells
    'show progress bar
	StatusBarText = "Filling Excel Cells..."
	ProgressBar = 0
	nCurrent = 0
	nTotal = ActiveDocument.Components.Count + ActiveDocument.Nets.Count
	xl.Workbooks.Add
	xl.Visible= True
	xl.ActiveWorkbook.Styles("Normal").NumberFormat = "@"
	Set cell = xl.Range("A1")
    cell.Item(1,1) = Title
    cell.Item(2,1) = "*PART*       ITEMS"
    cell.Item(3,1) = "*REMARK*"
    cell.Item(3,2) = ActiveDocument.Name
    xl.Rows("1:3").Font.Bold = True
    i = 4
    For Each comp In ActiveDocument.Components
    	cell.Item(i,1) = comp.Name
    	cell.Item(i,2) = comp.PartType
    	i = i + 1
    	IncProgress
    Next
    i = i + 2
    cell.Item(i,1) = "*NET*
    xl.Rows(i & ":" & i).Font.Bold = True
    i = i + 2
    For Each aNet In ActiveDocument.Nets
    	cell.Item(i,1) = "*SIG*"
    	cell.Item(i,2) = aNet.Name
	    xl.Rows(i & ":" & i).Font.Bold = True
    	i = i + 1
    	j = 1
    	For Each aPin In aNet.Pins
    		If j = 6 Then j = 1: i = i + 1
    		cell.Item(i,j) = aPin.Name
    		j = j + 1
    	Next
    	i = i + 2
    	IncProgress
    Next
    'hide progress bar
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


