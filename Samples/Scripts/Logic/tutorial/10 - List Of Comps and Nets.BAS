' Sample 10: List Of Comps and Nets.BAS
' 
' This sample demonstrates how to go through the list of PADS Logic database objects.
' It displays a dialog which lists all components and nets in the current design. Click on
' a component ot a net will select it.
'
' For more details, please refer to the PADS Logic Basic Editor Help File.
'

Dim ListComps$(10000)
Dim ListNets$(10000)
Sub Main
	index = 0
	Dim nextComp As Object
	For Each nextComp In ActiveDocument.Components
		ListComps$(index) = nextComp.Name
		index = index + 1
	Next nextComp
	
	index = 0
	Dim nextNet As Object
	For Each nextNet In ActiveDocument.Nets
		ListNets$(index) = nextNet.Name
		index = index + 1
	Next nextNet
	
	Begin Dialog UserDialog 250,336,"Yoman",.CallbackFunc ' %GRID:10,7,1,1
		ListBox 10,21,100,259,ListComps(),.ListBox1
		OKButton 10,308,230,21
		ListBox 120,21,120,259,ListNets(),.ListBox2
		Text 10,7,90,14,"Components",.Text1
		Text 120,7,90,14,"Nets",.Text2
		Text 10,287,230,14,"Click on listboxes to select item",.Text3
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub

Function CallbackFunc%(DlgItem$, Action%, SuppValue%)
	Select Case Action%
		Case 2 ' Value changing or button pressed
			If DlgItem$ = "ListBox1" Then
				ActiveDocument.SelectObjects(plogObjectTypeAll, "", False)
				ActiveDocument.SelectObjects(plogObjectTypeComponent, ListComps(SuppValue%), True)
				LocateSelection
			End If
			If DlgItem$ = "ListBox2" Then
				ActiveDocument.SelectObjects(plogObjectTypeAll, "", False)
				ActiveDocument.SelectObjects(plogObjectTypeNet, ListNets(SuppValue%), True)
				LocateSelection
			End If
	End Select
End Function
	
Sub LocateSelection
	'look for selection in current sheet: if it has selection then do nothing
	If ActiveDocument.ActiveSheet.GetObjects(,,True).Count > 0 Then Exit Sub
	For Each Sheet In ActiveDocument.Sheets
		If Sheet.GetObjects(,,True).Count > 0 Then 
			Sheet.Activate
			Exit Sub
		End If
	Next
End Sub
