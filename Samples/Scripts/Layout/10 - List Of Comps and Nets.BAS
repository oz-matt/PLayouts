' Sample 10: List Of Comps and Nets.BAS
' 
' This sample demonstrates how to go through the list of PADS Layout database objects,
' how to allocate arrays with size known only during script execution,
' how to use PADS Layout predefined constants such as ppcbObjectTypeAll
' It displays a dialog which lists all components and nets in the current design. Click on
' a component ot a net will select it.
'
' For more details, please refer to the PADS Layout Basic Editor Help File.
'

'Declare empty arrays
Dim ListComps$()
Dim ListNets$()

Sub Main
	'Allocate arrays component and net names
	ReDim ListComps$(ActiveDocument.Components.Count)
	ReDim ListNets$(ActiveDocument.Nets.Count)
	
	index = 0
	'Iterate through all components in the database and fill array of names
	For Each nextComp In ActiveDocument.Components
		ListComps$(index) = nextComp.Name
		index = index + 1
	Next nextComp
	
	index = 0
	'Iterate through all nets in the database and fill array of names
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
				'Unselect all objects
				ActiveDocument.SelectObjects(ppcbObjectTypeAll, "", False)
				'Select component by given name
				ActiveDocument.SelectObjects(ppcbObjectTypeComponent, ListComps(SuppValue%), True)
			End If
			If DlgItem$ = "ListBox2" Then
				'Unselect all objects
				ActiveDocument.SelectObjects(ppcbObjectTypeAll, "", False)
				'Select net by given name
				ActiveDocument.SelectObjects(ppcbObjectTypeNet, ListNets(SuppValue%), True)
			End If
	End Select
End Function
	

