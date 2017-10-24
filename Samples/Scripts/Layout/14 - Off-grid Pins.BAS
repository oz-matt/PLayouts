' Sample 14: Off-grid Pins.BAS
' 
' This sample demontsrates the PADS Layout Grid OLE features.
' It lists all pins and vias that are off the current PADS Layout grid. Click on one
' listed pin or via will select it and pan the PADS Layout view to it.
'
' For more details, please refer to the PADS Layout Basic Editor Help File.
'

'Declare empty arrays
Dim ListPins$()
Dim ListVias$()

Sub Main
	totalPins = ActiveDocument.Pins.Count
	totalVias = ActiveDocument.Vias.Count

	'Allocate arrays component and net names
	ReDim ListPins$(totalPins)
	ReDim ListVias$(totalVias)
	
	indexPins = 0
	indexVias = 0
	gX = ActiveDocument.GridX(1)
	gY = ActiveDocument.GridY(1)

	Dim nextPin As Object
	For Each nextPin In ActiveDocument.Pins
		x = nextPin.PositionX
		y = nextPin.PositionY
		statX = ((x/gX) = Fix(x/gX))
		statY = ((y/gY) = Fix(y/gY))
		If (statX = False) Or (statY = False) Then 
			If (nextPin.IsSMD = False) Then
				ListPins$(indexPins) = nextPin.Name
				indexPins = indexPins + 1
			End If
		End If
	Next nextPin

	Dim nextVia As Object
	For Each nextVia In ActiveDocument.Vias
		x = nextVia.PositionX
		y = nextVia.PositionY
		statX = ((x/gX) = Fix(x/gX))
		statY = ((y/gY) = Fix(y/gY))
		If (statX = False) Or (statY = False) Then 
			ListVias$(indexVias) = nextVia.Name
			indexVias = indexVias + 1
		End If
	Next nextVia

	Begin Dialog UserDialog 300,266,"Off Grid Pins",.CallbackFunc ' %GRID:10,7,1,1
		ListBox 10,7,100,189,ListPins(),.ListBox1
		OKButton 10,238,280,21
		Text 10,196,280,14,indexPins & "/" & totalPins & " pins off (" & gX & ", " & gY & ") grid",.Text1
		Text 10,210,280,14,indexVias & "/" & totalVias & " vias off (" & gX & ", " & gY & ") grid",.Text2
		Text 10,224,280,14,"Click item to select and locate",.Text3
		ListBox 120,7,170,189,ListVias(),.ListBox2
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg
End Sub

Rem See DialogFunc help topic for more information.
Private Function CallbackFunc(DlgItem$, Action%, SuppValue%) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		'Unselect all objects
		ActiveDocument.SelectObjects(ppcbObjectTypeAll, "", False)
		If DlgItem$ = "ListBox1" Then
			Set nextPin = ActiveDocument.Pins(ListPins(SuppValue%))
			nextPin.Selected = True
			ActiveDocument.ActiveView.Pan(nextPin.PositionX, nextPin.PositionY)
		End If
		If DlgItem$ = "ListBox2" Then
			Set nextVia = ActiveDocument.Vias(ListVias(SuppValue%))
			nextVia.Selected = True
			ActiveDocument.ActiveView.Pan(nextVia.PositionX, nextVia.PositionY)
		End If
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem CallbackFunc = True ' Continue getting idle actions
	End Select
End Function

