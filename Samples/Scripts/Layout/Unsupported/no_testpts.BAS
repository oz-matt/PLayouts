Dim ListNets$(5000)
Dim ListPins$(10000)
Sub Main
	index = 0
	index1 = 0
	NetCount = 0
	PinCount = 0
	b = ""
	Search_Text$ = "Begin Scanning Nets"
	Dim nextNet As Object
	Dim nextPin As Object
	
	Begin Dialog UserDialog 270,70," Nets with no Test Points" '%GRID:10,7,1,1
		OKButton 30,42,90,21
		Text 70,14,140,14,Search_Text$,.Text1
		CancelButton 150,42,90,21
	End Dialog
	Dim dlg1 As UserDialog
	Dialog dlg1
	Mousepointer = 11
	DBFileName$ = ActiveDocument.FullName
	For Each nextNet In ActiveDocument.GetObjects(2, "", False)
		'loading the array with net names in the current database
		ListNets$(index) = nextNet.Name
		index = index + 1
	Next nextNet
	' Go through each via in the design and get net names
	Dim nextVia As Object
	NetCount= index - 1 'arrays start at 0 not 1
	index = 0
	For Each nextVia In ActiveDocument.GetObjects(4, "", False)

		' Get Via Attached Net Property
		a = nextVia.Net.Name
		While index <= NetCount
		'begin searching thru array starting at the first (0) location
			If a = ListNets$(index) Then
				'If Via Net Name matches existing net name, move net names
				'from this point forward in ListNets$() up one overwriting the
				'existing net name at this location
				index1 = index
				'compare current via net name with every net name in array one
				'location at a time
				While index1 <= NetCount
					ListNets$(index1) = ListNets$(index1 + 1)
					index1 = index1 + 1
				Wend
				index1 = 0
				'this net name may exist on other vias, but we don't care if it does
				'so get next net name by setting index to make this loop think that all
				'nets have been checked
				index = NetCount - 1
				'the number of nets in the array has now been decreased by one
				NetCount = NetCount - 1
			Else
				'Otherwise no via net name matches exiting nets in ListNets$()
				'because it has been already removed. Get next array location.
			End If
			index = index + 1
		Wend
	'reset array counter to top because we are using a new via net name
	index = 0
	Next NextVia
	index = 0
	For Each nextPin In ActiveDocument.GetObjects(3, "", False)
	'Get Pin Attached Net Property
	If (Not(nextPin.Net Is Nothing))  Then 
		b = nextPin.Net
		
		If nextPin.IsSMD = False Then 
			While index <= NetCount
				'begin searching thru array starting at the first (0) location
				If b = ListNets$(index) Then
					'If Pin Net Name matches existing net name, move net names
					'from this point forward in ListNets$() up one overwriting the
					'existing net name at this location
					index1 = index
					'compare current pin net name with every via net name in array one
					'location at a time
					While index1 <= NetCount
						ListNets$(index1) = ListNets$(index1 + 1)
						index1 = index1 + 1
					Wend
					index1 = 0
					'the number of nets in the array has now been decreased by one
					NetCount = NetCount - 1
				Else
					'Otherwise no pin net name matches exiting nets in ListNets$()
					'because it has been already removed. Get next array location.
				End If
				index = index + 1
			Wend
			'reset array counter to top because we are using a new pin net name
			index = 0
			Else
		End If
	Else
	End If
	Next nextPin
	
	'Output a .txt file listing nets without testpoints
	FileName = Path & "tstptnet.txt"
	Open filename For Output As #1
	' Output Header
	Print #1, NetCount+1;" Net Names without a testpoint"
	Print #1,  "for database "; DBFileName$
	Print #1
		'Output net names
		For outindex = 0 To NetCount
			Print #1, ListNets$(outindex)
		Next outindex
	Close #1
	
	Begin Dialog UserDialog 270,392," Nets with no Test Points",.CallbackFunc '%GRID:10,7,1,1
		OKButton 20,357,230,21
		ListBox 30,21,210,259,ListNets(),.ListBox2
		Text 40,7,210,14,"Click on Nets to select them",.Text2
		Text 20,280,250,14,"Number of Nets without Testpoints",.Text3
		Text 120,294,40,14,Str$(NetCount+1),.Text4
		Text 20,322,230,14,"   File TSTPTNET.TXT generated",.Text1
	
	End Dialog
Dim dlg As UserDialog
Dialog dlg

End Sub

Function CallbackFunc%(DlgItem$, Action%, SuppValue%)
	Select Case Action%
		Case 2 ' Value changing or button pressed
			If DlgItem$ = "ListBox2" Then
				ActiveDocument.SelectObjects(9999, "", False)
				ActiveDocument.SelectObjects(2, ListNets(SuppValue%), True)
			End If
	End Select
End Function
