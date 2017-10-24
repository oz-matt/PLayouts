' Eliminate the painstaking task of finding that illusive hole size.  Simply run this VB script to generate
' seperate lists of all the different plated, non-plated, and via hole sizes used in a design.  Then click 
' on a hole size within any of these lists to select them.
' Written by Paul Woolford of Panacea Software part of Hydra Holdings PLC 
' Email: paulw_hydra@kbnet.co.uk
' Date: 03/06/98

Dim ListPltPins$(10000)
Dim ListNPltPins$(10000)
Dim ListVias$(10000)
Sub Main
	IndexV = 0 ' Set via list pointer to the top
	LastVia = "empty" ' Set LastVia to a dummy value
	Dim NextVia As Object
	For Each NextVia In ActiveDocument.GetObjects(4, "", False) ' Search through all vias in design
		Found = False
		Size = CStr(NextVia.DrillSize(0)) ' Avoid multiple calls to the CStr function
		If LastVia <> Size Then ' Check that next via is not the same as the last via
			For I = 0 To IndexV Step 1 ' Search through all vias within the list
				If (Size = ListVias$(I)) Then Found = True ' Set Found to be True if next via is already in the list 	
			Next I	
				If (Not Found) Then ' Make sure that next via was not already in the list
					ListVias$(IndexV) = Size ' Add next via to the list
					LastVia = Size ' Set LastVia to the last via to be added to the list
					IndexV = IndexV + 1 ' Increment the via list pointer		
				End If 		
		End If		
	Next NextVia
	
	IndexP = 0 ' Set plated list pointer to the top
	IndexNP = 0 ' Set non-plated list pointer to the top
	LastP = "empty" ' Set LastP to a dummy value
	LastNP = "empty" ' Set LastNP to a dummy value
			
	Dim NextPin As Object
	For Each NextPin In ActiveDocument.GetObjects(3, "", False) ' Search through all pins in design
		Size = CStr(NextPin.DrillSize(0)) ' Avoid multiple calls to the CStr function
			If (Size > 0) Then ' Check that next pin is thru-hole
				Found = False
					If (NextPin.Plated = 1) Then ' Check that next pin is plated
						If (LastP <> Size) Then ' Check that next pin is not the same as the last plated pin
							For I = 0 To IndexP Step 1 ' Search through all pins within the plated list
								If Size = ListPltPins$(I) Then Found = True ' Set Found to be True if next pin is already in the plated list 
							Next I		
								If (Not Found) Then ' Make sure that next pin is not already in the plated list 
									ListPltPins$(IndexP) = Size ' Add next pin to the plated list 
									LastP = Size	' Set LastP to the last pin added to the plated list
									IndexP = IndexP + 1 ' Increment the plated list pointer			
								End If
						End If 		
					Else ' NextPin must be non-plated
						If (LastNP <> Size) Then ' Check that next pin is not the same as the last non-plated pin
							For I = 0 To IndexNP Step 1 ' Search through all pins within the non-plated list
								If (Size = ListNPltPins$(I)) Then Found = True ' Set Found to be True if next pin is already in the non-plated list 
							Next I
								If (Not Found) Then ' Make sure that next pin is not already in the non-plated list
									ListNPltPins$(IndexNP) = Size ' Add next pin to the non-plated list
									LastNP = Size ' Set LastNP to the last pin added to the non-plated list
								 	IndexNP = IndexNP + 1 ' Increment the non-plated list item counter  
								End If
						End If
					End If	
			End If	
	Next NextPin
	
	' Sorting the lists 
	
	Bubble = True ' Set Bubble to True i.e. List is to be sorted
		While Bubble = True ' While list is still being sorted
			For i = 0 To indexP - 2 ' Check all plated pin hole sizes in the list
				If CSng(ListPltPins$(i)) > CSng(ListPltPins$(i + 1)) Then ' Check that the indexed plated pin hole size is not larger than the next in the list 
					Temp = ListPltPins$(i) ' Set a temporary register equal to the indexed plated pin hole size 
					ListPltPins$(i) = ListPltPins$(i + 1) ' Change the indexed plated pin hole size to the next in the list 
					ListPltPins$(i + 1) = Temp ' Change the next plated pin hole size in the list to the previously indexed plated pin hole size
					CountP = CountP + 1 ' Increment the plated list item counter
				End If 
			Next i
				If CountP = 0 Then ' Check that the list has been sorted
					Bubble = False ' Set Bubble to False i.e. List is sorted
				End If
			CountP = 0 ' Reset the plated list item counter 
		Wend
	Bubble = True ' Set Bubble to True i.e. List is to be sorted
		While Bubble = True ' While list is still being sorted
			For i = 0 To indexNP - 2 ' Check all non-plated pin hole sizes in the list
				If CSng(ListNPltPins$(i)) > CSng(ListNPltPins$(i + 1)) Then ' Check that the indexed non-plated pin hole size is not larger than the next in the list 
					Temp = ListNPltPins$(i) ' Set a temporary register equal to the indexed non-plated pin hole size 
					ListNPltPins$(i) = ListNPltPins$(i + 1) ' Change the indexed non-plated pin hole size to the next in the list 
					ListNPltPins$(i + 1) = Temp ' Change the next non-plated pin hole size in the list to the previously indexed non-plated pin hole size
					CountNP = CountNP + 1 ' Increment the non-plated list item counter
				End If	  
			Next i
				If CountNP = 0 Then ' Check that the list has been sorted
					Bubble = False ' Set Bubble to False i.e. List is sorted
				End If
			CountNP = 0 ' Reset the non-plated list item counter 
		Wend
	Bubble = True ' Set Bubble to True i.e. List is to be sorted
		While Bubble = True ' While list is still being sorted
			For i = 0 To index - 2 ' Check all via hole sizes in the list
				If CSng(ListVias$(i)) > CSng(ListVias$(i + 1)) Then ' Check that the indexed via hole size is not larger than the next in the list 
					Temp = ListVias$(i) ' Set a temporary register equal to the indexed via hole size 
					ListVias$(i) = ListVias$(i + 1) ' Change the indexed via hole size to the next in the list 
					ListVias$(i + 1) = Temp ' Change the next via hole size in the list to the previously indexed via hole size
					Count = Count + 1 ' Increment the via list item counter
				End If	  
			Next i
				If Count = 0 Then ' Check that the list has been sorted
					Bubble = False ' Set Bubble to False i.e. List is sorted
				End If
			Count = 0 ' Reset the via list item counter 
		Wend		
			
		ActiveDocument.SelectObjects(9999, "", False) ' Deselect all PowerPCB objects
	
	Begin Dialog UserDialog 370,203,"Select By Drill Size",.CallbackFunc ' %GRID:10,7,1,1
		Text 250,28,100,14,"Vias:",.Text2
		Text 10,28,100,14,"Plated Pins:",.Text3
		ListBox 250,42,110,91,ListVias(),.ListBox2
		ListBox 10,42,110,91,ListPltPins(),.ListBox1
		PushButton 250,175,110,21,"Close",.PushButton1
		Text 10,7,350,14,"Click on Drill Size in a List to Select in Design.",.Text1
		ListBox 130,42,110,91,ListNPltPins(),.ListBox3
		Text 130,28,110,14,"Non-Plated Pins:",.Text4
		TextBox 10,147,230,21,.Selection
		Text 250,133,30,14,"Qty:",.Text5
		Text 10,133,110,14,"Last Selected:",.Text6
		TextBox 250,147,110,21,.Qty
		PushButton 10,175,230,21,"Deselect",.PushButton2
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg

End Sub
Function CallbackFunc%(DlgItem$, Action%, SuppValue%)
	Select Case Action%
		Case 1 ' Dialog box initialization
		DlgText "Selection", "Nothing" ' Set the Selection Dialog Text Box to initially read Nothing
		DlgText "Qty", "0" ' Set the Qty Dialog Text Box to initially read 0
		Case 2 ' Value changing or button pressed

	If DlgItem$ = "ListBox1" Then ' Check that the selected hole size is within the plated list
		DlgText "Selection", "Plated " & ListPltPins(SuppValue%) ' Change the Selection dialog text box to read Plated followed by the hole size selected within the plated list 
		DlgText "Qty", "Calculating..." ' Change the Qty dialog text box to read Calculating...
			Dim NextPin As Object
				For Each NextPin In ActiveDocument.GetObjects(3, "", False) ' Search through all pins in design
					If (CStr(NextPin.DrillSize(0)) = ListPltPins(SuppValue%)) Then ' Check that next pin has the same hole size as the one selected within the plated list
						If (NextPin.Plated = 1) Then 
							NextPin.Selected = True ' Select next pin in design
							CountP = CountP + 1 ' Increment the plated pin count
						End If	
					End If	
				Next NextPin
		DlgText "Qty", CountP ' Change the Qty dialog text box to display the number of plated pins with the Selected hole size		
	ElseIf DlgItem$ = "ListBox3" Then ' Check that the selected hole size is within the non-plated list
		DlgText "Selection", "Non-Plated " & ListNPltPins(SuppValue%) ' Change the Selection dialog text box to read Non-Plated followed by the hole size selected within the non-plated list 
		DlgText "Qty", "Calculating..." ' Change the Qty dialog text box to read Calculating...
				For Each NextPin In ActiveDocument.GetObjects(3, "", False) ' Search through all pins in design
					If (CStr(NextPin.DrillSize(0)) = ListNPltPins(SuppValue%)) Then '  Check that next pin has the same hole size as the one selected within the non-plated list
						NextPin.Selected = True ' Select next pin in design	
						CountNP = CountNP + 1	' Increment the non-plated pin count
					End If
				Next NextPin
		DlgText "Qty", CountNP	' Change the Qty dialog text box to display the number of non-plated pins with the Selected hole size														
	Else
		DlgText "Selection", "Via " & ListVias(SuppValue%) ' Change the Selection dialog text box to read Via followed by the hole size selected within the via list 		
		DlgText "Qty", "Calculating..." ' Change the Qty dialog text box to read Calculating...
			Dim NextVia As Object
				For Each NextVia In ActiveDocument.GetObjects(4, "", False) ' Search through all vias in design
					If (CStr(NextVia.DrillSize(0)) = ListVias(SuppValue%)) Then ' Check that next via has the same hole size as the one selected within the via list
						NextVia.Selected = True ' Select next via in design
						Count = Count + 1 ' Increment the via count
					End If
			Next NextVia
		DlgText "Qty", Count ' Change the Qty dialog text box to display the number of vias with the selected hole size	
	End If	
	
	If DlgItem$ = "PushButton2" Then
		ActiveDocument.SelectObjects(9999, "", False) ' Deselect all PowerPCB objects
		DlgText "Selection", "Nothing" ' Reset the Selection Dialog Text Box to read Nothing
		DlgText "Qty", "0" ' Reset the Qty Dialog Text Box to read 0
		CallbackFunc = True	
	End If

	End Select
	
End Function
