Sub Main
	Begin Dialog UserDialog 240,98,"Add Part" ' %GRID:10,7,1,1
		Text 10,14,70,14,"&Part Type:",.Text1
		TextBox 90,7,140,21,.PartTypeName
		Text 10,42,70,14,"&Cost:",.Text2
		TextBox 90,35,140,21,.AttributeCost
		OKButton 10,70,100,21
		CancelButton 130,70,100,21
	End Dialog
	Dim dlg As UserDialog
	
	If Dialog(dlg) <> 0 And dlg.PartTypeName <> "" Then
		Set doc = ActiveDocument
		Set newGate = doc.ActiveSheet.AddGate(dlg.PartTypeName)
		Set part = newGate.Component
		If part.Attributes("Cost") Is Nothing Then
			part.Attributes.Add "Cost", dlg.AttributeCost
		Else
			part.Attributes("Cost") = dlg.AttributeCost
		End If
				
		doc.SelectObjects(,,False)	'Unselect all
		newGate.Selected = True		'select only new gate
		
		ProcessCommand 561 'Run Move command for selected object
	End If

End Sub
