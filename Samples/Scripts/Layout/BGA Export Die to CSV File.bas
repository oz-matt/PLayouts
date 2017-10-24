' Export DIE Component to CSV format

Public Function DesignUnits () As String
	Select Case ActiveDocument.unit
		Case ppcbUnitMils 				
				DesignUnits = "Mil"
		Case ppcbUnitInch 				
				DesignUnits = "Inch"
		Case ppcbUnitMetric 				
				DesignUnits = "MM"
		Case Else 
				DesignUnits = "Mil"
	End Select	
End Function

Public Sub ExportDieToCSV (die_comp, output_file_name)
	If (die_comp.IsDiePart) Then

		out_file = 1
		Open output_file_name For Output As #out_file
		Print #out_file, DesignUnits(); ",,,,,"
	
		Dim pad As CBP
		For Each pad In die_comp.CBPs
			is_round_pad = False
			If (pad.Shape = ppcbBondPadShapeOval) Then
				If (pad.Width = pad.Length) Then
					is_round_pad = True
				End If
			End If
			
			pad_name = pad.Name
			' Cut off leading component name and CBP prefix
			pad_name = Mid (pad_name, InStr(pad_name, ".CBP") + 4, Len(pad_name) )
			
			If (is_round_pad) Then
				'Name,Function,X,Y,radius
				Print #out_file, pad_name; ","; pad.Function; ","; pad.PositionX; ","; pad.PositionY; ","; pad.Length
			Else
				'Name,Function,X,Y,length,width
				Print #out_file, pad_name; ","; pad.Function; ","; pad.PositionX; ","; pad.PositionY; ","; pad.Length; ","; pad.Width
			End If
		Next pad
		Close #out_file
	End If
End Sub

Sub Main
	With ActiveDocument
		
		'Prepare list of all Die Components in design
		ReDim die_comps(.Components.Count)
		ReDim die_comp_names$(.Components.Count)
		
		n_die_comp = 0
		For Each comp In. Components
			If (comp.IsDiePart) Then
				die_comps(n_die_comp) = comp
				die_comp_names$(n_die_comp) = comp.Name
				n_die_comp = n_die_comp + 1
			End If
		Next comp
	End With
		
	ReDim Preserve die_comps(n_die_comp)
	ReDim Preserve die_comp_names$(n_die_comp)
		
	If (n_die_comp = 0) Then
		MsgBox "There are no Die Components in design"
		Exit All
	End If

	die_comp_id = 0
	If (n_die_comp > 0) Then ' Always show dialog
		' Ask witch Die Component should be exported
	    Begin Dialog UserDialog 180,95, "Export to CSV"
	        Text 10,10,180,15,"Die Component for Export:"
	        DropListBox 10,25,160,60,die_comp_names$(),.die_comp_id
	        OKButton 10,60,70,20
	        CancelButton 100,60,70,20
	    End Dialog
	    Dim dlg As UserDialog
	    dlg.die_comp_id = 0		' Choose first component by default
		    
	    If (Dialog(dlg) <> -1) Then
	    	' Dialog is cancelled
	    	Exit All
	    End If
	    die_comp_id = dlg.die_comp_id
	End If		
	
	' Export selected Die Component to CSV file
	out_file_name$ = DefaultFilePath & "\" & die_comps(die_comp_id).PartType & ".csv"
	
	' Prompt to overwrite existing file
	On Error GoTo FileNotExist	
	Open out_file_name$ For Input As #1
	Close #1
	If MsgBox("File '" & out_file_name$ & "' exists. Overwrite ?", vbOkCancel) = vbCancel Then
		GoTo ByeBye
	End If

FileNotExist:	
	
	On Error GoTo FileInUse
	ExportDieToCSV die_comps(die_comp_id), out_file_name$ 
	
	' Start Excel and loads the text file
	On Error GoTo noExcel
	' See resutls in Excel
	Dim excelApp As Object
	Set excelApp = CreateObject("Excel.Application")
	On Error GoTo 0
	excelApp.Visible = True
	excelApp.Workbooks.OpenText 	FileName:= out_file_name$

	excelApp.Columns("B:B").Select
	With excelApp.Selection
		.ColumnWidth = 20 ' Make the second column a little wider
	End With
		
	excelApp.Range("A1").Select
	Set excelApp = Nothing
	End
		
noExcel:		
	' See results in Notepad		
	Shell "Notepad " & out_file_name$, 1 
	End
	
FileInUse:
	MsgBox "Unable to create '" & out_file_name$ & "'"
	Exit All
ByeBye:	
	
End Sub

