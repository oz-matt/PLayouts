Const outCh = 499

Const maxColumnsArray = 50

Dim fileName As String

Dim currentLengthOfString As Integer

Dim currentIsColumn As Integer
Dim IsColumn(2 * maxColumnsArray) As Boolean

Dim firsttimeColumnsPassed As Boolean
Dim currentColumns As Integer
Dim currentColumn As Integer
Dim columnsPosition(1 To maxColumnsArray, 1 To maxColumnsArray) As Integer

Dim firsttimeBeetwensPassed As Boolean
Dim currentBetweens As Integer
Dim currentBeetwenColumn As Integer
Dim betweenColumns(1 To maxColumnsArray) As Integer

Dim maxColumns(1 To maxColumnsArray) As Integer


Public Function LayerCnt(doc As Object)
	LayerCnt = 0
	For i = 1 To doc.LayerCount
		If doc.LayerType(i) = ppcbLayerComponent Or _
		  doc.LayerType(i) = ppcbLayerRouting Then
			LayerCnt = LayerCnt + 1
		End If
	Next
End Function

Public Type Package
	Name As String
	Description As String
	LogicType As String
End Type

Public Sub GetPackageList(doc As Object, ByRef pkgList() As Package, _
                                                       Optional bSort As Boolean = False)
	MyCount = 0
	Set comps = doc.Components
	ReDim pkgList(comps.Count) As Package
	For Each nextComp In comps
		bFound = False
		For i=1 To myCount
			If pkgList(i).Name = nextComp.PartType Then
				bFound=True
				Exit For
			End If
		Next i
		If bFound=False Then
			myCount = myCount+1
			pkgList(myCount).Name = nextComp.PartType
			Set desc =  nextComp.PartTypeAttributes("Description")
			If Not desc Is Nothing Then
				pkgList(myCount).Description = desc
			End If
			pkgList(myCount).LogicType = nextComp.PartTypeLogic
		End If
	Next nextComp
	If (MyCount = 0) Then
		ReDim  pkgList(0) As Package
	Else
		ReDim Preserve pkgList(MyCount) As Package
	End If
	
	'Sort Package List
	If bSort Then
			lo = 1
			hi = MyCount
		    While hi > lo
		        max = lo
		        For p = lo+1 To hi
		            If StrNumCmp(pkgList(p).Name, pkgList(max).Name) > 0 Then
		                max = p
		            End If
				Next
				'swap max with hi
				Dim temp As Package
				temp = pkgList(max)
				pkgList(max) = pkgList(hi)
				pkgList(hi) = temp
				hi = hi - 1
		    Wend
	End If
End Sub

Public Function PkgCnt(doc As Object) As Integer
	Dim pkgs() As Package
	GetPackageList doc, pkgs, False  'Get unsorted list of packages
	PkgCnt = UBound(pkgs)
End Function

Public Function SymCnt(doc As Object) As Integer
	myCount = 0
	Set comps = doc.Components
	Dim symList() As String
	ReDim symList(comps.Count) As String
	For Each nextComp In comps
		bFound = False
		decals = nextComp.DecalCompatibleList
		For i=1 To UBound(decals)
			For j=0 To myCount - 1
				If symList(j) = decals(i) Then
					bFound=True
					Exit For
				End If
			Next j
			If bFound=False Then
				symList(myCount) = decals(i)
				myCount = myCount+1
			End If
		Next i
	Next nextComp
	SymCnt = myCount
End Function


Public Function EquivalentIC(doc As Object) As Double
	pinCount = doc.Pins.Count
	EquivalentIC = pinCount / 14
End Function

Public Function BdDensity(doc As Object) As Double
	boSurface = doc.BoardOutlineSurface
	BdDensity = boSurface / EquivalentIC(doc)
End Function

Public Function PSigCnt(doc As Object) As Integer
	myCount = 0
	
	For Each nextNet In doc.Nets
		If nextNet.Power = True Then myCount = myCount+1
	Next nextNet
	PSigCnt = myCount
End Function

Public Function SSigCnt(doc As Object) As Integer
	myCount = 0
	
	For Each nextNet In doc.Nets
		If nextNet.Power = False Then myCount = myCount+1
	Next nextNet
	SSigCnt = myCount
End Function

Public Function TopCompCnt(doc As Object) As Integer
	myCount = 0
	
	For Each nextComp In doc.Components
		If nextComp.layer = 1 Then myCount = myCount+1
	Next nextComp
	TopCompCnt = myCount
End Function

Public Function BottomCompCnt(doc As Object) As Integer
	myCount = 0

	For Each nextComp In doc.Components
		If nextComp.layer = LayerCnt(doc) Then 
			myCount = myCount+1
		End If
	Next nextComp
	BottomCompCnt = myCount
End Function

Public Function PadCnt(doc As Object) As Integer
	PadCnt = doc.Pins.Count
End Function

Public Function DrPadCnt(doc As Object) As Integer
	myCount = 0
	
	For Each nextPin In doc.Pins
		If nextPin.DrillSize <> 0.0 Then myCount = myCount+1
	Next nextPin
	
	DrPadCnt = myCount
End Function

Public Function NDPadCnt(doc As Object) As Integer
	myCount = 0
	
	For Each nextPin In doc.Pins
		If nextPin.DrillSize = 0.0 Then myCount = myCount+1
	Next nextPin
	
	NDPadCnt = myCount
End Function

Public Function TestPoints(doc As Object) As objects
	'make empty output collection
	Set tps = doc.GetObjects(ppcbObjectTypeUnknown)
	'Fill with pins and vias sorted by net
	For Each nextNet In doc.Nets
		'Pins go first
		For Each nextPin In nextNet.Pins
			If nextPin.TestPoint <> ppcbTestPointNone Then
				tps.Add nextPin
			End If
		Next
		'Vias go next
		For Each nextVia In nextNet.Vias
			If nextVia.TestPoint <> ppcbTestPointNone Then
				tps.Add nextVia
			End If
		Next
	Next
	'Add unused component pins go finally
	For Each nextPin In doc.Pins
		If nextPin.Net Is Nothing And nextPin.TestPoint <> ppcbTestPointNone Then
			tps.Add nextPin
		End If
	Next
	Set TestPoints = tps
End Function

Public Function NetPinsSortedByConnection(aNet As Object) As objects
	'make empty output collection
	Set netPins = aNet.Parent.GetObjects(ppcbObjectTypeUnknown)
	For Each nextCon In aNet.Connections
		For i = 1 To 2 
		   'find pin by name to avoid duplication in output pin collection
			If netPins(nextCon.Pins(i).Name) Is Nothing Then
				netPins.Add nextCon.Pins(i)
			End If
		Next i
	Next
	Set NetPinsSortedByConnection = netPins
End Function

Public Function TPPinCnt(Net As Object) As Integer
	myCount = 0
	
	For Each nextPin In Net.Pins
		If nextPin.TestPoint <> ppcbTestPointNone Then
			myCount = myCount+1
		End If
	Next nextPin
	TPPinCnt = myCount
End Function

Public Function TPViaCnt(Net As Object) As Integer
	myCount = 0

	For Each nextVia In Net.Vias
		If nextVia.TestPoint <> ppcbTestPointNone Then
			myCount = myCount+1
		End If
	Next nextVia
	TPViaCnt = myCount
End Function

Public Function TPAssigned(Net As Object) As Boolean
	myCount = TPPinCnt(Net)+TPViaCnt(Net)
	
	If myCount>0 Then 
		TPAssigned = True
	Else 
		TPAssigned = False
	End If

End Function

Public Function TPCnt(Net As Object) As Integer
	myCount = TPPinCnt(Net)+TPViaCnt(Net)
	TPCnt = myCount
End Function

Public Function PinTyp(aPin As Object) As String
		Select Case aPin.ElectricalType
			Case ppcbElectricalTypeUnknown 				
					PinTyp = "U"
			Case ppcbElectricalTypeSource
					PinTyp = "S"
			Case ppcbElectricalTypeBidirectional
					PinTyp = "B"
			Case ppcbElectricalTypeOpenCollector
					PinTyp = "C"
			Case ppcbElectricalTypeOrTieableSource
					PinTyp = "O"
			Case ppcbElectricalTypeTristate
					PinTyp = "T"
			Case ppcbElectricalTypeLoad
					PinTyp = "L"
			Case ppcbElectricalTypeTerminator
					PinTyp = "Z"
			Case ppcbElectricalTypePower
					PinTyp = "P"
			Case ppcbElectricalTypeGround
					PinTyp = "G"
			Case Else 
					PinTyp = "U"
		End Select	
End Function

Public Function Format2(value As Double) As String
	'Remove trailing zeros and floating points
	s = Format(value, "0.#####")
	l = Len(s)
	If Right(s,1) = "." Then 
		s = Left(s, l-1)
		Format2 = s
		Exit Function
	End If
	While Right(s,1) = "0" 
		s = Left(s, l-1)
		l = Len(s)
	Wend
	Format2 = s
End Function

Public Function GetTime As String
	GetTime = Format (Now, "ddd mmm dd hh:mm:ss yyyy")
End Function

Public Function OpenReport(file As String) As Integer
	fileName = file
	Open fileName For Output As #outCh
	OpenReport = outCh
End Function

Public Sub CloseReport(Optional bShowReport As Boolean = True)
	Close #outCh
	If bShowReport Then Shell "Notepad " & fileName, 1 ' See results by Notepad
End Sub

Public Sub Between(Optional betweenCol As Integer = 0)
	If firsttimeBeetwensPassed = False Then
		currentBetweens = 0
		firsttimeBeetwensPassed = True
	End If
	
	currentIsColumn = currentIsColumn + 1
	IsColumn(currentIsColumn) = False
	
	currentBetweens = currentBetweens + 1
	betweenColumns(currentBetweens) = betweenCol

	maxColumns(currentBetweens) = 0 
	currentBeetwenColumn = 0
	currentLengthOfString = 0
End Sub

Public Sub MaxCols(cols As Integer)
	maxColumns(currentBetweens) = cols
End Sub

Public Sub End_Between()
	currentBetweens = currentBetweens - 1
	If currentBetweens < 0 Then currentBetweens = 0
	currentBeetwenColumn = 0
	
	currentIsColumn = currentIsColumn - 1
	If currentIsColumn < 0 Then currentIsColumn = 0
	
	If currentLengthOfString <> 0 Then 
		currentLengthOfString = 0
		Print #outCh
	End If	
End Sub

Public Sub Columns(ParamArray formatParam())
	If firsttimeColumnsPassed = False Then
		currentColumns = 0
		firsttimeColumnsPassed = True
	End If
	
	If LBound(formatParam) > UBound(formatParam) Then Exit Sub 	'the ParamArray didn't receive any expressions. 
	
	If currentLengthOfString <> 0 Then Print #outCh
	
	currentIsColumn = currentIsColumn + 1
	currentColumn = 1
	IsColumn(currentIsColumn) = True
	
	currentColumns = currentColumns + 1
	For i = 0 To UBound(formatParam)
		columnsPosition(currentColumns, i+1) = formatParam(i)
	Next i
	columnsPosition(currentColumns, i+1) = -1
	
	currentLengthOfString = 0
End Sub

Public Sub End_Columns()
	currentColumns = currentColumns - 1
	If currentColumns < 0 Then currentColumns = 0
	currentColumn = 1
	
	currentIsColumn = currentIsColumn - 1
	If currentIsColumn < 0 Then currentIsColumn = 0
	
	If currentLengthOfString <> 0 Then 
		currentLengthOfString = 0
		Print #outCh
	End If
End Sub

Public Sub Out(ParamArray formatParam())
	If LBound(formatParam) > UBound(formatParam) Then 	'the ParamArray didn't receive any expressions. 
		Print #outCh
		Exit Sub 
	End If
	
	' we want just output string
	If currentBetweens = 0 And currentColumns = 0 Then
		For i = LBound(formatParam) To UBound(formatParam)
			Print #outCh,  formatParam(i);
		Next i
		Print #outCh
		Exit Sub
	End If
	
	If IsColumn(currentIsColumn) = True Then
		For i = LBound(formatParam) To UBound(formatParam)

			formatParam(i) = Trim(formatParam(i))
			
			If currentLengthOfString >= columnsPosition(currentColumns, currentColumn) Then 
				Print #outCh, formatParam(i);
				currentLengthOfString = currentLengthOfString + Len(formatParam(i))
			Else 
				Print #outCh, Space(columnsPosition(currentColumns, currentColumn)  - currentLengthOfString); formatParam(i);
				currentLengthOfString = currentLengthOfString + Len(Space(columnsPosition(currentColumns, currentColumn)  - currentLengthOfString)) + Len(formatParam(i))
			End If
			
			currentColumn = currentColumn + 1			
			If columnsPosition(currentColumns, currentColumn) = -1 Then
				currentColumn = 1
				currentLengthOfString = 0
				Print #outCh
			End If
		Next i
	Else
		For i = LBound(formatParam) To UBound(formatParam)
		
			formatParam(i) = Trim(formatParam(i))
			
			If currentLengthOfString >= currentBeetwenColumn * betweenColumns(currentBetweens) Then 
				Print #outCh, formatParam(i);
				currentLengthOfString = currentLengthOfString + Len(formatParam(i))
			Else 
				Print #outCh, Space(currentBeetwenColumn * betweenColumns(currentBetweens)  - currentLengthOfString); formatParam(i);
				currentLengthOfString = currentLengthOfString + Len(Space(currentBeetwenColumn * betweenColumns(currentBetweens)  - currentLengthOfString)) + Len(formatParam(i))
			End If
			
			currentBeetwenColumn = currentBeetwenColumn + 1			
			If maxColumns(currentBetweens) <> 0 And currentBeetwenColumn >= maxColumns(currentBetweens) Then
				currentBeetwenColumn = 0
				currentLengthOfString = 0
				Print #outCh
			End If		
		Next i
	End If
End Sub

