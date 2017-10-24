Public Const PLOGERROR_PIN_POSITION = -2147220987

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

Public Type UnusedInfoInPkg
	Name As String
	unusedGatesLst() As objects
	bAllPinsDisconnected() As Boolean
	bHaveUnusedPins() As Boolean
End Type

Public Sub GetUnusedInfoInPkg(doc As Object, ByRef unGatesInPkg() As UnusedInfoInPkg)
	Set allComp = doc.Components
	ReDim unGatesInPkg(0 To allComp.Count)
	MyCount = 0
	For Each nextComp In allComp
		Set usedGates = nextComp.Gates
		Set unusedGates = nextComp.UnusedGates
		totalCount = usedGates.Count + unusedGates.Count + 1 ' additional 1 for unused pins

		' get current package		
		bFound = False
		For i=1 To myCount
			If unGatesInPkg(i).Name = nextComp.PartType Then
				bFound=True
				Exit For
			End If
		Next i
		If bFound=False Then
			myCount = myCount + 1			
			curPkg = myCount
			lastIdx = 0
			
			unGatesInPkg(curPkg).Name = nextComp.PartType			
			ReDim unGatesInPkg(curPkg).unusedGatesLst(0 To totalCount) As objects
			ReDim unGatesInPkg(curPkg).bAllPinsDisconnected(0 To totalCount) As Boolean
			ReDim unGatesInPkg(curPkg).bHaveUnusedPins(0 To totalCount) As Boolean			
		Else
			curPkg = i
			lastIdx = UBound(unGatesInPkg(curPkg).unusedGatesLst)
			ReDim Preserve unGatesInPkg(curPkg).unusedGatesLst(0 To lastIdx + totalCount) As objects
			ReDim Preserve unGatesInPkg(curPkg).bAllPinsDisconnected(0 To lastIdx + totalCount) As Boolean
			ReDim Preserve unGatesInPkg(curPkg).bHaveUnusedPins(0 To lastIdx + totalCount) As Boolean			
		End If
								
		'	if unused gates existed,
		'	then this gate adding to list unused gates
		For i = 1 To unusedGates.Count
			lastIdx = lastIdx + 1
			Set unGatesInPkg(curPkg).unusedGatesLst(lastIdx) = unusedGates(i)
			unGatesInPkg(curPkg).bAllPinsDisconnected(lastIdx) = True
			unGatesInPkg(curPkg).bHaveUnusedPins(lastIdx) = False			
		Next i
		
		'	check used gates
		'	if all (or any) pins in used gate are not connected, 
		'	then this gate adding to list unused gates		
		For i = 1 To usedGates.Count
			bConnected = False
			bAnyDisconnected = False
			For Each nextPin In usedGates(i).Pins
				If Not (nextPin.Net Is Nothing) Then 
					bConnected = True
				ElseIf Not(nextPin.Gate Is Nothing) Then
					bAnyDisconnected = True
				End If
			Next nextPin
			
			If bConnected = False Then	' if all pins in current gate are not connected, then
				lastIdx = lastIdx + 1
				Set unGatesInPkg(curPkg).unusedGatesLst(lastIdx) = usedGates(i)
				unGatesInPkg(curPkg).bAllPinsDisconnected(lastIdx) = True
				unGatesInPkg(curPkg).bHaveUnusedPins(lastIdx) = False
			ElseIf bAnyDisconnected Then	' if any pin in gate is not connected
				lastIdx = lastIdx + 1
				Set unGatesInPkg(curPkg).unusedGatesLst(lastIdx) = usedGates(i)
				unGatesInPkg(curPkg).bAllPinsDisconnected(lastIdx) = False
				unGatesInPkg(curPkg).bHaveUnusedPins(lastIdx) = False
			End If
		Next i
	
		'	check component pins
		'	if pin in component is not belong to any gate and is not signal or power, 
		'	then first gate of component adding to list unused gates
		For Each nextPin In nextComp.Pins
			'	adding first gate of component to list of unused gates
			If (nextPin.Gate Is Nothing) And (nextPin.Net Is Nothing) Then
				lastIdx = lastIdx + 1
				Set unGatesInPkg(curPkg).unusedGatesLst(lastIdx) = usedGates(1)
				unGatesInPkg(curPkg).bAllPinsDisconnected(lastIdx) = False
				unGatesInPkg(curPkg).bHaveUnusedPins(lastIdx) = True
				Exit For
			End If
		Next nextPin
		
		If lastIdx = 0 Then
			MyCount = MyCount - 1
		Else
			ReDim Preserve unGatesInPkg(curPkg).unusedGatesLst(0 To lastIdx) As objects
			ReDim Preserve unGatesInPkg(curPkg).bAllPinsDisconnected(0 To lastIdx) As Boolean
			ReDim Preserve unGatesInPkg(curPkg).bHaveUnusedPins(0 To lastIdx) As Boolean
		End If
	Next nextComp
	ReDim Preserve unGatesInPkg(0 To MyCount) As UnusedInfoInPkg	
End Sub

Public Sub GetPackageList(doc As Object, ByRef pkgList() As String, _
                                                       Optional bSort As Boolean = False)
	MyCount = 0
	Set comps = doc.Components
	ReDim pkgList(1 To comps.Count) As String
	For Each nextComp In comps
		bFound = False
		For i=1 To myCount
			If pkgList(i) = nextComp.PartType Then
				bFound=True
				Exit For
			End If
		Next i
		If bFound=False Then
			myCount = myCount+1
			pkgList(myCount) = nextComp.PartType
		End If
	Next nextComp
	If (MyCount = 0) Then
		ReDim  pkgList(0) As String
	Else
		ReDim Preserve pkgList(1 To MyCount) As String
	End If

	'Sort Package List
	If bSort Then
		lo = 1
		hi = MyCount
	    While hi > lo
	        max = lo
	        For p = lo+1 To hi
	            If StrNumCmp(pkgList(p), pkgList(max)) > 0 Then
	                max = p
	            End If
			Next
			'swap max with hi
			Dim temp As String
			temp = pkgList(max)
			pkgList(max) = pkgList(hi)
			pkgList(hi) = temp
			hi = hi - 1
	    Wend
	End If
End Sub

Public Function PinTyp(aPin As Object) As String
		Select Case aPin.ElectricalType
			Case plogElectricalTypeUnknown 				
					PinTyp = "U"
			Case plogElectricalTypeSource
					PinTyp = "S"
			Case plogElectricalTypeBidirectional
					PinTyp = "B"
			Case plogElectricalTypeOpenCollector
					PinTyp = "C"
			Case plogElectricalTypeOrTieableSource
					PinTyp = "O"
			Case plogElectricalTypeTristate
					PinTyp = "T"
			Case plogElectricalTypeLoad
					PinTyp = "L"
			Case plogElectricalTypeTerminator
					PinTyp = "Z"
			Case plogElectricalTypePower
					PinTyp = "P"
			Case plogElectricalTypeGround
					PinTyp = "G"
			Case Else 
					PinTyp = "U"
		End Select	
End Function

Public Function SheetNumOrName(aSheet As Object) As String
	sname = aSheet.Name
	If Left(sname, 6) = "Sheet " Then
		sname = Right(sname, Len(sname) - 6)
	End If
	SheetNumOrName = sname
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
