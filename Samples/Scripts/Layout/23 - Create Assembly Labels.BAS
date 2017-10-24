' Sample 23: Create Assembly Labels.BAS
' 
' This sample demontsrates how to create new labels in PADS Layout
' It creates new Ref. Des. and Part Type labels for all components in the design
' on assembly drawing layer, centered.
'
' For more details, please refer to the PADS Layout Basic Editor Help File.
'


Sub Main
	For layer = 1 To ActiveDocument.LayerCount
		If ActiveDocument.LayerType(layer) = ppcbLayerAssembly Then
			Exit For
		End If
	Next layer

	For Each comp In ActiveDocument.Components
	    x = comp.CenterX
	    y = comp.CenterY
	    Set lab = comp.AddLabel(ppcbLabelTypeRefDesignator, , layer)
	    lab.HorzJustification = ppcbJustifyHCenter
	    lab.VertJustification = ppcbJustifyBottom
	    lab.PositionX(, ppcbOriginTypeDesign) = x
	    lab.PositionY(, ppcbOriginTypeDesign) = y
		lab.RightReading = ppcbRightReadingOrthogonal
		Set lab = comp.AddLabel(ppcbLabelTypePartType, , layer)
	    lab.HorzJustification = ppcbJustifyHCenter
	    lab.VertJustification = ppcbJustifyTop
	    lab.PositionX(, ppcbOriginTypeDesign) = x
	    lab.PositionY(, ppcbOriginTypeDesign) = y
		lab.RightReading = ppcbRightReadingOrthogonal
	Next comp
End Sub

