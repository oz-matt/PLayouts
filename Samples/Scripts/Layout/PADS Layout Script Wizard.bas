''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' VB Report Wizard for PADS Layout                                               '
' Copyright (C) 2003 Mentor Graphics Corp.                                       '  
'                                                                                '
' VB Script Version 1.02                                                         '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'This Script stores all settings in the Registry database
'You can reset all settings by deleting of the following key in REGEDIT application:
'
'HKEY_CURRENT_USER\Software\VB and VBA Program Settings\PowerPCB 3.0 Custom Report Wizard v1.02

Option Explicit

Const WizardTitle	= "VB Script Wizard"
Const AppName		= "PowerPCB"
Const WizName		= AppName & " 3.0 Custom Report Wizard v1.02"
Const AppId			= 2	' 1 for PADS Logic; 2 for PADS Layout

Const Notepad = 0
Const WordPad = 1
Const Excel   = 2
Const Browser = 3

Const extMap = Array("rep", "rtf", "xls", "html")

Const contColl = "ActiveDocument.AssemblyOptions"
Const contName = Array("Assembly Option", "assembly option")
Const contNameProp = "GetOptName(opt)"
Const contObjs = Array("ActiveDocument", "opt")

Const formatName = Array("Text format." & vbCrLf & "'For better look, turn off 'Word Wrap' item in the Edit menu of Notepad and use Courier or any other fixed width font.", _
						 "Rich-text format (RTF)." & vbCrLf & "'For better look, choose 'No Wrap' option in the 'Rich Text Tab' of the Options dialog in WordPad", _
						 "Microsoft Excel Format.", _
						 "HTML format. You can then put that reports directly to your Internet/Intranet site.")

Enum OutType
	Simple = 0
	Attr   = 1
	RteSeg = 5
End Enum

Enum AlignType
	Lf = 0
	Rt = 1
	Cn = 2
End Enum

Enum PrimObjTypePPCB
	ppcbParts			= 0
	ppcbNets			= 1
	ppcbPins			= 2
	ppcbRouteSegments	= 3
	ppcbVias			= 4
	ppcbJumpers			= 5
	ppcbConnections		= 6
	ppcbPartTypes		= 7
	ppcbTestPoints		= 8
End Enum

Enum PrimObjTypePLogic
	plogParts			= 0
	plogGates			= 1
	plogNets			= 2
	plogPins			= 3
	plogPartTypes		= 4
End Enum

'Array of primary ojects
'			Name0,		VarName1,	Collection2,		 SortColl3,				ObjType4					Attr5	AO6		Index7, SecObjs8, NameProp9
Const primData = Array( _
	Array("Part",		"part",		"xxx.Components",	"xxx.Components",		ppcbObjectTypeComponent,	True,	True,	True,	True,	"xxx.Name"), _
	Array("Net",		"aNet",		"xxx.Nets",			"xxx.Nets",				ppcbObjectTypeNet,			True,	False,	True,	True,	"xxx.Name"), _
	Array("Pin", 		"aPin",		"xxx.Pins",			"GetSortedPins(xxx)",	ppcbObjectTypePin,			True,	False,	True,	True,	"xxx.Name"), _
	Array("Routed Segment","seg",	"xxx.RouteSegments","xxx.RouteSegments",	ppcbObjectTypeRouteSegment,	False,	False,	False,	True,	"OutRSegName(xxx)"), _
	Array("Via",	 	"aVia",		"xxx.Vias",			"xxx.Vias",				ppcbObjectTypeVia,			True,	False,	False,	True,	"OutViaName(xxx)"), _
	Array("Jumper",		"jmp",		"xxx.Jumpers",		"xxx.Jumpers",			ppcbObjectTypeJumper,		False,	True,	True,	False,	"xxx.Name"), _
	Array("Connection",	"con",		"xxx.Connections",	"xxx.Connections",		ppcbObjectTypeConnection,	False,	False,	True,	True,	"xxx.Name"), _
	Array("Part Type",	"pkg",		"xxx.PartTypes",	"xxx.PartTypes",		ppcbObjectTypePartType,		True,	True,	True,	True,	"xxx.Name"), _
	Array("Test Point",	"tPnt",		"GetTestPoints(xxx)","GetTestPoints(xxx,True)",	ppcbObjectTypeUnknown,	False,	False,	True,	False,	"xxx.Name"), _
)

'Array of object properties 
'		Property Name ID,			Column Name, 				Width, OutType, Align,	Function/Property template,	Present in both type report,	Reserved
Const primObjProps = Array( _
	Array( _
		Array("Name", 						"", 					 8, Simple,		Lf,	"xxx.Name",										  		1), _
		Array("Installed (Yes/No)", 		"Installed",	 		 9,	Simple,		Cn,	"Format(xxx.Installed, ""Yes/No"")",					0), _
		Array("Substituted (Yes/No)", 		"Substituted", 			11, Simple,		Cn,	"Format(xxx.Substituted, ""Yes/No"")",					0), _
		Array("Part Type", 					"", 					10, Simple,		Lf,	"xxx.PartType",											1), _
		Array("Default Part Type", 			"", 			 		17, Simple,		Lf,	"ActiveDocument.Components(xxx).PartType",				0), _
		Array("PCB Decal", 					"", 					10, Simple,		Lf,	"xxx.Decal",											1), _
		Array("Orientation", 				"", 					10, Simple,		Lf,	"xxx.orientation",										1), _
		Array("Logic Family", 				"Logic", 		 		 5, Simple,		Cn,	"xxx.PartTypeLogic",									1), _
		Array("Pin Count", 					"Pins", 		 		 5, Simple,		Rt,	"xxx.Pins.Count",										1), _
		Array("Power Pin Count",			"Power Pins", 	 		10,	Simple,		Rt,	"GetPowerPins(cont, xxx).Count",						1), _
		Array("Unconnected Pin Count",		"Uncon.Pins",	 		 6,	Simple,		Rt,	"GetUnconPins(xxx).Count",								1), _
		Array("Value",						"", 		  			 8,	Attr,		Lf,	"AttrVal(xxx, ""Value"")",								1), _
		Array("Tolerance",					"", 			 		 9,	Attr,		Lf,	"AttrVal(xxx, ""Tolerance"")",							1), _
		Array("Glued (Yes/No)", 			"Glued", 			 	 5,	Simple,		Cn,	"Format(xxx.Glued, ""Yes/No"")", 						1), _
		Array("SMD (Yes/No)", 				"SMD", 			 		 3,	Simple,		Cn,	"Format(xxx.IsSMD, ""Yes/No"")", 						1), _
		Array("ECO Registry (Yes/No)", 		"ECO", 			 		 3,	Simple,		Cn,	"Format(xxx.PartTypeECORegistered, ""Yes/No"")", 		1), _
		Array("Layer Name", 				"", 			 		30,	Simple,		Lf,	"ActiveDocument.LayerName(xxx.layer)",					1), _
		Array("Layer Number", 				"", 			 		12,	Simple,		Lf,	"xxx.layer", 											1), _
		Array("Position X", 				"", 			 		10,	Simple,		Rt,	"Format(xxx.PositionX, ""0.000"")",						1), _
		Array("Position Y", 				"", 			 		10,	Simple,		Rt,	"Format(xxx.PositionY, ""0.000"")",						1)  _
	), _
	Array( _
		Array("Name", 						"", 					 8, Simple,		Lf,	"xxx.Name",												1), _
		Array("Length", 					"", 			 		 8,	Simple,		Rt,	"Format(xxx.Length(False), ""0.#####;;0"")", 			1), _
		Array("Routed Length", 				"", 			 		13, Simple,		Rt,	"Format(xxx.Length(True), ""0.#####;;0"")", 			1), _
		Array("Unrouted Length", 			"", 			 		15, Simple,		Rt,	"Format(xxx.Length - xxx.Length(True), ""0.#####;;0"")",1), _
		Array("Connection Count", 			"Connections", 			11, Simple,		Rt,	"xxx.Connections.Count", 								1), _
		Array("Pin Count", 					"Pins", 		 		 5, Simple,		Rt,	"xxx.Pins.Count",										1), _
		Array("Via Count", 					"Vias", 			 	 5,	Simple,		Rt,	"xxx.Vias.Count", 										1), _
		Array("Power (Yes/No)", 			"Power", 		 		 5,	Simple,		Cn,	"Format(xxx.Power, ""Yes/No"")",						1)  _
	), _
	Array( _
		Array("Name", 						"", 					 8, Simple,		Lf,	"xxx.Name",												1), _
		Array("Net", 						"", 			 		15,	Simple,		Lf,	"ObjName(xxx.Net)",										1), _
		Array("Number", 					"", 			 		 6,	Simple,		Rt,	"xxx.Number", 											1), _
		Array("Drill Size", 				"", 			 		10,	Simple,		Rt,	"xxx.DrillSize",										1), _
		Array("Glued (Yes/No)", 			"Glued", 		 		 5,	Simple,		Cn,	"Format(xxx.Glued, ""Yes/No"")", 						1), _
		Array("SMD (Yes/No)", 				"SMD", 			 		 3,	Simple,		Cn,	"Format(xxx.IsSMD, ""Yes/No"")", 						1), _
		Array("Electrical Type", 			"", 			 		15,	Simple,		Lf,	"PinType(xxx)", 										1), _
		Array("Function Name", 				"", 			 		13,	Simple,		Lf,	"xxx.FunctionName", 									1), _
		Array("Test Point (Side/No)",		"Test Point", 			10,	Simple,		Lf,	"IIf (xxx.TestPoint = ppcbTestPointBottomLayer, ""Bottom"", IIf (xxx.TestPoint = ppcbTestPointTopLayer, ""Top"", ""No""))",	1), _
		Array("Position X", 				"", 			 		10,	Simple,		Rt,	"Format(xxx.PositionX, ""0.000"")", 					1), _
		Array("Position Y", 				"", 			 		10,	Simple,		Rt,	"Format(xxx.PositionY, ""0.000"")",						1)  _
	), _
	Array( _
		Array("Net",	 					"", 			 		15,	Simple,		Lf,	"xxx.Net", 												1), _
		Array("Length", 					"", 			 		 8,	Simple,		Rt,	"Format(xxx.Length, ""0.000"")", 						1), _
		Array("Width", 						"", 			 		 8,	Simple,		Rt,	"Format(xxx.Width, ""0.#####"")", 						1), _
		Array("Type", 						"", 			 		 4,	Simple,		Lf,	"IIf(xxx.SegmentType = ppcbSegmentLine, ""Line"", IIf(xxx.SegmentType = ppcbSegmentArc, ""Arc"", """"))",	1), _
		Array("Layer Name", 				"", 			 		30,	Simple,		Lf,	"ActiveDocument.LayerName(xxx.layer)", 					1), _
		Array("Layer Number", 				"", 			 		12,	Simple,		Rt,	"xxx.layer", 											1), _
		Array("Point1 X", 					"", 			 		 8,	Simple,		Rt,	"GetPoint(xxx, 1, 1)", 									1), _
		Array("Point1 Y", 					"", 			 		 8,	Simple,		Rt,	"GetPoint(xxx, 1, 2)", 									1), _
		Array("Point2 X", 					"", 			 		 8,	Simple,		Rt,	"GetPoint(xxx, 2, 1)", 									1), _
		Array("Point2 Y", 					"", 			 		 8,	Simple,		Rt,	"GetPoint(xxx, 2, 2)", 									1), _
		Array("Point3 X", 					"", 			 		 8,	Simple,		Rt,	"GetPoint(xxx, 3, 1)", 									1), _
		Array("Point3 Y", 					"", 			 		 8,	Simple,		Rt,	"GetPoint(xxx, 3, 2)",									1)  _
	), _
	Array( _
		Array("Type", 						"", 			 		15,	Simple,		Lf,	"xxx.type", 											1), _
		Array("Drill Size", 				"", 			 		10, Simple,		Rt,	"xxx.DrillSize", 										1), _
		Array("Start Layer Name", 			"", 			 		30, Simple,		Lf,	"ActiveDocument.LayerName(xxx.StartLayer)", 			1), _
		Array("Start Layer Number", 		"", 			 		18, Simple,		Rt,	"xxx.StartLayer", 										1), _
		Array("End Layer Name", 			"", 			 		30, Simple,		Lf,	"ActiveDocument.LayerName(xxx.EndLayer)", 				1), _
		Array("End Layer Number", 			"", 			 		16,	Simple,		Rt,	"xxx.EndLayer", 										1), _
		Array("Glued (Yes/No)", 			"Glued", 		 		 5,	Simple,		Cn,	"Format(xxx.Glued, ""Yes/No"")", 						1), _
		Array("Test Point (Side/No)",		"Test Point", 			10,	Simple,		Lf,	"IIf (xxx.TestPoint = ppcbTestPointBottomLayer, ""Bottom"", IIf (xxx.TestPoint = ppcbTestPointTopLayer, ""Top"", ""No""))",	1), _
		Array("Position X", 				"", 			 		10,	Simple,		Rt,	"Format(xxx.PositionX, ""0.000"")", 					1), _
		Array("Position Y", 				"", 			 		10,	Simple,		Rt,	"Format(xxx.PositionY, ""0.000"")",						1)  _
	), _
	Array( _
		Array("Name", 						"", 					 8, Simple,		Lf,	"xxx.Name",												1), _
		Array("Installed (Yes/No)", 		"Installed",	 		 9,	Simple,		Cn,	"Format(xxx.Installed, ""Yes/No"")",					0), _
		Array("Net",	 					"", 			 		15,	Simple,		Lf,	"xxx.Net", 												1), _
		Array("Length", 					"", 			 		 8,	Simple,		Rt,	"Format(xxx.Length, ""0.#####"")", 						1), _
		Array("Pin1 X", 					"", 			 		 8,	Simple,		Rt,	"Format(xxx.Points(1).PositionX, ""0.000"")", 			1), _
		Array("Pin1 Y", 					"", 			 		 8,	Simple,		Rt,	"Format(xxx.Points(1).PositionY, ""0.000"")", 			1), _
		Array("Pin2 X", 					"", 			 		 8,	Simple,		Rt,	"Format(xxx.Points(2).PositionX, ""0.000"")", 			1), _
		Array("Pin2 Y", 					"", 			 		 8,	Simple,		Rt,	"Format(xxx.Points(2).PositionY, ""0.000"")",			1)  _
	), _
	Array( _
		Array("Name", 						"", 					15, Simple,		Lf,	"xxx.Name",												1), _
		Array("Net",	 					"", 			 		15,	Simple,		Lf,	"xxx.Net", 												1), _
		Array("Length", 					"", 			 		 8,	Simple,		Rt,	"Format(xxx.Length(False), ""0.000;;0"")", 				1), _
		Array("Routed Length", 				"", 			 		13,	Simple,		Rt,	"Format(ConnRoutedLength(xxx), ""0.000;;0"")", 			1), _
		Array("Unrouted Length", 			"", 			 		15,	Simple,		Rt,	"Format(Abs(xxx.Length - ConnRoutedLength(xxx)), ""0.000"")",	1), _
		Array("Pin Count", 					"Pins", 		 		 5, Simple,		Rt,	"xxx.Pins.Count",										1), _
		Array("Via Count", 					"Vias", 			 	 5,	Simple,		Rt,	"xxx.Vias.Count",										1), _
		Array("Route Segment Count", 		"Route Segments", 		14,	Simple,		Rt,	"xxx.RouteSegments.Count",								1)  _
	), _
	Array( _
		Array("Name", 						"", 					15, Simple,		Lf,	"xxx.Name",												1), _
		Array("Logic Family", 				"Logic", 				 5, Simple,		Cn,	"xxx.Logic",											1), _
		Array("Value",						"", 		  			 8,	Attr,		Lf,	"AttrVal(xxx, ""Value"")",								1), _
		Array("Tolerance",					"", 			 		 9,	Attr,		Lf,	"AttrVal(xxx, ""Tolerance"")",							1), _
		Array("ECO Registered (Yes/No)", 	"ECO", 			 		 3,	Simple,		Cn,	"Format(xxx.ECORegistered, ""Yes/No"")", 				1), _
		Array("Part Count", 				"Part Count",	 		10, Simple,		Rt,	"PartCountOfType(cont, xxx)",							1)  _
	), _
	Array( _
		Array("Net", 						"", 			 		15,	Simple,		Lf,	"ObjName(xxx.Net)",										1), _
		Array("Pin/Via", 					"", 					11, Simple,		Lf,	"GetTPName(xxx)",										1), _
		Array("Type (Pin/Via)", 			"", 					 4, Simple,		Lf,	"IIf (xxx.ObjectType = ppcbObjectTypePin, ""Pin"", ""Via"")",		1), _
		Array("Position X", 				"", 			 		10,	Simple,		Rt,	"Format(xxx.PositionX, ""0.000"")", 					1), _
		Array("Position Y", 				"", 			 		10,	Simple,		Rt,	"Format(xxx.PositionY, ""0.000"")",						1), _
		Array("Side (Top/Bottom)",			"Side", 				 6,	Simple,		Lf,	"IIf (xxx.TestPoint = ppcbTestPointTopLayer, ""Top"", ""Bottom"")",	1) _
	)  _
)

'Array of object lists
'		List Name ID,						List Header,   		   NumPerLine, OutType,VarName,	Collection,						Sorted Collection,			Filtered By Subsection Collection,	Sorted Filtered By Subsection Collection, Out																										Help String
Const primObjLists = Array( _
	Array( _
		Array("Part Pins",					"Part Pins:", 			 		5,	Simple,	"aPin",	"xxx.Pins",						"GetSortedPins(xxx)",		"",									"",										"yyy.Name\Pin Type:PinTypeShort(yyy)\Pin Signal:ObjName(yyy.Net)",											"List of all part pins (with optional type and signal info)"), _
		Array("Connected Pins",				"Connected Part Pins:", 		5,	Simple,	"aPin",	"GetConPins(xxx)",				"GetConPins(xxx,True)",		"",									"",										"yyy.Name\Pin Type:PinTypeShort(yyy)\Pin Signal:yyy.Net",													"List of only pins connected to a net (with optional type and signal info)"), _
		Array("Unconnected Pins",			"Unconnected Part Pins:", 		5,	Simple,	"aPin",	"GetUnconPins(xxx)",			"GetUnconPins(xxx,True)",	"", 								"",										"yyy.Name\Pin Type:PinTypeShort(yyy)",																		"List of only pins not connected to a net (with optional type info)"), _
		Array("Attributes", 				"Part Attributes:", 			1,	Attr,	"attr",	"xxx.Attributes",				"xxx.Attributes",			"",									"",										"yyy.Name",																									"List of part attributes")  _
	), _
	Array( _
		Array("Net Pins",					"Net Pins:", 			 		5,	Simple,	"aPin",	"xxx.Pins",						"GetSortedPins(xxx)",		"",									"",										"yyy.Name\Pin Type:PinTypeShort(yyy)",																		"List of all pins connected to the net (with optional type info)"), _
		Array("Net Vias",					"Net Vias:", 			 		5,	Simple,	"aVia",	"xxx.Vias",						"xxx.Vias",					"",									"",										"yyy.Name\Via Type:yyy.type",																				"List of all vias connected to the net (with optional type info)"), _
		Array("Net Connections",			"Net Connections:", 			5,	Simple,	"con",	"xxx.Connections",				"xxx.Connections",			"",									"",										"yyy.Name",																									"List of all connections combine to the net"), _
		Array("Attributes", 				"Net Attributes:", 				1,	Attr,	"attr",	"xxx.Attributes",				"xxx.Attributes",			"",									"",										"yyy.Name",																									"List of net attributes")  _
	), _
	Array( _
		Array("Attributes", 				"Pin Attributes:", 				1,	Attr,	"attr",	"xxx.Attributes",				"xxx.Attributes",			"",									"",										"yyy.Name",																									"List of pin attributes")  _
	), _
	Array( _
		Array("Route Segment Connections",	"Route Segment Connections:", 	5,	Simple,	"con",	"GetSegmentCons(xxx)",			"GetSegmentCons(xxx)",		"",									"",										"yyy.Name\Con. Net:yyy.Net",																				"List of all connections hold the route segment (With Optional net info)") _
	), _
	Array( _
		Array("Attributes", 				"Via Attributes:", 				1,	Attr,	"attr",	"xxx.Attributes",				"xxx.Attributes",			"",									"",										"OutViaName(yyy)",																							"List of via attributes")  _
	), _
	Array( _
	), _
	Array( _
		Array("Connection Route Segments",	"Connection Route Segments:", 	1,	RteSeg,	"seg",	"xxx.RouteSegments",			"xxx.RouteSegments",		"",									"",										"RteSeg\R.Seg. Type:IIf(yyy.SegmentType = ppcbSegmentLine, ""Line"", IIf(yyy.SegmentType = ppcbSegmentArc, ""Arc"", """"))",	"List of all route segments combine to the connection (With Optional type info)") _
	), _
	Array( _
		Array("Parts", 						"Parts:", 						5,	Simple,	"part",	"xxx.Components",				"",							"GetPartsByPartType(cont, xxx.PartType)",	"",								"yyy.Name\Value:ValPartAttr(yyy, ""Value"")\Tolerance:ValPartAttr(yyy, ""Tolerance"")",						"List of parts of this part type"), _
		Array("Attributes", 				"Part Type Attributes:", 		1,	Attr,	"attr",	"xxx.Attributes",				"xxx.Attributes",			"",									"",										"yyy.Name",																									"List of part type attributes")  _
	), _
	Array( _
	) _
)

'Array of primary ojects
'	List Name ID,			Hdr,NPL,OutType,VarName,Collection,				Sorted Collection,			Rsv	Out																																Help String
Const primList = Array( _
	Array( _
		Array("Parts",		"", 1,	Simple,	"aPin",	"cont.Components",		"",							,,	"xxx.Name\Part Type:xxx.PartType\Decal:xxx.Decal\Value:AttrVal(xxx, ""Value"")\Tolerance:AttrVal(xxx, ""Tolerance"")",			"") _
	), _
	Array( _
		Array("Nets",		"", 1,	Simple,	"aNet",	"cont.Nets",			"",							,,	"xxx.Name\Net Length:Format(xxx.Length, ""0.#####"")",																			"") _
	), _
	Array( _
		Array("Pins", 		"", 1,	Simple,	"aPin",	"GetSortedPins(cont)",	"",							,,	"xxx.Name\Pin Type:PinTypeShort(xxx)\Pin Signal:ObjName(xxx.Net)",																"") _
	), _
	Array( _
		Array("RSegments", 	"", 1,	RteSeg,	"seg",	"cont.RouteSegments",	"",							,,	"RteSeg\R.Seg. Signal:xxx.Net\R.Seg. Length:Format(xxx.Length, ""0.00000"")\R.Seg. Width:xxx.Width\R.Seg. Type:IIf(xxx.SegmentType = ppcbSegmentLine, ""Line"", IIf(xxx.SegmentType = ppcbSegmentArc, ""Arc"", """"))",	"") _
	), _
	Array( _
		Array("Vias", 		"", 1,	Simple,	"aVia",	"cont.Vias",			"",							,,	"xxx.Type\Via Coordinates:OutViaName(xxx)\Via Signal:ObjName(xxx.Net)",															"") _
	), _
	Array( _
		Array("Jumpers",	"", 1,	Simple,	"jmp",	"cont.Jumpers",			"",							,,	"xxx.Name\Jmp. Signal:ObjName(xxx.Net)\Jmp. Length:Format(xxx.Length, ""0.00000"")",											"") _
	), _
	Array( _
		Array("Connections", "", 1,	Simple,	"con",	"cont.Connections",		"",							,,	"xxx.Name\Con. Signal:ObjName(xxx.Net)\Con. Length:Format(xxx.Length, ""0.00000"")",											"") _
	), _
	Array( _
		Array("Part Types", "", 1,	Simple,	"pkg",	"cont.PartTypes",		"",							,,	"xxx.Name\Value:AttrVal(xxx, ""Value"")\Tolerance:AttrVal(xxx, ""Tolerance"")",																													"") _
	), _
	Array( _
		Array("Test Points", "", 1,	Simple,	"tPnt",	"GetTestPoints(cont)",	"GetTestPoints(cont,True)",	,,	"GetTPName(xxx)\TPnt. Signal:ObjName(xxx.Net)\Pos. X:Format(xxx.PositionX, ""0.000"")\Pos. Y:Format(xxx.PositionY, ""0.000"")\Side:IIf (xxx.TestPoint = ppcbTestPointTopLayer, ""Top"", ""Bottom"")",	"") _
	) _
)

'	condition text,					template,						nOperands,	Flags (1-Text, 2-Number)
Const CondList = Array ( _
	"is between", 					"_zzz_ >= _xxx_ And _zzz_ <= _yyy_",	2,  3, _
	"is not between", 				"(_zzz_ < _xxx_ Or _zzz_ > _yyy_)",		2,	3, _
	"contains", 					"InStr(_zzz_, _xxx_)",					1,	1, _
	"begins with", 					"Left(_zzz_, _lll_) = _xxx_",			1,	1, _
	"ends with", 					"Right(_zzz_, _lll_) = _xxx_",			1,	1, _
	"is equal to", 					"_zzz_ = _xxx_",						1,	3, _
	"is not equal to", 				"_zzz_ <> _xxx_",						1,	3, _
	"is greater than", 				"_zzz_ > _xxx_",						1,	3, _
	"is less than", 				"_zzz_ < _xxx_",						1,	3, _
	"is greater than or equal to", 	"_zzz_ >= _xxx_",						1,	3, _
	"is less than or equal to",		"_zzz_ <= _xxx_",						1,	3  _
)

Const objProps = Array("Attributes, Part Type, PCB Decal, Logic Family, Pin Count.", _
						"Pin Count, Length.", _
						"Number, ElectricalType, Net, etc.", _
						"Net, Length, Layer, etc.", _
						"Start Layer, End Layer, etc.", _
						"Net, Length, Point 1, Point 2", _
						"Net, Length, Pin Count, etc.", _
						"Logic Family, Pin Count, etc.", _
						"Signal, Name, Pos. X, Pos. Y, Side")
						
Const objLists = Array("For example: All Pins, Connected Pins, Unconnected Pins, Attributes", _
						"For example: Net Pins, Net Vias, Net Connections, Attributes.", _
						"For example: Attributes", _
						"For example: Connections", _
						"For example: Attributes", _
						"", _
						"For example: Route Segments", _
						"For example: List of parts of this type", _
						"")

Const attrW = 32
Const contHeader = "Assembly Option"
Const ACTIVE_SHEET = "Always for Current Sheet"
Const DefNumPerLine = 5
Const dblQt = Chr(34) & Chr(34)

Const targetApp = Array("Notepad", "Write")

'Global Declarations
Dim scriptDir, scriptName, scriptPath, header, repname, ext, txtFileExt, curDlg, dlgHandle, colorDlgHnd, preview, curStep, nSteps, curProperty
Dim Table, genHdr, incJob, incTime, primTotal, primContTotal, contContents, filterByContainer, shtFilter, filterByCount, enableAlign, bInstalledOnly
Dim	secTotal, sortObj, showProgress, showStatus, RTF, HTML, TXT, XLS, sortedPinsDone, colorTable, bkgDef, fontDef, tblHdrOut, drawVBorder, drawHBorder
Dim secObj, target, outputFile, tableSec, numPerLine, nSecObjs, nTabs, curAttribute, curAttrName
Dim primObj, primVarName, primName, primColl, primSortColl, primProps, primNameProp, contObj, givenCont, allCont
Dim Spc, NewRow, NewLine, ListIndent, OutLine, OutBoldLine, OutBoldUnLnLine, DLMT, AttrDLMT, RoutDLMT, colorColumns
Dim headerRule, repnameRule, scriptNameRule, oldAutoName
Dim bc0, fc0, bc1, fc1, bc2, fc2, lastScheme
Dim code As String
Dim	repByContainer	' sheet by sheet report (for PADS Logic); assembly options report (for PADS Layout)
Dim prefix

Dim ReportNameList() As String
Dim ScriptNameList() As String
Dim AttrArray() As String
Dim AvailPropArray() As String
Dim PropArray() As String
Dim ColConfigs() As String
Dim CondArray() As String
Dim ContArray() As String

'Main Entry
Sub Main
	RestoreWizardSettings
	oldAutoName = ReportAutoName
	curStep = 1
	nSteps = IIf (tableSec = 2, 7, 8)
	Page1
End Sub


'Wizard dialogs
Sub Page1
	Begin Dialog UserDialog 560,280,"Introduction",.dlgfunc ' %GRID:10,7,1,1
		Text 160,7,220,14,"Welcome to the VB Script Wizard.",.Text2
		Text 10,28,540,42,"This program is completely written in PADS Layout's VB script language. It will generate a skeleton VB script for extracting database information in different application formats including RTF, Excel, and HTML.",.Text6
		Text 10,154,540,28,"Now you can easily put your colored HTML reports on the company's Intranet or Internet sites to share them with your colleagues or customers!",.Text8
		Text 10,77,540,42,"The Wizard consists of a maximum of 8 steps. All settings you have made will be saved in the Windows Registry, so you can always rerun the Wizard and then tune settings using a previous report as a preview.",.Text1
		Text 10,119,540,28,"After completing this Wizard, you can extend the resulting VB script by manually writing additional code.",.Text7
		GroupBox 10,189,540,42,"VB Script Destination Directory",.Group1
		Text 20,210,430,14,"Text1",.vbDir
		PushButton 330,252,100,21,"&Next >",.NextBtn
		PushButton 450,203,90,21,"B&rowse...",.Browse
		CancelButton 450,252,100,21
		Text 10,231,540,14,"___________________________________________________________________",.Line
	End Dialog
	Dim n
	Dim dlg As UserDialog
	While True
		curDlg = "1"
		n = Dialog(dlg, 1)
		If n = 0 Then Abort
		curStep = curStep + 1
		Page2
		curStep = curStep - 1
	Wend
End Sub

Sub Page2
	Begin Dialog UserDialog 560,280,"Format",.dlgfunc ' %GRID:10,7,1,1
		Text 40,14,390,28,"Choose the target application format for the new VB script. Resulting VB code will vary depending on your choice.",.Text1
		OptionGroup .Target
			OptionButton 40,63,160,14,"&Notepad",.Notepad
			OptionButton 40,105,160,14,"&WordPad / MS Word",.WordPad
			OptionButton 40,147,160,14,"Microsoft &Excel",.Excel
			OptionButton 40,189,160,14,"&Internet Browser",.ie
		Text 220,63,320,14,"Plain text format. Fixed width font recommended.",.Text3
		Text 220,105,310,14,"Rich text format. ",.Text4
		Text 220,147,310,14,"Microsoft Excel Spreadsheet.",.Text5
		Text 220,189,160,14,"HTML format.",.Text6
		Text 220,77,310,14,"(Minimum formatting)",.Text7
		Text 220,119,310,14,"(Better formatting)",.Text8
		Text 220,161,310,14,"(Better formatting with tables)",.Text9
		Text 220,203,320,14,"(Best output quality, with colored tables)",.Text10
		PushButton 330,252,100,21,"&Next >",.NextBtn
		PushButton 230,252,100,21,"< &Back",.BackBtn
		CancelButton 450,252,100,21
		Text 10,231,540,14,"___________________________________________________________________",.Line
		PushButton 430,14,100,21,"More Info...",.HelpBtn
	End Dialog
	Dim n
	Dim dlg As UserDialog
	While True
		dlg.Target = target
		curDlg = "2"
		n =  Dialog(dlg, 1)
		If n = 0 Then Abort
		If target <> dlg.Target Then
			target = dlg.Target
			If target = Excel Then
				outputFile = 2
			ElseIf target = WordPad Or target = Browser Then
				If outputFile = 2 Then outputFile = 0
			End If
		End If
		If n = 2 Then Exit Sub
		curStep = curStep + 1
		Page3
		curStep = curStep - 1
	Wend
End Sub

Sub Page3
	Begin Dialog UserDialog 560,280,"Report Type",.dlgfunc ' %GRID:10,7,1,1
		Text 30,14,520,14,"Please choose which kind of report you want the new VB script to generate:",.Text1
		OptionGroup .Container
			OptionButton 30,49,210,14,"&PCB-Based Reports",.PCBBased
			OptionButton 30,119,200,14,"&Assembly Option Reports",.AssemblyBased
			OptionButton 30,189,270,14,"Reports for &Given Assembly Option",.Given
		DropListBox 60,210,490,70,ContArray(),.ContList,1
		Text 60,70,480,28,"Reports will contain database items collected from the entire design. Assembly Option will not be taken into account.",.Text2
		Text 60,140,480,28,"Part and jumper information will be extracted for each assembly option indivudually.",.Text5
		PushButton 330,252,100,21,"&Next >",.NextBtn
		PushButton 230,252,100,21,"< &Back",.BackBtn
		CancelButton 450,252,100,21
		Text 10,231,540,14,"___________________________________________________________________",.Line
	End Dialog
	Dim n, i, s
	Dim dlg As UserDialog
	While True
		If givenCont = "" Then
			dlg.Container = IIf (repByContainer, 1, 0)
		Else
			dlg.Container = 2
		End If
		curDlg = "3"
		n = ActiveDocument.AssemblyOptions.Count
		ReDim ContArray(n)
		For i = 1 To n
			s = ActiveDocument.AssemblyOptions(i)
			ContArray(i) = Left(s, Len(s) - Len(ActiveDocument.Name) - 1)
		Next
		dlg.ContList = ContArray(1)
		n =  Dialog(dlg, 1)
		If n = 0 Then Abort
		repByContainer = dlg.Container <> 0
		givenCont = IIf (dlg.Container = 2, dlg.ContList, "")
		allCont = dlg.Container = 1
		If givenCont = "" Then
			contObj = contObjs(IIf(repByContainer, 1, 0))
		Else 
			contObj = "ActiveDocument.AssemblyOptions(" & Chr(34) & givenCont & Chr(34) & ")"
		End If
		If repByContainer And Not primData(primObj)(6) Then primObj = 0	'is object allowed for sheet/ass.opt. report
		If repByContainer Then secObj = 0
		
		If n = 2 Then Exit Sub
		curStep = curStep + 1
		Page4
		curStep = curStep - 1
	Wend
End Sub

Sub Page4
	Begin Dialog UserDialog 560,280,"Database Object",.dlgfunc ' %GRID:10,7,1,1
		Text 40,21,510,14,"Please choose the type of database object that will be primary for the report.",.Text1
		OptionGroup .primObj
			OptionButton 60,49,60,14,"&Parts",.itPart
			OptionButton 220,49,60,14,"&Nets",.itNet
			OptionButton 220,77,60,14,"P&ins",.itPin
			OptionButton 350,77,140,14,"&Route Segments",.itRouteSegment
			OptionButton 220,105,60,14,"&Vias",.itVia
			OptionButton 60,77,80,14,"&Jumpers",.itJumper
			OptionButton 350,49,110,14,"&Connections",.itConnection
			OptionButton 60,105,130,14,"Part &Types",.itPartType
			OptionButton 350,105,140,14,"Te&st Points",.itTestPoint
		CheckBox 50,133,420,14,"Calculate the total number of primary objects in the design",.primTotal
		CheckBox 50,168,490,14,"Output installed objects only",.installedOnly
		CheckBox 50,203,490,14,"Calculate the number of objects for each assembly option individually",.primContTotal
		PushButton 330,252,100,21,"&Next >",.NextBtn
		PushButton 230,252,100,21,"< &Back",.BackBtn
		CancelButton 450,252,100,21
		Text 10,231,540,14,"___________________________________________________________________",.Line
	End Dialog
	Dim n
	Dim dlg As UserDialog
	While True
		dlg.primObj = primObj
		dlg.primTotal = primTotal
		dlg.installedOnly = bInstalledOnly
		dlg.primContTotal = primContTotal
		curDlg = "4"
		n =  Dialog(dlg, 1)
		If n = 0 Then Abort
		If primObj <> dlg.primObj Then
			SavePropArray	' save previous array values 
			ReDim AttrArray(0)
			primObj = dlg.primObj
			secObj = 0
			curAttribute = 0
			LoadPropArray
			If tableSec = 1 And Not primData(primObj)(8) Then tableSec = 2 'no related objects
			If Not primData(primObj)(7) Then contContents = False 'no index allowed for this object
		End If
		primName = primData(primObj)(0)
		primVarName = primData(primObj)(1)
		primNameProp = primData(primObj)(9)
		FormatTemplate primNameProp, "xxx", primVarName
		primColl = primData(primObj)(2)
		primSortColl = primData(primObj)(3)
		primTotal = dlg.primTotal = 1
		bInstalledOnly = dlg.installedOnly = 1 And primData(primObj)(6) 'object allowed for per container report
		primContTotal = dlg.primContTotal = 1
		If n = 2 Then Exit Sub
		curStep = curStep + 1
		Page5
		curStep = curStep - 1
	Wend
End Sub

Sub Page5
	Begin Dialog UserDialog 560,280,"Data Type",.dlgfunc ' %GRID:10,7,1,1
		Text 20,14,530,35,"You have chosen XXX as the primary type of database object for the VB report. Now you can choose what data the new VB script will extract from a XXX.",.Text1
		OptionGroup .TableSec
			OptionButton 50,56,490,14,"&General XXX properties in table format",.Opt1
			OptionButton 50,140,490,14,"&List(s) of secondary objects related to each XXX",.Opt2
			OptionButton 50,196,240,14,"List of XXX objects &only",.Opt3
		Text 90,105,450,14,"",.ObjProps
		Text 90,161,450,14,"",.ObjLists
		Text 90,77,450,28,"Each row of the table represents an individual XXX, and each column represents a XXX property, for example:",.Text5
		Text 10,231,540,14,"___________________________________________________________________",.Line
		PushButton 90,217,140,21,"&Edit List...",.EditList
		PushButton 330,252,100,21,"&Next >",.NextBtn
		PushButton 230,252,100,21,"< &Back",.BackBtn
		CancelButton 450,252,100,21
	End Dialog
	Dim n
	Dim dlg As UserDialog
	While True
		dlg.TableSec = tableSec
		curDlg = "5"
		n =  Dialog(dlg, 2)
		If n = 0 Then Abort
		SavePropArray
		tableSec = dlg.TableSec
		Table = tableSec = 0
		LoadPropArray
		If n = 3 Then Exit Sub
		curStep = curStep + 1
		If Table Then
			secObj = 0
			primProps = primObjProps(primObj)
			If primData(primObj)(5) Then 'has attribute
				Page6
			Else
				Page6A
			End If
		ElseIf tableSec = 2 Then
			secObj = 0
			primProps = primList(primObj)
			Page7
		Else
			primProps = primObjLists(primObj)
			Page6B
		End If
		curStep = curStep - 1
	Wend
End Sub

Sub Page6
	Begin Dialog UserDialog 560,280,"Object Properties",.dlgfunc ' %GRID:10,7,1,1
		Text 10,231,540,14,"___________________________________________________________________",.Line
		Text 10,7,540,14,"Add XXX properties or attributes to the list of table columns.",.Text1
		Text 350,28,190,14,"Table &Columns:",.Text2
		Text 10,28,190,14,"XXX Properties:",.Text3
		ListBox 10,42,200,154,AvailPropArray(),.AvailPropList
		ListBox 350,42,200,189,PropArray(),.PropList
		PushButton 220,42,120,21,"A&dd >>",.AddBtn
		PushButton 220,63,120,21,"<< &Remove",.RemoveBtn
		PushButton 280,119,60,21,"&Up",.UpBtn
		PushButton 280,140,60,21,"&Down",.DownBtn
		Text 10,196,330,14,"&Attribute (Select existing or type any valid name):",.Text4
		DropListBox 10,210,330,70,AttrArray(),.AttrName,1
		PushButton 280,168,60,21,"&Edit...",.EditBtn
		PushButton 220,91,120,21,"<< Remove A&ll",.ResetBtn
		PushButton 330,252,100,21,"&Next >",.NextBtn
		PushButton 230,252,100,21,"< &Back",.BackBtn
		CancelButton 450,252,100,21
	End Dialog
	Dim n
	Dim dlg As UserDialog
	While True
		curDlg = "6"
		n =  Dialog(dlg, 7)
		If n = 0 Then Abort
		If n = 8 Then Exit Sub
		curStep = curStep + 1
		Page7
		curStep = curStep - 1
	Wend
End Sub

Sub Page6A
	Begin Dialog UserDialog 560,280,"Object Properties",.dlgfunc ' %GRID:10,7,1,1
		Text 10,231,540,14,"___________________________________________________________________",.Line
		Text 10,7,540,14,"Add XXX properties to the list of table columns.",.Text1
		Text 10,28,200,14,"XXX Properties:",.Text3
		ListBox 10,42,200,189,AvailPropArray(),.AvailPropList
		Text 350,28,200,14,"Table &Columns:",.Text2
		ListBox 350,42,200,189,PropArray(),.PropList
		PushButton 220,42,120,21,"A&dd >>",.AddBtn
		PushButton 220,63,120,21,"<< &Remove",.RemoveBtn
		PushButton 280,147,60,21,"&Up",.UpBtn
		PushButton 280,168,60,21,"&Down",.DownBtn
		PushButton 280,203,60,21,"&Edit...",.EditBtn
		PushButton 220,105,120,21,"<< Remove A&ll",.ResetBtn
		PushButton 330,252,100,21,"&Next >",.NextBtn
		PushButton 230,252,100,21,"< &Back",.BackBtn
		CancelButton 450,252,100,21
	End Dialog
	Dim n
	Dim dlg As UserDialog
	While True
		curDlg = "6A"
		n =  Dialog(dlg, 7)
		If n = 0 Then Abort
		If n = 8 Then Exit Sub
		curStep = curStep + 1
		Page7
		curStep = curStep - 1
	Wend
End Sub

Sub Page6B
	Begin Dialog UserDialog 560,280,"Associated Objects",.dlgfunc ' %GRID:10,7,1,1
		Text 10,231,540,14,"___________________________________________________________________",.Line
		Text 10,7,540,14,"Please choose the list(s) of database objects related to a XXX.",.Text1
		Text 10,21,540,14,"VB Script will generate selected list(s) for each XXX entry in the report.",.Text12
		Text 10,42,210,14,"&Lists of XXX Related Objects:",.Text3
		ListBox 10,56,200,126,AvailPropArray(),.AvailPropList
		Text 350,42,200,14,"&Report:",.Text2
		ListBox 350,56,200,126,PropArray(),.PropList
		PushButton 220,49,120,21,"A&dd >>",.AddBtn
		PushButton 220,70,120,21,"<< &Remove",.RemoveBtn
		PushButton 280,119,60,21,"&Up",.UpBtn
		PushButton 280,140,60,21,"&Down",.DownBtn
		PushButton 280,161,60,21,"&Edit...",.EditBtn
		PushButton 220,91,120,21,"<< Remove A&ll",.ResetBtn
		Text 10,182,30,14,"Info:",.Text4
		Text 40,182,520,14,"",.Info
		CheckBox 10,203,190,14,"&Count objects in each list",.SecTotal
		CheckBox 240,203,280,14,"Objects always appear &Sorted in Lists",.SortObj
		CheckBox 10,224,540,14,"&Filter by sheet  ( Include only objects located on currently considered sheet )",.filterByContainer
		PushButton 330,252,100,21,"&Next >",.NextBtn
		PushButton 230,252,100,21,"< &Back",.BackBtn
		CancelButton 450,252,100,21
	End Dialog
	Dim n
	Dim dlg As UserDialog
	While True
		curDlg = "6B"
		dlg.secTotal = secTotal
		dlg.sortObj = sortObj
		dlg.filterByContainer = filterByContainer
		n =  Dialog(dlg, 7)
		If n = 0 Then Abort
		secObj = UBound(PropArray) > 0
		nSecObjs = UBound(PropArray)
		filterByContainer = dlg.filterByContainer = 1
		shtFilter = repByContainer And filterByContainer
		secTotal = dlg.secTotal = 1
		sortObj = dlg.sortObj =  1
		If n = 8 Then Exit Sub
		curStep = curStep + 1
		Page7
		curStep = curStep - 1
	Wend
End Sub

Sub Page7
	Begin Dialog UserDialog 560,280,"Report Options",.dlgfunc ' %GRID:10,7,1,1
		Text 10,7,540,14,"Select additional report options for the VB script.",.Text1
		CheckBox 20,28,200,14,"Output Report &Header",.genHdr
		DropListBox 50,42,500,56,ReportNameList(),.Header,1
		Text 50,70,120,14,"Header Includes:",.Text3
		CheckBox 200,70,100,14,"&Job Name",.IncJob
		CheckBox 320,70,180,14,"&Date and Time Stamp",.IncTime
		Text 20,105,240,14,"Show report generation progress in:",.Text2
		CheckBox 270,105,100,14,"&Status Bar",.ShowStatus
		CheckBox 380,105,120,14,"&Progress Bar",.ShowProgress
		CheckBox 20,140,530,14,"Enable Text &Alignment for Table Columns",.Align
		CheckBox 20,210,250,14,"&Use HTML tables color schemes",.colTbl
		PushButton 280,203,180,21,"Setup &Color Scheme...",.SetupColors
		CheckBox 20,175,530,14,"Output Index / &Table of Contents (for HTML reports only)",.Contents
		Text 20,210,230,14,"Draw Table Border ",.DrawBorderLab
		CheckBox 160,210,100,14,"Hori&zontal",.DrawHBorder
		CheckBox 270,210,110,14,"&Vertical",.DrawVBorder
		Text 10,231,540,14,"___________________________________________________________________",.Line
		PushButton 330,252,100,21,"&Next >",.NextBtn
		PushButton 230,252,100,21,"< &Back",.BackBtn
		CancelButton 450,252,100,21
		CheckBox 20,210,540,14,"Eliminate XXX entry if requested lists are empty (slower due to precalculations)",.FilterByCount
	End Dialog
	Dim n
	Dim dlg As UserDialog
	While True
		ReDim ReportNameList(2)
		ReportNameList(0) = ReportAutoName
		If ReportNameList(0) <> header And headerRule = 2 Then
			ReportNameList(1) = header
		End If
		If oldAutoName <> ReportAutoName Then
			headerRule = 1
			oldAutoName = ReportAutoName
		End If
		dlg.Header = Choose(headerRule, ReportAutoName, header)
		dlg.genHdr = genHdr
		dlg.incJob = incJob
		dlg.incTime = incTime
		dlg.ShowStatus = ShowStatus
		dlg.ShowProgress = ShowProgress
		dlg.colTbl = colorTable
		dlg.Contents = contContents
		dlg.DrawHBorder = drawHBorder
		dlg.DrawVBorder = drawVBorder
		dlg.Align = enableAlign
		dlg.FilterByCount = filterByCount
		curDlg = "7"
		n =  Dialog(dlg, 2)
		If n = 0 Then Abort
		genHdr = dlg.genHdr = 1
		incJob = dlg.incJob = 1
		incTime = dlg.incTime = 1
		headerRule = IIf (dlg.Header = ReportAutoName, 1, 2)
		header = dlg.Header
		ShowProgress = dlg.ShowProgress = 1
		ShowStatus = dlg.ShowStatus = 1
		colorTable = dlg.colTbl = 1
		contContents = dlg.Contents = 1
		drawHBorder = dlg.DrawHBorder = 1
		drawVBorder = dlg.DrawVBorder = 1
		enableAlign = dlg.Align = 1
		filterByCount = dlg.FilterByCount = 1
		If n = 3 Then Exit Sub
		curStep = curStep + 1
		Page8
		curStep = curStep - 1
	Wend
End Sub

Sub Page8
	Begin Dialog UserDialog 560,280,"Output Files",.dlgfunc ' %GRID:10,7,1,1
		Text 10,7,550,14,"Select the name of the report file to which the new VB Script will extract information.",.Text1
		OptionGroup .OutputFile
			OptionButton 20,28,480,14,"Always &based on job name (name of the schematic file will be used)",.jobBased
			OptionButton 20,63,420,14,"&Custom report file name (type without file extension):",.custom
			OptionButton 20,119,490,14,"&Create new untitled document and pass data via Clipboard",.Clipboard
		DropListBox 60,77,390,63,ReportNameList(),.CustomName,1
		Text 460,63,100,14,"File &Extension:",.ExtLabel
		TextBox 460,77,50,21,.Ext
		Text 60,133,480,14,"This option grays if you chose WordPad or Internet Browser in step 2.",.Text2
		Text 60,147,480,14,"(Recommended for Excel to avoid file share violations.)",.Text3
		Text 10,231,540,14,"___________________________________________________________________",.Line
		Text 10,182,460,14,"&Enter the name of the resulting VB script (without BAS extension):",.Text4
		DropListBox 10,196,540,63,ScriptNameList(),.scriptName,1
		PushButton 110,252,100,21,"&Finish",.FinishBtn
		PushButton 230,252,200,21,"Finish && &Run Report Now",.RunBtn
		PushButton 10,252,100,21,"< &Back",.BackBtn
		CancelButton 450,252,100,21
	End Dialog
	Dim n
	Dim dlg As UserDialog
	While True
		If target = Notepad Then
			dlg.Ext = txtFileExt
		Else
			dlg.Ext = ExtMap(target)
		End If
		'report name list		
		ReDim ReportNameList(3)
		ReportNameList(0) = header
		If ReportNameList(0) <> ReportAutoName Then
			ReportNameList(1) = ReportAutoName
		End If
		If ReportNameList(0) <> repname And ReportNameList(1) <> repname And repnameRule = 3 Then
			ReportNameList(2) = repname
		End If
		If oldAutoName <> ReportAutoName Then
			repnameRule = 1
		End If
		dlg.CustomName = Choose(repnameRule, header, ReportAutoName, repname)
		'script name list		
		ReDim ScriptNameList(3)
		ScriptNameList(0) = header
		If ScriptNameList(0) <> ReportAutoName Then
			ScriptNameList(1) = ReportAutoName
		End If
		If ScriptNameList(0) <> scriptName And ScriptNameList(1) <> scriptName And scriptNameRule = 3 Then
			ScriptNameList(2) = scriptName
		End If
		If oldAutoName <> ReportAutoName Then
			scriptNameRule = 1
		End If
		dlg.ScriptName = Choose(scriptNameRule, header, ReportAutoName, scriptName)
		
		dlg.OutputFile = outputFile
		curDlg = "8"
		n =  Dialog(dlg, 2)
		If n = 0 Then Abort
		If target = Notepad Then
			txtFileExt = dlg.Ext
		End If
		outputFile = dlg.OutputFile
		repnameRule = IIf (dlg.CustomName = header, 1, IIf (dlg.CustomName = ReportAutoName, 2, 3))
		repname = dlg.CustomName
		scriptNameRule = IIf (dlg.ScriptName = header, 1, IIf (dlg.ScriptName = ReportAutoName, 2, 3))
		scriptName = dlg.ScriptName
		If n = 3 Then Exit Sub
		GenerateCode
		SaveWizardSettings

		'Reload script by this internal command and param
		ProcessCommand 20197
		ProcessParameter 9108, scriptPath
		If n = 2 Then
			'Run script
			On Error Resume Next
			MacroRun scriptPath
		Else
			MsgBox "VB Script generated successfully in the location:" & vbCr & scriptPath, vbInformation, AppName
		End If
		
		Abort
	Wend
End Sub

Function CheckFileName(fname As String) As Boolean
	Dim i, BadChars
	BadChars = "\/:<>*?""|"
	CheckFileName = False
	For i = 1 To Len(BadChars)
		If InStr(fname, Mid(BadChars, i, 1)) > 0 Then Exit Function
	Next
	CheckFileName = True
End Function

Sub UpdateTitle
	DlgText -1, WizardTitle & " - Step " & Trim(Str(CurStep)) & " of " & Trim(Str(nSteps))
End Sub

Sub UpdateButtons
	Dim id
	id = DlgValue("PropList")
	If Left(curDlg, 1) = "6" Then
		DlgEnable "UpBtn", id > 0
		DlgEnable "DownBtn", id >= 0 And id < UBound(PropArray) - 1
		DlgEnable "RemoveBtn", id >= 0
		DlgEnable "EditBtn", id >= 0
		If curDlg <> "6" Then
			id = DlgValue("AvailPropList")
			DlgEnable "AddBtn", id >= 0
		End If
		ShowInfo
	End If
End Sub

Sub UpdateObjectName
	Dim i, pos, s
	For i = 0 To DlgCount - 1
		If DlgType(i) = "OptionButton" Or DlgType(i) = "Text" Or DlgType(i) = "CheckBox" Then 
			pos = InStr(DlgText(i), "XXX") 
			While pos > 0
				s = Left(DlgText(i), pos - 1) & PrimName & Mid(DlgText(i), pos + 3)
				DlgText i, s
				pos = InStr(DlgText(i), "XXX") 'find again
			Wend
		End If
	Next
End Sub

Sub ShowInfo
	Dim id
	If curDlg = "6B" Then
		id = DlgValue("AvailPropList")
		If id >= 0 Then
			id = FindProp(DlgText("AvailPropList"))
			DlgText "Info", primProps(id)(10)
		Else
			DlgText "Info", ""
		End If
	End If
End Sub

Sub UpdateAvailProps(Optional curProp As String = "")
	Dim count, i, j, n, cur, found
	EnableWindow dlgHandle, 0
	count = UBound(primProps)
	n = 0
	cur = IIf(curProp <> "", -1, 0)
	ReDim AvailPropArray(count)
	For i = 0 To count
		found = False
		If curDlg = "6B" Then
			If repByContainer And DlgValue("filterByContainer") = 1 And primProps(i)(7) = "" Then
				found = True
			End If
		End If
		If Not found Then
			For j = 0 To UBound(propArray)-1
				If UCase(propArray(j+1)) = UCase(primProps(i)(0)) Then 
					found = True
					Exit For
				End If
			Next
		End If
		If Not found And curDlg = "6" Then
			If Not repByContainer And primProps(i)(6) = 0 Then 'this prop only for cont reports
				found = True
			End If
		End If
		If Not found Then
			AvailPropArray(n) = primProps(i)(0)
			n = n + 1
			If curProp <> "" And UCase(curProp) = UCase(primProps(i)(0)) Then
				cur = n - 1
			End If
		End If
	Next
	ReDim Preserve AvailPropArray(n)
	DlgListBoxArray "AvailPropList", AvailPropArray()
	DlgValue "AvailPropList", cur
	UpdateButtons
	ShowInfo
	EnableWindow dlgHandle, 1
End Sub

Sub UpdateAvailAttrs(Optional curAttr As String = "")
	Dim i, s
	EnableWindow dlgHandle, 0
	If curAttr <> "" Then
		s = curAttr
	Else
		s = DlgText("AttrName")
	End If
	GetAttributeList
	DlgListBoxArray "AttrName", AttrArray()					
	DlgText "AttrName", s
	If s = "" Then
		DlgValue "AttrName", curAttribute
	Else
		DlgText "AttrName", s
		For i = 1 To UBound(AttrArray)
			If UCase(AttrArray(i)) = UCase(s) Then 
				DlgValue "AttrName", i - 1
				Exit For
			End If
		Next
	End If
	curAttrName = s
	curAttribute = DlgValue("AttrName")
	EnableWindow dlgHandle, 1
End Sub

Sub UpdateSelectedProps
	Dim cur, i, j, n, s
	EnableWindow dlgHandle, 0
	n = UBound(PropArray)
	For i = n To 1 Step -1
		If curDlg = "6B" Then
			j = FindProp(PropArray(i))
			If j < 0 Then 
				RemoveProp(i)
			ElseIf repByContainer And DlgValue("filterByContainer") = 1 And primProps(j)(7) = "" Then
				RemoveProp(i)				
			End If
		ElseIf Not repByContainer And curDlg = "6" Then
			j = FindProp(PropArray(i))
			If j >= 0 Then
				If primProps(j)(6) = 0 Then RemoveProp(i)	'this prop only for cont reports
			End If
		End If		
	Next
	n = UBound(PropArray)
	Dim DisplayPropArray() As String
	ReDim DisplayPropArray(n)
	For i = 1 To n
		s = PropArray(i)
		DisplayPropArray(i) = s & IIf(PropCondFormat(s), " (C)", "")
	Next
	cur = DlgValue("PropList")
	DlgListBoxArray "PropList", DisplayPropArray()
	DlgValue "PropList", cur
	EnableWindow dlgHandle, 1
End Sub

Sub RemoveProp(id)
	Dim cur, count, i
	EnableWindow dlgHandle, 0
	cur = DlgValue("PropList")
	count = UBound(PropArray) - 1
	For i = id To count 
		PropArray(i) = PropArray(i + 1)
	Next i
	ReDim Preserve PropArray(count)
	DlgValue "PropList", IIf (cur >= count, cur - 1, cur)
	UpdateButtons
	EnableWindow dlgHandle, 1
End Sub

Function FindProp(id As String) As Integer
	Dim j
	FindProp = -1
	For j = 0 To UBound(primProps)
		If id = primProps(j)(0) Then 
			FindProp = j
			Exit For
		End If
	Next
End Function

Function SetGetProp(section, entry, Optional newVal, Optional defVal)
	SetGetProp = GetSetting(WizName, section,  entry, defVal)
	If IsNull(newVal) Then
		On Error Resume Next
		DeleteSetting WizName, section,  entry
		On Error GoTo 0
	ElseIf Not IsMissing(newVal) Then
		SaveSetting WizName, section,  entry, newVal
	End If
End Function

Function PropHeader(id As String, Optional newVal)
	Dim j, defVal
	j = FindProp(id)
	If j >= 0 Then
		defVal = IIf (primProps(j)(1) <> "", primProps(j)(1), primProps(j)(0))
	Else 'attribute
		defVal = id
	End If
	PropHeader = SetGetProp(PropSection, primName & " Header " & id, newVal, defVal)
End Function

Function PropWidth(id As String, Optional newVal)
	Dim j, defVal
	j = FindProp(id)
	If j >= 0 Then
		defVal = primProps(j)(2)
	Else 'attribute
		defVal = attrW
	End If
	PropWidth = SetGetProp(PropSection, primName & " Width " & id, newVal, defVal)
End Function

Function PropAlign(id As String, Optional newVal)
	Dim j, defVal
	j = FindProp(id)
	If j >= 0 Then
		defVal = primProps(j)(4)
	Else 'attribute
		defVal = Lf
	End If
	PropAlign = SetGetProp(PropSection, primName & " Align " & id, newVal, defVal)
End Function

Function PropColor(id As String, Optional newVal)
	PropColor = SetGetProp(PropSection, primName & " Color " & id, newVal)
End Function

Function PropBold(id As String, Optional newVal)
	PropBold = CBool(SetGetProp(PropSection, primName & " Bold " & id, newVal, id = primProps(0)(0)))
End Function

Function PropCondFormat(id As String, Optional newVal)
	PropCondFormat = CBool(SetGetProp(PropSection, primName & " Conditional Format " & id, newVal, False))
End Function

Function PropCondition(id As String, Optional newVal)
	PropCondition = SetGetProp(PropSection, primName & " Condition " & id, newVal, 0)
End Function

Function PropValueType(id As String, Optional newVal)
	Dim j, defVal
	j = FindProp(id)
	If j >= 0 Then 
		If primProps(j)(3) = Attr Then
			defVal = 2 'measure by default if attribute type 
		ElseIf primProps(j)(4) = Rt Then
			defVal = 1 'number if is right aligned by default
		Else
			defVal = 0 'otherwise text
		End If
	Else 'attribute is text by default
		defVal = 0
	End If
	PropValueType = SetGetProp(PropSection, primName & " Value Type " & id, newVal, defVal)
End Function

Function PropValue1(id As String, Optional newVal)
	PropValue1 = SetGetProp(PropSection, primName & " Value1 " & id, newVal, "")
End Function

Function PropValue2(id As String, Optional newVal)
	PropValue2 = SetGetProp(PropSection, primName & " Value2 " & id, newVal, "")
End Function

Function PropCaseSensitive(id As String, Optional newVal)
	PropCaseSensitive = CBool(SetGetProp(PropSection, primName & " MatchCase " & id, newVal, True))
End Function

Function PropCondHide(id As String, Optional newVal)
	PropCondHide = CBool(SetGetProp(PropSection, primName & " Conditional Hide " & id, newVal, False))
End Function

Function ListHeader(id As String, Optional newVal)
	Dim j, defVal
	j = FindProp(id)
	defVal = primProps(j)(1)
	ListHeader = SetGetProp(ListSection, primName & " Header " & id, newVal, defVal)
End Function

Function ListNPL(id As String, Optional newVal)
	Dim j, defVal
	j = FindProp(id)
	defVal = primProps(j)(2)
	ListNPL = SetGetProp(ListSection, primName & " List NPL " & id, newVal, defVal)
End Function

Function ListWidth(id As String, Optional newVal)
	ListWidth = SetGetProp(ListSection, primName & " List Width " & id, newVal, 0)
End Function

Function ListItem(n, id As String, Optional newVal)
	ListItem = CBool(SetGetProp(ListSection, primName & " List Item " & n & " " & id, newVal, False))
End Function

Function ConditionalExpression(id As String, value As String)
	Dim v, v1, v2, s, j
	s = CondList (PropCondition(id) * 4 + 1)
	v1 = PropValue1(id)
	v2 = PropValue2(id)
	FormatTemplate s, "_lll_", Trim(Str(Len(v1)))
	If PropValueType(id) = 0 Then 'if text
		v1 = Chr(34) & v1 & Chr(34)
		v2 = Chr(34) & v2 & Chr(34)
		If PropCaseSensitive(id) Then
			FormatTemplate s, "_xxx_", v1
			FormatTemplate s, "_yyy_", v2
			FormatTemplate s, "_zzz_", value
		Else
			FormatTemplate s, "_xxx_", UCase(v1)
			FormatTemplate s, "_yyy_", UCase(v2)
			FormatTemplate s, "_zzz_", "UCase(" & value & ")"
		End If
	ElseIf PropValueType(id) = 1  Then 'numeric
		FormatTemplate s, "_xxx_", v1
		FormatTemplate s, "_yyy_", v2
		j = FindProp(id)
		If j >= 0 Then
			If primProps(j)(4) = Rt Then
				FormatTemplate s, "_zzz_", value
			Else
				FormatTemplate s, "_zzz_", "val(" & value & ")"
			End If
		Else
			FormatTemplate s, "_zzz_", "val(" & value & ")"
		End If
	Else  'measure
		v1 = "Measure(" & Chr(34) & v1 & Chr(34) & ")"
		v2 = "Measure(" & Chr(34) & v2 & Chr(34) & ")"
		FormatTemplate s, "_xxx_", v1
		FormatTemplate s, "_yyy_", v2
		FormatTemplate s, "_zzz_", value
		s = s + " And " + value + ".Name = " + v1 + ".Name"
	End If
	ConditionalExpression = s
End Function 

Function EditCondition
	Dim i
	ReDim CondArray(UBound(CondList) \ 4)
	For i = 0 To UBound(CondArray)
		CondArray(i) = CondList(i * 4)
	Next
	Begin Dialog UserDialog 580,189,"Condition",.EditCondFunc ' %GRID:10,7,1,1
		Text 10,7,560,14,"&Condition will be met if cell value",.ItemLab
		DropListBox 10,28,200,161,CondArray(),.CondList
		GroupBox 10,56,300,35,"Consider cell value as",.GroupBox1
		Text 380,35,30,14,"And",.Text2
		TextBox 220,28,150,21,.val1
		TextBox 420,28,150,21,.val2
		TextBox 220,28,350,21,.val0
		OptionGroup .ValueType
			OptionButton 30,70,70,14,"&Text",.TextOpt
			OptionButton 110,70,80,14,"&Numeric",.NumOpt
			OptionButton 210,70,90,14,"&Measure *",.PhysOpt
		CheckBox 370,63,140,14,"Match &Case",.MatchCase
		OKButton 310,161,120,21
		CancelButton 450,161,120,21
		Text 10,119,540,28,"Notice that you can always modify resulting VB code manually to assign conditions with unlimited level of complexity.",.Text1
	End Dialog
	Dim dlg As UserDialog
	dlg.CondList = PropCondition(curProperty)
	dlg.val0 = PropValue1(curProperty)
	dlg.val1 = PropValue1(curProperty)
	dlg.val2 = PropValue2(curProperty)
	dlg.ValueType = PropValueType(curProperty)
	dlg.MatchCase = PropCaseSensitive(curProperty)
	EditCondition = Dialog(dlg)
End Function

Function CheckNumericValue(value) As Boolean
	Dim v As Double
	On Error GoTo msg
	v = CDbl(value)
	On Error GoTo 0
	CheckNumericValue = True
	Exit Function
msg:
	MsgBox "Incorrect Numeric Value " & value, AppName
	CheckNumericValue = False
End Function

Function CheckMeasure(m) As Boolean
	Dim s1, s2
	CheckMeasure = True
	s1 = "Incorrect Measure Value: "
	s2 = vbCr & vbCr & "Correct measure is a numeric value followed by a physical unit" & vbCr & "e.g: 10mW, 100k, 2M, 200pF, 500nanofarad, etc."
	If m.Name = "Unknown" Or (m.Unit(ppcbMeasureFormatLong) = "" And m.Unit(ppcbMeasureFormatCurrent) = "") Then
		If m.Name = "Unknown" And m.unit <> "" Then
			MsgBox s1 & m.Text & vbCr & "Custom units (" & m.unit & ") are not allowed here" & s2, AppName
		ElseIf m.Name = "Unknown" Then
			MsgBox s1 & m.Text & s2, AppName
		Else
			MsgBox s1 & m.Text & vbCr & "Did you forget physical unit?" & s2, AppName
		End If
		CheckMeasure = False
	End If
End Function

Rem See DialogFunc help topic for more information.
Private Function EditCondFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim entry, j, cur, isTxt, isNum, isAttr, m1, m2, normalized
	Select Case Action%
	Case 1 ' Dialog box initialization
		j = FindProp(curProperty)
		cur = DlgValue("CondList")
		If j >= 0 Then 
			isAttr = primProps(j)(3) = Attr
		Else
			isAttr = True
		End If
		DlgText "ItemLab", "Format column cells only if " & IIf(isAttr, "attribute ", "") & curProperty
		DlgVisible "val0", CondList(cur*4 + 2) = 1
		DlgVisible "val1", CondList(cur*4 + 2) = 2
		DlgVisible "val2", CondList(cur*4 + 2) = 2
		isTxt = (CondList(cur*4 + 3) And 1) <> 0
		isNum = (CondList(cur*4 + 3) And 2) <> 0
		DlgEnable  "TextOpt", isTxt
		DlgEnable  "NumOpt",  isNum
		DlgEnable  "PhysOpt",  IsNum And isAttr 'physical units only for attributes!
		If isTxt And Not isNum Then DlgValue "ValueType", 0
		If isNum And Not isTxt Then DlgValue "ValueType", 1
		DlgEnable  "MatchCase", DlgValue ("ValueType") = 0
	Case 2
		If DlgItem = "OK" Then
			normalized = False
			If CondList(DlgValue("CondList") * 4 + 2) = 1 Then 'number of operands = 1
				If DlgValue("ValueType") = 1 Then 'if numeric
					If Not CheckNumericValue(DlgText("val0")) Then
						EditCondFunc = True
						DlgFocus "val0"
						Exit Function
					End If
					PropValue1 curProperty, Trim(Str(CDbl(DlgText("val0"))))
				ElseIf DlgValue("ValueType") = 2 Then 'if measure
					Set m1 = Measure(DlgText("val0"))
					If Not CheckMeasure(m1) Then
						EditCondFunc = True
						DlgFocus "val0"
						Exit Function
					End If
					m1.Normalize
					If m1.Text <> Measure(DlgText("val0")).Text Then normalized = True
					DlgText "val0", m1.Text
					PropValue1 curProperty, m1.Text
				Else
					PropValue1 curProperty, Trim(DlgText("val0"))
				End If
			Else '2 operands
				If DlgValue("ValueType") = 1 Then 'if numeric
					If Not CheckNumericValue(DlgText("val1")) Then
						EditCondFunc = True
						DlgFocus "val1"
						Exit Function
					ElseIf Not CheckNumericValue(DlgText("val2")) Then
						EditCondFunc = True
						DlgFocus "val2"
						Exit Function
					End If
					PropValue1 curProperty, Trim(Str(CDbl(DlgText("val1"))))
					PropValue2 curProperty, Trim(Str(CDbl(DlgText("val2"))))
				ElseIf DlgValue("ValueType") = 2 Then 'if measure then try to parse it
					Set m1 = Measure(DlgText("val1"))
					Set m2 = Measure(DlgText("val2"))
					If Not CheckMeasure(m1) Then
						EditCondFunc = True
						DlgFocus "val1"
						Exit Function
					End If
					If Not CheckMeasure(m2) Then
						EditCondFunc = True
						DlgFocus "val2"
						Exit Function
					End If
					If m1.Name <> m2.Name Then
						EditCondFunc = True
						MsgBox "Both measure values in the specified range require the same type" & vbCr & vbCr & _
								" The 1st value (" & m1.Text & ") is recognized as a measure of " & m1.Name & vbCr & _
								" The 2nd value (" & m2.Text & ") is recognized as a measure of " & m2.Name, AppName
						DlgFocus "val1"
						Exit Function
					End If
					m1.Normalize
					m2.Normalize
					If m1.Text <> Measure(DlgText("val1")).Text Then normalized = True
					If m2.Text <> Measure(DlgText("val2")).Text Then normalized = True
					DlgText "val1", m1.Text
					DlgText "val2", m2.Text
					PropValue1 curProperty, m1.Text
					PropValue2 curProperty, m2.Text
				Else
					PropValue1 curProperty, Trim(DlgText("val1"))
					PropValue2 curProperty, Trim(DlgText("val2"))
				End If
			End If
			PropCondition curProperty, DlgValue("CondList")
			PropValueType curProperty, DlgValue("ValueType")
			PropCaseSensitive curProperty, DlgValue("MatchCase")
			If normalized Then Wait 0.5 'show user normalized values
		ElseIf DlgItem = "ValueType" Then
			EditCondFunc("", 1, 0)
		ElseIf DlgItem = "CondList" Then
			EditCondFunc("", 1, 0)
			DlgFocus IIf(DlgValue("CondList") > 1, "val0", "val1")
		End If
	End Select
End Function

Sub EditColumn
	Begin Dialog UserDialog 500,287,"Table Column Properties",.ColumnPropFunc ' %GRID:10,7,1,1
		GroupBox 150,98,340,147,"Formatting",.Formatting
		Text 10,14,130,14,"Item:",.ItemLab
		Text 170,14,330,14,"",.ColItem
		Text 10,42,150,14,"Column &Header Text:",.Text3
		TextBox 170,35,320,21,.ColHeader
		Text 10,70,150,14,"&Width in Characters:",.WidLab1
		GroupBox 10,98,130,147,"Text Alignment",.GroupBox1
		TextBox 170,63,60,21,.ColWidth
		Text 240,70,240,14,"(for text formats only)",.WidLab2
		OptionGroup .Align
			OptionButton 30,126,60,14,"&Left",.OptionButton3
			OptionButton 30,196,70,14,"&Right",.OptionButton1
			OptionButton 30,161,80,14,"&Center",.OptionButton2
		OptionGroup .Condition
			OptionButton 170,126,80,14,"Always",.FormatAlways
			OptionButton 250,126,110,14,"Conditionally",.FormatCond
		PushButton 370,119,110,21,"Co&ndition...",.CondBtn
		CheckBox 170,154,120,14,"Colored &Text",.ColColor
		PushButton 310,147,170,21,"&Set Text Color...",.ColorBtn
		CheckBox 170,182,260,14,"&Bold Text ",.Bold
		CheckBox 170,210,300,14,"Use Specified Condition as a &Filter",.CondHide
		PushButton 10,259,140,21,"&Reset This Column",.Reset
		PushButton 160,259,140,21,"Reset &All Columns",.ResetAll
		PushButton 370,259,120,21,"Close",.CloseBtn
		Text 190,224,290,14,"(To hide entire table row if condition not met)",.Text1
	End Dialog
	Dim n
	Dim dlg As UserDialog
	n = Dialog(dlg, 5)
End Sub

Rem See DialogFunc help topic for more information.
Private Function ColumnPropFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim res, j, entry, color, colorsAllowed, widthAllowed, isAttr
	Static hwnd&
	Select Case Action%
	Case 1 ' Dialog box initialization
		j = FindProp(curProperty)
		If j >= 0 Then 
			isAttr = primProps(j)(3) = Attr
		Else
			isAttr = True
		End If
		If SuppValue <> 0 Then 
			hwnd = SuppValue
			DlgText "ItemLab", primName & " Item:"
			DlgText "ColItem", IIf(isAttr, "Attribute " & curProperty, curProperty)
			DlgText "ColHeader", PropHeader(curProperty)
			DlgText "ColWidth", Trim(Str(PropWidth(curProperty)))
			DlgValue "Align", PropAlign(curProperty)
			DlgValue "ColColor", PropColor(curProperty) <> ""
			DlgValue "Bold", IIf(PropBold(curProperty), 1, 0)
			DlgValue "Condition", IIf(PropCondFormat(curProperty), 1, 0)
			DlgValue "CondHide", IIf(PropCondHide(curProperty), 1, 0)
		End If
		widthAllowed = target = WordPad Or target = Notepad
		colorsAllowed = target = WordPad Or target = Browser
		DlgEnable "WidLab1", widthAllowed
		DlgEnable "WidLab2", widthAllowed
		DlgEnable "ColWidth", widthAllowed
		DlgEnable "ColColor", colorsAllowed
		DlgEnable "ColorBtn", colorsAllowed And DlgValue("ColColor") = 1
		DlgEnable "Bold", colorsAllowed
		DlgEnable "CondBtn", DlgValue("Condition") = 1
		DlgEnable "CondHide", DlgValue("Condition") = 1
	Case 2
		If DlgItem$ = "CloseBtn" Then
			If DlgEnable("ColWidth") And Val(DlgText("ColWidth")) <= 0 Then
				MsgBox "Incorrect Width Value", AppName
				ColumnPropFunc = True
				DlgFocus "ColWidth"
			ElseIf Len(Trim(DlgText("ColHeader"))) = 0 Then
				MsgBox "Incorrect Header Text", AppName
				ColumnPropFunc = True
				DlgFocus "ColHeader"
			ElseIf DlgValue("Condition") = 1 And DlgValue("Bold") = 0 And DlgValue("ColColor") = 0 And DlgValue("CondHide") = 0 Then
				MsgBox "You have defined a condition but did not specify formatting option" & vbCr & _
						"Select one of the available options (Bold, Color, or Filter) or turn formatting back to <Always>.", AppName
				ColumnPropFunc = True
				DlgFocus "Bold"
			Else
				PropHeader curProperty, Trim(DlgText("ColHeader"))
				PropWidth curProperty, Val(DlgText("ColWidth"))
				PropAlign curProperty, DlgValue("Align")
				PropBold curProperty, DlgValue("Bold")
				PropCondFormat curProperty, DlgValue("Condition") = 1
				PropCondHide curProperty, DlgValue("CondHide") = 1
			End If
		ElseIf DlgItem$ = "ColColor" Then
			If DlgValue("ColColor") = 0 Then
				PropColor curProperty, Null
			Else
				PropColor curProperty, 0
			End If
			ColumnPropFunc("", 1, 0)
		ElseIf DlgItem$ = "CondBtn" Then
			ColumnPropFunc = True
			EditCondition
		ElseIf DlgItem$ = "Condition" Then
			If DlgValue("Condition") = 1 Then
				If EditCondition = 0 And Not PropCondFormat(curProperty) Then
					DlgValue "Condition", 0
				End If
			End If
			ColumnPropFunc("", 1, 0)
		ElseIf DlgItem$ = "ColorBtn" Then
			ColumnPropFunc = True
			color = PropColor(curProperty)
			If color = "" Then color = 0
			color = AssignColor(hwnd, color)
			PropColor curProperty, color
		ElseIf DlgItem$ = "Reset" Then
			ColumnPropFunc = True
			PropHeader curProperty, Null
			PropWidth curProperty, Null
			PropAlign curProperty, Null
			PropColor curProperty, Null
			PropBold curProperty, Null
			PropCondFormat curProperty, Null
			PropCondition curProperty, Null
			PropValueType curProperty, Null
			PropValue1 curProperty, Null
			PropValue2 curProperty, Null
			PropCaseSensitive curProperty, Null
			PropCondHide curProperty, Null
			ColumnPropFunc("", 1, hwnd)
		ElseIf DlgItem$ = "ResetAll" Then
			ColumnPropFunc = True
			res = MsgBox ("All column formatting settings for all object properties and attributes will be reset to default values." & vbCr & _
							"Continue?", vbExclamation + vbYesNo, AppName)
			If res = vbYes Then
				On Error Resume Next
				DeleteSetting WizName, PropSection
				On Error GoTo 0
				ColumnPropFunc("", 1, hwnd)
			End If
		End If
	End Select
End Function

Sub EditList
	Begin Dialog UserDialog 670,231,"List Properties",.ListPropFunc ' %GRID:10,7,1,1
		Text 10,14,200,14,"Item:",.ItemLab
		Text 210,14,440,14,"",.ListItem
		Text 10,49,110,14,"List &Header Text:",.Text3
		TextBox 130,42,530,21,.ListHeader
		Text 10,84,320,14,"&Number of items per line ( Type 0 for unlimited ):",.WidLab1
		TextBox 340,77,60,21,.NumPerLine
		Text 10,119,310,14,"Maximum Item &Width  ( Type 0 for unlimited ):",.ItemWidthLab1
		TextBox 340,112,60,21,.ItemWidth
		Text 420,112,220,28,"Width > 0 guarantees proper alignment in text formats.",.ItemWidthLab2
		PushButton 10,203,120,21,"&Reset This List",.Reset
		PushButton 140,203,120,21,"Reset &All Lists",.ResetAll
		PushButton 540,203,120,21,"Close",.CloseBtn
		Text 10,154,380,14,"Output additional info associated with each item in the list:",.ItemsLab
		CheckBox 400,154,120,14,"Item1",.Item1
		CheckBox 530,154,140,14,"Item2",.Item2
		CheckBox 400,175,120,14,"Item3",.Item3
		CheckBox 530,175,140,14,"Item4",.Item4
	End Dialog
	Dim n
	Dim dlg As UserDialog
	n = Dialog(dlg, 3)
End Sub

Function ExtractListItem(ByVal s, n)
	Dim p, i
	ExtractListItem = ""
	p = InStr(s, "\")
	For i = 1 To n
		If p = 0 Then Exit Function
		s = Mid(s, p + 1)
		p = InStr(s, "\")
	Next
	If p > 0 Then
		ExtractListItem = Left(s, p-1)
	Else
		ExtractListItem = s
	End If
End Function

Function ExtractListItemLabel(ByVal s, n)
	s = ExtractListItem(s, n)
	ExtractListItemLabel = Left(s, InStr(s, ":") - 1)
End Function 

Function ExtractListItemTemplate(ByVal s, n)
	s = ExtractListItem(s, n)
	ExtractListItemTemplate = Mid(s, InStr(s, ":") + 1)
End Function 

Rem See DialogFunc help topic for more information.
Private Function ListPropFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim entry, i, j, t, res, s, item
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgText "ItemLab", IIf (secObj <> 0, "List of " & primName & " related items:", "Object List:")
		DlgText "ListItem", curProperty
		DlgText "ListHeader", ListHeader(curProperty)
		DlgText "NumPerLine", ListNPL(curProperty)
		DlgText "ItemWidth", ListWidth(curProperty)
		j = FindProp(curProperty)
		t = primProps(j)(3)
		DlgEnable "NumPerLine", t <> Attr And t <> RteSeg
		DlgEnable "ItemWidthLab1", target = Notepad Or target = WordPad
		DlgEnable "ItemWidthLab2", target = Notepad Or target = WordPad
		DlgEnable "ItemWidth", target = Notepad Or target = WordPad
		s = primProps(j)(9)
		DlgVisible "ItemsLab", ExtractListItem(s, 1) <> ""
		item = DlgNumber("Item1") - 1
		For i = 1 To 4
			DlgVisible item + i, ExtractListItem(s, i) <> ""
			If ExtractListItem(s, i) <> "" Then
				DlgText item + i, ExtractListItemLabel(s, i)
				DlgValue item + i, IIf (ListItem(i, curProperty), i, 0)
			End If
		Next
	Case 2
		If DlgItem$ = "CloseBtn" Then
			If Val(DlgText("NumPerLine")) < 0 Then
				MsgBox "Incorrect Number Per Line Value", AppName
				ListPropFunc = True
				DlgFocus "NumPerLine"
			Else
				ListHeader curProperty, Trim(DlgText("ListHeader"))
				ListNPL curProperty, Val(DlgText("NumPerLine"))
				ListWidth curProperty, Val(DlgText("ItemWidth"))
				item = DlgNumber("Item1") - 1
				For i = 1 To 4
					If DlgVisible(item + i) Then
						ListItem i, curProperty, DlgValue(item + i) = 1
					End If
				Next
			End If
		ElseIf DlgItem$ = "Reset" Then
			ListPropFunc = True
			ListHeader curProperty, Null
			ListNPL curProperty, Null
			ListWidth curProperty, Null
			item = DlgNumber("Item1") - 1
			For i = 1 To 4
				ListItem i, curProperty, Null
			Next
			ListPropFunc("", 1, 0)
		ElseIf DlgItem$ = "ResetAll" Then
			ListPropFunc = True
			res = MsgBox ("All custom list settings for all object types will be reset to default values." & vbCr &  _
							"Continue?", vbExclamation + vbYesNo, AppName)
			If res = vbYes Then
				On Error Resume Next
				DeleteSetting WizName, ListSection
				On Error GoTo 0
				ListPropFunc("", 1, 0)
			End If
		End If
	End Select
End Function

Rem See DialogFunc help topic for more information.
Private Function dlgfunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim i, j, n, s, scrName, res, id, count, prevFocus
	Select Case Action%
	Case 1 ' Dialog box initialization
		If SuppValue <> 0 Then
			dlgHandle = SuppValue
		End If
		'UpdateTitle
		'keep dialog position as it was in previous step
		If SuppValue <> 0 And Not (wr.x1 = 0 And wr.x2 = 0) Then
			MoveWindow(dlgHandle, wr.x1, wr.y1, wr.x2 - wr.x1, wr.y2 - wr.y1, False)
		End If
		DlgEnable "Line", False
		If curDlg = "5" Or Left(curDlg, 1) = "6" Or curDlg = "7" Then
			UpdateObjectName
		End If
		
		If curDlg = "1" Then
			DlgText "vbDir", scriptDir
		ElseIf curDlg = "3" Then
			DlgEnable "ContList", DlgValue("Container") = 2
		ElseIf curDlg = "4" Then
			i = DlgValue("primObj")
			DlgEnable "itNet", Not repByContainer 
			DlgEnable "itPin", Not repByContainer 
			DlgEnable "itRouteSegment", Not repByContainer
			DlgEnable "itVia", Not repByContainer 
			DlgEnable "itConnection", Not repByContainer
			DlgEnable "itTestPoint", Not repByContainer
			DlgVisible "installedOnly", AppId = 2 And repByContainer
			DlgVisible "primContTotal", AppId = 1 Or repByContainer
			DlgEnable "primContTotal", repByContainer And (AppId = 1 Or DlgValue("installedOnly"))
		ElseIf curDlg = "5" Then
			DlgText "ObjProps", objProps(primObj)
			DlgText "ObjLists", objLists(primObj)
			DlgEnable "Opt2", primData(primObj)(8) 'there are related objects
			DlgEnable "EditList", DlgValue("TableSec") = 2
		ElseIf  Left(curDlg, 1) =  "6" Then
			curAttrName = ""
			curAttribute = 0
			UpdateAvailProps
			UpdateSelectedProps
			If curDlg = "6B" Then
				DlgVisible "filterByContainer", repByContainer And AppId = 1
			End If
		ElseIf curDlg = "7" Then
			If SuppValue <> 0 Then DlgFocus "Header"
			DlgEnable "Header", DlgValue("GenHdr") = 1
			DlgEnable "IncTime", DlgValue("GenHdr") = 1
			DlgEnable "IncJob", DlgValue("GenHdr") = 1
			DlgEnable "Align", Table
			DlgEnable "Contents", target = Browser And (repByContainer Or secObj <> 0) And primData(primObj)(7) 'index allowed
			DlgVisible "colTbl", Table And target = Browser
			DlgVisible "SetupColors", Table And target = Browser
			DlgEnable  "SetupColors", DlgValue("colTbl") = 1
			DlgVisible "DrawBorderLab", Table And (target = Notepad Or target = WordPad)
			DlgVisible "DrawVBorder", Table And (target = Notepad Or target = WordPad)
			DlgVisible "DrawHBorder", Table And (target = Notepad Or target = WordPad)
			DlgVisible "FilterbyCount", Not Table And secObj <> 0
		ElseIf  curDlg = "8" Then
			DlgEnable "Ext", target = 0 And DlgValue("OutputFile") <> 2
			DlgEnable "CustomName", DlgValue("OutputFile") = 1
			DlgEnable "Clipboard", target <> Browser And target <> WordPad
		End If
	Case 2 ' Value changing or button pressed
		Rem dlgfunc = True ' Prevent button press from closing the dialog box
		If DlgItem$ = "BackBtn" Then
			GetWindowRect(dlgHandle, wr)
		ElseIf DlgItem$ = "NextBtn" Then
			GetWindowRect(dlgHandle, wr)
			If curDlg = "1" Then
				scriptDir = DlgText("vbDir")
			ElseIf Left(curDlg,1) = "6" Then
				If DlgFocus = "AvailPropList" And DlgValue("AvailPropList") >= 0 Then 'double click on left list
					dlgfunc("AddBtn", 2, 0)
					dlgfunc = True
					Exit Function					
				ElseIf DlgFocus = "PropList" And DlgValue("PropList") >= 0 Then 'double click on right list
					dlgfunc("EditBtn", 2, 0)
					dlgfunc = True
					Exit Function					
				ElseIf DlgFocus = "AttrName" Then
					If DlgText("AttrName") <> "" Then 
						UpdateAvailAttrs
						dlgfunc("AddBtn", 2, 0)
						dlgfunc = True
						Exit Function					
					End If
				End If
				
				If UBound(PropArray) = 0 Then
					res = MsgBox ("Please add at least one item", AppName)
					dlgfunc = True
				End If
			End If
		ElseIf DlgItem$ = "FinishBtn" Or DlgItem$ = "RunBtn" Then
			If Not CheckFileName(DlgText("CustomName")) And DlgEnable("CustomName") Then
				res = MsgBox ("Incorrect File Name", vbExclamation, AppName)
				dlgfunc = True
				DlgFocus "CustomName"
				SendKeys "{Home}+{End}"
				Exit Function
			End If
			If Not CheckFileName(DlgText("Ext")) Then
				res = MsgBox ("Incorrect File Extension", vbExclamation, AppName)
				dlgfunc = True
				DlgFocus "Ext"
				SendKeys "{Home}+{End}"
				Exit Function
			End If
			If Not CheckFileName(DlgText("ScriptName")) Then
				res = MsgBox ("Incorrect File Name", vbExclamation, AppName)
				dlgfunc = True
				DlgFocus "ScriptName"
				SendKeys "{Home}+{End}"
				Exit Function
			End If
			scrName = DlgText("ScriptName") & ".bas"
			If Dir(scriptDir & "\" & scrName) <> "" Then
				res = MsgBox ("The script file with the given name" & vbCr & _
								scriptDir & "\" & scrName & vbCr & _
								"already exists. Overwrite?", _
							vbYesNoCancel + vbDefaultButton2 + vbQuestion , AppName)
				If res = vbNo Or res = vbCancel Then 
					dlgfunc = True
					DlgFocus "ScriptName"
					SendKeys "{Home}+{End}"
					Exit Function
				End If
			End If
		ElseIf DlgItem$ = "HelpBtn" Then
			HelpPage
			dlgfunc = True
		ElseIf DlgItem$ = "Browse" Then
			s = BrowseDirectory
			If s <> "" Then
				DlgText "vbDir", s
			End If
			dlgfunc = True
		ElseIf DlgItem$ = "Container" Then
			dlgfunc("",1,0)
			If DlgValue("Container") = 2 Then
				DlgFocus "ContList"
				SendKeys "{Home}+{End}"
			End If
		ElseIf DlgItem$ = "TableSec" Then
			nSteps = IIf (DlgValue("TableSec") = 2, 7, 8)
			dlgfunc("",1,0)
		ElseIf DlgItem$ = "genHdr" Then
			dlgfunc("",1,0)
			If DlgValue("genHdr") = 1 Then
				DlgFocus "Header"
				SendKeys "{Home}+{End}"
			End If
		ElseIf DlgItem$ = "OutputFile" Then
			dlgfunc("",1,0)
			If DlgValue("OutputFile") = 1 Then
				DlgFocus "CustomName"
				SendKeys "{Home}+{End}"
			End If
		ElseIf DlgItem$ = "filterByContainer" Then
			UpdateSelectedProps
			UpdateAvailProps
		ElseIf DlgItem$ = "AvailPropList" Then
			ShowInfo
		ElseIf DlgItem$ = "AddBtn" Then
			dlgfunc = True
			id = DlgValue("AvailPropList")
			'get selected property
			s = ""
			If id >= 0 Then 
				s = DlgText("AvailPropList")
			ElseIf curDlg = "6" Then
				s = curAttrName
			End If
			If s = "" Then
				MsgBox "Please select a property or enter Attribute name", AppName
				Exit Function
			End If
			'is property/attribute already added?
			For i = 1 To UBound(PropArray)
				If UCase(s) = UCase(PropArray(i)) Then
					DlgValue "PropList", i - 1 'select this property
					UpdateButtons
					Exit Function
				End If
			Next
			n = UBound(AvailPropArray) - 1
			count = UBound(PropArray) + 1
			ReDim Preserve PropArray (count)
			PropArray(count) = s
			If id >= 0 Then
				UpdateAvailProps
				DlgValue "AvailPropList", IIf (id = n, id - 1, id)
				If DlgValue("AvailPropList") = -1 And curDlg = "6" Then
					UpdateAvailAttrs
					DlgFocus "AttrName"
				End If
			Else
				If curAttribute <> -1 Then
					curAttribute = curAttribute + 1
					If curAttribute = UBound(AttrArray) Then curAttribute = -1
					DlgValue "AttrName", curAttribute
				End If
				UpdateAvailProps
				UpdateAvailAttrs
				DlgFocus "AttrName"
			End If
			UpdateSelectedProps
			DlgValue "PropList", count - 1
			UpdateButtons
		ElseIf DlgItem$ = "RemoveBtn" Then
			dlgfunc = True
			id = DlgValue("PropList") + 1
			If id < 1 Then Exit Function
			s = PropArray(id)
			If Table And s = primProps(0)(0) Then
				MsgBox "This column is required. You cannot remove it", AppName
				Exit Function
			End If
			RemoveProp(id)
			UpdateSelectedProps
			UpdateAvailProps(s)
			If DlgValue("AvailPropList") <> -1 Then
				DlgFocus "AvailPropList"
			ElseIf curDlg = "6" Then
				UpdateAvailAttrs(s)
				DlgFocus "AttrName"
			End If
			UpdateButtons
		ElseIf DlgItem$ = "ResetBtn" Then
			dlgfunc = True
			If Table Then
				ReDim PropArray (1)
				PropArray(1) = primProps(0)(0)
			Else
				ReDim PropArray (0)
			End If
			dlgfunc("",1,0) 'update left list box
			UpdateSelectedProps
			UpdateButtons
		ElseIf DlgItem$ = "UpBtn" Then
			dlgfunc = True
			id = DlgValue("PropList") + 1
			If id < 1 Then Exit Function
			s = PropArray(id)
			PropArray(id) = PropArray(id - 1)
			PropArray(id - 1) = s
			UpdateSelectedProps
			DlgValue "PropList", id - 2
			UpdateButtons
		ElseIf DlgItem$ = "DownBtn" Then
			dlgfunc = True
			id = DlgValue("PropList") + 1
			If id < 1 Then Exit Function
			s = PropArray(id)
			PropArray(id) = PropArray(id + 1)
			PropArray(id + 1) = s
			UpdateSelectedProps
			DlgValue "PropList", id
			UpdateButtons
		ElseIf DlgItem$ = "EditList" Then
			dlgfunc = True
			primProps = primList(primObj)
			curProperty = primProps(0)(0)
			secObj = 0
			EditList
		ElseIf DlgItem$ = "EditBtn" Then
			dlgfunc = True
			id = DlgValue("PropList") + 1
			If id < 1 Then Exit Function
			curProperty = PropArray(id)
			If Table Then
				EditColumn
				UpdateSelectedProps
			Else
				EditList
			End If
			DlgFocus "PropList"
		ElseIf DlgItem$ = "PropList" Then
			UpdateButtons
		ElseIf DlgItem$ = "colTbl" Then
			DlgEnable "SetupColors", DlgValue("colTbl") = 1
		ElseIf DlgItem$ = "SetupColors" Then
			SetupColors
			dlgfunc = True
		ElseIf DlgItem$ = "primObj" Or DlgItem$ = "installedOnly" Then
			dlgfunc("", 1, 0)			
		End If
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
		If curDlg = "6" Then
			prevFocus = SuppValue
			If DlgFocus = "AttrName" Then
				DlgValue "AvailPropList", -1
				UpdateAvailAttrs
			ElseIf DlgFocus = "AddBtn" And prevFocus = DlgNumber("AttrName") Then
				UpdateAvailAttrs
				DlgValue "AttrName", -1
			ElseIf prevFocus = DlgNumber("AttrName") Then
				curAttrName = ""
				DlgValue "AttrName", -1
			End If
		End If
	Case 5 ' Idle
		Rem dlgfunc = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Sub Abort
	FreeCustomPalette
	End
End Sub

Function ReportAutoName As String
	Dim s1, s2, j
	s1 = IIf (allCont, IIf(AppId = 1, "Sheet by Sheet ", "Assembly Option "), "")
	If Table And FilterCondition <> "" Then
		s1 = s1 + "Conditional "
	End If
	If secObj <> 0 Then
		If nSecObjs = 1 Then
			s2 = PropHeader(PropArray(1))
			If s2 = "" Then 'get default list header
				j = FindProp(PropArray(1))
				s2 = primProps(j)(1)
			End If
			If Right(s2, 1) = ":" Then s2 = Left(s2, Len(s2) - 1)
			ReportAutoName = s1 + s2 + " Report"
		Else
			ReportAutoName = s1 + primName + " Report"
		End If
	Else
		ReportAutoName = s1 + primName + IIf(Table, "", " List") + " Report"
	End If
	If givenCont = ACTIVE_SHEET Then
		ReportAutoName = ReportAutoName + " (Current Sheet)"
	ElseIf givenCont <> "" Then
		ReportAutoName = ReportAutoName + " (" + IIf(AppId = 1, "Sheet ", "") + givenCont + ")"
	End If
End Function

Sub GenerateCode
	If target = Notepad Then
		ext = txtFileExt
	Else
		ext = ExtMap(target)
	End If
	If Len(ext) > 0 And  Left(ext, 1) <> "." Then ext = "." + ext
	If header = "" Then header = ReportAutoName
	If repname = "" Then repname = ReportAutoName
	If scriptName = "" Then scriptName = ReportAutoName
	scriptPath = scriptDir & "\" & scriptName & ".bas"
	TXT = False
	RTF = False
	HTML = False
	XLS = False
	If target = WordPad Then RTF = True
	If target = Browser Then HTML = True
	If target = Notepad Then TXT = True
	If target = Excel   Then XLS = True
	If RTF Or HTML Then
		NewLine = "OutLine"
		OutLine = "OutLine "
		OutBoldLine = "OutBoldLine "
		OutBoldUnLnLine = "OutBoldUnLnLine "
		DLMT= " & "
		Dim i
		colorColumns = False
		For i = 1 To UBound(PropArray)
			If PropColor(PropArray(i)) <> "" Then
				colorColumns = True
				Exit For
			End If
		Next
	Else
		NewLine = "Print #1"
		OutLine = "Print #1, "
		OutBoldLine = "Print #1, "
		OutBoldUnLnLine = "Print #1, "
		DLMT= "; "
	End If
	Spc = Chr(34) & " " & Chr(34)
	If HTML Then
		NewRow = "Print #1, NewRow"
		ListIndent = ""
		Spc = "Spc"
	Else
		NewRow = NewLine
		ListIndent = vbTab
	End If
	If XLS Then
		Spc = "vbTab"
		AttrDLMT = DLMT & Chr(34) & vbTab & Chr(34) & DLMT
		If Not table Then
			OutBoldLine = "Print #1, Space(1); "
			OutBoldUnLnLine = OutBoldLine
		End If
		ext = ".txt"
	Else
		AttrDLMT = DLMT & Chr(34) & vbTab & "=" & vbTab & Chr(34) & DLMT
	End If
	On Error GoTo OpenError
	Open scriptPath For Output As #1
	On Error GoTo 0
	GenTextCode
	Close #1
	Exit Sub
OpenError:
	MsgBox "Could not open file for writing", vbCritical, AppName
	End
End Sub

Sub 	GenTextCode
	Print #1, "'This script has been generated by "; AppName; "'s VB Script Wizard on "; Now
	Print #1, "'It will create reports in "; formatName(target)
	Print #1, "'You can use the following code as a skeleton for your own VB scripts"
	GenDeclarationCode
	Print #1
	Print #1, "Sub Main"
		nTabs = 1
		OpenOutputFile
		BeginFormatFile
		If Not XLS Or Not Table Then	
			OutputHeader
			OutputCondition
		End If
		Print #1
		BeginProgress
		TableOfContents
		If Table And XLS Then TableHeader
		ForEachContainer
			If Table Then
				BeginTable
					If Not XLS Then	TableHeader
					Print #1, Indent; "'Output table rows"
					ForPrimObj
						TableRow
					NextPrimObj
				EndTable
			Else
				ForPrimObj
					OutPrimObj
					SecondaryObjects
				NextPrimObj
			End If
		NextContainer
		Print #1
		If Not XLS Or Not Table Then
			OutputTotal
		End If
		EndProgress
		EndFormatFile
		CloseAndShowOutputFile
	Print #1, "End Sub"

	Close #1
	Open scriptPath For Input As #1
	Dim size As Long
	size = LOF(1)
	code = Input(size,1)
	Close #1
	Open scriptPath For Append As #1

	GenerateAdditionalFunctions	
End Sub

Sub ForEachContainer
	If allCont Then 
		Print #1, Indent; "For Each "; contObj; " in "; contColl
		nTabs = nTabs + 1
		If HTML Then
			Print #1, Indent; NewLine
			If contContents Then
				Print #1, Indent; "OutBookmark "; contNameProp
			End If
			Print #1, Indent; "OutTitle "; Chr(34); ContName(0); ":"; Chr(34); DLMT; Spc; DLMT; contNameProp; ", 5"
		ElseIf RTF Then
			Print #1, Indent; NewLine
			Print #1, Indent; OutBoldUnLnLine; Chr(34); ContName(0); ": "; Chr(34); DLMT; contNameProp
			Print #1, Indent; NewLine
		ElseIf TXT Then
			Print #1, Indent; NewLine
			Print #1, Indent; OutLine; Chr(34); ContName(0); ": "; Chr(34); DLMT; contNameProp
			If Not Table Then
				Print #1, Indent; OutLine; "String(7 + Len("; contNameProp; "), ""-"")"
			Else
				Print #1, Indent; NewLine
			End If
		ElseIf XLS And Not Table Then
			Print #1, Indent; NewLine
			Print #1, Indent; OutBoldLine; Chr(34); ContName(0); ": "; Chr(34); DLMT; contNameProp
			Print #1, Indent; NewLine
		End If
		If primContTotal And (Not XLS Or Not Table) Then
			Print #1, Indent; "Subtotal = 0"
		End If
	End If
End Sub

Sub NextContainer
	If repByContainer Then 
		If primContTotal And (Not XLS Or Not Table) Then
			If FilterCondition = "" Then
				Print #1, Indent; OutBoldLine; Chr(34); primName; " count in the "; ContName(1); ": "; ListIndent; Chr(34); DLMT; "Subtotal"
			Else
				Print #1, Indent; OutBoldLine; Chr(34); "Number of "; primName; "s in the "; ContName(1); " matching requested conditions: "; Chr(34); DLMT; "Subtotal"
			End If
			Print #1, Indent; NewLine
		End If
	End If
	If allCont Then
		nTabs = nTabs - 1
		Print #1, Indent; "Next "; contObj
	End If
End Sub

Sub ForPrimObj
	If secObj = 0 Then
		numPerLine = ListNPL(primProps(0)(0))
		If ListHeader(primProps(0)(0)) <> "" Then
			Print #1, Indent; OutBoldLine; Chr(34) & ListHeader(primProps(0)(0)) & Chr(34)
		End If
	End If
	If numPerLine > 1 And secObj = 0 And Not Table Then 
		Print #1, Indent; "n = 0"
		BeginTable
	End If
	FormatTemplate primSortColl, "xxx", contObj
	Print #1, Indent; "For Each "; primVarName; " in "; primSortColl
	nTabs = nTabs + 1
	If FilterCondition <> "" Then
		OutputMultiLineCode "If " & FilterCondition & " Then"
		nTabs = nTabs + 1
	End If
	If bInstalledOnly And primObj <> ppcbPartTypes Then
		Print #1, Indent; "If "; primVarName; ".Installed Then"
		nTabs = nTabs + 1
	End If	
	If Table And Not XLS And (enableAlign Or TXT Or RTF) Then
		Print #1, Indent; "CurCol = 0"
	End If
End Sub

Sub OutPrimObj
	If secObj <> 0 Then
		If HTML And contContents And (repByContainer Or secObj <> 0) Then
			Print #1, Indent; "OutBookmark "; IIf(allCont And repByContainer, contNameProp & " & ", ""); primNameProp
		End If
		Print #1, Indent; OutBoldLine; Chr(34) ; primName; Chr(34) ; DLMT; Spc; DLMT; primNameProp
	Else
		If numPerLine > 1 Then 
			Print #1, Indent; "if n > 0 And n Mod "; Trim(Str(numPerLine)); " = 0 then "; NewRow
		End If
		Dim item, width, curOutType, dlm, id
		id = primProps(0)(0)
		curOutType	= primProps(0)(3)
		width = ListWidth(id)
		item = ExtractItems(id)
		FormatTemplate item, "xxx", primVarName
		If curOutType = RteSeg Then
			dlm = " & "
			FormatTemplate item, "RteSeg" & dlm & Spc & dlm, ""
'			Print #1, Indent; "Print #1, OutRSegName("; primVarName; "); " & Spc & ";"
			Print #1, Indent; "Print #1, OutRSegName("; primVarName; "); "
			numPerLine = "2"
			If item <> "RteSeg" Then OutListItem item, width
			Print #1, Indent; OutLine
'			OutRouteSegment item, width	' TBD
			numPerLine = "1"
		Else
			OutListItem item, width
		End If
	End If
End Sub

Sub NextPrimObj
	If secObj <> 0  Then	
		Print #1, Indent; NewLine
	End If
	If numPerLine > 1 And secObj = 0 And Not Table Then 
		Print #1, Indent; "n = n + 1"
	End If
	If primTotal And Not repByContainer And FilterCondition <> "" Then
		Print #1, Indent; "Count = Count + 1"
	End If
	If repByContainer And primContTotal And (Not XLS Or Not Table) Then
		Print #1, Indent; "Subtotal = Subtotal + 1"
	End If
	If bInstalledOnly And primObj <> ppcbPartTypes Then 
		nTabs = nTabs - 1
		Print #1, Indent; "End If"		
	End If		
	If FilterCondition <> "" Then
		nTabs = nTabs - 1
		Print #1, Indent; "End If"
	End If
	If showProgress Then
		Print #1, Indent; "Processed = Processed + 1"
		Print #1, Indent; "ProgressBar = Processed * 100 / Total
	End If
	nTabs = nTabs - 1
	Print #1, Indent; "Next "; primVarName
	If secObj = 0 And Not Table Then 
		If numPerLine > 1 Then 
			If HTML Then
				EndTable
			Else
				Print #1, Indent; "if n > 0 then "; NewLine
			End If
		ElseIf numPerLine = 0 Then
			Print #1, Indent; NewLine
		End If
	End If
	If TXT Or RTF Then
		If Table And drawHBorder Then
			Print #1, Indent; OutLine; "String(L, ""-"")"
		Else
			Print #1, Indent; NewLine
		End If
	End If
End Sub

Sub PropertyTable
	TableHeader
End Sub

Sub BeginTable
	If HTML Then
		Print #1, Indent; "BeginTable"
	End If
End Sub

Sub EndTable
	If HTML Then
		Print #1, Indent; "EndTable"
	End If
End Sub

Sub BeginRow
	If HTML Then
		Print #1, Indent; "BeginRow"
	End If
End Sub

Sub EndRow
	If HTML Then
		Print #1, Indent; "EndRow"
	Else
		Print #1, Indent; NewLine
	End If
End Sub

Sub OutputCondition
	Dim i, id, s, j, cond, value
	If Not table Then Exit Sub 'condition header for list are not ready yet
	If FilterCondition <> "" Then
		If Not XLS Then Print #1, Indent; "'Output Conditions
		Print #1, Indent; OutBoldLine; Chr(34); "Requested Condition:"; Chr(34)
		s = Chr(34)
		For i = 1 To UBound(PropArray)
			id = PropArray(i)
			j = FindProp(id)
			If Table And PropCondFormat(id) And PropCondHide(id) Then
				If s <> Chr(34) Then s = s + "; "
				s = s + IIf(j >= 0, id, "Attribute " & "'" & id & "'")
				cond = PropCondition(id)
				s = s + " " + CondList(cond * 4)
				value = PropValue1(id)
				If PropValueType(id) = 0 Then value = "'" + value + "'"
				s = s + " " + value
				If CondList(cond * 4 + 2) = 2 Then
					value = PropValue2(id)
					If PropValueType(id) = 0 Then value = "'" + value + "'"
					s = s + " and " + value
				End If
			ElseIf Not Table And filterByCount Then
				If s <> Chr(34) Then s = s + "; "
				s = s + primName + " has " + primProps(j)(0)
			End If
		Next
		s = s + Chr(34)
		Print #1, Indent; OutLine; s
		Print #1, Indent; NewLine
	End If
End Sub

Function FilterCondition
	Dim i, j, id, s, template
	s = ""
	For i = 1 To UBound(PropArray)
		id = PropArray(i)
		j = FindProp(id)
		If Table And PropCondFormat(id) And PropCondHide(id) Then
			If s <> "" Then s = s + " _ And "
			If j >= 0 Then
				template = primProps(j)(5)
				FormatTemplate template, "AttrVal", "AttrMeas" 'in case of standard attribute such as Value, Tol. we need RealValue here instead of text
				FormatTemplate template, "xxx", primVarName
				FormatTemplate template, "cont", contObj
			ElseIf PropValueType(id) = 2 Then 'measure attribute
				template = "AttrMeas(" & primVarName & ", " & Chr(34) & id & Chr(34) & ")"
			Else
				template = "AttrVal(" & primVarName & ", " & Chr(34) & id & Chr(34) & ")"
			End If
			s = s + ConditionalExpression(id, template)
		ElseIf Not Table And secObj <> 0 And filterByCount Then
			If s <> "" Then s = s + " Or "
			template = primProps(j)(IIf(shtFilter, 7, 5))
			FormatTemplate template, "xxx", primVarName
			FormatTemplate template, "cont", contObj
			s = s + template + ".Count > 0"
		End If
	Next
	FilterCondition = s
End Function

Dim formatParam, formatCond
Sub OutCell (func As String, s As String)
	If (HTML Or RTF) And PropCondFormat(curProperty) And formatParam <> "" And formatCond <> "" Then
		If PropCondHide(curProperty) Then 'if this condition is also used as a filter then formatParam contain only True(s)
			Print #1, Indent; "OkFmt = "; formatCond
			Print #1, Indent; func; " "; s; formatParam
		ElseIf PropValueType(curProperty) = 2  Then 'measure attribute
			Print #1, Indent; "Set M = AttrMeas("; primVarName; ", "; Chr(34); curProperty; Chr(34);  ")"
			Print #1, Indent; "OkFmt = "; formatCond
			Print #1, Indent; func; " M.Text"; formatParam
		Else
			Print #1, Indent; "value = "; s
			Print #1, Indent; "OkFmt = "; formatCond
			Print #1, Indent; func; " value"; formatParam
		End If
	Else
		Print #1, Indent; func; " "; s; formatParam
	End If
End Sub

Sub OutputCell (s As String)
	OutCell "OutCell", s
End Sub

Sub OutputAttr (s As String)
	OutCell "OutCell", "AttrVal(" & primVarName & ", " & Chr(34) & s & Chr(34) & ")"
End Sub

Sub TableHeader
	curProperty = ""
	formatParam = ""
	formatCond = ""
	Print #1, Indent; "'Output table header"
	BeginRow
	If drawHBorder And (TXT Or RTF) Then
		If drawVBorder Then
			Print #1, Indent; "L = UBound(Columns) * 3 + 2"
		Else
			Print #1, Indent; "L = UBound(Columns)"
		End If
	End If
	If Not XLS And (enableAlign Or RTF Or TXT) Then
		Print #1, Indent; "CurCol = 0"
	End If
	Print #1, Indent; "For i = 0 to UBound(Columns)"
		nTabs = nTabs + 1
		If RTF Or HTML Then formatParam = ", True"
		OutputCell "Columns(i)"
		If drawHBorder And (TXT Or RTF) Then
			Print #1, Indent; "L = L + Widths(i)"
		End If
		nTabs = nTabs - 1
	Print #1, Indent; "Next"
	EndRow
	If drawHBorder And (TXT Or RTF) Then
		Print #1, Indent; OutLine; "String(L, ""-"")"
	End If
End Sub

Sub FormatTemplate(ByRef template, s1, s2)
	Dim p
	p = 1
	While InStr(p, template, s1) <> 0
		p = InStr(p, template, s1)
		template = Left(template, p - 1) & s2 & Mid(template, p + Len(s1))
		p = p + 2
	Wend
End Sub

Sub TableRow
	Dim i, j, id, curW, template, col, colIdx
	colIdx = 0
	If HTML And colorTable Then 'row color go first in table
		If Not bkgDef Then colIdx = colIdx + 3
		If Not fontDef Then colIdx = colIdx + 3
	End If 
	BeginRow
		If allCont And XLS Then
			OutputCell contNameProp
		End If
		For i = 1 To UBound(PropArray)
			curProperty = PropArray(i)
			formatParam = ""
			formatCond = ""
			If HTML Or RTF Then
				template = IIf(PropValueType(curProperty) = 2, "M", "value") 'if measure use M as object var
				If PropBold(curProperty) Then
					If PropCondFormat(curProperty) And Not PropCondHide(curProperty) Then
						formatCond = ConditionalExpression(curProperty, template)
						formatParam = ", OkFmt" 
					Else
						formatParam = ", True"
					End If
				End If
				If PropColor(curProperty) <> "" Then
					If formatParam = "" Then formatParam = ", "
					If PropCondFormat(curProperty) And Not PropCondHide(curProperty) Then
						formatCond = ConditionalExpression(curProperty, template)
						formatParam = formatParam + ", IIf(OkFmt, " & colIdx & ", -1)"
					Else
						formatParam = formatParam + ", " & colIdx
					End If
					colIdx = colIdx + 1
				End If
			End If
			id = FindProp(curProperty)
			If id >= 0 Then
				template = primProps(id)(5)
				FormatTemplate template, "xxx", primVarName
				FormatTemplate template, "cont", contObj
				OutputCell template
			Else						
				OutputAttr curProperty		'regular attribute
			End If
		Next
	EndRow
End Sub

Function ExtractItems(id)
	Dim n, s, j, subItem, dlm
	ExtractItems = ""
	j = FindProp(id)
	s = primProps(j)(9)
	n = 0
	While ExtractListItem(s, n) <> ""
		If n = 0 Or ListItem(n, id) Then
			subItem = ExtractListItemTemplate(s, n)
			dlm = IIf (NumPerLine = 1 Or (ListWidth(id) > 0 And (TXT Or RTF)), " & ", "; ")
			If ExtractItems <> "" Then ExtractItems = ExtractItems + dlm + Spc + dlm
			ExtractItems = ExtractItems + subItem
		End If
		n = n + 1
	Wend
End Function

Dim curOutType, curSecObj

Sub SecondaryObjects
	Dim i, item, width, id, j, coll, dlm
	If secObj = 0 Then Exit Sub
	BeginIndent
	For i = 1 To UBound(PropArray)
		id 			= PropArray(i)
		j	 		= FindProp(id)
		numPerLine  = ListNPL(id)
		width 		= ListWidth(id)
		curOutType	= primProps(j)(3)
		curSecObj	= primProps(j)(4)
		If Not shtFilter Then
			coll	= IIf (sortObj And primProps(j)(6) <> "", primProps(j)(6), primProps(j)(5))
		Else
			coll	= IIf (sortObj And primProps(j)(8) <> "", primProps(j)(8), primProps(j)(7))
		End If
		item		= ExtractItems(id)
		dlm = " & "
		FormatTemplate item, "RteSeg" & dlm & Spc & dlm, ""
		FormatTemplate item, "RteSeg", ""
		FormatTemplate item, "yyy", curSecObj
		FormatTemplate coll, "xxx", primVarName
		FormatTemplate coll, "cont", contObj
		ForObj curSecObj, ListHeader(id), coll
			If curOutType = Attr Then
				OutAttr
				EndCount
			ElseIf curOutType = RteSeg Then
				OutRouteSegment item, width
				EndCount
			Else
				BeginCount
				OutListItem item, width
				EndCount
			End If
		NextObj
	Next
	EndIndent
End Sub

Sub ForObj (obj As String, contHeader As String, InColl As String)
	If contHeader <> "" Then
		Print #1, Indent; OutBoldLine; Chr(34); ListIndent; contHeader; Chr(34)
	End If
	If numPerLine > 1 Or curOutType = Attr Or curOutType = RteSeg Then
		BeginTable
	End If
	If numPerLine > 1 Or secTotal Then 
		Print #1, Indent; "n = 0"
	End If
	Print #1, Indent; "For Each "; obj; " In "; InColl
	nTabs = nTabs + 1
End Sub

Sub NextObj
	nTabs = nTabs - 1
	Print #1, Indent; "Next"
	If numPerLine <> 1 And Not HTML And secTotal Then
		Print #1, Indent; "if n > 0 then "; NewLine
	ElseIf numPerLine = 0 Then
		Print #1, Indent; NewLine
	ElseIf numPerLine <> 1 And Not HTML Then
		Print #1, Indent; NewLine
	End If
	If numPerLine > 1 Or curOutType = Attr Or curOutType = RteSeg Then
		EndTable
	End If
	If secTotal Then
		Print #1, Indent; OutBoldLine; Chr(34); ListIndent; "Count:	"; Chr(34); DLMT; "n"
	End If
	If nSecObjs > 1 Then
		Print #1, Indent; NewLine
		If HTML Then Print #1, Indent; NewLine 'double new line for HTML
		Print #1
	End If
End Sub

Sub BeginCount
	If numPerLine > 1 Then 
		Print #1, Indent; "if n > 0 And n Mod "; Trim(Str(numPerLine)); " = 0 then "; NewRow
	End If
End Sub

Sub EndCount
	If numPerLine > 1 Or secTotal Then 
		Print #1, Indent; "n = n + 1"
	End If
End Sub

Sub OutListItem(item, width)
	Dim Spc1, Spc2
	If HTML Then
		Spc1 = IIf(numPerLine > 1, "NewCell; ", "Spc; ")
	Else
		Spc1 = IIf(secObj <> 0, "vbTab; ", "")
	End If
	If (RTF Or TXT) And numPerLine <> 1 Then 
		Spc2 = "; " & Spc & "; " & IIf (secObj = 0, "vbTab; ", "")
	Else
		Spc2 = "; "
	End If
	If numPerLine = 1 Then
		Print #1, Indent; OutLine; Spc; DLMT; item
	ElseIf width = 0 Or Not (TXT Or RTF) Then
		Print #1, Indent; "Print #1, "; Spc1; item; Spc2
	Else
		Print #1, Indent; "Print #1, "; Spc1; "Left("; item; " + Space(256), "; width; ")"; Spc2
	End If
End Sub

Sub OutRouteSegment(item, width)
	Print #1, Indent; "Print #1, OutRSegName(seg); "
	numPerLine = "2"
	If item <> "RteSeg" And item <> "" Then OutListItem item, width
	numPerLine = "1"
	If HTML And primObj = ppcbConnections Then
		Print #1, Indent; "Print #1, NewRow"
	Else
		Print #1, Indent; OutLine
	End If
End Sub

Sub OutAttr
	If HTML Then
		Print #1, Indent; "Print #1, NewCell; attr.Name; NewCell; ""=""; NewCell; attr.Value; NewRow"
	Else
		Print #1, Indent; OutLine;  "vbTab"; DLMT; "attr.Name"; AttrDLMT; "attr.Value"
	End If
End Sub

Function Indent As String
	Indent = String$(nTabs, vbTab)
End Function

Sub GenCodeForSortedPins
	Print #1
	Print #1, 	"'Pins are not sorted by default (performance issue), so sort them explicitly in report"
	Print #1, 	"Function GetSortedPins(obj As Object)"
	Print #1, 	"	Set GetSortedPins = obj.Pins"
	Print #1, 	"	GetSortedPins.Sort"
	Print #1, 	"End Function"
End Sub

Sub GenCodeForPinsBySheet
	Print #1
	Print #1,	"'Returns collection of object's pins located on the given sheet
	Print #1,  "Function GetPinsBySheet(obj As Object, aSheet As object, Optional Sorted As Boolean = False)"
	Print #1,	"	Set GetPinsBySheet = ActiveDocument.GetObjects(0)
	Print #1,	"	For Each aPin In obj.Pins
	Print #1,	"		If IsGateOnSheet(aPin.Gate, aSheet) Then
	Print #1,	"			GetPinsBySheet.Add aPin
	Print #1,	"		End If
	Print #1,	"	Next
	Print #1,	"	If Sorted Then GetPinsBySheet.Sort
	Print #1,  "End Function"
End Sub

Sub GenCodeForSignalPins
	Print #1
	Print #1,	"'Returns collection of signal pins in the given part
	Print #1,	"Function GetSignalPins(obj As Object, Optional Sorted As Boolean = False)
	Print #1,	"	Set GetSignalPins = ActiveDocument.GetObjects(0)
	Print #1,	"	For Each aPin In obj.Pins
	Print #1,	"		If aPin.Gate Is Nothing And Not aPin.Net Is Nothing Then
	Print #1,	"			GetSignalPins.Add aPin
	Print #1,	"		End If
	Print #1,	"	Next
	Print #1,	"	If Sorted Then GetSignalPins.Sort
	Print #1,  "End Function
End Sub

Sub GenCodeForSignalNetPins
	Print #1
	Print #1,	"'Returns collection of signal pins in the given net
	Print #1,	"Function GetSignalNetPins(obj As Object, Optional Sorted As Boolean = False)
	Print #1,	"	Set GetSignalNetPins = ActiveDocument.GetObjects(0)
	Print #1,	"	For Each aPin In obj.Pins
	Print #1,	"		If aPin.Gate Is Nothing Then
	Print #1,	"			GetSignalNetPins.Add aPin
	Print #1,	"		End If
	Print #1,	"	Next
	Print #1,	"	If Sorted Then GetSignalNetPins.Sort
	Print #1,  "End Function
End Sub

Sub GenCodeForRSegName
	Dim Spc0, Spc1, Spc2
	If HTML And primObj = ppcbConnections Then
		Spc0 = "NewCell & "
		Spc1 = " & NewCell & "
		Spc2 = " & NewCell"
	Else
		Spc0 = ""
		Spc1 = " & " & Spc & " & "
		Spc2 = " & " & Spc & " & " & Spc
	End If
	Print #1
	Print #1,	"Function OutRSegName(seg As Object)
	Print #1,	"	pnt = seg.Points"
	Print #1,	"	OutRSegName = " & Spc0 & "              Format(pnt(1,1), ""0.000"")" & Spc1 & "Format(pnt(1,2), ""0.000"")" & Spc2
	Print #1,	"	OutRSegName = OutRSegName + Format(pnt(2,1), ""0.000"")" & Spc1 & "Format(pnt(2,2), ""0.000"")" & Spc2
	Print #1,	"	If UBound(pnt, 1) = 3 Then
	Print #1,	"		OutRSegName = OutRSegName & Format(pnt(3,1), ""0.000"")" & Spc1 & "Format(pnt(3,2), ""0.000"")" & Spc2
	Print #1,	"	End If
	Print #1,  "End Function"
End Sub

Sub GenCodeForViaName
	Print #1
	Print #1,	"Function OutViaName(aVia As Object)
	Print #1,	"	OutViaName = Format(aVia.PositionX, ""0.000##"") & " & Spc & " & Format(aVia.PositionY, ""0.000##"")"
	Print #1,  "End Function"
End Sub

Sub GenCodeForGetOptName
	Print #1
	Print #1,  "Function GetOptName(opt As Object)"
	Print #1,  "	GetOptName = Left(opt.Name, Len(opt.Name) - (Len(ActiveDocument.Name) + 1))"
	Print #1,  "End Function"
End Sub

Sub GenCodeForGetSegmentCons
	Print #1
	Print #1,  "Function GetSegmentCons(seg As Object) As objects"
	Print #1,  "	'Create empty object collection"
	Print #1,  "	Set GetSegmentCons = ActiveDocument.GetObjects(0)"
	Print #1,  "	For Each con In seg.Net.Connections"
	Print #1,  "		Set aSeg = con.RouteSegments(seg.Name)"
	Print #1,  "		If Not aSeg Is Nothing Then GetSegmentCons.Add(con)"
	Print #1,  "	Next con"
	Print #1,  "End Function"
End Sub

Sub GenCodeForConnectedPins
	Print #1
	Print #1,	"'Returns collection of connected pins in the given object
	Print #1,	"Function GetConPins(obj as Object, Optional Sorted As Boolean = False)
	Print #1,	"	Set GetConPins = ActiveDocument.GetObjects(0) 'create empty collection
	Print #1,	"	For Each aPin In obj.Pins
	Print #1,	"		If Not aPin.Net Is Nothing Then
	Print #1,	"			GetConPins.Add aPin
	Print #1,	"		End If
	Print #1,	"	Next
	Print #1,	"	If Sorted Then GetConPins.Sort
	Print #1,  "End Function
End Sub

Sub GenCodeForConnectedPinsBySheet
	Print #1
	Print #1,	"'Returns collection of connected part/gate pins located on the given sheet
	Print #1,	"Function GetConPinsBySheet(obj As Object, aSheet As Object, Optional Sorted As Boolean = False)
	Print #1,	"	Set GetConPinsBySheet = ActiveDocument.GetObjects(0)
	Print #1,	"	For Each aPin In obj.Pins
	Print #1,	"		If IsGateOnSheet(aPin.Gate, aSheet) And Not aPin.Net Is Nothing Then
	Print #1,	"			GetConPinsBySheet.Add aPin
	Print #1,	"		End If
	Print #1,	"	Next
	Print #1,	"	If Sorted Then GetConPinsBySheet.Sort
	Print #1,  "End Function
End Sub

Sub GenCodeForUnconnectedPins
	Print #1
	Print #1,	"'Returns collection of unconnected pins in the given object
	Print #1,	"Function GetUnconPins(obj As Object, Optional Sorted As Boolean = False)
	Print #1,	"	Set GetUnconPins = ActiveDocument.GetObjects(0)
	Print #1,	"	For Each aPin In obj.Pins
	Print #1,	"		If aPin.Net Is Nothing Then
	Print #1,	"			GetUnconPins.Add aPin
	Print #1,	"		End If
	Print #1,	"	Next
	Print #1,	"	If Sorted Then GetUnconPins.Sort
	Print #1,  "End Function
End Sub

Sub GenCodeForUnconnectedPinsBySheet
	Print #1
	Print #1,	"'Returns collection of unconnected part/gate pins located on the given sheet
	Print #1,	"Function GetUnconPinsBySheet(obj As Object, aSheet As Object, Optional Sorted As Boolean = False)
	Print #1,	"	Set GetUnconPinsBySheet = ActiveDocument.GetObjects(0)
	Print #1,	"	For Each aPin In obj.Pins
	Print #1,	"		If IsGateOnSheet(aPin.Gate, aSheet) And aPin.Net Is Nothing Then
	Print #1,	"			GetUnconPinsBySheet.Add aPin
	Print #1,	"		End If
	Print #1,	"	Next
	Print #1,	"	If Sorted Then GetUnconPinsBySheet.Sort
	Print #1,  "End Function
End Sub

Sub GenCodeForUsedPins
	Print #1
	Print #1,	"'Returns collection of used pins in the given part
	Print #1,	"Function GetUsedPins(obj As Object, Optional Sorted As Boolean = False)
	Print #1,	"	Set GetUsedPins = ActiveDocument.GetObjects(0)
	Print #1,	"	For Each aPin In obj.Pins
	Print #1,	"		If IsPinUsed(aPin) Then
	Print #1,	"			GetUsedPins.Add aPin
	Print #1,	"		End If
	Print #1,	"	Next
	Print #1,	"	If Sorted Then GetUsedPins.Sort
	Print #1,  "End Function
End Sub

Sub GenCodeForUsedPinsBySheet
	Print #1
	Print #1,	"'Returns collection of used part pins located on the given sheet
	Print #1,	"Function GetUsedPinsBySheet(obj As Object, aSheet As Object, Optional Sorted As Boolean = False)
	Print #1,	"	Set GetUsedPinsBySheet = ActiveDocument.GetObjects(0)
	Print #1,	"	For Each aPin In obj.Pins
	Print #1,	"		If IsGateOnSheet(aPin.Gate, aSheet) And IsPinUsed(aPin) Then
	Print #1,	"			GetUsedPinsBySheet.Add aPin
	Print #1,	"		End If
	Print #1,	"	Next
	Print #1,	"	If Sorted Then GetUsedPinsBySheet.Sort
	Print #1,  "End Function
End Sub

Sub GenCodeForUnusedPins
	Print #1
	Print #1,	"'Returns collection of unused pins in the given part
	Print #1,	"Function GetUnusedPins(obj As Object, Optional Sorted As Boolean = False)
	Print #1,	"	Set GetUnusedPins = ActiveDocument.GetObjects(0)
	Print #1,	"	For Each aPin In obj.Pins
	Print #1,	"		If Not IsPinUsed(aPin) Then
	Print #1,	"			GetUnusedPins.Add aPin
	Print #1,	"		End If
	Print #1,	"	Next
	Print #1,	"	If Sorted Then GetUnusedPins.Sort
	Print #1,  "End Function
End Sub

Sub GenCodeForPartGatesBySheet
	Print #1
	Print #1,	"'Returns collection of used part gates located on the given sheet
	Print #1,	"Function GetPartGatesBySheet(obj As Object, aSheet As Object)
	Print #1,	"	Set GetPartGatesBySheet = ActiveDocument.GetObjects(0)
	Print #1,	"	For Each aGate In obj.Gates
	Print #1,	"		If IsGateOnSheet(aGate, aSheet) Then
	Print #1,	"			GetPartGatesBySheet.Add aGate
	Print #1,	"		End If
	Print #1,	"	Next
	Print #1,  "End Function
End Sub

Sub GenCodeForPowerPins
	Print #1,
	Print #1,	"'Returns collection of power pins in the given part
	Print #1,	"Function GetPowerPins(doc As Object, part As Object, Optional Sorted As Boolean = False)
	Print #1,	"	Set GetPowerPins = doc.GetObjects(0)
	Print #1,	"	For Each aPin In part.Pins
	Print #1,	"		If Not aPin.Net Is Nothing Then
	Print #1,	"			If aPin.Net.Power Then GetPowerPins.Add aPin
	Print #1,	"		End If
	Print #1,	"	Next
	Print #1,	"	If Sorted Then GetPowerPins.Sort
	Print #1,	"End Function
End Sub

Sub GenCodeForTPName
	Print #1
	Print #1,	"'Returns test point's name"
	Print #1,	"Function GetTPName(tp As Object) As String"
	Print #1,	"	If tp.ObjectType = ppcbObjectTypePin Then"
	Print #1,	"		GetTPName = tp.Name"
	Print #1,	"	Else"
	Print #1,	"		GetTPName = tp.type"
	Print #1,	"	End If"
	Print #1,	"End Function"
End Sub

Sub GenCodeForGetTPs
	Print #1
	Print #1,	"'Returns collection of test points in the given object
	Print #1,	"Function GetTestPoints(doc As Object, Optional Sorted As Boolean = False)
	Print #1,	"	Set GetTestPoints = doc.GetObjects(0)
	Print #1,	"	Set nets = doc.Nets
	Print #1,	"	If Sorted Then nets.Sort
	Print #1,	"	For Each aNet In nets
	Print #1,	"		Set pins = aNet.Pins
	Print #1,	"		If Sorted Then pins.Sort
	Print #1,	"		For Each aPin In pins
	Print #1,	"			If aPin.TestPoint <> ppcbTestPointNone Then
	Print #1,	"				GetTestPoints.Add aPin
	Print #1,	"			End If
	Print #1,	"		Next
	Print #1,	"		Set vias = aNet.Vias
	Print #1,	"		If Sorted Then vias.Sort
	Print #1,	"		For Each aVia In vias
	Print #1,	"			If aVia.TestPoint <> ppcbTestPointNone Then
	Print #1,	"				GetTestPoints.Add aVia
	Print #1,	"			End If
	Print #1,	"		Next
	Print #1,	"	Next
	Print #1,	"
	Print #1,	"	Set pins = doc.Pins
	Print #1,	"	If Sorted Then pins.Sort
	Print #1,	"	For Each aPin In pins
	Print #1,	"		If aPin.Net Is Nothing And aPin.TestPoint <> ppcbTestPointNone Then
	Print #1,	"			GetTestPoints.Add aPin
	Print #1,	"		End If
	Print #1,	"	Next
	Print #1,	"	Set vias = doc.Vias
	Print #1,	"	If Sorted Then vias.Sort
	Print #1,	"	For Each aVia In vias
	Print #1,	"		If aVia.Net Is Nothing And aVia.TestPoint <> ppcbTestPointNone Then
	Print #1,	"			GetTestPoints.Add aVia
	Print #1,	"		End If
	Print #1,	"	Next
	Print #1,	"End Function
End Sub

Sub GenCodeForIsPinUsed
	Print #1
	Print #1,	"'Returns whether pin is used. We consider pin as used if it is connected to a net or it belongs to used (installed) gate"
	Print #1,  "Function IsPinUsed(aPin As Object) As Boolean"
	Print #1,  "	IsPinUsed = False"
	Print #1,  "	'if pin connected it is considered as used"
	Print #1,  "	If Not aPin.Net Is Nothing Then IsPinUsed = True: Exit Function"
	Print #1,  "	'if pin not connected and does not belong to any gate it is considred as unused"
	Print #1,  "	If aPin.Gate Is Nothing Then Exit Function"
	Print #1,  "	'if pin belongs to unused gate it is considred as unused"
	Print #1,  "	If aPin.Gate.Sheet Is Nothing Then Exit Function"
	Print #1,  "	'othewise it is used
	Print #1,  "	IsPinUsed = True"
	Print #1,  "End Function"
End Sub

Sub GenCodeForIsGateOnSheet
	Print #1
	Print #1,	"'Returns whether given gate is located on the given sheet.
	Print #1,  "Function IsGateOnSheet(aGate As Object, aSheet As Object) As Boolean"
	Print #1,  "	IsGateOnSheet = False"
	Print #1,  "	If aGate Is Nothing Then Exit Function"
	Print #1,  "	If aGate.Sheet Is Nothing Then Exit Function"
	Print #1,  "	If aGate.Sheet <> aSheet Then Exit Function"
	Print #1,  "	IsGateOnSheet = True"
	Print #1,  "End Function"
End Sub

Sub GenCodeForNetGateCollection
	Print #1
	Print #1,	"'Returns collection of gates the given net is connected to"
	Print #1,  "Function GetNetGates(aNet As Object, Optional Sorted As Boolean = False) As Object"
	Print #1, "		'Create empty object collection"
	Print #1, "		set GetNetGates = ActiveDocument.GetObjects(0)"
	Print #1, "		'Prepare sorted collection of net gates"
	Print #1, "		For Each aPin in aNet.Pins"
	Print #1, "			'Make sure pin belongs to some gate"
	Print #1, "			If Not aPin.Gate Is Nothing Then"
	Print #1, "				GetNetGates.Add aPin.Gate"
	Print #1, "			End If"
	Print #1, "		Next aPin"
	Print #1, "		'Sort Gate Collection"
	Print #1, "		If Sorted Then GetNetGates.Sort"
	Print #1, "		'One more pass to remove repeated items"
	Print #1, "		prev = """
	Print #1, "		For Each aGate in GetNetGates"
	Print #1, "			if aGate.Name = prev then GetNetGates.Remove(prev)
	Print #1, "			prev = aGate.Name
	Print #1, "		Next aGate"
	Print #1,  "End Function"
End Sub

Sub GenCodeForNetGatesBySheet
	Print #1
	Print #1,	"'Returns collection of gates the given net is connected to; gates must be located on the given sheet"
	Print #1,  "Function GetNetGatesBySheet(aNet As Object, aSheet As Object, Optional Sorted As Boolean = False) As Object"
	Print #1, "		'Create empty object collection"
	Print #1, "		set GetNetGatesBySheet = ActiveDocument.GetObjects(0)"
	Print #1, "		'Prepare sorted collection of net gates"
	Print #1, "		For Each aPin in aNet.Pins"
	Print #1, "			If IsGateOnSheet(aPin.Gate, aSheet) Then"
	Print #1, "				GetNetGatesBySheet.Add aPin.Gate"
	Print #1, "			End If"
	Print #1, "		Next aPin"
	Print #1, "		'Sort Gate Collection"
	Print #1, "		If Sorted Then GetNetGatesBySheet.Sort"
	Print #1, "		'One more pass to remove repeated items"
	Print #1, "		prev = """
	Print #1, "		For Each aGate in GetNetGatesBySheet"
	Print #1, "			if aGate.Name = prev then GetNetGatesBySheet.Remove(prev)
	Print #1, "			prev = aGate.Name
	Print #1, "		Next aGate"
	Print #1,  "End Function"
End Sub

Sub GenCodeForPinType
	Print #1
	Print #1, "Function PinType(aPin As Object) As String"
	Print #1, "		Select Case aPin.ElectricalType"
	Print #1, "			Case "; prefix; "ElectricalTypeSource"
	Print #1, "				PinType = ""Source"""
	Print #1, "			Case "; prefix; "ElectricalTypeBidirectional"
	Print #1, "				PinType = ""Bidirectional"""
	Print #1, "			Case "; prefix; "ElectricalTypeOpenCollector"
	Print #1, "				PinType = ""Collector"""
	Print #1, "			Case "; prefix; "ElectricalTypeOrTieableSource"
	Print #1, "				PinType = ""Tieable Source"""
	Print #1, "			Case "; prefix; "ElectricalTypeTristate"
	Print #1, "				PinType = ""Tristate"""
	Print #1, "			Case "; prefix; "ElectricalTypeLoad"
	Print #1, "				PinType = ""Load"""
	Print #1, "			Case "; prefix; "ElectricalTypeTerminator"
	Print #1, "				PinType = ""Terminator"""
	Print #1, "			Case "; prefix; "ElectricalTypePower"
	Print #1, "				PinType = ""Power"""
	Print #1, "			Case "; prefix; "ElectricalTypeGround"
	Print #1, "				PinType = ""Ground"""
	Print #1, "			Case Else "
	Print #1, "				PinType = ""Unknown"""
	Print #1, "		End Select	"
	Print #1, "End Function"
End Sub

Sub GenCodeForPinTypeShort
	Print #1
	Print #1, "Function PinTypeShort(aPin As Object) As String"
	Print #1, "		Select Case aPin.ElectricalType"
	Print #1, "			Case "; prefix; "ElectricalTypeSource"
	Print #1, "				PinTypeShort = ""S"""
	Print #1, "			Case "; prefix; "ElectricalTypeBidirectional"
	Print #1, "				PinTypeShort = ""B"""
	Print #1, "			Case "; prefix; "ElectricalTypeOpenCollector"
	Print #1, "				PinTypeShort = ""C"""
	Print #1, "			Case "; prefix; "ElectricalTypeOrTieableSource"
	Print #1, "				PinTypeShort = ""O"""
	Print #1, "			Case "; prefix; "ElectricalTypeTristate"
	Print #1, "				PinTypeShort = ""T"""
	Print #1, "			Case "; prefix; "ElectricalTypeLoad"
	Print #1, "				PinTypeShort = ""L"""
	Print #1, "			Case "; prefix; "ElectricalTypeTerminator"
	Print #1, "				PinTypeShort = ""Z"""
	Print #1, "			Case "; prefix; "ElectricalTypePower"
	Print #1, "				PinTypeShort = ""P"""
	Print #1, "			Case "; prefix; "ElectricalTypeGround"
	Print #1, "				PinTypeShort = ""G"""
	Print #1, "			Case Else "
	Print #1, "				PinTypeShort = ""U"""
	Print #1, "		End Select	"
	Print #1, "End Function"
End Sub

Sub GenCodeForPartCountOfType
	Print #1
	Print #1,	"'Counts parts of the given type in the given container (Document or "; ContName(0); ")"
	Print #1, "Function PartCountOfType(Container As  Object, PartType As String)
	Print #1, "	PartCountOfType = 0
	Print #1, "	For Each part In Container.Components
	Print #1, "		If part.PartType = PartType"; IIf(AppId = 2 And bInstalledOnly, " And part.Installed", ""); " Then PartCountOfType = PartCountOfType + 1"	
	Print #1, "	Next
	Print #1, "End Function
End Sub

Sub GenCodeForPartsByPartType
	Print #1
	Print #1, "'Returns collection of parts by the given part type
	Print #1, "Function GetPartsByPartType(Container As Object, PartType As String) As Object
	Print #1, "	Set GetPartsByPartType = Container.GetObjects(0)
	Print #1, "	For Each part in Container.Components
	Print #1, "		If part.PartType = PartType"; IIf(AppId = 2 And bInstalledOnly, " And part.Installed", ""); " Then
	Print #1, "			GetPartsByPartType.Add part
	Print #1, "		End If
	Print #1, "	Next
	Print #1, "End Function
End Sub

Sub GenCodeForUnusedGatesByPartType
	Print #1
	Print #1, "'Returns collection of unused gates by the given part type
	Print #1, "Function GetUnusedGatesByPartType(Container As Object, PartType As String) As Object
	Print #1, "	Set GetUnusedGatesByPartType = ActiveDocument.GetObjects(0)
	Print #1, "	For Each part in Container.Components
	Print #1, "		If part.PartType = PartType Then
	Print #1, "			For Each gt in part.UnusedGates
	Print #1, "				GetUnusedGatesByPartType.Add gt
	Print #1, "			Next
	Print #1, "		End If
	Print #1, "	Next
	Print #1, "End Function
End Sub

Sub GenCodeForUnusedPinsByPartType
	Print #1
	Print #1, "'Returns collection of unused part pins by the given part type
	Print #1, "Function GetUnusedPinsByPartType(Container As Object, PartType As String, Optional Sorted = False) As Object
	Print #1, "	Set GetUnusedPinsByPartType = ActiveDocument.GetObjects(0)
	Print #1, "	For Each part in Container.Components
	Print #1, "		If part.PartType = PartType Then
	Print #1, "			For Each aPin in part.Pins
	Print #1, "				If aPin.Gate Is Nothing and aPin.Net Is Nothing Then
	Print #1, "					GetUnusedPinsByPartType.Add aPin
	Print #1, "				End If
	Print #1, "			Next
	Print #1, "		End If
	Print #1, "	Next
	Print #1, "	If Sorted then GetUnusedPinsByPartType.Sort
	Print #1, "End Function
End Sub

Sub GenCodeForPinGate
	Print #1
	Print #1, "'Returns pin gate name if the pin belongs to one of the multiple part gates otherwise return empty string
	Print #1, "Function PinGate (obj As Object)"
	Print #1, "	Gate = """"
	Print #1, "	If obj.Gate Is Nothing Then Exit Function
	Print #1, "	If obj.Gate = obj.Component Then Exit Function
	Print #1, "	PinGate = obj.Gate
	Print #1, "End Function
End Sub

Sub GenCodeForPinSheet
	Print #1
	Print #1, "'Returns pin's sheet name; for signal pin returns sheet of the first used part gate
	Print #1, "Function PinSheet (obj As Object)"
	Print #1, "	PinSheet = """"
	Print #1, "	If obj.Gate Is Nothing Then
	Print #1, "		If Not obj.Net Is Nothing Then PinSheet = obj.Component.Gates(1).Sheet 
	Print #1, "	Else
	Print #1, "		If Not obj.Gate.Sheet Is Nothing Then PinSheet = obj.Gate.Sheet
	Print #1, "	End If
	Print #1, "End Function
End Sub

Sub GenCodeForPinPosX
	Print #1
	Print #1, "'Returns pin's X position if pin is visible otherwise empty string
	Print #1, "Function PinPosX (obj As Object)
	Print #1, "	PinPosX = """"
	Print #1, "	If obj.Gate Is Nothing Then Exit Function
	Print #1, "	If obj.Gate.Sheet Is Nothing Then Exit Function
	Print #1, "	PinPosX = obj.PositionX
	Print #1, "End Function
End Sub

Sub GenCodeForPinPosY
	Print #1
	Print #1, "'Returns pin's Y position if pin is visible otherwise empty string
	Print #1, "Function PinPosY (obj As Object)
	Print #1, "	PinPosY = """"
	Print #1, "	If obj.Gate Is Nothing Then Exit Function
	Print #1, "	If obj.Gate.Sheet Is Nothing Then Exit Function
	Print #1, "	PinPosY = obj.PositionY
	Print #1, "End Function
End Sub

Sub GenCodeForPointsOutput
	Print #1
	Print #1, "Function GetPoint (seg As Object, i As Integer, j As Integer)
	Print #1, "	GetPoint = """"
	Print #1, "	If (UBound(seg.Points, 1) >= i) And (UBound(seg.Points, 2) >= j) Then GetPoint = Format(seg.Points(i, j), ""0.000"")
	Print #1, "End Function
End Sub

Sub GenCodeConnRLength
	Print #1
	Print #1,	"'Calculates routed length of connection"
	Print #1,  "Function ConnRoutedLength(conn As Object)"
	Print #1,  "	ConnRoutedLength = 0"
	Print #1,  "	For Each aRSeg In conn.RouteSegments"
	Print #1,  "		ConnRoutedLength = ConnRoutedLength + aRSeg.Length"
	Print #1,  "	Next"
	Print #1,  "End Function"
End Sub

Sub GenDeclarationCode
	Dim i, j, s, w
	If Table Then
		Print #1
		If TXT Or RTF Then
			Print #1, "'Arrays of column name and widths. You can modify them to rename, shrink, or expand columns"
		Else
			Print #1, "'Array of column names. You can modify it to rename columns"
		End If
		Print #1, "Const Columns = Array(";
		If allCont And XLS Then
			Print #1, Chr(34); contHeader; Chr(34); ", ";
		End If
		For i = 1 To UBound(PropArray)
			Print #1, Chr(34); PropHeader(PropArray(i)); Chr(34);
			If i <> UBound(PropArray) Then Print #1, ", ";
		Next
		Print #1, ")"
		If TXT Or RTF Then
			Print #1, "Const Widths  = Array(";
			For i = 1 To UBound(PropArray)
				s = PropArray(i)
				w = Trim(Str(PropWidth(s)))
				Print #1, Space(Len(PropHeader(s)) - Len(w) + 2); w;
				If i <> UBound(PropArray) Then Print #1, ", ";
			Next
			Print #1, ")"
		End If
		If enableAlign Then
			Print #1, "'Array of column alignment: 0 - Align Left, 1 - Align Right, 2 - Align Center."
			Print #1, "Const Align   = Array(";
			If allCont And XLS Then
				w = "0"
				Print #1, Space(Len(contHeader) - Len(w) + 2); w; ", ";
			End If
			For i = 1 To UBound(PropArray)
				s = PropArray(i)
				w = Trim(Str(PropAlign(s)))
				Print #1, Space(Len(PropHeader(s)) - Len(w) + 2); w;
				If i <> UBound(PropArray) Then Print #1, ", ";
			Next
			Print #1, ")"
		End If
	End If
	
	If Table And (HTML And colorTable) Or ((HTML Or RTF) And colorColumns) Then
		Print #1
		Print #1, "'Colors in RGB hex format: ";
		s = ""
		If HTML And colorTable And Not bkgDef Then
			s = s + " hdr bkg,  odd bkg, even bkg"
		End If
		If HTML And colorTable And Not fontDef Then
			If s <> "" Then s = s + ", "
			s = s + " hdr txt,  odd txt, even txt"
		End If
		If (HTML Or RTF) And colorColumns Then
			For i = 1 To UBound(PropArray)
				If PropColor(PropArray(i)) <> "" Then
					If s <> "" Then s = s + ", "
					w = Trim(Str(i))
					s = s + "colmn" + w + Space(3 - Len(w))
				End If
			Next
		End If
		Print #1, s
		s = ""
		Print #1, "Const Colors      = Array( ";
		If HTML And colorTable And Not bkgDef Then
			s = s + Chr(34) + MyHex(bc0) + Chr(34) + ", "
			s = s + Chr(34) + MyHex(bc1) + Chr(34) + ", "
			s = s + Chr(34) + MyHex(bc2) + Chr(34)
		End If
		If HTML And colorTable And Not fontDef Then
			If s <> "" Then s = s + ", "
			s = s + Chr(34) + MyHex(fc0) + Chr(34) + ", "
			s = s + Chr(34) + MyHex(fc1) + Chr(34) + ", "
			s = s + Chr(34) + MyHex(fc2) + Chr(34)
		End If
		If (HTML Or RTF) And colorColumns Then
			For i = 1 To UBound(PropArray)
				If PropColor(PropArray(i)) <> "" Then
					If s <> "" Then s = s + ", "
					s = s + Chr(34) + MyHex(PropColor(PropArray(i))) + Chr(34)
				End If
			Next
		End If
		Print #1, s;
		Print #1, ")"
	End If
	If XLS And genHdr Then
		Print #1, "Dim fname As String"
	End If
	If primTotal And Not repByContainer And FilterCondition <> "" Then
		Print #1, "Dim Count As Long"
	End If
End Sub

Sub GenerateAdditionalFunctions
	'generate additional functions if needed
	If InStr(code, "GetSortedPins") Then 
		GenCodeForSortedPins
	End If
	If InStr(code, "GetPinsBySheet") Then 
		GenCodeForPinsBySheet
	End If
	If InStr(code, "GetConPins(") Then 
		GenCodeForConnectedPins
	End If
	If InStr(code, "GetConPinsBySheet") Then 
		GenCodeForConnectedPinsBySheet
	End If
	If InStr(code, "GetUnconPins(") Then 
		GenCodeForUnconnectedPins
	End If
	If InStr(code, "GetUnconPinsBySheet") Then 
		GenCodeForUnconnectedPinsBySheet
	End If
	If InStr(code, "GetUsedPins(") Then 
		GenCodeForUsedPins
	End If
	If InStr(code, "GetUsedPinsBySheet") Then 
		GenCodeForUsedPinsBySheet
	End If
	If InStr(code, "GetUnusedPins") Then 
		GenCodeForUnusedPins
	End If
	If InStr(code, "GetSignalPins") Then 
		GenCodeForSignalPins
	End If
	If InStr(code, "GetPowerPins") Then
		GenCodeForPowerPins
	End If
	If InStr(code, "GetTPName") Then
		GenCodeForTPName
	End If
	If InStr(code, "GetTestPoints") Then
		GenCodeForGetTPs
	End If
	If InStr(code, "GetPoint") Then
		GenCodeForPointsOutput
	End If
	If InStr(code, "ConnRoutedLength") Then
		GenCodeConnRLength
	End If
	If InStr(code, "GetSignalNetPins") Then 
		GenCodeForSignalNetPins
	End If
	If InStr(code, "OutRSegName")  Then
		GenCodeForRSegName
	End If
	If InStr(code, "OutViaName")  Then
		GenCodeForViaName
	End If
	If InStr(code, "GetOptName") Then
		GenCodeForGetOptName
	End If
	If InStr(code, "GetSegmentCons") Then
		GenCodeForGetSegmentCons
	End If
	If InStr(code, "GetPartGatesBySheet") Then 
		GenCodeForPartGatesBySheet
	End If
	If InStr(code, "IsGateOnSheet") Or InStr(code, "GetPinsBySheet") Or InStr(code, "GetConPinsBySheet") Or _
		InStr(code, "GetUsedPinsBySheet") Or InStr(code, "GetPartGatesBySheet") Or InStr(code, "GetNetGatesBySheet") Then 
		GenCodeForIsGateOnSheet
	End If
	If InStr(code, "IsPinUsed") Or InStr(code, "GetUsedPins") Or InStr(code, "GetUnusedPins") Or _
		InStr(code, "GetUsedPinsBySheet") Or InStr(code, "GetUnusedPinsByPartType") Then 
		GenCodeForIsPinUsed
	End If
	If InStr(code, "PinType(") Then
		GenCodeForPinType
	End If
	If InStr(code, "PinTypeShort") Then
		GenCodeForPinTypeShort
	End If
	If InStr(code, "GetNetGates(") Then 
		GenCodeForNetGateCollection
	End If
	If InStr(code, "GetNetGatesBySheet") Then 
		GenCodeForNetGatesBySheet
	End If
	If InStr(code, "PartCountOfType") Then
		GenCodeForPartCountOfType
	End If
	If InStr(code, "GetPartsByPartType") Then
		GenCodeForPartsByPartType
	End If
	If InStr(code, "GetUnusedGatesByPartType") Then
		GenCodeForUnusedGatesByPartType
	End If
	If InStr(code, "GetUnusedPinsByPartType") Then
		GenCodeForUnusedPinsByPartType
	End If
	If InStr(code, "AttrVal") Then
		Print #1
		Print #1, "Function AttrVal (obj As Object, nm As String)"
		Print #1, "	AttrVal = IIf(obj.Attributes(nm) Is Nothing, """", obj.Attributes(nm))"
		Print #1, "End Function"
	End If
	If InStr(code, "AttrMeas") Then
		Print #1
		Print #1, "Function AttrMeas (obj As Object, nm As String)
		Print #1, "	If obj.Attributes(nm) Is Nothing Then
		Print #1, "		Set AttrMeas = Measure	'just empty measure
		Print #1, "	Else
		Print #1, "		Set AttrMeas = obj.Attributes(nm).Measure
		Print #1, "	End If
		Print #1, "End Function
	End If
	If InStr(code, "ObjName") Then
		Print #1
		Print #1, "Function ObjName (obj As Object)"
		Print #1, "	ObjName = IIf(obj Is Nothing, """", obj)"
		Print #1, "End Function"
	End If
	If InStr(code, "PinGate") Then
		GenCodeForPinGate
	End If
	If InStr(code, "PinSheet") Then
		GenCodeForPinSheet
	End If
	If InStr(code, "PinPosX") Then
		GenCodeForPinPosX
	End If
	If InStr(code, "PinPosY") Then
		GenCodeForPinPosY
	End If
	If Table And Not XLS And (TXT Or RTF Or enableAlign) Then
		Print #1
		Print #1, "Dim CurCol As Integer	'Current column index staring from 0
	End If
	If HTML And Table And colorTable Then
		Print #1
		Print #1, "Dim RowType As Integer	'Type of the current table row: header row = 0, odd row = 1, even row = 2"
	End If
	If TXT Then
		GenCodeForText
	End If
	If XLS Then
		GenCodeForExcel		
	End If
	If RTF Then
		GenCodeForRTF
	End If
	If HTML Then
		GenCodeForHTML
	End If
	If outputFile = 2 Then
		GenCodeForFillClipboard
	End If
End Sub

Sub OpenOutputFile
	If outputFile = 1 Or outputFile = 2 Then
		If genHdr And incJob Then 
			Print #1, "	fname = ActiveDocument"
			Print #1, "	If fname = """" Then"
			Print #1, "		fname = ""Untitled"""
			Print #1, "	End If"
		End If
		If outputFile <> 2 Then Print #1, "	report = DefaultFilePath & ""\"" & """; repname; ext; """"
	ElseIf outputFile = 0 Then
		Print #1, "	'Make report file name from current schematic file name"
		Print #1, "	fname = ActiveDocument"
		Print #1, "	If fname = """" Then"
		Print #1, "		fname = ""Untitled"""
		Print #1, "		report = DefaultFilePath & ""\default"; ext; """"
		Print #1, "	Else"
		Print #1, "		nm = Left(fname, Len( fname) - 4)"
		Print #1, "		report = DefaultFilePath & ""\"" & nm & """; ext; """"
		Print #1, "	End If"
	End If
	If outputFile <> 2 Then
		Print #1, "	Open report For Output As #1"
	Else
		Print #1, "	tempFile = DefaultFilePath & ""\temp.txt"""
		Print #1, "	Open tempFile For Output As #1"
	End If
End Sub

Sub BeginFormatFile
	If RTF Then
		Print #1, "	'Write RTF file header"
		Print #1, "	BeginRTF"
	ElseIf HTML Then
		Print #1, "	'Write HTML file header"
		Print #1, "	BeginHTML"
	End If
End Sub

Sub EndFormatFile
	If RTF Then
		Print #1, "	EndRTF"
	ElseIf HTML Then
		Print #1, "	EndHTML"
	End If
End Sub

Sub BeginIndent
	If HTML  And secObj <> 0 Then 
		Print #1, Indent; "BeginIndent"
		nTabs = nTabs + 1
	End If
End Sub

Sub EndIndent
	If HTML  And secObj <> 0 Then 
		nTabs = nTabs - 1
		Print #1, Indent; "EndIndent"
	End If
End Sub

Sub OutputMultiLineCode(code As String)
	Dim i, s
	i = -1
	While i <> 0
		i = InStr(code, " _ ")
		If i > 0 Then 
			s = Left(code, i + 1)
			code = Mid(code, i + 3)
		Else
			s = code
		End If
		Print #1, Indent; s
	Wend
End Sub

Sub OutputHeader
	If genHdr Then
		If Not XLS Then Print #1, "	'Output report header"
		If HTML Then
			'special version of title for HTML in professional style
			Print #1, Indent; "OutTitle "; Chr(34); header; Chr(34); ", 6"
			If incJob  Then 
				Print #1,  Indent; "OutTitle "; Chr(34); "For "; Chr(34); DLMT; "fname"; 
				If incTime Then 
					Print #1, DLMT; Chr(34); " on "; Chr(34); DLMT; "Now";
				End If
				Print #1, ", 2"
			ElseIf incTime Then 
				Print #1,  Indent; "OutTitle "; Chr(34); "On "; Chr(34); DLMT; "Now";
				Print #1, ", 2"
			End If
		Else
			Print #1, Indent; OutBoldLine; Chr(34); header;
			If incJob  Then 
				Print #1,  " for "; Chr(34); DLMT; "fname"; 
				If incTime Then 
					Print #1, DLMT; Chr(34); " on "; Chr(34); DLMT; "Now";
				End If
			ElseIf incTime Then 
				Print #1, " on "; Chr(34); DLMT; "Now";
			Else
				Print #1, Chr(34); 
			End If
			Print #1
		End If
		Print #1, Indent; NewLine
	End If
End Sub

Sub BeginProgress
	If ShowStatus Then
		Print #1, Indent; "StatusBarText = "; Chr(34); "Generating report..."; Chr(34)
	End If
	If ShowProgress Then
		Dim coll
		coll = primColl
		Print #1, Indent; "Processed = 0"
		Print #1, Indent; "ProgressBar = 0"
		If allCont And AppId = 2 Then 'AO report for PADS Layout
			FormatTemplate coll, "xxx", "ActiveDocument"
			Print #1, Indent; "Total = "; coll; ".Count * ActiveDocument.AssemblyOptions.Count"
		ElseIf allCont And primObj <> plogGates Then 'Sheet report for PADS Logic objects except gates
			FormatTemplate coll, "xxx", contObj
			Print #1, Indent; "'Calculate total number of "; primName; " entries in all "; ContName(1); "s for correct progress info (this value can be different from the number of "; primName; "s in entire design)
			Print #1, Indent; "Total = 0"
			Print #1, Indent; "For Each "; contObj; " In "; contColl
			Print #1, Indent; "	Total = Total + "; coll; ".Count"
			Print #1, Indent; "Next"
		ElseIf allCont Then 'Sheet report for PADS Logic gates
			FormatTemplate coll, "xxx", "ActiveDocument"
			Print #1, Indent; "Total = "; coll; ".Count"
		ElseIf givenCont <> "" Then 'report for the given ass.opt./sheet
			FormatTemplate coll, "xxx", contObj
			Print #1, Indent; "Total = "; coll; ".Count"
		Else 'Whole Document
			FormatTemplate coll, "xxx", "ActiveDocument"
			Print #1, Indent; "Total = "; coll; ".Count"
		End If
	End If
End Sub

Sub EndProgress
	If ShowProgress Then
		Print #1, Indent; "ProgressBar = -1"
	End If
	If ShowStatus Then
		Print #1, Indent; "StatusBarText = "; Chr(34); Chr(34) 
	End If
End Sub

Sub OutputTotal
	If primTotal Then
		Dim coll
		coll = primColl
		FormatTemplate coll, "xxx", "ActiveDocument"
		Print #1, Indent; NewLine
		If Not repByContainer And FilterCondition <> "" Then
			Print #1, Indent; OutBoldLine; Chr(34);  "Number of "; primName; "s in the design matching the request: "; Chr(34); DLMT; "Count"
		End If
		Print #1, Indent; OutBoldLine; Chr(34);  "Design "; primName; " count: "; Chr(34); DLMT; coll; ".Count"
	End If
End Sub

Sub CloseAndShowOutputFile
	Print #1, "	Close #1"
	If TXT Or RTF Then
		If outputFile <> 2 Then
			Print #1, "	'Do not forget quotes for file name!"
			Print #1, "	Shell """; targetApp(target); " "" & Chr(34) & report & Chr(34), 1"
		Else
			Print #1, "	FillClipboard"
			Print #1, "	App = Shell ("; Chr(34); targetApp(target); Chr(34); ", 1)"
			Print #1, "	AppActivate App"
			Print #1, "	SendKeys "; Chr(34); "+{Ins}"; Chr(34); " 'Shift-Ins  Paste"
			Print #1, "	SendKeys "; Chr(34); "^{Home}"; Chr(34); " 'Ctrl+Home  Go to top of the document"
		End If
	ElseIf XLS Then
		If outputFile = 2 Then
			Print #1, "	ExportToExcel"
		Else
			Print #1, "	ExportToExcel report"
			Print #1, "	Kill report"
		End If		
	ElseIf HTML Then
		Print #1, "	ShellExecute (0, ""Open"", report, """", """", 5)"
	End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Fill string array of all object attributes at the moment of Wizard execution
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetAttributeList()
	Dim i, objType, attrType
	objType = primData(primObj)(4)
	ReDim AttrArray(ActiveDocument.AttributeTypes.Count)
	i = 1
	For Each attrType In ActiveDocument.AttributeTypes
		If attrType.IsAllowed(objType) Then
			AttrArray(i) = attrType
			i = i + 1
		End If
	Next
End Sub

'''''''''''''''''''''''''''''''''''''''''
' Load/Restore Registry Setting of Wizard 
'''''''''''''''''''''''''''''''''''''''''
Const Section1		= "Step1"
Const Key1_1 		= "Script dir"
Const Section2		= "Step2"
Const Key2_1 		= "Target"
Const Section3 		= "Step3"
Const Key3_1 		= "Type of report"
Const Key3_2 		= "Given sheet"
Const Section4 		= "Step4"
Const Key4_1 		= "Primary object"
Const Key4_2 		= "Primary total"
Const Key4_3 		= "Primary subtotal"
Const Key4_4 		= "Installed Only"
Const Section5 		= "Step5"
Const Key5_1 		= "Output format"
Const Section6 	    = "Step6"
Const Key6_1 		= "List count"
Const Section6B		= "Step6B"
Const Key6B_1 		= "Object types"
Const Key6B_2 		= "Calculate totals"
Const Key6B_3 		= "Filter by sheet"
Const Key6B_4 		= "Sort secondary objects"
Const Section7 		= "Step7"
Const Key7_1 		= "Output Report Header"
Const Key7_2 		= "Header"
Const Key7_3 		= "Include job name"
Const Key7_4 		= "Include data and time"
Const Key7_5 		= "Show status bar"
Const Key7_6 		= "Show progress bar"
Const Key7_8 		= "Count per line"
Const Key7_9 		= "Color HTML tables"
Const Key7_10 		= ContName(0) & " Contents"
Const Key7_11 		= "Last Used Color Scheme"
Const Key7_12 		= "Draw Horizontal Border"
Const Key7_13		= "Draw Vertical Border"
Const Key7_14 		= "Header Rule"
Const Key7_15 		= "Enable Alignment"
Const Key7_16 		= "Filter By Count"
Const Section8 		= "Step8"
Const Key8_1 		= "Creation method"
Const Key8_2 		= "Custom report file name"
Const Key8_3 		= "Text File extension"
Const Key8_4 		= "Script name"
Const Key8_5 		= "Report Name Rule"
Const Key8_6 		= "Script Name Rule"
Const ColConfSection= "Color Config"
Const PropSection	= "Column Properties"
Const ListSection	= "List Properties"

Sub RestoreWizardSettings
	Dim i, count, s
	
	scriptDir		= GetSetting(WizName, Section1, Key1_1, MacroDir)
	'check that specified directory exists, otherwise use current wizard directory
	On Error Resume Next
	s = Dir(scriptDir, vbDirectory)
	On Error GoTo 0
	If s = "" Then scriptDir = MacroDir
	
	target			= CInt (GetSetting(WizName, Section2, Key2_1, Browser))
	repByContainer	= CBool(GetSetting(WizName, Section3, Key3_1, False))
	givenCont		= Trim (GetSetting(WizName, Section3, Key3_2, ""))
	allCont			= givenCont = "" And repByContainer
	primObj			= CInt (GetSetting(WizName, Section4, Key4_1, 0))
	primTotal		= CBool(GetSetting(WizName, Section4, Key4_2, False))
	primContTotal	= CBool(GetSetting(WizName, Section4, Key4_3, False))
	bInstalledOnly 	= CBool(GetSetting(WizName, Section4, Key4_4, False))
	tableSec		= CInt (GetSetting(WizName, Section5, Key5_1, 2))
	Table 			= (tableSec = 0)
	primProps 		= Choose(tableSec + 1, primObjProps(primObj), primObjLists(primObj), primList(primObj))
	primName 		= primData(primObj)(0)
	LoadPropArray

	secObj 				= CBool (GetSetting(WizName, Section6B, Key6B_1, False))
	secTotal 			= CBool(GetSetting(WizName, Section6B, Key6B_2, False))
	filterByContainer	= CBool(GetSetting(WizName, Section6B, Key6B_3, False))
	sortObj				= CBool(GetSetting(WizName, Section6B, Key6B_4, True))
	genHdr				= CBool(GetSetting(WizName, Section7,  Key7_1, True))
	header				= Trim (GetSetting(WizName, Section7,  Key7_2, ""))
	incJob				= CBool(GetSetting(WizName, Section7,  Key7_3, True))
	incTime				= CBool(GetSetting(WizName, Section7,  Key7_4, True))
	ShowStatus			= CBool(GetSetting(WizName, Section7,  Key7_5, True))
	ShowProgress		= CBool(GetSetting(WizName, Section7,  Key7_6, False))
	colorTable			= CBool(GetSetting(WizName, Section7,  Key7_9, False))
	contContents		= CBool(GetSetting(WizName, Section7,  Key7_10, False))
	lastScheme			= Trim (GetSetting(WizName, Section7,  Key7_11, "Spruce"))
	drawHBorder			= CBool(GetSetting(WizName, Section7,  Key7_12, True))
	drawVBorder			= CBool(GetSetting(WizName, Section7,  Key7_13, False))
	headerRule			= CInt (GetSetting(WizName, Section7,  Key7_14, 1))
	enableAlign			= CBool(GetSetting(WizName, Section7,  Key7_15, False))
	filterByCount		= CBool(GetSetting(WizName, Section7,  Key7_16, False))
	outputFile			= CInt (GetSetting(WizName, Section8,  Key8_1, 0))
	repname				= Trim (GetSetting(WizName, Section8,  Key8_2, ""))
	txtFileExt			= Trim (GetSetting(WizName, Section8,  Key8_3, "rep"))
	scriptName			= Trim (GetSetting(WizName, Section8,  Key8_4, ""))
	repnameRule			= CInt (GetSetting(WizName, Section8,  Key8_5, 1))
	scriptNameRule		= CInt (GetSetting(WizName, Section8,  Key8_6, 1))

	ReDim Preserve AttrArray (0)
	curAttribute = 0
	prefix = IIf(AppId = 1, "plog", "ppcb")

	PredefineColorConfigs
	LoadColorConfig "Current"
End Sub

Sub LoadPropArray
	Dim i, count, entry
	entry = primData(primObj)(0) & " " & IIf(Table, "Columns", "Lists") & " "
	count = CInt(GetSetting(WizName, Section6, entry & "Count", 0))
	ReDim PropArray (count)
	For i = 1 To count
		PropArray(i) = GetSetting(WizName, Section6, entry & "property" & i, "")
	Next i
	If Table And count = 0 Then
		ReDim PropArray (1)
		PropArray(1) = primObjProps(primObj)(0)(0)
	End If
End Sub

Sub SavePropArray
	Dim i, count, entry	
	entry = primData(primObj)(0) & " " & IIf(Table, "Columns", "Lists") & " "
	
	'	delete previous list
	count = CInt(GetSetting(WizName, Section6, entry & "Count", 0))
	For i = 1 To count
		DeleteSetting WizName, Section6, entry & "property" & i
	Next i

	'	fill list with new values
	SaveSetting WizName, Section6, entry & "Count", UBound(PropArray)
	For i = 1 To UBound(PropArray)
		SaveSetting WizName, Section6, entry & "property" & i, PropArray(i)
	Next i
End Sub

Sub SaveWizardSettings
	SaveSetting WizName, Section1, Key1_1, scriptDir
	SaveSetting WizName, Section2, Key2_1, target
	SaveSetting WizName, Section3, Key3_1, repByContainer
	SaveSetting WizName, Section3, Key3_2, givenCont
	SaveSetting WizName, Section4, Key4_1, primObj
	SaveSetting WizName, Section4, Key4_2, primTotal
	SaveSetting WizName, Section4, Key4_3, primContTotal
	SaveSetting WizName, Section4, Key4_4, bInstalledOnly
	SaveSetting WizName, Section5, Key5_1, tableSec

	SavePropArray
	
	SaveSetting WizName, Section6B, Key6B_1, secObj
	SaveSetting WizName, Section6B, Key6B_2, secTotal
	SaveSetting WizName, Section6B, Key6B_3, filterByContainer
	SaveSetting WizName, Section6B, Key6B_4, sortObj
	SaveSetting WizName, Section7, Key7_1, genHdr
	SaveSetting WizName, Section7, Key7_2, header
	SaveSetting WizName, Section7, Key7_3, incJob
	SaveSetting WizName, Section7, Key7_4, incTime
	SaveSetting WizName, Section7, Key7_5, ShowStatus
	SaveSetting WizName, Section7, Key7_6, ShowProgress
	SaveSetting WizName, Section7, Key7_9, colorTable
	SaveSetting WizName, Section7, Key7_10, contContents
	SaveSetting WizName, Section7, Key7_11, lastScheme
	SaveSetting WizName, Section7, Key7_12, drawHBorder
	SaveSetting WizName, Section7, Key7_13, drawVBorder
	SaveSetting WizName, Section7, Key7_14, headerRule
	SaveSetting WizName, Section7, Key7_15, enableAlign
	SaveSetting WizName, Section7, Key7_16, filterByCount
	SaveSetting WizName, Section8, Key8_1, outputFile
	SaveSetting WizName, Section8, Key8_2, repname
	SaveSetting WizName, Section8, Key8_3, txtFileExt
	SaveSetting WizName, Section8, Key8_4, scriptName
	SaveSetting WizName, Section8, Key8_5, repnameRule
	SaveSetting WizName, Section8, Key8_6, scriptNameRule

	SaveColorConfig "Current"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''
'Color configuration functions 
'''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SetupColors
	Begin Dialog UserDialog 450,217,"Table Colors",.SetupColorsFunc ' %GRID:10,7,1,1
		CheckBox 10,28,210,14,"Browser Default Background",.bdef
		CheckBox 230,28,200,14,"Browser Default &Text Color",.fdef
		PushButton 10,49,210,21,"Table &Header Background",.bc0
		PushButton 230,49,210,21,"Table H&eader Text",.fc0
		PushButton 10,77,210,21,"&Odd Row Background",.bc1
		PushButton 230,77,210,21,"Odd &Row Text",.fc1
		PushButton 10,105,210,21,"&Even Row Background",.bc2
		PushButton 230,105,210,21,"E&ven Row Text",.fc2
		Text 10,7,420,14,"Use the buttons below to assign table row colors:",.Text1
		Text 10,140,180,14,"&Color Scheme:",.Text2
		DropListBox 10,154,210,63,ColConfigs(),.ConfigList
		PushButton 230,154,100,21,"&Save...",.SaveBtn
		PushButton 340,154,100,21,"&Delete",.DeleteBtn
		PushButton 230,189,140,21,"&Close",.CloseBtn
		PushButton 80,189,140,21,"&Preview",.PreviewBtn
	End Dialog
	Dim dlg As UserDialog

	AllocateCustomPalette
	dlg.bdef = bkgDef
	dlg.fdef = fontDef
	Dialog dlg
	bkgDef = dlg.bdef = 1
	fontDef = dlg.fdef = 1
	If preview <> "" Then
		Kill preview
		preview = ""
	End If
End Sub

Sub PreviewColors
	preview = "Table Color Preview.html"
	Open "Table Color Preview.html" For Output As #2
	Print #2,	"<HTML>"
	Print #2,	"<HEAD>"
	Print #2,	"<TITLE>Table Colors Preview</TITLE>"
	Print #2,	"</HEAD>"
	Print #2,	"<BODY>"
	Print #2,	"<TABLE CELLSPACING=1 BORDER=1 CELLPADDING=2 WIDTH=100%>"
	Dim bc, fc, txt, j, row
	For row = 0 To 10
		If row = 0 Then
			bc = MyHex(bc0)
			fc = MyHex(fc0)
			txt = "<B>Table Header Text</B>"
		ElseIf row Mod 2 = 1 Then
			bc = MyHex(bc1)
			fc = MyHex(fc1)
			txt = "Odd Row Text"
		Else
			bc = MyHex(bc2)
			fc = MyHex(fc2)
			txt = "Even Row Text"
		End If
		If DlgValue("bdef")  = 1 Then 
			Print #2, "<TR>";
		Else
			Print #2, "<TR BGCOLOR = #"; bc; ">";
		End If
		For j = 1 To 3
			Print #2, "<TD>";
			If DlgValue("fdef")  = 0 Then Print #2, "<FONT COLOR = #"; fc; ">";
			Print #2, txt;
			If DlgValue("fdef")  = 0 Then Print #2,	"</FONT>"
			Print #2, "</TD>";
		Next
		Print #2, "</TR>";
	Next
	Print #2,	"</TABLE>"
	Print #2,	"</BODY>"
	Print #2,	"</HTML>"
	Close #2
	ShellExecute (0, "Open", preview, "", "", 5)
End Sub

Sub CheckScheme
	If IsCurrentScheme(lastScheme) Then 
		DlgText "ConfigList", lastScheme
		DlgEnable "DeleteBtn", True
	Else
		DlgValue "ConfigList", -1
		DlgEnable "DeleteBtn", False
	End If
End Sub

Rem See DialogFunc help topic for more information.
Private Function SetupColorsFunc(DlgItem$, Action%, SuppValue&) As Boolean
	Dim configName
	Select Case Action%
	Case 1 ' Dialog box initialization
		If SuppValue <> 0 Then colorDlgHnd = SuppValue
		DlgEnable "bc0", DlgValue("bdef") = 0
		DlgEnable "bc1", DlgValue("bdef") = 0
		DlgEnable "bc2", DlgValue("bdef") = 0
		DlgEnable "fc0", DlgValue("fdef") = 0
		DlgEnable "fc1", DlgValue("fdef") = 0
		DlgEnable "fc2", DlgValue("fdef") = 0
		GetColorConfigs
		DlgListBoxArray "ConfigList", ColConfigs()
		CheckScheme
	Case 2 ' Value changing or button pressed
		SetupColorsFunc = True
		If DlgItem$ = "bc0" Then
			bc0 = AssignColor(colorDlgHnd, bc0)
			CheckScheme
		ElseIf DlgItem$ = "fc0" Then
			fc0 = AssignColor(colorDlgHnd, fc0)
			CheckScheme
		ElseIf DlgItem$ = "bc1" Then
			bc1 = AssignColor(colorDlgHnd, bc1)
			CheckScheme
		ElseIf DlgItem$ = "fc1" Then
			fc1 = AssignColor(colorDlgHnd, fc1)
			CheckScheme
		ElseIf DlgItem$ = "bc2" Then
			bc2 = AssignColor(colorDlgHnd, bc2)
			CheckScheme
		ElseIf DlgItem$ = "fc2" Then
			fc2 = AssignColor(colorDlgHnd, fc2)
			CheckScheme
		ElseIf DlgItem$ = "bdef" Then
			bkgDef = DlgValue("bdef") = 1
			SetupColorsFunc("", 1, 0)
		ElseIf DlgItem$ = "fdef" Then
			If DlgValue("fdef") = 0 Then
				MsgBox "Changing of default text color may result in dramatic increase of your HTML files." & vbCr & _
						"To keep HTML report as compact as possible we recommend using of Browser default colors." & vbCr & _
						"Notice that row text color has a lower priority than column text color you could assign on previous step." _
						, vbExclamation, AppName
			End If
			fontDef = DlgValue("fdef") = 1
			SetupColorsFunc("", 1, 0)
		ElseIf DlgItem$ = "CloseBtn" Then
			SetupColorsFunc = False
		ElseIf DlgItem$ = "PreviewBtn" Then
			PreviewColors
		ElseIf DlgItem$ = "ConfigList" Then
			configName = DlgText("ConfigList")
			LoadColorConfig configName
			DlgValue "bdef", bkgDef
			DlgValue "fdef", fontDef
			lastScheme = configName
			SetupColorsFunc("", 1, 0)
		ElseIf DlgItem$ = "SaveBtn" Then
			configName = InputBox("Save this color scheme as:", AppName, lastScheme)
			If configName <> "" Then
				SaveColorConfig configName
				lastScheme = configName
				SetupColorsFunc("", 1, 0)
			End If
		ElseIf DlgItem$ = "DeleteBtn" Then
			DeleteColorConfig DlgText("ConfigList")
			DlgValue "bdef", bkgDef
			DlgValue "fdef", fontDef
			SetupColorsFunc("", 1, 0)
		End If
	End Select
End Function

Const SysSchemes = Array("Current", "D5D2E3 BBDBBD A6D0A8 000000 000000 000000 0 1", _
						 "Desert", 	"F3B9B6 F1EFC7 EAE7AC 000000 000000 000000 0 1", _
						 "Eggplant","A4A7C1 80C1AC A5D3C5 000000 000000 000000 0 1", _
						 "Lilac", 	"AEA8D9 C8C4E6 ADAEE2 000000 000000 000000 0 1", _
						 "Navy", 	"B4BECD 9CCBCD C2E4E7 000000 000000 000000 0 1", _
						 "Rose", 	"D3A3B5 EBC7D6 E6B0C4 000000 000000 000000 0 1", _
						 "Spruce",	"D5D2E3 BBDBBD A6D0A8 000000 000000 000000 0 1")

Sub PredefineColorConfigs 
	Dim i
	For i = 0 To UBound(SysSchemes) \ 2
		PredefineColorConfig SysSchemes(i*2 + 1), SysSchemes(i*2)
	Next
End Sub

Sub PredefineColorConfig (value As String, configName As String)
	If GetSetting(WizName, ColConfSection,  configName, "") = "" Then
		SaveSetting WizName, ColConfSection, configName, value
	End If
End Sub

Sub LoadColorConfig (configName As String)
	Dim s
	If configName = "" Then Exit Sub
	s = GetSetting(WizName, ColConfSection,  configName, "")
	bc0 = Val("&H" & RGBHex(Mid(s, 1, 6)))
	bc1 = Val("&H" & RGBHex(Mid(s, 8, 6)))
	bc2 = Val("&H" & RGBHex(Mid(s, 15, 6)))
	fc0 = Val("&H" & RGBHex(Mid(s, 22, 6)))
	fc1 = Val("&H" & RGBHex(Mid(s, 29, 6)))
	fc2 = Val("&H" & RGBHex(Mid(s, 36, 6)))
	bkgDef = CBool(Mid(s, 43, 1))
	fontDef	= CBool(Mid(s, 45, 1))
End Sub

Function GetConfigString As String
	GetConfigString = 	MyHex(bc0) & " " & MyHex(bc1) & " " & MyHex(bc2) & " " & _
						MyHex(fc0) & " " & MyHex(fc1) & " " & MyHex(fc2) & " " & _
						IIf(bkgDef, "1", "0") & " " & IIf(fontDef, "1", "0")
End Function

Sub SaveColorConfig (configName As String)
	If configName = "" Then Exit Sub
	SaveSetting WizName, ColConfSection,  configName, GetConfigString
End Sub

Function IsCurrentScheme (configName As String) As Boolean
	Dim s
	IsCurrentScheme = False
	If configName = "" Then Exit Function
	s = GetSetting(WizName, ColConfSection,  configName, "")
	IsCurrentScheme = (s = GetConfigString)
End Function

Sub DeleteColorConfig (configName As String)
	Dim i, res
	If configName = "" Then Exit Sub
	For i = 0 To UBound(SysSchemes) \ 2
		If UCase(configName) = UCase(SysSchemes(i*2)) Then
			res = MsgBox ("You cannot delete predefined color scheme " & Chr(34) & SysSchemes(i*2) & Chr(34) & vbCr & _
						"However, this action will reset it to original state." & vbCr & _
						"Continue?", vbYesNo Or vbExclamation, AppName)
			If res = vbYes Then
				DeleteSetting WizName, ColConfSection,  SysSchemes(i*2)
				PredefineColorConfig SysSchemes(i*2 + 1), SysSchemes(i*2)
				LoadColorConfig SysSchemes(i*2)
			End If
			Exit Sub
		End If
	Next
	DeleteSetting WizName, ColConfSection,  configName
	lastScheme = ""
End Sub

Sub GetColorConfigs
	Dim Settings, i
	ReDim ColConfigs(0)
	Settings = GetAllSettings(WizName, ColConfSection)
	If IsEmpty(Settings) Then Exit Sub
	ReDim ColConfigs(UBound(Settings))
	For i = 0 To UBound(Settings)
		If Settings(i,0) <> "Current" Then
			ColConfigs(i) = Settings(i,0)
		End If
	Next
End Sub

'Declare Win32 functions and structures not supported in VB laguage directly
Type CHOOSECOLOR
    lStructSize As Long 
    hwndOwner As Long 
    hInstance As Long
    rgbResult As Long 
    lpCustColors As Long 
    Flags As Long 
    lCustData As Long 
    lpfnHook As Long 
    lpTemplateName As Long
End Type
Dim cs As CHOOSECOLOR

Type Rect
	x1 As Long
	y1 As Long
	x2 As Long
	y2 As Long
End Type
Dim wr As Rect

Type BROWSEINFO
    hwndOwner As Long 
    pidlRoot As Long
    pszDisplayName As Long 
    lpszTitle As Long 
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Dim heap As Long

Private Declare Function ChooseColor Lib "comdlg32.dll" (ByRef cs As CHOOSECOLOR) As Long
Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long
Private Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal paint As Long) As Long
Private Declare Function EnableWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal enable As Boolean) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef r As Rect) As Boolean
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long , ByVal op As String, ByVal fname As String, ByVal par As String, ByVal lpDirectory As String, ByVal cmdShow As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (ByRef bi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal dirName As String) As Long

Sub AllocateCustomPalette
	If heap = 0 Then
		heap = GetProcessHeap
		cs.lStructSize =  36
		cs.Flags = 3
		cs.lpCustColors = HeapAlloc(heap, 8, 64)
	End If
End Sub

Sub FreeCustomPalette
	If heap <> 0 Then
		HeapFree(heap, 0, cs.lpCustColors)
	End If
End Sub

Function AssignColor(hwnd, color)
	AllocateCustomPalette
	cs.rgbResult = color
	cs.hwndOwner = hwnd
	ChooseColor(cs)
	AssignColor = cs.rgbResult
End Function

''''''''''''''''''''''''''''''
' Misc Utilities
''''''''''''''''''''''''''''''
Function BrowseDirectory As String
	Dim bi As BROWSEINFO, s As String*256, pos
	bi.ulFlags = 9
	bi.hwndOwner = dlgHandle
	Dim pidl As Long
	pidl = SHBrowseForFolder(bi)
	If pidl <> 0 Then
		If SHGetPathFromIDList(pidl, s) <> 0 Then
			pos = InStr(s, Chr(0))
			BrowseDirectory = Left(s, pos-1)
		End If
	End If
End Function

Sub GenCodeForFillClipboard
	Print #1
	Print #1,	"Sub FillClipboard"
    If ShowStatus Then
		Print #1,	"	StatusBarText = ""Export Data To Clipboard..."""
	End If
	Print #1,	"	' Load whole file to string variable    "
	Print #1,	"	tempFile = DefaultFilePath & ""\temp.txt"""
	Print #1,	"	Open tempFile  For Input As #1"
	Print #1,	"	L = LOF(1)"
	Print #1,	"	AllData$ = Input$(L,1)"
	Print #1,	"	Close #1"
	Print #1,	"	'Copy whole data to clipboard"
	Print #1,	"	Clipboard AllData$ "
	Print #1,	"	Kill tempFile"
    If ShowStatus Then
		Print #1,	"	StatusBarText = """""
	End If
	Print #1,	"End Sub"
End Sub

Function RGBHex(a) As String
	RGBHex = Right(a, 2) & Mid(a, 3, 2) & Left(a, 2)
End Function

Function MyHex(a) As String
	Dim l, s
	l = Len(Hex(a))
	If l < 6 Then
		s = String(6 - l, "0") & Hex(a)
	Else
		s = Hex(a)
	End If
	'change order of bytes
	MyHex = RGBHex(s)
End Function 

'****************************************************************
'  Here is the code generator for different application formats
'****************************************************************

'''''''''''''''''''''''''''''''''''''''''''''''
' Excel Format Code Generator
'''''''''''''''''''''''''''''''''''''''''''''''
Sub GenCodeForExcel
	Print #1
	If outputFile = 2 Then
		Print #1,	"Sub ExportToExcel"
	    Print #1,	"	FillClipboard"
	Else
		Print #1,	"Sub ExportToExcel (TextFile As String)"
	End If
	If outputFile = 2 Then
		Print #1,	"	Dim xl As Object"
		Print #1,	"	On Error Resume Next"
		Print #1,	"	Set xl =  GetObject(,""Excel.Application"")"
		Print #1,	"	On Error GoTo ExcelError	' Enable error trapping."
		Print #1,	"	If xl Is Nothing Then"
		Print #1,	"		Set xl =  CreateObject(""Excel.Application"")"
		Print #1,	"	End If"
		Print #1,	"	xl.Visible = True"
		Print #1,	"	xl.Workbooks.Add"
		Print #1,	"	xl.ActiveSheet.Paste"
	Else
		Print #1,	"	On Error GoTo ExcelError	' Enable error trapping."
		Print #1,	"	Set xl =  CreateObject(""Excel.Application"")"
		Print #1,	"	xl.Visible = True"
		Print #1,	"	xl.Workbooks.OpenText FileName:=TextFile, Tab:=True"
		Print #1,	"	'save worksheet as native Excel document"
		Print #1,	"	NativeFile = Left(TextFile, Len(TextFile) - 4) + "".xls"""
		Print #1,	"	xl.ActiveWorkbook.SaveAs Filename:= NativeFile, FileFormat:= -4143"
	End If
	If Table Then
		Dim range, nameCol, i
		range = Chr(34) & "A1:" & GetLastColName & "1" & Chr(34)
		Print #1,	"	xl.Range("; range; ").Font.Bold = True"
		Print #1,	"	xl.Range("; range; ").NumberFormat = " & Chr(34) & "@"  & Chr(34)
		Print #1,	"	xl.Range("; range; ").AutoFilter"
		If enableAlign Then
			Print #1,	"	For i = 0 To UBound(Align)"
			Print #1,	"		xl.Columns(i + 1).HorizontalAlignment = Choose(Align(i)+1, -4131, -4152, -4108)"
			Print #1,	"	Next"
		End If
		'calculate subtotals
		If allCont And primContTotal Then 
			nameCol = ""
			For i = 1 To UBound(PropArray)
				If PropArray(i) = primProps(0)(0) Then
					nameCol = Trim(Str(i + 1))
					Exit For
				End If
			Next
			If nameCol <> "" Then
				Print #1,	"	xl.Range(""A1"").Subtotal GroupBy:=1, Function:=2, TotalList:=Array("; nameCol; ")"
			End If
		End If
		'autofit all
		Print #1,	"	xl.ActiveSheet.UsedRange.Columns.AutoFit"
		'output condition after table is inserted
		If Table And FilterCondition <> "" Then
			Print #1,	"	'Output Report Condition
			OutBoldLine = "xl.Rows(1).Cells(1) = "
			OutLine 	= "xl.Rows(2).Cells(1) = "
			NewLine 	= "xl.Rows(3).Insert"
			Print #1,	"	xl.Rows(1).Insert"
			Print #1,	"	xl.Rows(1).Insert"
			OutputCondition
			Print #1,	"	xl.Rows(1).Font.bold = True"
		End If
		'Insert Header after table is inserted
		If genHdr Then 
			Print #1,	"	'Output Report Header
			OutBoldLine = "xl.Rows(1).Cells(1) = Space(1) & "
			NewLine 	= "xl.Rows(2).Insert"
			Print #1,	"	xl.Rows(1).Insert"
			DLMT 		= " & "
			OutputHeader
			If Table Then Print #1,	"	xl.Rows(1).Font.bold = True"
		End If
		'Insert Totals after table is inserted
		If primTotal Then
			Print #1,	"	'Output Design Totals
			Print #1,	"	lastRow = xl.ActiveSheet.UsedRange.Rows.Count + 1
			DLMT 		= " & "
			If Not repByContainer And FilterCondition <> "" Then
				Print #1,	"	xl.Rows(lastRow + 1).Font.bold = True
				OutBoldLine = "lastRow = lastRow + 1: xl.Rows(lastRow).Cells(1) = Space(1) & "
				NewLine = "xl.Rows(lastRow + 2).Font.bold = True"
			Else
				OutBoldLine = "xl.Rows(lastRow + 1).Cells(1) = Space(1) & "
				NewLine = "xl.Rows(lastRow + 1).Font.bold = True"
			End If
			OutputTotal
		End If
	End If 'Table
	If outputFile = 2 Then 'goto top after paste
		Print #1,	"	xl.Range(""A1"").Select"
	End If
	'Formatting bold lines
	If Not Table Then
		Print #1,	"	'Format bold lines"
		Print #1,	"	For Each row In xl.ActiveSheet.UsedRange.Rows"
		Print #1,	"		If Left(row.Cells(1), 1) = Space(1) Then"
		Print #1,	"			row.Font.bold = True"
		Print #1,	"		End If"
		Print #1,	"	Next"
	End If
	Print #1,	"	On Error GoTo 0 ' Disable error trapping. "
	Print #1,	"	Exit Sub    "
	Print #1
	Print #1,	"ExcelError:"
	If outputFile <> 2 Then
		Print #1,	"	if Err.Number = 1004& then xl.Quit	'Excel could not save the file. Terminate Excel
	End If
	Print #1,	"	MsgBox Err.Description, vbExclamation, ""Error Running Excel"""
	Print #1,	"	On Error GoTo 0 ' Disable error trapping.    "
	Print #1,	"	Exit Sub"
	Print #1,	"End Sub"
	If Table Then
		Print #1
		Print #1, "Sub OutCell (txt As String)"
		Print #1, "	Print #1, txt; vbTab;"
		Print #1, "End Sub"
	End If
End Sub

Function GetLastColName As String
	Dim n1, n2, n
	n = UBound(PropArray)
	If allCont Then n = n + 1
	n1 =  Asc("Z") - Asc("A") + 1
	GetLastColName = ""
	If n > n1 Then
		n2 = (n - 1) \ n1
		GetLastColName = Chr(n2 - 1 + Asc("A"))
		n = n - n2 * n1
	End If
	GetLastColName = GetLastColName + Chr(n - 1 + Asc("A"))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''
' RTF Format Code Generator
'''''''''''''''''''''''''''''''''''''''''''''''
Sub GenCodeForRTF
	Print #1
	Print #1,	"Sub BeginRTF"
	Print #1,	"	Print #1,	""{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fmodern Courier New;}}"""
	If colorColumns Then
		Print #1,	"	'Create Color Table. First color is always black."
		Print #1,	"	Print #1, ""{\colortbl;\red0\green0\blue0;"";"
		Print #1,	"	For i = 0 To UBound(Colors)"
		Print #1,	"		Print #1, ""\red"";   Trim(Str(Val( ""&H"" & Left (Colors(i), 2))));"
		Print #1,	"		Print #1, ""\green""; Trim(Str(Val( ""&H"" & Mid  (Colors(i), 3, 2))));"
		Print #1,	"		Print #1, ""\blue"";  Trim(Str(Val( ""&H"" & Right(Colors(i), 2)))); "";"";"
		Print #1,	"	Next"
		Print #1,	"	Print #1, ""}"""
	End If
	Print #1,	"End Sub"
	Print #1
	Print #1,	"Sub EndRTF"
	Print #1,	"	Print #1, ""}"""
	Print #1,	"End Sub"
	If InStr(code, "OutBoldLine") Then
		Print #1
		Print #1,	"Sub OutBoldLine (txt As String)"
		Print #1,	"	Print #1, ""\plain\f0\b ""; FormatText(txt); ""\plain\f0"""
		Print #1,	"	Print #1, ""\par "";"
		Print #1,	"End Sub"
	End If
	If InStr(code, "OutBoldUnLnLine") Then
		Print #1
		Print #1,	"Sub OutBoldUnLnLine (txt As String)"
		Print #1,	"	Print #1, ""\plain\f0\b\ul ""; FormatText(txt); ""\plain\f0"""
		Print #1,	"	Print #1, ""\par "";"
		Print #1,	"End Sub"
	End If
	If Table Then
		Print #1
		Print #1, "Sub OutCell (txt As String, Optional bold = False"; IIf(colorColumns, ", Optional color = -1)", ")")
		Print #1, "	If bold then Print #1, ""\b """
		If colorColumns Then
			Print #1, "	If color <> -1 then Print #1, ""\cf""; Trim(Str(color+2))"
		End If
		Print #1, "	w   = Widths(CurCol)"
		If Not enableAlign Then
			Print #1, "	txt = Left(txt, w)"
			Print #1, "	Print #1, FormatText(txt); Space(w - Len(txt)"; IIf(drawVBorder, "", " + 1"); ");"
		Else
			Print #1, "	fmt = FormatText(txt)"
			Print #1, "	txt = Left(txt, w)"
			Print #1, "	n   = w - Len(txt)"
			Print #1, "	If Align(CurCol) = 0 Then"
			Print #1, "		Print #1, fmt; Space(n"; IIf(drawVBorder, "", " + 1"); ");"
			Print #1, "	ElseIf Align(curCol) = 1 Then
			Print #1, "		Print #1, Space(n); fmt;"; IIf(drawVBorder, "", " Space(1);")
			Print #1, "	Else
			Print #1, "		Print #1, Space(n\2); fmt; Space(n - n\2"; IIf(drawVBorder, "", " + 1"); ");"
			Print #1, "	End If
		End If
		If colorColumns Then
			Print #1, "	If bold or color <> -1 then Print #1, ""\plain """
		Else
			Print #1, "	If bold then Print #1, ""\plain """
		End If
		If drawVBorder Then
			Print #1, "	Print #1, "; Chr(34); IIf(drawVBorder, " | ", " "); Chr(34); ";"
		End If
		Print #1, "	CurCol = CurCol + 1"
		Print #1, "End Sub"
	End If
	Print #1
	Print #1,	"Sub OutLine (Optional txt As String = """")"
	Print #1,	"	Print #1, FormatText(txt)"
	Print #1,	"	Print #1, ""\par "";"
	Print #1,	"End Sub"
	Print #1
	Print #1,	"'Replace every single '\' symbol with '\\' otherwise it will be considered as RTF keyword
	Print #1,	"Function FormatText(ByVal txt As String) As String"
	Print #1,	"	p = 1"
	Print #1,	"	While InStr(p, txt, ""\"") <> 0"
	Print #1,	"		p = InStr(p, txt, ""\"")"
	Print #1,	"		txt = Left(txt, p - 1) & ""\\"" & Mid(txt, p+1)"
	Print #1,	"		p = p + 2"
	Print #1,	"	Wend"
	Print #1,	"	FormatText = txt"
	Print #1,	"End Function"
End Sub	

'''''''''''''''''''''''''''''''''''''''''''''''
' HTML Format Code Generator
'''''''''''''''''''''''''''''''''''''''''''''''
Sub GenCodeForHTML
	If Not Table And (InStr(code, "NewCell") Or InStr(code, "NewRow")) Then
		Print #1
		Print #1, "Const NewCell = ""<TD>"""
		Print #1, "Const NewRow = ""<TR>"""
	End If
	Print #1
	Print #1, "Const Spc = "" &nbsp;""	'HTML space"
	Print #1
	Print #1,	"Sub BeginHTML"
	Print #1,	"	Print #1, ""<HTML>"""
	Print #1,	"	Print #1, ""<HEAD>"""
	Print #1,	"	Print #1, ""<TITLE>"; header; "</TITLE>"""
	Print #1,	"	Print #1, ""</HEAD>"""
	Print #1,	"	Print #1, ""<BODY>"""
	Print #1,	"End Sub"
	Print #1
	Print #1,	"Sub EndHTML"
	Print #1,	"	Print #1, ""</BODY>"""
	Print #1,	"	Print #1, ""</HTML>"""
	Print #1,	"End Sub"
	If InStr(code, "OutBoldLine") Then
		Print #1
		Print #1,	"Sub OutBoldLine (txt As String)"
		Print #1,	"	Print #1, ""<B>""; txt; ""</B><BR>"""
		Print #1,	"End Sub"
	End If
	If InStr(code, "OutBoldUnLnLine") Then
		Print #1
		Print #1,	"Sub OutBoldUnLnLine (txt As String)"
		Print #1,	"	Print #1, ""<B><U>""; txt; ""</B></U><BR>"""
		Print #1,	"End Sub"
	End If
	If InStr(code, "OutTitle") Then
		Print #1
		Print #1,	"Sub OutTitle (txt As String, size As Integer)"
		Print #1,	"	Print #1, ""<FONT SIZE = ""; size; ""><P ALIGN=CENTER><B>""; txt; ""</B></P></FONT>"""
		Print #1,	"End Sub"
	End If
	Print #1
	Print #1,	"Sub OutLine (Optional txt As String = """")"
	Print #1,	"	Print #1, txt; ""<BR>"""
	Print #1,	"End Sub"
	If secObj <> 0 Then
		Print #1
		Print #1,	"Sub BeginIndent"
		Print #1,	"	Print #1, ""<DIR>"""
		Print #1,	"End Sub"
		Print #1
		Print #1,	"Sub EndIndent"
		Print #1,	"	Print #1, ""</DIR>"""
		Print #1,	"End Sub"
	End If
	If Not Table Then
		If InStr(code, "BeginTable") Then
			Print #1
			Print #1,	"Sub BeginTable"
			Print #1,	"	Print #1, ""<TABLE CELLPADDING=5><TR>"""
			Print #1,	"End Sub"
		End If
		If InStr(code, "EndTable") Then
			Print #1
			Print #1,	"Sub EndTable"
			Print #1,	"	Print #1, ""</TABLE>"""
			Print #1,	"End Sub"
		End If
	End If
	If Table Then
		Print #1
		Print #1,	"Sub BeginTable"
		If colorTable Then
			Print #1,	"	RowType = 0"
		End If
		Print #1,	"	Print #1, ""<TABLE CELLSPACING=1 BORDER=1 CELLPADDING=2>"""
		Print #1,	"End Sub"
		Print #1
		Print #1,	"Sub EndTable"
		Print #1,	"	Print #1, ""</TABLE>"""
		Print #1,	"End Sub"
		Print #1
		Print #1,	"Sub BeginRow"
		If Not colorTable Then
			Print #1,	"	Print #1, ""<TR>"""
		Else
			If Not bkgDef Then
				Print #1,	"	Print #1, ""<TR BGCOLOR = #""; Colors(RowType); "">"""
			End If
		End If
		Print #1,	"End Sub"
		Print #1
		Print #1,	"Sub EndRow"
		Print #1,	"	Print #1, ""</TR>"""
		If colorTable Then
			Print #1,	"	RowType = IIf (RowType = 0, 1, 3 - RowType)"
		End If
		Print #1,	"End Sub"
		'Output Cell Sub
		Print #1
		Print #1,	"Sub OutCell (txt As String, Optional bold = False"; IIf(colorColumns, ", Optional color = -1)", ")")
		Print #1,	"	If txt = """" Then txt = Spc"
		Print #1,	"	Print #1, ""<TD>""; "
		Print #1,	"	If bold Then Print #1, ""<B>"""
		If enableAlign Then
			Print #1,	"	Print #1, ""<P ALIGN=""; Choose(Align(CurCol) + 1, ""LEFT"", ""RIGHT"", ""CENTER""); "">"""
		End If
		If colorTable Or colorColumns Then
			If colorColumns Then
				If colorTable And bkgDef And Not fontDef Then
					Print #1,	"	Print #1, ""<FONT COLOR = #""; Colors(IIf(color = -1, RowType, color)); "">"""
				ElseIf colorTable And Not bkgDef And Not fontDef Then
					Print #1,	"	Print #1, ""<FONT COLOR = #""; Colors(IIf(color = -1, 3 + RowType, color)); "">"""
				Else
					Print #1,	"	If color <> -1 Then Print #1, ""<FONT COLOR = #""; Colors(color); "">"""
				End If
			ElseIf Not fontDef Then
				If bkgDef Then
					Print #1,	"	Print #1, ""<FONT COLOR = #""; Colors(RowType); "">"""
				Else
					Print #1,	"	Print #1, ""<FONT COLOR = #""; Colors(3 + RowType); "">"""
				End If
			End If
		End If
		Print #1,	"	Print #1, txt;"
		If colorTable And Not fontDef Then
			Print #1,	"	Print #1, ""</FONT>"""
		ElseIf colorColumns Then
			Print #1,	"	If color <> -1 Then Print #1, ""</FONT>"""
		End If
		If enableAlign Then
			Print #1,	"	Print #1, ""</P>"""
			Print #1, 	"	CurCol = CurCol + 1"
		End If
		Print #1,	"	If bold Then Print #1, ""</B>"""
		Print #1,	"	Print #1, ""</TD>"""
		Print #1,	"End Sub"
	End If
	If InStr(code, "OutHyperLink") Then
			Print #1
			Print #1,	"Sub OutHyperLink (LinkName As String, LinkText As String)"
			Print #1,	"	Print #1, ""<A HREF=""; Chr(34); ""#""; LinkName; Chr(34); "">""; LinkText; ""</A>"""
			Print #1,	"End Sub"
	End If
	If InStr(code, "OutBookmark") Then
			Print #1
			Print #1,	"Sub OutBookmark (LinkName As String)"
			Print #1,	"	Print #1, ""<A Name=""; Chr(34); LinkName; Chr(34); ""></A>"""
			Print #1,	"End Sub"
	End If
	Print #1
	Print #1, "'Declare Win32 function to run the application registered to handle a given file type (HTML in our case)"
	Print #1, "Private Declare Function ShellExecute Lib ""shell32.dll"" Alias ""ShellExecuteA"" (ByVal hwnd As Long , ByVal op As String, ByVal fname As String, ByVal par As String, ByVal lpDirectory As String, ByVal cmdShow As Long) As Long
End Sub	

Sub TableOfContents
	If HTML And contContents And (allCont Or secObj <> 0) Then
		Dim indexHeader, coll
		indexHeader = IIf(secObj <> 0, "Index", "Table of Contents")
		Print #1
		Print #1, Indent; "'Output "; indexHeader
		Print #1, Indent; "OutBoldLine "; Chr(34); indexHeader; Chr(34)
		If showProgress And secObj <> 0 And FilterCondition <> "" Then
			Print #1, Indent; "Total = Total * 2"
		End If
		If allCont Then
			Print #1, Indent; "For Each "; contObj; " In "; contColl
			nTabs = nTabs + 1
			Print #1, Indent; "OutHyperLink "; contNameProp; ", "; contNameProp
			Print #1, Indent; NewLine
		End If
		If secObj <> 0 Then
			Print #1, Indent; "BeginIndent"
			coll = primSortColl
			FormatTemplate coll, "xxx", contObj
			Print #1, Indent; "For Each "; primVarName; " In "; coll
			If FilterCondition <> "" Then
				OutputMultiLineCode "If " & FilterCondition & " Then"
				nTabs = nTabs + 1
			End If
			If bInstalledOnly Then 		
				Print #1, Indent; "	If "; primVarName; ".Installed Then"
				nTabs = nTabs + 1
			End If				
			Print #1, Indent; "	Print #1, Spc;"
			Print #1, Indent; "	OutHyperLink "; IIf(allCont And repByContainer, contNameProp & " & ", ""); primNameProp; ", "; primNameProp
			If bInstalledOnly Then
				nTabs = nTabs - 1
				Print #1, Indent; "	End If"
			End If
			If FilterCondition <> "" Then
				nTabs = nTabs - 1
				Print #1, Indent; "	End If"
			End If
			If showProgress And FilterCondition <> "" Then
				Print #1, Indent; "	Processed = Processed + 1
				Print #1, Indent; "	ProgressBar = Processed * 100 / Total
			End If
			Print #1, Indent; "Next"
			Print #1, Indent; "EndIndent"
		End If
		If allCont Then
			nTabs = nTabs - 1
			Print #1, Indent; "Next"
		End If
		Print #1
	End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''
' Text Format Code Generator
'''''''''''''''''''''''''''''''''''''''''''''''
Sub GenCodeForText
	If InStr(code, "OutCell") Then
		Dim brd
		brd = IIf(drawVBorder, "); "" | "";", " + 1);")
		Print #1
		Print #1, "Sub OutCell (txt As String)"
		Print #1, "	w   = Widths(CurCol)"
		Print #1, "	txt = Left(txt, w)"
		If Not enableAlign Then
			Print #1, "	Print #1, txt; Space(w - Len(txt)"; brd
		Else
			Print #1, "	n   = w - Len(txt)"
			Print #1, "	If Align(CurCol) = 0 Then"
			Print #1, "		Print #1, txt; Space(n"; brd
			Print #1, "	ElseIf Align(curCol) = 1 Then
			Print #1, "		Print #1, Space(n); txt;"; IIf(drawVBorder, " "" | "";", " Space(1);")
			Print #1, "	Else
			Print #1, "		Print #1, Space(n\2); txt; Space(n - n\2"; brd
			Print #1, "	End If
		End If
		Print #1, "	CurCol = CurCol + 1"
		Print #1, "End Sub"
	End If
End Sub

Sub HelpPage
	Dim bc, col, row, rows, chart
	chart = "chart.html"
	Open chart For Output As #2
	Print #2,	"<HTML>"
	Print #2,	"<HEAD>"
	Print #2,	"<TITLE>Wizard Features</TITLE>"
	Print #2,	"</HEAD>"
	Print #2,	"<BODY>"
	Print #2, 	"<FONT SIZE = 5><B>Availability of Wizard Features Depending on Format</B></P></FONT>"
	Print #2,	"<TABLE CELLSPACING=1 BORDER=1 CELLPADDING=2 WIDTH=100%>"
	rows = Array( _
		Array("Wizard Features / Options",				"Plain Text (Notepad)",	"Rich Text (WordPad)",	"Spreadsheet (MS&nbsp;Excel)",	"HTML (Browser)"), _
		Array("Tables", 								"Yes",			"Yes",			"Yes",		"Yes"), _
		Array("Colored Schemes for Tables", 			"No",			"No",			"No",		"Yes"), _
		Array("Format Headers to be in Bold", 			"No",			"Yes",			"Yes",		"Yes"), _
		Array("Conditional Reports", 					"Yes",			"Yes",			"Yes",		"Yes"), _
		Array("Conditional Table Formatting", 			"No",			"Yes",			"No",		"Yes"), _
		Array("Hyperlink Table Of Contents / Index", 	"No",			"No",			"No",		"Yes"), _
		Array("Dynamic Table Data Filters", 			"No",			"No",			"Always",	"No"), _
		Array("Auto Subtotals", 						"No",			"No",			"Yes",		"No"), _
		Array("Column Text Alignment", 					"Yes",			"Yes",			"Yes",		"Yes"), _
		Array("Setting Exact Field Width in Characters","Yes",			"Yes",			"No",		"No"), _
		Array("Auto-fit Columns to Maximum Text", 		"No",			"No",			"Yes",		"Yes"), _
		Array("Option to Draw Table Borders", 			"Yes",			"Yes",			"No",		"Always"), _
		Array("Pass Output Data Through Clipboard",		"Yes",			"No",			"Yes",		"No") _
	)
	For row = 0 To UBound(rows)
		If row = 0 Then
			bc = "D5D2E3"
		ElseIf row Mod 2 = 1 Then
			bc = "BBDBBD"
		Else
			bc = "A6D0A8"
		End If
		Print #2, "<TR BGCOLOR = #"; bc; ">";
		For col = 0 To UBound(rows(0))
			Print #2, "<TD>"; 
			If col <> 0 Or row = 0 Then
				Print #2, "<B>"; 
			End If
			If col > 0 Then
				Print #2, "<P ALIGN=CENTER>"
			End If
			txt = rows(row)(col)
			If txt = "No" Then
				Print #2, "<FONT COLOR = red>"
			ElseIf txt = "Yes" Or txt = "Always" Then
				Print #2, "<FONT COLOR = green>"
			End If
			Print #2, txt;
			Print #2, "</TD>"
		Next
		Print #2, "</TR>";
	Next
	Print #2,	"</TABLE>"
	Print #2,	"</BODY>"
	Print #2,	"</HTML>"
	Close #2
	ShellExecute (0, "Open", chart, "", "", 5)
End Sub

