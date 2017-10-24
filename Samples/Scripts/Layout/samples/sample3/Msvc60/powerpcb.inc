// this ALWAYS GENERATED file contains the definitions for the Automation interfaces

// Enums

typedef enum
    {	ppcbObjectTypeUnknown	= 0,
	ppcbObjectTypeComponent	= 1,
	ppcbObjectTypeNet	= 2,
	ppcbObjectTypePin	= 3,
	ppcbObjectTypeVia	= 4,
	ppcbObjectTypeConnection	= 5,
	ppcbObjectTypeRouteSegment	= 6,
	ppcbObjectTypeJumper	= 7,
	ppcbObjectTypePartType	= 8,
	ppcbObjectTypeCBP	= 9,
	ppcbObjectTypeSBP	= 10,
	ppcbObjectTypeWirebond	= 11,
	ppcbObjectTypeNetClass	= 12,
	ppcbObjectTypeDrawing	= 13,
	ppcbObjectTypeText	= 14,
	ppcbObjectTypeLabel	= 15,
	ppcbObjectTypePolyline	= 16,
	ppcbObjectTypeCircle	= 17,
	ppcbObjectTypeLibrary	= 18,
	ppcbObjectTypeLibraryItem	= 19,
	ppcbObjectTypeAll	= 9999,
	ppcbObjectTypeApplication	= 20,
	ppcbObjectTypeAttribute	= 21,
	ppcbObjectTypeAttributeType	= 22,
	ppcbObjectTypeDocument	= 23,
	ppcbObjectTypeMeasure	= 24,
	ppcbObjectTypeView	= 25,
	ppcbObjectTypeAssemblyOptions	= 26,
	ppcbObjectTypeAttributes	= 27,
	ppcbObjectTypeAttributeTypes	= 28,
	ppcbObjectTypeObjects	= 29
    }	PPcbObjectType;

typedef enum
    {	ppcbUnitCurrent	= 0,
	ppcbUnitDatabase	= 1,
	ppcbUnitMils	= 2,
	ppcbUnitInch	= 3,
	ppcbUnitMetric	= 4
    }	PPcbUnit;

typedef enum
    {	ppcbJustifyLeft	= 0,
	ppcbJustifyHCenter	= 1,
	ppcbJustifyRight	= 2
    }	PPcbHorizontalJustification;

typedef enum
    {	ppcbJustifyTop	= 0,
	ppcbJustifyVCenter	= 1,
	ppcbJustifyBottom	= 2
    }	PPcbVerticalJustification;

typedef enum
    {	ppcbOutlineTypeCenter	= 0,
	ppcbOutlineTypeOuter	= 1,
	ppcbOutlineTypeInner	= 2
    }	PPcbOutlineType;

typedef enum
    {	ppcbDrw2Dline	= 0,
	ppcbDrwBoard	= 1,
	ppcbDrwCopper	= 3,
	ppcbDrwCopperPour	= 6,
	ppcbDrwCopperHatch	= 7,
	ppcbDrwCopperThermal	= 8,
	ppcbDrwKeepout	= 9
    }	PPcbDrawingType;

typedef enum
    {	ppcbLabelTypeRefDesignator	= 0,
	ppcbLabelTypePartType	= 1,
	ppcbLabelTypeAttribute	= 2
    }	PPcbLabelType;

typedef enum
    {	ppcbOriginTypeDesign	= 0,
	ppcbOriginTypeParent	= 1
    }	PPcbOriginType;

typedef enum
    {	ppcbShapeTypeOpen	= 0,
	ppcbShapeTypeHollow	= 1,
	ppcbShapeTypeFilled	= 2,
	ppcbShapeTypeVoid	= 3
    }	PPcbShapeType;

typedef enum
    {	ppcbSegmentUnknown	= 0,
	ppcbSegmentLine	= 1,
	ppcbSegmentArc	= 2
    }	PPcbSegmentType;

typedef enum
    {	ppcbTestPointNone	= 0,
	ppcbTestPointTopLayer	= 1,
	ppcbTestPointBottomLayer	= 2
    }	PPcbTestPointType;

typedef enum
    {	ppcbGridNone	= 0,
	ppcbGridDesign	= 1,
	ppcbGridVia	= 2,
	ppcbGridDisplay	= 3,
	ppcbGridAll	= 9999
    }	PPcbGridType;

typedef enum
    {	ppcbLayerUnknown	= 0,
	ppcbLayerComponent	= 1,
	ppcbLayerRouting	= 2,
	ppcbLayerDrill	= 3,
	ppcbLayerSilkscreen	= 4,
	ppcbLayerPasteMask	= 5,
	ppcbLayerSolderMask	= 6,
	ppcbLayerAssembly	= 7,
	ppcbLayerGeneral	= 8,
	ppcbLayerAll	= 9999
    }	PPcbLayerType;

typedef enum
    {	ppcbASCIIVerCurrent	= 0,
	ppcbASCIIVer1_1	= 2,
	ppcbASCIIVer1_5	= 3,
	ppcbASCIIVer2_0	= 4,
	ppcbASCIIVer2_1	= 5,
	ppcbASCIIVer3_0	= 7,
	ppcbASCIIVer3_5	= 8,
	ppcbASCIIVer4_0	= 9
    }	PPcbASCIIVersion;

typedef enum
    {	ppcbASCIISectionPCB	= 0x1,
	ppcbASCIISectionReuse	= 0x2,
	ppcbASCIISectionText	= 0x4,
	ppcbASCIISectionLines	= 0x8,
	ppcbASCIISectionClusters	= 0x10,
	ppcbASCIISectionVias	= 0x20,
	ppcbASCIISectionDecals	= 0x40,
	ppcbASCIISectionPartTypes	= 0x80,
	ppcbASCIISectionParts	= 0x100,
	ppcbASCIISectionJumpers	= 0x200,
	ppcbASCIISectionConnections	= 0x400,
	ppcbASCIISectionRoutes	= 0x800,
	ppcbASCIISectionTeardrops	= 0x1000,
	ppcbASCIISectionMisc	= 0x2000,
	ppcbASCIISectionRules	= 0x4000,
	ppcbASCIISectionCAM	= 0x8000,
	ppcbASCIISectionPour	= 0x10000,
	ppcbASCIISectionAssemblyOptions	= 0x20000,
	ppcbASCIISectionTestPoints	= 0x40000,
	ppcbASCIISectionAttributes	= 0x80000,
	ppcbASCIISectionAll	= 0xffffffff
    }	PPcbASCIISections;

typedef enum
    {	ppcbAttrNone	= 0,
	ppcbAttrPart	= 0x1,
	ppcbAttrNet	= 0x2,
	ppcbAttrPin	= 0x4,
	ppcbAttrVia	= 0x8,
	ppcbAttrPCB	= 0x10,
	ppcbAttrPartType	= 0x20,
	ppcbAttrDecal	= 0x40,
	ppcbAttrNetClass	= 0x80,
	ppcbAttrAll	= 0xffffffff
    }	PPcbAttrFlags;

typedef enum
    {	ppcbElectricalTypeUnknown	= 0,
	ppcbElectricalTypeSource	= 1,
	ppcbElectricalTypeBidirectional	= 2,
	ppcbElectricalTypeOpenCollector	= 3,
	ppcbElectricalTypeOrTieableSource	= 4,
	ppcbElectricalTypeTristate	= 5,
	ppcbElectricalTypeLoad	= 6,
	ppcbElectricalTypeTerminator	= 7,
	ppcbElectricalTypePower	= 8,
	ppcbElectricalTypeGround	= 9
    }	PPcbPinElectricalType;

typedef enum
    {	ppcbMeasureFormatStandard	= 0,
	ppcbMeasureFormatCurrent	= 1,
	ppcbMeasureFormatShort	= 2,
	ppcbMeasureFormatLong	= 3
    }	PPcbMeasureFormat;

typedef enum
    {	ppcbBondPadShapeUnknown	= 0,
	ppcbBondPadShapeRectangle	= 1,
	ppcbBondPadShapeOval	= 2
    }	PPcbBondPadShape;

typedef enum
    {	ppcbBondPadEdgeUnknown	= 0,
	ppcbBondPadEdgeLeft	= 1,
	ppcbBondPadEdgeTop	= 2,
	ppcbBondPadEdgeRight	= 3,
	ppcbBondPadEdgeBottom	= 4
    }	PPcbBondPadEdge;

typedef enum
    {	ppcbLabelDisplayNone	= 0,
	ppcbLabelDisplayValue	= 1,
	ppcbLabelDisplayNameAndValue	= 2,
	ppcbLabelDisplayFullNameAndValue	= 3
    }	PPcbLabelDisplayMode;

typedef enum
    {	ppcbRightReadingNone	= 0,
	ppcbRightReadingOrthogonal	= 1,
	ppcbRightReadingAngled	= 2
    }	PPcbRightReadingStatus;

typedef enum
    {	ppcbLibraryItemTypePartType	= 0,
	ppcbLibraryItemTypeDecal	= 1,
	ppcbLibraryItemTypeLogicDrawing	= 2,
	ppcbLibraryItemTypeDrawing	= 3,
	ppcbLibraryItemTypeAll	= 9999
    }	PPcbLibraryItemType;

typedef enum
    {	ppcbDRCNone	= 0,
	ppcbDRCOff	= 1,
	ppcbDRCWarn	= 2,
	ppcbDRCIgnoreClearance	= 3,
	ppcbDRCPrevent	= 4
    }	PPcbDRCMode;

typedef enum
    {	ppcbNudgeNone	= 0,
	ppcbNudgeOff	= 1,
	ppcbNudgeWarn	= 2,
	ppcbNudgeAuto	= 3
    }	PPcbNudgeMode;


// IIDs/CLSIDs
const IID IID_IPowerPCBApp = {0x0821EF02,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBDoc = {0x0821EF04,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBView = {0x0821EF06,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBObjs = {0x0821EF08,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBComp = {0x0821EF09,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBPartType = {0x0821EF12,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBNet = {0x0821EF0A,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBPin = {0x0821EF0B,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBVia = {0x0821EF0C,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBConnection = {0x0821EF0D,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBRouteSegment = {0x0821EF10,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBJumper = {0x0821EF11,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBAttrs = {0x0821EF0E,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBAttr = {0x0821EF0F,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBAttrTypes = {0x0821EF2D,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBAttrType = {0x0821EF2F,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBMeasure = {0x0821EF2B,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBAssOpts = {0x0821EF07,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID IID_IPowerPCBCBP = {0x20A2BB62,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const IID IID_IPowerPCBSBP = {0x20A2BB60,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const IID IID_IPowerPCBWirebond = {0x20A2BB64,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const IID IID_IPowerPCBNetClass = {0x20A2BB66,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const IID IID_IPowerPCBPolyline = {0x20A2BB68,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const IID IID_IPowerPCBCircle = {0x20A2BB6A,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const IID IID_IPowerPCBDrawing = {0x20A2BB6E,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const IID IID_IPowerPCBTextBase = {0x20A2BB6C,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const IID IID_IPowerPCBText = {0x20A2BB70,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const IID IID_IPowerPCBLabel = {0x20A2BB72,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const IID IID_IPowerPCBLibrary = {0x20A2BB74,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const IID IID_IPowerPCBLibraryItem = {0x20A2BB76,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const IID LIBID_PowerPCB = {0x0821EF00,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID DIID__PowerPCBAppEvents = {0x0821EF22,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBApplication = {0x0821EF01,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID DIID__PowerPCBDocEvents = {0x0821EF24,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBDocument = {0x0821EF14,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const IID DIID__PowerPCBViewEvents = {0x0821EF26,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBView = {0x0821EF16,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBObjects = {0x0821EF18,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBComponent = {0x0821EF19,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBNet = {0x0821EF1A,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBPin = {0x0821EF1B,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBVia = {0x0821EF1C,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBConnection = {0x0821EF1D,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBRouteSegment = {0x0821EF20,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBJumper = {0x0821EF21,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBPartType = {0x0821EF2A,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBAttributes = {0x0821EF1E,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBAttribute = {0x0821EF1F,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBAttributeTypes = {0x0821EF2E,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBAttributeType = {0x0821EF30,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBMeasure = {0x0821EF2C,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBAssemblyOptions = {0x0821EF17,0xB880,0x11CF,{0x82,0x68,0x00,0x20,0xAF,0x93,0x79,0x87}};
const CLSID CLSID_PowerPCBCBP = {0x20A2BB63,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const CLSID CLSID_PowerPCBSBP = {0x20A2BB61,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const CLSID CLSID_PowerPCBWirebond = {0x20A2BB65,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const CLSID CLSID_PowerPCBNetClass = {0x20A2BB67,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const CLSID CLSID_PowerPCBPolyline = {0x20A2BB69,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const CLSID CLSID_PowerPCBCircle = {0x20A2BB6B,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const CLSID CLSID_PowerPCBDrawing = {0x20A2BB6F,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const CLSID CLSID_PowerPCBText = {0x20A2BB71,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const CLSID CLSID_PowerPCBLabel = {0x20A2BB73,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const CLSID CLSID_PowerPCBLibrary = {0x20A2BB75,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};
const CLSID CLSID_PowerPCBLibraryItem = {0x20A2BB77,0xE608,0x11d3,{0x93,0xB5,0x32,0x82,0xA8,0x00,0x00,0x00}};

// Dispatch Ids for methods/properties/events

// IPowerPCBApp Properties
#define IPOWERPCBAPP_NAME                                  0
#define IPOWERPCBAPP_APPLICATION                           10001
#define IPOWERPCBAPP_PARENT                                10002
#define IPOWERPCBAPP_ACTIVEDOCUMENT                        10003
#define IPOWERPCBAPP_DEFAULTFILEPATH                       10004
#define IPOWERPCBAPP_FULLNAME                              10005
#define IPOWERPCBAPP_VISIBLE                               10006
#define IPOWERPCBAPP_PREFERENCE                            10007
#define IPOWERPCBAPP_STATUSBARTEXT                         10008
#define IPOWERPCBAPP_PROGRESSBAR                           10009
#define IPOWERPCBAPP_VERSION                               10010
#define IPOWERPCBAPP_OBJECTTYPE                            10011
#define IPOWERPCBAPP_LIBRARIES                             10012
// IPowerPCBApp Methods
#define IPOWERPCBAPP_OPENDOCUMENT                          10500
#define IPOWERPCBAPP_QUIT                                  10501
#define IPOWERPCBAPP_RUNMACRO                              10502
#define IPOWERPCBAPP_REGISTERSINK                          10503
#define IPOWERPCBAPP_REGISTERSINKNOTIFICATION              10504
#define IPOWERPCBAPP_UNREGISTERSINK                        10505
#define IPOWERPCBAPP_UNREGISTERSINKNOTIFICATION            10506
#define IPOWERPCBAPP_LOCKSERVER                            10507
#define IPOWERPCBAPP_UNLOCKSERVER                          10508
#define IPOWERPCBAPP_PROCESSCOMMAND                        10509
#define IPOWERPCBAPP_PROCESSPARAMETER                      10510
#define IPOWERPCBAPP_PROCESSPOINTER                        10511
#define IPOWERPCBAPP_STRNUMCMP                             10512
#define IPOWERPCBAPP_MEASURE                               10513
#define IPOWERPCBAPP_GETLIBRARYITEMS                       10514
#define IPOWERPCBAPP_EXPORTLIBRARYITEMS                    10515
// IPowerPCBApp Events
#define IPOWERPCBAPP_OPENDOCUMENTEVENT                     10700
#define IPOWERPCBAPP_QUITEVENT                             10701
#define IPOWERPCBAPP_PROGRESSEVENT                         10702

// IPowerPCBDoc Properties
#define IPOWERPCBDOC_NAME                                  0
#define IPOWERPCBDOC_APPLICATION                           11001
#define IPOWERPCBDOC_PARENT                                11002
#define IPOWERPCBDOC_PATH                                  11003
#define IPOWERPCBDOC_FULLNAME                              11004
#define IPOWERPCBDOC_ACTIVEVIEW                            11005
#define IPOWERPCBDOC_SAVED                                 11006
#define IPOWERPCBDOC_UNIT                                  11007
#define IPOWERPCBDOC_ORIGINX                               11008
#define IPOWERPCBDOC_ORIGINY                               11009
#define IPOWERPCBDOC_GRIDX                                 11010
#define IPOWERPCBDOC_GRIDY                                 11011
#define IPOWERPCBDOC_LAYERNAME                             11012
#define IPOWERPCBDOC_ASSEMBLYOPTIONS                       11013
#define IPOWERPCBDOC_PREFERENCE                            11014
#define IPOWERPCBDOC_ATTRIBUTES                            11017
#define IPOWERPCBDOC_LAYERCOUNT                            11018
#define IPOWERPCBDOC_LAYERTYPE                             11019
#define IPOWERPCBDOC_COMPONENTS                            11020
#define IPOWERPCBDOC_NETS                                  11021
#define IPOWERPCBDOC_PINS                                  11022
#define IPOWERPCBDOC_VIAS                                  11023
#define IPOWERPCBDOC_CONNECTIONS                           11024
#define IPOWERPCBDOC_ROUTESEGMENTS                         11025
#define IPOWERPCBDOC_BOARDOUTLINESURFACE                   11026
#define IPOWERPCBDOC_JUMPERS                               11027
#define IPOWERPCBDOC_PARTTYPES                             11029
#define IPOWERPCBDOC_ATTRIBUTETYPES                        11030
#define IPOWERPCBDOC_OBJECTTYPE                            11031
#define IPOWERPCBDOC_NETCLASSES                            11032
#define IPOWERPCBDOC_DRAWINGS                              11033
#define IPOWERPCBDOC_TEXTS                                 11034
#define IPOWERPCBDOC_LAYER_ENABLED                         11035
#define IPOWERPCBDOC_ELECTRICALLAYERCOUNT                  11036
// IPowerPCBDoc Methods
#define IPOWERPCBDOC_ACTIVATE                              11500
#define IPOWERPCBDOC_GETOBJECTS                            11501
#define IPOWERPCBDOC_SELECTOBJECTS                         11502
#define IPOWERPCBDOC_SAVE                                  11503
#define IPOWERPCBDOC_SAVEAS                                11504
#define IPOWERPCBDOC_IMPORTNETLIST                         11505
#define IPOWERPCBDOC_EXPORTNETLIST                         11506
#define IPOWERPCBDOC_IMPORTECOFILE                         11507
#define IPOWERPCBDOC_EXPORTECOFILE                         11508
#define IPOWERPCBDOC_CHECKASCII                            11509
#define IPOWERPCBDOC_EXPORTRULES                           11510
#define IPOWERPCBDOC_GETPLACEDOBJECTS                      11511
#define IPOWERPCBDOC_EXPORTASCII                           11512
#define IPOWERPCBDOC_ADDTEXT                               11513
// IPowerPCBDoc Events
#define IPOWERPCBDOC_SAVEEVENT                             11700
#define IPOWERPCBDOC_SELECTIONCHANGEEVENT                  11701
#define IPOWERPCBDOC_POSITIONSCHANGEEVENT                  11702
#define IPOWERPCBDOC_SECURITYLIMITEVENT                    11703

// IPowerPCBView Properties
#define IPOWERPCBVIEW_NAME                                 0
#define IPOWERPCBVIEW_APPLICATION                          12001
#define IPOWERPCBVIEW_PARENT                               12002
#define IPOWERPCBVIEW_TOPLEFTX                             12003
#define IPOWERPCBVIEW_TOPLEFTY                             12004
#define IPOWERPCBVIEW_BOTTOMRIGHTX                         12005
#define IPOWERPCBVIEW_BOTTOMRIGHTY                         12006
#define IPOWERPCBVIEW_CENTERX                              12007
#define IPOWERPCBVIEW_CENTERY                              12008
#define IPOWERPCBVIEW_ZOOM                                 12009
#define IPOWERPCBVIEW_OBJECTTYPE                           12010
// IPowerPCBView Methods
#define IPOWERPCBVIEW_REFRESH                              12500
#define IPOWERPCBVIEW_PAN                                  12501
#define IPOWERPCBVIEW_SETEXTENTSTOALL                      12502
#define IPOWERPCBVIEW_SETEXTENTSTOBOARD                    12503
#define IPOWERPCBVIEW_SETEXTENTS                           12504
#define IPOWERPCBVIEW_SETSCALE                             12505
#define IPOWERPCBVIEW_SETEXTENTSTOSELECTION                12506
// IPowerPCBView Events
#define IPOWERPCBVIEW_CHANGEEVENT                          12700

// IPowerPCBObjs Properties
#define IPOWERPCBOBJS_ITEM                                 0
#define IPOWERPCBOBJS_APPLICATION                          13001
#define IPOWERPCBOBJS_PARENT                               13002
#define IPOWERPCBOBJS_COUNT                                13003
#define IPOWERPCBOBJS_ITEMTYPE                             13004
#define IPOWERPCBOBJS_NEXT                                 13005
#define IPOWERPCBOBJS_NEWENUM                              13006
#define IPOWERPCBOBJS_OBJECTTYPE                           13007
#define IPOWERPCBOBJS_PARENTOBJECT                         13008
// IPowerPCBObjs Methods
#define IPOWERPCBOBJS_ADD                                  13500
#define IPOWERPCBOBJS_MERGE                                13501
#define IPOWERPCBOBJS_REMOVE                               13502
#define IPOWERPCBOBJS_RESET                                13503
#define IPOWERPCBOBJS_SELECT                               13504
#define IPOWERPCBOBJS_SORT                                 13505

// IPowerPCBComp Properties
#define IPOWERPCBCOMP_NAME                                 0
#define IPOWERPCBCOMP_APPLICATION                          14001
#define IPOWERPCBCOMP_PARENT                               14002
#define IPOWERPCBCOMP_SELECTED                             14003
#define IPOWERPCBCOMP_PARTTYPE                             14004
#define IPOWERPCBCOMP_DECAL                                14005
#define IPOWERPCBCOMP_LAYER                                14006
#define IPOWERPCBCOMP_LAYERNAME                            14007
#define IPOWERPCBCOMP_PINS                                 14008
#define IPOWERPCBCOMP_POSITIONX                            14009
#define IPOWERPCBCOMP_POSITIONY                            14010
#define IPOWERPCBCOMP_ORIENTATION                          14011
#define IPOWERPCBCOMP_ISSMD                                14012
#define IPOWERPCBCOMP_GLUED                                14013
#define IPOWERPCBCOMP_INSTALLED                            14014
#define IPOWERPCBCOMP_DECALCOMPATIBLELIST                  14015
#define IPOWERPCBCOMP_PLACED                               14016
#define IPOWERPCBCOMP_ATTRIBUTES                           14017
#define IPOWERPCBCOMP_DECALATTRIBUTES                      14018
#define IPOWERPCBCOMP_PARTTYPEATTRIBUTES                   14019
#define IPOWERPCBCOMP_SUBSTITUTED                          14020
#define IPOWERPCBCOMP_OBJECTTYPE                           14021
#define IPOWERPCBCOMP_PARTTYPELOGIC                        14022
#define IPOWERPCBCOMP_PARTTYPEECOREGISTERED                14024
#define IPOWERPCBCOMP_PARTTYPEOBJECT                       14025
#define IPOWERPCBCOMP_ISDIEPART                            14026
#define IPOWERPCBCOMP_DIELENGTH                            14027
#define IPOWERPCBCOMP_DIEWIDTH                             14028
#define IPOWERPCBCOMP_DIEHEIGHT                            14029
#define IPOWERPCBCOMP_WIREBONDRULESLENGTHMINIMUM           14030
#define IPOWERPCBCOMP_WIREBONDRULESLENGTHMAXIMUM           14031
#define IPOWERPCBCOMP_WIREBONDRULESANGLEMAXIMUM            14032
#define IPOWERPCBCOMP_WIREBONDRULESCLEARANCEWIRETOWIRE     14033
#define IPOWERPCBCOMP_WIREBONDRULESCLEARANCEWIRETOPAD      14034
#define IPOWERPCBCOMP_CBPS                                 14035
#define IPOWERPCBCOMP_SBPS                                 14036
#define IPOWERPCBCOMP_WIREBONDS                            14037
#define IPOWERPCBCOMP_LABELS                               14038
#define IPOWERPCBCOMP_CENTERX                              14039
#define IPOWERPCBCOMP_CENTERY                              14040
// IPowerPCBComp Methods
#define IPOWERPCBCOMP_MOVE                                 14500
#define IPOWERPCBCOMP_MOVECENTER                           14501
#define IPOWERPCBCOMP_ADDLABEL                             14502

// IPowerPCBNet Properties
#define IPOWERPCBNET_NAME                                  0
#define IPOWERPCBNET_APPLICATION                           15001
#define IPOWERPCBNET_PARENT                                15002
#define IPOWERPCBNET_SELECTED                              15003
#define IPOWERPCBNET_ATTRIBUTES                            15004
#define IPOWERPCBNET_NETCLASSATTRIBUTES                    15005
#define IPOWERPCBNET_OBJECTTYPE                            15006
#define IPOWERPCBNET_CONNECTIONS                           15007
#define IPOWERPCBNET_PINS                                  15008
#define IPOWERPCBNET_LENGTH                                15009
#define IPOWERPCBNET_POWER                                 15010
#define IPOWERPCBNET_VIAS                                  15011
#define IPOWERPCBNET_NETCLASS                              15012
#define IPOWERPCBNET_DRAWINGS                              15013
// IPowerPCBNet Methods

// IPowerPCBPin Properties
#define IPOWERPCBPIN_NAME                                  0
#define IPOWERPCBPIN_APPLICATION                           16001
#define IPOWERPCBPIN_PARENT                                16002
#define IPOWERPCBPIN_SELECTED                              16003
#define IPOWERPCBPIN_POSITIONX                             16004
#define IPOWERPCBPIN_POSITIONY                             16005
#define IPOWERPCBPIN_DRILLSIZE                             16006
#define IPOWERPCBPIN_COMPONENT                             16007
#define IPOWERPCBPIN_NET                                   16008
#define IPOWERPCBPIN_ISSMD                                 16009
#define IPOWERPCBPIN_PLATED                                16010
#define IPOWERPCBPIN_TESTPOINT                             16011
#define IPOWERPCBPIN_GLUED                                 16012
#define IPOWERPCBPIN_PLANETHERMAL                          16013
#define IPOWERPCBPIN_FUNCTIONNAME                          16014
#define IPOWERPCBPIN_ATTRIBUTES                            16015
#define IPOWERPCBPIN_OBJECTTYPE                            16016
#define IPOWERPCBPIN_ELECTRICALTYPE                        16017
#define IPOWERPCBPIN_NUMBER                                16018
#define IPOWERPCBPIN_SLOTLENGTH                            16019
#define IPOWERPCBPIN_SLOTOFFSET                            16020
#define IPOWERPCBPIN_SLOTORIENTATION                       16021
// IPowerPCBPin Methods

// IPowerPCBVia Properties
#define IPOWERPCBVIA_NAME                                  0
#define IPOWERPCBVIA_APPLICATION                           17001
#define IPOWERPCBVIA_PARENT                                17002
#define IPOWERPCBVIA_SELECTED                              17003
#define IPOWERPCBVIA_POSITIONX                             17004
#define IPOWERPCBVIA_POSITIONY                             17005
#define IPOWERPCBVIA_DRILLSIZE                             17006
#define IPOWERPCBVIA_NET                                   17007
#define IPOWERPCBVIA_PLATED                                17008
#define IPOWERPCBVIA_TESTPOINT                             17009
#define IPOWERPCBVIA_GLUED                                 17010
#define IPOWERPCBVIA_STARTLAYER                            17011
#define IPOWERPCBVIA_ENDLAYER                              17012
#define IPOWERPCBVIA_PLANETHERMAL                          17013
#define IPOWERPCBVIA_TYPE                                  17014
#define IPOWERPCBVIA_ATTRIBUTES                            17015
#define IPOWERPCBVIA_OBJECTTYPE                            17016
// IPowerPCBVia Methods

// IPowerPCBDocs Properties
// IPowerPCBDocs Methods
// IPowerPCBAttrs Properties
#define IPOWERPCBATTRS_ITEM                                0
#define IPOWERPCBATTRS_APPLICATION                         19001
#define IPOWERPCBATTRS_PARENT                              19002
#define IPOWERPCBATTRS_COUNT                               19003
#define IPOWERPCBATTRS_NEWENUM                             19006
#define IPOWERPCBATTRS_ITEMTYPE                            19007
#define IPOWERPCBATTRS_NEXT                                19008
#define IPOWERPCBATTRS_OBJECTTYPE                          19009
#define IPOWERPCBATTRS_PARENTOBJECT                        19010
// IPowerPCBAttrs Methods
#define IPOWERPCBATTRS_ADD                                 19500
#define IPOWERPCBATTRS_DELETE                              19501
#define IPOWERPCBATTRS_MERGE                               19502
#define IPOWERPCBATTRS_REMOVE                              19503
#define IPOWERPCBATTRS_RESET                               19504
#define IPOWERPCBATTRS_SELECT                              19505
#define IPOWERPCBATTRS_SORT                                19506

// IPowerPCBAttr Properties
#define IPOWERPCBATTR_VALUE                                0
#define IPOWERPCBATTR_APPLICATION                          20001
#define IPOWERPCBATTR_PARENT                               20002
#define IPOWERPCBATTR_NAME                                 20003
#define IPOWERPCBATTR_MEASURE                              20004
#define IPOWERPCBATTR_OBJECTTYPE                           20005
// IPowerPCBAttr Methods

// IPowerPCBConnection  Properties
#define IPOWERPCBCONNECTION_NAME                           0
#define IPOWERPCBCONNECTION_APPLICATION                    21001
#define IPOWERPCBCONNECTION_PARENT                         21002
#define IPOWERPCBCONNECTION_PINS                           21003
#define IPOWERPCBCONNECTION_VIAS                           21004
#define IPOWERPCBCONNECTION_NET                            21005
#define IPOWERPCBCONNECTION_ROUTESEGMENTS                  21006
#define IPOWERPCBCONNECTION_SELECTED                       21007
#define IPOWERPCBCONNECTION_OBJECTTYPE                     21010
#define IPOWERPCBCONNECTION_LENGTH                         21011
// IPowerPCBConnection  Methods

// IPowerPCBRouteSegment  Properties
#define IPOWERPCBROUTESEGMENT_NAME                         0
#define IPOWERPCBROUTESEGMENT_APPLICATION                  22001
#define IPOWERPCBROUTESEGMENT_PARENT                       22002
#define IPOWERPCBROUTESEGMENT_OBJECTTYPE                   22003
#define IPOWERPCBROUTESEGMENT_NET                          22004
#define IPOWERPCBROUTESEGMENT_LAYER                        22005
#define IPOWERPCBROUTESEGMENT_WIDTH                        22006
#define IPOWERPCBROUTESEGMENT_SELECTED                     22007
#define IPOWERPCBROUTESEGMENT_SEGMENTTYPE                  22008
#define IPOWERPCBROUTESEGMENT_LENGTH                       22009
#define IPOWERPCBROUTESEGMENT_POINTS                       22010
// IPowerPCBRouteSegment  Methods

// IPowerPCBAssOpts Properties
#define IPOWERPCBAOS_ITEM                                  0
#define IPOWERPCBAOS_APPLICATION                           23001
#define IPOWERPCBAOS_PARENT                                23002
#define IPOWERPCBAOS_COUNT                                 23003
#define IPOWERPCBAOS_NEWENUM                               23006
#define IPOWERPCBAOS_OBJECTTYPE                            23007
#define IPOWERPCBAOS_PARENTOBJECT                          23008
#define IPOWERPCBAOS_ITEMTYPE                              23009
#define IPOWERPCBAOS_NEXT                                  23010
// IPowerPCBAssOpts Methods
#define IPOWERPCBAOS_ADD                                   23500
#define IPOWERPCBAOS_DELETE                                23501
#define IPOWERPCBAOS_MERGE                                 23502
#define IPOWERPCBAOS_REMOVE                                23503
#define IPOWERPCBAOS_RESET                                 23504
#define IPOWERPCBAOS_SELECT                                23505
#define IPOWERPCBAOS_SORT                                  23506

// IPowerPCBJumper Properties
#define IPOWERPCBJUMPER_NAME                               0
#define IPOWERPCBJUMPER_APPLICATION                        24001
#define IPOWERPCBJUMPER_PARENT                             24002
#define IPOWERPCBJUMPER_OBJECTTYPE                         24003
#define IPOWERPCBJUMPER_ORIENTATION                        24004
#define IPOWERPCBJUMPER_POINTS                             24005
#define IPOWERPCBJUMPER_SELECTED                           24006
#define IPOWERPCBJUMPER_NET                                24007
#define IPOWERPCBJUMPER_LENGTH                             24008
#define IPOWERPCBJUMPER_INSTALLED                          24009
// IPowerPCBJumper Methods

// IPowerPCBPartType Properties
#define IPOWERPCBPARTTYPE_NAME                             0
#define IPOWERPCBPARTTYPE_APPLICATION                      25001
#define IPOWERPCBPARTTYPE_PARENT                           25002
#define IPOWERPCBPARTTYPE_SELECTED                         25003
#define IPOWERPCBPARTTYPE_COMPONENTS                       25004
#define IPOWERPCBPARTTYPE_ATTRIBUTES                       25005
#define IPOWERPCBPARTTYPE_OBJECTTYPE                       25006
#define IPOWERPCBPARTTYPE_LOGIC                            25007
#define IPOWERPCBPARTTYPE_ECOREGISTERED                    25008
// IPowerPCBPartType Methods

// IPowerPCBMeasure Properties
#define IPOWERPCBMEASURE_VALUE                             0
#define IPOWERPCBMEASURE_APPLICATION                       31001
#define IPOWERPCBMEASURE_PARENT                            31002
#define IPOWERPCBMEASURE_NAME                              31003
#define IPOWERPCBMEASURE_TEXT                              31004
#define IPOWERPCBMEASURE_UNIT                              31005
#define IPOWERPCBMEASURE_PREFIX                            31006
#define IPOWERPCBMEASURE_NUMBER                            31007
// IPowerPCBMeasure Properties
#define IPOWERPCBMEASURE_NORMALIZE                         31008
#define IPOWERPCBMEASURE_OBJECTTYPE                        31009

// IPowerPCBAttrTypes Properties
#define IPOWERPCBATTRTYPES_ITEM                            0
#define IPOWERPCBATTRTYPES_APPLICATION                     32001
#define IPOWERPCBATTRTYPES_PARENT                          32002
#define IPOWERPCBATTRTYPES_COUNT                           32003
#define IPOWERPCBATTRTYPES_NEWENUM                         32006
#define IPOWERPCBATTRTYPES_OBJECTTYPE                      32007
#define IPOWERPCBATTRTYPES_PARENTOBJECT                    32008
#define IPOWERPCBATTRTYPES_ITEMTYPE                        32009
#define IPOWERPCBATTRTYPES_NEXT                            32010
// IPowerPCBAttrTypes Methods
#define IPOWERPCBATTRTYPES_ADD                             32500
#define IPOWERPCBATTRTYPES_MERGE                           32502
#define IPOWERPCBATTRTYPES_REMOVE                          32503
#define IPOWERPCBATTRTYPES_RESET                           32504
#define IPOWERPCBATTRTYPES_SELECT                          32505
#define IPOWERPCBATTRTYPES_SORT                            32506

// IPowerPCBAttrType Properties
#define IPOWERPCBATTRTYPE_NAME                             0
#define IPOWERPCBATTRTYPE_APPLICATION                      33001
#define IPOWERPCBATTRTYPE_PARENT                           33002
#define IPOWERPCBATTRTYPE_ISALLOWED                        33003
#define IPOWERPCBATTRTYPE_OBJECTTYPE                       33004
// IPowerPCBAttrType Methods

// IPowerPCBCBP Properties
#define IPOWERPCBCBP_NAME                                  0
#define IPOWERPCBCBP_APPLICATION                           34001
#define IPOWERPCBCBP_PARENT                                34002
#define IPOWERPCBCBP_COMPONENT                             34003
#define IPOWERPCBCBP_SBPS                                  34004
#define IPOWERPCBCBP_WIREBONDS                             34005
#define IPOWERPCBCBP_POSITIONX                             34006
#define IPOWERPCBCBP_POSITIONY                             34007
#define IPOWERPCBCBP_EDGE                                  34008
#define IPOWERPCBCBP_LENGTH                                34009
#define IPOWERPCBCBP_WIDTH                                 34010
#define IPOWERPCBCBP_SHAPE                                 34011
#define IPOWERPCBCBP_FUNCTION                              34012
#define IPOWERPCBCBP_LAYER                                 34013
#define IPOWERPCBCBP_OBJECTTYPE                            34014
// IPowerPCBCBP Methods

// IPowerPCBSBP Properties
#define IPOWERPCBSBP_NAME                                  0
#define IPOWERPCBSBP_APPLICATION                           35001
#define IPOWERPCBSBP_PARENT                                35002
#define IPOWERPCBSBP_COMPONENT                             35003
#define IPOWERPCBSBP_CBPS                                  35004
#define IPOWERPCBSBP_WIREBONDS                             35005
#define IPOWERPCBSBP_POSITIONX                             35006
#define IPOWERPCBSBP_POSITIONY                             35007
#define IPOWERPCBSBP_ORIENTATION                           35008
#define IPOWERPCBSBP_LENGTH                                35009
#define IPOWERPCBSBP_WIDTH                                 35010
#define IPOWERPCBSBP_SHAPE                                 35011
#define IPOWERPCBSBP_FUNCTION                              35012
#define IPOWERPCBSBP_LAYER                                 35013
#define IPOWERPCBSBP_OBJECTTYPE                            35014
#define IPOWERPCBSBP_TIER                                  35015
// IPowerPCBSBP Methods

// IPowerPCBWirebond Properties
#define IPOWERPCBWIREBOND_NAME                             0
#define IPOWERPCBWIREBOND_APPLICATION                      36001
#define IPOWERPCBWIREBOND_PARENT                           36002
#define IPOWERPCBWIREBOND_COMPONENT                        36003
#define IPOWERPCBWIREBOND_STARTPAD                         36004
#define IPOWERPCBWIREBOND_ENDPAD                           36005
#define IPOWERPCBWIREBOND_STARTX                           36006
#define IPOWERPCBWIREBOND_STARTY                           36007
#define IPOWERPCBWIREBOND_ENDX                             36008
#define IPOWERPCBWIREBOND_ENDY                             36009
#define IPOWERPCBWIREBOND_STARTOFFSETX                     36010
#define IPOWERPCBWIREBOND_STARTOFFSETY                     36011
#define IPOWERPCBWIREBOND_ENDOFFSETX                       36012
#define IPOWERPCBWIREBOND_ENDOFFSETY                       36013
#define IPOWERPCBWIREBOND_OBJECTTYPE                       36014
#define IPOWERPCBWIREBOND_ANGLE                            36015
// IPowerPCBWirebond Methods

// IPowerPCBNetClass Properties
#define IPOWERPCBNETCLASS_APPLICATION                      0
#define IPOWERPCBNETCLASS_PARENT                           37001
#define IPOWERPCBNETCLASS_OBJECTTYPE                       37002
#define IPOWERPCBNETCLASS_NAME                             37003
#define IPOWERPCBNETCLASS_NETS                             37004
#define IPOWERPCBNETCLASS_ATTRIBUTES                       37005
// IPowerPCBNetClass Methods

// IPowerPCBPolyline Properties
#define IPOWERPCBPOLYLINE_APPLICATION                      0
#define IPOWERPCBPOLYLINE_PARENT                           38001
#define IPOWERPCBPOLYLINE_OBJECTTYPE                       38002
#define IPOWERPCBPOLYLINE_GEOMETRY                         38003
#define IPOWERPCBPOLYLINE_LINEWIDTH                        38004
#define IPOWERPCBPOLYLINE_OUTLINETYPE                      38005
#define IPOWERPCBPOLYLINE_SHAPETYPE                        38006
#define IPOWERPCBPOLYLINE_LAYER                            38007
#define IPOWERPCBPOLYLINE_POINTS                           38008
#define IPOWERPCBPOLYLINE_CENTERX                          38009
#define IPOWERPCBPOLYLINE_CENTERY                          38010
#define IPOWERPCBPOLYLINE_RADIUS                           38011
// IPowerPCBPolyline Methods

// IPowerPCBCircle Properties
#define IPOWERPCBCIRCLE_APPLICATION                        0
#define IPOWERPCBCIRCLE_PARENT                             39001
#define IPOWERPCBCIRCLE_OBJECTTYPE                         39002
#define IPOWERPCBCIRCLE_GEOMETRY                           39003
#define IPOWERPCBCIRCLE_LINEWIDTH                          39004
#define IPOWERPCBCIRCLE_OUTLINETYPE                        39005
#define IPOWERPCBCIRCLE_SHAPETYPE                          39006
#define IPOWERPCBCIRCLE_RADIUS                             39007
#define IPOWERPCBCIRCLE_CENTERX                            39008
#define IPOWERPCBCIRCLE_CENTERY                            39009
#define IPOWERPCBCIRCLE_LAYER                              39010
// IPowerPCBCircle Methods

// IPowerPCBDrawing Properties
#define IPOWERPCBDRAWING_APPLICATION                       0
#define IPOWERPCBDRAWING_PARENT                            40001
#define IPOWERPCBDRAWING_OBJECTTYPE                        40002
#define IPOWERPCBDRAWING_NAME                              40003
#define IPOWERPCBDRAWING_SELECTED                          40004
#define IPOWERPCBDRAWING_POSITIONX                         40005
#define IPOWERPCBDRAWING_POSITIONY                         40006
#define IPOWERPCBDRAWING_NET                               40007
#define IPOWERPCBDRAWING_TEXTS                             40008
#define IPOWERPCBDRAWING_DRAWINGTYPE                       40009
#define IPOWERPCBDRAWING_GEOMETRY                          40010
// IPowerPCBDrawing Methods

// IPowerPCBText Properties
#define IPOWERPCBTEXT_APPLICATION                          0
#define IPOWERPCBTEXT_PARENT                               41001
#define IPOWERPCBTEXT_OBJECTTYPE                           41002
#define IPOWERPCBTEXT_NAME                                 41003
#define IPOWERPCBTEXT_DRAWING                              41004
#define IPOWERPCBTEXT_SELECTED                             41005
#define IPOWERPCBTEXT_TEXT                                 41006
// IPowerPCBText Methods
#define IPOWERPCBTEXT_DELETE                               41500

// IPowerPCBTextBase Properties (common for IPowerPCBText and IPowerPCBLabel)
#define IPOWERPCBTEXTBASE_LINEWIDTH                        0
#define IPOWERPCBTEXTBASE_POSITIONX                        42001
#define IPOWERPCBTEXTBASE_POSITIONY                        42002
#define IPOWERPCBTEXTBASE_HEIGHT                           42003
#define IPOWERPCBTEXTBASE_MIRROR                           42004
#define IPOWERPCBTEXTBASE_ORIENTATION                      42005
#define IPOWERPCBTEXTBASE_HORIZONTALJUSTIFICATION          42006
#define IPOWERPCBTEXTBASE_VERTICALJUSTIFICATION            42007
#define IPOWERPCBTEXTBASE_LAYER                            42008
// IPowerPCBTextBase Methods

// IPowerPCBLabel Properties
#define IPOWERPCBLABEL_APPLICATION                         0
#define IPOWERPCBLABEL_PARENT                              43001
#define IPOWERPCBLABEL_OBJECTTYPE                          43002
#define IPOWERPCBLABEL_NAME                                43003
#define IPOWERPCBLABEL_TYPE                                43004
#define IPOWERPCBLABEL_DISPLAY                             43005
#define IPOWERPCBLABEL_COMPONENT                           43006
#define IPOWERPCBLABEL_ATTRIBUTE                           43007
#define IPOWERPCBLABEL_RIGHTREADING                        43008
#define IPOWERPCBLABEL_SELECTED                            43009
#define IPOWERPCBLABEL_TEXT                                43010
// IPowerPCBLabel Methods
#define IPOWERPCBLABEL_DELETE                              43500

// IPowerPCBLibrary Properties
#define IPOWERPCBLIBRARY_NAME                              0
#define IPOWERPCBLIBRARY_APPLICATION                       44001
#define IPOWERPCBLIBRARY_PARENT                            44002
#define IPOWERPCBLIBRARY_OBJECTTYPE                        44003
#define IPOWERPCBLIBRARY_PATH                              44004
#define IPOWERPCBLIBRARY_FULLNAME                          44005
// IPowerPCBLibrary Methods
#define IPOWERPCBLIBRARY_IMPORTLIBRARYITEMS                44500
#define IPOWERPCBLIBRARY_GETLIBRARYITEMS                   44502

// IPowerPCBLibraryItem Properties
#define IPOWERPCBLIBRARYITEM_NAME                          0
#define IPOWERPCBLIBRARYITEM_APPLICATION                   45001
#define IPOWERPCBLIBRARYITEM_PARENT                        45002
#define IPOWERPCBLIBRARYITEM_OBJECTTYPE                    45003
#define IPOWERPCBLIBRARYITEM_TYPE                          45004
#define IPOWERPCBLIBRARYITEM_LIBRARY                       45005
// IPowerPCBLibraryItem Methods


// Error Ids
#define POWERPCB_E_UNEXPECTED                     0x80040200
#define POWERPCB_E_OUTOFMEMORY                    0x80040201
#define POWERPCB_E_COLLECTIONINVALIDINDEX         0x80040202
#define POWERPCB_E_COLLECTIONFULL                 0x80040203
#define POWERPCB_E_OBJECTFROMOTHERINSTANCE        0x80040204
#define POWERPCB_E_SINKLISTISFULL                 0x80040205
#define POWERPCB_E_SINKNOTFOUND                   0x80040206
#define POWERPCB_E_INVALIDARGUMENT                0x80040207
#define POWERPCB_E_COMPGLUED                      0x80040208
#define POWERPCB_E_RULESVIOLATION                 0x80040209
#define POWERPCB_E_LOCKEDTESTPOINT                0x8004020A
#define POWERPCB_E_UNIONMEMBERMOVINGOFF           0x8004020B
#define POWERPCB_E_INVALIDLAYERINDEX              0x8004020C
#define POWERPCB_E_LAYERNOTCOMPONENT              0x8004020D
#define POWERPCB_E_INVALIDDECAL                   0x8004020E
#define POWERPCB_E_COULDNOTMODIFYDECAL            0x8004020F
#define POWERPCB_E_INVALIDOBJECT                  0x80040210
#define POWERPCB_E_INCONSISTENTNUDGEMODE          0x80040211
#define POWERPCB_E_INCONSISTENTDRCMODE            0x80040212
#define POWERPCB_E_DRCSETFAILED                   0x80040213
#define POWERPCB_E_ATTRTYPEMISMATCH               0x80040214
#define POWERPCB_E_COLLECTIONINVALIDNAMEINDEX     0x80040215
#define POWERPCB_E_ATTRALREADYEXISTS              0x80040216
#define POWERPCB_E_DELALIENATR                    0x80040217
#define POWERPCB_E_BADATTRNAME                    0x80040218
#define POWERPCB_E_COULDNOTADDATTR                0x80040219
#define POWERPCB_E_NOTASSOPTCOMP                  0x8004021A
#define POWERPCB_E_BADAONAME                      0x8004021B
#define POWERPCB_E_AOALREADYEXISTS                0x8004021C
#define POWERPCB_E_COULDNOTSETTP                  0x8004021D
#define POWERPCB_E_OBJECTFROMOTHERDOC             0x8004021E
#define POWERPCB_E_INVALIDPARTTYPE                0x8004021F
#define POWERPCB_E_SECNOTACQUIRED                 0x80040220
#define POWERPCB_E_SECOVERLIMIT                   0x80040221
#define POWERPCB_E_BADINFILE                      0x80040222
#define POWERPCB_E_BADOUTFILE                     0x80040223
#define POWERPCB_E_MEMBEROFREUSE                  0x80040224
#define POWERPCB_E_COULDNOTPARSEMEASURE           0x80040225
#define POWERPCB_E_UNIT_UNKNOWN                   0x80040226
#define POWERPCB_E_INVALIDDIE                     0x80040227
#define POWERPCB_E_INVALIDHEIGHT                  0x80040228
#define POWERPCB_E_INVALIDLINEWIDTH               0x80040229
#define POWERPCB_E_INVALIDUNIT                    0x8004022B
#define POWERPCB_E_INVALIDORIGIN                  0x8004022C
#define POWERPCB_E_OUTOFRANGE                     0x8004022D
#define POWERPCB_E_NEEDDRCOFF                     0x8004022E
#define POWERPCB_E_NOTLAYOUTEDITOR                0x8004022F
#define POWERPCB_E_CANNOTIMPORTFILE               0x80040230
