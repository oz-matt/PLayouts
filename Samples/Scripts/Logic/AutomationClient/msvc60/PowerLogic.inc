// this ALWAYS GENERATED file contains the definitions for the Automation interfaces

// Enums

typedef enum
    {	plogObjectTypeUnknown	= 0,
	plogObjectTypeComponent	= 1,
	plogObjectTypeNet	= 2,
	plogObjectTypePin	= 3,
	plogObjectTypeGate	= 4,
	plogObjectTypePartType	= 8,
	plogObjectTypeAll	= 9999
    }	PlogObjectType;

typedef enum
    {	plogUnitCurrent	= 0,
	plogUnitDatabase	= 1,
	plogUnitMils	= 2,
	plogUnitInch	= 3,
	plogUnitMetric	= 4,
	plogUnitDrawArea	= 99
    }	PlogUnit;

typedef enum
    {	plogGridNone	= 0,
	plogGridDesign	= 1,
	plogGridDisplay	= 3,
	plogGridAll	= 9999
    }	PlogGridType;

typedef enum
    {	plogASCIIVerCurrent	= 0,
	plogASCIIVer1_2	= 1,
	plogASCIIVer3_0	= 2
    }	PlogASCIIVersion;

typedef enum
    {	plogPowerPCBNetListVerCurrent	= 0,
	plogPowerPCBNetListVer2_1	= 2,
	plogPowerPCBNetListVer3_0	= 3
    }	PlogNetListVersion;

typedef enum
    {	plogElectricalTypeUnknown	= 0,
	plogElectricalTypeSource	= 1,
	plogElectricalTypeBidirectional	= 2,
	plogElectricalTypeOpenCollector	= 3,
	plogElectricalTypeOrTieableSource	= 4,
	plogElectricalTypeTristate	= 5,
	plogElectricalTypeLoad	= 6,
	plogElectricalTypeTerminator	= 7,
	plogElectricalTypePower	= 8,
	plogElectricalTypeGround	= 9
    }	PlogPinElectricalType;

typedef enum
    {	plogAttrVisibility	= 0,
	plogAttrNameVisibility	= 1,
	plogRefDesVisibility	= 2,
	plogPartTypeVisibility	= 3,
	plogPinNumberVisibility	= 4,
	plogPinNameVisibility	= 5,
	plogPCBDecalVisibility	= 6,
	plogPCBDecalNameVisibility	= 7
    }	PlogGateVisibility;

typedef enum
    {	plogDefaultPositionX	= 0x80000000,
	plogDefaultPositionY	= 0x80000000
    }	PlogDefaultPosition;

typedef enum
    {	plogMeasureFormatStandard	= 0,
	plogMeasureFormatCurrent	= 1,
	plogMeasureFormatShort	= 2,
	plogMeasureFormatLong	= 3
    }	PlogMeasureFormat;


// IIDs/CLSIDs
const IID LIBID_PowerLogic = {0x91D6BD00,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicApp = {0x91D6BD02,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicDoc = {0x91D6BD04,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicView = {0x91D6BD06,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicSheet = {0x91D6BD0D,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicSheets = {0x91D6BD0C,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicComp = {0x91D6BD09,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicPartType = {0x91D6BD11,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicGate = {0x91D6BD10,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicNet = {0x91D6BD0A,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicPin = {0x91D6BD0B,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicObjs = {0x91D6BD08,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicAttr = {0x91D6BD0F,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicAttrs = {0x91D6BD0E,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicAttrTypes = {0x91D6BD2D,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicAttrType = {0x91D6BD2F,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID IID_IPowerLogicMeasure = {0x91D6BD2B,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID DIID__PowerLogicAppEvents = {0x91D6BD22,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicApplication = {0x91D6BD01,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicSheets = {0x91D6BD1C,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicSheet = {0x91D6BD1D,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID DIID__PowerLogicDocEvents = {0x91D6BD24,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicDocument = {0x91D6BD14,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const IID DIID__PowerLogicViewEvents = {0x91D6BD26,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicView = {0x91D6BD16,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicComponent = {0x91D6BD19,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicPartType = {0x91D6BD21,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicObjects = {0x91D6BD18,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicAttribute = {0x91D6BD1F,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicAttributes = {0x91D6BD1E,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicAttributeTypes = {0x91D6BD2E,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicAttributeType = {0x91D6BD30,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicMeasure = {0x91D6BD2C,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicNet = {0x91D6BD1A,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicPin = {0x91D6BD1B,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};
const CLSID CLSID_PowerLogicGate = {0x91D6BD20,0x246D,0x11d2,{0x86,0x43,0x3A,0xE4,0xEC,0x00,0x00,0x00}};

// Dispatch Ids for methods/properties/events

// IPowerLogicApp Properties
#define IPOWERLOGICAPP_NAME                                0
#define IPOWERLOGICAPP_APPLICATION                         10001
#define IPOWERLOGICAPP_PARENT                              10002
#define IPOWERLOGICAPP_FULLNAME                            10003
#define IPOWERLOGICAPP_VISIBLE                             10004
#define IPOWERLOGICAPP_ACTIVEDOCUMENT                      10005
#define IPOWERLOGICAPP_DEFAULTFILEPATH                     10006
#define IPOWERLOGICAPP_INTERACTIVEMODE                     10007
#define IPOWERLOGICAPP_PREFERENCE                          10008
#define IPOWERLOGICAPP_STATUSBARTEXT                       10009
#define IPOWERLOGICAPP_VERSION                             10010
#define IPOWERLOGICAPP_PROGRESSBAR                         10011
#define IPOWERLOGICAPP_MEASURE                             10012
// IPowerLogicApp Methods
#define IPOWERLOGICAPP_QUIT                                10500
#define IPOWERLOGICAPP_OPEN                                10501
#define IPOWERLOGICAPP_RUNMACRO                            10502
#define IPOWERLOGICAPP_PROCESSCOMMAND                      10503
#define IPOWERLOGICAPP_PROCESSPARAMETER                    10504
#define IPOWERLOGICAPP_PROCESSPOINTER                      10505
#define IPOWERLOGICAPP_STRNUMCMP                           10506
#define IPOWERLOGICAPP_LOCKSERVER                          10507
#define IPOWERLOGICAPP_UNLOCKSERVER                        10508
// IPowerLogicApp Events
#define IPOWERLOGICAPP_OPENDOCUMENTEVENT                   10700
#define IPOWERLOGICAPP_QUITEVENT                           10701
#define IPOWERLOGICAPP_PROGRESSEVENT                       10702

// IPowerLogicDoc Properties
#define IPOWERLOGICDOC_NAME                                0
#define IPOWERLOGICDOC_APPLICATION                         11001
#define IPOWERLOGICDOC_PARENT                              11002
#define IPOWERLOGICDOC_FULLNAME                            11003
#define IPOWERLOGICDOC_PATH                                11004
#define IPOWERLOGICDOC_SAVED                               11005
#define IPOWERLOGICDOC_ACTIVEVIEW                          11006
#define IPOWERLOGICDOC_ACTIVESHEET                         11007
#define IPOWERLOGICDOC_SHEETS                              11008

#define IPOWERLOGICDOC_ANCESTORSHEETS                      11010
#define IPOWERLOGICDOC_COMPONENTS                          11011
#define IPOWERLOGICDOC_NETS                                11012
#define IPOWERLOGICDOC_GATES                               11013
#define IPOWERLOGICDOC_GRIDX                               11014
#define IPOWERLOGICDOC_GRIDY                               11015
#define IPOWERLOGICDOC_ORIGINX                             11016
#define IPOWERLOGICDOC_ORIGINY                             11017
#define IPOWERLOGICDOC_PINS                                11018
#define IPOWERLOGICDOC_PARTTYPES                           11019
#define IPOWERLOGICDOC_ATTRIBUTELIST                       11020
#define IPOWERLOGICDOC_ATTRIBUTETYPES                      11021
// IPowerLogicDoc Methods
#define IPOWERLOGICDOC_ACTIVATE                            11500
#define IPOWERLOGICDOC_EXPORTASCII                         11501
#define IPOWERLOGICDOC_EXPORTNETLIST                       11502
#define IPOWERLOGICDOC_GENERATEECO                         11503
#define IPOWERLOGICDOC_GETOBJECTS                          11504
#define IPOWERLOGICDOC_IMPORTASCII                         11505
#define IPOWERLOGICDOC_IMPORTECO                           11506
#define IPOWERLOGICDOC_SAVE                                11507
#define IPOWERLOGICDOC_SELECTOBJECTS                       11508
#define IPOWERLOGICDOC_SAVEAS                              11509
// IPowerLogicDoc Events
#define IPOWERLOGICDOC_SAVEEVENT                           11700
#define IPOWERLOGICDOC_SELECTIONCHANGEEVENT                11701

// IPowerLogicView Properties
#define IPOWERLOGICVIEW_NAME                               0
#define IPOWERLOGICVIEW_APPLICATION                        12001
#define IPOWERLOGICVIEW_PARENT                             12002
#define IPOWERLOGICVIEW_TOPLEFTX                           12003
#define IPOWERLOGICVIEW_TOPLEFTY                           12004
#define IPOWERLOGICVIEW_BOTTOMRIGHTX                       12005
#define IPOWERLOGICVIEW_BOTTOMRIGHTY                       12006
// IPowerLogicView Methods
#define IPOWERLOGICVIEW_REFRESH                            12500
#define IPOWERLOGICVIEW_PAN                                12501
#define IPOWERLOGICVIEW_SETEXTENTSTOALL                    12502
#define IPOWERLOGICVIEW_SETEXTENTSTOSHEET                  12503
#define IPOWERLOGICVIEW_SETEXTENTS                         12504
// IPowerLogicView Events
#define IPOWERLOGICVIEW_CHANGEEVENT                        12700

// IPowerLogicSheet Properties
#define IPOWERLOGICSHEET_NAME                              0
#define IPOWERLOGICSHEET_APPLICATION                       13001
#define IPOWERLOGICSHEET_PARENT                            13002
#define IPOWERLOGICSHEET_CHILDSHEETS                       13003
#define IPOWERLOGICSHEET_PARENTSHEET                       13004
#define IPOWERLOGICSHEET_GATES                             13005
#define IPOWERLOGICSHEET_NETS                              13006
#define IPOWERLOGICSHEET_COMPONENTS                        13007
#define IPOWERLOGICSHEET_PINS                              13008
#define IPOWERLOGICSHEET_PARTTYPES                         13009
// IPowerLogicSheet Methods
#define IPOWERLOGICSHEET_ACTIVATE                          13500
#define IPOWERLOGICSHEET_GETOBJECTS                        13501
#define IPOWERLOGICSHEET_ADDGATE                           13502
#define IPOWERLOGICSHEET_ADDCOMPONENT                      13503

// IPowerLogicSheets Properties
#define IPOWERLOGICSHEETS_ITEM                             0
#define IPOWERLOGICSHEETS_APPLICATION                      14001
#define IPOWERLOGICSHEETS_PARENT                           14002
#define IPOWERLOGICSHEETS_COUNT                            14003
#define IPOWERLOGICSHEETS_NEWENUM                          14004
// IPowerLogicSheets Methods
#define IPOWERLOGICSHEETS_ADD                              14500
#define IPOWERLOGICSHEETS_DELETE                           14501

// IPowerLogicComp Properties
#define IPOWERLOGICCOMP_NAME                               0
#define IPOWERLOGICCOMP_APPLICATION                        15001
#define IPOWERLOGICCOMP_PARENT                             15002
#define IPOWERLOGICCOMP_PARTTYPE                           15003
#define IPOWERLOGICCOMP_ATTRIBUTES                         15004
#define IPOWERLOGICCOMP_PCBDECAL                           15005
#define IPOWERLOGICCOMP_PINS                               15006
#define IPOWERLOGICCOMP_GATES                              15007
#define IPOWERLOGICCOMP_OBJECTTYPE                         15008
#define IPOWERLOGICCOMP_SELECTED                           15009
#define IPOWERLOGICCOMP_SHEETS                             15010
#define IPOWERLOGICCOMP_UNUSEDGATES                        15012
#define IPOWERLOGICCOMP_PARTTYPELOGIC                      15013
#define IPOWERLOGICCOMP_PARTTYPEOBJECT                     15014
// IPowerLogicComp Methods
#define IPOWERLOGICCOMP_DELETE                             15501

// IPowerLogicGate Properties
#define IPOWERLOGICGATE_NAME                               0
#define IPOWERLOGICGATE_APPLICATION                        16001
#define IPOWERLOGICGATE_PARENT                             16002
#define IPOWERLOGICGATE_COMPONENT                          16003
#define IPOWERLOGICGATE_OBJECTTYPE                         16004
#define IPOWERLOGICGATE_PINS                               16005
#define IPOWERLOGICGATE_SELECTED                           16006
#define IPOWERLOGICGATE_SHEET                              16007
#define IPOWERLOGICGATE_SWAPCLASS                          16008
#define IPOWERLOGICGATE_POSITIONX                          16009
#define IPOWERLOGICGATE_POSITIONY                          16010
#define IPOWERLOGICGATE_ORIENTATION                        16011
#define IPOWERLOGICGATE_ISCONNECTOR                        16012
#define IPOWERLOGICGATE_REFLECTEDX                         16013
#define IPOWERLOGICGATE_REFLECTEDY                         16014
#define IPOWERLOGICGATE_ROTATED90                          16015
#define IPOWERLOGICGATE_VISIBILITY                         16016
#define IPOWERLOGICGATE_NUMBER                             16017
// IPowerLogicGate Methods
#define IPOWERLOGICGATE_MOVE                               16501
#define IPOWERLOGICGATE_DELETE                             16502

// IPowerLogicNet Properties
#define IPOWERLOGICNET_NAME                                0
#define IPOWERLOGICNET_APPLICATION                         17001
#define IPOWERLOGICNET_PARENT                              17002
#define IPOWERLOGICNET_WIDTH                               17003
#define IPOWERLOGICNET_PINS                                17004
#define IPOWERLOGICNET_OBJECTTYPE                          17005
#define IPOWERLOGICNET_SELECTED                            17006
#define IPOWERLOGICNET_SHEETS                              17007
#define IPOWERLOGICNET_ATTRIBUTES                          17008
// IPowerLogicNet Methods

// IPowerLogicPin Properties
#define IPOWERLOGICPIN_NAME                                0
#define IPOWERLOGICPIN_APPLICATION                         18001
#define IPOWERLOGICPIN_PARENT                              18002
#define IPOWERLOGICPIN_GATEPINNAME                         18003
#define IPOWERLOGICPIN_NUMBER                              18004
#define IPOWERLOGICPIN_GATE                                18005
#define IPOWERLOGICPIN_NET                                 18006
#define IPOWERLOGICPIN_OBJECTTYPE                          18007
#define IPOWERLOGICPIN_SELECTED                            18008
#define IPOWERLOGICPIN_FUNCTIONNAME                        18009
#define IPOWERLOGICPIN_ELECTRICALTYPE                      18010
#define IPOWERLOGICPIN_SWAPCLASS                           18011
#define IPOWERLOGICPIN_POSITIONX                           18012
#define IPOWERLOGICPIN_POSITIONY                           18013
#define IPOWERLOGICPIN_COMPONENT                           18014
// IPowerLogicPin Methods

// IPowerLogicObjs Properties
#define IPOWERLOGICOBJS_NAME                               0
#define IPOWERLOGICOBJS_APPLICATION                        19001
#define IPOWERLOGICOBJS_PARENT                             19002
#define IPOWERLOGICOBJS_COUNT                              19003
#define IPOWERLOGICOBJS_NEWENUM                            19004
#define IPOWERLOGICOBJS_ITEMTYPE                           19005
#define IPOWERLOGICOBJS_NEXT                               19006
// IPowerLogicObjs Methods
#define IPOWERLOGICOBJS_ADD                                19500
#define IPOWERLOGICOBJS_MERGE                              19501
#define IPOWERLOGICOBJS_REMOVE                             19502
#define IPOWERLOGICOBJS_RESET                              19503
#define IPOWERLOGICOBJS_SELECT                             19504
#define IPOWERLOGICOBJS_SORT                               19505

// IPowerLogicAttr Properties
#define IPOWERLOGICATTR_VALUE                              0
#define IPOWERLOGICATTR_NAME                               20001
#define IPOWERLOGICATTR_APPLICATION                        20002
#define IPOWERLOGICATTR_PARENT                             20003
#define IPOWERLOGICATTR_MEASURE                            20004
// IPowerLogicAttr Methods

// IPowerLogicAttrs Properties
#define IPOWERLOGICATTRS_ITEM                              0
#define IPOWERLOGICATTRS_APPLICATION                       21001
#define IPOWERLOGICATTRS_PARENT                            21002
#define IPOWERLOGICATTRS_COUNT                             21003
#define IPOWERLOGICATTRS_NEWENUM                           21004
// IPowerLogicAttrs Methods
#define IPOWERLOGICATTRS_ADD                               21500
#define IPOWERLOGICATTRS_DELETE                            21501
#define IPOWERLOGICATTRS_ADDMEASURE                        21502

// IPowerLogicPartType Properties
#define IPOWERLOGICPARTTYPE_NAME                           0
#define IPOWERLOGICPARTTYPE_APPLICATION                    25001
#define IPOWERLOGICPARTTYPE_PARENT                         25002
#define IPOWERLOGICPARTTYPE_SELECTED                       25003
#define IPOWERLOGICPARTTYPE_COMPONENTS                     25004
#define IPOWERLOGICPARTTYPE_OBJECTTYPE                     25005
#define IPOWERLOGICPARTTYPE_LOGIC                          25006
// IPowerLogicPartType Methods

// IPowerLogicMeasure Properties
#define IPOWERLOGICMEASURE_VALUE                           0
#define IPOWERLOGICMEASURE_APPLICATION                     31001
#define IPOWERLOGICMEASURE_PARENT                          31002
#define IPOWERLOGICMEASURE_NAME                            31003
#define IPOWERLOGICMEASURE_TEXT                            31004
#define IPOWERLOGICMEASURE_UNIT                            31005
#define IPOWERLOGICMEASURE_PREFIX                          31006
#define IPOWERLOGICMEASURE_NUMBER                          31007
// IPowerLogicMeasure Properties
#define IPOWERLOGICMEASURE_NORMALIZE                       31008

// IPowerLogicAttrTypes Properties
#define IPOWERLOGICATTRTYPES_ITEM                          0
#define IPOWERLOGICATTRTYPES_APPLICATION                   32001
#define IPOWERLOGICATTRTYPES_PARENT                        32002
#define IPOWERLOGICATTRTYPES_COUNT                         32003
#define IPOWERLOGICATTRTYPES_NEWENUM                       32006
// IPowerLogicAttrTypes Methods
#define IPOWERLOGICATTRTYPES_ADD                           32500
#define IPOWERLOGICATTRTYPES_DELETE                        32501

// IPowerLogicAttrType Properties
#define IPOWERLOGICATTRTYPE_NAME                           0
#define IPOWERLOGICATTRTYPE_APPLICATION                    33001
#define IPOWERLOGICATTRTYPE_PARENT                         33002
#define IPOWERLOGICATTRTYPE_ISALLOWED                      33003
// IPowerLogicAttrType Methods

// Error Ids
#define POWERLOGIC_E_UNEXPECTED                     0x80040200
#define POWERLOGIC_E_OUTOFMEMORY                    0x80040201
#define POWERLOGIC_E_COLLECTIONINVALIDINDEX         0x80040202
#define POWERLOGIC_E_COLLECTIONFULL                 0x80040203
#define POWERLOGIC_E_OBJECTFROMOTHERINSTANCE        0x80040204
#define POWERLOGIC_E_OBJECTHASNOPOSITION            0x80040205
#define POWERLOGIC_E_OBJECTISUNUSED                 0x80040206
#define POWERLOGIC_E_INVALIDARGUMENT                0x80040207
#define POWERLOGIC_E_INVALIDCOORDINATES             0x80040208
#define POWERLOGIC_E_DIFFERENTPACKAGE               0x80040209
#define POWERLOGIC_E_INVALIDLAYERINDEX              0x8004020C
#define POWERLOGIC_E_INVALIDOBJECT                  0x80040210
#define POWERLOGIC_E_ATTRTYPEMISMATCH               0x80040211
#define POWERLOGIC_E_COULDNOTSETVALUE               0x80040212
#define POWERLOGIC_E_COLLECTIONINVALIDNAMEINDEX     0x80040215
#define POWERLOGIC_E_OBJALREADYEXISTS               0x80040216
#define POWERLOGIC_E_BADOBJNAME                     0x80040218
#define POWERLOGIC_E_COULDNOTADDOBJ                 0x80040219
#define POWERLOGIC_E_COULDNOTDELOBJ                 0x8004021A
#define POWERLOGIC_E_OBJECTFROMOTHERDOC             0x8004021E
#define POWERLOGIC_E_COULDNOTRENAMEOBJ              0x8004021F
#define POWERLOGIC_E_NOMOREOBJSALLOWED              0x80040220
