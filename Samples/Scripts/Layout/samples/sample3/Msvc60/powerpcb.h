/* Copyright Mentor Graphics Corporation 2003

    All Rights Reserved.

 THIS WORK CONTAINS TRADE SECRET
 AND PROPRIETARY INFORMATION WHICH IS THE
 PROPERTY OF MENTOR GRAPHICS
 CORPORATION OR ITS LICENSORS AND IS
 SUBJECT TO LICENSE TERMS. 
*/
// Machine generated IDispatch wrapper class(es) created with ClassWizard
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBDoc wrapper class

class IPowerPCBDoc : public COleDispatchDriver
{
public:
	IPowerPCBDoc() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBDoc(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBDoc(const IPowerPCBDoc& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	CString GetPath();
	CString GetFullName();
	LPDISPATCH GetActiveView();
	void SetSaved(BOOL bNewValue);
	BOOL GetSaved();
	void SetUnit(long nNewValue);
	long GetUnit();
	double GetOriginX(long unit);
	double GetOriginY(long unit);
	void SetGridX(long type, long unit, double newValue);
	double GetGridX(long type, long unit);
	void SetGridY(long type, long unit, double newValue);
	double GetGridY(long type, long unit);
	CString GetLayerName(long layer);
	LPDISPATCH GetAssemblyOptions();
	void SetPreference(LPCTSTR Name, const VARIANT& newValue);
	VARIANT GetPreference(LPCTSTR Name);
	LPDISPATCH GetAttributes();
	long GetLayerCount();
	long GetLayerType(long layer);
	LPDISPATCH GetComponents();
	LPDISPATCH GetNets();
	LPDISPATCH GetPins();
	LPDISPATCH GetVias();
	LPDISPATCH GetConnections();
	LPDISPATCH GetRouteSegments();
	LPDISPATCH GetJumpers();
	double GetBoardOutlineSurface(long unit);
	LPDISPATCH GetPartTypes();
	void Activate();
	LPDISPATCH GetObjects(long type, LPCTSTR value, BOOL selected);
	void SelectObjects(long type, LPCTSTR value, BOOL Select);
	void Save();
	void SaveAs(LPCTSTR Name);
	long ImportNetList(LPCTSTR Name);
	long ExportNetList(LPCTSTR Name, long ver);
	long ExportASCII(LPCTSTR Name, long sections, long ver, long expandAttrs);
	long ImportECOFile(LPCTSTR Name);
	long ExportECOFile(LPCTSTR Name);
	long CheckASCII(LPCTSTR Name, LPCTSTR ignoreNet);
	long ExportRules(LPCTSTR Name, long ver);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBApp wrapper class

class IPowerPCBApp : public COleDispatchDriver
{
public:
	IPowerPCBApp() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBApp(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBApp(const IPowerPCBApp& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	CString GetVersion();
	LPDISPATCH GetActiveDocument();
	void SetDefaultFilePath(LPCTSTR lpszNewValue);
	CString GetDefaultFilePath();
	CString GetFullName();
	void SetVisible(BOOL bNewValue);
	BOOL GetVisible();
	void SetPreference(LPCTSTR Name, const VARIANT& newValue);
	VARIANT GetPreference(LPCTSTR Name);
	void SetStatusBarText(LPCTSTR lpszNewValue);
	CString GetStatusBarText();
	short GetProgressBar();
	void SetProgressBar(short nNewValue);
	LPDISPATCH OpenDocument(LPCTSTR filename);
	void Quit();
	void RunMacro(LPCTSTR filename, LPCTSTR macroname);
	void LockServer();
	void UnlockServer();
	void ProcessCommand(long ID);
	void ProcessParameter(long ID, const VARIANT& value);
	void ProcessPointer(long ID, double x, double y, long unit);
	long StrNumCmp(LPCTSTR str1, LPCTSTR str2);
	LPDISPATCH Measure(const VARIANT& value, LPCTSTR DefaultUnit);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBMeasure wrapper class

class IPowerPCBMeasure : public COleDispatchDriver
{
public:
	IPowerPCBMeasure() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBMeasure(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBMeasure(const IPowerPCBMeasure& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	void SetValue(double newValue);
	double GetValue();
	void SetText(LPCTSTR lpszNewValue);
	CString GetText();
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	CString GetUnit(long Format);
	CString GetPrefix(long Format);
	double GetNumber();
	CString Normalize();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBView wrapper class

class IPowerPCBView : public COleDispatchDriver
{
public:
	IPowerPCBView() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBView(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBView(const IPowerPCBView& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	double GetTopLeftX(long unit);
	double GetTopLeftY(long unit);
	double GetBottomRightX(long unit);
	double GetBottomRightY(long unit);
	void Refresh();
	void Pan(double x, double y, long unit);
	void SetExtentsToAll();
	void SetExtentsToBoard();
	void SetExtents(double tlx, double tly, double brx, double bry, long unit);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBAssOpts wrapper class

class IPowerPCBAssOpts : public COleDispatchDriver
{
public:
	IPowerPCBAssOpts() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBAssOpts(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBAssOpts(const IPowerPCBAssOpts& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetItem(const VARIANT& index);
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Add(LPCTSTR Name);
	void Delete(const VARIANT& index);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBAttrs wrapper class

class IPowerPCBAttrs : public COleDispatchDriver
{
public:
	IPowerPCBAttrs() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBAttrs(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBAttrs(const IPowerPCBAttrs& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetItem(const VARIANT& index);
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Add(LPCTSTR Name, const VARIANT& value);
	void Delete(const VARIANT& index);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBAttr wrapper class

class IPowerPCBAttr : public COleDispatchDriver
{
public:
	IPowerPCBAttr() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBAttr(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBAttr(const IPowerPCBAttr& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	void SetValue(const VARIANT& newValue);
	VARIANT GetValue();
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	LPDISPATCH GetMeasure();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBAttrTypes wrapper class

class IPowerPCBAttrTypes : public COleDispatchDriver
{
public:
	IPowerPCBAttrTypes() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBAttrTypes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBAttrTypes(const IPowerPCBAttrTypes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetItem(const VARIANT& index);
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Add(LPCTSTR Name);
	void Delete(const VARIANT& index);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBAttrType wrapper class

class IPowerPCBAttrType : public COleDispatchDriver
{
public:
	IPowerPCBAttrType() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBAttrType(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBAttrType(const IPowerPCBAttrType& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	BOOL GetIsAllowed(long type);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBObjs wrapper class

class IPowerPCBObjs : public COleDispatchDriver
{
public:
	IPowerPCBObjs() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBObjs(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBObjs(const IPowerPCBObjs& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetItem(const VARIANT& index);
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	long GetCount();
	long GetItemType(const VARIANT& index);
	long GetNext(long index, long type);
	void Add(LPDISPATCH object);
	void Merge(LPDISPATCH objects);
	void Remove(const VARIANT& index);
	void Reset();
	void Select(BOOL bSelect);
	void Sort();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBComp wrapper class

class IPowerPCBComp : public COleDispatchDriver
{
public:
	IPowerPCBComp() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBComp(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBComp(const IPowerPCBComp& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	void SetSelected(BOOL bNewValue);
	BOOL GetSelected();
	void SetPartType(LPCTSTR lpszNewValue);
	CString GetPartType();
	void SetDecal(LPCTSTR lpszNewValue);
	CString GetDecal();
	void SetLayer(long nNewValue);
	long GetLayer();
	LPDISPATCH GetPins();
	double GetPositionX(long unit);
	double GetPositionY(long unit);
	void SetOrientation(double newValue);
	double GetOrientation();
	BOOL GetIsSMD();
	void SetGlued(BOOL bNewValue);
	BOOL GetGlued();
	void SetInstalled(BOOL bNewValue);
	BOOL GetInstalled();
	BOOL GetSubstituted();
	VARIANT GetDecalCompatibleList();
	BOOL GetPlaced();
	LPDISPATCH GetAttributes();
	LPDISPATCH GetDecalAttributes();
	LPDISPATCH GetPartTypeAttributes();
	long GetObjectType();
	CString GetPartTypeLogic();
	void SetPartTypeECORegistered(BOOL bNewValue);
	BOOL GetPartTypeECORegistered();
	LPDISPATCH GetPartTypeObject();
	void Move(double x, double y, long unit);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBPartType wrapper class

class IPowerPCBPartType : public COleDispatchDriver
{
public:
	IPowerPCBPartType() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBPartType(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBPartType(const IPowerPCBPartType& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	void SetSelected(BOOL bNewValue);
	BOOL GetSelected();
	LPDISPATCH GetComponents();
	LPDISPATCH GetAttributes();
	long GetObjectType();
	CString GetLogic();
	void SetECORegistered(BOOL bNewValue);
	BOOL GetECORegistered();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBNet wrapper class

class IPowerPCBNet : public COleDispatchDriver
{
public:
	IPowerPCBNet() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBNet(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBNet(const IPowerPCBNet& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	void SetSelected(BOOL bNewValue);
	BOOL GetSelected();
	LPDISPATCH GetAttributes();
	LPDISPATCH GetNetClassAttributes();
	long GetObjectType();
	LPDISPATCH GetConnections();
	double GetLength(BOOL bRouted, long unit);
	BOOL GetPower();
	LPDISPATCH GetPins();
	LPDISPATCH GetVias();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBPin wrapper class

class IPowerPCBPin : public COleDispatchDriver
{
public:
	IPowerPCBPin() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBPin(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBPin(const IPowerPCBPin& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	void SetSelected(BOOL bNewValue);
	BOOL GetSelected();
	double GetPositionX(long unit);
	double GetPositionY(long unit);
	double GetDrillSize(long unit);
	LPDISPATCH GetComponent();
	LPDISPATCH GetNet();
	BOOL GetIsSMD();
	BOOL GetPlated();
	void SetTestPoint(long nNewValue);
	long GetTestPoint();
	void SetGlued(BOOL bNewValue);
	BOOL GetGlued();
	BOOL GetPlaneThermal();
	CString GetFunctionName();
	LPDISPATCH GetAttributes();
	long GetObjectType();
	long GetElectricalType();
	long GetNumber();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBVia wrapper class

class IPowerPCBVia : public COleDispatchDriver
{
public:
	IPowerPCBVia() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBVia(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBVia(const IPowerPCBVia& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	void SetSelected(BOOL bNewValue);
	BOOL GetSelected();
	double GetPositionX(long unit);
	double GetPositionY(long unit);
	double GetDrillSize(long unit);
	LPDISPATCH GetNet();
	BOOL GetPlated();
	void SetTestPoint(long nNewValue);
	long GetTestPoint();
	void SetGlued(BOOL bNewValue);
	BOOL GetGlued();
	long GetStartLayer();
	long GetEndLayer();
	BOOL GetPlaneThermal();
	CString GetType();
	LPDISPATCH GetAttributes();
	long GetObjectType();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBConnection wrapper class

class IPowerPCBConnection : public COleDispatchDriver
{
public:
	IPowerPCBConnection() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBConnection(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBConnection(const IPowerPCBConnection& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	LPDISPATCH GetPins();
	LPDISPATCH GetVias();
	LPDISPATCH GetNet();
	LPDISPATCH GetRouteSegments();
	BOOL GetSelected();
	void SetSelected(BOOL bNewValue);
	long GetObjectType();
	double GetLength(long unit);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBRouteSegment wrapper class

class IPowerPCBRouteSegment : public COleDispatchDriver
{
public:
	IPowerPCBRouteSegment() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBRouteSegment(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBRouteSegment(const IPowerPCBRouteSegment& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	long GetObjectType();
	LPDISPATCH GetNet();
	long GetLayer();
	double GetWidth(long unit);
	void SetSelected(BOOL bNewValue);
	BOOL GetSelected();
	long GetSegmentType();
	double GetLength(long unit);
	VARIANT GetPoints(long unit);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBJumper wrapper class

class IPowerPCBJumper : public COleDispatchDriver
{
public:
	IPowerPCBJumper() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBJumper(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBJumper(const IPowerPCBJumper& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	long GetObjectType();
	double GetOrientation();
	LPDISPATCH GetPoints();
	BOOL GetSelected();
	void SetSelected(BOOL bNewValue);
	LPDISPATCH GetNet();
	double GetLength(long unit);
	void SetInstalled(BOOL bNewValue);
	BOOL GetInstalled();
};
