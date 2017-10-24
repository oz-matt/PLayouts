/* Copyright Mentor Graphics Corporation 2007

All Rights Reserved.

THIS WORK CONTAINS TRADE SECRET
AND PROPRIETARY INFORMATION WHICH IS THE
PROPERTY OF MENTOR GRAPHICS
CORPORATION OR ITS LICENSORS AND IS
SUBJECT TO LICENSE TERMS. 
*/

// Machine generated IDispatch wrapper class(es) created with ClassWizard
/////////////////////////////////////////////////////////////////////////////
// IPowerLogicApp wrapper class

class IPowerLogicApp : public COleDispatchDriver
{
public:
	IPowerLogicApp() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicApp(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicApp(const IPowerLogicApp& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	CString GetFullName();
	BOOL GetVisible();
	void SetVisible(BOOL bNewValue);
	LPDISPATCH GetActiveDocument();
	CString GetDefaultFilePath();
	void SetDefaultFilePath(LPCTSTR lpszNewValue);
	VARIANT GetPreference(LPCTSTR Name);
	void SetPreference(LPCTSTR Name, const VARIANT& newValue);
	CString GetStatusBarText();
	void SetStatusBarText(LPCTSTR lpszNewValue);
	short GetProgressBar();
	void SetProgressBar(short nNewValue);
	CString GetVersion();
	void Quit();
	LPDISPATCH OpenDocument(LPCTSTR filename);
	void RunMacro(LPCTSTR filename, LPCTSTR macroname);
	void ProcessCommand(long cmdId);
	void ProcessParameter(long parId, const VARIANT& Value);
	void ProcessPointer(long typeID, double x, double y, long unit);
	long StrNumCmp(LPCTSTR str1, LPCTSTR str2);
	void LockServer();
	void UnlockServer();
	LPDISPATCH Measure(const VARIANT& Value, LPCTSTR DefaultUnit);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerLogicDoc wrapper class

class IPowerLogicDoc : public COleDispatchDriver
{
public:
	IPowerLogicDoc() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicDoc(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicDoc(const IPowerLogicDoc& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	CString GetFullName();
	CString GetPath();
	BOOL GetSaved();
	void SetSaved(BOOL bNewValue);
	LPDISPATCH GetActiveView();
	LPDISPATCH GetActiveSheet();
	LPDISPATCH GetSheets();
	LPDISPATCH GetAncestorSheets();
	LPDISPATCH GetComponents();
	LPDISPATCH GetNets();
	LPDISPATCH GetGates();
	void SetGridX(long type, long unit, double newValue);
	double GetGridX(long type, long unit);
	void SetGridY(long type, long unit, double newValue);
	double GetGridY(long type, long unit);
	double GetOriginX(long unit);
	double GetOriginY(long unit);
	LPDISPATCH GetPins();
	LPDISPATCH GetPartTypes();
	void Activate();
	void ExportASCII(LPCTSTR Name, long ver);
	void ExportNetList(LPCTSTR Path, long ver);
	void GenerateECO(LPCTSTR fileToComp, LPCTSTR Path);
	LPDISPATCH GetObjects(long type, LPCTSTR Name, BOOL selected);
	void ImportASCII(LPCTSTR Path);
	void ImportECO(LPCTSTR Path);
	void Save();
	void SaveAs(LPCTSTR Name);
	void SelectObjects(long type, LPCTSTR Name, BOOL Select);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerLogicView wrapper class

class IPowerLogicView : public COleDispatchDriver
{
public:
	IPowerLogicView() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicView(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicView(const IPowerLogicView& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

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
	void SetExtentsToSheet();
	void SetExtents(double TopLeftX, double TopLeftY, double BottomRightX, double BottomRightY, long unit);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerLogicSheet wrapper class

class IPowerLogicSheet : public COleDispatchDriver
{
public:
	IPowerLogicSheet() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicSheet(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicSheet(const IPowerLogicSheet& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	LPDISPATCH GetChildSheets();
	LPDISPATCH GetParentSheet();
	LPDISPATCH GetGates();
	LPDISPATCH GetNets();
	LPDISPATCH GetComponents();
	LPDISPATCH GetPins();
	LPDISPATCH GetPartTypes();
	void Activate();
	LPDISPATCH GetObjects(long type, LPCTSTR Name, BOOL selected);
	LPDISPATCH AddGate(LPCTSTR partTypeName, LPCTSTR refDes, long gateIndex, double PositionX, double PositionY, long unit);
	LPDISPATCH AddComponent(LPCTSTR partTypeName, LPCTSTR refDes, double PositionX, double PositionY, double DeltaX, double DeltaY, long unit);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerLogicSheets wrapper class

class IPowerLogicSheets : public COleDispatchDriver
{
public:
	IPowerLogicSheets() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicSheets(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicSheets(const IPowerLogicSheets& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetItem(const VARIANT& index);
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Add();
	void Delete(const VARIANT& index);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerLogicObjs wrapper class

class IPowerLogicObjs : public COleDispatchDriver
{
public:
	IPowerLogicObjs() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicObjs(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicObjs(const IPowerLogicObjs& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

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
	void Select(BOOL Select);
	void Sort();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerLogicGate wrapper class

class IPowerLogicGate : public COleDispatchDriver
{
public:
	IPowerLogicGate() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicGate(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicGate(const IPowerLogicGate& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	LPDISPATCH GetComponent();
	long GetObjectType();
	LPDISPATCH GetPins();
	BOOL GetSelected();
	void SetSelected(BOOL bNewValue);
	LPDISPATCH GetSheet();
	long GetSwapClass();
	double GetPositionX(long unit);
	double GetPositionY(long unit);
	BOOL GetIsConnector();
	BOOL GetReflectedX();
	void SetReflectedX(BOOL bNewValue);
	BOOL GetReflectedY();
	void SetReflectedY(BOOL bNewValue);
	BOOL GetRotated90();
	void SetRotated90(BOOL bNewValue);
	BOOL GetVisibility(long Item, LPCTSTR attrName);
	void SetVisibility(long Item, LPCTSTR attrName, BOOL bNewValue);
	long GetNumber();
	void Move(double x, double y, long unit);
	void Delete();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerLogicComp wrapper class

class IPowerLogicComp : public COleDispatchDriver
{
public:
	IPowerLogicComp() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicComp(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicComp(const IPowerLogicComp& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	CString GetPartType();
	CString GetPartTypeLogic();
	LPDISPATCH GetAttributes();
	void SetPCBDecal(LPCTSTR lpszNewValue);
	CString GetPCBDecal();
	LPDISPATCH GetPins();
	LPDISPATCH GetGates();
	long GetObjectType();
	BOOL GetSelected();
	void SetSelected(BOOL bNewValue);
	LPDISPATCH GetSheets();
	LPDISPATCH GetUnusedGates();
	LPDISPATCH GetPartTypeObject();
	void Delete();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerLogicAttrs wrapper class

class IPowerLogicAttrs : public COleDispatchDriver
{
public:
	IPowerLogicAttrs() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicAttrs(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicAttrs(const IPowerLogicAttrs& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetItem(const VARIANT& index);
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	long GetCount();
	LPDISPATCH Add(LPCTSTR Name, const VARIANT& Value);
	LPDISPATCH AddMeasure(LPCTSTR Name, double Value, LPCTSTR PrefixUnit);
	void Delete(const VARIANT& varIndex);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerLogicAttr wrapper class

class IPowerLogicAttr : public COleDispatchDriver
{
public:
	IPowerLogicAttr() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicAttr(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicAttr(const IPowerLogicAttr& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	VARIANT GetValue();
	void SetValue(const VARIANT& newValue);
	CString GetName();
	void SetName(LPCTSTR lpszNewValue);
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	LPDISPATCH GetMeasure();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerLogicMeasure wrapper class

class IPowerLogicMeasure : public COleDispatchDriver
{
public:
	IPowerLogicMeasure() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicMeasure(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicMeasure(const IPowerLogicMeasure& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

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
// IPowerLogicPartType wrapper class

class IPowerLogicPartType : public COleDispatchDriver
{
public:
	IPowerLogicPartType() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicPartType(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicPartType(const IPowerLogicPartType& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

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
	long GetObjectType();
	CString GetLogic();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerLogicAttrTypes wrapper class

class IPowerLogicAttrTypes : public COleDispatchDriver
{
public:
	IPowerLogicAttrTypes() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicAttrTypes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicAttrTypes(const IPowerLogicAttrTypes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetItem(const VARIANT& index);
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	long GetCount();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerLogicAttrType wrapper class

class IPowerLogicAttrType : public COleDispatchDriver
{
public:
	IPowerLogicAttrType() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicAttrType(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicAttrType(const IPowerLogicAttrType& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

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
// IPowerLogicNet wrapper class

class IPowerLogicNet : public COleDispatchDriver
{
public:
	IPowerLogicNet() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicNet(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicNet(const IPowerLogicNet& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	double GetWidth(long unit);
	LPDISPATCH GetPins();
	long GetObjectType();
	BOOL GetSelected();
	void SetSelected(BOOL bNewValue);
	LPDISPATCH GetSheets();
	LPDISPATCH GetAttributes();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerLogicPin wrapper class

class IPowerLogicPin : public COleDispatchDriver
{
public:
	IPowerLogicPin() {}		// Calls COleDispatchDriver default constructor
	IPowerLogicPin(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerLogicPin(const IPowerLogicPin& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	CString GetName();
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	CString GetGatePinName();
	long GetNumber();
	LPDISPATCH GetGate();
	LPDISPATCH GetNet();
	long GetObjectType();
	BOOL GetSelected();
	void SetSelected(BOOL bNewValue);
	CString GetFunctionName();
	long GetElectricalType();
	long GetSwapClass();
	double GetPositionX(long unit);
	double GetPositionY(long unit);
	LPDISPATCH GetComponent();
};
