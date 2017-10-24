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
	void SetTopLeftX(long unit, double newValue);
	double GetTopLeftY(long unit);
	void SetTopLeftY(long unit, double newValue);
	double GetBottomRightX(long unit);
	void SetBottomRightX(long unit, double newValue);
	double GetBottomRightY(long unit);
	void SetBottomRightY(long unit, double newValue);
	void Refresh();
	void Pan(double x, double y, long unit);
	void SetExtentsToAll();
	void SetExtentsToBoard();
	void SetExtents(double tlx, double tly, double brx, double bry, long unit);
	double GetCenterX(long unit);
	double GetCenterY(long unit);
	double GetZoom();
	long GetObjectType();
	void SetScale(double Zoom, double CenterX, double CenterY, long unit);
	void ShowRectangle(double left, double top, double right, double bottom, long unit);
	void ZoomToSelection();
};
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
	LPDISPATCH GetPartTypes();
	long GetObjectType();
	LPDISPATCH GetNetClasses();
	LPDISPATCH GetDrawings();
	LPDISPATCH GetTexts();
	LPDISPATCH AddText(LPCTSTR Text, long layer, double posX, double posY, double Height, double LineWidth, long unit, double Orientation, BOOL Mirror);
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
	LPDISPATCH Measure(const VARIANT& value, LPCTSTR DefaultUnit);
	long GetObjectType();
	LPDISPATCH GetLibraries();
	LPDISPATCH GetLibraryItems(long type, LPCTSTR Name);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBText wrapper class

class IPowerPCBText : public COleDispatchDriver
{
public:
	IPowerPCBText() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBText(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBText(const IPowerPCBText& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	double GetLineWidth(long unit);
	void SetLineWidth(long unit, double newValue);
	double GetPositionX(long unit, long origin);
	void SetPositionX(long unit, long origin, double newValue);
	double GetPositionY(long unit, long origin);
	void SetPositionY(long unit, long origin, double newValue);
	double GetHeight(long unit);
	void SetHeight(long unit, double newValue);
	BOOL GetMirror(long origin);
	void SetMirror(long origin, BOOL bNewValue);
	double GetOrientation(long origin);
	void SetOrientation(long origin, double newValue);
	long GetHorzJustification();
	void SetHorzJustification(long nNewValue);
	long GetVertJustification();
	void SetVertJustification(long nNewValue);
	long GetLayer();
	void SetLayer(long nNewValue);
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	long GetObjectType();
	CString GetName();
	LPDISPATCH GetDrawing();
	BOOL GetSelected();
	void SetSelected(BOOL bNewValue);
	CString GetText();
	void SetText(LPCTSTR lpszNewValue);
	BOOL Delete();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBLabel wrapper class

class IPowerPCBLabel : public COleDispatchDriver
{
public:
	IPowerPCBLabel() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBLabel(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBLabel(const IPowerPCBLabel& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	double GetLineWidth(long unit);
	void SetLineWidth(long unit, double newValue);
	double GetPositionX(long unit, long origin);
	void SetPositionX(long unit, long origin, double newValue);
	double GetPositionY(long unit, long origin);
	void SetPositionY(long unit, long origin, double newValue);
	double GetHeight(long unit);
	void SetHeight(long unit, double newValue);
	BOOL GetMirror(long origin);
	void SetMirror(long origin, BOOL bNewValue);
	double GetOrientation(long origin);
	void SetOrientation(long origin, double newValue);
	long GetHorzJustification();
	void SetHorzJustification(long nNewValue);
	long GetVertJustification();
	void SetVertJustification(long nNewValue);
	long GetLayer();
	void SetLayer(long nNewValue);
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	long GetObjectType();
	CString GetName();
	long GetLabelType();
	long GetDisplay();
	void SetDisplay(long nNewValue);
	LPDISPATCH GetComponent();
	LPDISPATCH GetAttribute();
	long GetRightReading();
	void SetRightReading(long nNewValue);
	BOOL GetSelected();
	void SetSelected(BOOL bNewValue);
	CString GetText();
	BOOL Delete();
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
// IPowerPCBPolyline wrapper class

class IPowerPCBPolyline : public COleDispatchDriver
{
public:
	IPowerPCBPolyline() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBPolyline(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBPolyline(const IPowerPCBPolyline& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	long GetObjectType();
	LPDISPATCH GetGeometry();
	double GetLineWidth(long unit);
	long GetOutlineType();
	long GetShapeType();
	long GetLayer();
	VARIANT GetPoints(long unit, long origin);
	double GetCenterX(long corner, long unit, long origin);
	double GetCenterY(long corner, long unit, long origin);
	double GetRadius(long corner, long unit);
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBCircle wrapper class

class IPowerPCBCircle : public COleDispatchDriver
{
public:
	IPowerPCBCircle() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBCircle(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBCircle(const IPowerPCBCircle& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	long GetObjectType();
	LPDISPATCH GetGeometry();
	double GetLineWidth(long unit);
	long GetOutlineType();
	long GetShapeType();
	double GetRadius(long unit);
	double GetCenterX(long unit, long origin);
	double GetCenterY(long unit, long origin);
	long GetLayer();
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
	long GetObjectType();
	LPDISPATCH GetParentObject();
};
/////////////////////////////////////////////////////////////////////////////
// IPowerPCBDrawing wrapper class

class IPowerPCBDrawing : public COleDispatchDriver
{
public:
	IPowerPCBDrawing() {}		// Calls COleDispatchDriver default constructor
	IPowerPCBDrawing(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	IPowerPCBDrawing(const IPowerPCBDrawing& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

// Attributes
public:

// Operations
public:
	LPDISPATCH GetApplication();
	LPDISPATCH GetParent();
	long GetObjectType();
	CString GetName();
	BOOL GetSelected();
	void SetSelected(BOOL bNewValue);
	double GetPositionX(long unit, long origin);
	double GetPositionY(long unit, long origin);
	LPDISPATCH GetNet();
	LPDISPATCH GetTexts();
	long GetDrawingType();
	LPDISPATCH GetGeometry();
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
