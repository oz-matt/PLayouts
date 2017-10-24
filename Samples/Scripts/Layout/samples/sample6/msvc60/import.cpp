/* Copyright Mentor Graphics Corporation 2003

    All Rights Reserved.

 THIS WORK CONTAINS TRADE SECRET
 AND PROPRIETARY INFORMATION WHICH IS THE
 PROPERTY OF MENTOR GRAPHICS
 CORPORATION OR ITS LICENSORS AND IS
 SUBJECT TO LICENSE TERMS. 
*/
// Machine generated IDispatch wrapper class(es) created with ClassWizard

#include "stdafx.h"
#include "import.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif



/////////////////////////////////////////////////////////////////////////////
// IPowerPCBView properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBView operations

CString IPowerPCBView::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBView::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x2ee1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBView::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x2ee2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

double IPowerPCBView::GetTopLeftX(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2ee3, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

void IPowerPCBView::SetTopLeftX(long unit, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_R8;
	InvokeHelper(0x2ee3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 unit, newValue);
}

double IPowerPCBView::GetTopLeftY(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2ee4, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

void IPowerPCBView::SetTopLeftY(long unit, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_R8;
	InvokeHelper(0x2ee4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 unit, newValue);
}

double IPowerPCBView::GetBottomRightX(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2ee5, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

void IPowerPCBView::SetBottomRightX(long unit, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_R8;
	InvokeHelper(0x2ee5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 unit, newValue);
}

double IPowerPCBView::GetBottomRightY(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2ee6, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

void IPowerPCBView::SetBottomRightY(long unit, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_R8;
	InvokeHelper(0x2ee6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 unit, newValue);
}

void IPowerPCBView::Refresh()
{
	InvokeHelper(0x30d4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void IPowerPCBView::Pan(double x, double y, long unit)
{
	static BYTE parms[] =
		VTS_R8 VTS_R8 VTS_I4;
	InvokeHelper(0x30d5, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 x, y, unit);
}

void IPowerPCBView::SetExtentsToAll()
{
	InvokeHelper(0x30d6, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void IPowerPCBView::SetExtentsToBoard()
{
	InvokeHelper(0x30d7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void IPowerPCBView::SetExtents(double tlx, double tly, double brx, double bry, long unit)
{
	static BYTE parms[] =
		VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_I4;
	InvokeHelper(0x30d8, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 tlx, tly, brx, bry, unit);
}

double IPowerPCBView::GetCenterX(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2ee7, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

double IPowerPCBView::GetCenterY(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2ee8, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

double IPowerPCBView::GetZoom()
{
	double result;
	InvokeHelper(0x2ee9, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

long IPowerPCBView::GetObjectType()
{
	long result;
	InvokeHelper(0x2eea, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void IPowerPCBView::SetScale(double Zoom, double CenterX, double CenterY, long unit)
{
	static BYTE parms[] =
		VTS_R8 VTS_R8 VTS_R8 VTS_I4;
	InvokeHelper(0x30d9, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Zoom, CenterX, CenterY, unit);
}

void IPowerPCBView::ShowRectangle(double left, double top, double right, double bottom, long unit)
{
	static BYTE parms[] =
		VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_I4;
	InvokeHelper(0x30da, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 left, top, right, bottom, unit);
}

void IPowerPCBView::ZoomToSelection()
{
	InvokeHelper(0x30db, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBDoc properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBDoc operations

CString IPowerPCBDoc::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x2af9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x2afa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

CString IPowerPCBDoc::GetPath()
{
	CString result;
	InvokeHelper(0x2afb, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

CString IPowerPCBDoc::GetFullName()
{
	CString result;
	InvokeHelper(0x2afc, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetActiveView()
{
	LPDISPATCH result;
	InvokeHelper(0x2afd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerPCBDoc::SetSaved(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x2afe, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerPCBDoc::GetSaved()
{
	BOOL result;
	InvokeHelper(0x2afe, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerPCBDoc::SetUnit(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2aff, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

long IPowerPCBDoc::GetUnit()
{
	long result;
	InvokeHelper(0x2aff, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

double IPowerPCBDoc::GetOriginX(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2b00, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

double IPowerPCBDoc::GetOriginY(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2b01, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

void IPowerPCBDoc::SetGridX(long type, long unit, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_I4 VTS_R8;
	InvokeHelper(0x2b02, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 type, unit, newValue);
}

double IPowerPCBDoc::GetGridX(long type, long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0x2b02, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		type, unit);
	return result;
}

void IPowerPCBDoc::SetGridY(long type, long unit, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_I4 VTS_R8;
	InvokeHelper(0x2b03, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 type, unit, newValue);
}

double IPowerPCBDoc::GetGridY(long type, long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0x2b03, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		type, unit);
	return result;
}

CString IPowerPCBDoc::GetLayerName(long layer)
{
	CString result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2b04, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms,
		layer);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetAssemblyOptions()
{
	LPDISPATCH result;
	InvokeHelper(0x2b05, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerPCBDoc::SetPreference(LPCTSTR Name, const VARIANT& newValue)
{
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT;
	InvokeHelper(0x2b06, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 Name, &newValue);
}

VARIANT IPowerPCBDoc::GetPreference(LPCTSTR Name)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2b06, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms,
		Name);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetAttributes()
{
	LPDISPATCH result;
	InvokeHelper(0x2b09, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBDoc::GetLayerCount()
{
	long result;
	InvokeHelper(0x2b0a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

long IPowerPCBDoc::GetLayerType(long layer)
{
	long result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2b0b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, parms,
		layer);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetComponents()
{
	LPDISPATCH result;
	InvokeHelper(0x2b0c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetNets()
{
	LPDISPATCH result;
	InvokeHelper(0x2b0d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetPins()
{
	LPDISPATCH result;
	InvokeHelper(0x2b0e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetVias()
{
	LPDISPATCH result;
	InvokeHelper(0x2b0f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetConnections()
{
	LPDISPATCH result;
	InvokeHelper(0x2b10, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetRouteSegments()
{
	LPDISPATCH result;
	InvokeHelper(0x2b11, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetJumpers()
{
	LPDISPATCH result;
	InvokeHelper(0x2b13, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

double IPowerPCBDoc::GetBoardOutlineSurface(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2b12, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

void IPowerPCBDoc::Activate()
{
	InvokeHelper(0x2cec, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LPDISPATCH IPowerPCBDoc::GetObjects(long type, LPCTSTR value, BOOL selected)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_I4 VTS_BSTR VTS_BOOL;
	InvokeHelper(0x2ced, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		type, value, selected);
	return result;
}

void IPowerPCBDoc::SelectObjects(long type, LPCTSTR value, BOOL Select)
{
	static BYTE parms[] =
		VTS_I4 VTS_BSTR VTS_BOOL;
	InvokeHelper(0x2cee, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 type, value, Select);
}

void IPowerPCBDoc::Save()
{
	InvokeHelper(0x2cef, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void IPowerPCBDoc::SaveAs(LPCTSTR Name)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2cf0, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Name);
}

long IPowerPCBDoc::ImportNetList(LPCTSTR Name)
{
	long result;
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2cf1, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
		Name);
	return result;
}

long IPowerPCBDoc::ExportNetList(LPCTSTR Name, long ver)
{
	long result;
	static BYTE parms[] =
		VTS_BSTR VTS_I4;
	InvokeHelper(0x2cf2, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
		Name, ver);
	return result;
}

long IPowerPCBDoc::ExportASCII(LPCTSTR Name, long sections, long ver, long expandAttrs)
{
	long result;
	static BYTE parms[] =
		VTS_BSTR VTS_I4 VTS_I4 VTS_I4;
	InvokeHelper(0x2cf8, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
		Name, sections, ver, expandAttrs);
	return result;
}

long IPowerPCBDoc::ImportECOFile(LPCTSTR Name)
{
	long result;
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2cf3, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
		Name);
	return result;
}

long IPowerPCBDoc::ExportECOFile(LPCTSTR Name)
{
	long result;
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2cf4, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
		Name);
	return result;
}

long IPowerPCBDoc::CheckASCII(LPCTSTR Name, LPCTSTR ignoreNet)
{
	long result;
	static BYTE parms[] =
		VTS_BSTR VTS_BSTR;
	InvokeHelper(0x2cf5, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
		Name, ignoreNet);
	return result;
}

long IPowerPCBDoc::ExportRules(LPCTSTR Name, long ver)
{
	long result;
	static BYTE parms[] =
		VTS_BSTR VTS_I4;
	InvokeHelper(0x2cf6, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
		Name, ver);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetPartTypes()
{
	LPDISPATCH result;
	InvokeHelper(0x2b15, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBDoc::GetObjectType()
{
	long result;
	InvokeHelper(0x2b17, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetNetClasses()
{
	LPDISPATCH result;
	InvokeHelper(0x2b18, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetDrawings()
{
	LPDISPATCH result;
	InvokeHelper(0x2b19, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDoc::GetTexts()
{
	LPDISPATCH result;
	InvokeHelper(0x2b1a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDoc::AddText(LPCTSTR Text, long layer, double posX, double posY, double Height, double LineWidth, long unit, double Orientation, BOOL Mirror)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR VTS_I4 VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_I4 VTS_R8 VTS_BOOL;
	InvokeHelper(0x2cf9, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Text, layer, posX, posY, Height, LineWidth, unit, Orientation, Mirror);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBApp properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBApp operations

CString IPowerPCBApp::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBApp::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x2711, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBApp::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x2712, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

CString IPowerPCBApp::GetVersion()
{
	CString result;
	InvokeHelper(0x271a, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBApp::GetActiveDocument()
{
	LPDISPATCH result;
	InvokeHelper(0x2713, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerPCBApp::SetDefaultFilePath(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2714, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

CString IPowerPCBApp::GetDefaultFilePath()
{
	CString result;
	InvokeHelper(0x2714, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

CString IPowerPCBApp::GetFullName()
{
	CString result;
	InvokeHelper(0x2715, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void IPowerPCBApp::SetVisible(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x2716, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerPCBApp::GetVisible()
{
	BOOL result;
	InvokeHelper(0x2716, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerPCBApp::SetPreference(LPCTSTR Name, const VARIANT& newValue)
{
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT;
	InvokeHelper(0x2717, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 Name, &newValue);
}

VARIANT IPowerPCBApp::GetPreference(LPCTSTR Name)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2717, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms,
		Name);
	return result;
}

void IPowerPCBApp::SetStatusBarText(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2718, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

CString IPowerPCBApp::GetStatusBarText()
{
	CString result;
	InvokeHelper(0x2718, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

short IPowerPCBApp::GetProgressBar()
{
	short result;
	InvokeHelper(0x2719, DISPATCH_PROPERTYGET, VT_I2, (void*)&result, NULL);
	return result;
}

void IPowerPCBApp::SetProgressBar(short nNewValue)
{
	static BYTE parms[] =
		VTS_I2;
	InvokeHelper(0x2719, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

LPDISPATCH IPowerPCBApp::OpenDocument(LPCTSTR filename)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2904, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		filename);
	return result;
}

void IPowerPCBApp::Quit()
{
	InvokeHelper(0x2905, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void IPowerPCBApp::RunMacro(LPCTSTR filename, LPCTSTR macroname)
{
	static BYTE parms[] =
		VTS_BSTR VTS_BSTR;
	InvokeHelper(0x2906, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 filename, macroname);
}

void IPowerPCBApp::LockServer()
{
	InvokeHelper(0x290b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void IPowerPCBApp::UnlockServer()
{
	InvokeHelper(0x290c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LPDISPATCH IPowerPCBApp::Measure(const VARIANT& value, LPCTSTR DefaultUnit)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT VTS_BSTR;
	InvokeHelper(0x2911, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		&value, DefaultUnit);
	return result;
}

long IPowerPCBApp::GetObjectType()
{
	long result;
	InvokeHelper(0x271b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBApp::GetLibraries()
{
	LPDISPATCH result;
	InvokeHelper(0x271c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBApp::GetLibraryItems(long type, LPCTSTR Name)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_I4 VTS_BSTR;
	InvokeHelper(0x60020020, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		type, Name);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBText properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBText operations

double IPowerPCBText::GetLineWidth(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa410, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

void IPowerPCBText::SetLineWidth(long unit, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_R8;
	InvokeHelper(0xa410, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 unit, newValue);
}

double IPowerPCBText::GetPositionX(long unit, long origin)
{
	double result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0xa411, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit, origin);
	return result;
}

void IPowerPCBText::SetPositionX(long unit, long origin, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_I4 VTS_R8;
	InvokeHelper(0xa411, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 unit, origin, newValue);
}

double IPowerPCBText::GetPositionY(long unit, long origin)
{
	double result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0xa412, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit, origin);
	return result;
}

void IPowerPCBText::SetPositionY(long unit, long origin, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_I4 VTS_R8;
	InvokeHelper(0xa412, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 unit, origin, newValue);
}

double IPowerPCBText::GetHeight(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa413, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

void IPowerPCBText::SetHeight(long unit, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_R8;
	InvokeHelper(0xa413, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 unit, newValue);
}

BOOL IPowerPCBText::GetMirror(long origin)
{
	BOOL result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa414, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, parms,
		origin);
	return result;
}

void IPowerPCBText::SetMirror(long origin, BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_BOOL;
	InvokeHelper(0xa414, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 origin, bNewValue);
}

double IPowerPCBText::GetOrientation(long origin)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa415, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		origin);
	return result;
}

void IPowerPCBText::SetOrientation(long origin, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_R8;
	InvokeHelper(0xa415, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 origin, newValue);
}

long IPowerPCBText::GetHorzJustification()
{
	long result;
	InvokeHelper(0xa416, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void IPowerPCBText::SetHorzJustification(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa416, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

long IPowerPCBText::GetVertJustification()
{
	long result;
	InvokeHelper(0xa417, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void IPowerPCBText::SetVertJustification(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa417, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

long IPowerPCBText::GetLayer()
{
	long result;
	InvokeHelper(0xa418, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void IPowerPCBText::SetLayer(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa418, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

LPDISPATCH IPowerPCBText::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0xa028, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBText::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0xa029, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBText::GetObjectType()
{
	long result;
	InvokeHelper(0xa02a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

CString IPowerPCBText::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBText::GetDrawing()
{
	LPDISPATCH result;
	InvokeHelper(0xa02c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL IPowerPCBText::GetSelected()
{
	BOOL result;
	InvokeHelper(0xa02d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerPCBText::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0xa02d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

CString IPowerPCBText::GetText()
{
	CString result;
	InvokeHelper(0xa02e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void IPowerPCBText::SetText(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0xa02e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

BOOL IPowerPCBText::Delete()
{
	BOOL result;
	InvokeHelper(0x60030009, DISPATCH_METHOD, VT_BOOL, (void*)&result, NULL);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBLabel properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBLabel operations

double IPowerPCBLabel::GetLineWidth(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa410, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

void IPowerPCBLabel::SetLineWidth(long unit, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_R8;
	InvokeHelper(0xa410, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 unit, newValue);
}

double IPowerPCBLabel::GetPositionX(long unit, long origin)
{
	double result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0xa411, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit, origin);
	return result;
}

void IPowerPCBLabel::SetPositionX(long unit, long origin, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_I4 VTS_R8;
	InvokeHelper(0xa411, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 unit, origin, newValue);
}

double IPowerPCBLabel::GetPositionY(long unit, long origin)
{
	double result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0xa412, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit, origin);
	return result;
}

void IPowerPCBLabel::SetPositionY(long unit, long origin, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_I4 VTS_R8;
	InvokeHelper(0xa412, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 unit, origin, newValue);
}

double IPowerPCBLabel::GetHeight(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa413, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

void IPowerPCBLabel::SetHeight(long unit, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_R8;
	InvokeHelper(0xa413, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 unit, newValue);
}

BOOL IPowerPCBLabel::GetMirror(long origin)
{
	BOOL result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa414, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, parms,
		origin);
	return result;
}

void IPowerPCBLabel::SetMirror(long origin, BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_BOOL;
	InvokeHelper(0xa414, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 origin, bNewValue);
}

double IPowerPCBLabel::GetOrientation(long origin)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa415, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		origin);
	return result;
}

void IPowerPCBLabel::SetOrientation(long origin, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_R8;
	InvokeHelper(0xa415, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 origin, newValue);
}

long IPowerPCBLabel::GetHorzJustification()
{
	long result;
	InvokeHelper(0xa416, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void IPowerPCBLabel::SetHorzJustification(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa416, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

long IPowerPCBLabel::GetVertJustification()
{
	long result;
	InvokeHelper(0xa417, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void IPowerPCBLabel::SetVertJustification(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa417, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

long IPowerPCBLabel::GetLayer()
{
	long result;
	InvokeHelper(0xa418, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void IPowerPCBLabel::SetLayer(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa418, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

LPDISPATCH IPowerPCBLabel::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0xa7f8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBLabel::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0xa7f9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBLabel::GetObjectType()
{
	long result;
	InvokeHelper(0xa7fa, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

CString IPowerPCBLabel::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

long IPowerPCBLabel::GetLabelType()
{
	long result;
	InvokeHelper(0xa7fc, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

long IPowerPCBLabel::GetDisplay()
{
	long result;
	InvokeHelper(0xa7fd, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void IPowerPCBLabel::SetDisplay(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa7fd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

LPDISPATCH IPowerPCBLabel::GetComponent()
{
	LPDISPATCH result;
	InvokeHelper(0xa7fe, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBLabel::GetAttribute()
{
	LPDISPATCH result;
	InvokeHelper(0xa7ff, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBLabel::GetRightReading()
{
	long result;
	InvokeHelper(0xa800, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void IPowerPCBLabel::SetRightReading(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0xa800, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

BOOL IPowerPCBLabel::GetSelected()
{
	BOOL result;
	InvokeHelper(0xa801, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerPCBLabel::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0xa801, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

CString IPowerPCBLabel::GetText()
{
	CString result;
	InvokeHelper(0xa802, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

BOOL IPowerPCBLabel::Delete()
{
	BOOL result;
	InvokeHelper(0x6003000e, DISPATCH_METHOD, VT_BOOL, (void*)&result, NULL);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBRouteSegment properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBRouteSegment operations

CString IPowerPCBRouteSegment::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBRouteSegment::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x55f1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBRouteSegment::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x55f2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBRouteSegment::GetObjectType()
{
	long result;
	InvokeHelper(0x55f3, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBRouteSegment::GetNet()
{
	LPDISPATCH result;
	InvokeHelper(0x55f4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBRouteSegment::GetLayer()
{
	long result;
	InvokeHelper(0x55f5, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

double IPowerPCBRouteSegment::GetWidth(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x55f6, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

void IPowerPCBRouteSegment::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x55f7, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerPCBRouteSegment::GetSelected()
{
	BOOL result;
	InvokeHelper(0x55f7, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

long IPowerPCBRouteSegment::GetSegmentType()
{
	long result;
	InvokeHelper(0x55f8, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

double IPowerPCBRouteSegment::GetLength(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x55f9, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

VARIANT IPowerPCBRouteSegment::GetPoints(long unit)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x55fa, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms,
		unit);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBPolyline properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBPolyline operations

LPDISPATCH IPowerPCBPolyline::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x9470, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBPolyline::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x9471, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBPolyline::GetObjectType()
{
	long result;
	InvokeHelper(0x9472, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBPolyline::GetGeometry()
{
	LPDISPATCH result;
	InvokeHelper(0x9473, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

double IPowerPCBPolyline::GetLineWidth(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x9474, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

long IPowerPCBPolyline::GetOutlineType()
{
	long result;
	InvokeHelper(0x9475, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

long IPowerPCBPolyline::GetShapeType()
{
	long result;
	InvokeHelper(0x9476, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

long IPowerPCBPolyline::GetLayer()
{
	long result;
	InvokeHelper(0x9477, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

VARIANT IPowerPCBPolyline::GetPoints(long unit, long origin)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0x9478, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms,
		unit, origin);
	return result;
}

double IPowerPCBPolyline::GetCenterX(long corner, long unit, long origin)
{
	double result;
	static BYTE parms[] =
		VTS_I4 VTS_I4 VTS_I4;
	InvokeHelper(0x9479, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		corner, unit, origin);
	return result;
}

double IPowerPCBPolyline::GetCenterY(long corner, long unit, long origin)
{
	double result;
	static BYTE parms[] =
		VTS_I4 VTS_I4 VTS_I4;
	InvokeHelper(0x947a, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		corner, unit, origin);
	return result;
}

double IPowerPCBPolyline::GetRadius(long corner, long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0x947b, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		corner, unit);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBCircle properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBCircle operations

LPDISPATCH IPowerPCBCircle::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x9858, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBCircle::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x9859, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBCircle::GetObjectType()
{
	long result;
	InvokeHelper(0x985a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBCircle::GetGeometry()
{
	LPDISPATCH result;
	InvokeHelper(0x985b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

double IPowerPCBCircle::GetLineWidth(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x985c, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

long IPowerPCBCircle::GetOutlineType()
{
	long result;
	InvokeHelper(0x985d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

long IPowerPCBCircle::GetShapeType()
{
	long result;
	InvokeHelper(0x985e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

double IPowerPCBCircle::GetRadius(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x985f, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

double IPowerPCBCircle::GetCenterX(long unit, long origin)
{
	double result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0x9860, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit, origin);
	return result;
}

double IPowerPCBCircle::GetCenterY(long unit, long origin)
{
	double result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0x9861, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit, origin);
	return result;
}

long IPowerPCBCircle::GetLayer()
{
	long result;
	InvokeHelper(0x9862, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBObjs properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBObjs operations

LPDISPATCH IPowerPCBObjs::GetItem(const VARIANT& index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		&index);
	return result;
}

LPDISPATCH IPowerPCBObjs::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x32c9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBObjs::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x32ca, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBObjs::GetCount()
{
	long result;
	InvokeHelper(0x32cb, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

long IPowerPCBObjs::GetItemType(const VARIANT& index)
{
	long result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x32cc, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, parms,
		&index);
	return result;
}

long IPowerPCBObjs::GetNext(long index, long type)
{
	long result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0x32cd, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, parms,
		index, type);
	return result;
}

void IPowerPCBObjs::Add(LPDISPATCH object)
{
	static BYTE parms[] =
		VTS_DISPATCH;
	InvokeHelper(0x34bc, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 object);
}

void IPowerPCBObjs::Merge(LPDISPATCH objects)
{
	static BYTE parms[] =
		VTS_DISPATCH;
	InvokeHelper(0x34bd, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 objects);
}

void IPowerPCBObjs::Remove(const VARIANT& index)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x34be, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &index);
}

void IPowerPCBObjs::Reset()
{
	InvokeHelper(0x34bf, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void IPowerPCBObjs::Select(BOOL bSelect)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x34c0, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 bSelect);
}

void IPowerPCBObjs::Sort()
{
	InvokeHelper(0x34c1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

long IPowerPCBObjs::GetObjectType()
{
	long result;
	InvokeHelper(0x32cf, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBObjs::GetParentObject()
{
	LPDISPATCH result;
	InvokeHelper(0x32d0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBDrawing properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBDrawing operations

LPDISPATCH IPowerPCBDrawing::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x9c40, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDrawing::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x9c41, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBDrawing::GetObjectType()
{
	long result;
	InvokeHelper(0x9c42, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

CString IPowerPCBDrawing::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

BOOL IPowerPCBDrawing::GetSelected()
{
	BOOL result;
	InvokeHelper(0x9c44, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerPCBDrawing::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x9c44, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

double IPowerPCBDrawing::GetPositionX(long unit, long origin)
{
	double result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0x9c45, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit, origin);
	return result;
}

double IPowerPCBDrawing::GetPositionY(long unit, long origin)
{
	double result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0x9c46, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit, origin);
	return result;
}

LPDISPATCH IPowerPCBDrawing::GetNet()
{
	LPDISPATCH result;
	InvokeHelper(0x9c47, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDrawing::GetTexts()
{
	LPDISPATCH result;
	InvokeHelper(0x9c48, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBDrawing::GetDrawingType()
{
	long result;
	InvokeHelper(0x9c49, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBDrawing::GetGeometry()
{
	LPDISPATCH result;
	InvokeHelper(0x9c4a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBPin properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBPin operations

CString IPowerPCBPin::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBPin::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x3e81, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBPin::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x3e82, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerPCBPin::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x3e83, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerPCBPin::GetSelected()
{
	BOOL result;
	InvokeHelper(0x3e83, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

double IPowerPCBPin::GetPositionX(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x3e84, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

double IPowerPCBPin::GetPositionY(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x3e85, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

double IPowerPCBPin::GetDrillSize(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x3e86, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

LPDISPATCH IPowerPCBPin::GetComponent()
{
	LPDISPATCH result;
	InvokeHelper(0x3e87, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBPin::GetNet()
{
	LPDISPATCH result;
	InvokeHelper(0x3e88, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL IPowerPCBPin::GetIsSMD()
{
	BOOL result;
	InvokeHelper(0x3e89, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

BOOL IPowerPCBPin::GetPlated()
{
	BOOL result;
	InvokeHelper(0x3e8a, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerPCBPin::SetTestPoint(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x3e8b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

long IPowerPCBPin::GetTestPoint()
{
	long result;
	InvokeHelper(0x3e8b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void IPowerPCBPin::SetGlued(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x3e8c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerPCBPin::GetGlued()
{
	BOOL result;
	InvokeHelper(0x3e8c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

BOOL IPowerPCBPin::GetPlaneThermal()
{
	BOOL result;
	InvokeHelper(0x3e8d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

CString IPowerPCBPin::GetFunctionName()
{
	CString result;
	InvokeHelper(0x3e8e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBPin::GetAttributes()
{
	LPDISPATCH result;
	InvokeHelper(0x3e8f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBPin::GetObjectType()
{
	long result;
	InvokeHelper(0x3e90, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

long IPowerPCBPin::GetElectricalType()
{
	long result;
	InvokeHelper(0x3e91, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

long IPowerPCBPin::GetNumber()
{
	long result;
	InvokeHelper(0x3e92, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBVia properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBVia operations

CString IPowerPCBVia::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBVia::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x4269, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBVia::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x426a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerPCBVia::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x426b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerPCBVia::GetSelected()
{
	BOOL result;
	InvokeHelper(0x426b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

double IPowerPCBVia::GetPositionX(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x426c, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

double IPowerPCBVia::GetPositionY(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x426d, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

double IPowerPCBVia::GetDrillSize(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x426e, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

LPDISPATCH IPowerPCBVia::GetNet()
{
	LPDISPATCH result;
	InvokeHelper(0x426f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL IPowerPCBVia::GetPlated()
{
	BOOL result;
	InvokeHelper(0x4270, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerPCBVia::SetTestPoint(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x4271, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

long IPowerPCBVia::GetTestPoint()
{
	long result;
	InvokeHelper(0x4271, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void IPowerPCBVia::SetGlued(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x4272, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerPCBVia::GetGlued()
{
	BOOL result;
	InvokeHelper(0x4272, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

long IPowerPCBVia::GetStartLayer()
{
	long result;
	InvokeHelper(0x4273, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

long IPowerPCBVia::GetEndLayer()
{
	long result;
	InvokeHelper(0x4274, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

BOOL IPowerPCBVia::GetPlaneThermal()
{
	BOOL result;
	InvokeHelper(0x4275, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

CString IPowerPCBVia::GetType()
{
	CString result;
	InvokeHelper(0x4276, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBVia::GetAttributes()
{
	LPDISPATCH result;
	InvokeHelper(0x4277, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBVia::GetObjectType()
{
	long result;
	InvokeHelper(0x4278, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}
