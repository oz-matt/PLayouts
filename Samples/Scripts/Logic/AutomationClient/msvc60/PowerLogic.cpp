/* Copyright Mentor Graphics Corporation 2007

All Rights Reserved.

THIS WORK CONTAINS TRADE SECRET
AND PROPRIETARY INFORMATION WHICH IS THE
PROPERTY OF MENTOR GRAPHICS
CORPORATION OR ITS LICENSORS AND IS
SUBJECT TO LICENSE TERMS. 
*/

// Machine generated IDispatch wrapper class(es) created with ClassWizard

#include "stdafx.h"
#include "powerlogic.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif



/////////////////////////////////////////////////////////////////////////////
// IPowerLogicApp properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicApp operations

CString IPowerLogicApp::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicApp::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x2711, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicApp::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x2712, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

CString IPowerLogicApp::GetFullName()
{
	CString result;
	InvokeHelper(0x2713, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

BOOL IPowerLogicApp::GetVisible()
{
	BOOL result;
	InvokeHelper(0x2714, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerLogicApp::SetVisible(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x2714, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH IPowerLogicApp::GetActiveDocument()
{
	LPDISPATCH result;
	InvokeHelper(0x2715, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

CString IPowerLogicApp::GetDefaultFilePath()
{
	CString result;
	InvokeHelper(0x2716, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void IPowerLogicApp::SetDefaultFilePath(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2716, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

VARIANT IPowerLogicApp::GetPreference(LPCTSTR Name)
{
	VARIANT result;
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2718, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms,
		Name);
	return result;
}

void IPowerLogicApp::SetPreference(LPCTSTR Name, const VARIANT& newValue)
{
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT;
	InvokeHelper(0x2718, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 Name, &newValue);
}

CString IPowerLogicApp::GetStatusBarText()
{
	CString result;
	InvokeHelper(0x2719, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void IPowerLogicApp::SetStatusBarText(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2719, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

short IPowerLogicApp::GetProgressBar()
{
	short result;
	InvokeHelper(0x271b, DISPATCH_PROPERTYGET, VT_I2, (void*)&result, NULL);
	return result;
}

void IPowerLogicApp::SetProgressBar(short nNewValue)
{
	static BYTE parms[] =
		VTS_I2;
	InvokeHelper(0x271b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

CString IPowerLogicApp::GetVersion()
{
	CString result;
	InvokeHelper(0x271a, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void IPowerLogicApp::Quit()
{
	InvokeHelper(0x2904, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LPDISPATCH IPowerLogicApp::OpenDocument(LPCTSTR filename)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2905, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		filename);
	return result;
}

void IPowerLogicApp::RunMacro(LPCTSTR filename, LPCTSTR macroname)
{
	static BYTE parms[] =
		VTS_BSTR VTS_BSTR;
	InvokeHelper(0x2906, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 filename, macroname);
}

void IPowerLogicApp::ProcessCommand(long cmdId)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2907, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 cmdId);
}

void IPowerLogicApp::ProcessParameter(long parId, const VARIANT& Value)
{
	static BYTE parms[] =
		VTS_I4 VTS_VARIANT;
	InvokeHelper(0x2908, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 parId, &Value);
}

void IPowerLogicApp::ProcessPointer(long typeID, double x, double y, long unit)
{
	static BYTE parms[] =
		VTS_I4 VTS_R8 VTS_R8 VTS_I4;
	InvokeHelper(0x2909, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 typeID, x, y, unit);
}

long IPowerLogicApp::StrNumCmp(LPCTSTR str1, LPCTSTR str2)
{
	long result;
	static BYTE parms[] =
		VTS_BSTR VTS_BSTR;
	InvokeHelper(0x290a, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
		str1, str2);
	return result;
}

void IPowerLogicApp::LockServer()
{
	InvokeHelper(0x290b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void IPowerLogicApp::UnlockServer()
{
	InvokeHelper(0x290c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LPDISPATCH IPowerLogicApp::Measure(const VARIANT& Value, LPCTSTR DefaultUnit)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT VTS_BSTR;
	InvokeHelper(0x290d, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		&Value, DefaultUnit);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerLogicDoc properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicDoc operations

CString IPowerLogicDoc::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicDoc::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x2af9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicDoc::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x2afa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

CString IPowerLogicDoc::GetFullName()
{
	CString result;
	InvokeHelper(0x2afb, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

CString IPowerLogicDoc::GetPath()
{
	CString result;
	InvokeHelper(0x2afc, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

BOOL IPowerLogicDoc::GetSaved()
{
	BOOL result;
	InvokeHelper(0x2afd, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerLogicDoc::SetSaved(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x2afd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH IPowerLogicDoc::GetActiveView()
{
	LPDISPATCH result;
	InvokeHelper(0x2afe, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicDoc::GetActiveSheet()
{
	LPDISPATCH result;
	InvokeHelper(0x2aff, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicDoc::GetSheets()
{
	LPDISPATCH result;
	InvokeHelper(0x2b00, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicDoc::GetAncestorSheets()
{
	LPDISPATCH result;
	InvokeHelper(0x2b02, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicDoc::GetComponents()
{
	LPDISPATCH result;
	InvokeHelper(0x2b03, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicDoc::GetNets()
{
	LPDISPATCH result;
	InvokeHelper(0x2b04, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicDoc::GetGates()
{
	LPDISPATCH result;
	InvokeHelper(0x2b05, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerLogicDoc::SetGridX(long type, long unit, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_I4 VTS_R8;
	InvokeHelper(0x2b06, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 type, unit, newValue);
}

double IPowerLogicDoc::GetGridX(long type, long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0x2b06, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		type, unit);
	return result;
}

void IPowerLogicDoc::SetGridY(long type, long unit, double newValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_I4 VTS_R8;
	InvokeHelper(0x2b07, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 type, unit, newValue);
}

double IPowerLogicDoc::GetGridY(long type, long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0x2b07, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		type, unit);
	return result;
}

double IPowerLogicDoc::GetOriginX(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2b08, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

double IPowerLogicDoc::GetOriginY(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2b09, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

LPDISPATCH IPowerLogicDoc::GetPins()
{
	LPDISPATCH result;
	InvokeHelper(0x2b0a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicDoc::GetPartTypes()
{
	LPDISPATCH result;
	InvokeHelper(0x2b0b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerLogicDoc::Activate()
{
	InvokeHelper(0x2cec, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void IPowerLogicDoc::ExportASCII(LPCTSTR Name, long ver)
{
	static BYTE parms[] =
		VTS_BSTR VTS_I4;
	InvokeHelper(0x2ced, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Name, ver);
}

void IPowerLogicDoc::ExportNetList(LPCTSTR Path, long ver)
{
	static BYTE parms[] =
		VTS_BSTR VTS_I4;
	InvokeHelper(0x2cee, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Path, ver);
}

void IPowerLogicDoc::GenerateECO(LPCTSTR fileToComp, LPCTSTR Path)
{
	static BYTE parms[] =
		VTS_BSTR VTS_BSTR;
	InvokeHelper(0x2cef, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 fileToComp, Path);
}

LPDISPATCH IPowerLogicDoc::GetObjects(long type, LPCTSTR Name, BOOL selected)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_I4 VTS_BSTR VTS_BOOL;
	InvokeHelper(0x2cf0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		type, Name, selected);
	return result;
}

void IPowerLogicDoc::ImportASCII(LPCTSTR Path)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2cf1, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Path);
}

void IPowerLogicDoc::ImportECO(LPCTSTR Path)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2cf2, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Path);
}

void IPowerLogicDoc::Save()
{
	InvokeHelper(0x2cf3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void IPowerLogicDoc::SaveAs(LPCTSTR Name)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x2cf5, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Name);
}

void IPowerLogicDoc::SelectObjects(long type, LPCTSTR Name, BOOL Select)
{
	static BYTE parms[] =
		VTS_I4 VTS_BSTR VTS_BOOL;
	InvokeHelper(0x2cf4, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 type, Name, Select);
}


/////////////////////////////////////////////////////////////////////////////
// IPowerLogicView properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicView operations

CString IPowerLogicView::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicView::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x2ee1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicView::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x2ee2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

double IPowerLogicView::GetTopLeftX(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2ee3, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

double IPowerLogicView::GetTopLeftY(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2ee4, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

double IPowerLogicView::GetBottomRightX(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2ee5, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

double IPowerLogicView::GetBottomRightY(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2ee6, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

void IPowerLogicView::Refresh()
{
	InvokeHelper(0x30d4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void IPowerLogicView::Pan(double x, double y, long unit)
{
	static BYTE parms[] =
		VTS_R8 VTS_R8 VTS_I4;
	InvokeHelper(0x30d5, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 x, y, unit);
}

void IPowerLogicView::SetExtentsToAll()
{
	InvokeHelper(0x30d6, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void IPowerLogicView::SetExtentsToSheet()
{
	InvokeHelper(0x30d7, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void IPowerLogicView::SetExtents(double TopLeftX, double TopLeftY, double BottomRightX, double BottomRightY, long unit)
{
	static BYTE parms[] =
		VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_I4;
	InvokeHelper(0x30d8, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 TopLeftX, TopLeftY, BottomRightX, BottomRightY, unit);
}


/////////////////////////////////////////////////////////////////////////////
// IPowerLogicSheet properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicSheet operations

CString IPowerLogicSheet::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void IPowerLogicSheet::SetName(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

LPDISPATCH IPowerLogicSheet::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x32c9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicSheet::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x32ca, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicSheet::GetChildSheets()
{
	LPDISPATCH result;
	InvokeHelper(0x32cb, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicSheet::GetParentSheet()
{
	LPDISPATCH result;
	InvokeHelper(0x32cc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicSheet::GetGates()
{
	LPDISPATCH result;
	InvokeHelper(0x32cd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicSheet::GetNets()
{
	LPDISPATCH result;
	InvokeHelper(0x32ce, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicSheet::GetComponents()
{
	LPDISPATCH result;
	InvokeHelper(0x32cf, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicSheet::GetPins()
{
	LPDISPATCH result;
	InvokeHelper(0x32d0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicSheet::GetPartTypes()
{
	LPDISPATCH result;
	InvokeHelper(0x32d1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerLogicSheet::Activate()
{
	InvokeHelper(0x34bc, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

LPDISPATCH IPowerLogicSheet::GetObjects(long type, LPCTSTR Name, BOOL selected)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_I4 VTS_BSTR VTS_BOOL;
	InvokeHelper(0x34bd, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		type, Name, selected);
	return result;
}

LPDISPATCH IPowerLogicSheet::AddGate(LPCTSTR partTypeName, LPCTSTR refDes, long gateIndex, double PositionX, double PositionY, long unit)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR VTS_BSTR VTS_I4 VTS_R8 VTS_R8 VTS_I4;
	InvokeHelper(0x34be, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		partTypeName, refDes, gateIndex, PositionX, PositionY, unit);
	return result;
}

LPDISPATCH IPowerLogicSheet::AddComponent(LPCTSTR partTypeName, LPCTSTR refDes, double PositionX, double PositionY, double DeltaX, double DeltaY, long unit)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR VTS_BSTR VTS_R8 VTS_R8 VTS_R8 VTS_R8 VTS_I4;
	InvokeHelper(0x34bf, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		partTypeName, refDes, PositionX, PositionY, DeltaX, DeltaY, unit);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerLogicSheets properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicSheets operations

LPDISPATCH IPowerLogicSheets::GetItem(const VARIANT& index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		&index);
	return result;
}

LPDISPATCH IPowerLogicSheets::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x36b1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicSheets::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x36b2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerLogicSheets::GetCount()
{
	long result;
	InvokeHelper(0x36b3, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicSheets::Add()
{
	LPDISPATCH result;
	InvokeHelper(0x38a4, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerLogicSheets::Delete(const VARIANT& index)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x38a5, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &index);
}


/////////////////////////////////////////////////////////////////////////////
// IPowerLogicObjs properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicObjs operations

LPDISPATCH IPowerLogicObjs::GetItem(const VARIANT& index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		&index);
	return result;
}

LPDISPATCH IPowerLogicObjs::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x4a39, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicObjs::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x4a3a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerLogicObjs::GetCount()
{
	long result;
	InvokeHelper(0x4a3b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

long IPowerLogicObjs::GetItemType(const VARIANT& index)
{
	long result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x4a3d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, parms,
		&index);
	return result;
}

long IPowerLogicObjs::GetNext(long index, long type)
{
	long result;
	static BYTE parms[] =
		VTS_I4 VTS_I4;
	InvokeHelper(0x4a3e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, parms,
		index, type);
	return result;
}

void IPowerLogicObjs::Add(LPDISPATCH object)
{
	static BYTE parms[] =
		VTS_DISPATCH;
	InvokeHelper(0x4c2c, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 object);
}

void IPowerLogicObjs::Merge(LPDISPATCH objects)
{
	static BYTE parms[] =
		VTS_DISPATCH;
	InvokeHelper(0x4c2d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 objects);
}

void IPowerLogicObjs::Remove(const VARIANT& index)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x4c2e, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &index);
}

void IPowerLogicObjs::Reset()
{
	InvokeHelper(0x4c2f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}

void IPowerLogicObjs::Select(BOOL Select)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x4c30, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 Select);
}

void IPowerLogicObjs::Sort()
{
	InvokeHelper(0x4c31, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// IPowerLogicGate properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicGate operations

CString IPowerLogicGate::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void IPowerLogicGate::SetName(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

LPDISPATCH IPowerLogicGate::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x3e81, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicGate::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x3e82, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicGate::GetComponent()
{
	LPDISPATCH result;
	InvokeHelper(0x3e83, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerLogicGate::GetObjectType()
{
	long result;
	InvokeHelper(0x3e84, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicGate::GetPins()
{
	LPDISPATCH result;
	InvokeHelper(0x3e85, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL IPowerLogicGate::GetSelected()
{
	BOOL result;
	InvokeHelper(0x3e86, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerLogicGate::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x3e86, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH IPowerLogicGate::GetSheet()
{
	LPDISPATCH result;
	InvokeHelper(0x3e87, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerLogicGate::GetSwapClass()
{
	long result;
	InvokeHelper(0x3e88, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

double IPowerLogicGate::GetPositionX(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x3e89, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

double IPowerLogicGate::GetPositionY(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x3e8a, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

BOOL IPowerLogicGate::GetIsConnector()
{
	BOOL result;
	InvokeHelper(0x3e8c, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

BOOL IPowerLogicGate::GetReflectedX()
{
	BOOL result;
	InvokeHelper(0x3e8d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerLogicGate::SetReflectedX(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x3e8d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerLogicGate::GetReflectedY()
{
	BOOL result;
	InvokeHelper(0x3e8e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerLogicGate::SetReflectedY(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x3e8e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerLogicGate::GetRotated90()
{
	BOOL result;
	InvokeHelper(0x3e8f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerLogicGate::SetRotated90(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x3e8f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerLogicGate::GetVisibility(long Item, LPCTSTR attrName)
{
	BOOL result;
	static BYTE parms[] =
		VTS_I4 VTS_BSTR;
	InvokeHelper(0x3e90, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, parms,
		Item, attrName);
	return result;
}

void IPowerLogicGate::SetVisibility(long Item, LPCTSTR attrName, BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_I4 VTS_BSTR VTS_BOOL;
	InvokeHelper(0x3e90, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 Item, attrName, bNewValue);
}

long IPowerLogicGate::GetNumber()
{
	long result;
	InvokeHelper(0x3e91, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

void IPowerLogicGate::Move(double x, double y, long unit)
{
	static BYTE parms[] =
		VTS_R8 VTS_R8 VTS_I4;
	InvokeHelper(0x4075, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 x, y, unit);
}

void IPowerLogicGate::Delete()
{
	InvokeHelper(0x4076, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// IPowerLogicComp properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicComp operations

CString IPowerLogicComp::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void IPowerLogicComp::SetName(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

LPDISPATCH IPowerLogicComp::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x3a99, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicComp::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x3a9a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

CString IPowerLogicComp::GetPartType()
{
	CString result;
	InvokeHelper(0x3a9b, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

CString IPowerLogicComp::GetPartTypeLogic()
{
	CString result;
	InvokeHelper(0x3aa5, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicComp::GetAttributes()
{
	LPDISPATCH result;
	InvokeHelper(0x3a9c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerLogicComp::SetPCBDecal(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x3a9d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

CString IPowerLogicComp::GetPCBDecal()
{
	CString result;
	InvokeHelper(0x3a9d, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicComp::GetPins()
{
	LPDISPATCH result;
	InvokeHelper(0x3a9e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicComp::GetGates()
{
	LPDISPATCH result;
	InvokeHelper(0x3a9f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerLogicComp::GetObjectType()
{
	long result;
	InvokeHelper(0x3aa0, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

BOOL IPowerLogicComp::GetSelected()
{
	BOOL result;
	InvokeHelper(0x3aa1, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerLogicComp::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x3aa1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH IPowerLogicComp::GetSheets()
{
	LPDISPATCH result;
	InvokeHelper(0x3aa2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicComp::GetUnusedGates()
{
	LPDISPATCH result;
	InvokeHelper(0x3aa4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicComp::GetPartTypeObject()
{
	LPDISPATCH result;
	InvokeHelper(0x3aa6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerLogicComp::Delete()
{
	InvokeHelper(0x3c8d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// IPowerLogicAttrs properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicAttrs operations

LPDISPATCH IPowerLogicAttrs::GetItem(const VARIANT& index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		&index);
	return result;
}

LPDISPATCH IPowerLogicAttrs::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x5209, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicAttrs::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x520a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerLogicAttrs::GetCount()
{
	long result;
	InvokeHelper(0x520b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicAttrs::Add(LPCTSTR Name, const VARIANT& Value)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT;
	InvokeHelper(0x53fc, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Name, &Value);
	return result;
}

LPDISPATCH IPowerLogicAttrs::AddMeasure(LPCTSTR Name, double Value, LPCTSTR PrefixUnit)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR VTS_R8 VTS_BSTR;
	InvokeHelper(0x53fe, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Name, Value, PrefixUnit);
	return result;
}

void IPowerLogicAttrs::Delete(const VARIANT& varIndex)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x53fd, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &varIndex);
}


/////////////////////////////////////////////////////////////////////////////
// IPowerLogicAttr properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicAttr operations

VARIANT IPowerLogicAttr::GetValue()
{
	VARIANT result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
	return result;
}

void IPowerLogicAttr::SetValue(const VARIANT& newValue)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 &newValue);
}

CString IPowerLogicAttr::GetName()
{
	CString result;
	InvokeHelper(0x4e21, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void IPowerLogicAttr::SetName(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x4e21, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

LPDISPATCH IPowerLogicAttr::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x4e22, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicAttr::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x4e23, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicAttr::GetMeasure()
{
	LPDISPATCH result;
	InvokeHelper(0x4e24, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerLogicMeasure properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicMeasure operations

void IPowerLogicMeasure::SetValue(double newValue)
{
	static BYTE parms[] =
		VTS_R8;
	InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 newValue);
}

double IPowerLogicMeasure::GetValue()
{
	double result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

void IPowerLogicMeasure::SetText(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x791c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

CString IPowerLogicMeasure::GetText()
{
	CString result;
	InvokeHelper(0x791c, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

CString IPowerLogicMeasure::GetName()
{
	CString result;
	InvokeHelper(0x791b, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicMeasure::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x7919, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicMeasure::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x791a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

CString IPowerLogicMeasure::GetUnit(long Format)
{
	CString result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x791d, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms,
		Format);
	return result;
}

CString IPowerLogicMeasure::GetPrefix(long Format)
{
	CString result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x791e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms,
		Format);
	return result;
}

double IPowerLogicMeasure::GetNumber()
{
	double result;
	InvokeHelper(0x791f, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

CString IPowerLogicMeasure::Normalize()
{
	CString result;
	InvokeHelper(0x7920, DISPATCH_METHOD, VT_BSTR, (void*)&result, NULL);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerLogicPartType properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicPartType operations

CString IPowerLogicPartType::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicPartType::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x61a9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicPartType::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x61aa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerLogicPartType::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x61ab, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerLogicPartType::GetSelected()
{
	BOOL result;
	InvokeHelper(0x61ab, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicPartType::GetComponents()
{
	LPDISPATCH result;
	InvokeHelper(0x61ac, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerLogicPartType::GetObjectType()
{
	long result;
	InvokeHelper(0x61ad, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

CString IPowerLogicPartType::GetLogic()
{
	CString result;
	InvokeHelper(0x61ae, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerLogicAttrTypes properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicAttrTypes operations

LPDISPATCH IPowerLogicAttrTypes::GetItem(const VARIANT& index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		&index);
	return result;
}

LPDISPATCH IPowerLogicAttrTypes::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x7d01, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicAttrTypes::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x7d02, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerLogicAttrTypes::GetCount()
{
	long result;
	InvokeHelper(0x7d03, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerLogicAttrType properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicAttrType operations

CString IPowerLogicAttrType::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicAttrType::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x80e9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicAttrType::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x80ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL IPowerLogicAttrType::GetIsAllowed(long type)
{
	BOOL result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x80eb, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, parms,
		type);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerLogicNet properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicNet operations

CString IPowerLogicNet::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicNet::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x4269, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicNet::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x426a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

double IPowerLogicNet::GetWidth(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x426b, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

LPDISPATCH IPowerLogicNet::GetPins()
{
	LPDISPATCH result;
	InvokeHelper(0x426c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerLogicNet::GetObjectType()
{
	long result;
	InvokeHelper(0x426d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

BOOL IPowerLogicNet::GetSelected()
{
	BOOL result;
	InvokeHelper(0x426e, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerLogicNet::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x426e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH IPowerLogicNet::GetSheets()
{
	LPDISPATCH result;
	InvokeHelper(0x426f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicNet::GetAttributes()
{
	LPDISPATCH result;
	InvokeHelper(0x4270, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerLogicPin properties

/////////////////////////////////////////////////////////////////////////////
// IPowerLogicPin operations

CString IPowerLogicPin::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicPin::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x4651, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicPin::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x4652, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

CString IPowerLogicPin::GetGatePinName()
{
	CString result;
	InvokeHelper(0x4653, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

long IPowerLogicPin::GetNumber()
{
	long result;
	InvokeHelper(0x4654, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicPin::GetGate()
{
	LPDISPATCH result;
	InvokeHelper(0x4655, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerLogicPin::GetNet()
{
	LPDISPATCH result;
	InvokeHelper(0x4656, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerLogicPin::GetObjectType()
{
	long result;
	InvokeHelper(0x4657, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

BOOL IPowerLogicPin::GetSelected()
{
	BOOL result;
	InvokeHelper(0x4658, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerLogicPin::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x4658, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

CString IPowerLogicPin::GetFunctionName()
{
	CString result;
	InvokeHelper(0x4659, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

long IPowerLogicPin::GetElectricalType()
{
	long result;
	InvokeHelper(0x465a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

long IPowerLogicPin::GetSwapClass()
{
	long result;
	InvokeHelper(0x465b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

double IPowerLogicPin::GetPositionX(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x465c, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

double IPowerLogicPin::GetPositionY(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x465d, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

LPDISPATCH IPowerLogicPin::GetComponent()
{
	LPDISPATCH result;
	InvokeHelper(0x465e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}
