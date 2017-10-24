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
#include "powerpcb.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif



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

LPDISPATCH IPowerPCBDoc::GetPartTypes()
{
	LPDISPATCH result;
	InvokeHelper(0x2b15, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
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

void IPowerPCBApp::ProcessCommand(long ID)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x290d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 ID);
}

void IPowerPCBApp::ProcessParameter(long ID, const VARIANT& value)
{
	static BYTE parms[] =
		VTS_I4 VTS_VARIANT;
	InvokeHelper(0x290e, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 ID, &value);
}

void IPowerPCBApp::ProcessPointer(long ID, double x, double y, long unit)
{
	static BYTE parms[] =
		VTS_I4 VTS_R8 VTS_R8 VTS_I4;
	InvokeHelper(0x290f, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 ID, x, y, unit);
}

long IPowerPCBApp::StrNumCmp(LPCTSTR str1, LPCTSTR str2)
{
	long result;
	static BYTE parms[] =
		VTS_BSTR VTS_BSTR;
	InvokeHelper(0x2910, DISPATCH_METHOD, VT_I4, (void*)&result, parms,
		str1, str2);
	return result;
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


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBMeasure properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBMeasure operations

void IPowerPCBMeasure::SetValue(double newValue)
{
	static BYTE parms[] =
		VTS_R8;
	InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 newValue);
}

double IPowerPCBMeasure::GetValue()
{
	double result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

void IPowerPCBMeasure::SetText(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x791c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

CString IPowerPCBMeasure::GetText()
{
	CString result;
	InvokeHelper(0x791c, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

CString IPowerPCBMeasure::GetName()
{
	CString result;
	InvokeHelper(0x791b, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBMeasure::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x7919, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBMeasure::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x791a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

CString IPowerPCBMeasure::GetUnit(long Format)
{
	CString result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x791d, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms,
		Format);
	return result;
}

CString IPowerPCBMeasure::GetPrefix(long Format)
{
	CString result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x791e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms,
		Format);
	return result;
}

double IPowerPCBMeasure::GetNumber()
{
	double result;
	InvokeHelper(0x791f, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

CString IPowerPCBMeasure::Normalize()
{
	CString result;
	InvokeHelper(0x7920, DISPATCH_METHOD, VT_BSTR, (void*)&result, NULL);
	return result;
}


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

double IPowerPCBView::GetTopLeftY(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2ee4, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
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

double IPowerPCBView::GetBottomRightY(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x2ee6, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
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


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBAssOpts properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBAssOpts operations

LPDISPATCH IPowerPCBAssOpts::GetItem(const VARIANT& index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		&index);
	return result;
}

LPDISPATCH IPowerPCBAssOpts::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x59d9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBAssOpts::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x59da, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBAssOpts::GetCount()
{
	long result;
	InvokeHelper(0x59db, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBAssOpts::Add(LPCTSTR Name)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x5bcc, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Name);
	return result;
}

void IPowerPCBAssOpts::Delete(const VARIANT& index)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x5bcd, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &index);
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBAttrs properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBAttrs operations

LPDISPATCH IPowerPCBAttrs::GetItem(const VARIANT& index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		&index);
	return result;
}

LPDISPATCH IPowerPCBAttrs::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x4a39, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBAttrs::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x4a3a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBAttrs::GetCount()
{
	long result;
	InvokeHelper(0x4a3b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBAttrs::Add(LPCTSTR Name, const VARIANT& value)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR VTS_VARIANT;
	InvokeHelper(0x4c2c, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Name, &value);
	return result;
}

void IPowerPCBAttrs::Delete(const VARIANT& index)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x4c2d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &index);
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBAttr properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBAttr operations

void IPowerPCBAttr::SetValue(const VARIANT& newValue)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 &newValue);
}

VARIANT IPowerPCBAttr::GetValue()
{
	VARIANT result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
	return result;
}

CString IPowerPCBAttr::GetName()
{
	CString result;
	InvokeHelper(0x4e23, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBAttr::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x4e21, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBAttr::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x4e22, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBAttr::GetMeasure()
{
	LPDISPATCH result;
	InvokeHelper(0x4e24, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBAttrTypes properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBAttrTypes operations

LPDISPATCH IPowerPCBAttrTypes::GetItem(const VARIANT& index)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms,
		&index);
	return result;
}

LPDISPATCH IPowerPCBAttrTypes::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x7d01, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBAttrTypes::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x7d02, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBAttrTypes::GetCount()
{
	long result;
	InvokeHelper(0x7d03, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBAttrTypes::Add(LPCTSTR Name)
{
	LPDISPATCH result;
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x7ef4, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms,
		Name);
	return result;
}

void IPowerPCBAttrTypes::Delete(const VARIANT& index)
{
	static BYTE parms[] =
		VTS_VARIANT;
	InvokeHelper(0x4c2d, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 &index);
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBAttrType properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBAttrType operations

CString IPowerPCBAttrType::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBAttrType::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x80e9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBAttrType::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x80ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL IPowerPCBAttrType::GetIsAllowed(long type)
{
	BOOL result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x80eb, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, parms,
		type);
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


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBComp properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBComp operations

CString IPowerPCBComp::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBComp::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x36b1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBComp::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x36b2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerPCBComp::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x36b3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerPCBComp::GetSelected()
{
	BOOL result;
	InvokeHelper(0x36b3, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerPCBComp::SetPartType(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x36b4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

CString IPowerPCBComp::GetPartType()
{
	CString result;
	InvokeHelper(0x36b4, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void IPowerPCBComp::SetDecal(LPCTSTR lpszNewValue)
{
	static BYTE parms[] =
		VTS_BSTR;
	InvokeHelper(0x36b5, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 lpszNewValue);
}

CString IPowerPCBComp::GetDecal()
{
	CString result;
	InvokeHelper(0x36b5, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void IPowerPCBComp::SetLayer(long nNewValue)
{
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x36b6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 nNewValue);
}

long IPowerPCBComp::GetLayer()
{
	long result;
	InvokeHelper(0x36b6, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBComp::GetPins()
{
	LPDISPATCH result;
	InvokeHelper(0x36b8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

double IPowerPCBComp::GetPositionX(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x36b9, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

double IPowerPCBComp::GetPositionY(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x36ba, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

void IPowerPCBComp::SetOrientation(double newValue)
{
	static BYTE parms[] =
		VTS_R8;
	InvokeHelper(0x36bb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 newValue);
}

double IPowerPCBComp::GetOrientation()
{
	double result;
	InvokeHelper(0x36bb, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

BOOL IPowerPCBComp::GetIsSMD()
{
	BOOL result;
	InvokeHelper(0x36bc, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerPCBComp::SetGlued(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x36bd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerPCBComp::GetGlued()
{
	BOOL result;
	InvokeHelper(0x36bd, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerPCBComp::SetInstalled(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x36be, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerPCBComp::GetInstalled()
{
	BOOL result;
	InvokeHelper(0x36be, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

BOOL IPowerPCBComp::GetSubstituted()
{
	BOOL result;
	InvokeHelper(0x36c4, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

VARIANT IPowerPCBComp::GetDecalCompatibleList()
{
	VARIANT result;
	InvokeHelper(0x36bf, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
	return result;
}

BOOL IPowerPCBComp::GetPlaced()
{
	BOOL result;
	InvokeHelper(0x36c0, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBComp::GetAttributes()
{
	LPDISPATCH result;
	InvokeHelper(0x36c1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBComp::GetDecalAttributes()
{
	LPDISPATCH result;
	InvokeHelper(0x36c2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBComp::GetPartTypeAttributes()
{
	LPDISPATCH result;
	InvokeHelper(0x36c3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBComp::GetObjectType()
{
	long result;
	InvokeHelper(0x36c5, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

CString IPowerPCBComp::GetPartTypeLogic()
{
	CString result;
	InvokeHelper(0x36c6, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void IPowerPCBComp::SetPartTypeECORegistered(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x36c8, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerPCBComp::GetPartTypeECORegistered()
{
	BOOL result;
	InvokeHelper(0x36c8, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBComp::GetPartTypeObject()
{
	LPDISPATCH result;
	InvokeHelper(0x36c9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerPCBComp::Move(double x, double y, long unit)
{
	static BYTE parms[] =
		VTS_R8 VTS_R8 VTS_I4;
	InvokeHelper(0x38a4, DISPATCH_METHOD, VT_EMPTY, NULL, parms,
		 x, y, unit);
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBPartType properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBPartType operations

CString IPowerPCBPartType::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBPartType::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x61a9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBPartType::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x61aa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerPCBPartType::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x61ab, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerPCBPartType::GetSelected()
{
	BOOL result;
	InvokeHelper(0x61ab, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBPartType::GetComponents()
{
	LPDISPATCH result;
	InvokeHelper(0x61ac, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBPartType::GetAttributes()
{
	LPDISPATCH result;
	InvokeHelper(0x61ad, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBPartType::GetObjectType()
{
	long result;
	InvokeHelper(0x61ae, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

CString IPowerPCBPartType::GetLogic()
{
	CString result;
	InvokeHelper(0x61af, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

void IPowerPCBPartType::SetECORegistered(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x61b0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerPCBPartType::GetECORegistered()
{
	BOOL result;
	InvokeHelper(0x61b0, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBNet properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBNet operations

CString IPowerPCBNet::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBNet::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x3a99, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBNet::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x3a9a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

void IPowerPCBNet::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x3a9b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerPCBNet::GetSelected()
{
	BOOL result;
	InvokeHelper(0x3a9b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBNet::GetAttributes()
{
	LPDISPATCH result;
	InvokeHelper(0x3a9c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBNet::GetNetClassAttributes()
{
	LPDISPATCH result;
	InvokeHelper(0x3a9d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBNet::GetObjectType()
{
	long result;
	InvokeHelper(0x3a9e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBNet::GetConnections()
{
	LPDISPATCH result;
	InvokeHelper(0x3a9f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

double IPowerPCBNet::GetLength(BOOL bRouted, long unit)
{
	double result;
	static BYTE parms[] =
		VTS_BOOL VTS_I4;
	InvokeHelper(0x3aa1, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		bRouted, unit);
	return result;
}

BOOL IPowerPCBNet::GetPower()
{
	BOOL result;
	InvokeHelper(0x3aa2, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBNet::GetPins()
{
	LPDISPATCH result;
	InvokeHelper(0x3aa0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBNet::GetVias()
{
	LPDISPATCH result;
	InvokeHelper(0x3aa3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
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


/////////////////////////////////////////////////////////////////////////////
// IPowerPCBConnection properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBConnection operations

CString IPowerPCBConnection::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBConnection::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x5209, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBConnection::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x520a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBConnection::GetPins()
{
	LPDISPATCH result;
	InvokeHelper(0x520b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBConnection::GetVias()
{
	LPDISPATCH result;
	InvokeHelper(0x520c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBConnection::GetNet()
{
	LPDISPATCH result;
	InvokeHelper(0x520d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBConnection::GetRouteSegments()
{
	LPDISPATCH result;
	InvokeHelper(0x520e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL IPowerPCBConnection::GetSelected()
{
	BOOL result;
	InvokeHelper(0x520f, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerPCBConnection::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x520f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

long IPowerPCBConnection::GetObjectType()
{
	long result;
	InvokeHelper(0x5212, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

double IPowerPCBConnection::GetLength(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x5213, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
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
// IPowerPCBJumper properties

/////////////////////////////////////////////////////////////////////////////
// IPowerPCBJumper operations

CString IPowerPCBJumper::GetName()
{
	CString result;
	InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBJumper::GetApplication()
{
	LPDISPATCH result;
	InvokeHelper(0x5dc1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBJumper::GetParent()
{
	LPDISPATCH result;
	InvokeHelper(0x5dc2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

long IPowerPCBJumper::GetObjectType()
{
	long result;
	InvokeHelper(0x5dc3, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
	return result;
}

double IPowerPCBJumper::GetOrientation()
{
	double result;
	InvokeHelper(0x5dc4, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
	return result;
}

LPDISPATCH IPowerPCBJumper::GetPoints()
{
	LPDISPATCH result;
	InvokeHelper(0x5dc5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

BOOL IPowerPCBJumper::GetSelected()
{
	BOOL result;
	InvokeHelper(0x5dc6, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}

void IPowerPCBJumper::SetSelected(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x5dc6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

LPDISPATCH IPowerPCBJumper::GetNet()
{
	LPDISPATCH result;
	InvokeHelper(0x5dc7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
	return result;
}

double IPowerPCBJumper::GetLength(long unit)
{
	double result;
	static BYTE parms[] =
		VTS_I4;
	InvokeHelper(0x5dc8, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, parms,
		unit);
	return result;
}

void IPowerPCBJumper::SetInstalled(BOOL bNewValue)
{
	static BYTE parms[] =
		VTS_BOOL;
	InvokeHelper(0x5dc9, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms,
		 bNewValue);
}

BOOL IPowerPCBJumper::GetInstalled()
{
	BOOL result;
	InvokeHelper(0x5dc9, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
	return result;
}
