/* Copyright Mentor Graphics Corporation 2007

All Rights Reserved.

THIS WORK CONTAINS TRADE SECRET
AND PROPRIETARY INFORMATION WHICH IS THE
PROPERTY OF MENTOR GRAPHICS
CORPORATION OR ITS LICENSORS AND IS
SUBJECT TO LICENSE TERMS. 
*/

//////////////////////////////////////////////////////////////////////////////
//
// SINK.CPP : implementation file
//
//////////////////////////////////////////////////////////////////////////////
// This is a part of the PADS Logic OLE Automation server SAMPLE2 sample.
// Copyright Mentor Graphics Corporation 2007
// All rights reserved.
//
// This source code is only intended as a supplement to the PADS Logic OLE 
// Automation Server API Help file.
//////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "oleapp.h"
#include "sink.h"
#include "oleappdlg.h"
#include "PowerLogic.inc"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// PowerLogicSink

// IMPLEMENT_DYNCREATE(PowerLogicSink, CCmdTarget)

PowerLogicSink::PowerLogicSink()
{
	EnableAutomation();
}

PowerLogicSink::~PowerLogicSink()
{
}


void PowerLogicSink::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(PowerLogicSink, CCmdTarget)
	//{{AFX_MSG_MAP(PowerLogicSink)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(PowerLogicSink, CCmdTarget)
	//{{AFX_DISPATCH_MAP(PowerLogicSink)
	DISP_FUNCTION_ID(PowerLogicSink, "OnOpenDocument", IPOWERLOGICAPP_OPENDOCUMENTEVENT, OnOpenDocument, VT_EMPTY, VTS_DISPATCH)
	DISP_FUNCTION_ID(PowerLogicSink, "OnSelectionChange", IPOWERLOGICDOC_SELECTIONCHANGEEVENT, OnSelectionChange, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(PowerLogicSink, "OnQuit", IPOWERLOGICAPP_QUITEVENT, OnQuit, VT_EMPTY, VTS_NONE)
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IPowerLogicSink to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {667FDFC1-D446-11D0-BCBF-444553540000}
static const IID IID_IPowerLogicSink =
{ 0x667fdfc1, 0xd446, 0x11d0, { 0xbc, 0xbf, 0x44, 0x45, 0x53, 0x54, 0x0, 0x0 } };

BEGIN_INTERFACE_MAP(PowerLogicSink, CCmdTarget)
	INTERFACE_PART(PowerLogicSink, DIID__PowerLogicAppEvents, Dispatch)
	INTERFACE_PART(PowerLogicSink, DIID__PowerLogicDocEvents, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// PowerLogicSink message handlers

void PowerLogicSink::OnOpenDocument(LPDISPATCH pDoc)
{
	((COleappDlg *)AfxGetMainWnd())->SetDlgItemText(IDC_TXT_STATUS, "Received OnOpenDocument from server.");
	((COleappDlg *)AfxGetMainWnd())->Refresh();
}

void PowerLogicSink::OnSelectionChange()
{
	((COleappDlg *)AfxGetMainWnd())->SetDlgItemText(IDC_TXT_STATUS, "Received OnSelectionChange from server.");
	((COleappDlg *)AfxGetMainWnd())->RefreshGeneralInfos();
	((COleappDlg *)AfxGetMainWnd())->RefreshSelectionInfos();
}

void PowerLogicSink::OnQuit()
{
	((COleappDlg *)AfxGetMainWnd())->SetDlgItemText(IDC_TXT_STATUS, "Received OnQuit from server.");
	((COleappDlg *)AfxGetMainWnd())->SafeDisconnect();
	((COleappDlg *)AfxGetMainWnd())->Refresh();
}
