//////////////////////////////////////////////////////////////////////////////
//
// SINK.CPP : implementation file
//
//////////////////////////////////////////////////////////////////////////////
// This is a part of the PADS-PowerPCB OLE Automation server SAMPLE2 sample.
// Copyright (C) 2003 Mentor Graphics Corp.
// All rights reserved.
//
// This source code is only intended as a supplement to the PADS-PowerPCB OLE 
// Automation Server API Help file.
//////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "client.h"
#include "sink.h"
#include "clientDl.h"
#include "powerpcb.inc"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// PowerPCBSink

// IMPLEMENT_DYNCREATE(PowerPCBSink, CCmdTarget)

PowerPCBSink::PowerPCBSink()
{
	EnableAutomation();
}

PowerPCBSink::~PowerPCBSink()
{
}


void PowerPCBSink::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}


BEGIN_MESSAGE_MAP(PowerPCBSink, CCmdTarget)
	//{{AFX_MSG_MAP(PowerPCBSink)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(PowerPCBSink, CCmdTarget)
	//{{AFX_DISPATCH_MAP(PowerPCBSink)
	DISP_FUNCTION_ID(PowerPCBSink, "OnOpenDocument", IPOWERPCBAPP_OPENDOCUMENTEVENT, OnOpenDocument, VT_EMPTY, VTS_DISPATCH)
	DISP_FUNCTION_ID(PowerPCBSink, "OnSelectionChange", IPOWERPCBDOC_SELECTIONCHANGEEVENT, OnSelectionChange, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION_ID(PowerPCBSink, "OnQuit", IPOWERPCBAPP_QUITEVENT, OnQuit, VT_EMPTY, VTS_NONE)
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IPowerPCBSink to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {667FDFC1-D446-11D0-BCBF-444553540000}
static const IID IID_IPowerPCBSink =
{ 0x667fdfc1, 0xd446, 0x11d0, { 0xbc, 0xbf, 0x44, 0x45, 0x53, 0x54, 0x0, 0x0 } };

BEGIN_INTERFACE_MAP(PowerPCBSink, CCmdTarget)
	INTERFACE_PART(PowerPCBSink, DIID__PowerPCBAppEvents, Dispatch)
	INTERFACE_PART(PowerPCBSink, DIID__PowerPCBDocEvents, Dispatch)
END_INTERFACE_MAP()

/////////////////////////////////////////////////////////////////////////////
// PowerPCBSink message handlers

void PowerPCBSink::OnOpenDocument(LPDISPATCH pDoc)
{
	((CClientDlg *)AfxGetMainWnd())->SetDlgItemText(IDC_TXT_STATUS, "Received OnOpenDocument from server.");
	((CClientDlg *)AfxGetMainWnd())->Refresh();
}

void PowerPCBSink::OnSelectionChange()
{
	((CClientDlg *)AfxGetMainWnd())->SetDlgItemText(IDC_TXT_STATUS, "Received OnSelectionChange from server.");
	((CClientDlg *)AfxGetMainWnd())->RefreshGeneralInfos();
	((CClientDlg *)AfxGetMainWnd())->RefreshSelectionInfos();
}

void PowerPCBSink::OnQuit()
{
	((CClientDlg *)AfxGetMainWnd())->SetDlgItemText(IDC_TXT_STATUS, "Received OnQuit from server.");
	((CClientDlg *)AfxGetMainWnd())->SafeDisconnect();
	((CClientDlg *)AfxGetMainWnd())->Refresh();
}
