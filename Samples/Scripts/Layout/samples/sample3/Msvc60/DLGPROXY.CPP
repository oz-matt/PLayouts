/* Copyright Mentor Graphics Corporation 2003

    All Rights Reserved.

 THIS WORK CONTAINS TRADE SECRET
 AND PROPRIETARY INFORMATION WHICH IS THE
 PROPERTY OF MENTOR GRAPHICS
 CORPORATION OR ITS LICENSORS AND IS
 SUBJECT TO LICENSE TERMS. 
*/
// DlgProxy.cpp : implementation file
//

#include "stdafx.h"
#include "Client.h"
#include "DlgProxy.h"
#include "sink.h"
#include "ClientDl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CClientDlgAutoProxy

IMPLEMENT_DYNCREATE(CClientDlgAutoProxy, CCmdTarget)

CClientDlgAutoProxy::CClientDlgAutoProxy()
{
	EnableAutomation();
	
	// To keep the application running as long as an OLE automation 
	//	object is active, the constructor calls AfxOleLockApp.
	AfxOleLockApp();

	// Get access to the dialog through the application's
	//  main window pointer.  Set the proxy's internal pointer
	//  to point to the dialog, and set the dialog's back pointer to
	//  this proxy.
	ASSERT (AfxGetApp()->m_pMainWnd != NULL);
	ASSERT_VALID (AfxGetApp()->m_pMainWnd);
	ASSERT_KINDOF(CClientDlg, AfxGetApp()->m_pMainWnd);
	m_pDialog = (CClientDlg*) AfxGetApp()->m_pMainWnd;
	m_pDialog->m_pAutoProxy = this;
}

CClientDlgAutoProxy::~CClientDlgAutoProxy()
{
	// To terminate the application when all objects created with
	// 	with OLE automation, the destructor calls AfxOleUnlockApp.
	//  Among other things, this will destroy the main dialog
	AfxOleUnlockApp();
}

void CClientDlgAutoProxy::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CClientDlgAutoProxy, CCmdTarget)
	//{{AFX_MSG_MAP(CClientDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CClientDlgAutoProxy, CCmdTarget)
	//{{AFX_DISPATCH_MAP(CClientDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IClient to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {EFD42725-D443-11D0-BCBF-444553540000}
static const IID IID_IClient =
{ 0xefd42725, 0xd443, 0x11d0, { 0xbc, 0xbf, 0x44, 0x45, 0x53, 0x54, 0x0, 0x0 } };

BEGIN_INTERFACE_MAP(CClientDlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CClientDlgAutoProxy, IID_IClient, Dispatch)
END_INTERFACE_MAP()

// The IMPLEMENT_OLECREATE2 macro is defined in StdAfx.h of this project
// {EFD42723-D443-11D0-BCBF-444553540000}
IMPLEMENT_OLECREATE2(CClientDlgAutoProxy, "Client.Application", 0xefd42723, 0xd443, 0x11d0, 0xbc, 0xbf, 0x44, 0x45, 0x53, 0x54, 0x0, 0x0)

/////////////////////////////////////////////////////////////////////////////
// CClientDlgAutoProxy message handlers
