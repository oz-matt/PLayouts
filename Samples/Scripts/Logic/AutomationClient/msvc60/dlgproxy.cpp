/* Copyright Mentor Graphics Corporation 2007

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
#include "oleapp.h"
#include "DlgProxy.h"
#include "oleappDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// COleappDlgAutoProxy

IMPLEMENT_DYNCREATE(COleappDlgAutoProxy, CCmdTarget)

COleappDlgAutoProxy::COleappDlgAutoProxy()
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
	ASSERT_KINDOF(COleappDlg, AfxGetApp()->m_pMainWnd);
	m_pDialog = (COleappDlg*) AfxGetApp()->m_pMainWnd;
	m_pDialog->m_pAutoProxy = this;
}

COleappDlgAutoProxy::~COleappDlgAutoProxy()
{
	// To terminate the application when all objects created with
	// 	with OLE automation, the destructor calls AfxOleUnlockApp.
	//  Among other things, this will destroy the main dialog
	AfxOleUnlockApp();
}

void COleappDlgAutoProxy::OnFinalRelease()
{
	// When the last reference for an automation object is released
	// OnFinalRelease is called.  The base class will automatically
	// deletes the object.  Add additional cleanup required for your
	// object before calling the base class.

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(COleappDlgAutoProxy, CCmdTarget)
	//{{AFX_MSG_MAP(COleappDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(COleappDlgAutoProxy, CCmdTarget)
	//{{AFX_DISPATCH_MAP(COleappDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove mapping macros here.
	//}}AFX_DISPATCH_MAP
END_DISPATCH_MAP()

// Note: we add support for IID_IOleapp to support typesafe binding
//  from VBA.  This IID must match the GUID that is attached to the 
//  dispinterface in the .ODL file.

// {647950C5-5221-11D2-81CA-0000A00F0ECE}
static const IID IID_IOleapp =
{ 0x647950c5, 0x5221, 0x11d2, { 0x81, 0xca, 0x0, 0x0, 0xa0, 0xf, 0xe, 0xce } };

BEGIN_INTERFACE_MAP(COleappDlgAutoProxy, CCmdTarget)
	INTERFACE_PART(COleappDlgAutoProxy, IID_IOleapp, Dispatch)
END_INTERFACE_MAP()

// The IMPLEMENT_OLECREATE2 macro is defined in StdAfx.h of this project
// {647950C3-5221-11D2-81CA-0000A00F0ECE}
IMPLEMENT_OLECREATE2(COleappDlgAutoProxy, "Oleapp.Application", 0x647950c3, 0x5221, 0x11d2, 0x81, 0xca, 0x0, 0x0, 0xa0, 0xf, 0xe, 0xce)

/////////////////////////////////////////////////////////////////////////////
// COleappDlgAutoProxy message handlers
