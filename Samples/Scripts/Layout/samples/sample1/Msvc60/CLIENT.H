//////////////////////////////////////////////////////////////////////////////
//
// CLIENT.H : Main header file for the CLIENT application
//
//////////////////////////////////////////////////////////////////////////////
// This is a part of the PADS-PowerPCB OLE Automation server SAMPLE1 sample.
// Copyright (C) 2003 Mentor Graphics Corp.
// All rights reserved.
//
// This source code is only intended as a supplement to the PADS-PowerPCB OLE 
// Automation Server API Help file.
//////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_CLIENT_H__EFD42728_D443_11D0_BCBF_444553540000__INCLUDED_)
#define AFX_CLIENT_H__EFD42728_D443_11D0_BCBF_444553540000__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CClientApp:
// See client.cpp for the implementation of this class
//

class CClientApp : public CWinApp
{
public:
	CClientApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CClientApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CClientApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CLIENT_H__EFD42728_D443_11D0_BCBF_444553540000__INCLUDED_)
