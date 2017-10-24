//////////////////////////////////////////////////////////////////////////////////////
//
// PPCBAutoGeom.h : main header file for the PPCBAUTOGEOM application
//
//////////////////////////////////////////////////////////////////////////////////////
// This is a part of the Innoveda-PowerPCB OLE Automation server PPCBAutoGeom sample.
// Copyright (C) 2003 Mentor Graphics Corp.
// All rights reserved.
//
// This source code is only intended as a supplement to the Innoveda-PowerPCB OLE
// Automation Server API Help file.
//////////////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_PPCBAUTOGEOM_H__AE78D51B_4830_4F5A_80B1_2EB13939737E__INCLUDED_)
#define AFX_PPCBAUTOGEOM_H__AE78D51B_4830_4F5A_80B1_2EB13939737E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"       // main symbols

#include "MainFrm.h"

/////////////////////////////////////////////////////////////////////////////
// CPPCBAutoGeomApp:
// See PPCBAutoGeom.cpp for the implementation of this class
//

class CPPCBAutoGeomApp : public CWinApp
{
public:
	CPPCBAutoGeomApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CPPCBAutoGeomApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	void SetStatusBarText(UINT strId ) { ((CMainFrame *)m_pMainWnd)->SetStatusBarText(strId); }

public:
	//{{AFX_MSG(CPPCBAutoGeomApp)
	afx_msg void OnAppAbout();
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_PPCBAUTOGEOM_H__AE78D51B_4830_4F5A_80B1_2EB13939737E__INCLUDED_)
