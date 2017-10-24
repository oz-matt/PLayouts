/* Copyright Mentor Graphics Corporation 2007

All Rights Reserved.

THIS WORK CONTAINS TRADE SECRET
AND PROPRIETARY INFORMATION WHICH IS THE
PROPERTY OF MENTOR GRAPHICS
CORPORATION OR ITS LICENSORS AND IS
SUBJECT TO LICENSE TERMS. 
*/

// oleapp.h : main header file for the OLEAPP application
//

#if !defined(AFX_OLEAPP_H__647950C8_5221_11D2_81CA_0000A00F0ECE__INCLUDED_)
#define AFX_OLEAPP_H__647950C8_5221_11D2_81CA_0000A00F0ECE__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// COleappApp:
// See oleapp.cpp for the implementation of this class
//

class COleappApp : public CWinApp
{
public:
	COleappApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(COleappApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(COleappApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_OLEAPP_H__647950C8_5221_11D2_81CA_0000A00F0ECE__INCLUDED_)
