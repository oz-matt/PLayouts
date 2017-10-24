/* Copyright Mentor Graphics Corporation 2007

All Rights Reserved.

THIS WORK CONTAINS TRADE SECRET
AND PROPRIETARY INFORMATION WHICH IS THE
PROPERTY OF MENTOR GRAPHICS
CORPORATION OR ITS LICENSORS AND IS
SUBJECT TO LICENSE TERMS. 
*/

// DlgProxy.h : header file
//

#if !defined(AFX_DLGPROXY_H__647950CC_5221_11D2_81CA_0000A00F0ECE__INCLUDED_)
#define AFX_DLGPROXY_H__647950CC_5221_11D2_81CA_0000A00F0ECE__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

class COleappDlg;

/////////////////////////////////////////////////////////////////////////////
// COleappDlgAutoProxy command target

class COleappDlgAutoProxy : public CCmdTarget
{
	DECLARE_DYNCREATE(COleappDlgAutoProxy)

	COleappDlgAutoProxy();           // protected constructor used by dynamic creation

// Attributes
public:
	COleappDlg* m_pDialog;

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(COleappDlgAutoProxy)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:
	virtual ~COleappDlgAutoProxy();

	// Generated message map functions
	//{{AFX_MSG(COleappDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	DECLARE_OLECREATE(COleappDlgAutoProxy)

	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(COleappDlgAutoProxy)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_DLGPROXY_H__647950CC_5221_11D2_81CA_0000A00F0ECE__INCLUDED_)
