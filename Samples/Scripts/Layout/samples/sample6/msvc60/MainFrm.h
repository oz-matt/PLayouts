//////////////////////////////////////////////////////////////////////////////////////
//
// MainFrm.h : interface of the CMainFrame class
//
//////////////////////////////////////////////////////////////////////////////////////
// This is a part of the Innoveda-PowerPCB OLE Automation server PPCBAutoGeom sample.
// Copyright (C) 2003 Mentor Graphics Corp.
// All rights reserved.
//
// This source code is only intended as a supplement to the Innoveda-PowerPCB OLE
// Automation Server API Help file.
//////////////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_MAINFRM_H__CD126FFC_65F2_4F2E_8FAC_8B8F9A1DECC1__INCLUDED_)
#define AFX_MAINFRM_H__CD126FFC_65F2_4F2E_8FAC_8B8F9A1DECC1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ChildView.h"

class CMainFrame : public CFrameWnd
{
	
public:
	CMainFrame();
protected: 
	DECLARE_DYNAMIC(CMainFrame)

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMainFrame)
	virtual BOOL PreCreateWindow(CREATESTRUCT& cs);
	virtual BOOL OnCmdMsg(UINT nID, int nCode, void* pExtra, AFX_CMDHANDLERINFO* pHandlerInfo);
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CMainFrame();
#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif

	void SetStatusBarText(UINT strId )
	{
		CString str;
		str.LoadString( strId );
		m_wndStatusBar.SetWindowText(str);
	}

	CWnd &GetViewWnd() { return m_wndView; }

protected:  // control bar embedded members
	CStatusBar  m_wndStatusBar;
	CChildView    m_wndView;

// Generated message map functions
protected:
	//{{AFX_MSG(CMainFrame)
	afx_msg int OnCreate(LPCREATESTRUCT lpCreateStruct);
	afx_msg void OnSetFocus(CWnd *pOldWnd);
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MAINFRM_H__CD126FFC_65F2_4F2E_8FAC_8B8F9A1DECC1__INCLUDED_)
