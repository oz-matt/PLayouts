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
// SINK.H : header file
//
//////////////////////////////////////////////////////////////////////////////
// This is a part of the PADS Logic OLE Automation server SAMPLE2 sample.
// Copyright Mentor Graphics Corporation 2007
// All rights reserved.
//
// This source code is only intended as a supplement to the PADS Logic OLE 
// Automation Server API Help file.
//////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_SINK_H__667FDFC2_D446_11D0_BCBF_444553540000__INCLUDED_)
#define AFX_SINK_H__667FDFC2_D446_11D0_BCBF_444553540000__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

/////////////////////////////////////////////////////////////////////////////
// PowerLogicSink command target

class PowerLogicSink : public CCmdTarget
{
	//DECLARE_DYNCREATE(PowerLogicSink)

	//PowerLogicSink();           // protected constructor used by dynamic creation
	//virtual ~PowerLogicSink();

// Attributes
public:
	PowerLogicSink();           // public constructor used by dynamic creation
	~PowerLogicSink();

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(PowerLogicSink)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(PowerLogicSink)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(PowerLogicSink)
	afx_msg void OnOpenDocument(LPDISPATCH pDoc);
	afx_msg void OnSelectionChange();
	afx_msg void OnQuit();
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()
	DECLARE_INTERFACE_MAP()
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_SINK_H__667FDFC2_D446_11D0_BCBF_444553540000__INCLUDED_)
