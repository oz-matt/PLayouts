//////////////////////////////////////////////////////////////////////////////
//
// SINK.H : header file
//
//////////////////////////////////////////////////////////////////////////////
// This is a part of the PADS-PowerPCB OLE Automation server SAMPLE2 sample.
// Copyright (C) 2003 Mentor Graphics Corp.
// All rights reserved.
//
// This source code is only intended as a supplement to the PADS-PowerPCB OLE 
// Automation Server API Help file.
//////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_SINK_H__667FDFC2_D446_11D0_BCBF_444553540000__INCLUDED_)
#define AFX_SINK_H__667FDFC2_D446_11D0_BCBF_444553540000__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

/////////////////////////////////////////////////////////////////////////////
// PowerPCBSink command target

class PowerPCBSink : public CCmdTarget
{
	//DECLARE_DYNCREATE(PowerPCBSink)

	//PowerPCBSink();           // protected constructor used by dynamic creation
	//virtual ~PowerPCBSink();

// Attributes
public:
	PowerPCBSink();           // public constructor used by dynamic creation
	~PowerPCBSink();

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(PowerPCBSink)
	public:
	virtual void OnFinalRelease();
	//}}AFX_VIRTUAL

// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(PowerPCBSink)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	// Generated OLE dispatch map functions
	//{{AFX_DISPATCH(PowerPCBSink)
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
