//////////////////////////////////////////////////////////////////////////////
//
// CLIENTDLG.H : header file
//
//////////////////////////////////////////////////////////////////////////////
// This is a part of the PADS-PowerPCB OLE Automation server SAMPLE2 sample.
// Copyright (C) 2003 Mentor Graphics Corp.
// All rights reserved.
//
// This source code is only intended as a supplement to the PADS-PowerPCB OLE 
// Automation Server API Help file.
//////////////////////////////////////////////////////////////////////////////

#if !defined(AFX_CLIENTDLG_H__EFD4272A_D443_11D0_BCBF_444553540000__INCLUDED_)
#define AFX_CLIENTDLG_H__EFD4272A_D443_11D0_BCBF_444553540000__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

class CClientDlgAutoProxy;

/////////////////////////////////////////////////////////////////////////////
// CClientDlg dialog
class IPowerPCBApp;
class CClientDlg : public CDialog
{
	DECLARE_DYNAMIC(CClientDlg);
	friend class CClientDlgAutoProxy;

// Construction
public:
	CClientDlg(CWnd* pParent = NULL);	// standard constructor
	virtual ~CClientDlg();

	void Refresh();
	void RefreshGeneralInfos();
	void RefreshObjectsInfos();
	void Disconnect();

// Dialog Data
	//{{AFX_DATA(CClientDlg)
	enum { IDD = IDD_CLIENT_DIALOG };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CClientDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	CClientDlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	// Pointer to the PowerPCB server's top level object
	IPowerPCBApp *m_pPowerPCBApplication;
	PowerPCBSink m_PowerPCBSink;
	BOOL SetSink(BOOL bStatus);
	DWORD m_dwAppConnectionID;
	DWORD m_dwDocConnectionID;

	BOOL CanExit();

	// Generated message map functions
	//{{AFX_MSG(CClientDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnConnect();
	afx_msg void OnDisconnect();
	afx_msg void OnRefresh();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_CLIENTDLG_H__EFD4272A_D443_11D0_BCBF_444553540000__INCLUDED_)
