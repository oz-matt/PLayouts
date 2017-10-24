/* Copyright Mentor Graphics Corporation 2007

All Rights Reserved.

THIS WORK CONTAINS TRADE SECRET
AND PROPRIETARY INFORMATION WHICH IS THE
PROPERTY OF MENTOR GRAPHICS
CORPORATION OR ITS LICENSORS AND IS
SUBJECT TO LICENSE TERMS. 
*/

// oleappDlg.h : header file
//

#if !defined(AFX_OLEAPPDLG_H__647950CA_5221_11D2_81CA_0000A00F0ECE__INCLUDED_)
#define AFX_OLEAPPDLG_H__647950CA_5221_11D2_81CA_0000A00F0ECE__INCLUDED_

#include "sink.h"

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

class COleappDlgAutoProxy;

/////////////////////////////////////////////////////////////////////////////
// COleappDlg dialog

class IPowerLogicApp;
class PowerLogicSink;
class COleappDlg : public CDialog
{
	DECLARE_DYNAMIC(COleappDlg);
	friend class COleappDlgAutoProxy;

// Construction
public:
	COleappDlg(CWnd* pParent = NULL);	// standard constructor
	virtual ~COleappDlg();

	void Refresh();
	void RefreshGeneralInfos();
	void RefreshObjectsInfos();
	void RefreshSelectionInfos();
	void ProcessException(COleException* e);
	void Connect();
	void SendSelection();
	void ShowObjectInfo();
	void Disconnect();
	void SafeDisconnect();
	void LocateSelection();

	// Dialog Data
	//{{AFX_DATA(COleappDlg)
	enum { IDD = IDD_OLEAPP_DIALOG };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(COleappDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	COleappDlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	IPowerLogicApp *m_pPowerLogicApplication;
	PowerLogicSink m_PowerLogicSink;
	BOOL SetSink(BOOL bStatus);
	DWORD m_dwAppConnectionID;
	DWORD m_dwDocConnectionID;

	void OnConnect();
	void OnDisconnect();


	BOOL CanExit();

	// Generated message map functions
	//{{AFX_MSG(COleappDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	afx_msg void OnBtnConnect();
	afx_msg void OnBtnDisconnect();
	afx_msg void OnBtnRefresh();
	afx_msg void OnDblclkLstDocobjects();
	afx_msg void OnSelchangeLstDocobjects();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_OLEAPPDLG_H__647950CA_5221_11D2_81CA_0000A00F0ECE__INCLUDED_)
