//////////////////////////////////////////////////////////////////////////////////////
//
// ChildView.cpp : implementation of the CChildView class
//
//////////////////////////////////////////////////////////////////////////////////////
// This is a part of the Innoveda-PowerPCB OLE Automation server PPCBAutoGeom sample.
// Copyright (C) 2003 Mentor Graphics Corp.
// All rights reserved.
//
// This source code is only intended as a supplement to the Innoveda-PowerPCB OLE
// Automation Server API Help file.
//////////////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "PPCBAutoGeom.h"
#include "ChildView.h"
#include "Picture.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CChildView

CChildView::CChildView()
{
}

CChildView::~CChildView()
{
}


BEGIN_MESSAGE_MAP(CChildView,CWnd)
	//{{AFX_MSG_MAP(CChildView)
	ON_COMMAND(ID_VIEW_REFRESH, OnViewRefresh)
	ON_WM_LBUTTONDBLCLK()
	ON_WM_PAINT()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// CChildView message handlers

BOOL CChildView::PreCreateWindow(CREATESTRUCT& cs) 
{
	if (!CWnd::PreCreateWindow(cs))
		return FALSE;

	cs.dwExStyle |= WS_EX_CLIENTEDGE;
	cs.style &= ~WS_BORDER;
	cs.lpszClass = AfxRegisterWndClass(CS_HREDRAW|CS_VREDRAW|CS_DBLCLKS,
		::LoadCursor(NULL, IDC_ARROW), HBRUSH(COLOR_WINDOW+1), NULL);

	return TRUE;
}

void CChildView::OnPaint()
{
	CPaintDC dc(this); // device context for painting

	// TODO: Add your message handler code here

	AfxGetPicture()->Draw(dc);

	// Do not call CWnd::OnPaint() for painting messages
}

void CChildView::OnViewRefresh() 
{
	// TODO: Add your command handler code here
	((CPPCBAutoGeomApp *)AfxGetApp())->SetStatusBarText(IDS_LOADDATAMESSAGE);
	{
		CWaitCursor cwc;
		AfxGetPicture()->Refresh();
	}
	((CPPCBAutoGeomApp *)AfxGetApp())->SetStatusBarText(AFX_IDS_IDLEMESSAGE);

	Invalidate();
}

void CChildView::OnLButtonDblClk(UINT nFlags, CPoint point) 
{
	// TODO: Add your message handler code here and/or call default
	AfxGetMainWnd()->SendMessage(WM_COMMAND,MAKEWPARAM(ID_VIEW_REFRESH,0),NULL);
}
