//////////////////////////////////////////////////////////////////////////////
//
// CLIENTDLG.CPP : implementation file
//
//////////////////////////////////////////////////////////////////////////////
// This is a part of the PADS-PowerPCB OLE Automation server SAMPLE1 sample.
// Copyright (C) 2003 Mentor Graphics Corp.
// All rights reserved.
//
// This source code is only intended as a supplement to the PADS-PowerPCB OLE 
// Automation Server API Help file.
//////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "client.h"
#include "clientDl.h"
#include "DlgProxy.h"
#include "powerpcb.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
	//{{AFX_MSG_MAP(CAboutDlg)
		// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CClientDlg dialog

IMPLEMENT_DYNAMIC(CClientDlg, CDialog);

CClientDlg::CClientDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CClientDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CClientDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = NULL;
	m_pPowerPCBApplication = NULL;
}

CClientDlg::~CClientDlg()
{
	// If there is an automation proxy for this dialog, set
	//  its back pointer to this dialog to NULL, so it knows
	//  the dialog has been deleted.
	if (m_pAutoProxy != NULL)
		m_pAutoProxy->m_pDialog = NULL;
}

void CClientDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CClientDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CClientDlg, CDialog)
	//{{AFX_MSG_MAP(CClientDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_CLOSE()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BTN_CONNECT, OnConnect)
	ON_BN_CLICKED(IDC_BTN_DISCONNECT, OnDisconnect)
	ON_BN_CLICKED(IDC_BTN_REFRESH, OnRefresh)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CClientDlg message handlers

BOOL CClientDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	SetDlgItemText(IDC_TXT_STATUS, "Not connected to PowerPCB.");
	((CButton *) GetDlgItem(IDC_BTN_CONNECT))->EnableWindow(TRUE);
	((CButton *) GetDlgItem(IDC_BTN_DISCONNECT))->EnableWindow(FALSE);
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CClientDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CClientDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CClientDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

// Automation servers should not exit when a user closes the UI
//  if a controller still holds on to one of its objects.  These
//  message handlers make sure that if the proxy is still in use,
//  then the UI is hidden but the dialog remains around if it
//  is dismissed.

void CClientDlg::OnClose() 
{
	if (CanExit())
		CDialog::OnClose();
}

void CClientDlg::OnOK() 
{
	if (CanExit())
		CDialog::OnOK();
}

void CClientDlg::OnCancel() 
{
	if (CanExit())
		CDialog::OnCancel();
}

BOOL CClientDlg::CanExit()
{
	// If the proxy object is still around, then the automation
	//  controller is still holding on to this application.  Leave
	//  the dialog around, but hide its UI.
	if (m_pAutoProxy != NULL)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	OnConnect()
// Desc:		Connect button callback. Connects to PowerPCB Server.
/////////////////////////////////////////////////////////////////////////////
void CClientDlg::OnConnect() 
{
	if (m_pPowerPCBApplication != NULL) { ASSERT(FALSE); return; }

 	BOOL rc;
	HRESULT hResult;

	SetDlgItemText(IDC_TXT_STATUS, "Connecting to PowerPCB server...");

	/////////////////////////////////////////////////////////////////////////////
	// Get the CLSID from the PowerPCB name string ID (system registry query)
	CLSID clsID;
	hResult = ::CLSIDFromProgID(OLESTR("PowerPCB.Application"), &clsID);
	if (hResult != NOERROR) { ASSERT(FALSE); return; }

	/////////////////////////////////////////////////////////////////////////////
	// Try to get an active instance of PowerPCB. The ::GetActiveObject() function
	// will query the Running Object Table and search for a registered PowerPCB
	// instance currently running on the system. If the return value of ::GetActiveObject()
	// is not S_OK, it means that the server is not running on the system at this time.
	LPUNKNOWN pUnknown;
	SetDlgItemText(IDC_TXT_STATUS, "Querying Running Object Table...");
	hResult = ::GetActiveObject(clsID, NULL, &pUnknown);

	/////////////////////////////////////////////////////////////////////////////
	// If such an instance was found, connect to it.
	if (hResult == S_OK) 
	{
		BeginWaitCursor();
		SetDlgItemText(IDC_TXT_STATUS, "Attaching to running PowerPCB server...");

		// Get IDispatch interface of the PowerPCB application object.
		LPDISPATCH pDispatch = NULL;
		hResult = pUnknown->QueryInterface(IID_IDispatch, (void **) &pDispatch);
		ASSERT(hResult == S_OK);

		// Allocate the interface
		m_pPowerPCBApplication = (IPowerPCBApp *) new IPowerPCBApp;
		ASSERT(m_pPowerPCBApplication != NULL);

		// Attach this dispatch pointer our pPowerPCBApp interface pointer.
		m_pPowerPCBApplication->AttachDispatch(pDispatch, FALSE);

		// Release the pUnknown we have used.
		pUnknown->Release();	   

		EndWaitCursor();
	}
	/////////////////////////////////////////////////////////////////////////////
	// If no such instance was found, start a new one and connect to it.
	else
	{
		BeginWaitCursor();
		SetDlgItemText(IDC_TXT_STATUS, "Starting and connecting to PowerPCB server...");

		// Allocate the interface
		m_pPowerPCBApplication = (IPowerPCBApp *) new IPowerPCBApp;
		ASSERT(m_pPowerPCBApplication != NULL);

		// Create pPowerPCBApp dispatch
		COleException e;

		////////////////////////////////////////////////////////////////////////
		// The CreateDispatch() function is very powerful because it encapsulate
		//	everything to establich the connection. However, if the function fails,
		// or if the system gets stuck inside this function, it is very hard to 
		// debug and diagnostic what went wrong. The #else section contains the 
		// expanded code of CreateDispatch() for easier debugging in case of a problem.
		////////////////////////////////////////////////////////////////////////
#if 1
		rc = m_pPowerPCBApplication->CreateDispatch(clsID, &e);

#else
		ASSERT(m_pPowerPCBApplication->m_lpDispatch == NULL);

		m_pPowerPCBApplication->m_bAutoRelease = TRUE;  // Good default is to auto-release

		// Create an instance of the object
		LPUNKNOWN lpUnknown = NULL;
		IClassFactory *pCF;
		hResult = CoGetClassObject(clsID, CLSCTX_ALL, NULL, IID_IClassFactory, (void **)&pCF);
		if (hResult != S_OK) { ASSERT(FALSE); return; }

		hResult = pCF->CreateInstance(NULL, IID_IUnknown, (void **)&lpUnknown);
		if (hResult != S_OK) { ASSERT(FALSE); return; }
		pCF->Release();

		// Make sure it is running
		hResult = OleRun(lpUnknown);
		if (hResult != S_OK) { ASSERT(FALSE); lpUnknown->Release(); e.m_sc = hResult; return; }

		// Query for IDispatch interface
		hResult = lpUnknown->QueryInterface(IID_IDispatch, (void **) &(m_pPowerPCBApplication->m_lpDispatch));
		if (hResult != S_OK) { ASSERT(FALSE); lpUnknown->Release(); e.m_sc = hResult; return; }

		lpUnknown->Release();
		ASSERT(m_pPowerPCBApplication->m_lpDispatch != NULL);

#endif

		EndWaitCursor();

		if (rc == FALSE) 
		{
			ASSERT(FALSE);	
			delete m_pPowerPCBApplication; m_pPowerPCBApplication = NULL;
			SetDlgItemText(IDC_TXT_STATUS, "Not connected to PowerPCB.");
			return;
		}
	}
	
	((CButton *) GetDlgItem(IDC_BTN_CONNECT))->EnableWindow(FALSE);
	((CButton *) GetDlgItem(IDC_BTN_DISCONNECT))->EnableWindow(TRUE);
	SetDlgItemText(IDC_TXT_STATUS, "Connected to PowerPCB.");

	m_pPowerPCBApplication->SetVisible(TRUE);
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	OnDisconnect()
// Desc:		Disconnect button callback. Disconnects from PowerPCB Server.
/////////////////////////////////////////////////////////////////////////////
void CClientDlg::OnDisconnect() 
{
	if (m_pPowerPCBApplication == NULL) { ASSERT(FALSE); return; }

	// Delete the connection with PowerPCB by deleting it's reference.
	delete m_pPowerPCBApplication; 
	m_pPowerPCBApplication = NULL;

	SetDlgItemText(IDC_TXT_STATUS, "Not connected to PowerPCB.");

	((CButton *) GetDlgItem(IDC_BTN_CONNECT))->EnableWindow(TRUE);
	((CButton *) GetDlgItem(IDC_BTN_DISCONNECT))->EnableWindow(FALSE);
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	OnRefresh()
// Desc:		Refresh button callback. Get data from PowerPCB server.
/////////////////////////////////////////////////////////////////////////////
void CClientDlg::OnRefresh() 
{
	if (m_pPowerPCBApplication == NULL)
	{
		// Not connected to server -> print dummy strings
		SetDlgItemText(IDC_EDT_DOCNAME, "???????");
	}
	else
	{
		// Call the server Application ActiveDocument property which returns
		// a pointer to its IDispatch interface. PowerPCB garantees that the
		// returned pointed object implements the IPowerPCBDoc interface as
		// well, so we attach this IDispatch pointer to a IPowerPCBDoc object.
		LPDISPATCH pDispatch = m_pPowerPCBApplication->GetActiveDocument();
		ASSERT(pDispatch);
		IPowerPCBDoc doc; doc.AttachDispatch(pDispatch);

		// Call the server Document Name property and outpu its value.
		CString docname = doc.GetName();
		SetDlgItemText(IDC_EDT_DOCNAME, docname.GetBuffer(0));
	}
}
