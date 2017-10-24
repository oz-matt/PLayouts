//////////////////////////////////////////////////////////////////////////////
//
// CLIENTDLG.CPP : implementation file
//
//////////////////////////////////////////////////////////////////////////////
// This is a part of the PADS-PowerPCB OLE Automation server SAMPLE2 sample.
// Copyright (C) 2003 Mentor Graphics Corp.
// All rights reserved.
//
// This source code is only intended as a supplement to the PADS-PowerPCB OLE 
// Automation Server API Help file.
//////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "client.h"
#include "sink.h"
#include "clientDl.h"
#include "DlgProxy.h"
#include "powerpcb.h"
#include "powerpcb.inc"

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
	m_dwAppConnectionID		= 0;
	m_dwDocConnectionID		= 0;
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
	ON_LBN_SELCHANGE(IDC_LST_DOCOBJECTS, OnListBoxSelectionChange)
	ON_LBN_DBLCLK(IDC_LST_DOCOBJECTS, OnListBoxDoubleClick)
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

void CClientDlg::Connect() 
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
	
	// Connect sink
	SetDlgItemText(IDC_TXT_STATUS, "Connecting sink...");
	SetSink(TRUE);

	((CButton *) GetDlgItem(IDC_BTN_CONNECT))->EnableWindow(FALSE);
	((CButton *) GetDlgItem(IDC_BTN_DISCONNECT))->EnableWindow(TRUE);
	SetDlgItemText(IDC_TXT_STATUS, "Connected to PowerPCB.");

	m_pPowerPCBApplication->SetVisible(TRUE);
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	OnConnect()
// Desc:		Connect button callback. Connects to PowerPCB Server.
/////////////////////////////////////////////////////////////////////////////
void CClientDlg::OnConnect() 
{
	try {
		Connect();
	}
	catch (COleException* e) {
		ProcessException(e);
	}
}

//ProcessException
void CClientDlg::ProcessException(COleException* e)
{
	LPTSTR   szMessage;
	HRESULT hr = e->m_sc;

    if (hr == S_OK) return;

    if (HRESULT_FACILITY(hr) == FACILITY_WINDOWS)
        hr = HRESULT_CODE(hr);

    FormatMessage(
        FORMAT_MESSAGE_ALLOCATE_BUFFER |
        FORMAT_MESSAGE_FROM_SYSTEM,
        NULL,
        hr,
        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), //The user default language
        (LPTSTR)&szMessage,
        0,
        NULL);

	if ((hr & 0xFFFF) == RPC_S_SERVER_UNAVAILABLE) {
		//probably PowerPCB is crashed or ubnormal terminated
		delete m_pPowerPCBApplication; 
		m_pPowerPCBApplication = NULL;
		SetDlgItemText(IDC_TXT_STATUS, "Not connected to PowerPCB.");
		((CButton *) GetDlgItem(IDC_BTN_CONNECT))->EnableWindow(TRUE);
		((CButton *) GetDlgItem(IDC_BTN_DISCONNECT))->EnableWindow(FALSE);
	}

	CString str;
	str.Format("OLE exception: %s (%lx)", szMessage, hr);
	AfxMessageBox(str);
	LocalFree(szMessage);
}

void CClientDlg::SafeDisconnect()
{
	try {
		Disconnect();
	}
	catch (COleException* e) {
		ProcessException(e);
	}
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	Disconnect()
// Desc:		Disconnects from PowerPCB Server.
/////////////////////////////////////////////////////////////////////////////
void CClientDlg::Disconnect()
{
	if (m_pPowerPCBApplication == NULL) { ASSERT(FALSE); return; }

	// Disconnect sink
	SetSink(FALSE);

	// Delete the connection with PowerPCB by deleting it's reference.
	delete m_pPowerPCBApplication; 
	m_pPowerPCBApplication = NULL;

	SetDlgItemText(IDC_TXT_STATUS, "Not connected to PowerPCB.");

	((CButton *) GetDlgItem(IDC_BTN_CONNECT))->EnableWindow(TRUE);
	((CButton *) GetDlgItem(IDC_BTN_DISCONNECT))->EnableWindow(FALSE);
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	OnDisconnect()
// Desc:		Disconnect button callback. 
/////////////////////////////////////////////////////////////////////////////
void CClientDlg::OnDisconnect() 
{
	SafeDisconnect();
}

BOOL CClientDlg::SetSink(BOOL bStatus)
{
	HRESULT hr;
	ASSERT (m_pPowerPCBApplication->m_lpDispatch != NULL);
	IPowerPCBDoc doc; doc.AttachDispatch(m_pPowerPCBApplication->GetActiveDocument());

	LPCONNECTIONPOINTCONTAINER pAppConnPtCont = NULL;
	LPCONNECTIONPOINTCONTAINER pDocConnPtCont = NULL;
	hr = m_pPowerPCBApplication->m_lpDispatch->QueryInterface(IID_IConnectionPointContainer, (LPVOID*)&pAppConnPtCont);
	if (hr != S_OK) return FALSE;

	hr = doc.m_lpDispatch->QueryInterface(IID_IConnectionPointContainer, (LPVOID*)&pDocConnPtCont);
	if (hr != S_OK) { doc.ReleaseDispatch(); return FALSE; }
	

	LPCONNECTIONPOINT pAppConnPt = NULL;
	hr = pAppConnPtCont->FindConnectionPoint(DIID__PowerPCBAppEvents, &pAppConnPt);
	if (hr != S_OK) return FALSE;

	LPCONNECTIONPOINT pDocConnPt = NULL;
	hr = pDocConnPtCont->FindConnectionPoint(DIID__PowerPCBDocEvents, &pDocConnPt);
	if (hr != S_OK) { doc.ReleaseDispatch(); return FALSE; }
		
	LPDISPATCH lpDispatch = m_PowerPCBSink.GetIDispatch(FALSE); ASSERT(lpDispatch);
	if (bStatus == TRUE) {
		hr = pAppConnPt->Advise(lpDispatch, &m_dwAppConnectionID);
		hr = pDocConnPt->Advise(lpDispatch, &m_dwDocConnectionID);
		pDocConnPtCont->AddRef();
	} else {
		pDocConnPtCont->Release();
		hr = pAppConnPt->Unadvise(m_dwAppConnectionID);
		hr = pDocConnPt->Unadvise(m_dwDocConnectionID);
		m_dwAppConnectionID = 0; m_dwDocConnectionID = 0;
	}
	pAppConnPt->Release();
	pDocConnPt->Release();
		
	pAppConnPtCont->Release();
	pDocConnPtCont->Release();
	return SUCCEEDED(hr);
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	OnRefresh()
// Desc:		Refresh button callback. Get data from PowerPCB server.
/////////////////////////////////////////////////////////////////////////////
void CClientDlg::OnRefresh() 
{
	try {
		// We know that the RefreshObjectsInfos() function performs a lengthy
		// operation, which is mainly OLE calls to PowerPCB server. We will lock
		// the server so that these OLE calls are very fast.
		if (m_pPowerPCBApplication) m_pPowerPCBApplication->LockServer();
		RefreshGeneralInfos();
		RefreshObjectsInfos();
		RefreshSelectionInfos();
		if (m_pPowerPCBApplication) m_pPowerPCBApplication->UnlockServer();
	}
	catch (COleException* e) {
		ProcessException(e);
	}
}

void CClientDlg::Refresh()
{
	// We know that the RefreshObjectsInfos() function performs a lengthy
	// operation, which is mainly OLE calls to PowerPCB server. We will lock
	// the server so that these OLE calls are very fast.
	if (m_pPowerPCBApplication) m_pPowerPCBApplication->LockServer();
	OnRefresh();
	if (m_pPowerPCBApplication) m_pPowerPCBApplication->UnlockServer();
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	RefreshGeneralInfos()
// Desc:		Get data from PowerPCB server (fast access data)
/////////////////////////////////////////////////////////////////////////////
void CClientDlg::RefreshGeneralInfos()
{
	if (m_pPowerPCBApplication == NULL)
	{
		// Not connected to server -> print dummy strings
		SetDlgItemText(IDC_EDT_DOCNAME, "???????");
		SetDlgItemText(IDC_EDT_COMPS, "???????");
		SetDlgItemText(IDC_EDT_SELECTEDCOMPS, "???????");
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
		SetDlgItemText(IDC_EDT_DOCNAME, docname);

		// Get the number of components
		pDispatch = doc.GetComponents();
		ASSERT(pDispatch);
		IPowerPCBObjs objs1; objs1.AttachDispatch(pDispatch);
		char selCountStr[255];
		sprintf(selCountStr, "%i component object(s)", objs1.GetCount());
		SetDlgItemText(IDC_EDT_COMPS, selCountStr);

		// Get the number of selected components
		pDispatch = doc.GetObjects(ppcbObjectTypeComponent, NULL, TRUE);
		ASSERT(pDispatch);
		IPowerPCBObjs objs2; objs2.AttachDispatch(pDispatch);
		sprintf(selCountStr, "%i selected component object(s)", objs2.GetCount());
		SetDlgItemText(IDC_EDT_SELECTEDCOMPS, selCountStr);
	}
	SetDlgItemText(IDC_TXT_STATUS, "");
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	RefreshObjectsInfos()
// Desc:		Get data from PowerPCB server (slow access data)
/////////////////////////////////////////////////////////////////////////////
void CClientDlg::RefreshObjectsInfos()
{
	if (m_pPowerPCBApplication == NULL)
	{
		CListBox *listbox = (CListBox *)GetDlgItem(IDC_LST_DOCOBJECTS);
		listbox->ResetContent();
	}
	else
	{
		BeginWaitCursor();

		// Call the server Application ActiveDocument property which returns
		// a pointer to its IDispatch interface. PowerPCB garantees that the
		// returned pointed object implements the IPowerPCBDoc interface as
		// well, so we attach this IDispatch pointer to a IPowerPCBDoc object.
		LPDISPATCH pDispatch = m_pPowerPCBApplication->GetActiveDocument();
		ASSERT(pDispatch);
		IPowerPCBDoc doc; doc.AttachDispatch(pDispatch);

		// Reset our listbox
		CListBox *listbox = (CListBox *)GetDlgItem(IDC_LST_DOCOBJECTS);
		listbox->ResetContent();

		LPDISPATCH pObjs;
		
		// Get all PowerPCB objects of type COMPONENT from PowerPCB database
		// and output each component in the listbox
		int countComp = 0;
		pObjs = doc.GetComponents(); ASSERT(pObjs != NULL);
		IPowerPCBObjs comps; comps.AttachDispatch(pObjs);
		// For each component returned by PowerPCB...
		for (int i=1; i<=comps.GetCount(); i++)
		{
			// ... get the component object itself...
			COleVariant v((long)i);
			LPDISPATCH pObj = comps.GetItem(v); ASSERT(pObj != NULL);
			IPowerPCBComp comp; comp.AttachDispatch(pObj);
					
			// ... extract its name and add it to the listbox.
			int index = listbox->AddString(comp.GetName());
			countComp++; char msg[255]; sprintf(msg, "Loading %i component(s) from PowerPCB...", countComp);
			SetDlgItemText(IDC_TXT_STATUS, msg);
		}

		EndWaitCursor();
	
	}
	SetDlgItemText(IDC_TXT_STATUS, "");
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	RefreshSelectionInfos()
// Desc:		Get selection data from PowerPCB server
/////////////////////////////////////////////////////////////////////////////
void CClientDlg::RefreshSelectionInfos()
{
	if (m_pPowerPCBApplication == NULL) return;

	BeginWaitCursor();

	LPDISPATCH pDispatch = NULL;
	pDispatch= m_pPowerPCBApplication->GetActiveDocument(); ASSERT(pDispatch != NULL);
	IPowerPCBDoc doc; doc.AttachDispatch(pDispatch);

	// Reset listbox selection
	CListBox *listbox = (CListBox *)GetDlgItem(IDC_LST_DOCOBJECTS);
	BOOL rc = listbox->SetSel(-1, FALSE); ASSERT(rc != LB_ERR);

	// Get the selected COMPONENTS
	pDispatch = doc.GetObjects(ppcbObjectTypeComponent, NULL, TRUE); ASSERT(pDispatch != NULL);
	IPowerPCBObjs comps; comps.AttachDispatch(pDispatch);
	for (int i=1; i<=comps.GetCount(); i++) {
		COleVariant v((long)i);
		pDispatch = comps.GetItem(v); ASSERT(pDispatch != NULL);
		IPowerPCBComp comp; comp.AttachDispatch(pDispatch);
		// Find this component in the listbox.
		int index = listbox->FindString(0, comp.GetName()); // ASSERT(index != LB_ERR);
		// Select it in the listbox.
		rc = listbox->SetSel(index, TRUE); ASSERT(rc != LB_ERR);
	}
		
	EndWaitCursor();
} // RefreshSelectionInfos()


/////////////////////////////////////////////////////////////////////////////
// Routine:	SendSelection()
// Desc:		This function retrieves
//				what has been selected in the listbox and selects the same thing in 
//				PowerPCB. It demonstrate the SelectObjects() routine as well as
//				PowerPCB object collections manipulation.
/////////////////////////////////////////////////////////////////////////////
void CClientDlg::SendSelection() 
{
	if (m_pPowerPCBApplication == NULL) return;
	
	BeginWaitCursor();

	
	LPDISPATCH pDispatch = NULL;

	// Get the number of items selected in the listbox. If something is selected, process...
	CListBox *listbox = (CListBox *)GetDlgItem(IDC_LST_DOCOBJECTS);
	int count = listbox->GetSelCount();
	if (count >= 0)
	{
		pDispatch = m_pPowerPCBApplication->GetActiveDocument(); ASSERT(pDispatch != NULL);
		IPowerPCBDoc doc; doc.AttachDispatch(pDispatch);

		///////////////////////////////////////////////////////////////////////////
		// We have to avoid re-entrant loops here. The problem is that client and server 
		// are interconnected in a way that if the client changes its object selection, 
		// it will notify the server which will updates its own selection and notify 
		// back the client about it, etc, etc, we will loop. To avoid the loop, the 
		// client NULLs the top level server pointer so that all subsequent functions
		// called will not process anything.
		///////////////////////////////////////////////////////////////////////////
		IPowerPCBApp *tmpPtr = m_pPowerPCBApplication;
		m_pPowerPCBApplication = NULL;
		
		// Reset the server selection list
		doc.SelectObjects(ppcbObjectTypeAll, NULL, FALSE);

		if (count > 0)
		{
			// Get the selected items in the listbox
			int *items = (int *) new int[count];
			listbox->GetSelItems(count, items);

			for (int i=0; i<count; i++)
			{
				char objectName[255];
				BOOL rc = listbox->GetText(items[i], objectName); ASSERT(rc != LB_ERR);
				doc.SelectObjects(ppcbObjectTypeComponent, objectName, TRUE);
			}
			delete [] items;
		}

		// Put it back
		m_pPowerPCBApplication = tmpPtr;
		RefreshGeneralInfos();
	}

	EndWaitCursor();
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	OnListBoxSelectionChange()
// Desc:		Object list box selection change callback.
/////////////////////////////////////////////////////////////////////////////
void CClientDlg::OnListBoxSelectionChange()
{
	try {
		SendSelection();
	}
	catch (COleException* e) {
		ProcessException(e);
	}
}


//////////////////////////////////////////////////////////////////////////////
// Routine:	ShowObjectInfo()
// Desc:		This function retrieves	what has been selected in the listbox 
//				and gets infos on the selected
//				item. It demonstrates how to get infos on a PowerPCB object.
//////////////////////////////////////////////////////////////////////////////
void CClientDlg::ShowObjectInfo()
{
	if (m_pPowerPCBApplication == NULL) return;
	
	BeginWaitCursor();

	LPDISPATCH pDispatch = NULL;

	// Get the number of items selected in the listbox. If something is selected, process...
	CListBox *listbox = (CListBox *)GetDlgItem(IDC_LST_DOCOBJECTS);
	int count = listbox->GetSelCount(); ASSERT(count = 1);

	pDispatch = m_pPowerPCBApplication->GetActiveDocument(); ASSERT(pDispatch != NULL);
	IPowerPCBDoc doc; doc.AttachDispatch(pDispatch);

	// Get the selected item in the listbox
	int index = listbox->GetCurSel();

	char objectName[255];
	BOOL rc = listbox->GetText(index, objectName); ASSERT(rc != LB_ERR);
			
	// Get the collection of matching objects (should be a single object in the collection)
	pDispatch = doc.GetObjects(ppcbObjectTypeComponent, objectName, FALSE); ASSERT(pDispatch != NULL);
	IPowerPCBObjs matchingObjects; matchingObjects.AttachDispatch(pDispatch);
	ASSERT(matchingObjects.GetCount() == 1);
	
	// Get the current unit
	CString UnitString; long unit = doc.GetUnit();
	switch (unit)
	{
		case 2:	UnitString = "Mils"; break;
		case 3:	UnitString = "Inch"; break;
		case 4:	UnitString = "Metric"; break;
		default:							UnitString = "Unknown unit"; break;
	}
	
	
	char line1[255];
	char line2[255];
	char line3[255];
	char line4[255];
	char line5[255];
	char line6[255];
	char line7[255];
	char line8[255];
	char pInfoMessage[2048];

	COleVariant v((long)1);
	pDispatch = matchingObjects.GetItem(v); ASSERT(pDispatch != NULL);
	IPowerPCBComp comp; comp.AttachDispatch(pDispatch);

	// Line 1: Object
	sprintf(line1, "Object: %s component.", (comp.GetIsSMD() == TRUE) ? "SMD" : (comp.GetIsSMD() == FALSE) ? "Through Hole" : "Unknown");

	// Line 2: Object Name
	CString name = comp.GetName();
	sprintf(line2, "Object Name: %s.", name);

	// Line 3: Object Location
	double locX = comp.GetPositionX(0);
	double locY = comp.GetPositionY(0);
	sprintf(line3, "Object Location (%s): (%.2f, %.2f).", UnitString, locX, locY);

	// Line 4: Object Orientation
	double angle = comp.GetOrientation();
	sprintf(line4, "Object Orientation (drg): %.2f.", angle);

	// Line 5: Object Part Type
	CString type = comp.GetPartType();
	sprintf(line5, "Object Type: %s.", type);

	// Line 6: Object Part Decal
	CString decal = comp.GetDecal();
	sprintf(line6, "Object Decal: %s.", decal);

	// Line 7: Object Layer
	int layer = comp.GetLayer();
	CString layername = doc.GetLayerName(layer);
	sprintf(line7, "Object Layer: %i (%s).", layer, layername);
			
	// Line 8: Object Pin count
	IPowerPCBObjs pins; pins.AttachDispatch(comp.GetPins());
	int countPins = pins.GetCount();
	sprintf(line8, "Object Pins: %i pins.", countPins);
			
	sprintf(pInfoMessage, "%s\n%s\n%s\n%s\n%s\n%s\n%s\n%s\n", line1, line2, line3, line4, line5, line6, line7, line8); 

	AfxMessageBox(pInfoMessage, MB_OK);
}

//////////////////////////////////////////////////////////////////////////////
// Routine:	OnListBoxDoubleClick()
// Desc:		Objects list box double click callback. 
//////////////////////////////////////////////////////////////////////////////
void CClientDlg::OnListBoxDoubleClick() 
{
	try {
		ShowObjectInfo();
	}
	catch (COleException* e) {
		ProcessException(e);
	}
} // OnListBoxDoubleClick()

