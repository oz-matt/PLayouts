/* Copyright Mentor Graphics Corporation 2007

All Rights Reserved.

THIS WORK CONTAINS TRADE SECRET
AND PROPRIETARY INFORMATION WHICH IS THE
PROPERTY OF MENTOR GRAPHICS
CORPORATION OR ITS LICENSORS AND IS
SUBJECT TO LICENSE TERMS. 
*/

// oleappDlg.cpp : implementation file
//

#include "stdafx.h"
#include "oleapp.h"
#include "oleappDlg.h"
#include "DlgProxy.h"
#include "PowerLogic.h"
#include "PowerLogic.inc"
#include "sink.h"

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
// COleappDlg dialog

IMPLEMENT_DYNAMIC(COleappDlg, CDialog);

COleappDlg::COleappDlg(CWnd* pParent /*=NULL*/)
	: CDialog(COleappDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(COleappDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = NULL;

	m_pPowerLogicApplication = NULL;
	m_dwAppConnectionID		= 0;
	m_dwDocConnectionID		= 0;
}

COleappDlg::~COleappDlg()
{
	// If there is an automation proxy for this dialog, set
	//  its back pointer to this dialog to NULL, so it knows
	//  the dialog has been deleted.
	if (m_pAutoProxy != NULL)
		m_pAutoProxy->m_pDialog = NULL;
}

void COleappDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(COleappDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(COleappDlg, CDialog)
	//{{AFX_MSG_MAP(COleappDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_CLOSE()
	ON_BN_CLICKED(IDC_BTN_CONNECT, OnBtnConnect)
	ON_BN_CLICKED(IDC_BTN_DISCONNECT, OnBtnDisconnect)
	ON_BN_CLICKED(IDC_BTN_REFRESH, OnBtnRefresh)
	ON_LBN_DBLCLK(IDC_LST_DOCOBJECTS, OnDblclkLstDocobjects)
	ON_LBN_SELCHANGE(IDC_LST_DOCOBJECTS, OnSelchangeLstDocobjects)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// COleappDlg message handlers

BOOL COleappDlg::OnInitDialog()
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
	SetDlgItemText(IDC_TXT_STATUS, "Not connected to PADS Logic.");
	((CButton *) GetDlgItem(IDC_BTN_CONNECT))->EnableWindow(TRUE);
	((CButton *) GetDlgItem(IDC_BTN_DISCONNECT))->EnableWindow(FALSE);
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void COleappDlg::OnSysCommand(UINT nID, LPARAM lParam)
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

void COleappDlg::OnPaint() 
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
HCURSOR COleappDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

// Automation servers should not exit when a user closes the UI
//  if a controller still holds on to one of its objects.  These
//  message handlers make sure that if the proxy is still in use,
//  then the UI is hidden but the dialog remains around if it
//  is dismissed.

void COleappDlg::OnClose() 
{
	if (CanExit())
		CDialog::OnClose();
}

void COleappDlg::OnOK() 
{
	if (CanExit())
		CDialog::OnOK();
}

void COleappDlg::OnCancel() 
{
	if (CanExit())
		CDialog::OnCancel();
}

BOOL COleappDlg::CanExit()
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

void COleappDlg::Connect() 
{
	if (m_pPowerLogicApplication != NULL) { ASSERT(FALSE); return; }

 	BOOL rc;
	HRESULT hResult;

	SetDlgItemText(IDC_TXT_STATUS, "Connecting to PADS Logic server...");

	/////////////////////////////////////////////////////////////////////////////
	// Get the CLSID from the PADS Logic name string ID (system registry query)
	CLSID clsID;
	hResult = ::CLSIDFromProgID(OLESTR("PowerLogic.Application"), &clsID);
	if (hResult != NOERROR) { ASSERT(FALSE); return; }

	/////////////////////////////////////////////////////////////////////////////
	// Try to get an active instance of PADS Logic. The ::GetActiveObject() function
	// will query the Running Object Table and search for a registered PADS Logic
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
		SetDlgItemText(IDC_TXT_STATUS, "Attaching to running PADS Logic server...");

		// Get IDispatch interface of the PADS Logic application object.
		LPDISPATCH pDispatch = NULL;
		hResult = pUnknown->QueryInterface(IID_IDispatch, (void **) &pDispatch);
		ASSERT(hResult == S_OK);

		// Allocate the interface
		m_pPowerLogicApplication = (IPowerLogicApp *) new IPowerLogicApp;
		ASSERT(m_pPowerLogicApplication != NULL);

		// Attach this dispatch pointer our pPowerLogicApp interface pointer.
		m_pPowerLogicApplication->AttachDispatch(pDispatch, FALSE);

		// Release the pUnknown we have used.
		pUnknown->Release();	   

		EndWaitCursor();
	}
	/////////////////////////////////////////////////////////////////////////////
	// If no such instance was found, start a new one and connect to it.
	else
	{
		BeginWaitCursor();
		SetDlgItemText(IDC_TXT_STATUS, "Starting and connecting to PADS Logic server...");

		// Allocate the interface
		m_pPowerLogicApplication = (IPowerLogicApp *) new IPowerLogicApp;
		ASSERT(m_pPowerLogicApplication != NULL);

		// Create pPowerLogicApp dispatch
		COleException e;

		////////////////////////////////////////////////////////////////////////
		// The CreateDispatch() function is very powerful because it encapsulate
		//	everything to establich the connection. However, if the function fails,
		// or if the system gets stuck inside this function, it is very hard to 
		// debug and diagnostic what went wrong. The #else section contains the 
		// expanded code of CreateDispatch() for easier debugging in case of a problem.
		////////////////////////////////////////////////////////////////////////
#if 1
		rc = m_pPowerLogicApplication->CreateDispatch(clsID, &e);

#else
		ASSERT(m_pPowerLogicApplication->m_lpDispatch == NULL);

		m_pPowerLogicApplication->m_bAutoRelease = TRUE;  // Good default is to auto-release

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
		hResult = lpUnknown->QueryInterface(IID_IDispatch, (void **) &(m_pPowerLogicApplication->m_lpDispatch));
		if (hResult != S_OK) { ASSERT(FALSE); lpUnknown->Release(); e.m_sc = hResult; return; }

		lpUnknown->Release();
		ASSERT(m_pPowerLogicApplication->m_lpDispatch != NULL);

#endif

		EndWaitCursor();

		if (rc == FALSE) 
		{
			ASSERT(FALSE);	
			delete m_pPowerLogicApplication; m_pPowerLogicApplication = NULL;
			SetDlgItemText(IDC_TXT_STATUS, "Not connected to PowerLogic.");
			return;
		}
	}
	
	// Connect sink
	SetDlgItemText(IDC_TXT_STATUS, "Connecting sink...");
	SetSink(TRUE);

	((CButton *) GetDlgItem(IDC_BTN_CONNECT))->EnableWindow(FALSE);
	((CButton *) GetDlgItem(IDC_BTN_DISCONNECT))->EnableWindow(TRUE);
	SetDlgItemText(IDC_TXT_STATUS, "Connected to PADS Logic.");

	m_pPowerLogicApplication->SetVisible(TRUE);
}

void COleappDlg::OnBtnConnect() 
{
	try {
		Connect();
	}
	catch (COleException* e) {
		ProcessException(e);
	}
}

//ProcessException
void COleappDlg::ProcessException(COleException* e)
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
		//probably PADS Logic is crashed or ubnormal terminated
		// Delete the connection with PADS Logic by deleting it's reference.
		delete m_pPowerLogicApplication; 
		m_pPowerLogicApplication = NULL;
		SetDlgItemText(IDC_TXT_STATUS, "Not connected to PADS Logic.");
		((CButton *) GetDlgItem(IDC_BTN_CONNECT))->EnableWindow(TRUE);
		((CButton *) GetDlgItem(IDC_BTN_DISCONNECT))->EnableWindow(FALSE);
	}

	CString str;
	str.Format("OLE exception: %s (%lx)", szMessage, hr);
	AfxMessageBox(str);
	LocalFree(szMessage);
}


/////////////////////////////////////////////////////////////////////////////
// Routine:	SafeDisconnect()
// Desc:		Disconnects from PADS Logic Server safely. Stop routing exceptions here.
/////////////////////////////////////////////////////////////////////////////
void COleappDlg::SafeDisconnect()
{
	try {
		Disconnect();
	}
	catch (COleException* e) {
		ProcessException(e);
	}
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	OnDisconnect()
// Desc:		Disconnects from PADS Logic Server.
/////////////////////////////////////////////////////////////////////////////
void COleappDlg::Disconnect()
{
	if (m_pPowerLogicApplication == NULL) { ASSERT(FALSE); return; }

	// Disconnect sink
	SetSink(FALSE);

	// Delete the connection with PADS Logic by deleting it's reference.
	delete m_pPowerLogicApplication; 
	m_pPowerLogicApplication = NULL;

	SetDlgItemText(IDC_TXT_STATUS, "Not connected to PADS Logic.");

	((CButton *) GetDlgItem(IDC_BTN_CONNECT))->EnableWindow(TRUE);
	((CButton *) GetDlgItem(IDC_BTN_DISCONNECT))->EnableWindow(FALSE);
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	OnBtnDisconnect()
// Desc:		Disconnect button callback.
/////////////////////////////////////////////////////////////////////////////
void COleappDlg::OnBtnDisconnect() 
{
	SafeDisconnect();
}

BOOL COleappDlg::SetSink(BOOL bStatus)
{
	HRESULT hr;
	ASSERT (m_pPowerLogicApplication->m_lpDispatch != NULL);
	IPowerLogicDoc doc; doc.AttachDispatch(m_pPowerLogicApplication->GetActiveDocument());

	LPCONNECTIONPOINTCONTAINER pAppConnPtCont = NULL;
	LPCONNECTIONPOINTCONTAINER pDocConnPtCont = NULL;
	hr = m_pPowerLogicApplication->m_lpDispatch->QueryInterface(IID_IConnectionPointContainer, (LPVOID*)&pAppConnPtCont);
	if (hr != S_OK) return FALSE;

	hr = doc.m_lpDispatch->QueryInterface(IID_IConnectionPointContainer, (LPVOID*)&pDocConnPtCont);
	if (hr != S_OK) { doc.ReleaseDispatch(); return FALSE; }
	

	LPCONNECTIONPOINT pAppConnPt = NULL;
	hr = pAppConnPtCont->FindConnectionPoint(DIID__PowerLogicAppEvents, &pAppConnPt);
	if (hr != S_OK) return FALSE;

	LPCONNECTIONPOINT pDocConnPt = NULL;
	hr = pDocConnPtCont->FindConnectionPoint(DIID__PowerLogicDocEvents, &pDocConnPt);
	if (hr != S_OK) { doc.ReleaseDispatch(); return FALSE; }
		
	LPDISPATCH lpDispatch = m_PowerLogicSink.GetIDispatch(FALSE); ASSERT(lpDispatch);
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

void COleappDlg::Refresh()
{
	OnBtnRefresh();
}

void COleappDlg::OnBtnRefresh() 
{
	try {
		// We know that the RefreshObjectsInfos() function performs a lengthy
		// operation, which is mainly OLE calls to PADS Logic server. We will lock
		// the server so that these OLE calls are very fast.
		if (m_pPowerLogicApplication) m_pPowerLogicApplication->LockServer();
		RefreshGeneralInfos();
		RefreshObjectsInfos();
		RefreshSelectionInfos();
		if (m_pPowerLogicApplication) m_pPowerLogicApplication->UnlockServer();
	}
	catch (COleException* e) {
		ProcessException(e);
	}
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	RefreshGeneralInfos()
// Desc:		Get data from PADS Logic server (fast access data)
/////////////////////////////////////////////////////////////////////////////
void COleappDlg::RefreshGeneralInfos()
{
	if (m_pPowerLogicApplication == NULL)
	{
		// Not connected to server -> print dummy strings
		SetDlgItemText(IDC_EDT_DOCNAME, "???????");
		SetDlgItemText(IDC_EDT_COMPS, "???????");
		SetDlgItemText(IDC_EDT_SELECTEDCOMPS, "???????");
	}
	else
	{
		// Call the server Application ActiveDocument property which returns
		// a pointer to its IDispatch interface. PADS Logic garantees that the
		// returned pointed object implements the IPowerLogicDoc interface as
		// well, so we attach this IDispatch pointer to a IPowerLogicDoc object.
		LPDISPATCH pDispatch = m_pPowerLogicApplication->GetActiveDocument();
		ASSERT(pDispatch);
		IPowerLogicDoc doc; doc.AttachDispatch(pDispatch);

		// Call the server Document Name property and outpu its value.
		CString docname = doc.GetName();
		SetDlgItemText(IDC_EDT_DOCNAME, docname);

		// Get the number of components
		pDispatch = doc.GetComponents();
		ASSERT(pDispatch);
		IPowerLogicObjs objs1; objs1.AttachDispatch(pDispatch);
		char selCountStr[255];
		sprintf(selCountStr, "%i component object(s)", objs1.GetCount());
		SetDlgItemText(IDC_EDT_COMPS, selCountStr);

		// Get the number of selected components
		pDispatch = doc.GetObjects(plogObjectTypeComponent, NULL, TRUE);
		ASSERT(pDispatch);
		IPowerLogicObjs objs2; objs2.AttachDispatch(pDispatch);
		sprintf(selCountStr, "%i selected component object(s)", objs2.GetCount());
		SetDlgItemText(IDC_EDT_SELECTEDCOMPS, selCountStr);
	}
	SetDlgItemText(IDC_TXT_STATUS, "");
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	RefreshObjectsInfos()
// Desc:		Get data from PADS Logic server (slow access data)
/////////////////////////////////////////////////////////////////////////////
void COleappDlg::RefreshObjectsInfos()
{
	if (m_pPowerLogicApplication == NULL)
	{
		CListBox *listbox = (CListBox *)GetDlgItem(IDC_LST_DOCOBJECTS);
		listbox->ResetContent();
	}
	else
	{
		BeginWaitCursor();

		// Call the server Application ActiveDocument property which returns
		// a pointer to its IDispatch interface. PADS Logic garantees that the
		// returned pointed object implements the IPowerLogicDoc interface as
		// well, so we attach this IDispatch pointer to a IPowerLogicDoc object.
		LPDISPATCH pDispatch = m_pPowerLogicApplication->GetActiveDocument();
		ASSERT(pDispatch);
		IPowerLogicDoc doc; doc.AttachDispatch(pDispatch);

		// Reset our listbox
		CListBox *listbox = (CListBox *)GetDlgItem(IDC_LST_DOCOBJECTS);
		listbox->ResetContent();

		LPDISPATCH pObjs;
		
		// Get all PADS Logic objects of type COMPONENT from PADS Logic database
		// and output each component in the listbox
		int countComp = 0;
		pObjs = doc.GetComponents(); ASSERT(pObjs != NULL);
		IPowerLogicObjs comps; comps.AttachDispatch(pObjs);
		int count = comps.GetCount();
		int percent = 0;
		// Use PADS Logic Progress Bar
		m_pPowerLogicApplication->SetStatusBarText("Sending components to client...");
		m_pPowerLogicApplication->SetProgressBar(percent);
		// For each component returned by PADS Logic...
		for (int i=1; i<=count; i++)
		{
			// ... get the component object itself...
			COleVariant v((long)i);
			LPDISPATCH pObj = comps.GetItem(v); ASSERT(pObj != NULL);
			IPowerLogicComp comp; comp.AttachDispatch(pObj);
					
			// ... extract its name and add it to the listbox.
			int index = listbox->InsertString(i-1, comp.GetName());
			countComp++; char msg[255]; sprintf(msg, "Loading %i component(s) from PADS Logic...", countComp);
			SetDlgItemText(IDC_TXT_STATUS, msg);
			//update progress bar but avoid extra calls
			int newPercent = i * 100 / count;
			if (newPercent != percent) {
				percent = newPercent;
				m_pPowerLogicApplication->SetProgressBar(percent);
			}
		}
		m_pPowerLogicApplication->SetProgressBar(-1);

		EndWaitCursor();
	
	}
	SetDlgItemText(IDC_TXT_STATUS, "");
}

/////////////////////////////////////////////////////////////////////////////
// Routine:	RefreshSelectionInfos()
// Desc:		Get selection data from PADS Logic server
/////////////////////////////////////////////////////////////////////////////
void COleappDlg::RefreshSelectionInfos()
{
	if (m_pPowerLogicApplication == NULL) return;

	BeginWaitCursor();

	LPDISPATCH pDispatch = NULL;
	pDispatch= m_pPowerLogicApplication->GetActiveDocument(); ASSERT(pDispatch != NULL);
	IPowerLogicDoc doc; doc.AttachDispatch(pDispatch);

	// Reset listbox selection
	CListBox *listbox = (CListBox *)GetDlgItem(IDC_LST_DOCOBJECTS);
	BOOL rc = listbox->SetSel(-1, FALSE); ASSERT(rc != LB_ERR);

	// Get the selected COMPONENTS
	pDispatch = doc.GetObjects(1, NULL, TRUE); ASSERT(pDispatch != NULL);
	IPowerLogicObjs comps; comps.AttachDispatch(pDispatch);
	for (int i=1; i<=comps.GetCount(); i++) {
		COleVariant v((long)i);
		pDispatch = comps.GetItem(v); ASSERT(pDispatch != NULL);
		IPowerLogicComp comp; comp.AttachDispatch(pDispatch);
		// Find this component in the listbox.
		int index = listbox->FindString(0, comp.GetName()); // ASSERT(index != LB_ERR);
		// Select it in the listbox.
		rc = listbox->SetSel(index, TRUE); ASSERT(rc != LB_ERR);
	}
		
	EndWaitCursor();
} // RefreshSelectionInfos()



//////////////////////////////////////////////////////////////////////////////
// Desc:		This function retrieves
//				what has been selected in the listbox and gets infos on the selected
//				item. It demonstrates how to get infos on a PADS Logic object.
//////////////////////////////////////////////////////////////////////////////
void COleappDlg::ShowObjectInfo()
{
	if (m_pPowerLogicApplication == NULL) return;
	
	BeginWaitCursor();

	LPDISPATCH pDispatch = NULL;

	// Get the number of items selected in the listbox. If something is selected, process...
	CListBox *listbox = (CListBox *)GetDlgItem(IDC_LST_DOCOBJECTS);
	int count = listbox->GetSelCount(); ASSERT(count = 1);

	pDispatch = m_pPowerLogicApplication->GetActiveDocument(); ASSERT(pDispatch != NULL);
	IPowerLogicDoc doc; doc.AttachDispatch(pDispatch);

	// Get the selected item in the listbox
	int index = listbox->GetCurSel();

	char objectName[255];
	BOOL rc = listbox->GetText(index, objectName); ASSERT(rc != LB_ERR);
			
	// Get the collection of matching objects (should be a single object in the collection)
	pDispatch = doc.GetObjects(plogObjectTypeComponent, objectName, FALSE); ASSERT(pDispatch != NULL);
	IPowerLogicObjs matchingObjects; matchingObjects.AttachDispatch(pDispatch);
	ASSERT(matchingObjects.GetCount() == 1);
	
	char line1[255];
	char line2[255];
	char line3[255];
	char line4[255];
//	char line5[255];
//	char line6[255];
//	char line7[255];
//	char line8[255];
	char pInfoMessage[2048];

	COleVariant v((long)1);
	pDispatch = matchingObjects.GetItem(v); ASSERT(pDispatch != NULL);
	IPowerLogicComp comp; comp.AttachDispatch(pDispatch);

	// Line 1: Object Name
	CString name = comp.GetName();
	sprintf(line1, "Object Name: %s.", name);

	// Line 2: Object Part Type
	CString type = comp.GetPartType();
	sprintf(line2, "Object Type: %s.", type);

	// Line 3: Object Part Decal
	CString decal = comp.GetPCBDecal();
	sprintf(line3, "Object Decal: %s.", decal);

	// Line 4: Object Pin count
	IPowerLogicObjs pins; pins.AttachDispatch(comp.GetPins());
	int countPins = pins.GetCount();
	sprintf(line4, "Object Pins: %i pins.", countPins);
			
	sprintf(pInfoMessage, "%s\n%s\n%s\n%s\n", line1, line2, line3, line4); 

	AfxMessageBox(pInfoMessage, MB_OK);
}

//////////////////////////////////////////////////////////////////////////////
// Desc:		Objects list box double click callback. 
//////////////////////////////////////////////////////////////////////////////
void COleappDlg::OnDblclkLstDocobjects() 
{
	try {
		ShowObjectInfo();
	}
	catch (COleException* e) {
		ProcessException(e);
	}
}

/////////////////////////////////////////////////////////////////////////////
// Desc:		This function retrieves
//				what has been selected in the listbox and selects the same thing in 
//				PADS Logic. It demonstrate the SelectObjects() routine as well as
//				PADS Logic object collections manipulation.
/////////////////////////////////////////////////////////////////////////////
void COleappDlg::SendSelection()
{
	if (m_pPowerLogicApplication == NULL) return;
	
	BeginWaitCursor();
	
	LPDISPATCH pDispatch = NULL;

	// Get the number of items selected in the listbox. If something is selected, process...
	CListBox *listbox = (CListBox *)GetDlgItem(IDC_LST_DOCOBJECTS);
	int count = listbox->GetSelCount();
	if (count > 0)
	{
		pDispatch = m_pPowerLogicApplication->GetActiveDocument(); ASSERT(pDispatch != NULL);
		IPowerLogicDoc doc; doc.AttachDispatch(pDispatch);

		///////////////////////////////////////////////////////////////////////////
		// We have to avoid re-entrant loops here. The problem is that client and server 
		// are interconnected in a way that if the client changes its object selection, 
		// it will notify the server which will updates its own selection and notify 
		// back the client about it, etc, etc, we will loop. To avoid the loop, the 
		// client NULLs the top level server pointer so that all subsequent functions
		// called will not process anything.
		///////////////////////////////////////////////////////////////////////////
		//IPowerLogicApp *tmpPtr = m_pPowerLogicApplication;
		//m_pPowerLogicApplication = NULL;
		SetSink(FALSE);
		
		// Reset the server selection list
		doc.SelectObjects(plogObjectTypeAll, NULL, FALSE);

		// Get the selected items in the listbox
		int *items = (int *) new int[count];
		listbox->GetSelItems(count, items);

		// Use PADS Logic Progress Bar
		int percent = 0;
		if (count > 1) {
			m_pPowerLogicApplication->SetStatusBarText("Receiving selection from client...");
			m_pPowerLogicApplication->SetProgressBar(percent);
		}

		for (int i=0; i<count; i++)
		{
			char objectName[255];
			BOOL rc = listbox->GetText(items[i], objectName); ASSERT(rc != LB_ERR);
			doc.SelectObjects(plogObjectTypeComponent, objectName, TRUE);
			//update progress bar but avoid extra calls
			int newPercent = i * 100 / count;
			if (count > 1 && newPercent != percent) {
				percent = newPercent;
				m_pPowerLogicApplication->SetProgressBar(percent);
			}
		}
		m_pPowerLogicApplication->SetProgressBar(-1);
		delete [] items;

		// Put it back
		//m_pPowerLogicApplication = tmpPtr;
		SetSink(TRUE);

		RefreshGeneralInfos();

		//Switch to PADS Logic sheet containing selection if needed
		LocateSelection();
	}

	EndWaitCursor();
}

/////////////////////////////////////////////////////////////////////////////
// Desc:		Object list box selection change callback. 
/////////////////////////////////////////////////////////////////////////////
void COleappDlg::OnSelchangeLstDocobjects() 
{
	try {
		SendSelection();
	}
	catch (COleException* e) {
		ProcessException(e);
	}
}

/////////////////////////////////////////////////////////////////////////////
// Desc:		This function searchs first sheet containing selected objects
//				and activates it
/////////////////////////////////////////////////////////////////////////////
void COleappDlg::LocateSelection() 
{
	if (m_pPowerLogicApplication == NULL) return;
	LPDISPATCH pDispatch = NULL;
	pDispatch = m_pPowerLogicApplication->GetActiveDocument(); ASSERT(pDispatch != NULL);
	IPowerLogicDoc doc; doc.AttachDispatch(pDispatch);
	//get active sheet
	pDispatch = doc.GetActiveSheet(); ASSERT(pDispatch != NULL);
	IPowerLogicSheet sheet; sheet.AttachDispatch(pDispatch);
	//is there selection on active sheet?
	pDispatch = sheet.GetObjects(plogObjectTypeComponent, NULL, TRUE); ASSERT(pDispatch != NULL);
	IPowerLogicObjs comps; comps.AttachDispatch(pDispatch);
	//if current sheet has selection do not change sheet
	if (comps.GetCount() > 0) {
		return;
	}
	//iterate through all sheets in order to find first one with selection
	pDispatch = doc.GetSheets(); ASSERT(pDispatch != NULL);
	IPowerLogicSheets sheets; sheets.AttachDispatch(pDispatch);
	for (int i = 1; i <= sheets.GetCount(); i++) {
		//get next sheet object
		COleVariant v((long)i);
		pDispatch = sheets.GetItem(v); ASSERT(pDispatch != NULL);
		IPowerLogicSheet sheet; sheet.AttachDispatch(pDispatch);
		//is there selection on this sheet?
		pDispatch = sheet.GetObjects(plogObjectTypeComponent, NULL, TRUE); ASSERT(pDispatch != NULL);
		IPowerLogicObjs comps; comps.AttachDispatch(pDispatch);
		if (comps.GetCount() > 0) {
			//activate sheet and exit
			sheet.Activate();
			break;
		}
	}
}
