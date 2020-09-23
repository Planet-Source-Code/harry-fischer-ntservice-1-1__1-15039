// NtsvcCtl.cpp : Implementation of the CNtsvcCtrl ActiveX Control class.

#include "stdafx.h"
#include "Ntsvc.h"
#include "NtsvcCtl.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CNtsvcCtrl, COleControl)

// private window messages
#define CWM_START   WM_USER+10
#define CWM_HANDLER	WM_USER+11

/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CNtsvcCtrl, COleControl)
	//{{AFX_MSG_MAP(CNtsvcCtrl)
	ON_WM_ERASEBKGND()
	//}}AFX_MSG_MAP
	ON_MESSAGE(CWM_START, OnStart)
	ON_MESSAGE(CWM_HANDLER, OnHandler)
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Dispatch map

BEGIN_DISPATCH_MAP(CNtsvcCtrl, COleControl)
	//{{AFX_DISPATCH_MAP(CNtsvcCtrl)
	DISP_PROPERTY_NOTIFY(CNtsvcCtrl, "Account", m_account, OnAccountChanged, VT_BSTR)
	DISP_PROPERTY_NOTIFY(CNtsvcCtrl, "ControlsAccepted", m_controlsAccepted, OnControlsAcceptedChanged, VT_I4)
	DISP_PROPERTY_NOTIFY(CNtsvcCtrl, "LoadOrderGroup", m_LoadOrderGroup, OnLoadOrderGroupChanged, VT_BSTR)
	DISP_PROPERTY_NOTIFY(CNtsvcCtrl, "Dependencies", m_Dependencies, OnDependenciesChanged, VT_BSTR)
	DISP_PROPERTY_NOTIFY(CNtsvcCtrl, "DisplayName", m_displayName, OnDisplayNameChanged, VT_BSTR)
	DISP_PROPERTY_NOTIFY(CNtsvcCtrl, "Interactive", m_interactive, OnInteractiveChanged, VT_BOOL)
	DISP_PROPERTY_NOTIFY(CNtsvcCtrl, "Password", m_password, OnPasswordChanged, VT_BSTR)
	DISP_PROPERTY_NOTIFY(CNtsvcCtrl, "WaitHint", m_waitHint, OnWaitHintChanged, VT_I4)
	DISP_PROPERTY_EX(CNtsvcCtrl, "Debug", GetDebug, SetDebug, VT_BOOL)
	DISP_PROPERTY_EX(CNtsvcCtrl, "ServiceName", GetServiceName, SetServiceName, VT_BSTR)
	DISP_PROPERTY_EX(CNtsvcCtrl, "StartMode", GetStartMode, SetStartMode, VT_I2)
	DISP_FUNCTION(CNtsvcCtrl, "Install", Install, VT_BOOL, VTS_NONE)
	DISP_FUNCTION(CNtsvcCtrl, "LogEvent", LogEvent, VT_BOOL, VTS_I2 VTS_I4 VTS_BSTR)
	DISP_FUNCTION(CNtsvcCtrl, "Running", Running, VT_BOOL, VTS_NONE)
	DISP_FUNCTION(CNtsvcCtrl, "Uninstall", Uninstall, VT_BOOL, VTS_NONE)
	DISP_FUNCTION(CNtsvcCtrl, "StopService", StopService, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CNtsvcCtrl, "StartService", StartService, VT_BOOL, VTS_NONE)
	DISP_FUNCTION(CNtsvcCtrl, "GetAllSettings", GetAllSettings, VT_VARIANT, VTS_BSTR)
	DISP_FUNCTION(CNtsvcCtrl, "SaveSetting", SaveSetting, VT_EMPTY, VTS_BSTR VTS_BSTR VTS_BSTR)
	DISP_FUNCTION(CNtsvcCtrl, "DeleteSetting", DeleteSetting, VT_EMPTY, VTS_BSTR VTS_VARIANT)
	DISP_FUNCTION(CNtsvcCtrl, "GetSetting", GetSetting, VT_BSTR, VTS_BSTR VTS_BSTR VTS_VARIANT)
	DISP_FUNCTION(CNtsvcCtrl, "ReportStatus", ReportSrvStatus, VT_EMPTY, VTS_NONE)
	DISP_FUNCTION(CNtsvcCtrl, "GetServiceStatusHandle", GetServiceStatusHandle, VT_I4, VTS_NONE)
	DISP_DEFVALUE(CNtsvcCtrl, "DisplayName")
	//}}AFX_DISPATCH_MAP
	DISP_FUNCTION_ID(CNtsvcCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()


/////////////////////////////////////////////////////////////////////////////
// Event map

BEGIN_EVENT_MAP(CNtsvcCtrl, COleControl)
	//{{AFX_EVENT_MAP(CNtsvcCtrl)
	EVENT_CUSTOM("Continue", FireContinue, VTS_PBOOL)
	EVENT_CUSTOM("Control", FireControl, VTS_I4)
	EVENT_CUSTOM("Pause", FirePause, VTS_PBOOL)
	EVENT_CUSTOM("Start", FireStart, VTS_PBOOL)
	EVENT_CUSTOM("Stop", FireStop, VTS_NONE)
	//}}AFX_EVENT_MAP
END_EVENT_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CNtsvcCtrl, "NTService.Control.1",
	0xc32049a5, 0xe8d3, 0x11d4, 0x9d, 0x82, 0, 0xd0, 0xb7, 0x3b, 0x61, 0xa4)


/////////////////////////////////////////////////////////////////////////////
// Type library ID and version

IMPLEMENT_OLETYPELIB(CNtsvcCtrl, _tlid, _wVerMajor, _wVerMinor)


/////////////////////////////////////////////////////////////////////////////
// Interface IDs

const IID BASED_CODE IID_DNtsvc =
		{ 0xc32049a3, 0xe8d3, 0x11d4, { 0x9d, 0x82, 0, 0xd0, 0xb7, 0x3b, 0x61, 0xa4 } };
const IID BASED_CODE IID_DNtsvcEvents =
		{ 0xc32049a4, 0xe8d3, 0x11d4, { 0x9d, 0x82, 0, 0xd0, 0xb7, 0x3b, 0x61, 0xa4 } };


/////////////////////////////////////////////////////////////////////////////
// Control type information

static const DWORD BASED_CODE _dwNtsvcOleMisc =
	OLEMISC_INVISIBLEATRUNTIME |	// control has no GUI
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_SIMPLEFRAME |			// Guarantees that the control window exists and messages can be posted
	OLEMISC_ALWAYSRUN |				// The control generates events without being visible
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE /* |
	OLEMISC_RECOMPOSEONRESIZE */	// Disabled because control doesn't resize
	;

IMPLEMENT_OLECTLTYPE(CNtsvcCtrl, IDS_NTSVC, _dwNtsvcOleMisc)


class CNtsvcCtrl* CNtsvcCtrl::m_pThis;


/////////////////////////////////////////////////////////////////////////////
// CNtsvcCtrl::CNtsvcCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CNtsvcCtrl

BOOL CNtsvcCtrl::CNtsvcCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_NTSVC,
			IDB_NTSVC,
			FALSE,                      //  Not insertable
			_dwNtsvcOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


/////////////////////////////////////////////////////////////////////////////
// Helper painting functions
//
static void CreateMask(CBitmap *pbmp, CBitmap *pbmpMask)
{
	BITMAP bmp;
	pbmp->GetObject(sizeof (BITMAP), &bmp);
	pbmpMask->CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, NULL);

	CDC dcDst;
	dcDst.CreateCompatibleDC(NULL);
	CDC dcSrc;
	dcSrc.CreateCompatibleDC(NULL);
	CBitmap* pbmpSaveDst = dcDst.SelectObject(pbmpMask);
	CBitmap* pbmpSaveSrc = dcSrc.SelectObject(pbmp);

	COLORREF oldBkColor = dcSrc.SetBkColor(dcSrc.GetPixel(0, 0));
	dcDst.BitBlt(0, 0, bmp.bmWidth, bmp.bmHeight, &dcSrc, 0, 0, NOTSRCCOPY);
	dcSrc.SetBkColor(oldBkColor);
}

#define DSx     0x00660046L
#define DSna    0x00220326L

static void DrawMaskedBitmap(CDC* pDC, CBitmap* pbmp, CBitmap* pbmpMask,
	int x, int y, int cx, int cy)
{
	COLORREF oldBkColor = pDC->SetBkColor(RGB(255, 255, 255));
	COLORREF oldTextColor = pDC->SetTextColor(RGB(0, 0, 0));

	CDC dcCompat;
	dcCompat.CreateCompatibleDC(pDC);
	CBitmap* pbmpSave = dcCompat.SelectObject(pbmp);
	pDC->BitBlt(x, y, cx, cy, &dcCompat, 0, 0, DSx);
	dcCompat.SelectObject(pbmpMask);
	pDC->BitBlt(x, y, cx, cy, &dcCompat, 0, 0, DSna);
	dcCompat.SelectObject(pbmp);
	pDC->BitBlt(x, y, cx, cy, &dcCompat, 0, 0, DSx);
	dcCompat.SelectObject(pbmpSave);

	pDC->SetBkColor(oldBkColor);
	pDC->SetTextColor(oldTextColor);
}


/////////////////////////////////////////////////////////////////////////////
// CNtsvcCtrl::CNtsvcCtrl - Constructor

CNtsvcCtrl::CNtsvcCtrl()
{
	InitializeIIDs(&IID_DNtsvc, &IID_DNtsvcEvents);

	EnableSimpleFrame();

	// Initialize control's instance data.
    m_pThis = this;			// WARNING: Only one service per application


	m_hStopEvent = NULL;
	m_hServiceThread = NULL;

	m_hEventSource = NULL;

	ZeroMemory (&m_Status, sizeof(SERVICE_STATUS));	
	ZeroMemory (m_szServiceName, sizeof(m_szServiceName));
  m_hServiceStatus = NULL;
  m_bIsRunning = FALSE;
  m_LoadOrderGroup = "";
  m_Dependencies = "";
	m_controlsAccepted = 0;
	m_bIsDebug = -1;
	m_interactive = FALSE;
	m_account = "";
	m_password = "";
	m_Dependencies = "";
	m_LoadOrderGroup = "";
  m_waitHint = 0;


	// keep fixed control size
	m_bmIcon.LoadBitmap(IDB_NTSVC);
	CreateMask (&m_bmIcon, &m_bmMask);

	BITMAP bmInfo;
	m_bmIcon.GetObject(sizeof(bmInfo), &bmInfo);
	m_pixSize.cx = bmInfo.bmWidth;
	m_pixSize.cy = bmInfo.bmHeight;

	SetInitialSize(28, 28);
	m_Size.cx = m_cxExtent;
	m_Size.cy = m_cyExtent;

	m_brBkgnd.CreateSolidBrush(GetSysColor(COLOR_3DFACE));
}


/////////////////////////////////////////////////////////////////////////////
// CNtsvcCtrl::~CNtsvcCtrl - Destructor

CNtsvcCtrl::~CNtsvcCtrl()
{
    if (m_hEventSource) {
        ::DeregisterEventSource(m_hEventSource);
    }
}


/////////////////////////////////////////////////////////////////////////////
// CNtsvcCtrl::OnDraw - Drawing function

void CNtsvcCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	CRect rc(rcBounds);
	// draw the control icon
	::DrawEdge(pdc->m_hDC, &rc, BDR_INNER,
		BF_RECT | BF_ADJUST | BF_FLAT );

	::DrawEdge(pdc->m_hDC, &rc, EDGE_RAISED, BF_RECT | BF_ADJUST);

	rc.left += (rc.Width() - m_pixSize.cx) / 2;
	rc.top += (rc.Height() - m_pixSize.cy) / 2;
	DrawMaskedBitmap(pdc, &m_bmIcon, &m_bmMask, rc.left, rc.top, m_pixSize.cx, m_pixSize.cy);
}


BOOL CNtsvcCtrl::OnEraseBkgnd(CDC* pDC) 
{
	CRect rc;
	GetClientRect (&rc);
	pDC->FillRect (&rc, &m_brBkgnd);

	return COleControl::OnEraseBkgnd(pDC);
}


BOOL CNtsvcCtrl::OnSetExtent(LPSIZEL lpSizeL)
{
	*lpSizeL = m_Size;

	return COleControl::OnSetExtent (lpSizeL);
}


/////////////////////////////////////////////////////////////////////////////
// CNtsvcCtrl::DoPropExchange - Persistence support

void CNtsvcCtrl::DoPropExchange(CPropExchange* pPX)
{
	BOOL bLoading = pPX->IsLoading();

	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	PX_String(pPX, _T("DisplayName"), m_displayName, "Simple Service");
	PX_String(pPX, _T("Dependencies"), m_Dependencies, "");
	PX_String(pPX, _T("LoadOrderGroup"), m_LoadOrderGroup, "");
	PX_Bool(pPX, _T("Debug"), m_bIsDebug, 0);
	PX_Long(pPX, _T("ControlsAccepted"), m_controlsAccepted, 0);
	PX_String(pPX, _T("Account"), m_account, "");
	PX_String(pPX, _T("Password"), m_password, "");
	PX_Bool(pPX, _T("Interactive"), m_interactive, 0);
	PX_Long(pPX, _T("WaitHint"), m_waitHint, 0);

	if (bLoading)
	{
		CString strProp;
		PX_String(pPX, _T("ServiceName"), strProp, "Simple");
		InternalSetService (strProp);

		long lProp;
		PX_Long(pPX, _T("StartMode"), lProp, SERVICE_DEMAND_START);
		m_dwStartMode = (DWORD) lProp;
	}
	else
	{
		long lProp;

		PX_String(pPX, _T("ServiceName"), CString(m_szServiceName));
		
		lProp = m_dwStartMode;
		PX_Long(pPX, _T("StartMode"), lProp);
	}
}


/////////////////////////////////////////////////////////////////////////////
// CNtsvcCtrl::OnResetState - Reset control to default state

void CNtsvcCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}


/////////////////////////////////////////////////////////////////////////////
// CNtsvcCtrl::AboutBox - Display an "About" box to the user

void CNtsvcCtrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_NTSVC);
	dlgAbout.DoModal();
}


/////////////////////////////////////////////////////////////////////////////
// CNtsvcCtrl message handlers

LRESULT CNtsvcCtrl::OnStart(WPARAM wParam, LPARAM lParam)
{
    // Get a pointer to the C++ object
    CNtsvcCtrl* pService = m_pThis;

	// Start the initialisation
	BOOL bSuccess = FALSE;

	pService->FireStart(&bSuccess);
	if (bSuccess) 
	{
		// Do the real work. 
		pService->m_bIsRunning = TRUE;
        pService->SetStatus(SERVICE_RUNNING);

		if (!m_bIsDebug)
			LogEvent(EVENTLOG_INFORMATION_TYPE, EVMSG_STARTED, "");
	}
	else
	{
		pService->m_bIsRunning = FALSE;
        pService->SetStatus(SERVICE_STOPPED);

		if (!m_bIsDebug)
			LogEvent(EVENTLOG_INFORMATION_TYPE, EVMSG_FAILEDINIT, "");
	}

	// report status to service control manager
	ReportStatus();

	return 0L;
}

LRESULT CNtsvcCtrl::OnHandler(WPARAM wParam, LPARAM lParam)
{
    // Get a pointer to the C++ object
    CNtsvcCtrl* pService = m_pThis;
	BOOL bSuccess = FALSE;

	// Report current status
    ReportStatus();

	switch (wParam)
	{
    case SERVICE_CONTROL_STOP: // 1
		m_pThis->StopService();
        break;

    case SERVICE_CONTROL_PAUSE: // 2
		pService->FirePause(&bSuccess);
		if (bSuccess)
			SetStatus (SERVICE_PAUSED);
        break;

    case SERVICE_CONTROL_CONTINUE: // 3
		pService->FireContinue(&bSuccess);
		if (bSuccess)
			SetStatus (SERVICE_RUNNING);
        break;

    case SERVICE_CONTROL_INTERROGATE: // 4
        break;

    case SERVICE_CONTROL_SHUTDOWN: // 5
		break;

	default:
		pService->FireControl(wParam);
		break;
	}

	// Report current status
    ReportStatus();

	return 0L;
}


/////////////////////////////////////////////////////////////////////////////
// CNtsvcCtrl support member functions
//

void CNtsvcCtrl::InternalSetService(LPCTSTR lpszNewValue)
{
	_tcsncpy(m_szServiceName, lpszNewValue, sizeof(m_szServiceName));
	m_szServiceName[SERVICE_NAME_LEN] = '\0';

    wsprintf(m_szSvcRegKey, TEXT("SYSTEM\\CurrentControlSet\\Services\\%s\\"), m_szServiceName);
}

long CNtsvcCtrl::GetStatus() 
{
	return m_Status.dwCurrentState;
}

void CNtsvcCtrl::SetStatus(long nNewValue) 
{
	m_Status.dwCurrentState = (DWORD) nNewValue;
}

void CNtsvcCtrl::ReportStatus()
{
	// Report current status
	if (m_hServiceStatus)
	    ::SetServiceStatus(m_hServiceStatus, &m_Status);
}


void CNtsvcCtrl::DebugMsg(const TCHAR* pszFormat, ...)
{
    TCHAR buf[1024];
#ifdef UNICODE
#define DBG_FMT TEXT("[%ls](%lu): ")
#else
#define DBG_FMT TEXT("[%s](%lu): ")
#endif
    wsprintf(buf, DBG_FMT, m_szServiceName, GetCurrentThreadId());
	va_list arglist;
	va_start(arglist, pszFormat);
    _vstprintf(&buf[_tcslen(buf)], pszFormat, arglist);
	va_end(arglist);
    _tcscat(buf, TEXT("\n"));
    OutputDebugString(buf);
}


/////////////////////////////////////////////////////////////////////////////
// CNtsvcCtrl static members

DWORD CNtsvcCtrl::ServiceThread(LPVOID)
{
	BOOL b;

    // Get a pointer to the C++ object
    CNtsvcCtrl* pService = m_pThis;

    SERVICE_TABLE_ENTRY st[] = {
        {m_pThis->m_szServiceName, ServiceMain},
        {NULL, NULL}
    };

    m_pThis->DebugMsg(TEXT("Calling StartServiceCtrlDispatcher()"));

	// Call the services dispatcher. This call blocks
    // until the service stops
    b = ::StartServiceCtrlDispatcher(st);

    m_pThis->DebugMsg(TEXT("Returned from StartServiceCtrlDispatcher()"));

    return b;
}


// 
void CNtsvcCtrl::ServiceMain(DWORD dwArgc, LPTSTR* lpszArgv)
{
    // Get a pointer to the C++ object
    CNtsvcCtrl* pService = m_pThis;

    pService->DebugMsg(TEXT("Entering CNtsvcCtrl::ServiceMain()"));

    // Register the control request handler
    pService->m_Status.dwCurrentState = SERVICE_START_PENDING;
    pService->m_hServiceStatus = RegisterServiceCtrlHandler(pService->m_szServiceName,
                                                           Handler);
    if (pService->m_hServiceStatus == NULL) 
	{
        pService->LogEvent(EVENTLOG_ERROR_TYPE, EVMSG_CTRLHANDLERNOTINSTALLED, NULL);
        return;
    }

	pService->m_hStopEvent = CreateEvent(
        NULL,    // no security attributes
        TRUE,    // manual reset event
        FALSE,   // not-signalled
        NULL);   // no name

	// notify the user that the service is starting
	m_pThis->PostMessage(CWM_START, 0, 0);

	// The service thread will wait indefinately until
	// the service control manager or the user signals
	// the stop event. This event is set inside the StopService
	// method.
	WaitForSingleObject(m_pThis->m_hStopEvent, INFINITE);

    // Tell the service manager we are stopped
    pService->SetStatus(SERVICE_STOPPED);
	pService->ReportStatus();

    pService->DebugMsg(TEXT("Leaving CNtsvcCtrl::ServiceMain()"));
}


void CNtsvcCtrl::Handler(DWORD dwOpcode)
{
    // Get a pointer to the object
    CNtsvcCtrl* pService = m_pThis;


    pService->DebugMsg("CNTService::Handler(%lu)", dwOpcode);
    switch (dwOpcode) {
    case SERVICE_CONTROL_STOP: // 1
        pService->SetStatus(SERVICE_STOP_PENDING);
        break;

    case SERVICE_CONTROL_PAUSE:
        pService->SetStatus(SERVICE_PAUSE_PENDING);
        break;

    default:
		// custom service control
        break;
    }
	
	// Report current status
    m_pThis->ReportStatus();

	// deffer message
	m_pThis->PostMessage (CWM_HANDLER,dwOpcode, 0);
}



/////////////////////////////////////////////////////////////////////////////
// CNtsvcCtrl methods
//
BOOL CNtsvcCtrl::Install() 
{
    // Open the Service Control Manager
    SC_HANDLE hSCM = ::OpenSCManager(NULL, // local machine
                                     NULL, // ServicesActive database
                                     SC_MANAGER_ALL_ACCESS); // full access
    if (!hSCM) return FALSE;

    // Get the executable file path
    TCHAR szFilePath[_MAX_PATH];
    ::GetModuleFileName(NULL, szFilePath, sizeof(szFilePath));

	DWORD dwDesiredAccess = SERVICE_WIN32_OWN_PROCESS;
	if (m_interactive)
		dwDesiredAccess |= SERVICE_INTERACTIVE_PROCESS;

	LPSTR lpServiceAccount = NULL;
	if (m_account.GetLength() > 0 )
		lpServiceAccount = m_account.GetBuffer(m_account.GetLength()+1);

	LPSTR lpPassword = NULL;
	if (m_password.GetLength() > 0)
		lpPassword = m_password.GetBuffer(m_password.GetLength()+1);

    // Create the service
    SC_HANDLE hService = ::CreateService(hSCM,
                                         m_szServiceName,
                                         m_displayName,
                                         SERVICE_ALL_ACCESS,
                                         dwDesiredAccess,
                                         m_dwStartMode,        // start condition
                                         SERVICE_ERROR_NORMAL,
                                         szFilePath,
                                         NULL,
                                         NULL,
                                         NULL,
                                         lpServiceAccount,
                                         lpPassword);

	if (lpServiceAccount)
		m_account.ReleaseBuffer();

	if (lpPassword)
		m_password.ReleaseBuffer();

    if (!hService) {
        ::CloseServiceHandle(hSCM);
        return FALSE;
    }

    // make registry entries to support logging messages
    // Add the source name as a subkey under the Application
    // key in the EventLog service portion of the registry.
	
	// this time, we get the path of the control
    ::GetModuleFileName(AfxGetInstanceHandle(), szFilePath, sizeof(szFilePath));

    TCHAR szKey[256];
    HKEY hKey = NULL;
    _tcscpy(szKey, TEXT("SYSTEM\\CurrentControlSet\\Services\\EventLog\\Application\\"));
    _tcscat(szKey, m_szServiceName);
    if (::RegCreateKey(HKEY_LOCAL_MACHINE, szKey, &hKey) != ERROR_SUCCESS) {
        ::CloseServiceHandle(hService);
        ::CloseServiceHandle(hSCM);
        return FALSE;
    }

    // Add the Event ID message-file name to the 'EventMessageFile' subkey.
    ::RegSetValueEx(hKey,
                    TEXT("EventMessageFile"),
                    0,
                    REG_EXPAND_SZ, 
                    (CONST BYTE*)szFilePath,
                    _tcslen(szFilePath) + 1);     

    // Set the supported types flags.
    DWORD dwData = EVENTLOG_ERROR_TYPE | EVENTLOG_WARNING_TYPE | EVENTLOG_INFORMATION_TYPE;
    ::RegSetValueEx(hKey,
                    TEXT("TypesSupported"),
                    0,
                    REG_DWORD,
                    (CONST BYTE*)&dwData,
                     sizeof(DWORD));
    ::RegCloseKey(hKey);

    LogEvent(EVENTLOG_INFORMATION_TYPE, EVMSG_INSTALLED, m_displayName);

    // tidy up
    ::CloseServiceHandle(hService);
    ::CloseServiceHandle(hSCM);

    return TRUE;
}

BOOL CNtsvcCtrl::Uninstall() 
{
    BOOL bResult = FALSE;

    // Open the Service Control Manager
    SC_HANDLE hSCM = ::OpenSCManager(NULL, // local machine
                                     NULL, // ServicesActive database
                                     SC_MANAGER_ALL_ACCESS); // full access
    if (!hSCM) return FALSE;

    SC_HANDLE hService = ::OpenService(hSCM,
                                       m_szServiceName,
                                       DELETE);
    if (hService) {
        if (::DeleteService(hService)) {
            LogEvent(EVENTLOG_INFORMATION_TYPE, EVMSG_REMOVED, m_szServiceName);
            bResult = TRUE;
        } else {
            LogEvent(EVENTLOG_ERROR_TYPE, EVMSG_NOTREMOVED, m_szServiceName);
        }
        ::CloseServiceHandle(hService);
    }
    
    ::CloseServiceHandle(hSCM);
    
	return bResult;
}


BOOL CNtsvcCtrl::StartService()
{
	BOOL bSuccess = FALSE;

	// initialize service status
    m_Status.dwServiceType = SERVICE_WIN32_OWN_PROCESS;
    m_Status.dwCurrentState = SERVICE_STOPPED;
    m_Status.dwControlsAccepted = SERVICE_ACCEPT_STOP | m_controlsAccepted;
    m_Status.dwWin32ExitCode = 0;
    m_Status.dwServiceSpecificExitCode = 0;
    m_Status.dwCheckPoint = 0;
    m_Status.dwWaitHint = (DWORD) m_waitHint;

    // in debug mode, don’t create a service thread
	if (m_bIsDebug)
	{
		m_pThis->PostMessage(CWM_START, 0, 0);
		bSuccess = TRUE;
	}
	else
	{
		DWORD dwThreadId;
		
        // create a service thread that to communicate with
        // the services dispatcher

		m_hServiceThread = CreateThread(
			NULL,		// pointer to thread security attributes  
			0,			// initial thread stack size, in bytes 
			ServiceThread,	// pointer to thread function 
			NULL,		// argument for new thread 
			0,			// creation flags 
			&dwThreadId 	// pointer to returned thread identifier 
		);	

		if (m_hServiceThread)
		{
			bSuccess = TRUE;
		}
	}

	return bSuccess;
}

void CNtsvcCtrl::StopService() 
{
    FireStop();

	m_bIsRunning = FALSE;

	// Keep service contorl manager informed
	ReportStatus();

	// Release the thread created by the
	// service control manager. When that thread exits,
	// the service will be considered as 'stopped'
    if ( m_hStopEvent )
        SetEvent(m_hStopEvent);

	// Wait for the service thread to exit. This ensures
	// that service control manager gets all its notifications
	// before the container frees the control
	if ( m_hServiceThread)
	{
		WaitForSingleObject (m_hServiceThread, INFINITE);
		CloseHandle (m_hServiceThread);
		m_hServiceThread = NULL;
	}

	// Signal that the user wants to end the service.
	if (!m_bIsDebug)
		LogEvent(EVENTLOG_INFORMATION_TYPE, EVMSG_STOPPED, "");

	PostQuitMessage(0);

}

BOOL CNtsvcCtrl::Running() 
{
	return m_bIsRunning;
}


BOOL CNtsvcCtrl::LogEvent(short wType, long dwID, LPCTSTR Message) 
{
    // Check the event source has been registered and if
    // not then register it now
    if (!m_hEventSource) {
        m_hEventSource = ::RegisterEventSource(NULL,  // local machine
                                               m_szServiceName); // source name
    }

    if (m_hEventSource) {
        ::ReportEvent(m_hEventSource,
                      wType,
                      0,
                      dwID,
                      NULL, // sid
                      1,
                      0,
                      &Message,
                      NULL);
    }
	else {
		return FALSE;
	}

	return TRUE;
}


VARIANT CNtsvcCtrl::GetAllSettings(LPCTSTR section) 
{
	VARIANT vaResult;
	VariantInit(&vaResult);

	SAFEARRAYBOUND rgsaBounds[2];
	SAFEARRAY* psaSettings;

	DWORD cbValue;
	DWORD cbData;
	DWORD dwType;
    HKEY hKey = NULL;
	TCHAR szKey[256];
	TCHAR szValue [256];
	int n;

	HRESULT hr;

	rgsaBounds[0].lLbound = 0;
	rgsaBounds[0].cElements = 2;
	rgsaBounds[1].lLbound = 0;
	rgsaBounds[1].cElements = 0;

	psaSettings = SafeArrayCreate (VT_BSTR, 2, rgsaBounds);

	wsprintf (szKey, "%s%s\\", m_szSvcRegKey, section);
    if (::RegOpenKey(HKEY_LOCAL_MACHINE, szKey, &hKey) == ERROR_SUCCESS )
	{
		n = 0;

		cbValue = sizeof(szValue) - 1;
		while (::RegEnumValue(hKey, n, szValue, &cbValue, 0, &dwType, NULL, &cbData) == ERROR_SUCCESS )
		{
			// terminate value string
			szValue [cbValue] = '\0';

			CString strValue (szValue);

			// get data
			CString strData;
			LPSTR szBuffer = strData.GetBuffer(cbData + 1);

			if (szBuffer)
			{
				::RegEnumValue(hKey, n, szValue, &cbValue, 0, &dwType, (BYTE*)szBuffer, &cbData);
				strData.ReleaseBuffer();
			}

			// re-dimension least significant dimension (rows)
			rgsaBounds[1].cElements = n + 1;
			hr = SafeArrayRedim (psaSettings, &rgsaBounds[1]);

			if (SUCCEEDED(hr))
			{
				long rgIndices[2];
				BSTR bstrElement;
				rgIndices[1] = n;

				rgIndices[0] = 0;
				bstrElement = strValue.AllocSysString();
				hr = SafeArrayPutElement (psaSettings, rgIndices, bstrElement);

				rgIndices[0] = 1;
				bstrElement = strData.AllocSysString();
				hr = SafeArrayPutElement (psaSettings, rgIndices, bstrElement);
			}
			else
			{
				// TODO: should throw error
				break;
			}


			// increment enum counter
			n ++;

			// reset input parameters for while loop
			cbValue = sizeof(szValue) - 1;
		}
	}
    
    ::RegCloseKey(hKey);
	
	vaResult.vt = VT_ARRAY|VT_BSTR;
	vaResult.parray = psaSettings;

	return vaResult;
}

void CNtsvcCtrl::SaveSetting(LPCTSTR section, LPCTSTR key, LPCTSTR setting) 
{
    HKEY hKey = NULL;
	TCHAR szKey[256];

	wsprintf (szKey, "%s%s\\", m_szSvcRegKey, section);
    if (::RegCreateKey(HKEY_LOCAL_MACHINE, szKey, &hKey) != ERROR_SUCCESS) 
	{
		// cannot create key
		ThrowError(ERROR_ACCESS_DENIED);
        return;
    }

    // Add the Event ID message-file name to the 'EventMessageFile' subkey.
    ::RegSetValueEx(hKey,
                    key,
                    0,
                    REG_EXPAND_SZ, 
                    (CONST BYTE*)setting,
                    _tcslen(setting) + 1);     

    ::RegCloseKey(hKey);
}


void CNtsvcCtrl::DeleteSetting(LPCTSTR section, const VARIANT FAR& key) 
{
    HKEY hKey = NULL;
	SCODE sc = 0;

	if (::RegOpenKey(HKEY_LOCAL_MACHINE, m_szSvcRegKey, &hKey) != ERROR_SUCCESS)
		ThrowError(ERROR_ACCESS_DENIED);

	if (key.vt != VT_ERROR)
	{
		if (key.vt == VT_BSTR)
		{
			HKEY hSubKey;
			CString strKey = key.bstrVal;
			if (::RegOpenKey(hKey, m_szSvcRegKey, &hSubKey) == ERROR_SUCCESS)
			{
				RegDeleteKey (hSubKey, strKey);
				RegCloseKey (hKey);
			}
			else
			{
				sc = ERROR_ACCESS_DENIED;
			}
		}
		else
		{
			sc = DISP_E_TYPEMISMATCH;
		}
	}
	else
	{
		::RegDeleteKey (hKey, section);
	}
    ::RegCloseKey (hKey);

	if (sc)
	{
		ThrowError(sc);
	}
}

BSTR CNtsvcCtrl::GetSetting(LPCTSTR section, LPCTSTR key, const VARIANT FAR& def) 
{
	CString strResult;

	DWORD cbValue;
	DWORD dwType;
    HKEY hKey = NULL;
	TCHAR szKey[256];

	wsprintf (szKey, "%s%s\\", m_szSvcRegKey, section);
    if (::RegOpenKey(HKEY_LOCAL_MACHINE, szKey, &hKey) == ERROR_SUCCESS
		&& ::RegQueryValueEx(hKey, key, 0, &dwType, NULL, &cbValue) == ERROR_SUCCESS )
	{
		LPSTR szBuffer = strResult.GetBuffer(cbValue + 1);

		if (szBuffer)
		{
			::RegQueryValueEx(hKey, key, 0, &dwType, (BYTE*)szBuffer, &cbValue);
			strResult.ReleaseBuffer();
		}
    }
	else
	{
		if (def.vt == VT_BSTR)
			strResult = def.bstrVal;
		else
			ThrowError(DISP_E_TYPEMISMATCH);
	}
    
    ::RegCloseKey(hKey);
	
	return strResult.AllocSysString();
}

void CNtsvcCtrl::ReportSrvStatus() 
{
	ReportStatus();
}

long CNtsvcCtrl::GetServiceStatusHandle() 
{
	return (long)m_hServiceStatus;
}

/////////////////////////////////////////////////////////////////////////////
// CNtsvcCtrl property handlers

void CNtsvcCtrl::OnDisplayNameChanged() 
{
	SetModifiedFlag();
}


void CNtsvcCtrl::OnDependenciesChanged() 
{
	SetModifiedFlag();
}

void CNtsvcCtrl::OnLoadOrderGroupChanged() 
{
	SetModifiedFlag();
}

long CNtsvcCtrl::GetStartMode() 
{
	return m_dwStartMode;
}

void CNtsvcCtrl::SetStartMode(long nNewValue) 
{
	m_dwStartMode = (DWORD) nNewValue;

	SetModifiedFlag();
}

BSTR CNtsvcCtrl::GetServiceName() 
{
	CString strResult;

	strResult = m_szServiceName;

	return strResult.AllocSysString();
}

void CNtsvcCtrl::SetServiceName(LPCTSTR lpszNewValue) 
{
	InternalSetService(lpszNewValue);

	SetModifiedFlag();
}

BOOL CNtsvcCtrl::GetDebug() 
{
	return m_bIsDebug;
}

void CNtsvcCtrl::SetDebug(BOOL bNewValue) 
{
	m_bIsDebug = bNewValue;

	SetModifiedFlag();
}

void CNtsvcCtrl::OnControlsAcceptedChanged() 
{
	SetModifiedFlag();
}

void CNtsvcCtrl::OnInteractiveChanged() 
{
	SetModifiedFlag();
}

void CNtsvcCtrl::OnAccountChanged() 
{
	SetModifiedFlag();
}

void CNtsvcCtrl::OnPasswordChanged() 
{
	SetModifiedFlag();
}

void CNtsvcCtrl::OnWaitHintChanged() 
{
  m_Status.dwWaitHint = (DWORD) m_waitHint;
	SetModifiedFlag();
}
