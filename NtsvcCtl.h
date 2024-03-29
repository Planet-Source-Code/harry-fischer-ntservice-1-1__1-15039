#if !defined(AFX_NTSVCCTL_H__C32049B3_E8D3_11D4_9D82_00D0B73B61A4__INCLUDED_)
#define AFX_NTSVCCTL_H__C32049B3_E8D3_11D4_9D82_00D0B73B61A4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// NtsvcCtl.h : Declaration of the CNtsvcCtrl ActiveX Control class.

/////////////////////////////////////////////////////////////////////////////
// CNtsvcCtrl : See NtsvcCtl.cpp for implementation.

class CNtsvcCtrl : public COleControl
{
	DECLARE_DYNCREATE(CNtsvcCtrl)

// Constructor
public:
	CNtsvcCtrl();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CNtsvcCtrl)
	public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	//}}AFX_VIRTUAL

// Implementation
protected:
	~CNtsvcCtrl();

	enum {SERVICE_NAME_LEN = 64};
	enum {SERVICE_REG_LEN = 256};

	DECLARE_OLECREATE_EX(CNtsvcCtrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CNtsvcCtrl)      // GetTypeInfo
	DECLARE_OLECTLTYPE(CNtsvcCtrl)		// Type name and misc status

// Members
  TCHAR m_szServiceName[SERVICE_NAME_LEN];
  HANDLE m_hEventSource;
  SERVICE_STATUS_HANDLE m_hServiceStatus;
  SERVICE_STATUS m_Status;
  DWORD m_dwStartMode;
  BOOL m_bIsRunning;
  BOOL m_bIsDebug;

  TCHAR m_szSvcRegKey[SERVICE_REG_LEN];

	HANDLE m_hStopEvent;			// signaled when the user wants to stop the service
	HANDLE m_hServiceThread;		// handle to service control thread

  // static data
  static CNtsvcCtrl* m_pThis; // nasty hack to get object ptr

	// painting stuff
	SIZEL m_Size;
	CBitmap m_bmIcon;
	CBitmap m_bmMask;
	CBrush m_brBkgnd;
	CSize m_pixSize;


	// member functions
	long GetStatus();
	void SetStatus(long nNewValue);
	void ReportStatus();
	void InternalSetService(LPCTSTR lpszNewValue);
	void DebugMsg(const TCHAR* pszFormat, ...);

  // static member functions
  static DWORD WINAPI ServiceThread(LPVOID);
  static void WINAPI ServiceMain(DWORD dwArgc, LPTSTR* lpszArgv);
  static void WINAPI Handler(DWORD dwOpcode);

// Message maps
	//{{AFX_MSG(CNtsvcCtrl)
	afx_msg BOOL OnEraseBkgnd(CDC* pDC);
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	BOOL OnSetExtent(LPSIZEL lpSizeL);

// Dispatch maps
	//{{AFX_DISPATCH(CNtsvcCtrl)
	CString m_account;
	afx_msg void OnAccountChanged();
	long m_controlsAccepted;
	afx_msg void OnControlsAcceptedChanged();
	CString m_LoadOrderGroup;
	afx_msg void OnLoadOrderGroupChanged();
	CString m_Dependencies;
	afx_msg void OnDependenciesChanged();
	CString m_displayName;
	afx_msg void OnDisplayNameChanged();
	BOOL m_interactive;
	afx_msg void OnInteractiveChanged();
	CString m_password;
	afx_msg void OnPasswordChanged();
	long m_waitHint;
	afx_msg void OnWaitHintChanged();
	afx_msg BOOL GetDebug();
	afx_msg void SetDebug(BOOL bNewValue);
	afx_msg BSTR GetServiceName();
	afx_msg void SetServiceName(LPCTSTR lpszNewValue);
	afx_msg long GetStartMode();
	afx_msg void SetStartMode(long nNewValue);
	afx_msg BOOL Install();
	afx_msg BOOL LogEvent(short EventType, long ID, LPCTSTR Message);
	afx_msg BOOL Running();
	afx_msg BOOL Uninstall();
	afx_msg void StopService();
	afx_msg BOOL StartService();
	afx_msg VARIANT GetAllSettings(LPCTSTR section);
	afx_msg void SaveSetting(LPCTSTR section, LPCTSTR key, LPCTSTR setting);
	afx_msg void DeleteSetting(LPCTSTR section, const VARIANT FAR& key);
	afx_msg BSTR GetSetting(LPCTSTR section, LPCTSTR key, const VARIANT FAR& def);
	afx_msg void ReportSrvStatus();
	afx_msg long GetServiceStatusHandle();
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()

	afx_msg void AboutBox();

	afx_msg LRESULT OnStart(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnHandler(WPARAM wParam, LPARAM lParam);

// Event maps
	//{{AFX_EVENT(CNtsvcCtrl)
	void FireStop()
		{FireEvent(eventidStop,EVENT_PARAM(VTS_NONE));}
	void FirePause(BOOL FAR* Success)
		{FireEvent(eventidPause,EVENT_PARAM(VTS_PBOOL), Success);}
	void FireContinue(BOOL FAR* Success)
		{FireEvent(eventidContinue,EVENT_PARAM(VTS_PBOOL), Success);}
	void FireStart(BOOL FAR* Success)
		{FireEvent(eventidStart,EVENT_PARAM(VTS_PBOOL), Success);}
	void FireControl(long lEvent)
		{FireEvent(eventidControl,EVENT_PARAM(VTS_I4), lEvent);}
	//}}AFX_EVENT
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	enum {
	//{{AFX_DISP_ID(CNtsvcCtrl)
	dispidAccount = 1L,
	dispidControlsAccepted = 2L,
	dispidDebug = 9L,
	dispidLoadOrderGroup = 3L,
	dispidDependencies = 4L,
	dispidDisplayName = 5L,
	dispidInteractive = 6L,
	dispidPassword = 7L,
	dispidServiceName = 10L,
	dispidStartMode = 11L,
	dispidWaitHint = 8L,
	dispidDeleteSetting = 20L,
	dispidGetAllSettings = 18L,
	dispidGetSetting = 21L,
	dispidInstall = 12L,
	dispidLogEvent = 13L,
	dispidRunning = 14L,
	dispidSaveSetting = 19L,
	dispidStartService = 17L,
	dispidStopService = 16L,
	dispidUninstall = 15L,
	dispidReportStatus = 22L,
	dispidGetServiceStatusHandle = 23L,
	eventidContinue = 3L,
	eventidControl = 5L,
	eventidPause = 2L,
	eventidStart = 4L,
	eventidStop = 1L,
	//}}AFX_DISP_ID
	};
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_NTSVCCTL_H__C32049B3_E8D3_11D4_9D82_00D0B73B61A4__INCLUDED)
