#if !defined(AFX_NTSVC_H__C32049AB_E8D3_11D4_9D82_00D0B73B61A4__INCLUDED_)
#define AFX_NTSVC_H__C32049AB_E8D3_11D4_9D82_00D0B73B61A4__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// Ntsvc.h : main header file for NTSVC.DLL

#if !defined( __AFXCTL_H__ )
	#error include 'afxctl.h' before including this file
#endif

#include "resource.h"       // main symbols
#include "ntservmsg.h"

/////////////////////////////////////////////////////////////////////////////
// CNtsvcApp : See Ntsvc.cpp for implementation.

class CNtsvcApp : public COleControlModule
{
public:
	BOOL InitInstance();
	int ExitInstance();
};

extern const GUID CDECL _tlid;
extern const WORD _wVerMajor;
extern const WORD _wVerMinor;

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_NTSVC_H__C32049AB_E8D3_11D4_9D82_00D0B73B61A4__INCLUDED)
