// dllmain.cpp : DllMain 的实现。

#include "stdafx.h"
#include "resource.h"
#include "PPTAddin_i.h"
#include "dllmain.h"
#include "xdlldata.h"

CPPTAddinModule _AtlModule;

class CPPTAddinApp : public CWinApp
{
public:

// 重写
	virtual BOOL InitInstance();
	virtual int ExitInstance();

	DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(CPPTAddinApp, CWinApp)
END_MESSAGE_MAP()

CPPTAddinApp theApp;

BOOL CPPTAddinApp::InitInstance()
{
#ifdef _MERGE_PROXYSTUB
	if (!PrxDllMain(m_hInstance, DLL_PROCESS_ATTACH, NULL))
		return FALSE;
#endif
	return CWinApp::InitInstance();
}

int CPPTAddinApp::ExitInstance()
{
	return CWinApp::ExitInstance();
}
