// SimpleAddin.cpp : CSimpleAddin µÄÊµÏÖ

#include "stdafx.h"
#include "SimpleAddin.h"


// CSimpleAddin


STDMETHODIMP CSimpleAddin::OnConnection(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom)
{
	OutputDebugString(TEXT("PPT OnConnection"));
	m_Application = Application;
	PresentationOpenEvent::DispEventAdvise(m_Application, &__uuidof(PowerPoint::EApplication));
	PresentationCloseEvent::DispEventAdvise(m_Application, &__uuidof(PowerPoint::EApplication));
	return S_OK;
}
STDMETHODIMP CSimpleAddin::OnDisconnection(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
{
	OutputDebugString(TEXT("PPT OnDisconnection"));
	PresentationOpenEvent::DispEventUnadvise(m_Application, &__uuidof(PowerPoint::EApplication));
	PresentationCloseEvent::DispEventUnadvise(m_Application, &__uuidof(PowerPoint::EApplication));
	return S_OK;
}
STDMETHODIMP CSimpleAddin::OnAddInsUpdate(SAFEARRAY * * custom)
{
	OutputDebugString(TEXT("PPT OnAddInsUpdate"));
	return S_OK;
}
STDMETHODIMP CSimpleAddin::OnStartupComplete(SAFEARRAY * * custom)
{
	OutputDebugString(TEXT("PPT OnStartupComplete"));
	return S_OK;
}
STDMETHODIMP CSimpleAddin::OnBeginShutdown(SAFEARRAY * * custom)
{
	OutputDebugString(TEXT("PPT OnBeginShutdown"));
	return S_OK;
}

void CSimpleAddin::PresentationOpen(_In_ LPDISPATCH Pres)
{
	OutputDebugString(TEXT("PPT PresentationOpen"));
}
void CSimpleAddin::PresentationClose(_In_ LPDISPATCH Pres)
{
	OutputDebugString(TEXT("PPT PresentationClose"));
}