// SimpleAddin.h : CSimpleAddin ������

#pragma once
#include "resource.h"       // ������



#include "PPTAddin_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE ƽ̨(�粻�ṩ��ȫ DCOM ֧�ֵ� Windows Mobile ƽ̨)���޷���ȷ֧�ֵ��߳� COM ���󡣶��� _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA ��ǿ�� ATL ֧�ִ������߳� COM ����ʵ�ֲ�����ʹ���䵥�߳� COM ����ʵ�֡�rgs �ļ��е��߳�ģ���ѱ�����Ϊ��Free����ԭ���Ǹ�ģ���Ƿ� DCOM Windows CE ƽ̨֧�ֵ�Ψһ�߳�ģ�͡�"
#endif

using namespace ATL;
# include "CApplication.h"
# include "CPresentation.h"
# include "CPresentations.h"

#import "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\MSO.DLL" \
	no_implementation \
	rename("RGB", "PPTRGB") \
	rename("DocumentProperties", "PPTDocumentProperties") \
	rename("SearchPath", "PPTSearchPath")

#import "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.OLB" \
	no_implementation

#import "C:\Program Files (x86)\Microsoft Office\Office15\MSPPT.OLB" \
	rename("RGB", "PPTRGB") 

_ATL_FUNC_INFO PresentationOpenInfo = { CC_STDCALL, VT_EMPTY, 1,{VT_DISPATCH | VT_BYREF}};
_ATL_FUNC_INFO PresentationCloseInfo = { CC_STDCALL, VT_EMPTY, 1,{VT_DISPATCH | VT_BYREF}};

// CSimpleAddin

class ATL_NO_VTABLE CSimpleAddin :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CSimpleAddin, &CLSID_SimpleAddin>,
	public IDispatchImpl<ISimpleAddin, &IID_ISimpleAddin, &LIBID_PPTAddinLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &LIBID_AddInDesignerObjects, /* wMajor = */ 1>,
	public IDispEventSimpleImpl<10, CSimpleAddin, &__uuidof(PowerPoint::EApplication)>,
	public IDispEventSimpleImpl<11, CSimpleAddin, &__uuidof(PowerPoint::EApplication)>
{
public:
	CSimpleAddin()
	{
	}

	DECLARE_REGISTRY_RESOURCEID(IDR_SIMPLEADDIN)


	BEGIN_COM_MAP(CSimpleAddin)
		COM_INTERFACE_ENTRY(ISimpleAddin)
		COM_INTERFACE_ENTRY2(IDispatch, _IDTExtensibility2)
		COM_INTERFACE_ENTRY(_IDTExtensibility2)
	END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:




	// _IDTExtensibility2 Methods
public:
	STDMETHOD(OnConnection)(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom);
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom);
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom);
	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom);
	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom);

public:
	typedef IDispEventSimpleImpl<10, CSimpleAddin, &__uuidof(PowerPoint::EApplication)> PresentationOpenEvent;
	typedef IDispEventSimpleImpl<11, CSimpleAddin, &__uuidof(PowerPoint::EApplication)> PresentationCloseEvent;
	
	LPDISPATCH m_Application;

	void _stdcall PresentationOpen(_In_ LPDISPATCH Pres);
	void _stdcall PresentationClose(_In_ LPDISPATCH Pres);

	BEGIN_SINK_MAP(CSimpleAddin)
		SINK_ENTRY_INFO(10, __uuidof(PowerPoint::EApplication), 0x000007D6, PresentationOpen, &PresentationOpenInfo)
		SINK_ENTRY_INFO(11, __uuidof(PowerPoint::EApplication), 0x000007D4, PresentationClose, &PresentationCloseInfo)
	END_SINK_MAP()
};

OBJECT_ENTRY_AUTO(__uuidof(SimpleAddin), CSimpleAddin)
