// dllmain.h : 模块类的声明。

class CPPTAddinModule : public ATL::CAtlDllModuleT< CPPTAddinModule >
{
public :
	DECLARE_LIBID(LIBID_PPTAddinLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_PPTADDIN, "{BB6B1CCA-07D5-4F0D-86C9-18658D0C7F17}")
};

extern class CPPTAddinModule _AtlModule;
