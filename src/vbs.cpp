#include <Windows.h>
#include <activscp.h>
#include <atlbase.h>
#include <tchar.h>
#include <iostream>
#include "rust/cxx.h"
#include "inline-vbs/include/vbs.h"
#include "comdef.h"

class CSimpleScriptSite :
        public IActiveScriptSite,
        public IActiveScriptSiteWindow
{
public:
    CSimpleScriptSite() : m_cRefCount(1), m_hWnd(NULL) { }

    STDMETHOD_(ULONG, AddRef)();
    STDMETHOD_(ULONG, Release)();
    STDMETHOD(QueryInterface)(REFIID riid, void **ppvObject);

    STDMETHOD(GetLCID)(LCID *plcid){ *plcid = 0; return S_OK; }
    STDMETHOD(GetItemInfo)(LPCOLESTR pstrName, DWORD dwReturnMask, IUnknown **ppiunkItem, ITypeInfo **ppti) { return TYPE_E_ELEMENTNOTFOUND; }
    STDMETHOD(GetDocVersionString)(BSTR *pbstrVersion) { *pbstrVersion = SysAllocString(L"1.0"); return S_OK; }
    STDMETHOD(OnScriptTerminate)(const VARIANT *pvarResult, const EXCEPINFO *pexcepinfo) { return S_OK; }
    STDMETHOD(OnStateChange)(SCRIPTSTATE ssScriptState) { return S_OK; }
    STDMETHOD(OnScriptError)(IActiveScriptError *pIActiveScriptError) { return S_OK; }
    STDMETHOD(OnEnterScript)(void) { return S_OK; }
    STDMETHOD(OnLeaveScript)(void) { return S_OK; }

    STDMETHOD(GetWindow)(HWND *phWnd) { *phWnd = m_hWnd; return S_OK; }
    STDMETHOD(EnableModeless)(BOOL fEnable) { return S_OK; }

    HRESULT SetWindow(HWND hWnd) { m_hWnd = hWnd; return S_OK; }

public:
    LONG m_cRefCount;
    HWND m_hWnd;
};

STDMETHODIMP_(ULONG) CSimpleScriptSite::AddRef()
{
    return InterlockedIncrement(&m_cRefCount);
}

STDMETHODIMP_(ULONG) CSimpleScriptSite::Release()
{
    if (!InterlockedDecrement(&m_cRefCount))
    {
        delete this;
        return 0;
    }

    return m_cRefCount;
}

STDMETHODIMP CSimpleScriptSite::QueryInterface(REFIID riid, void **ppvObject)
{
    if (riid == IID_IUnknown || riid == IID_IActiveScriptSiteWindow)
    {
        *ppvObject = (IActiveScriptSiteWindow *) this;
        AddRef();
        return NOERROR;
    }

    if (riid == IID_IActiveScriptSite)
    {
        *ppvObject = (IActiveScriptSite *) this;
        AddRef();
        return NOERROR;
    }

    return E_NOINTERFACE;
}
CComPtr<IActiveScriptParse> script_parser;
CComPtr<IActiveScript> script_engine;
CSimpleScriptSite* script_site;
bool initialized = false;

#define TRY(x) if (FAILED(hr = x)) return hr;

rust::String error_to_string(int hr)
{
    _com_error err(hr);
    return std::string(err.ErrorMessage());
}

int init()
{
    if (initialized)
        return S_OK;

    HRESULT hr;
    TRY(CoInitializeEx(nullptr, COINIT_MULTITHREADED));

    script_site = new CSimpleScriptSite();

    TRY(script_engine.CoCreateInstance(OLESTR("VBScript")));
    TRY(script_engine->SetScriptSite(script_site));
    TRY(script_engine->QueryInterface(&script_parser));
    TRY(script_parser->InitNew());

    initialized = true;

    return S_OK;
}

wchar_t* to_wide(rust::Str str)
{
    int wide_len = MultiByteToWideChar(CP_UTF8, 0, str.data(), str.length(), nullptr, 0);
    wchar_t* wcode = new wchar_t[wide_len + 1];
    MultiByteToWideChar(CP_UTF8, 0, str.data(), -1, wcode, wide_len);
    wcode[wide_len] = 0;
    return wcode;
}

int parse(rust::Str code, VARIANT* output)
{
    wchar_t* wcode = to_wide(code);

    CComVariant result;
    EXCEPINFO ei = { };

    int hr = script_parser->ParseScriptText(
            &wcode[0],
            nullptr,
            nullptr,
            nullptr,
            0,
            0,
            output ? SCRIPTTEXT_ISEXPRESSION : 0,
            output ? output : &result,
            &ei);

    delete[] wcode;

    return hr;
}

int set_variable(rust::Str name, char* val)
{
    HRESULT hr;

    IDispatch* objPtr;
    script_engine->GetScriptDispatch(nullptr, &objPtr);

    DISPID varid;
    wchar_t* wname = to_wide(name);

    TRY(objPtr->GetIDsOfNames(IID_NULL, &wname, 1, LOCALE_USER_DEFAULT, &varid));

    DISPPARAMS dspp;
    ZeroMemory(&dspp, sizeof(dspp));
    dspp.cArgs = dspp.cNamedArgs = 1;
    DISPID dispPropPut = DISPID_PROPERTYPUT;
    dspp.rgdispidNamedArgs = &dispPropPut;
    VARIANT* var = (VARIANT*) val;
    dspp.rgvarg = var;

    TRY(objPtr->Invoke(varid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dspp, nullptr, nullptr, nullptr));
    VariantClear(var);

    TRY(objPtr->Release());

    return S_OK;
}

int parse_wrapper(rust::Str code, char* output)
{
    return parse(code, (VARIANT*) output);
}

int close()
{
    if (!initialized)
        return S_OK;

    HRESULT hr;

    script_parser.p = nullptr; // TODO: this is a hack
    script_engine.p = nullptr; // TODO: but so is COM anyway
    TRY(script_site->Release());

    ::CoUninitialize();

    initialized = false;
    return S_OK;
}

class VBSContext
{
public:
    ~VBSContext()
    {
        close();
    }
};

VBSContext sentinel; // cleaner atexit() implementation