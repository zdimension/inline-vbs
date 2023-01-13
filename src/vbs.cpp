#include <Windows.h>
#include <activscp.h>
#include <atlbase.h>
#include <tchar.h>
#include <iostream>
#include <unordered_map>
#include <string>
#include "rust/cxx.h"
#include "inline-vbs/src/lib.rs.h"
#include "comdef.h"

#define TRY(x) if (FAILED(hr = x)) { std::cout << "FAIL at line" << __LINE__ << std::endl; return hr; }

const wchar_t *LangName(ScriptLang lang) {
    switch (lang) {
        case ScriptLang::VBScript:
            return L"VBScript";
        case ScriptLang::JScript:
            return L"JScript";
        case ScriptLang::Perl:
            return L"Perl";
        case ScriptLang::Ruby:
            return L"Ruby";
        default:
            std::cout << "INVALID LANG" << std::endl;
            exit(1);
    }
}

const wchar_t *LangClsId(ScriptLang lang) {
    switch (lang) {
        case ScriptLang::VBScript:
            return L"{B54F3741-5B07-11CF-A4B0-00AA004A55E8}";
        case ScriptLang::JScript:
            return L"{16D51579-A30B-4C8B-A276-0FF4DC41E755}"; // JScript9.dll (Chakra)
        case ScriptLang::Perl:
            return L"{F8D77580-0F09-11D0-AA61-3C284E000000}"; // PerlSE.dll from ActivePerl 5.20
        case ScriptLang::Ruby:
            //return L"{39D7243A-AF85-46BB-B70C-200EE1021A9B}"; // ActiveRuby 2.4
            return L"{5DBEF578-E278-11D3-8E7A-0000F45A3C05}"; // ActiveRuby 1.8
            // return L"{0AC0D358-E866-11D3-8E82-0000F45A3C05}"; // ActiveRuby 1.8 (global) -> this doesn't support expression evaluation
        default:
            std::cout << "INVALID LANG" << std::endl;
            exit(1);
    }
}

struct LanguageData {
  CComPtr <IActiveScriptParse> parser;
  CComPtr <IActiveScript> script;
  std::unordered_map<std::wstring, VARIANT *> vars;
};

std::unordered_map <ScriptLang, LanguageData> languages;
HRESULT get_data(ScriptLang lang, LanguageData **output);

void show_exc(EXCEPINFO& ei)
{
    std::cout << "Code: " << ei.wCode << std::endl;
    std::wcout << L"Source: " << ei.bstrSource << std::endl;
    std::wcout << L"Description: " << ei.bstrDescription << std::endl;
}

class CSimpleScriptSite:
    public IActiveScriptSite,
    public IActiveScriptSiteWindow {
 public:
  CSimpleScriptSite() : m_cRefCount(1), m_hWnd(NULL) {}

  STDMETHOD_(ULONG, AddRef
  )();
  STDMETHOD_(ULONG, Release
  )();
  STDMETHOD (QueryInterface)(REFIID riid, void **ppvObject);

  STDMETHOD (GetLCID)(LCID *plcid) {
      *plcid = 0;
      return S_OK;
  }
  STDMETHOD (GetItemInfo)(LPCOLESTR pstrName,
                          DWORD dwReturnMask,
                          IUnknown **ppiunkItem,
                          ITypeInfo **ppti) {
      HRESULT hr = E_FAIL;

      if (dwReturnMask & SCRIPTINFO_ITYPEINFO) {
          *ppti = NULL;
          return S_OK;
      }
      std::wcout << "GetItemInfo: " << pstrName << std::endl;

      if (dwReturnMask & SCRIPTINFO_IUNKNOWN) {
          *ppiunkItem = nullptr;
      }

      for (ScriptLang lng = ScriptLang::VBScript; lng < ScriptLang::Last; lng = (ScriptLang)((int)lng + 1)) {
          if (!lstrcmpW(pstrName, LangName(lng))) {
              LanguageData *data;
              TRY(get_data(lng, &data));
              TRY(data->script->GetScriptDispatch(nullptr,
                                                  (IDispatch **) ppiunkItem));
          }
      }

      return hr;
  }
  STDMETHOD (GetDocVersionString)(BSTR *pbstrVersion) {
      *pbstrVersion = SysAllocString(L"1.0");
      return S_OK;
  }
  STDMETHOD (OnScriptTerminate)(const VARIANT *pvarResult,
                                const EXCEPINFO *pexcepinfo) { return S_OK; }
  STDMETHOD (OnStateChange)(SCRIPTSTATE ssScriptState) { return S_OK; }
  STDMETHOD
  (OnScriptError)(IActiveScriptError *pIActiveScriptError) {
      EXCEPINFO exc;
        pIActiveScriptError->GetExceptionInfo(&exc);
      show_exc(exc);
      return S_OK;
  }
  STDMETHOD (OnEnterScript)(void) { return S_OK; }
  STDMETHOD (OnLeaveScript)(void) { return S_OK; }

  STDMETHOD (GetWindow)(HWND *phWnd) {
      *phWnd = m_hWnd;
      return S_OK;
  }
  STDMETHOD (EnableModeless)(BOOL fEnable) { return S_OK; }

  HRESULT SetWindow(HWND hWnd) {
      m_hWnd = hWnd;
      return S_OK;
  }

 public:
  LONG m_cRefCount;
  HWND m_hWnd;
};
CSimpleScriptSite *script_site;
HRESULT get_data(ScriptLang lang, LanguageData **output) {
    int hr = S_OK;
    auto it = languages.find(lang);
    if (it == languages.end()) {
        std::cout << "Language not found!" << std::endl;
        return E_FAIL;
    } else {
        *output = &it->second;
    }
    return hr;
}
STDMETHODIMP_(ULONG)
CSimpleScriptSite::AddRef() {
    return InterlockedIncrement(&m_cRefCount);
}

STDMETHODIMP_(ULONG)
CSimpleScriptSite::Release() {
    if (!InterlockedDecrement(&m_cRefCount)) {
        delete this;
        return 0;
    }

    return m_cRefCount;
}

STDMETHODIMP CSimpleScriptSite::QueryInterface(REFIID riid, void **ppvObject) {
    if (riid == IID_IUnknown || riid == IID_IActiveScriptSiteWindow) {
        *ppvObject = (IActiveScriptSiteWindow * )
        this;
        AddRef();
        return NOERROR;
    }

    if (riid == IID_IActiveScriptSite) {
        *ppvObject = (IActiveScriptSite * )
        this;
        AddRef();
        return NOERROR;
    }

    return E_NOINTERFACE;
}


wchar_t *to_wide_raw(const char* ptr, size_t len, UINT cp) {
    int wide_len =
        MultiByteToWideChar(cp, 0, ptr, len, nullptr, 0);
    wchar_t *wcode = new wchar_t[wide_len + 1];
    MultiByteToWideChar(cp, 0, ptr, -1, wcode, wide_len);
    wcode[wide_len] = 0;
    return wcode;
}

wchar_t *to_wide(rust::Str str) {
    return to_wide_raw(str.data(), str.length(), CP_UTF8);
}

char* from_wide(wchar_t* str) {
    int len = WideCharToMultiByte(CP_ACP, 0, str, -1, nullptr, 0, nullptr, nullptr);
    char *code = new char[len + 1];
    WideCharToMultiByte(CP_ACP, 0, str, -1, code, len, nullptr, nullptr);
    code[len] = 0;
    return code;
}

bool initialized = false;

rust::String error_to_string(int hr) {
    _com_error err(hr);
    const TCHAR* message = err.ErrorMessage();
    std::wcout << "HR error: " << message << std::endl;
    try {
        return rust::String((char16_t*) to_wide_raw((char*)message, strlen(message), CP_ACP));
    } catch (std::exception &e) {
        std::cout << "Error converting to string: " << e.what() << std::endl;
        return "Error converting to string, see stdout";
    }
}

int init() {
    if (initialized)
        return S_OK;

    HRESULT hr;
    TRY(CoInitializeEx(nullptr, COINIT_MULTITHREADED));

    script_site = new CSimpleScriptSite();

    for (ScriptLang lng = ScriptLang::VBScript; lng < ScriptLang::Last; lng = (ScriptLang)((int)lng + 1)) {
        CComPtr <IActiveScript> script_engine;
        CLSID id;
        CLSIDFromString(LangClsId(lng), &id);
        HRESULT inter = script_engine.CoCreateInstance(id);
        if (inter == CO_E_CLASSSTRING || inter == REGDB_E_CLASSNOTREG) {
            std::wcout << "Not loading " << LangName(lng) << std::endl;
            continue;
        }
        CComPtr <IActiveScriptParse> script_parser;
        TRY(inter);
        TRY(script_engine->SetScriptSite(script_site));
        TRY(script_engine->QueryInterface(&script_parser));
        if (script_parser == nullptr)
            return E_NOINTERFACE;
        TRY(script_parser->InitNew());

        std::wcout << "Initialized " << LangName(lng) << std::endl;
        languages.insert({lng, {script_parser,
                                            script_engine}}).first->second;
    }

    for (auto first : languages) {
        for (auto engine : languages) {
            if (first.first != engine.first) {
                std::wcout << "Adding " << LangName(first.first) << " to " << LangName(engine.first) << std::endl;
                TRY(engine.second.script->AddNamedItem(LangName(first.first),
                                                       SCRIPTITEM_GLOBALMEMBERS
                                                           | SCRIPTITEM_ISVISIBLE));
            }
        }
    }

    initialized = true;

    return S_OK;
}

int parse_internal(wchar_t *wcode, VARIANT *output, ScriptLang lang) {
    wchar_t* wcode_trim = wcode;
    while (*wcode_trim == '\n')
        wcode_trim++;
    std::wcout << "Run: " << wcode_trim << std::endl;

    CComVariant result;
    EXCEPINFO ei = {};

    int hr;

    LanguageData *data;
    TRY(get_data(lang, &data));

    hr = data->parser->ParseScriptText(
        &wcode[0],
        nullptr,
        nullptr,
        nullptr,
        0,
        0,
        output ? SCRIPTTEXT_ISEXPRESSION : 0,
        output ? output : &result,
        &ei);

    if (hr == DISP_E_EXCEPTION) {
        std::cout << "Error: " << error_to_string(hr) << std::endl;
        show_exc(ei);
    }

    return hr;
}

int parse(rust::Str code, VARIANT *output, ScriptLang lang) {
    wchar_t *wcode = to_wide(code);

    int hr = parse_internal(wcode, output, lang);

    delete[] wcode;

    return hr;
}

int set_variable(rust::Str name, char *val, ScriptLang var_lang) {
    HRESULT hr;

    wchar_t *wname = to_wide(name);

    var_lang = ScriptLang::VBScript;
    IDispatch *objPtr;
    LanguageData *data;
    TRY(get_data(var_lang, &data));
    data->script->GetScriptDispatch(nullptr, &objPtr);

    DISPID varid;

    wchar_t dim_line[256];
    swprintf(dim_line, 256, L"Dim %s", wname);
    TRY(parse_internal(dim_line, nullptr, var_lang));

    TRY(objPtr->GetIDsOfNames(IID_NULL,
                              &wname,
                              1,
                              LOCALE_USER_DEFAULT,
                              &varid));

    std::wcout << "Set: " << wname << " (ID = " << varid << ")" << std::endl;

    DISPPARAMS dspp;
    ZeroMemory(&dspp, sizeof(dspp));
    dspp.cArgs = dspp.cNamedArgs = 1;
    DISPID dispPropPut = DISPID_PROPERTYPUT;
    dspp.rgdispidNamedArgs = &dispPropPut;
    VARIANT *var = (VARIANT *) val;
    dspp.rgvarg = var;

    TRY(objPtr->Invoke(varid,
                       IID_NULL,
                       LOCALE_USER_DEFAULT,
                       DISPATCH_PROPERTYPUT,
                       &dspp,
                       nullptr,
                       nullptr,
                       nullptr));
    VariantClear(var);

    TRY(objPtr->Release());

    return S_OK;
}

int parse_wrapper(rust::Str code, char *output, ScriptLang lang) {
    return parse(code, (VARIANT *) output, lang);
}

int close() {
    if (!initialized)
        return S_OK;

    HRESULT hr;

    for (auto &pair: languages) {
        pair.second.parser.p = nullptr; // TODO: this is a hack
        pair.second.script.p = nullptr; // TODO: but so is COM anyway
    }
    TRY(script_site->Release());

    ::CoUninitialize();

    initialized = false;
    return S_OK;
}

class VBSContext {
 public:
  ~VBSContext() {
      close();
  }
};

VBSContext sentinel; // cleaner atexit() implementation