/**
  tCOMUtil.cpp: implementation of the tCOMUtil class.
*/

#include <stdio.h>
#include <ocidl.h>
#include <shlwapi.h>
#include <wchar.h>

#include "tCOMUtil.h"
#include "tLuaCOMException.h"
#include "tUtil.h"
#include <string>



///
/// Construction/Destruction
///



tCOMUtil::tCOMUtil()
{

}

tCOMUtil::~tCOMUtil()
{

}

ITypeInfo *tCOMUtil::GetCoClassTypeInfo(CLSID clsid)
{
  ITypeLib* typelib = tCOMUtil::LoadTypeLibFromCLSID(clsid);

  if(typelib == NULL)
    return NULL;

  return tCOMUtil::GetCoClassTypeInfo(typelib, clsid);
}

ITypeInfo *tCOMUtil::GetCoClassTypeInfo(IUnknown* punk)
{
  HRESULT hr = S_OK;
  
  tCOMPtr<IProvideClassInfo> pIProvideClassInfo;
  hr = punk->QueryInterface(IID_IProvideClassInfo,
    (void **) &pIProvideClassInfo);
  if (FAILED(hr))
    return NULL;

  ITypeInfo* coclassinfo = NULL;
  hr = pIProvideClassInfo->GetClassInfo(&coclassinfo);
  if(SUCCEEDED(hr))
    return coclassinfo;
  else
    return NULL;
}

ITypeInfo *tCOMUtil::GetCoClassTypeInfo(IDispatch* pdisp, CLSID clsid)
{
  CHECKPARAM(pdisp);

  HRESULT hr = S_OK;

  unsigned int typeinfocount = 0;
  hr = pdisp->GetTypeInfoCount(&typeinfocount);
  if(FAILED(hr) || typeinfocount == 0)
    return NULL;

  tCOMPtr<ITypeInfo> typeinfo;
  hr = pdisp->GetTypeInfo(0, 0, &typeinfo);
  if(FAILED(hr))
    return NULL;

  tCOMPtr<ITypeLib> typelib;
  unsigned int dumb_index = (unsigned int)-1;
  hr = typeinfo->GetContainingTypeLib(&typelib, &dumb_index);
  if(FAILED(hr))
    return NULL;

  ITypeInfo* coclasstypeinfo = tCOMUtil::GetCoClassTypeInfo(typelib, clsid);

  return coclasstypeinfo;
}

ITypeInfo* tCOMUtil::GetCoClassTypeInfo(ITypeLib* typelib,
    const char* coclassname)
{
    if (!typelib || !coclassname || !*coclassname)
        return NULL;

    // --- Конвертация ANSI → UTF‑16 ---
    int n = MultiByteToWideChar(CP_ACP, 0, coclassname, -1, NULL, 0);
    if (n <= 0)
        return NULL;

    std::wstring wcCoClass;
    wcCoClass.resize(n); // включает нулевой терминатор

    if (MultiByteToWideChar(CP_ACP, 0, coclassname, -1,
        &wcCoClass[0], n) <= 0)
        return NULL;

    // --- Поиск имени в TypeLib ---
    const USHORT max_typeinfos = 1000;
    MEMBERID dumb[max_typeinfos];
    ITypeInfo* typeinfos[max_typeinfos];
    USHORT number = max_typeinfos;

    HRESULT hr = typelib->FindName(&wcCoClass[0], 0,
        typeinfos, dumb, &number);

    if (FAILED(hr) || number == 0)
        return NULL;

    // --- Поиск COCLASS среди найденных ---
    ITypeInfo* found = NULL;

    for (USHORT i = 0; i < number; i++)
    {
        TYPEATTR* typeattr = NULL;
        hr = typeinfos[i]->GetTypeAttr(&typeattr);

        if (FAILED(hr) || !typeattr)
        {
            typeinfos[i]->Release();
            continue;
        }

        TYPEKIND typekind = typeattr->typekind;
        typeinfos[i]->ReleaseTypeAttr(typeattr);

        if (!found && typekind == TKIND_COCLASS)
        {
            // оставляем найденный интерфейс, НЕ Release()
            found = typeinfos[i];
        }
        else
        {
            // остальные освобождаем
            typeinfos[i]->Release();
        }
    }

    return found; // может быть NULL, если COCLASS не найден
}



ITypeInfo *tCOMUtil::GetDefaultInterfaceTypeInfo(ITypeInfo* pCoClassinfo,
                                                 bool source)
           
{
  ITypeInfo* typeinfo = NULL;

  // if the component does not have a dispinterface typeinfo 
  // for events, we stay with an interface typeinfo
  ITypeInfo* interface_typeinfo = NULL;

  TYPEATTR* pTA = NULL;
  HRESULT hr = S_OK;

  if (SUCCEEDED(pCoClassinfo->GetTypeAttr(&pTA)))
  {
    UINT i = 0;
    int iFlags = 0;

    for (i=0; i < pTA->cImplTypes; i++)
    {
      //Get the implementation type for this interface
      hr = pCoClassinfo->GetImplTypeFlags(i, &iFlags);

      if (FAILED(hr))
        continue;

      if (iFlags & IMPLTYPEFLAG_FDEFAULT || pTA->cImplTypes == 1)
      {
        if(source == false && !(iFlags & IMPLTYPEFLAG_FSOURCE)
        || source == true && (iFlags & IMPLTYPEFLAG_FSOURCE))
        {
          HREFTYPE    hRefType=0;

          /*
           * This is the interface we want.  Get a handle to
           * the type description from which we can then get
           * the ITypeInfo.
           */
          pCoClassinfo->GetRefTypeOfImplType(i, &hRefType);
          hr = pCoClassinfo->GetRefTypeInfo(hRefType, &typeinfo);

          // gets typeattr info
          TYPEATTR *ptypeattr = NULL;
          GUID guid;
          TYPEKIND typekind;

          hr = typeinfo->GetTypeAttr(&ptypeattr);

          if(FAILED(hr))
          {
            COM_RELEASE(typeinfo);
            break;
          }

          guid = ptypeattr->guid;
          typekind = ptypeattr->typekind;

          typeinfo->ReleaseTypeAttr(ptypeattr);

          COM_RELEASE(interface_typeinfo);
          if(typekind == TKIND_DISPATCH)
          {
            break;  // found!
          }
          else // hold this pointer. If we do not find
              // anything better, we stay with this typeinfo
          {
            interface_typeinfo = typeinfo;
            typeinfo = NULL;
          }
        }
      }
    }

    pCoClassinfo->ReleaseTypeAttr(pTA);
  }

  if(!typeinfo)
    return interface_typeinfo;
  else
    return typeinfo;
}



ITypeInfo *tCOMUtil::GetDispatchTypeInfo(IDispatch* pdisp)
{
  ITypeInfo* typeinfo = NULL;
  HRESULT hr = pdisp->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &typeinfo);
  if(FAILED(hr))
    return NULL;

  TYPEATTR *ptypeattr = NULL;
  typeinfo->GetTypeAttr(&ptypeattr);

  TYPEKIND typekind = ptypeattr->typekind;

  if(typekind == TKIND_DISPATCH)
  {
    typeinfo->ReleaseTypeAttr(ptypeattr);
    return typeinfo;
  }

  // tries to find another description of the same
  // interface in the typelib with TKIND_DISPATCH

  IID iid = ptypeattr->guid;
  
  tCOMPtr<ITypeLib> ptypelib;
  unsigned int dumb = 0;
  hr = typeinfo->GetContainingTypeLib(&ptypelib, &dumb);

  typeinfo->ReleaseTypeAttr(ptypeattr);

  // if there's no containing type lib, we have to
  // trust this one is the right type info
  if(FAILED(hr))
    return typeinfo;

  // obtem a typeinfo do iid fornecido
  // caso haja uma implementacao dispinterface,
  // esta' e' que sera' retornada (segundo
  // documentacao do ActiveX

  ITypeInfo* typeinfo_guid = NULL;
  hr = ptypelib->GetTypeInfoOfGuid(iid, &typeinfo_guid);
  if(FAILED(hr))
    return typeinfo;

  // verifica se e' dispinterface
  TYPEATTR *ptypeattr_iface = NULL;
  hr = typeinfo_guid->GetTypeAttr(&ptypeattr_iface);
  TYPEKIND typekind_iface = ptypeattr_iface->typekind;
  typeinfo_guid->ReleaseTypeAttr(ptypeattr_iface);

  if(typekind_iface == TKIND_DISPATCH)
  {
    // releases original type information
    COM_RELEASE(typeinfo);

    return typeinfo_guid;
  }
  else
  {
    COM_RELEASE(typeinfo_guid);
    
    // returns original type info
    return typeinfo;
  }
}


ITypeInfo* tCOMUtil::GetInterfaceTypeInfo(ITypeLib* typelib,
    const char* interface_name)
{
    if (!typelib || !interface_name || !*interface_name)
        return NULL;

    // --- Конвертация ANSI → UTF‑16 ---
    int n = MultiByteToWideChar(CP_ACP, 0, interface_name, -1, NULL, 0);
    if (n <= 0)
        return NULL;

    std::wstring wcInterface;
    wcInterface.resize(n); // включает нулевой терминатор

    if (MultiByteToWideChar(CP_ACP, 0, interface_name, -1,
        &wcInterface[0], n) <= 0)
        return NULL;

    // --- Поиск имени в TypeLib ---
    const USHORT max_typeinfos = 30;
    MEMBERID dumb[max_typeinfos];
    ITypeInfo* typeinfos[max_typeinfos];
    USHORT number = max_typeinfos;

    HRESULT hr = typelib->FindName(&wcInterface[0], 0,
        typeinfos, dumb, &number);

    if (FAILED(hr) || number == 0)
        return NULL;

    // --- Поиск интерфейса TKIND_DISPATCH ---
    ITypeInfo* found = NULL;

    for (USHORT i = 0; i < number; i++)
    {
        TYPEATTR* typeattr = NULL;
        hr = typeinfos[i]->GetTypeAttr(&typeattr);

        if (FAILED(hr) || !typeattr)
        {
            typeinfos[i]->Release();
            continue;
        }

        TYPEKIND typekind = typeattr->typekind;
        typeinfos[i]->ReleaseTypeAttr(typeattr);

        if (!found && typekind == TKIND_DISPATCH)
        {
            // оставляем найденный интерфейс
            found = typeinfos[i];
        }
        else
        {
            // остальные освобождаем
            typeinfos[i]->Release();
        }
    }

    return found; // может быть NULL, если TKIND_DISPATCH не найден
}


/**
  Carrega uma Typelib associada a um ProgID
*/
ITypeLib* tCOMUtil::LoadTypeLibFromProgID(const char* ProgID,
                                          unsigned short major_version)
{
  CLSID clsid = IID_NULL;
  HRESULT hr = tCOMUtil::ProgID2CLSID(&clsid, ProgID);

  if(FAILED(hr))
    return NULL;

  bool version_found = false;
  int version = -1;

  if(major_version <= 0)
  {
      std::string s = ProgID;
      auto pos = s.rfind('.');
      if (pos != std::string::npos) {
          int v = std::stoi(s.substr(pos + 1));
          if (v > 0) {
              major_version = static_cast<unsigned short>(v);
              version_found = true;
          }
      }
  }

  // tries to get some version information to help finding
  // the right type library
  if(version_found)
    return tCOMUtil::LoadTypeLibFromCLSID(clsid, major_version);
  else
    return tCOMUtil::LoadTypeLibFromCLSID(clsid);

}


/**
  LoadTypeLibByName
  Carrega uma typelib a partir de um arquivo TLB
*/
ITypeLib* tCOMUtil::LoadTypeLibByName(const char* pcFilename)
{
    if (!pcFilename || !*pcFilename)
        return NULL;

    // --- Конвертация ANSI → UTF‑16 ---
    int n = MultiByteToWideChar(CP_ACP, 0, pcFilename, -1, NULL, 0);
    if (n <= 0)
        return NULL;

    std::wstring wcFilename;
    wcFilename.resize(n); // включает нулевой терминатор

    if (MultiByteToWideChar(CP_ACP, 0, pcFilename, -1,
        &wcFilename[0], n) <= 0)
        return NULL;

    // --- Загрузка TypeLib ---
    ITypeLib* ptlib = NULL;
    HRESULT hr = LoadTypeLibEx(&wcFilename[0], REGKIND_NONE, &ptlib);

    if (FAILED(hr))
        return NULL;

    return ptlib;
}




ITypeLib* tCOMUtil::LoadTypeLibFromCLSID(CLSID clsid,
    unsigned short major_version)
{   
    // --- CLSID → wide string ---
    LPOLESTR wcClsid = nullptr; 
    HRESULT hr = StringFromCLSID(clsid, &wcClsid);  
    std::string clsidA;
    if (FAILED(hr) || !wcClsid) 
        return nullptr; 

    bool ok = true; 
    do { 
        // --- wide CLSID → ANSI string (для RegOpenKeyExA) ---
        int lenA = WideCharToMultiByte(CP_ACP, 0, wcClsid, -1, NULL, 0, NULL, NULL);
        if (lenA <= 0) { 
            ok = false; 
            break; 
        } 
        
        clsidA.resize(lenA); // включает нулевой терминатор
        if (WideCharToMultiByte(CP_ACP, 0, wcClsid, -1,
            &clsidA[0], lenA, NULL, NULL) <= 0) {
            ok = false; 
            break; 
        } 
    } while (false); 
    CoTaskMemFree(wcClsid); 
    if (!ok) return nullptr;

    // --- Чтение из реестра: TypeLib и version ---
    BYTE bVersion[100] = {};
    DWORD versionSize = sizeof(bVersion);

    std::string libIdA;
    bool version_info_found = true;

    HKEY iid_key = NULL, obj_key = NULL, typelib_key = NULL, version_key = NULL;

    try
    {
        LONG res = RegOpenKeyExA(HKEY_CLASSES_ROOT, "CLSID", 0, KEY_READ, &iid_key);
        WINCHECK(res == ERROR_SUCCESS);

        res = RegOpenKeyExA(iid_key, clsidA.c_str(), 0, KEY_READ, &obj_key);
        RegCloseKey(iid_key);
        iid_key = NULL;
        WINCHECK(res == ERROR_SUCCESS);

        res = RegOpenKeyExA(obj_key, "TypeLib", 0, KEY_READ, &typelib_key);
        if (res != ERROR_SUCCESS)
        {
            RegCloseKey(obj_key);
            obj_key = NULL;
            LUACOM_EXCEPTION(WINDOWS_ERROR);
        }

        // сначала узнаём размер строки libID
        DWORD libIdSize = 0;
        res = RegQueryValueExA(typelib_key, NULL, NULL, NULL, NULL, &libIdSize);
        WINCHECK(res == ERROR_SUCCESS && libIdSize > 1);

        libIdA.resize(libIdSize); // включает нулевой терминатор
        res = RegQueryValueExA(typelib_key, NULL, NULL, NULL,
            reinterpret_cast<BYTE*>(&libIdA[0]), &libIdSize);
        RegCloseKey(typelib_key);
        typelib_key = NULL;
        WINCHECK(res == ERROR_SUCCESS);

        // убираем возможный лишний нулевой символ в конце
        if (!libIdA.empty() && libIdA.back() == '\0')
            libIdA.pop_back();

        // version
        res = RegOpenKeyExA(obj_key, "version", 0, KEY_READ, &version_key);
        RegCloseKey(obj_key);
        obj_key = NULL;

        if (res != ERROR_SUCCESS)
        {
            version_info_found = false;
        }
        else
        {
            versionSize = sizeof(bVersion);
            res = RegQueryValueExA(version_key, NULL, NULL, NULL, bVersion, &versionSize);
            RegCloseKey(version_key);
            version_key = NULL;

            if (res != ERROR_SUCCESS)
                version_info_found = false;
        }
    }
    catch (class tLuaCOMException& e)
    {
        UNUSED(e);

        if (iid_key) RegCloseKey(iid_key);
        if (obj_key) RegCloseKey(obj_key);
        if (typelib_key) RegCloseKey(typelib_key);
        if (version_key) RegCloseKey(version_key);

        return NULL;
    }

    // --- Разбор версии ---
    int version_major = 0, version_minor = 0;
    int found = 0;

    if (version_info_found)
    {
        std::string ver(reinterpret_cast<const char*>(bVersion), versionSize); 
        auto pos = ver.find('.'); 
        if (pos != std::string::npos) { 
            try { 
                version_major = std::stoi(ver.substr(0, pos)); 
                version_minor = std::stoi(ver.substr(pos + 1)); 
                found = 2; 
            } catch (...) { found = 0; } 
        }
    }

    if (major_version > 0 &&
        ((!version_info_found) || (version_info_found && found == 0)))
    {
        version_major = major_version;
        version_minor = 0;
    }
    else if (!version_info_found)
    {
        bool result = tCOMUtil::GetDefaultTypeLibVersion(
            libIdA.c_str(),
            &version_major,
            &version_minor);

        if (!result)
            return NULL;
    }

    // --- libIdA (ANSI GUID) → wide string ---
    int lenW = MultiByteToWideChar(CP_ACP, 0, libIdA.c_str(), -1, NULL, 0);
    if (lenW <= 0)
        return NULL;

    std::wstring wcTypelib;
    wcTypelib.resize(lenW);

    if (MultiByteToWideChar(CP_ACP, 0, libIdA.c_str(), -1,
        &wcTypelib[0], lenW) <= 0)
        return NULL;

    // --- wide GUID string → GUID ---
    GUID libid = IID_NULL;
    HRESULT hrCLSID = CLSIDFromString(&wcTypelib[0], &libid);
    if (FAILED(hrCLSID))
        return NULL;

    // --- Загрузка TypeLib ---
    ITypeLib* typelib = NULL;
    hr = LoadRegTypeLib(libid, version_major, version_minor, 0, &typelib);

    if (FAILED(hr))
        return NULL;

    return typelib;
}


ITypeInfo* tCOMUtil::GetCoClassTypeInfo(ITypeLib *typelib, CLSID clsid)
{
  ITypeInfo* coclassinfo = NULL;
  
  HRESULT hr = typelib->GetTypeInfoOfGuid(clsid, &coclassinfo);
  
  if(FAILED(hr))
    return NULL;

  return coclassinfo;
}

HRESULT tCOMUtil::ProgID2CLSID(CLSID* pClsid, const char* ProgID)
{
    CHECKPARAM(pClsid);

    if (!ProgID || !*ProgID)
        return E_INVALIDARG;

    HRESULT hr = S_OK;

    // --- Если строка уже GUID в фигурных скобках ---
    if (ProgID[0] == '{')
    {
        int n = MultiByteToWideChar(CP_ACP, 0, ProgID, -1, NULL, 0);
        if (n <= 0)
            return E_FAIL;

        std::wstring wcProgId;
        wcProgId.resize(n);

        if (MultiByteToWideChar(CP_ACP, 0, ProgID, -1, &wcProgId[0], n) <= 0)
            return E_FAIL;

        return CLSIDFromString(&wcProgId[0], pClsid);
    }

    // --- Обычный ProgID → wide string ---
    int n = MultiByteToWideChar(CP_ACP, 0, ProgID, -1, NULL, 0);
    if (n <= 0)
        return E_FAIL;

    std::wstring wcProgId;
    wcProgId.resize(n);

    if (MultiByteToWideChar(CP_ACP, 0, ProgID, -1, &wcProgId[0], n) <= 0)
        return E_FAIL;

    return CLSIDFromProgID(&wcProgId[0], pClsid);
}


CLSID tCOMUtil::GetCLSID(ITypeInfo *coclassinfo)
{
  TYPEATTR* ptypeattr = NULL;

  HRESULT hr = coclassinfo->GetTypeAttr(&ptypeattr);

  if(FAILED(hr))
    return IID_NULL;

  CLSID clsid = ptypeattr->guid;

  coclassinfo->ReleaseTypeAttr(ptypeattr);

  return clsid;
}

bool tCOMUtil::GetDefaultTypeLibVersion(const char* libid,
                                        int *version_major,
                                        int *version_minor)
{
  LONG res = 0;
  HKEY typelib_key, this_typelib_key;

  res = RegOpenKeyExA(HKEY_CLASSES_ROOT,"TypeLib", 0, KEY_READ, &typelib_key);
  RegCloseKey(HKEY_CLASSES_ROOT);

  if(res != ERROR_SUCCESS)
    return false;

  res = RegOpenKeyExA(typelib_key, libid, 0, KEY_READ, &this_typelib_key);
  RegCloseKey(typelib_key);

  if(res != ERROR_SUCCESS)
    return false;

  const int bufsize = 1000;
  char version_info[bufsize];
  DWORD size = bufsize;

  res = RegEnumKeyExA(this_typelib_key, 0, version_info, &size,
    NULL, NULL, NULL, NULL); 
  RegCloseKey(this_typelib_key);

  if(res != ERROR_SUCCESS)
    return false;

  std::string s(version_info);
  auto pos = s.find('.');
  if (pos != std::string::npos) {
      try {
          *version_major = std::stoi(s.substr(0, pos));
          *version_minor = std::stoi(s.substr(pos + 1));
          return true;
      }
      catch (...) {
          return false;
      }
  }
  return false;
}

bool tCOMUtil::GetRegKeyValue(const char* key, char** pValue)
{
    if (!key || !pValue)
        return false;

    *pValue = nullptr;

    DWORD cb = 0;
    LONG ec = RegGetValueA(
        HKEY_CLASSES_ROOT,
        key,
        NULL,               // default value
        RRF_RT_REG_SZ,      // only strings
        NULL,
        NULL,
        &cb
    );

    if (ec != ERROR_SUCCESS || cb == 0)
        return false;

    char* buffer = new char[cb];
    ec = RegGetValueA(
        HKEY_CLASSES_ROOT,
        key,
        NULL,
        RRF_RT_REG_SZ,
        NULL,
        buffer,
        &cb
    );

    if (ec != ERROR_SUCCESS)
    {
        delete[] buffer;
        return false;
    }

    *pValue = buffer;
    return true;
}



bool tCOMUtil::SetRegKeyValue(const char *key,
                              const char *subkey,
                              const char *value)
{
   if (!key) return false;

  std::string Key = key;
  if (subkey && *subkey) {
    Key.push_back('\\');
    Key += subkey;
  }

  HKEY hKey = NULL;
  LONG ec = RegCreateKeyExA(HKEY_CLASSES_ROOT, Key.c_str(), 0, NULL,
                            REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS,
                            NULL, &hKey, NULL);
  if (ec != ERROR_SUCCESS) return false;

  bool ok = false;
  if (value) {
    ec = RegSetValueExA(hKey, NULL, 0, REG_SZ,
                        reinterpret_cast<const BYTE*>(value),
                        static_cast<DWORD>(strlen(value) + 1));
    ok = (ec == ERROR_SUCCESS);
  } else {
    ok = true;
  }

  RegCloseKey(hKey);
  return ok;
}

bool tCOMUtil::DelRegKey(const char* key,
    const char* subkey)
{
    if (!key) return false;

    std::string Key = key;
    if (subkey && *subkey) {
        Key.push_back('\\');
        Key += subkey;
    }

    LONG ec = SHDeleteKeyA(HKEY_CLASSES_ROOT, Key.c_str());
    return (ec == ERROR_SUCCESS);
}

void tCOMUtil::DumpTypeInfo(ITypeInfo *typeinfo)
{
  HRESULT hr = S_OK;
  TYPEATTR* pta = NULL;

  CHECKPARAM(typeinfo);

  CHK_COM_CODE(typeinfo->GetTypeAttr(&pta));

  // prints IID
  LPOLESTR lpsz = NULL;

#ifdef __WINE__
  hr = 0;
  MessageBoxA(NULL, "FIX - not implemented - StringFromIID", "LuaCOM", MB_ICONEXCLAMATION);
  #warning FIX - not implemented - StringFromIID
#else
  hr = StringFromIID(pta->guid, &lpsz);
#endif

  if(FAILED(hr))
  {
    hr = StringFromCLSID(pta->guid, &lpsz);
  }

  if(SUCCEEDED(hr))
  {
    wprintf(L"\nInterface:  %s\n\n", lpsz);

    CoTaskMemFree(lpsz);
  }

  for(int i = 0; i < pta->cFuncs; i++)
  {
    FUNCDESC *pfd = NULL;
    typeinfo->GetFuncDesc(i, &pfd);

    BSTR names[1];
    unsigned int dumb;
    typeinfo->GetNames(pfd->memid, names, 1, &dumb);

    printf("%.3d: %-30s\tid=0x%d\t%d param(s)\n", i,
      tUtil::bstr2string(names[0]).getBuffer(), pfd->memid, pfd->cParams);

    typeinfo->ReleaseFuncDesc(pfd);
    SysFreeString(names[0]);
  }

  typeinfo->ReleaseTypeAttr(pta);
}


const char* tCOMUtil::getPrintableInvokeKind(const INVOKEKIND invkind)
{
  switch(invkind)
  {
  case INVOKE_PROPERTYGET:
    return "propget";

  case INVOKE_PROPERTYPUT:
    return "propput";

  case INVOKE_PROPERTYPUTREF:
    return "propputref";

  case INVOKE_FUNC:
    return "func";
  }

  return NULL;
}

tStringBuffer tCOMUtil::getPrintableTypeDesc(const TYPEDESC& tdesc)
{
    std::string s;

    switch (tdesc.vt & ~(VT_ARRAY | VT_BYREF))
    {
    case VT_VOID:      s += "void"; break;
    case VT_I2:        s += "short"; break;
    case VT_I4:        s += "long"; break;
    case VT_R4:        s += "float"; break;
    case VT_R8:        s += "double"; break;
    case VT_CY:        s += "CY"; break;
    case VT_DATE:      s += "DATE"; break;
    case VT_BSTR:      s += "BSTR"; break;
    case VT_DISPATCH:  s += "IDispatch*"; break;
    case VT_BOOL:      s += "VARIANT_BOOL"; break;
    case VT_VARIANT:   s += "VARIANT"; break;
    case VT_UNKNOWN:   s += "IUnknown*"; break;
    case VT_DECIMAL:   s += "Decimal"; break;
    case VT_UI1:       s += "unsigned char"; break;
    case VT_INT:       s += "int"; break;
    case VT_HRESULT:   s += "void"; break;
    default:           break;
    }

    if (tdesc.vt & VT_BYREF) s.push_back('*');
    if (tdesc.vt & VT_ARRAY) s += "[]";

    return tStringBuffer(s.c_str());
}

const char* tCOMUtil::getPrintableTypeKind(const TYPEKIND tkind)
{
  switch(tkind)
  {
  case TKIND_COCLASS:
    return "coclass";
    break;

  case TKIND_ENUM:
    return "enumeration";
    break;

  case TKIND_RECORD:
    return "record";
    break;

  case TKIND_MODULE:
    return "module";
    break;

  case TKIND_INTERFACE:
    return "interface";
    break;

  case TKIND_DISPATCH:
    return "dispinterface";
    break;

  case TKIND_ALIAS:
    return "alias";
    break;

  case TKIND_UNION:
    return "union";
    break;

  default:
    return "";
    break;
  }
}


HRESULT tCOMUtil::GUID2String(GUID& Guid, char** ppGuid)
{
    if (!ppGuid)
        return E_INVALIDARG;

    *ppGuid = nullptr;

    // --- GUID → wide string ---
    LPOLESTR wcGuid = nullptr;
    HRESULT hr = StringFromCLSID(Guid, &wcGuid);
    if (FAILED(hr) || !wcGuid)
        return hr;

    // --- wide → ANSI ---
    int n = WideCharToMultiByte(CP_ACP, 0, wcGuid, -1, NULL, 0, NULL, NULL);
    if (n <= 0)
    {
        CoTaskMemFree(wcGuid);
        return E_FAIL;
    }

    char* guidA = new char[n]; // n включает нулевой терминатор

    if (WideCharToMultiByte(CP_ACP, 0, wcGuid, -1, guidA, n, NULL, NULL) <= 0)
    {
        CoTaskMemFree(wcGuid);
        delete[] guidA;
        return E_FAIL;
    }

    CoTaskMemFree(wcGuid);

    *ppGuid = guidA;
    return S_OK;
}


CLSID tCOMUtil::FindCLSID(ITypeInfo* interface_typeinfo)
{
  // gets IID
  TYPEATTR* ptypeattr = NULL;
  interface_typeinfo->GetTypeAttr(&ptypeattr);
  IID iid = ptypeattr->guid;
  interface_typeinfo->ReleaseTypeAttr(ptypeattr);
  ptypeattr = NULL;

  // Gets type library
  tCOMPtr<ITypeLib> ptypelib;
  interface_typeinfo->GetContainingTypeLib(&ptypelib, NULL);

  // iterates looking for IID inside some coclass
  long count = ptypelib->GetTypeInfoCount();
  bool found = false;
  CLSID clsid = IID_NULL;
  while(count-- && !found)
  {
    TYPEKIND tkind;
    ptypelib->GetTypeInfoType(count, &tkind);

    if(tkind != TKIND_COCLASS)
      continue;

    // look inside
    tCOMPtr<ITypeInfo> ptypeinfo;
    ptypelib->GetTypeInfo(count, &ptypeinfo);

    // gets counts and clsid
    TYPEATTR* ptypeattr = NULL;
    ptypeinfo->GetTypeAttr(&ptypeattr);
    long ifaces_count   = ptypeattr->cImplTypes;
    clsid = ptypeattr->guid;
    ptypeinfo->ReleaseTypeAttr(ptypeattr);
    ptypeattr = NULL;

    while(ifaces_count-- && !found)
    {
      HREFTYPE RefType;
      ptypeinfo->GetRefTypeOfImplType(ifaces_count, &RefType);
      tCOMPtr<ITypeInfo> piface_typeinfo;
      ptypeinfo->GetRefTypeInfo(RefType, &piface_typeinfo);
      piface_typeinfo->GetTypeAttr(&ptypeattr);

      if(IsEqualIID(ptypeattr->guid, iid))
      {
        found = true;
      }

      piface_typeinfo->ReleaseTypeAttr(ptypeattr);
      ptypeattr = NULL;
    }
  }

  return clsid;
}
