#include <windows.h>
#include <psapi.h>

__declspec(dllexport) SAFEARRAY* __stdcall loadedModules() {

  HMODULE process = GetCurrentProcess();

#define maxNofModules 256
  HMODULE modules[maxNofModules];
  DWORD   bytesNeeded;

  if (! EnumProcessModules(
      process,
      modules, // HMODULE *lphModule,
      sizeof(modules),
     &bytesNeeded
  )) {
      MessageBox(0, "EnumProcessModules", 0, 0);
  }

  int nofModules = bytesNeeded/sizeof(HANDLE);

  SAFEARRAYBOUND dim;

  dim.lLbound   = 0;
  dim.cElements = nofModules;

  SAFEARRAY *ret;
  ret = SafeArrayCreate(
       VT_BSTR,
       1      , // Array to be returned has one dimension …
      &dim      // … which is described in this (dim) parameter
  );

  for (LONG i=0; i<nofModules; i++) {

    wchar_t baseName[MAX_PATH];
    GetModuleBaseNameW(process, modules[i], baseName, MAX_PATH);

    SafeArrayPutElement(ret, &i, SysAllocString(baseName));

  }

  return ret;

}
