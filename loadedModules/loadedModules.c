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

  SAFEARRAYBOUND dimensions[2];

  dimensions[0].lLbound = 0; dimensions[0].cElements = 2;
  dimensions[1].lLbound = 0; dimensions[1].cElements = nofModules;

  SAFEARRAY *ret;
  ret = SafeArrayCreate(
       VT_VARIANT,
       2         ,
       dimensions
  );

  VARIANT vBaseName;
  vBaseName.vt = VT_BSTR;

  VARIANT vHandle;
  vHandle.vt = VT_I4;

  LONG putIndices[2];

  for (LONG i=0; i<nofModules; i++) {


    wchar_t baseName[MAX_PATH];
    GetModuleBaseNameW(process, modules[i], baseName, MAX_PATH);

    vBaseName.bstrVal = SysAllocString(baseName);
    vHandle.lVal      =(long) modules[i];

    putIndices[0] = 0;
    putIndices[1] = i;
    SafeArrayPutElement(ret, putIndices, &vBaseName);

    putIndices[0] = 1;
//  putIndices[1] = 1;

    SafeArrayPutElement(ret, putIndices, &vHandle);

  }

  return ret;

}
