#define WIN32_LEAN_AND_MEAN
#define EXCELNATIVE_EXPORTS
#include "ExcelNative.h"
#include <shlwapi.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "shlwapi.lib")

static const char s_version[] = "1.0.1";
static BOOL       s_initialized = FALSE;

BOOL WINAPI DllMain(HMODULE hMod, DWORD reason, LPVOID reserved)
{
    UNREFERENCED_PARAMETER(hMod);
    UNREFERENCED_PARAMETER(reserved);
    if (reason == DLL_PROCESS_ATTACH)
        DisableThreadLibraryCalls(hMod);
    return TRUE;
}

EXCELNATIVE_API BOOL WINAPI ExcelNative_Init(HWND hParent)
{
    UNREFERENCED_PARAMETER(hParent);
    if (s_initialized) return TRUE;

    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr) && hr != RPC_E_CHANGED_MODE)
        return FALSE;

    s_initialized = TRUE;
    return TRUE;
}

EXCELNATIVE_API INT WINAPI ExcelNative_Convert(LPCWSTR lpSrc, LPCWSTR lpDst, DWORD flags)
{
    UNREFERENCED_PARAMETER(flags);

    if (!lpSrc || !lpDst)          return ERROR_INVALID_PARAMETER;
    if (!PathFileExistsW(lpSrc))   return ERROR_FILE_NOT_FOUND;

    
    return ERROR_CALL_NOT_IMPLEMENTED;
}

EXCELNATIVE_API VOID WINAPI ExcelNative_Free()
{
    if (s_initialized)
    {
        CoUninitialize();
        s_initialized = FALSE;
    }
}

EXCELNATIVE_API LPCSTR WINAPI ExcelNative_Version()
{
    return s_version;
}
