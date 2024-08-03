#pragma once
#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#ifdef EXCELNATIVE_EXPORTS
#define EXCELNATIVE_API __declspec(dllexport)
#else
#define EXCELNATIVE_API __declspec(dllimport)
#endif

#define EXCELNATIVE_FLAG_KEEP_FORMULAS  0x01
#define EXCELNATIVE_FLAG_STRIP_MACROS   0x02
#define EXCELNATIVE_FLAG_FORCE_RECALC   0x04

#ifdef __cplusplus
extern "C" {
#endif

EXCELNATIVE_API BOOL  WINAPI ExcelNative_Init(HWND hParent);

EXCELNATIVE_API INT   WINAPI ExcelNative_Convert(LPCWSTR lpSrc, LPCWSTR lpDst, DWORD flags);

EXCELNATIVE_API VOID  WINAPI ExcelNative_Free();

EXCELNATIVE_API LPCSTR WINAPI ExcelNative_Version();

#ifdef __cplusplus
}
#endif
