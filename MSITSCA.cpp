#include "StdAfx.h"


////////////////////////////////////////////////////////////////////////////
// Globalne spremenljivke
////////////////////////////////////////////////////////////////////////////

HINSTANCE MSITSCA::hInstance = NULL;


////////////////////////////////////////////////////////////////////////////
// Globalne funkcije
////////////////////////////////////////////////////////////////////////////

extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    UNREFERENCED_PARAMETER(lpReserved);

    switch (dwReason) {
        case DLL_PROCESS_ATTACH:
            MSITSCA::hInstance = hInstance;
        case DLL_THREAD_ATTACH:
        case DLL_THREAD_DETACH:
        case DLL_PROCESS_DETACH:
            break;
    }

    return TRUE;
}


////////////////////////////////////////////////////////////////////
// Javne funkcije
////////////////////////////////////////////////////////////////////

UINT MSITSCA_API InstallScheduledTask(MSIHANDLE hSession)
{
    UNREFERENCED_PARAMETER(hSession);

    return ERROR_SUCCESS;
}
