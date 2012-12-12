#include "StdAfx.h"

#pragma comment(lib, "msi.lib")


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

UINT MSITSCA_API EvaluateScheduledTasks(MSIHANDLE hInstall)
{
    UNREFERENCED_PARAMETER(hInstall);

    UINT uiResult;
    BOOL bIsCoInitialized = SUCCEEDED(::CoInitialize(NULL));
    CString msg;

    msg.Format(_T("Pripni razhrošèevalnik na proces %u."), ::GetCurrentProcessId());
    ::MessageBox(NULL, msg, _T("MSITSCA"), MB_OK);
    uiResult = ERROR_SUCCESS;

    if (bIsCoInitialized) ::CoUninitialize();
    return uiResult;
}


UINT MSITSCA_API InstallScheduledTasks(MSIHANDLE hInstall)
{
    UNREFERENCED_PARAMETER(hInstall);
    assert(::MsiGetMode(hInstall, MSIRUNMODE_SCHEDULED));

    UINT uiResult;
    BOOL bIsCoInitialized = SUCCEEDED(::CoInitialize(NULL));
    CString msg;

    msg.Format(_T("Pripni razhrošèevalnik na proces %u."), ::GetCurrentProcessId());
    ::MessageBox(NULL, msg, _T("MSITSCA"), MB_OK);
    uiResult = ERROR_SUCCESS;

    if (bIsCoInitialized) ::CoUninitialize();
    return uiResult;
}


UINT MSITSCA_API CommitScheduledTasks(MSIHANDLE hInstall)
{
    UNREFERENCED_PARAMETER(hInstall);
    assert(::MsiGetMode(hInstall, MSIRUNMODE_COMMIT));

    UINT uiResult;
    BOOL bIsCoInitialized = SUCCEEDED(::CoInitialize(NULL));
    CString msg;

    msg.Format(_T("Pripni razhrošèevalnik na proces %u."), ::GetCurrentProcessId());
    ::MessageBox(NULL, msg, _T("MSITSCA"), MB_OK);
    uiResult = ERROR_SUCCESS;

    if (bIsCoInitialized) ::CoUninitialize();
    return uiResult;
}


UINT MSITSCA_API RollbackScheduledTasks(MSIHANDLE hInstall)
{
    UNREFERENCED_PARAMETER(hInstall);
    assert(::MsiGetMode(hInstall, MSIRUNMODE_ROLLBACK));

    UINT uiResult;
    BOOL bIsCoInitialized = SUCCEEDED(::CoInitialize(NULL));
    CString msg;

    msg.Format(_T("Pripni razhrošèevalnik na proces %u."), ::GetCurrentProcessId());
    ::MessageBox(NULL, msg, _T("MSITSCA"), MB_OK);
    uiResult = ERROR_SUCCESS;

    if (bIsCoInitialized) ::CoUninitialize();
    return uiResult;
}
