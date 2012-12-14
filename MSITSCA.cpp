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
    CString sComponent;
    PMSIHANDLE hDatabase, hViewST;

    assert(0); // Attach debugger here or press "Ignore"!

    hDatabase = ::MsiGetActiveDatabase(hInstall);
    if (!hDatabase) {
        uiResult = ERROR_INVALID_HANDLE;
        goto error1;
    }

    uiResult = ::MsiDatabaseOpenView(hDatabase, _T("SELECT Component_ FROM ScheduledTask"), &hViewST);
    if (uiResult != ERROR_SUCCESS) goto error1;

    uiResult = ::MsiViewExecute(hViewST, NULL);
    if (uiResult != ERROR_SUCCESS) goto error1;

    for (;;) {
        PMSIHANDLE hRecord;
        INSTALLSTATE iInstalled, iAction;

        // Fetch one record from the view.
        uiResult = ::MsiViewFetch(hViewST, &hRecord);
        if (uiResult == ERROR_NO_MORE_ITEMS)
            break;
        else if (uiResult != ERROR_SUCCESS)
            goto error1;

        uiResult = ::MsiRecordGetString(hRecord, 1, sComponent);
        if (uiResult != ERROR_SUCCESS) goto error1;

        uiResult = ::MsiGetComponentState(hInstall, sComponent, &iInstalled, &iAction);
        if (uiResult != ERROR_SUCCESS) goto error1;
    }

    verify(::MsiViewClose(hViewST) == ERROR_SUCCESS);
    uiResult = ERROR_SUCCESS;

error1:
    if (bIsCoInitialized) ::CoUninitialize();
    return uiResult;
}


UINT MSITSCA_API InstallScheduledTasks(MSIHANDLE hInstall)
{
    UNREFERENCED_PARAMETER(hInstall);
    assert(::MsiGetMode(hInstall, MSIRUNMODE_SCHEDULED));

    UINT uiResult;
    BOOL bIsCoInitialized = SUCCEEDED(::CoInitialize(NULL));

    assert(0); // Attach debugger here or press "Ignore"!
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

    assert(0); // Attach debugger here or press "Ignore"!
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

    assert(0); // Attach debugger here or press "Ignore"!
    uiResult = ERROR_SUCCESS;

    if (bIsCoInitialized) ::CoUninitialize();
    return uiResult;
}
