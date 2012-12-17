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
    CString sSequenceFilename, sComponent;
    CStringW sDisplayName;
    CFile fSequence;
    PMSIHANDLE hDatabase, hViewST;

    assert(0); // Attach debugger here or press "Ignore"!

    {
        // Prepare our own "TODO" script.
        LPTSTR szBuffer = sSequenceFilename.GetBuffer(MAX_PATH);
        assert(szBuffer);
        ::GetTempPath(MAX_PATH, szBuffer);
        ::GetTempFileName(szBuffer, _T("TS"), 0, szBuffer);
        sSequenceFilename.ReleaseBuffer();
    }
    if (!fSequence.Open(sSequenceFilename, CFile::modeWrite | CFile::shareDenyWrite | CFile::modeCreate | CFile::osSequentialScan)) {
        uiResult = ERROR_CREATE_FAILED;
        goto error1;
    }

    hDatabase = ::MsiGetActiveDatabase(hInstall);
    if (!hDatabase) {
        uiResult = ERROR_INVALID_HANDLE;
        goto error2;
    }

    uiResult = ::MsiDatabaseOpenView(hDatabase, _T("SELECT Component_, DisplayName FROM ScheduledTask"), &hViewST);
    if (uiResult != ERROR_SUCCESS) goto error2;

    uiResult = ::MsiViewExecute(hViewST, NULL);
    if (uiResult != ERROR_SUCCESS) goto error2;

    for (;;) {
        PMSIHANDLE hRecord;
        INSTALLSTATE iInstalled, iAction;

        // Fetch one record from the view.
        uiResult = ::MsiViewFetch(hViewST, &hRecord);
        if (uiResult == ERROR_NO_MORE_ITEMS)
            break;
        else if (uiResult != ERROR_SUCCESS)
            goto error2;

        // Read Component ID.
        uiResult = ::MsiRecordGetString(hRecord, 1, sComponent);
        if (uiResult != ERROR_SUCCESS) goto error2;

        // Get component state and save for deferred processing.
        uiResult = ::MsiGetComponentState(hInstall, sComponent, &iInstalled, &iAction);
        if (uiResult != ERROR_SUCCESS) goto error2;
        fSequence << iAction;

        // Get task's display name and save for deferred processing.
        uiResult = ::MsiRecordGetStringW(hRecord, 2, sDisplayName);
        if (uiResult != ERROR_SUCCESS) goto error2;
        fSequence << sDisplayName;
    }

    verify(::MsiViewClose(hViewST) == ERROR_SUCCESS);
    uiResult = ERROR_SUCCESS;

error2:
    fSequence.Close();
error1:
    if (uiResult != ERROR_SUCCESS) ::DeleteFile(sSequenceFilename);
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
