#include "StdAfx.h"

#pragma comment(lib, "msi.lib")
#pragma comment(lib, "mstask.lib")


////////////////////////////////////////////////////////////////////////////
// Local constants
////////////////////////////////////////////////////////////////////////////

#define MSITSCA_TASK_TICK_SIZE  2048


////////////////////////////////////////////////////////////////////////////
// Local function declarations
////////////////////////////////////////////////////////////////////////////

inline UINT MsiFormatAndStoreString(MSIHANDLE hInstall, MSIHANDLE hRecord, unsigned int iField, CAtlFile &fSequence, CStringW &sValue);
inline UINT MsiStoreInteger(MSIHANDLE hRecord, unsigned int iField, BOOL bIsNullable, CAtlFile &fSequence, int *piValue = NULL);


////////////////////////////////////////////////////////////////////////////
// Global functions
////////////////////////////////////////////////////////////////////////////

extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    UNREFERENCED_PARAMETER(hInstance);
    UNREFERENCED_PARAMETER(lpReserved);

    switch (dwReason) {
        case DLL_PROCESS_ATTACH:
            // Randomize!
            srand((unsigned)time(NULL));

        case DLL_THREAD_ATTACH:
        case DLL_THREAD_DETACH:
        case DLL_PROCESS_DETACH:
            break;
    }

    return TRUE;
}


////////////////////////////////////////////////////////////////////
// Exported functions
////////////////////////////////////////////////////////////////////

UINT MSITSCA_API EvaluateScheduledTasks(MSIHANDLE hInstall)
{
    UNREFERENCED_PARAMETER(hInstall);

    HRESULT hr;
    BOOL bIsCoInitialized = SUCCEEDED(::CoInitialize(NULL));
    CString sSequenceFilename, sComponent;
    CStringW sValue;
    CAtlFile fSequence;
    PMSIHANDLE hDatabase, hViewST;
    PMSIHANDLE hRecordProg = ::MsiCreateRecord(2);
    BOOL bRollbackEnabled;

    assert(hRecordProg);
    assert(0); // Attach debugger here, or press "Ignore"!

    // The amount of tick space to add for each task to progress indicator.
    verify(::MsiRecordSetInteger(hRecordProg, 1, 3) == ERROR_SUCCESS);
    verify(::MsiRecordSetInteger(hRecordProg, 2, MSITSCA_TASK_TICK_SIZE) == ERROR_SUCCESS);

    // Prepare our own sequence script.
    {
        LPTSTR szBuffer = sSequenceFilename.GetBuffer(MAX_PATH);
        assert(szBuffer);
        ::GetTempPath(MAX_PATH, szBuffer);
        ::GetTempFileName(szBuffer, _T("TS"), 0, szBuffer);
        sSequenceFilename.ReleaseBuffer();
    }
    hr = fSequence.Create(sSequenceFilename, GENERIC_WRITE, FILE_SHARE_READ, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN);
    if (FAILED(hr)) goto error1;

    // Save the rollback enabled state.
    hr = HRESULT_FROM_WIN32(::MsiGetProperty(hInstall, _T("RollbackDisabled"), sValue));
    bRollbackEnabled = SUCCEEDED(hr) ?
        _wtoi(sValue) || !sValue.IsEmpty() && towlower(sValue.GetAt(0)) == L'y' ? FALSE : TRUE :
        TRUE;
    hr = fSequence << (int)bRollbackEnabled;

    hDatabase = ::MsiGetActiveDatabase(hInstall);
    if (!hDatabase) { hr = E_INVALIDARG; goto error2; }

    // Execute a query to get tasks.
    hr = HRESULT_FROM_WIN32(::MsiDatabaseOpenView(hDatabase, _T("SELECT Task,DisplayName,Application,Parameters,WorkingDir,Flags,Priority,User,Password,Description,IdleMin,IdleDeadline,MaxRuntime,Condition,Component_ FROM ScheduledTask"), &hViewST));
    if (FAILED(hr)) goto error2;
    hr = HRESULT_FROM_WIN32(::MsiViewExecute(hViewST, NULL));
    if (FAILED(hr)) goto error2;

    for (;;) {
        PMSIHANDLE hRecord, hViewTT;
        INSTALLSTATE iInstalled, iAction;
        MSICONDITION condition;
        ULONGLONG posTriggerCount, posEOF;
        UINT nTriggers = 0;

        // Fetch one record from the view.
        hr = HRESULT_FROM_WIN32(::MsiViewFetch(hViewST, &hRecord));
        if (hr == HRESULT_FROM_WIN32(ERROR_NO_MORE_ITEMS)) {
            hr = S_OK;
            break;
        } else if (FAILED(hr))
            break;

        // Read and evaluate condition.
        hr = HRESULT_FROM_WIN32(::MsiRecordGetString(hRecord, 14, sValue));
        if (FAILED(hr)) break;
        condition = ::MsiEvaluateCondition(hInstall, sValue);
        if (condition == MSICONDITION_FALSE)
            continue;
        else if (condition == MSICONDITION_ERROR) {
            hr = HRESULT_FROM_WIN32(ERROR_INVALID_FIELD);
            break;
        }

        // Read Component ID.
        hr = HRESULT_FROM_WIN32(::MsiRecordGetString(hRecord, 15, sComponent));
        if (FAILED(hr)) break;

        // Get component state and save for deferred processing.
        hr = HRESULT_FROM_WIN32(::MsiGetComponentState(hInstall, sComponent, &iInstalled, &iAction));
        if (FAILED(hr)) break;
        hr = fSequence << (int)iInstalled;
        if (FAILED(hr)) break;
        hr = fSequence << (int)iAction;
        if (FAILED(hr)) break;

        // Read and store task's data to script file.
        hr = HRESULT_FROM_WIN32(::MsiFormatAndStoreString(hInstall, hRecord,  2,        fSequence, sValue)); // DisplayName
        if (FAILED(hr)) break;
        if (iAction >= INSTALLSTATE_LOCAL) {
            hr = HRESULT_FROM_WIN32(::MsiFormatAndStoreString(hInstall, hRecord,  3,        fSequence, sValue)); // Application
            if (FAILED(hr)) break;
            hr = HRESULT_FROM_WIN32(::MsiFormatAndStoreString(hInstall, hRecord,  4,        fSequence, sValue)); // Parameters
            if (FAILED(hr)) break;
            hr = HRESULT_FROM_WIN32(::MsiFormatAndStoreString(hInstall, hRecord,  5,        fSequence, sValue)); // WorkingDir
            if (FAILED(hr)) break;
            hr = HRESULT_FROM_WIN32(::MsiStoreInteger(                  hRecord,  6, FALSE, fSequence        )); // Flags
            if (FAILED(hr)) break;
            hr = HRESULT_FROM_WIN32(::MsiStoreInteger(                  hRecord,  7, FALSE, fSequence        )); // Priority
            if (FAILED(hr)) break;
            hr = HRESULT_FROM_WIN32(::MsiFormatAndStoreString(hInstall, hRecord,  8,        fSequence, sValue)); // User
            if (FAILED(hr)) break;
            hr = HRESULT_FROM_WIN32(::MsiFormatAndStoreString(hInstall, hRecord,  9,        fSequence, sValue)); // Password
            if (FAILED(hr)) break;
            hr = HRESULT_FROM_WIN32(::MsiFormatAndStoreString(hInstall, hRecord, 10,        fSequence, sValue)); // Description
            if (FAILED(hr)) break;
            hr = HRESULT_FROM_WIN32(::MsiStoreInteger(                  hRecord, 11,  TRUE, fSequence        )); // IdleMin
            if (FAILED(hr)) break;
            hr = HRESULT_FROM_WIN32(::MsiStoreInteger(                  hRecord, 12,  TRUE, fSequence        )); // IdleDeadline
            if (FAILED(hr)) break;
            hr = HRESULT_FROM_WIN32(::MsiStoreInteger(                  hRecord, 13, FALSE, fSequence        )); // MaxRuntime
            if (FAILED(hr)) break;

            // Write 0 for the number of triggers and store file position to update this count later.
            verify(SUCCEEDED(fSequence.GetPosition(posTriggerCount)));
            hr = fSequence << (int)nTriggers;
            if (FAILED(hr)) break;

            // Perform another query to get task's triggers.
            hr = HRESULT_FROM_WIN32(::MsiDatabaseOpenView(hDatabase, _T("SELECT Trigger,BeginDate,EndDate,StartTime,StartTimeRand,MinutesDuration,MinutesInterval,Flags,Type,DaysInterval,WeeksInterval,DaysOfTheWeek,DaysOfMonth,WeekOfMonth,MonthsOfYear FROM TaskTrigger WHERE Task_=?"), &hViewTT));
            if (FAILED(hr)) break;
            hr = HRESULT_FROM_WIN32(::MsiViewExecute(hViewTT, hRecord));
            if (FAILED(hr)) break;

            for (;; nTriggers++) {
                PMSIHANDLE hRecord;
                int iStartTime, iStartTimeRand;

                // Fetch one record from the view.
                hr = HRESULT_FROM_WIN32(::MsiViewFetch(hViewTT, &hRecord));
                if (hr == HRESULT_FROM_WIN32(ERROR_NO_MORE_ITEMS)) {
                    hr = S_OK;
                    break;
                } else if (FAILED(hr))
                    break;

                // Read StartTime and StartTimeRand
                iStartTime = ::MsiRecordGetInteger(hRecord, 4);
                if (iStartTime == MSI_NULL_INTEGER) {
                    hr = HRESULT_FROM_WIN32(ERROR_INVALID_FIELD);
                    break;
                }
                iStartTimeRand = ::MsiRecordGetInteger(hRecord, 5);
                if (iStartTimeRand != MSI_NULL_INTEGER) {
                    // Add random delay to StartTime.
                    iStartTime += ::MulDiv(rand(), iStartTimeRand, RAND_MAX);
                }

                hr = HRESULT_FROM_WIN32(::MsiStoreInteger(hRecord,  2, FALSE, fSequence)); // BeginDate
                if (FAILED(hr)) break;
                hr = HRESULT_FROM_WIN32(::MsiStoreInteger(hRecord,  3,  TRUE, fSequence)); // EndDate
                if (FAILED(hr)) break;
                hr = fSequence << iStartTime;                                              // StartTime
                if (FAILED(hr)) break;
                hr = HRESULT_FROM_WIN32(::MsiStoreInteger(hRecord,  6,  TRUE, fSequence)); // MinutesDuration
                if (FAILED(hr)) break;
                hr = HRESULT_FROM_WIN32(::MsiStoreInteger(hRecord,  7,  TRUE, fSequence)); // MinutesInterval
                if (FAILED(hr)) break;
                hr = HRESULT_FROM_WIN32(::MsiStoreInteger(hRecord,  8, FALSE, fSequence)); // Flags
                if (FAILED(hr)) break;
                hr = HRESULT_FROM_WIN32(::MsiStoreInteger(hRecord,  9, FALSE, fSequence)); // Type
                if (FAILED(hr)) break;
                hr = HRESULT_FROM_WIN32(::MsiStoreInteger(hRecord, 10,  TRUE, fSequence)); // DaysInterval
                if (FAILED(hr)) break;
                hr = HRESULT_FROM_WIN32(::MsiStoreInteger(hRecord, 11,  TRUE, fSequence)); // WeeksInterval
                if (FAILED(hr)) break;
                hr = HRESULT_FROM_WIN32(::MsiStoreInteger(hRecord, 12,  TRUE, fSequence)); // DaysOfTheWeek
                if (FAILED(hr)) break;
                hr = HRESULT_FROM_WIN32(::MsiStoreInteger(hRecord, 13,  TRUE, fSequence)); // DaysOfMonth
                if (FAILED(hr)) break;
                hr = HRESULT_FROM_WIN32(::MsiStoreInteger(hRecord, 14,  TRUE, fSequence)); // WeekOfMonth
                if (FAILED(hr)) break;
                hr = HRESULT_FROM_WIN32(::MsiStoreInteger(hRecord, 15,  TRUE, fSequence)); // MonthsOfYear
                if (FAILED(hr)) break;
            }

            verify(::MsiViewClose(hViewTT) == ERROR_SUCCESS);
            if (FAILED(hr)) break;

            // Seek back to update actual trigger count.
            verify(SUCCEEDED(fSequence.GetPosition(posEOF)));
            verify(SUCCEEDED(fSequence.Seek(posTriggerCount, FILE_BEGIN)));
            fSequence << nTriggers;
            verify(SUCCEEDED(fSequence.Seek(posEOF, FILE_BEGIN)));
        }

        if (::MsiProcessMessage(hInstall, INSTALLMESSAGE_PROGRESS, hRecordProg) == IDCANCEL) { hr = E_ABORT; goto error2; }
    }

    verify(::MsiViewClose(hViewST) == ERROR_SUCCESS);
    if (FAILED(hr)) goto error2;

    // Store sequence script file names for deferred custiom actions.
    hr = HRESULT_FROM_WIN32(::MsiSetProperty(hInstall, _T("caInstallScheduledTasks"),  sSequenceFilename));
    if (FAILED(hr)) goto error2;
    {
        LPCTSTR pszExtension = ::PathFindExtension(sSequenceFilename);
        CString sSequenceFilename2;
        sSequenceFilename2.Format(_T("%.*ls-rb%ls"), pszExtension - (LPCTSTR)sSequenceFilename, (LPCTSTR)sSequenceFilename, pszExtension);
        hr = HRESULT_FROM_WIN32(::MsiSetProperty(hInstall, _T("caRollbackScheduledTasks"), sSequenceFilename2));
        if (FAILED(hr)) goto error2;
        sSequenceFilename2.Format(_T("%.*ls-cm%ls"), pszExtension - (LPCTSTR)sSequenceFilename, (LPCTSTR)sSequenceFilename, pszExtension);
        hr = HRESULT_FROM_WIN32(::MsiSetProperty(hInstall, _T("caCommitScheduledTasks"),   sSequenceFilename2));
        if (FAILED(hr)) goto error2;
    }

    hr = S_OK;

error2:
    fSequence.Close();
error1:
    if (FAILED(hr)) ::DeleteFile(sSequenceFilename);
    if (bIsCoInitialized) ::CoUninitialize();
    return SUCCEEDED(hr) ? ERROR_SUCCESS : hr == E_ABORT ? ERROR_INSTALL_USEREXIT : ERROR_INVALID_DATA;
}


UINT MSITSCA_API InstallScheduledTasks(MSIHANDLE hInstall)
{
    assert(::MsiGetMode(hInstall, MSIRUNMODE_SCHEDULED));

    HRESULT hr;
    BOOL bIsCoInitialized = SUCCEEDED(::CoInitialize(NULL));
    CString sSequenceFilename, sSequenceFilenameRB, sSequenceFilenameCM;
    LPCTSTR pszExtension;
    CStringW sDisplayName;
    CAtlFile fSequence, fSequence2;
    CComPtr<ITaskScheduler> pTaskScheduler;
    CMSITSCAOperationList lstOperationsRB, lstOperationsCM;
    PMSIHANDLE hRecordAction = ::MsiCreateRecord(3), hRecordProg = ::MsiCreateRecord(3);
    BOOL bRollbackEnabled;

    assert(hRecordProg);

    assert(0); // Attach debugger here, or press "Ignore"!

    // Tell the installer to use explicit progress messages.
    verify(::MsiRecordSetInteger(hRecordProg, 1, 1) == ERROR_SUCCESS);
    verify(::MsiRecordSetInteger(hRecordProg, 2, 1) == ERROR_SUCCESS);
    verify(::MsiRecordSetInteger(hRecordProg, 3, 0) == ERROR_SUCCESS);
    if (::MsiProcessMessage(hInstall, INSTALLMESSAGE_PROGRESS, hRecordProg) == IDCANCEL) { hr = E_ABORT; goto error1; }

    // Specify that an update of the progress bar's position in this case means to move it forward by one increment.
    verify(::MsiRecordSetInteger(hRecordProg, 1, 2) == ERROR_SUCCESS);
    verify(::MsiRecordSetInteger(hRecordProg, 2, MSITSCA_TASK_TICK_SIZE) == ERROR_SUCCESS);
    verify(::MsiRecordSetInteger(hRecordProg, 3, 0) == ERROR_SUCCESS);

    // Determine sequence script file names.
    hr = HRESULT_FROM_WIN32(::MsiGetProperty(hInstall, _T("CustomActionData"), sSequenceFilename));
    if (FAILED(hr)) goto error1;
    pszExtension = ::PathFindExtension(sSequenceFilename);
    assert(pszExtension);
    sSequenceFilenameRB.Format(_T("%.*ls-rb%ls"), pszExtension - (LPCTSTR)sSequenceFilename, (LPCTSTR)sSequenceFilename, pszExtension);
    sSequenceFilenameCM.Format(_T("%.*ls-cm%ls"), pszExtension - (LPCTSTR)sSequenceFilename, (LPCTSTR)sSequenceFilename, pszExtension);

    // Open sequence script file.
    hr = fSequence.Create(sSequenceFilename, GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN);
    if (FAILED(hr)) goto error2;

    // Get rollback-enabled state.
    {
        int iValue;
        hr = fSequence >> iValue;
        if (FAILED(hr)) goto error3;
        bRollbackEnabled = iValue ? TRUE : FALSE;
    }

    hr = pTaskScheduler.CoCreateInstance(CLSID_CTaskScheduler, NULL, CLSCTX_ALL);
    if (FAILED(hr)) goto error3;

    for (;;) {
        INSTALLSTATE iInstalled, iAction;
        CComPtr<ITask> pTaskOrig;

        hr = fSequence >> (int&)iInstalled;
        if (hr == HRESULT_FROM_WIN32(ERROR_HANDLE_EOF)) {
            hr = S_OK;
            break;
        } else if (FAILED(hr)) goto error4;
        hr = fSequence >> (int&)iAction;
        if (FAILED(hr)) goto error4;
        hr = fSequence >> sDisplayName;
        if (FAILED(hr)) goto error4;

        verify(::MsiRecordSetString(hRecordAction, 1, sDisplayName) == ERROR_SUCCESS);
        if (::MsiProcessMessage(hInstall, INSTALLMESSAGE_ACTIONDATA, hRecordAction) == IDCANCEL) { hr = E_ABORT; goto error1; }

        // See if the task with this name already exists.
        hr = pTaskScheduler->Activate(sDisplayName, IID_ITask, (IUnknown**)&pTaskOrig);
        if (hr != COR_E_FILENOTFOUND && FAILED(hr)) goto error4;
        else if (SUCCEEDED(hr)) {
            // The task exists.
            if (iInstalled < INSTALLSTATE_LOCAL || iAction < INSTALLSTATE_LOCAL) {
                // This component either was, or is going to be uninstalled.
                // In first case this task shouldn't exist. Perhaps it is from some failed previous install/uninstall attempt.
                // In second case we should delete it.
                // So delete it anyway!
                DWORD dwFlags;
                CStringW sDisplayNameOrig;
                UINT uiCount = 0;

                if (bRollbackEnabled) {
                    hr = pTaskOrig->GetFlags(&dwFlags);
                    if (FAILED(hr)) goto error4;
                    if ((dwFlags & TASK_FLAG_DISABLED) == 0) {
                        // The task is enabled.

                        // In case the task disabling fails halfway, try to re-enable it anyway on rollback.
                        lstOperationsRB.AddHead(new CMSITSCAOpEnableTask(sDisplayName, TRUE));

                        // Disable it.
                        dwFlags |= TASK_FLAG_DISABLED;
                        hr = pTaskOrig->SetFlags(dwFlags);
                        if (FAILED(hr)) goto error4;
                    }

                    // Prepare a backup copy of task.
                    do {
                        sDisplayNameOrig.Format(L"%ls (orig %u)", (LPCWSTR)sDisplayName, ++uiCount);
                        hr = pTaskScheduler->AddWorkItem(sDisplayNameOrig, pTaskOrig);
                    } while (hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS));
                    // In case the backup copy creation failed halfway, try to delete its remains anyway on rollback.
                    lstOperationsRB.AddHead(new CMSITSCAOpDeleteTask(sDisplayNameOrig));
                    if (FAILED(hr)) goto error4;

                    // Save the backup copy.
                    CComQIPtr<IPersistFile> pTaskFile(pTaskOrig);
                    hr = pTaskFile->Save(NULL, TRUE);
                    if (FAILED(hr)) goto error4;

                    // Order rollback action to restore from backup copy.
                    lstOperationsRB.AddHead(new CMSITSCAOpCopyTask(sDisplayNameOrig, sDisplayName));
                }

                // Delete it.
                hr = pTaskScheduler->Delete(sDisplayName);
                if (FAILED(hr)) goto error4;

                if (bRollbackEnabled) {
                    // Order commit action to delete backup copy.
                    lstOperationsCM.AddTail(new CMSITSCAOpDeleteTask(sDisplayNameOrig));
                }
            }
        }

        if (iAction >= INSTALLSTATE_LOCAL) {
            // The component is being installed.
            CComPtr<ITask> pTask;
            CComPtr<IPersistFile> pTaskFile;

            if (bRollbackEnabled) {
                // In case the task creation fails halfway, try to delete its remains anyway on rollback.
                lstOperationsRB.AddHead(new CMSITSCAOpDeleteTask(sDisplayName));
            }

            // Create the new task.
            hr = pTaskScheduler->NewWorkItem(sDisplayName, CLSID_CTask, IID_ITask, (IUnknown**)&pTask);
            if (FAILED(hr)) goto error4;

            // Load the task from sequence file.
            hr = fSequence >> pTask;
            if (FAILED(hr)) goto error4;

            // Save the task.
            hr = pTask.QueryInterface<IPersistFile>(&pTaskFile);
            if (FAILED(hr)) goto error4;
            hr = pTaskFile->Save(NULL, TRUE);
            if (FAILED(hr)) goto error4;
        }

        if (::MsiProcessMessage(hInstall, INSTALLMESSAGE_PROGRESS, hRecordProg) == IDCANCEL) { hr = E_ABORT; goto error4; }
    }

    if (bRollbackEnabled) {
        // Save commit script.
        hr = fSequence2.Create(sSequenceFilenameCM, GENERIC_WRITE, FILE_SHARE_READ, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN);
        if (SUCCEEDED(hr)) {
            // Save rollback filename too, since commit script should delete it.
            hr = fSequence2 << sSequenceFilenameRB;
            if (FAILED(hr)) goto error4;
            // Save commit script.
            hr = lstOperationsCM.Save(fSequence2);
            if (FAILED(hr)) goto error4;
            fSequence2.Close();
        }
    }

    hr = S_OK;

error4:
    if (bRollbackEnabled) {
        if (SUCCEEDED(hr)) {
            // This custom action completed successfully. Save rollback script in case of error.
            if (SUCCEEDED(fSequence2.Create(sSequenceFilenameRB, GENERIC_WRITE, FILE_SHARE_READ, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN))) {
                // Save commit filename too, since rollback script should delete it.
                if (SUCCEEDED(fSequence2 << sSequenceFilenameRB))
                    lstOperationsRB.Save(fSequence2);
                fSequence2.Close();
            }
        } else {
            // This very custom action encountered an error. Its rollback won't be called later (if scheduled later than this CA) so do the cleanup immediately.
            lstOperationsRB.Execute(pTaskScheduler, TRUE);
            ::DeleteFile(sSequenceFilenameCM);
        }
    }
    lstOperationsRB.Free();
    lstOperationsCM.Free();
error3:
    fSequence.Close();
error2:
    ::DeleteFile(sSequenceFilename);
error1:
    if (bIsCoInitialized) ::CoUninitialize();
    return SUCCEEDED(hr) ? ERROR_SUCCESS : hr == E_ABORT ? ERROR_INSTALL_USEREXIT : ERROR_INVALID_DATA;
}


UINT MSITSCA_API FinalizeScheduledTasks(MSIHANDLE hInstall)
{
    HRESULT hr;
    BOOL bIsCoInitialized = SUCCEEDED(::CoInitialize(NULL));
    CString sSequenceFilename, sSequenceFilename2;
    CAtlFile fSequence;
    CComPtr<ITaskScheduler> pTaskScheduler;
    CMSITSCAOperationList lstOperations;

    assert(0); // Attach debugger here, or press "Ignore"!

    hr = HRESULT_FROM_WIN32(::MsiGetProperty(hInstall, _T("CustomActionData"), sSequenceFilename));
    if (SUCCEEDED(hr)) {
        // Open sequence script file.
        hr = fSequence.Create(sSequenceFilename, GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN);
        if (SUCCEEDED(hr)) {
            // Read "the other" script filename and delete it.
            hr = fSequence >> sSequenceFilename2;
            if (SUCCEEDED(hr)) {
                ::DeleteFile(sSequenceFilename2);

                // Get task scheduler object.
                hr = pTaskScheduler.CoCreateInstance(CLSID_CTaskScheduler, NULL, CLSCTX_ALL);
                if (SUCCEEDED(hr)) {
                    // Load operation sequence.
                    hr = lstOperations.Load(fSequence);
                    if (SUCCEEDED(hr)) {
                        // Execute the cleanup operations.
                        hr = lstOperations.Execute(pTaskScheduler, TRUE);
                    }

                    lstOperations.Free();
                }
            }

            fSequence.Close();
        } else
            hr = S_OK;

        ::DeleteFile(sSequenceFilename);
    }

    if (bIsCoInitialized) ::CoUninitialize();
    return SUCCEEDED(hr) ? ERROR_SUCCESS : ERROR_INVALID_DATA;
}


////////////////////////////////////////////////////////////////////////////
// Local functions
////////////////////////////////////////////////////////////////////////////

inline UINT MsiFormatAndStoreString(MSIHANDLE hInstall, MSIHANDLE hRecord, unsigned int iField, CAtlFile &fSequence, CStringW &sValue)
{
    UINT uiResult;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord, iField, sValue);
    if (uiResult != ERROR_SUCCESS) return uiResult;
    if (FAILED(fSequence << sValue)) return ERROR_WRITE_FAULT;
    return ERROR_SUCCESS;
}


inline UINT MsiStoreInteger(MSIHANDLE hRecord, unsigned int iField, BOOL bIsNullable, CAtlFile &fSequence, int *piValue)
{
    int iValue;

    iValue = ::MsiRecordGetInteger(hRecord, iField);
    if (!bIsNullable && iValue == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;
    if (FAILED(fSequence << iValue)) return ERROR_WRITE_FAULT;
    if (piValue) *piValue = iValue;
    return ERROR_SUCCESS;
}
