#include "StdAfx.h"

#pragma comment(lib, "msi.lib")
#pragma comment(lib, "mstask.lib")


////////////////////////////////////////////////////////////////////////////
// Local constants
////////////////////////////////////////////////////////////////////////////

#define MSITSCA_TASK_TICK_SIZE  (16*1024)


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
    UINT uiResult;
    HRESULT hr;
    BOOL bIsCoInitialized = SUCCEEDED(::CoInitialize(NULL));
    CMSITSCAOpList olExecute;
    BOOL bRollbackEnabled;
    PMSIHANDLE
        hDatabase,
        hRecordProg = ::MsiCreateRecord(3);
    CString sValue;

    assert(hRecordProg);
#ifdef ASSERT_TO_DEBUG
    assert(0); // Attach debugger here, or press "Ignore"!
#endif

    // Check and add the rollback enabled state.
    uiResult = ::MsiGetProperty(hInstall, _T("RollbackDisabled"), sValue);
    bRollbackEnabled = uiResult == ERROR_SUCCESS ?
        _ttoi(sValue) || !sValue.IsEmpty() && _totlower(sValue.GetAt(0)) == _T('y') ? FALSE : TRUE :
        TRUE;
    olExecute.AddTail(new CMSITSCAOpRollbackEnable(bRollbackEnabled));

    // Open MSI database.
    hDatabase = ::MsiGetActiveDatabase(hInstall);
    if (hDatabase) {
        // Check if ScheduledTask table exists. If it doesn't exist, there's nothing to do.
        MSICONDITION condition = ::MsiDatabaseIsTablePersistent(hDatabase, _T("ScheduledTask"));
        if (condition == MSICONDITION_FALSE || condition == MSICONDITION_TRUE) {
            PMSIHANDLE hViewST;

            // Prepare a query to get a list/view of tasks.
            uiResult = ::MsiDatabaseOpenView(hDatabase, _T("SELECT Task,DisplayName,Application,Parameters,WorkingDir,Flags,Priority,User,Password,Author,Description,IdleMin,IdleDeadline,MaxRuntime,Condition,Component_ FROM ScheduledTask"), &hViewST);
            if (uiResult == ERROR_SUCCESS) {
                // Execute query!
                uiResult = ::MsiViewExecute(hViewST, NULL);
                if (uiResult == ERROR_SUCCESS) {
                    //CString sComponent;
                    CStringW sDisplayName;

                    for (;;) {
                        PMSIHANDLE hRecord;
                        INSTALLSTATE iInstalled, iAction;

                        // Fetch one record from the view.
                        uiResult = ::MsiViewFetch(hViewST, &hRecord);
                        if (uiResult == ERROR_NO_MORE_ITEMS) {
                            uiResult = ERROR_SUCCESS;
                            break;
                        } else if (uiResult != ERROR_SUCCESS)
                            break;

                        // Read and evaluate task's condition.
                        uiResult = ::MsiRecordGetString(hRecord, 15, sValue);
                        if (uiResult != ERROR_SUCCESS) break;
                        condition = ::MsiEvaluateCondition(hInstall, sValue);
                        if (condition == MSICONDITION_FALSE)
                            continue;
                        else if (condition == MSICONDITION_ERROR) {
                            uiResult = ERROR_INVALID_FIELD;
                            break;
                        }

                        // Read task's Component ID.
                        uiResult = ::MsiRecordGetString(hRecord, 16, sValue);
                        if (uiResult != ERROR_SUCCESS) break;

                        // Get the component state.
                        uiResult = ::MsiGetComponentState(hInstall, sValue, &iInstalled, &iAction);
                        if (uiResult != ERROR_SUCCESS) break;

                        // Get task's DisplayName.
                        uiResult = ::MsiRecordFormatStringW(hInstall, hRecord, 2, sDisplayName);

                        if (iAction >= INSTALLSTATE_LOCAL) {
                            // Component is or should be installed. Create the task.
                            PMSIHANDLE hViewTT;
                            CMSITSCAOpTaskCreate *opCreateTask = new CMSITSCAOpTaskCreate(sDisplayName, MSITSCA_TASK_TICK_SIZE);
                            assert(opCreateTask);

                            // Populate the operation with task's data.
                            uiResult = opCreateTask->SetFromRecord(hInstall, hRecord);
                            if (uiResult != ERROR_SUCCESS) break;

                            // Perform another query to get task's triggers.
                            uiResult = ::MsiDatabaseOpenView(hDatabase, _T("SELECT Trigger,BeginDate,EndDate,StartTime,StartTimeRand,MinutesDuration,MinutesInterval,Flags,Type,DaysInterval,WeeksInterval,DaysOfTheWeek,DaysOfMonth,WeekOfMonth,MonthsOfYear FROM TaskTrigger WHERE Task_=?"), &hViewTT);
                            if (uiResult != ERROR_SUCCESS) break;

                            // Execute query!
                            uiResult = ::MsiViewExecute(hViewTT, hRecord);
                            if (uiResult == ERROR_SUCCESS) {
                                // Populate trigger list.
                                uiResult = opCreateTask->SetTriggersFromView(hViewTT);
                                verify(::MsiViewClose(hViewTT) == ERROR_SUCCESS);
                                if (uiResult != ERROR_SUCCESS) break;
                            } else
                                break;

                            olExecute.AddTail(opCreateTask);
                        } else if (iAction >= INSTALLSTATE_REMOVED) {
                            // Component is installed, but should be degraded to advertised/removed. Delete the task.
                            olExecute.AddTail(new CMSITSCAOpTaskDelete(sDisplayName, MSITSCA_TASK_TICK_SIZE));
                        }

                        // The amount of tick space to add for each task to progress indicator.
                        verify(::MsiRecordSetInteger(hRecordProg, 1, 3                     ) == ERROR_SUCCESS);
                        verify(::MsiRecordSetInteger(hRecordProg, 2, MSITSCA_TASK_TICK_SIZE) == ERROR_SUCCESS);
                        if (::MsiProcessMessage(hInstall, INSTALLMESSAGE_PROGRESS, hRecordProg) == IDCANCEL) { uiResult = ERROR_INSTALL_USEREXIT; break; }
                    }
                    verify(::MsiViewClose(hViewST) == ERROR_SUCCESS);

                    if (SUCCEEDED(uiResult)) {
                        CString sSequenceFilename;
                        CAtlFile fSequence;

                        // Prepare our own sequence script file.
                        // The InstallScheduledTasks is a deferred custom action, thus all this information will be unavailable to it.
                        // Therefore save all required info to file now.
                        {
                            LPTSTR szBuffer = sSequenceFilename.GetBuffer(MAX_PATH);
                            assert(szBuffer);
                            ::GetTempPath(MAX_PATH, szBuffer);
                            ::GetTempFileName(szBuffer, _T("TS"), 0, szBuffer);
                            sSequenceFilename.ReleaseBuffer();
                        }
                        // Save execute sequence to file.
                        hr = olExecute.SaveToFile(sSequenceFilename);
                        if (SUCCEEDED(hr)) {
                            // Store sequence script file names to properties for deferred custiom actions.
                            uiResult = ::MsiSetProperty(hInstall, _T("InstallScheduledTasks"),  sSequenceFilename);
                            if (uiResult == ERROR_SUCCESS) {
                                LPCTSTR pszExtension = ::PathFindExtension(sSequenceFilename);
                                CString sSequenceFilename2;

                                sSequenceFilename2.Format(_T("%.*ls-rb%ls"), pszExtension - (LPCTSTR)sSequenceFilename, (LPCTSTR)sSequenceFilename, pszExtension);
                                uiResult = ::MsiSetProperty(hInstall, _T("RollbackScheduledTasks"), sSequenceFilename2);
                                if (uiResult == ERROR_SUCCESS) {
                                    sSequenceFilename2.Format(_T("%.*ls-cm%ls"), pszExtension - (LPCTSTR)sSequenceFilename, (LPCTSTR)sSequenceFilename, pszExtension);
                                    uiResult = ::MsiSetProperty(hInstall, _T("CommitScheduledTasks"),   sSequenceFilename2);
                                    if (uiResult != ERROR_SUCCESS) {
                                        verify(::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_SCHEDULED_TASKS_PROPERTY_SET) == ERROR_SUCCESS);
                                        verify(::MsiRecordSetString (hRecordProg, 2, _T("CommitScheduledTasks")                ) == ERROR_SUCCESS);
                                        verify(::MsiRecordSetInteger(hRecordProg, 3, uiResult                                  ) == ERROR_SUCCESS);
                                        ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                                    }
                                } else {
                                    verify(::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_SCHEDULED_TASKS_PROPERTY_SET) == ERROR_SUCCESS);
                                    verify(::MsiRecordSetString (hRecordProg, 2, _T("RollbackScheduledTasks")              ) == ERROR_SUCCESS);
                                    verify(::MsiRecordSetInteger(hRecordProg, 3, uiResult                                  ) == ERROR_SUCCESS);
                                    ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                                }
                            } else {
                                verify(::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_SCHEDULED_TASKS_PROPERTY_SET) == ERROR_SUCCESS);
                                verify(::MsiRecordSetString (hRecordProg, 2, _T("InstallScheduledTasks")               ) == ERROR_SUCCESS);
                                verify(::MsiRecordSetInteger(hRecordProg, 3, uiResult                                  ) == ERROR_SUCCESS);
                                ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                            }
                            if (uiResult != ERROR_SUCCESS) ::DeleteFile(sSequenceFilename);
                        } else {
                            uiResult = ERROR_INSTALL_SCHEDULED_TASKS_SCRIPT_WRITE;
                            verify(::MsiRecordSetInteger(hRecordProg, 1, uiResult         ) == ERROR_SUCCESS);
                            verify(::MsiRecordSetString (hRecordProg, 2, sSequenceFilename) == ERROR_SUCCESS);
                            verify(::MsiRecordSetInteger(hRecordProg, 3, hr               ) == ERROR_SUCCESS);
                            ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                        }
                    } else if (uiResult != ERROR_INSTALL_USEREXIT) {
                        verify(::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_SCHEDULED_TASKS_OPLIST_CREATE) == ERROR_SUCCESS);
                        verify(::MsiRecordSetInteger(hRecordProg, 2, uiResult                                   ) == ERROR_SUCCESS);
                        ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                    }
                } else {
                    verify(::MsiRecordSetInteger(hRecordProg, 1, uiResult) == ERROR_SUCCESS);
                    ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                }
            } else {
                verify(::MsiRecordSetInteger(hRecordProg, 1, uiResult) == ERROR_SUCCESS);
                ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
            }
        }
    } else {
        uiResult = ERROR_INSTALL_SCHEDULED_TASKS_DATABASE_OPEN;
        verify(::MsiRecordSetInteger(hRecordProg, 1, uiResult) == ERROR_SUCCESS);
        ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
    }

    olExecute.Free();
    if (bIsCoInitialized) ::CoUninitialize();
    return uiResult;
}


UINT MSITSCA_API InstallScheduledTasks(MSIHANDLE hInstall)
{
    UINT uiResult;
    HRESULT hr;
    BOOL bIsCoInitialized = SUCCEEDED(::CoInitialize(NULL));
    CString sSequenceFilename;

#ifdef ASSERT_TO_DEBUG
    assert(0); // Attach debugger here, or press "Ignore"!
#endif

    uiResult = ::MsiGetProperty(hInstall, _T("CustomActionData"), sSequenceFilename);
    if (uiResult == ERROR_SUCCESS) {
        CMSITSCAOpList lstOperations;

        // Load operation sequence.
        hr = lstOperations.LoadFromFile(sSequenceFilename);
        if (SUCCEEDED(hr)) {
            CMSITSCASession session;
            BOOL bIsCleanup = ::MsiGetMode(hInstall, MSIRUNMODE_COMMIT) || ::MsiGetMode(hInstall, MSIRUNMODE_ROLLBACK);

            session.m_hInstall = hInstall;

            // In case of commit/rollback, continue sequence on error, to do as much cleanup as possible.
            session.m_bContinueOnError = bIsCleanup;

            // Execute the operations.
            hr = lstOperations.Execute(&session);
            if (SUCCEEDED(hr)) {
                if (!bIsCleanup && session.m_bRollbackEnabled) {
                    // Save cleanup scripts.
                    LPCTSTR pszExtension = ::PathFindExtension(sSequenceFilename);
                    CString sSequenceFilenameCM, sSequenceFilenameRB;

                    sSequenceFilenameRB.Format(_T("%.*ls-rb%ls"), pszExtension - (LPCTSTR)sSequenceFilename, (LPCTSTR)sSequenceFilename, pszExtension);
                    sSequenceFilenameCM.Format(_T("%.*ls-cm%ls"), pszExtension - (LPCTSTR)sSequenceFilename, (LPCTSTR)sSequenceFilename, pszExtension);

                    // After end of commit, delete rollback file too. After end of rollback, delete commit file too.
                    session.m_olCommit.AddTail(new CMSITSCAOpFileDelete(
#ifdef _UNICODE
                        sSequenceFilenameRB
#else
                        CStringW(sSequenceFilenameRB)
#endif
                        ));
                    session.m_olRollback.AddTail(new CMSITSCAOpFileDelete(
#ifdef _UNICODE
                        sSequenceFilenameCM
#else
                        CStringW(sSequenceFilenameCM)
#endif
                        ));

                    // Save commit file first.
                    hr = session.m_olCommit.SaveToFile(sSequenceFilenameCM);
                    if (SUCCEEDED(hr)) {
                        // Save rollback file next.
                        hr = session.m_olRollback.SaveToFile(sSequenceFilenameRB);
                        if (SUCCEEDED(hr)) {
                            uiResult = ERROR_SUCCESS;
                        } else {
                            // Saving rollback file failed.
                            uiResult = HRESULT_CODE(hr);
                        }
                    } else {
                        // Saving commit file failed.
                        uiResult = HRESULT_CODE(hr);
                    }
                } else {
                    // No cleanup support required.
                    uiResult = ERROR_SUCCESS;
                }
            } else {
                // Execution failed.
                uiResult = HRESULT_CODE(hr);
            }

            if (uiResult != ERROR_SUCCESS) {
                // Perform the cleanup now. The rollback script might not have been written to file yet.
                // And even if it was, the rollback action might not get invoked at all (if scheduled in InstallExecuteSequence later than this action).
                session.m_bContinueOnError = TRUE;
                session.m_bRollbackEnabled = FALSE;
                verify(SUCCEEDED(session.m_olRollback.Execute(&session)));
            }
        } else {
            // Sequence loading failed. Probably, LOCAL SYSTEM doesn't have read access to user's temp directory.
            PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
            assert(hRecordProg);
            uiResult = ERROR_INSTALL_SCHEDULED_TASKS_SCRIPT_READ;
            verify(::MsiRecordSetInteger(hRecordProg, 1, uiResult         ) == ERROR_SUCCESS);
            verify(::MsiRecordSetString (hRecordProg, 2, sSequenceFilename) == ERROR_SUCCESS);
            verify(::MsiRecordSetInteger(hRecordProg, 3, hr               ) == ERROR_SUCCESS);
            ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        }

        lstOperations.Free();
        ::DeleteFile(sSequenceFilename);
    } else {
        // Couldn't get CustomActionData property.
    }

    if (bIsCoInitialized) ::CoUninitialize();
    return uiResult;
}
