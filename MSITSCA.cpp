#include "StdAfx.h"


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
            //assert(0); // Attach debugger here, or press "Ignore"!

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
    AMSICA::COpList olExecute;
    BOOL bRollbackEnabled;
    PMSIHANDLE
        hDatabase,
        hRecordProg = ::MsiCreateRecord(3);
    ATL::CAtlString sValue;

    // Check and add the rollback enabled state.
    uiResult = ::MsiGetProperty(hInstall, _T("RollbackDisabled"), sValue);
    bRollbackEnabled = uiResult == ERROR_SUCCESS ?
        _ttoi(sValue) || !sValue.IsEmpty() && _totlower(sValue.GetAt(0)) == _T('y') ? FALSE : TRUE :
        TRUE;
    olExecute.AddTail(new AMSICA::COpRollbackEnable(bRollbackEnabled));

    // Open MSI database.
    hDatabase = ::MsiGetActiveDatabase(hInstall);
    if (hDatabase) {
        // Check if ScheduledTask table exists. If it doesn't exist, there's nothing to do.
        MSICONDITION condition = ::MsiDatabaseIsTablePersistent(hDatabase, _T("ScheduledTask"));
        if (condition == MSICONDITION_FALSE || condition == MSICONDITION_TRUE) {
            PMSIHANDLE hViewST;

            // Prepare a query to get a list/view of tasks.
            uiResult = ::MsiDatabaseOpenView(hDatabase, _T("SELECT Task,DisplayName,Application,Parameters,Directory_,Flags,Priority,User,Password,Author,Description,IdleMin,IdleDeadline,MaxRuntime,Condition,Component_ FROM ScheduledTask"), &hViewST);
            if (uiResult == ERROR_SUCCESS) {
                // Execute query!
                uiResult = ::MsiViewExecute(hViewST, NULL);
                if (uiResult == ERROR_SUCCESS) {
                    //ATL::CAtlString sComponent;
                    ATL::CAtlStringW sDisplayName;

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
                            AMSICA::COpTaskCreate *opCreateTask = new AMSICA::COpTaskCreate(sDisplayName, MSITSCA_TASK_TICK_SIZE);

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
                                ::MsiViewClose(hViewTT);
                                if (uiResult != ERROR_SUCCESS) break;
                            } else
                                break;

                            olExecute.AddTail(opCreateTask);
                        } else if (iAction >= INSTALLSTATE_REMOVED) {
                            // Component is installed, but should be degraded to advertised/removed. Delete the task.
                            olExecute.AddTail(new AMSICA::COpTaskDelete(sDisplayName, MSITSCA_TASK_TICK_SIZE));
                        }

                        // The amount of tick space to add for each task to progress indicator.
                        ::MsiRecordSetInteger(hRecordProg, 1, 3                     );
                        ::MsiRecordSetInteger(hRecordProg, 2, MSITSCA_TASK_TICK_SIZE);
                        if (::MsiProcessMessage(hInstall, INSTALLMESSAGE_PROGRESS, hRecordProg) == IDCANCEL) { uiResult = ERROR_INSTALL_USEREXIT; break; }
                    }
                    ::MsiViewClose(hViewST);

                    if (SUCCEEDED(uiResult)) {
                        ATL::CAtlString sSequenceFilename;
                        ATL::CAtlFile fSequence;

                        // Prepare our own sequence script file.
                        // The InstallScheduledTasks is a deferred custom action, thus all this information will be unavailable to it.
                        // Therefore save all required info to file now.
                        {
                            LPTSTR szBuffer = sSequenceFilename.GetBuffer(MAX_PATH);
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
                                ATL::CAtlString sSequenceFilename2;

                                sSequenceFilename2.Format(_T("%.*ls-rb%ls"), pszExtension - (LPCTSTR)sSequenceFilename, (LPCTSTR)sSequenceFilename, pszExtension);
                                uiResult = ::MsiSetProperty(hInstall, _T("RollbackScheduledTasks"), sSequenceFilename2);
                                if (uiResult == ERROR_SUCCESS) {
                                    sSequenceFilename2.Format(_T("%.*ls-cm%ls"), pszExtension - (LPCTSTR)sSequenceFilename, (LPCTSTR)sSequenceFilename, pszExtension);
                                    uiResult = ::MsiSetProperty(hInstall, _T("CommitScheduledTasks"),   sSequenceFilename2);
                                    if (uiResult != ERROR_SUCCESS) {
                                        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_PROPERTY_SET);
                                        ::MsiRecordSetString (hRecordProg, 2, _T("CommitScheduledTasks"));
                                        ::MsiRecordSetInteger(hRecordProg, 3, uiResult                  );
                                        ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                                    }
                                } else {
                                    ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_PROPERTY_SET  );
                                    ::MsiRecordSetString (hRecordProg, 2, _T("RollbackScheduledTasks"));
                                    ::MsiRecordSetInteger(hRecordProg, 3, uiResult                    );
                                    ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                                }
                            } else {
                                ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_PROPERTY_SET );
                                ::MsiRecordSetString (hRecordProg, 2, _T("InstallScheduledTasks"));
                                ::MsiRecordSetInteger(hRecordProg, 3, uiResult                   );
                                ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                            }
                            if (uiResult != ERROR_SUCCESS) ::DeleteFile(sSequenceFilename);
                        } else {
                            uiResult = ERROR_INSTALL_SCRIPT_WRITE;
                            ::MsiRecordSetInteger(hRecordProg, 1, uiResult         );
                            ::MsiRecordSetString (hRecordProg, 2, sSequenceFilename);
                            ::MsiRecordSetInteger(hRecordProg, 3, hr               );
                            ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                        }
                    } else if (uiResult != ERROR_INSTALL_USEREXIT) {
                        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_SCHEDULED_TASKS_OPLIST_CREATE);
                        ::MsiRecordSetInteger(hRecordProg, 2, uiResult                                   );
                        ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                    }
                } else {
                    ::MsiRecordSetInteger(hRecordProg, 1, uiResult);
                    ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                }
            } else {
                ::MsiRecordSetInteger(hRecordProg, 1, uiResult);
                ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
            }
        }
    } else {
        uiResult = ERROR_INSTALL_DATABASE_OPEN;
        ::MsiRecordSetInteger(hRecordProg, 1, uiResult);
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
    ATL::CAtlString sSequenceFilename;

    uiResult = ::MsiGetProperty(hInstall, _T("CustomActionData"), sSequenceFilename);
    if (uiResult == ERROR_SUCCESS) {
        AMSICA::COpList lstOperations;
        BOOL bIsCleanup = ::MsiGetMode(hInstall, MSIRUNMODE_COMMIT) || ::MsiGetMode(hInstall, MSIRUNMODE_ROLLBACK);

        // Load operation sequence.
        hr = lstOperations.LoadFromFile(sSequenceFilename);
        if (SUCCEEDED(hr)) {
            AMSICA::CSession session;

            session.m_hInstall = hInstall;

            // In case of commit/rollback, continue sequence on error, to do as much cleanup as possible.
            session.m_bContinueOnError = bIsCleanup;

            // Execute the operations.
            hr = lstOperations.Execute(&session);
            if (!bIsCleanup) {
                // Save cleanup scripts of delayed action regardless of action's execution status.
                // Rollback action MUST be scheduled in InstallExecuteSequence before this action! Otherwise cleanup won't be performed in case this action execution failed.
                LPCTSTR pszExtension = ::PathFindExtension(sSequenceFilename);
                ATL::CAtlString sSequenceFilenameCM, sSequenceFilenameRB;
                HRESULT hr;

                sSequenceFilenameRB.Format(_T("%.*ls-rb%ls"), pszExtension - (LPCTSTR)sSequenceFilename, (LPCTSTR)sSequenceFilename, pszExtension);
                sSequenceFilenameCM.Format(_T("%.*ls-cm%ls"), pszExtension - (LPCTSTR)sSequenceFilename, (LPCTSTR)sSequenceFilename, pszExtension);

                // After commit, delete rollback file. After rollback, delete commit file.
                session.m_olCommit.AddTail(new AMSICA::COpFileDelete(
#ifdef _UNICODE
                    sSequenceFilenameRB
#else
                    ATL::CAtlStringW(sSequenceFilenameRB)
#endif
                    ));
                session.m_olRollback.AddTail(new AMSICA::COpFileDelete(
#ifdef _UNICODE
                    sSequenceFilenameCM
#else
                    ATL::CAtlStringW(sSequenceFilenameCM)
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
                        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
                        uiResult = ERROR_INSTALL_SCRIPT_WRITE;
                        ::MsiRecordSetInteger(hRecordProg, 1, uiResult           );
                        ::MsiRecordSetString (hRecordProg, 2, sSequenceFilenameRB);
                        ::MsiRecordSetInteger(hRecordProg, 3, hr                 );
                        ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                    }
                } else {
                    // Saving commit file failed.
                    PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
                    uiResult = ERROR_INSTALL_SCRIPT_WRITE;
                    ::MsiRecordSetInteger(hRecordProg, 1, uiResult           );
                    ::MsiRecordSetString (hRecordProg, 2, sSequenceFilenameCM);
                    ::MsiRecordSetInteger(hRecordProg, 3, hr                 );
                    ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                }

                if (uiResult != ERROR_SUCCESS) {
                    // The commit and/or rollback scripts were not written to file successfully. Perform the cleanup immediately.
                    session.m_bContinueOnError = TRUE;
                    session.m_bRollbackEnabled = FALSE;
                    session.m_olRollback.Execute(&session);
                    ::DeleteFile(sSequenceFilenameRB);
                }
            } else {
                // No cleanup after cleanup support.
                uiResult = ERROR_SUCCESS;
            }

            if (FAILED(hr)) {
                // Execution of the action failed.
                uiResult = HRESULT_CODE(hr);
            }

            ::DeleteFile(sSequenceFilename);
        } else if (hr == ERROR_FILE_NOT_FOUND && bIsCleanup) {
            // Sequence file not found and this is rollback/commit action. Either of the following scenarios are possible:
            // - The delayed action failed to save the rollback/commit file. The delayed action performed cleanup itself. No further action is required.
            // - Somebody removed the rollback/commit file between delayed action and rollback/commit action. No further action is possible.
            uiResult = ERROR_SUCCESS;
        } else {
            // Sequence loading failed. Probably, LOCAL SYSTEM doesn't have read access to user's temp directory.
            PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
            uiResult = ERROR_INSTALL_SCRIPT_READ;
            ::MsiRecordSetInteger(hRecordProg, 1, uiResult         );
            ::MsiRecordSetString (hRecordProg, 2, sSequenceFilename);
            ::MsiRecordSetInteger(hRecordProg, 3, hr               );
            ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        }

        lstOperations.Free();
    } else {
        // Couldn't get CustomActionData property. uiResult has the error code.
    }

    if (bIsCoInitialized) ::CoUninitialize();
    return uiResult;
}
