#include "StdAfx.h"

#pragma comment(lib, "taskschd.lib")


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOp
////////////////////////////////////////////////////////////////////////////

CMSITSCAOp::CMSITSCAOp(int iTicks) :
    m_iTicks(iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpSingleStringOperation
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpSingleStringOperation::CMSITSCAOpSingleStringOperation(LPCWSTR pszValue, int iTicks) :
    m_sValue(pszValue),
    CMSITSCAOp(iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpDoubleStringOperation
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpSrcDstStringOperation::CMSITSCAOpSrcDstStringOperation(LPCWSTR pszValue1, LPCWSTR pszValue2, int iTicks) :
    m_sValue1(pszValue1),
    m_sValue2(pszValue2),
    CMSITSCAOp(iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpBooleanOperation
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpBooleanOperation::CMSITSCAOpBooleanOperation(BOOL bValue, int iTicks) :
    m_bValue(bValue),
    CMSITSCAOp(iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpEnableRollback
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpEnableRollback::CMSITSCAOpEnableRollback(BOOL bEnable, int iTicks) :
    CMSITSCAOpBooleanOperation(bEnable, iTicks)
{
}


HRESULT CMSITSCAOpEnableRollback::Execute(CMSITSCASession *pSession)
{
    assert(pSession);

    pSession->m_bRollbackEnabled = m_bValue;
    return S_OK;
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpDeleteFile
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpDeleteFile::CMSITSCAOpDeleteFile(LPCWSTR pszFileName, int iTicks) :
    CMSITSCAOpSingleStringOperation(pszFileName, iTicks)
{
}


HRESULT CMSITSCAOpDeleteFile::Execute(CMSITSCASession *pSession)
{
    assert(pSession);
    DWORD dwError;

    if (pSession->m_bRollbackEnabled) {
        CStringW sBackupName;
        UINT uiCount = 0;

        do {
            // Rename the file to make a backup.
            sBackupName.Format(L"%ls (orig %u)", (LPCWSTR)m_sValue, ++uiCount);
            dwError = ::MoveFileW(m_sValue, sBackupName) ? ERROR_SUCCESS : ::GetLastError();
        } while (dwError == ERROR_ALREADY_EXISTS);
        if (dwError == ERROR_SUCCESS) {
            // Order rollback action to restore from backup copy.
            pSession->m_olRollback.AddHead(new CMSITSCAOpMoveFile(sBackupName, m_sValue));

            // Order commit action to delete backup copy.
            pSession->m_olCommit.AddTail(new CMSITSCAOpDeleteFile(sBackupName));
        }
    } else {
        // Delete the file.
        dwError = ::DeleteFileW(m_sValue) ? ERROR_SUCCESS : ::GetLastError();
    }

    if (dwError == ERROR_SUCCESS || dwError == ERROR_FILE_NOT_FOUND)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        verify(::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_DELETE_FAILED) == ERROR_SUCCESS);
        verify(::MsiRecordSetStringW(hRecordProg, 2, m_sValue                   ) == ERROR_SUCCESS);
        verify(::MsiRecordSetInteger(hRecordProg, 3, dwError                    ) == ERROR_SUCCESS);
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(dwError);
    }
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpMoveFile
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpMoveFile::CMSITSCAOpMoveFile(LPCWSTR pszFileSrc, LPCWSTR pszFileDst, int iTicks) :
    CMSITSCAOpSrcDstStringOperation(pszFileSrc, pszFileDst, iTicks)
{
}


HRESULT CMSITSCAOpMoveFile::Execute(CMSITSCASession *pSession)
{
    assert(pSession);
    DWORD dwError;

    // Move the file.
    dwError = ::MoveFileW(m_sValue1, m_sValue2) ? ERROR_SUCCESS : ::GetLastError();
    if (dwError == ERROR_SUCCESS) {
        if (pSession->m_bRollbackEnabled) {
            // Order rollback action to move it back.
            pSession->m_olRollback.AddHead(new CMSITSCAOpMoveFile(m_sValue2, m_sValue1));
        }

        return S_OK;
    } else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(4);
        verify(::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_MOVE_FAILED) == ERROR_SUCCESS);
        verify(::MsiRecordSetStringW(hRecordProg, 2, m_sValue1                ) == ERROR_SUCCESS);
        verify(::MsiRecordSetStringW(hRecordProg, 3, m_sValue2                ) == ERROR_SUCCESS);
        verify(::MsiRecordSetInteger(hRecordProg, 4, dwError                  ) == ERROR_SUCCESS);
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(dwError);
    }
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpCreateTask
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpCreateTask::CMSITSCAOpCreateTask(LPCWSTR pszTaskName, int iTicks) :
    m_dwFlags(0),
    m_dwPriority(NORMAL_PRIORITY_CLASS),
    m_wIdleMinutes(0),
    m_wDeadlineMinutes(0),
    m_dwMaxRuntimeMS(INFINITE),
    CMSITSCAOpSingleStringOperation(pszTaskName, iTicks)
{
}


CMSITSCAOpCreateTask::~CMSITSCAOpCreateTask()
{
    // Clear the password in memory.
    int iLength = m_sPassword.GetLength();
    ::SecureZeroMemory(m_sPassword.GetBuffer(iLength), sizeof(WCHAR) * iLength);
    m_sPassword.ReleaseBuffer(0);
}


HRESULT CMSITSCAOpCreateTask::Execute(CMSITSCASession *pSession)
{
    HRESULT hr;
    POSITION pos;
    PMSIHANDLE hRecordMsg = ::MsiCreateRecord(1);
    CComPtr<ITaskService> pService;

    // Display our custom message in the progress bar.
    verify(::MsiRecordSetStringW(hRecordMsg, 1, m_sValue) == ERROR_SUCCESS);
    if (MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ACTIONDATA, hRecordMsg) == IDCANCEL)
        return AtlHresultFromWin32(ERROR_INSTALL_USEREXIT);

    {
        // Delete existing task first.
        // Since task deleting is a complicated job (when rollback/commit support is required), and we do have an operation just for that, we use it.
        // Don't worry, CMSITSCAOpDeleteTask::Execute() returns S_OK if task doesn't exist.
        CMSITSCAOpDeleteTask opDeleteTask(m_sValue);
        hr = opDeleteTask.Execute(pSession);
        if (FAILED(hr)) return hr;
    }

    hr = pService.CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER);
    if (SUCCEEDED(hr)) {
        // Windows Vista or newer.
        CComVariant vEmpty;
        CComPtr<ITaskDefinition> pTaskDefinition;
        CComPtr<ITaskSettings> pTaskSettings;
        CComPtr<IPrincipal> pPrincipal;
        CComPtr<IActionCollection> pActionCollection;
        CComPtr<IAction> pAction;
        CComPtr<IIdleSettings> pIdleSettings;
        CComPtr<IExecAction> pExecAction;
        CComPtr<IRegistrationInfo> pRegististrationInfo;
        CComPtr<ITriggerCollection> pTriggerCollection;
        CComPtr<IRegisteredTask> pTask;
        CStringW str;
        UINT iTrigger;
        TASK_LOGON_TYPE logonType;
        CComBSTR bstrContext(L"Author");

        // Connect to local task service.
        hr = pService->Connect(vEmpty, vEmpty, vEmpty, vEmpty);
        if (FAILED(hr)) return hr;

        // Prepare the definition for a new task.
        hr = pService->NewTask(0, &pTaskDefinition);
        if (FAILED(hr)) return hr;

        // Get the task's settings.
        hr = pTaskDefinition->get_Settings(&pTaskSettings);
        if (FAILED(hr)) return hr;

        // Configure Task Scheduler 1.0 compatible task.
        hr = pTaskSettings->put_Compatibility(TASK_COMPATIBILITY_V1);
        if (FAILED(hr)) return hr;

        // Enable/disable task.
        if (pSession->m_bRollbackEnabled && (m_dwFlags & TASK_FLAG_DISABLED) == 0) {
            // The task is supposed to be enabled.
            // However, we shall enable it at commit to prevent it from accidentally starting in the very installation phase.
            pSession->m_olCommit.AddTail(new CMSITSCAOpEnableTask(m_sValue, TRUE));
            m_dwFlags |= TASK_FLAG_DISABLED;
        }
        hr = pTaskSettings->put_Enabled(m_dwFlags & TASK_FLAG_DISABLED ? VARIANT_FALSE : VARIANT_TRUE); if (FAILED(hr)) return hr;

        // Get task actions.
        hr = pTaskDefinition->get_Actions(&pActionCollection);
        if (FAILED(hr)) return hr;

        // Add execute action.
        hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
        if (FAILED(hr)) return hr;
        hr = pAction.QueryInterface(&pExecAction);
        if (FAILED(hr)) return hr;

        // Configure the action.
        hr = pExecAction->put_Path            (CComBSTR(m_sApplicationName )); if (FAILED(hr)) return hr;
        hr = pExecAction->put_Arguments       (CComBSTR(m_sParameters      )); if (FAILED(hr)) return hr;
        hr = pExecAction->put_WorkingDirectory(CComBSTR(m_sWorkingDirectory)); if (FAILED(hr)) return hr;

        // Set task description.
        hr = pTaskDefinition->get_RegistrationInfo(&pRegististrationInfo);
        if (FAILED(hr)) return hr;
        hr = pRegististrationInfo->put_Description(CComBSTR(m_sComment));
        if (FAILED(hr)) return hr;

        hr = pRegististrationInfo->put_Author(CComBSTR(L"Amebis, d. o. o., Kamnik"));

        // Configure task "flags".
        if (m_dwFlags & TASK_FLAG_DELETE_WHEN_DONE) {
            hr = pTaskSettings->put_DeleteExpiredTaskAfter(CComBSTR(L"PT24H"));
            if (FAILED(hr)) return hr;
        }
        hr = pTaskSettings->put_Hidden(m_dwFlags & TASK_FLAG_HIDDEN ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) return hr;
        hr = pTaskSettings->put_WakeToRun(m_dwFlags & TASK_FLAG_SYSTEM_REQUIRED ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) return hr;
        hr = pTaskSettings->put_DisallowStartIfOnBatteries(m_dwFlags & TASK_FLAG_DONT_START_IF_ON_BATTERIES ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) return hr;
        hr = pTaskSettings->put_StopIfGoingOnBatteries(m_dwFlags & TASK_FLAG_KILL_IF_GOING_ON_BATTERIES ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) return hr;

        // Configure task priority.
        hr = pTaskSettings->put_Priority(
            m_dwPriority == REALTIME_PRIORITY_CLASS     ? 0 :
            m_dwPriority == HIGH_PRIORITY_CLASS         ? 1 :
            m_dwPriority == ABOVE_NORMAL_PRIORITY_CLASS ? 2 :
            m_dwPriority == NORMAL_PRIORITY_CLASS       ? 4 :
            m_dwPriority == BELOW_NORMAL_PRIORITY_CLASS ? 7 :
            m_dwPriority == IDLE_PRIORITY_CLASS         ? 9 : 7);
        if (FAILED(hr)) return hr;

        // Get task principal.
        hr = pTaskDefinition->get_Principal(&pPrincipal);
        if (FAILED(hr)) return hr;

        if (m_sAccountName.IsEmpty()) {
            logonType = TASK_LOGON_SERVICE_ACCOUNT;
            hr = pPrincipal->put_LogonType(logonType);             if (FAILED(hr)) return hr;
            hr = pPrincipal->put_UserId(CComBSTR(L"S-1-5-18"));    if (FAILED(hr)) return hr;
            hr = pPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);  if (FAILED(hr)) return hr;
        } else if (m_dwFlags & TASK_FLAG_RUN_ONLY_IF_LOGGED_ON) {
            logonType = TASK_LOGON_INTERACTIVE_TOKEN;
            hr = pPrincipal->put_LogonType(logonType);             if (FAILED(hr)) return hr;
            hr = pPrincipal->put_RunLevel(TASK_RUNLEVEL_LUA);      if (FAILED(hr)) return hr;
        } else {
            logonType = TASK_LOGON_PASSWORD;
            hr = pPrincipal->put_LogonType(logonType);             if (FAILED(hr)) return hr;
            hr = pPrincipal->put_UserId(CComBSTR(m_sAccountName)); if (FAILED(hr)) return hr;
            hr = pPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);  if (FAILED(hr)) return hr;
        }

        // Connect principal and action collection.
        hr = pPrincipal->put_Id(bstrContext);
        if (FAILED(hr)) return hr;
        hr = pActionCollection->put_Context(bstrContext);
        if (FAILED(hr)) return hr;

        // Configure idle settings.
        hr = pTaskSettings->put_RunOnlyIfIdle(m_dwFlags & TASK_FLAG_START_ONLY_IF_IDLE ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) return hr;
        hr = pTaskSettings->get_IdleSettings(&pIdleSettings);
        if (FAILED(hr)) return hr;
        str.Format(L"PT%uS", m_wIdleMinutes*60);
        hr = pIdleSettings->put_IdleDuration(CComBSTR(str));
        if (FAILED(hr)) return hr;
        str.Format(L"PT%uS", m_wDeadlineMinutes*60);
        hr = pIdleSettings->put_WaitTimeout(CComBSTR(str));
        if (FAILED(hr)) return hr;
        hr = pIdleSettings->put_RestartOnIdle(m_dwFlags & TASK_FLAG_RESTART_ON_IDLE_RESUME ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) return hr;
        hr = pIdleSettings->put_StopOnIdleEnd(m_dwFlags & TASK_FLAG_KILL_ON_IDLE_END ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) return hr;

        // Configure task runtime limit.
        str.Format(L"PT%uS", m_dwMaxRuntimeMS != INFINITE ? (m_dwMaxRuntimeMS + 500) / 1000 : 0);
        hr = pTaskSettings->put_ExecutionTimeLimit(CComBSTR(str));
        if (FAILED(hr)) return hr;

        // Get task trigger colection.
        hr = pTaskDefinition->get_Triggers(&pTriggerCollection);
        if (FAILED(hr)) return hr;

        // Add triggers.
        for (pos = m_lTriggers.GetHeadPosition(), iTrigger = 0; pos; iTrigger++) {
            CComPtr<ITrigger> pTrigger;
            TASK_TRIGGER &ttData = m_lTriggers.GetNext(pos);

            switch (ttData.TriggerType) {
                case TASK_TIME_TRIGGER_ONCE: {
                    CComPtr<ITimeTrigger> pTriggerTime;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_TIME, &pTrigger); if (FAILED(hr)) return hr;
                    hr = pTrigger.QueryInterface(&pTriggerTime);                   if (FAILED(hr)) return hr;
                    str.Format(L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerTime->put_RandomDelay(CComBSTR(str));             if (FAILED(hr)) return hr;
                    break;
                }

                case TASK_TIME_TRIGGER_DAILY: {
                    CComPtr<IDailyTrigger> pTriggerDaily;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_DAILY, &pTrigger);       if (FAILED(hr)) return hr;
                    hr = pTrigger.QueryInterface(&pTriggerDaily);                         if (FAILED(hr)) return hr;
                    hr = pTriggerDaily->put_DaysInterval(ttData.Type.Daily.DaysInterval); if (FAILED(hr)) return hr;
                    str.Format(L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerDaily->put_RandomDelay(CComBSTR(str));                   if (FAILED(hr)) return hr;
                    break;
                }

                case TASK_TIME_TRIGGER_WEEKLY: {
                    CComPtr<IWeeklyTrigger> pTriggerWeekly;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_WEEKLY, &pTrigger);          if (FAILED(hr)) return hr;
                    hr = pTrigger.QueryInterface(&pTriggerWeekly);                            if (FAILED(hr)) return hr;
                    hr = pTriggerWeekly->put_WeeksInterval(ttData.Type.Weekly.WeeksInterval); if (FAILED(hr)) return hr;
                    hr = pTriggerWeekly->put_DaysOfWeek(ttData.Type.Weekly.rgfDaysOfTheWeek); if (FAILED(hr)) return hr;
                    str.Format(L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerWeekly->put_RandomDelay(CComBSTR(str));                      if (FAILED(hr)) return hr;
                    break;
                }

                case TASK_TIME_TRIGGER_MONTHLYDATE: {
                    CComPtr<IMonthlyTrigger> pTriggerMonthly;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_MONTHLY, &pTrigger);          if (FAILED(hr)) return hr;
                    hr = pTrigger.QueryInterface(&pTriggerMonthly);                            if (FAILED(hr)) return hr;
                    hr = pTriggerMonthly->put_DaysOfMonth(ttData.Type.MonthlyDate.rgfDays);    if (FAILED(hr)) return hr;
                    hr = pTriggerMonthly->put_MonthsOfYear(ttData.Type.MonthlyDate.rgfMonths); if (FAILED(hr)) return hr;
                    str.Format(L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerMonthly->put_RandomDelay(CComBSTR(str));                      if (FAILED(hr)) return hr;
                    break;
                }

                case TASK_TIME_TRIGGER_MONTHLYDOW: {
                    CComPtr<IMonthlyDOWTrigger> pTriggerMonthlyDOW;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_MONTHLYDOW, &pTrigger);              if (FAILED(hr)) return hr;
                    hr = pTrigger.QueryInterface(&pTriggerMonthlyDOW);                                if (FAILED(hr)) return hr;
                    hr = pTriggerMonthlyDOW->put_WeeksOfMonth(
                        ttData.Type.MonthlyDOW.wWhichWeek == TASK_FIRST_WEEK  ? 0x01 :
                        ttData.Type.MonthlyDOW.wWhichWeek == TASK_SECOND_WEEK ? 0x02 :
                        ttData.Type.MonthlyDOW.wWhichWeek == TASK_THIRD_WEEK  ? 0x04 :
                        ttData.Type.MonthlyDOW.wWhichWeek == TASK_FOURTH_WEEK ? 0x08 :
                        ttData.Type.MonthlyDOW.wWhichWeek == TASK_LAST_WEEK   ? 0x10 : 0x00);         if (FAILED(hr)) return hr;
                    hr = pTriggerMonthlyDOW->put_DaysOfWeek(ttData.Type.MonthlyDOW.rgfDaysOfTheWeek); if (FAILED(hr)) return hr;
                    hr = pTriggerMonthlyDOW->put_MonthsOfYear(ttData.Type.MonthlyDOW.rgfMonths);      if (FAILED(hr)) return hr;
                    str.Format(L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerMonthlyDOW->put_RandomDelay(CComBSTR(str));                          if (FAILED(hr)) return hr;
                    break;
                }

                case TASK_EVENT_TRIGGER_ON_IDLE: {
                    hr = pTriggerCollection->Create(TASK_TRIGGER_IDLE, &pTrigger); if (FAILED(hr)) return hr;
                    break;
                }

                case TASK_EVENT_TRIGGER_AT_SYSTEMSTART: {
                    CComPtr<IBootTrigger> pTriggerBoot;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_BOOT, &pTrigger); if (FAILED(hr)) return hr;
                    hr = pTrigger.QueryInterface(&pTriggerBoot);                   if (FAILED(hr)) return hr;
                    str.Format(L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerBoot->put_Delay(CComBSTR(str));                   if (FAILED(hr)) return hr;
                    break;
                }

                case TASK_EVENT_TRIGGER_AT_LOGON: {
                    CComPtr<ILogonTrigger> pTriggerLogon;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_LOGON, &pTrigger); if (FAILED(hr)) return hr;
                    hr = pTrigger.QueryInterface(&pTriggerLogon);                   if (FAILED(hr)) return hr;
                    str.Format(L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerLogon->put_Delay(CComBSTR(str));                   if (FAILED(hr)) return hr;
                    break;
                }
            }

            // Set trigger ID.
            str.Format(L"%u", iTrigger);
            hr = pTrigger->put_Id(CComBSTR(str));
            if (FAILED(hr)) return hr;

            // Set trigger start date.
            str.Format(L"%04u-%02u-%02uT%02u:%02u:00", ttData.wBeginYear, ttData.wBeginMonth, ttData.wBeginDay, ttData.wStartHour, ttData.wStartMinute);
            hr = pTrigger->put_StartBoundary(CComBSTR(str));
            if (FAILED(hr)) return hr;

            if (ttData.rgFlags & TASK_TRIGGER_FLAG_HAS_END_DATE) {
                // Set trigger end date.
                str.Format(L"%04u-%02u-%02u", ttData.wEndYear, ttData.wEndMonth, ttData.wEndDay);
                hr = pTrigger->put_EndBoundary(CComBSTR(str));
                if (FAILED(hr)) return hr;
            }

            // Set trigger repetition duration and interval.
            if (ttData.MinutesDuration || ttData.MinutesInterval) {
                CComPtr<IRepetitionPattern> pRepetitionPattern;

                hr = pTrigger->get_Repetition(&pRepetitionPattern);
                if (FAILED(hr)) return hr;
                str.Format(L"PT%uM", ttData.MinutesDuration);
                hr = pRepetitionPattern->put_Duration(CComBSTR(str));
                if (FAILED(hr)) return hr;
                str.Format(L"PT%uM", ttData.MinutesInterval);
                hr = pRepetitionPattern->put_Interval(CComBSTR(str));
                if (FAILED(hr)) return hr;
                hr = pRepetitionPattern->put_StopAtDurationEnd(ttData.rgFlags & TASK_TRIGGER_FLAG_KILL_AT_DURATION_END ? VARIANT_TRUE : VARIANT_FALSE);
                if (FAILED(hr)) return hr;
            }

            // Enable/disable trigger.
            hr = pTrigger->put_Enabled(ttData.rgFlags & TASK_TRIGGER_FLAG_DISABLED ? VARIANT_FALSE : VARIANT_TRUE);
            if (FAILED(hr)) return hr;
        }

        // Get the task folder.
        CComPtr<ITaskFolder> pTaskFolder;
        hr = pService->GetFolder(CComBSTR(L"\\"), &pTaskFolder);
        if (FAILED(hr)) return hr;

#ifdef _DEBUG
        CComBSTR xml;
        hr = pTaskDefinition->get_XmlText(&xml);
#endif

        // Register the task.
        hr = pTaskFolder->RegisterTaskDefinition(
            CComBSTR(m_sValue), // path
            pTaskDefinition,    // pDefinition
            TASK_CREATE,        // flags
            vEmpty,             // userId
            logonType != TASK_LOGON_SERVICE_ACCOUNT && !m_sPassword.IsEmpty() ? CComVariant(m_sPassword) : vEmpty, // password
            logonType,          // logonType
            vEmpty,             // sddl
            &pTask);            // ppTask
        if (FAILED(hr)) return hr;
    } else {
        // Windows XP or older.
        CComPtr<ITask> pTask;

        // Create the new task.
        hr = pSession->m_pTaskScheduler->NewWorkItem(m_sValue, CLSID_CTask, IID_ITask, (IUnknown**)&pTask);
        if (pSession->m_bRollbackEnabled) {
            // Order rollback action to delete the task. ITask::NewWorkItem() might made a blank task already.
            pSession->m_olRollback.AddHead(new CMSITSCAOpDeleteTask(m_sValue));
        }
        if (FAILED(hr)) return hr;

        // Set its properties.
        hr = pTask->SetApplicationName (m_sApplicationName ); if (FAILED(hr)) return hr;
        hr = pTask->SetParameters      (m_sParameters      ); if (FAILED(hr)) return hr;
        hr = pTask->SetWorkingDirectory(m_sWorkingDirectory); if (FAILED(hr)) return hr;
        hr = pTask->SetComment         (m_sComment         ); if (FAILED(hr)) return hr;
        if (pSession->m_bRollbackEnabled && (m_dwFlags & TASK_FLAG_DISABLED) == 0) {
            // The task is supposed to be enabled.
            // However, we shall enable it at commit to prevent it from accidentally starting in the very installation phase.
            pSession->m_olCommit.AddTail(new CMSITSCAOpEnableTask(m_sValue, TRUE));
            m_dwFlags |= TASK_FLAG_DISABLED;
        }
        hr = pTask->SetFlags           (m_dwFlags          ); if (FAILED(hr)) return hr;
        hr = pTask->SetPriority        (m_dwPriority       ); if (FAILED(hr)) return hr;

        if (!m_sAccountName.IsEmpty()) {
            hr = pTask->SetAccountInformation(m_sAccountName, m_sPassword.IsEmpty() ? NULL : m_sPassword);
            if (FAILED(hr)) return hr;
        }

        hr = pTask->SetIdleWait  (m_wIdleMinutes, m_wDeadlineMinutes); if (FAILED(hr)) return hr;
        hr = pTask->SetMaxRunTime(m_dwMaxRuntimeMS                  ); if (FAILED(hr)) return hr;

        // Add triggers.
        for (pos = m_lTriggers.GetHeadPosition(); pos;) {
            WORD wTriggerIdx;
            CComPtr<ITaskTrigger> pTrigger;
            TASK_TRIGGER &ttData = m_lTriggers.GetNext(pos);

            hr = pTask->CreateTrigger(&wTriggerIdx, &pTrigger);
            if (FAILED(hr)) return hr;

            hr = pTrigger->SetTrigger(&ttData);
            if (FAILED(hr)) return hr;
        }

        // Save the task.
        CComQIPtr<IPersistFile> pTaskFile(pTask);
        if (!pTaskFile) return E_NOINTERFACE;

        hr = pTaskFile->Save(NULL, TRUE);
        if (FAILED(hr)) return hr;
    }

    return S_OK;
}


UINT CMSITSCAOpCreateTask::SetFromRecord(MSIHANDLE hInstall, MSIHANDLE hRecord)
{
    UINT uiResult;
    int iValue;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord,  3, m_sApplicationName);
    if (uiResult != ERROR_SUCCESS) return uiResult;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord,  4, m_sParameters);
    if (uiResult != ERROR_SUCCESS) return uiResult;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord,  5, m_sWorkingDirectory);
    if (uiResult != ERROR_SUCCESS) return uiResult;
    if (!m_sWorkingDirectory.IsEmpty() && m_sWorkingDirectory.GetAt(m_sWorkingDirectory.GetLength() - 1) == L'\\') {
        // Trim trailing backslash.
        m_sWorkingDirectory.Truncate(m_sWorkingDirectory.GetLength() - 1);
    }

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord, 10, m_sComment);
    if (uiResult != ERROR_SUCCESS) return uiResult;

    m_dwFlags = ::MsiRecordGetInteger(hRecord,  6);
    if (m_dwFlags == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;

    m_dwPriority = ::MsiRecordGetInteger(hRecord,  7);
    if (m_dwPriority == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord,  8, m_sAccountName);
    if (uiResult != ERROR_SUCCESS) return uiResult;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord,  9, m_sPassword);
    if (uiResult != ERROR_SUCCESS) return uiResult;

    iValue = ::MsiRecordGetInteger(hRecord, 11);
    m_wIdleMinutes = iValue != MSI_NULL_INTEGER ? (WORD)iValue : 0;

    iValue = ::MsiRecordGetInteger(hRecord, 12);
    m_wDeadlineMinutes = iValue != MSI_NULL_INTEGER ? (WORD)iValue : 0;

    m_dwMaxRuntimeMS = ::MsiRecordGetInteger(hRecord, 13);
    if (m_dwMaxRuntimeMS == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;

    return S_OK;
}


UINT CMSITSCAOpCreateTask::SetTriggersFromView(MSIHANDLE hView)
{
    for (;;) {
        UINT uiResult;
        PMSIHANDLE hRecord;
        TASK_TRIGGER ttData;
        ULONGLONG ullValue;
        FILETIME ftValue;
        SYSTEMTIME stValue;
        int iValue;

        // Fetch one record from the view.
        uiResult = ::MsiViewFetch(hView, &hRecord);
        if (uiResult == ERROR_NO_MORE_ITEMS) return ERROR_SUCCESS;
        else if (uiResult != ERROR_SUCCESS)  return uiResult;

        ZeroMemory(&ttData, sizeof(TASK_TRIGGER));
        ttData.cbTriggerSize = sizeof(TASK_TRIGGER);

        // Get StartDate.
        iValue = ::MsiRecordGetInteger(hRecord, 2);
        if (iValue == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;
        ullValue = ((ULONGLONG)iValue + 138426) * 864000000000;
        ftValue.dwHighDateTime = ullValue >> 32;
        ftValue.dwLowDateTime  = ullValue & 0xffffffff;
        if (!::FileTimeToSystemTime(&ftValue, &stValue))
            return ::GetLastError();
        ttData.wBeginYear  = stValue.wYear;
        ttData.wBeginMonth = stValue.wMonth;
        ttData.wBeginDay   = stValue.wDay;

        // Get EndDate.
        iValue = ::MsiRecordGetInteger(hRecord, 3);
        if (iValue != MSI_NULL_INTEGER) {
            ullValue = ((ULONGLONG)iValue + 138426) * 864000000000;
            ftValue.dwHighDateTime = ullValue >> 32;
            ftValue.dwLowDateTime  = ullValue & 0xffffffff;
            if (!::FileTimeToSystemTime(&ftValue, &stValue))
                return ::GetLastError();
            ttData.wEndYear  = stValue.wYear;
            ttData.wEndMonth = stValue.wMonth;
            ttData.wEndDay   = stValue.wDay;
            ttData.rgFlags  |= TASK_TRIGGER_FLAG_HAS_END_DATE;
        }

        // Get StartTime.
        iValue = ::MsiRecordGetInteger(hRecord, 4);
        if (iValue == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;
        ttData.wStartHour   = (WORD)(iValue / 60);
        ttData.wStartMinute = (WORD)(iValue % 60);

        // Get StartTimeRand.
        iValue = ::MsiRecordGetInteger(hRecord, 5);
        ttData.wRandomMinutesInterval = iValue != MSI_NULL_INTEGER ? (WORD)iValue : 0;

        // Get MinutesDuration.
        iValue = ::MsiRecordGetInteger(hRecord, 6);
        ttData.MinutesDuration = iValue != MSI_NULL_INTEGER ? iValue : 0;

        // Get MinutesInterval.
        iValue = ::MsiRecordGetInteger(hRecord, 7);
        ttData.MinutesInterval = iValue != MSI_NULL_INTEGER ? iValue : 0;

        // Get Flags.
        iValue = ::MsiRecordGetInteger(hRecord, 8);
        if (iValue == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;
        ttData.rgFlags |= iValue & ~TASK_TRIGGER_FLAG_HAS_END_DATE;

        // Get Type.
        iValue = ::MsiRecordGetInteger(hRecord,  9);
        if (iValue == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;
        ttData.TriggerType = (TASK_TRIGGER_TYPE)iValue;

        switch (ttData.TriggerType) {
        case TASK_TIME_TRIGGER_DAILY:
            // Get DaysInterval.
            iValue = ::MsiRecordGetInteger(hRecord, 10);
            if (iValue == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;
            ttData.Type.Daily.DaysInterval = (WORD)iValue;
            break;

        case TASK_TIME_TRIGGER_WEEKLY:
            // Get WeeksInterval.
            iValue = ::MsiRecordGetInteger(hRecord, 11);
            if (iValue == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;
            ttData.Type.Weekly.WeeksInterval = (WORD)iValue;

            // Get DaysOfTheWeek.
            iValue = ::MsiRecordGetInteger(hRecord, 12);
            if (iValue == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;
            ttData.Type.Weekly.rgfDaysOfTheWeek = (WORD)iValue;
            break;

        case TASK_TIME_TRIGGER_MONTHLYDATE:
            // Get DaysOfMonth.
            iValue = ::MsiRecordGetInteger(hRecord, 13);
            if (iValue == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;
            ttData.Type.MonthlyDate.rgfDays = (WORD)iValue;

            // Get MonthsOfYear.
            iValue = ::MsiRecordGetInteger(hRecord, 15);
            if (iValue == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;
            ttData.Type.MonthlyDate.rgfMonths = (WORD)iValue;
            break;

        case TASK_TIME_TRIGGER_MONTHLYDOW:
            // Get WeekOfMonth.
            iValue = ::MsiRecordGetInteger(hRecord, 14);
            if (iValue == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;
            ttData.Type.MonthlyDOW.wWhichWeek = (WORD)iValue;

            // Get DaysOfTheWeek.
            iValue = ::MsiRecordGetInteger(hRecord, 12);
            if (iValue == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;
            ttData.Type.MonthlyDOW.rgfDaysOfTheWeek = (WORD)iValue;

            // Get MonthsOfYear.
            iValue = ::MsiRecordGetInteger(hRecord, 15);
            if (iValue == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;
            ttData.Type.MonthlyDOW.rgfMonths = (WORD)iValue;
            break;
        }

        m_lTriggers.AddTail(ttData);
    }

    return S_OK;
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpDeleteTask
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpDeleteTask::CMSITSCAOpDeleteTask(LPCWSTR pszTaskName, int iTicks) :
    CMSITSCAOpSingleStringOperation(pszTaskName, iTicks)
{
}


HRESULT CMSITSCAOpDeleteTask::Execute(CMSITSCASession *pSession)
{
    assert(pSession);
    HRESULT hr;

    if (pSession->m_bRollbackEnabled) {
        CComPtr<ITask> pTask;
        DWORD dwFlags;
        CString sDisplayNameOrig;
        UINT uiCount = 0;

        // Load the task.
        hr = pSession->m_pTaskScheduler->Activate(m_sValue, IID_ITask, (IUnknown**)&pTask);
        if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) return S_OK;
        else if (FAILED(hr)) return hr;

        // Disable the task.
        hr = pTask->GetFlags(&dwFlags);
        if (FAILED(hr)) return hr;
        if ((dwFlags & TASK_FLAG_DISABLED) == 0) {
            // The task is enabled.

            // In case the task disabling fails halfway, try to re-enable it anyway on rollback.
            pSession->m_olRollback.AddHead(new CMSITSCAOpEnableTask(m_sValue, TRUE));

            // Disable it.
            dwFlags |= TASK_FLAG_DISABLED;
            hr = pTask->SetFlags(dwFlags);
            if (FAILED(hr)) return hr;
        }

        // Prepare a backup copy of task.
        do {
            sDisplayNameOrig.Format(L"%ls (orig %u)", (LPCWSTR)m_sValue, ++uiCount);
            hr = pSession->m_pTaskScheduler->AddWorkItem(sDisplayNameOrig, pTask);
        } while (hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS));
        // In case the backup copy creation failed halfway, try to delete its remains anyway on rollback.
        pSession->m_olRollback.AddHead(new CMSITSCAOpDeleteTask(sDisplayNameOrig));
        if (FAILED(hr)) return hr;

        // Save the backup copy.
        CComQIPtr<IPersistFile> pTaskFile(pTask);
        if (!pTaskFile) return E_NOINTERFACE;

        hr = pTaskFile->Save(NULL, TRUE);
        if (FAILED(hr)) return hr;

        // Order rollback action to restore from backup copy.
        pSession->m_olRollback.AddHead(new CMSITSCAOpCopyTask(sDisplayNameOrig, m_sValue));

        // Delete it.
        hr = pSession->m_pTaskScheduler->Delete(m_sValue);
        if (FAILED(hr)) return hr;

        // Order commit action to delete backup copy.
        pSession->m_olCommit.AddTail(new CMSITSCAOpDeleteTask(sDisplayNameOrig));

        return S_OK;
    } else {
        // Delete the task.
        hr = pSession->m_pTaskScheduler->Delete(m_sValue);
        if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) return S_OK;
        else return hr;
    }
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpEnableTask
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpEnableTask::CMSITSCAOpEnableTask(LPCWSTR pszTaskName, BOOL bEnable, int iTicks) :
    m_bEnable(bEnable),
    CMSITSCAOpSingleStringOperation(pszTaskName, iTicks)
{
}


HRESULT CMSITSCAOpEnableTask::Execute(CMSITSCASession *pSession)
{
    assert(pSession);
    HRESULT hr;
    CComPtr<ITask> pTask;

    // Load the task.
    hr = pSession->m_pTaskScheduler->Activate(m_sValue, IID_ITask, (IUnknown**)&pTask);
    if (SUCCEEDED(hr)) {
        DWORD dwFlags;

        // Get task's current flags.
        hr = pTask->GetFlags(&dwFlags);
        if (SUCCEEDED(hr)) {
            // Modify flags.
            if (m_bEnable) {
                if (pSession->m_bRollbackEnabled && (dwFlags & TASK_FLAG_DISABLED)) {
                    // The task is disabled and supposed to be enabled.
                    // However, we shall enable it at commit to prevent it from accidentally starting in the very installation phase.
                    pSession->m_olCommit.AddTail(new CMSITSCAOpEnableTask(m_sValue, TRUE));
                    return S_OK;
                } else
                    dwFlags &= ~TASK_FLAG_DISABLED;
            } else {
                if (pSession->m_bRollbackEnabled && !(dwFlags & TASK_FLAG_DISABLED)) {
                    // The task is enabled and we will disable it now.
                    // Order rollback to re-enable it.
                    pSession->m_olRollback.AddHead(new CMSITSCAOpEnableTask(m_sValue, TRUE));
                }
                dwFlags |=  TASK_FLAG_DISABLED;
            }

            // Set flags.
            hr = pTask->SetFlags(dwFlags);
            if (SUCCEEDED(hr)) {
                // Save the task.
                CComQIPtr<IPersistFile> pTaskFile(pTask);
                hr = pTaskFile ? pTaskFile->Save(NULL, TRUE) : E_NOINTERFACE;
            }
        }
    }

    return hr;
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpCopyTask
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpCopyTask::CMSITSCAOpCopyTask(LPCWSTR pszTaskSrc, LPCWSTR pszTaskDst, int iTicks) :
    CMSITSCAOpSrcDstStringOperation(pszTaskSrc, pszTaskDst, iTicks)
{
}


HRESULT CMSITSCAOpCopyTask::Execute(CMSITSCASession *pSession)
{
    assert(pSession);
    HRESULT hr;
    CComPtr<ITask> pTask;

    // Load the source task.
    hr = pSession->m_pTaskScheduler->Activate(m_sValue1, IID_ITask, (IUnknown**)&pTask);
    if (FAILED(hr)) return hr;

    // Add with different name.
    hr = pSession->m_pTaskScheduler->AddWorkItem(m_sValue2, pTask);
    if (pSession->m_bRollbackEnabled) {
        // Order rollback action to delete the new copy. ITask::AddWorkItem() might made a blank task already.
        pSession->m_olRollback.AddHead(new CMSITSCAOpDeleteTask(m_sValue2));
    }
    if (FAILED(hr)) return hr;

    // Save the task.
    CComQIPtr<IPersistFile> pTaskFile(pTask);
    if (!pTaskFile) return E_NOINTERFACE;

    hr = pTaskFile->Save(NULL, TRUE);
    if (FAILED(hr)) return hr;

    return S_OK;
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpList
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpList::CMSITSCAOpList(int iTicks) :
    CMSITSCAOp(iTicks),
    CAtlList<CMSITSCAOp*>(sizeof(CMSITSCAOp*))
{
}


void CMSITSCAOpList::Free()
{
    POSITION pos;

    for (pos = GetHeadPosition(); pos;) {
        CMSITSCAOp *pOp = GetNext(pos);
        CMSITSCAOpList *pOpList = dynamic_cast<CMSITSCAOpList*>(pOp);

        if (pOpList) {
            // Recursivelly free sublists.
            pOpList->Free();
        }
        delete pOp;
    }

    RemoveAll();
}


HRESULT CMSITSCAOpList::LoadFromFile(LPCTSTR pszFileName)
{
    assert(pszFileName);
    HRESULT hr;
    CAtlFile fSequence;

    hr = fSequence.Create(pszFileName, GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN);
    if (FAILED(hr)) return hr;

    // Load operation sequence.
    return fSequence >> *this;
}


HRESULT CMSITSCAOpList::SaveToFile(LPCTSTR pszFileName) const
{
    assert(pszFileName);
    HRESULT hr;
    CAtlFile fSequence;

    hr = fSequence.Create(pszFileName, GENERIC_WRITE, FILE_SHARE_READ, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN);
    if (FAILED(hr)) return hr;

    // Save execute sequence to file.
    hr = fSequence << *this;
    fSequence.Close();

    if (FAILED(hr)) ::DeleteFile(pszFileName);
    return hr;
}


HRESULT CMSITSCAOpList::Execute(CMSITSCASession *pSession)
{
    assert(pSession);
    POSITION pos;
    HRESULT hr;
    PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);

    // Tell the installer to use explicit progress messages.
    verify(::MsiRecordSetInteger(hRecordProg, 1, 1) == ERROR_SUCCESS);
    verify(::MsiRecordSetInteger(hRecordProg, 2, 1) == ERROR_SUCCESS);
    verify(::MsiRecordSetInteger(hRecordProg, 3, 0) == ERROR_SUCCESS);
    ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_PROGRESS, hRecordProg);

    // Prepare hRecordProg for progress messages.
    verify(::MsiRecordSetInteger(hRecordProg, 1, 2) == ERROR_SUCCESS);
    verify(::MsiRecordSetInteger(hRecordProg, 3, 0) == ERROR_SUCCESS);

    for (pos = GetHeadPosition(); pos;) {
        CMSITSCAOp *pOp = GetNext(pos);
        assert(pOp);

        hr = pOp->Execute(pSession);
        if (!pSession->m_bContinueOnError && FAILED(hr)) return hr;

        verify(::MsiRecordSetInteger(hRecordProg, 2, pOp->m_iTicks) == ERROR_SUCCESS);
        if (::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_PROGRESS, hRecordProg) == IDCANCEL)
            return AtlHresultFromWin32(ERROR_INSTALL_USEREXIT);
    }

    verify(::MsiRecordSetInteger(hRecordProg, 2, m_iTicks) == ERROR_SUCCESS);
    ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_PROGRESS, hRecordProg);

    return S_OK;
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCASession
////////////////////////////////////////////////////////////////////////////

CMSITSCASession::CMSITSCASession() :
    m_bContinueOnError(FALSE),
    m_bRollbackEnabled(FALSE)
{
}
