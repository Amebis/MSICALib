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
// CMSITSCAOpTypeSingleString
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpTypeSingleString::CMSITSCAOpTypeSingleString(LPCWSTR pszValue, int iTicks) :
    m_sValue(pszValue),
    CMSITSCAOp(iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpDoubleStringOperation
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpTypeSrcDstString::CMSITSCAOpTypeSrcDstString(LPCWSTR pszValue1, LPCWSTR pszValue2, int iTicks) :
    m_sValue1(pszValue1),
    m_sValue2(pszValue2),
    CMSITSCAOp(iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpTypeBoolean
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpTypeBoolean::CMSITSCAOpTypeBoolean(BOOL bValue, int iTicks) :
    m_bValue(bValue),
    CMSITSCAOp(iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpRollbackEnable
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpRollbackEnable::CMSITSCAOpRollbackEnable(BOOL bEnable, int iTicks) :
    CMSITSCAOpTypeBoolean(bEnable, iTicks)
{
}


HRESULT CMSITSCAOpRollbackEnable::Execute(CMSITSCASession *pSession)
{
    assert(pSession);

    pSession->m_bRollbackEnabled = m_bValue;
    return S_OK;
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpFileDelete
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpFileDelete::CMSITSCAOpFileDelete(LPCWSTR pszFileName, int iTicks) :
    CMSITSCAOpTypeSingleString(pszFileName, iTicks)
{
}


HRESULT CMSITSCAOpFileDelete::Execute(CMSITSCASession *pSession)
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
            pSession->m_olRollback.AddHead(new CMSITSCAOpFileMove(sBackupName, m_sValue));

            // Order commit action to delete backup copy.
            pSession->m_olCommit.AddTail(new CMSITSCAOpFileDelete(sBackupName));
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
// CMSITSCAOpFileMove
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpFileMove::CMSITSCAOpFileMove(LPCWSTR pszFileSrc, LPCWSTR pszFileDst, int iTicks) :
    CMSITSCAOpTypeSrcDstString(pszFileSrc, pszFileDst, iTicks)
{
}


HRESULT CMSITSCAOpFileMove::Execute(CMSITSCASession *pSession)
{
    assert(pSession);
    DWORD dwError;

    // Move the file.
    dwError = ::MoveFileW(m_sValue1, m_sValue2) ? ERROR_SUCCESS : ::GetLastError();
    if (dwError == ERROR_SUCCESS) {
        if (pSession->m_bRollbackEnabled) {
            // Order rollback action to move it back.
            pSession->m_olRollback.AddHead(new CMSITSCAOpFileMove(m_sValue2, m_sValue1));
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
// CMSITSCAOpTaskCreate
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpTaskCreate::CMSITSCAOpTaskCreate(LPCWSTR pszTaskName, int iTicks) :
    m_dwFlags(0),
    m_dwPriority(NORMAL_PRIORITY_CLASS),
    m_wIdleMinutes(0),
    m_wDeadlineMinutes(0),
    m_dwMaxRuntimeMS(INFINITE),
    CMSITSCAOpTypeSingleString(pszTaskName, iTicks)
{
}


CMSITSCAOpTaskCreate::~CMSITSCAOpTaskCreate()
{
    // Clear the password in memory.
    int iLength = m_sPassword.GetLength();
    ::SecureZeroMemory(m_sPassword.GetBuffer(iLength), sizeof(WCHAR) * iLength);
    m_sPassword.ReleaseBuffer(0);
}


HRESULT CMSITSCAOpTaskCreate::Execute(CMSITSCASession *pSession)
{
    HRESULT hr;
    PMSIHANDLE hRecordMsg = ::MsiCreateRecord(1);
    CComPtr<ITaskService> pService;
    POSITION pos;

    // Display our custom message in the progress bar.
    verify(::MsiRecordSetStringW(hRecordMsg, 1, m_sValue) == ERROR_SUCCESS);
    if (MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ACTIONDATA, hRecordMsg) == IDCANCEL)
        return AtlHresultFromWin32(ERROR_INSTALL_USEREXIT);

    {
        // Delete existing task first.
        // Since task deleting is a complicated job (when rollback/commit support is required), and we do have an operation just for that, we use it.
        // Don't worry, CMSITSCAOpTaskDelete::Execute() returns S_OK if task doesn't exist.
        CMSITSCAOpTaskDelete opDeleteTask(m_sValue);
        hr = opDeleteTask.Execute(pSession);
        if (FAILED(hr)) goto finish;
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
        CComPtr<ITaskFolder> pTaskFolder;
        CComPtr<IRegisteredTask> pTask;
        CStringW str;
        UINT iTrigger;
        TASK_LOGON_TYPE logonType;
        CComBSTR bstrContext(L"Author");

        // Connect to local task service.
        hr = pService->Connect(vEmpty, vEmpty, vEmpty, vEmpty);
        if (FAILED(hr)) goto finish;

        // Prepare the definition for a new task.
        hr = pService->NewTask(0, &pTaskDefinition);
        if (FAILED(hr)) goto finish;

        // Get the task's settings.
        hr = pTaskDefinition->get_Settings(&pTaskSettings);
        if (FAILED(hr)) goto finish;

        // Enable/disable task.
        if (pSession->m_bRollbackEnabled && (m_dwFlags & TASK_FLAG_DISABLED) == 0) {
            // The task is supposed to be enabled.
            // However, we shall enable it at commit to prevent it from accidentally starting in the very installation phase.
            pSession->m_olCommit.AddTail(new CMSITSCAOpTaskEnable(m_sValue, TRUE));
            m_dwFlags |= TASK_FLAG_DISABLED;
        }
        hr = pTaskSettings->put_Enabled(m_dwFlags & TASK_FLAG_DISABLED ? VARIANT_FALSE : VARIANT_TRUE); if (FAILED(hr)) goto finish;

        // Get task actions.
        hr = pTaskDefinition->get_Actions(&pActionCollection);
        if (FAILED(hr)) goto finish;

        // Add execute action.
        hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
        if (FAILED(hr)) goto finish;
        hr = pAction.QueryInterface(&pExecAction);
        if (FAILED(hr)) goto finish;

        // Configure the action.
        hr = pExecAction->put_Path            (CComBSTR(m_sApplicationName )); if (FAILED(hr)) goto finish;
        hr = pExecAction->put_Arguments       (CComBSTR(m_sParameters      )); if (FAILED(hr)) goto finish;
        hr = pExecAction->put_WorkingDirectory(CComBSTR(m_sWorkingDirectory)); if (FAILED(hr)) goto finish;

        // Set task description.
        hr = pTaskDefinition->get_RegistrationInfo(&pRegististrationInfo);
        if (FAILED(hr)) goto finish;
        hr = pRegististrationInfo->put_Author(CComBSTR(m_sAuthor));
        if (FAILED(hr)) goto finish;
        hr = pRegististrationInfo->put_Description(CComBSTR(m_sComment));
        if (FAILED(hr)) goto finish;

        // Configure task "flags".
        if (m_dwFlags & TASK_FLAG_DELETE_WHEN_DONE) {
            hr = pTaskSettings->put_DeleteExpiredTaskAfter(CComBSTR(L"PT24H"));
            if (FAILED(hr)) goto finish;
        }
        hr = pTaskSettings->put_Hidden(m_dwFlags & TASK_FLAG_HIDDEN ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) goto finish;
        hr = pTaskSettings->put_WakeToRun(m_dwFlags & TASK_FLAG_SYSTEM_REQUIRED ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) goto finish;
        hr = pTaskSettings->put_DisallowStartIfOnBatteries(m_dwFlags & TASK_FLAG_DONT_START_IF_ON_BATTERIES ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) goto finish;
        hr = pTaskSettings->put_StopIfGoingOnBatteries(m_dwFlags & TASK_FLAG_KILL_IF_GOING_ON_BATTERIES ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) goto finish;

        // Configure task priority.
        hr = pTaskSettings->put_Priority(
            m_dwPriority == REALTIME_PRIORITY_CLASS     ? 0 :
            m_dwPriority == HIGH_PRIORITY_CLASS         ? 1 :
            m_dwPriority == ABOVE_NORMAL_PRIORITY_CLASS ? 2 :
            m_dwPriority == NORMAL_PRIORITY_CLASS       ? 4 :
            m_dwPriority == BELOW_NORMAL_PRIORITY_CLASS ? 7 :
            m_dwPriority == IDLE_PRIORITY_CLASS         ? 9 : 7);
        if (FAILED(hr)) goto finish;

        // Get task principal.
        hr = pTaskDefinition->get_Principal(&pPrincipal);
        if (FAILED(hr)) goto finish;

        if (m_sAccountName.IsEmpty()) {
            logonType = TASK_LOGON_SERVICE_ACCOUNT;
            hr = pPrincipal->put_LogonType(logonType);             if (FAILED(hr)) goto finish;
            hr = pPrincipal->put_UserId(CComBSTR(L"S-1-5-18"));    if (FAILED(hr)) goto finish;
            hr = pPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);  if (FAILED(hr)) goto finish;
        } else if (m_dwFlags & TASK_FLAG_RUN_ONLY_IF_LOGGED_ON) {
            logonType = TASK_LOGON_INTERACTIVE_TOKEN;
            hr = pPrincipal->put_LogonType(logonType);             if (FAILED(hr)) goto finish;
            hr = pPrincipal->put_RunLevel(TASK_RUNLEVEL_LUA);      if (FAILED(hr)) goto finish;
        } else {
            logonType = TASK_LOGON_PASSWORD;
            hr = pPrincipal->put_LogonType(logonType);             if (FAILED(hr)) goto finish;
            hr = pPrincipal->put_UserId(CComBSTR(m_sAccountName)); if (FAILED(hr)) goto finish;
            hr = pPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);  if (FAILED(hr)) goto finish;
        }

        // Connect principal and action collection.
        hr = pPrincipal->put_Id(bstrContext);
        if (FAILED(hr)) goto finish;
        hr = pActionCollection->put_Context(bstrContext);
        if (FAILED(hr)) goto finish;

        // Configure idle settings.
        hr = pTaskSettings->put_RunOnlyIfIdle(m_dwFlags & TASK_FLAG_START_ONLY_IF_IDLE ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) goto finish;
        hr = pTaskSettings->get_IdleSettings(&pIdleSettings);
        if (FAILED(hr)) goto finish;
        str.Format(L"PT%uS", m_wIdleMinutes*60);
        hr = pIdleSettings->put_IdleDuration(CComBSTR(str));
        if (FAILED(hr)) goto finish;
        str.Format(L"PT%uS", m_wDeadlineMinutes*60);
        hr = pIdleSettings->put_WaitTimeout(CComBSTR(str));
        if (FAILED(hr)) goto finish;
        hr = pIdleSettings->put_RestartOnIdle(m_dwFlags & TASK_FLAG_RESTART_ON_IDLE_RESUME ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) goto finish;
        hr = pIdleSettings->put_StopOnIdleEnd(m_dwFlags & TASK_FLAG_KILL_ON_IDLE_END ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) goto finish;

        // Configure task runtime limit.
        str.Format(L"PT%uS", m_dwMaxRuntimeMS != INFINITE ? (m_dwMaxRuntimeMS + 500) / 1000 : 0);
        hr = pTaskSettings->put_ExecutionTimeLimit(CComBSTR(str));
        if (FAILED(hr)) goto finish;

        // Get task trigger colection.
        hr = pTaskDefinition->get_Triggers(&pTriggerCollection);
        if (FAILED(hr)) goto finish;

        // Add triggers.
        for (pos = m_lTriggers.GetHeadPosition(), iTrigger = 0; pos; iTrigger++) {
            CComPtr<ITrigger> pTrigger;
            TASK_TRIGGER &ttData = m_lTriggers.GetNext(pos);

            switch (ttData.TriggerType) {
                case TASK_TIME_TRIGGER_ONCE: {
                    CComPtr<ITimeTrigger> pTriggerTime;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_TIME, &pTrigger); if (FAILED(hr)) goto finish;
                    hr = pTrigger.QueryInterface(&pTriggerTime);                   if (FAILED(hr)) goto finish;
                    str.Format(L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerTime->put_RandomDelay(CComBSTR(str));             if (FAILED(hr)) goto finish;
                    break;
                }

                case TASK_TIME_TRIGGER_DAILY: {
                    CComPtr<IDailyTrigger> pTriggerDaily;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_DAILY, &pTrigger);       if (FAILED(hr)) goto finish;
                    hr = pTrigger.QueryInterface(&pTriggerDaily);                         if (FAILED(hr)) goto finish;
                    hr = pTriggerDaily->put_DaysInterval(ttData.Type.Daily.DaysInterval); if (FAILED(hr)) goto finish;
                    str.Format(L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerDaily->put_RandomDelay(CComBSTR(str));                   if (FAILED(hr)) goto finish;
                    break;
                }

                case TASK_TIME_TRIGGER_WEEKLY: {
                    CComPtr<IWeeklyTrigger> pTriggerWeekly;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_WEEKLY, &pTrigger);          if (FAILED(hr)) goto finish;
                    hr = pTrigger.QueryInterface(&pTriggerWeekly);                            if (FAILED(hr)) goto finish;
                    hr = pTriggerWeekly->put_WeeksInterval(ttData.Type.Weekly.WeeksInterval); if (FAILED(hr)) goto finish;
                    hr = pTriggerWeekly->put_DaysOfWeek(ttData.Type.Weekly.rgfDaysOfTheWeek); if (FAILED(hr)) goto finish;
                    str.Format(L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerWeekly->put_RandomDelay(CComBSTR(str));                      if (FAILED(hr)) goto finish;
                    break;
                }

                case TASK_TIME_TRIGGER_MONTHLYDATE: {
                    CComPtr<IMonthlyTrigger> pTriggerMonthly;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_MONTHLY, &pTrigger);          if (FAILED(hr)) goto finish;
                    hr = pTrigger.QueryInterface(&pTriggerMonthly);                            if (FAILED(hr)) goto finish;
                    hr = pTriggerMonthly->put_DaysOfMonth(ttData.Type.MonthlyDate.rgfDays);    if (FAILED(hr)) goto finish;
                    hr = pTriggerMonthly->put_MonthsOfYear(ttData.Type.MonthlyDate.rgfMonths); if (FAILED(hr)) goto finish;
                    str.Format(L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerMonthly->put_RandomDelay(CComBSTR(str));                      if (FAILED(hr)) goto finish;
                    break;
                }

                case TASK_TIME_TRIGGER_MONTHLYDOW: {
                    CComPtr<IMonthlyDOWTrigger> pTriggerMonthlyDOW;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_MONTHLYDOW, &pTrigger);              if (FAILED(hr)) goto finish;
                    hr = pTrigger.QueryInterface(&pTriggerMonthlyDOW);                                if (FAILED(hr)) goto finish;
                    hr = pTriggerMonthlyDOW->put_WeeksOfMonth(
                        ttData.Type.MonthlyDOW.wWhichWeek == TASK_FIRST_WEEK  ? 0x01 :
                        ttData.Type.MonthlyDOW.wWhichWeek == TASK_SECOND_WEEK ? 0x02 :
                        ttData.Type.MonthlyDOW.wWhichWeek == TASK_THIRD_WEEK  ? 0x04 :
                        ttData.Type.MonthlyDOW.wWhichWeek == TASK_FOURTH_WEEK ? 0x08 :
                        ttData.Type.MonthlyDOW.wWhichWeek == TASK_LAST_WEEK   ? 0x10 : 0x00);         if (FAILED(hr)) goto finish;
                    hr = pTriggerMonthlyDOW->put_DaysOfWeek(ttData.Type.MonthlyDOW.rgfDaysOfTheWeek); if (FAILED(hr)) goto finish;
                    hr = pTriggerMonthlyDOW->put_MonthsOfYear(ttData.Type.MonthlyDOW.rgfMonths);      if (FAILED(hr)) goto finish;
                    str.Format(L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerMonthlyDOW->put_RandomDelay(CComBSTR(str));                          if (FAILED(hr)) goto finish;
                    break;
                }

                case TASK_EVENT_TRIGGER_ON_IDLE: {
                    hr = pTriggerCollection->Create(TASK_TRIGGER_IDLE, &pTrigger); if (FAILED(hr)) goto finish;
                    break;
                }

                case TASK_EVENT_TRIGGER_AT_SYSTEMSTART: {
                    CComPtr<IBootTrigger> pTriggerBoot;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_BOOT, &pTrigger); if (FAILED(hr)) goto finish;
                    hr = pTrigger.QueryInterface(&pTriggerBoot);                   if (FAILED(hr)) goto finish;
                    str.Format(L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerBoot->put_Delay(CComBSTR(str));                   if (FAILED(hr)) goto finish;
                    break;
                }

                case TASK_EVENT_TRIGGER_AT_LOGON: {
                    CComPtr<ILogonTrigger> pTriggerLogon;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_LOGON, &pTrigger); if (FAILED(hr)) goto finish;
                    hr = pTrigger.QueryInterface(&pTriggerLogon);                   if (FAILED(hr)) goto finish;
                    str.Format(L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerLogon->put_Delay(CComBSTR(str));                   if (FAILED(hr)) goto finish;
                    break;
                }
            }

            // Set trigger ID.
            str.Format(L"%u", iTrigger);
            hr = pTrigger->put_Id(CComBSTR(str));
            if (FAILED(hr)) goto finish;

            // Set trigger start date.
            str.Format(L"%04u-%02u-%02uT%02u:%02u:00", ttData.wBeginYear, ttData.wBeginMonth, ttData.wBeginDay, ttData.wStartHour, ttData.wStartMinute);
            hr = pTrigger->put_StartBoundary(CComBSTR(str));
            if (FAILED(hr)) goto finish;

            if (ttData.rgFlags & TASK_TRIGGER_FLAG_HAS_END_DATE) {
                // Set trigger end date.
                str.Format(L"%04u-%02u-%02u", ttData.wEndYear, ttData.wEndMonth, ttData.wEndDay);
                hr = pTrigger->put_EndBoundary(CComBSTR(str));
                if (FAILED(hr)) goto finish;
            }

            // Set trigger repetition duration and interval.
            if (ttData.MinutesDuration || ttData.MinutesInterval) {
                CComPtr<IRepetitionPattern> pRepetitionPattern;

                hr = pTrigger->get_Repetition(&pRepetitionPattern);
                if (FAILED(hr)) goto finish;
                str.Format(L"PT%uM", ttData.MinutesDuration);
                hr = pRepetitionPattern->put_Duration(CComBSTR(str));
                if (FAILED(hr)) goto finish;
                str.Format(L"PT%uM", ttData.MinutesInterval);
                hr = pRepetitionPattern->put_Interval(CComBSTR(str));
                if (FAILED(hr)) goto finish;
                hr = pRepetitionPattern->put_StopAtDurationEnd(ttData.rgFlags & TASK_TRIGGER_FLAG_KILL_AT_DURATION_END ? VARIANT_TRUE : VARIANT_FALSE);
                if (FAILED(hr)) goto finish;
            }

            // Enable/disable trigger.
            hr = pTrigger->put_Enabled(ttData.rgFlags & TASK_TRIGGER_FLAG_DISABLED ? VARIANT_FALSE : VARIANT_TRUE);
            if (FAILED(hr)) goto finish;
        }

        // Get the task folder.
        hr = pService->GetFolder(CComBSTR(L"\\"), &pTaskFolder);
        if (FAILED(hr)) goto finish;

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
    } else {
        // Windows XP or older.
        CComPtr<ITaskScheduler> pTaskScheduler;
        CComPtr<ITask> pTask;

        // Get task scheduler object.
        hr = pTaskScheduler.CoCreateInstance(CLSID_CTaskScheduler, NULL, CLSCTX_ALL);
        if (FAILED(hr)) goto finish;

        // Create the new task.
        hr = pTaskScheduler->NewWorkItem(m_sValue, CLSID_CTask, IID_ITask, (IUnknown**)&pTask);
        if (pSession->m_bRollbackEnabled) {
            // Order rollback action to delete the task. ITask::NewWorkItem() might made a blank task already.
            pSession->m_olRollback.AddHead(new CMSITSCAOpTaskDelete(m_sValue));
        }
        if (FAILED(hr)) goto finish;

        // Set its properties.
        hr = pTask->SetApplicationName (m_sApplicationName ); if (FAILED(hr)) goto finish;
        hr = pTask->SetParameters      (m_sParameters      ); if (FAILED(hr)) goto finish;
        hr = pTask->SetWorkingDirectory(m_sWorkingDirectory); if (FAILED(hr)) goto finish;
        hr = pTask->SetComment         (m_sComment         ); if (FAILED(hr)) goto finish;
        if (pSession->m_bRollbackEnabled && (m_dwFlags & TASK_FLAG_DISABLED) == 0) {
            // The task is supposed to be enabled.
            // However, we shall enable it at commit to prevent it from accidentally starting in the very installation phase.
            pSession->m_olCommit.AddTail(new CMSITSCAOpTaskEnable(m_sValue, TRUE));
            m_dwFlags |= TASK_FLAG_DISABLED;
        }
        hr = pTask->SetFlags           (m_dwFlags          ); if (FAILED(hr)) goto finish;
        hr = pTask->SetPriority        (m_dwPriority       ); if (FAILED(hr)) goto finish;

        // Set task credentials.
        hr = m_sAccountName.IsEmpty() ?
            pTask->SetAccountInformation(L"",            NULL       ) :
            pTask->SetAccountInformation(m_sAccountName, m_sPassword);
        if (FAILED(hr)) goto finish;

        hr = pTask->SetIdleWait  (m_wIdleMinutes, m_wDeadlineMinutes); if (FAILED(hr)) goto finish;
        hr = pTask->SetMaxRunTime(m_dwMaxRuntimeMS                  ); if (FAILED(hr)) goto finish;

        // Add triggers.
        for (pos = m_lTriggers.GetHeadPosition(); pos;) {
            WORD wTriggerIdx;
            CComPtr<ITaskTrigger> pTrigger;
            TASK_TRIGGER ttData = m_lTriggers.GetNext(pos); // Don't use reference! We don't want to modify original trigger data by adding random startup delay.

            hr = pTask->CreateTrigger(&wTriggerIdx, &pTrigger);
            if (FAILED(hr)) goto finish;

            if (ttData.wRandomMinutesInterval) {
                // Windows XP doesn't support random startup delay. However, we can add fixed "random" delay when creating the trigger.
                WORD wStartTime = ttData.wStartHour * 60 + ttData.wStartMinute + (WORD)::MulDiv(rand(), ttData.wRandomMinutesInterval, RAND_MAX);
                FILETIME ftValue;
                SYSTEMTIME stValue;
                ULONGLONG ullValue;

                // Convert MDY date to numerical date (SYSTEMTIME -> FILETIME -> ULONGLONG).
                memset(&stValue, 0, sizeof(SYSTEMTIME));
                stValue.wYear  = ttData.wBeginYear;
                stValue.wMonth = ttData.wBeginMonth;
                stValue.wDay   = ttData.wBeginDay;
                verify(::SystemTimeToFileTime(&stValue, &ftValue));
                ullValue = ((ULONGLONG)(ftValue.dwHighDateTime) << 32) | ftValue.dwLowDateTime;

                // Wrap days.
                while (wStartTime >= 1440) {
                    ullValue   += (ULONGLONG)864000000000;
                    wStartTime -= 1440;
                }

                // Convert numerical date to DMY (ULONGLONG -> FILETIME -> SYSTEMTIME).
                ftValue.dwHighDateTime = ullValue >> 32;
                ftValue.dwLowDateTime  = ullValue & 0xffffffff;
                verify(::FileTimeToSystemTime(&ftValue, &stValue));

                // Set new trigger date and time.
                ttData.wBeginYear   = stValue.wYear;
                ttData.wBeginMonth  = stValue.wMonth;
                ttData.wBeginDay    = stValue.wDay;
                ttData.wStartHour   = wStartTime / 60;
                ttData.wStartMinute = wStartTime % 60;
            }

            hr = pTrigger->SetTrigger(&ttData);
            if (FAILED(hr)) goto finish;
        }

        // Save the task.
        CComQIPtr<IPersistFile> pTaskFile(pTask);
        if (!pTaskFile) { hr = E_NOINTERFACE; goto finish; }
        hr = pTaskFile->Save(NULL, TRUE);
    }

finish:
    if (FAILED(hr)) {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        verify(::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_TASK_CREATE_FAILED) == ERROR_SUCCESS);
        verify(::MsiRecordSetStringW(hRecordProg, 2, m_sValue                        ) == ERROR_SUCCESS);
        verify(::MsiRecordSetInteger(hRecordProg, 3, hr                              ) == ERROR_SUCCESS);
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
    }
    return hr;
}


UINT CMSITSCAOpTaskCreate::SetFromRecord(MSIHANDLE hInstall, MSIHANDLE hRecord)
{
    UINT uiResult;
    int iValue;
    CStringW sFolder;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord,  3, m_sApplicationName);
    if (uiResult != ERROR_SUCCESS) return uiResult;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord,  4, m_sParameters);
    if (uiResult != ERROR_SUCCESS) return uiResult;

    uiResult = ::MsiRecordGetStringW(hRecord, 5, sFolder);
    if (uiResult != ERROR_SUCCESS) return uiResult;
    uiResult = ::MsiGetTargetPathW(hInstall, sFolder, m_sWorkingDirectory);
    if (uiResult != ERROR_SUCCESS) return uiResult;
    if (!m_sWorkingDirectory.IsEmpty() && m_sWorkingDirectory.GetAt(m_sWorkingDirectory.GetLength() - 1) == L'\\') {
        // Trim trailing backslash.
        m_sWorkingDirectory.Truncate(m_sWorkingDirectory.GetLength() - 1);
    }

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord, 10, m_sAuthor);
    if (uiResult != ERROR_SUCCESS) return uiResult;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord, 11, m_sComment);
    if (uiResult != ERROR_SUCCESS) return uiResult;

    m_dwFlags = ::MsiRecordGetInteger(hRecord,  6);
    if (m_dwFlags == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;

    m_dwPriority = ::MsiRecordGetInteger(hRecord,  7);
    if (m_dwPriority == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord,  8, m_sAccountName);
    if (uiResult != ERROR_SUCCESS) return uiResult;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord,  9, m_sPassword);
    if (uiResult != ERROR_SUCCESS) return uiResult;

    iValue = ::MsiRecordGetInteger(hRecord, 12);
    m_wIdleMinutes = iValue != MSI_NULL_INTEGER ? (WORD)iValue : 0;

    iValue = ::MsiRecordGetInteger(hRecord, 13);
    m_wDeadlineMinutes = iValue != MSI_NULL_INTEGER ? (WORD)iValue : 0;

    m_dwMaxRuntimeMS = ::MsiRecordGetInteger(hRecord, 14);
    if (m_dwMaxRuntimeMS == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;

    return ERROR_SUCCESS;
}


UINT CMSITSCAOpTaskCreate::SetTriggersFromView(MSIHANDLE hView)
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

    return ERROR_SUCCESS;
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpTaskDelete
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpTaskDelete::CMSITSCAOpTaskDelete(LPCWSTR pszTaskName, int iTicks) :
    CMSITSCAOpTypeSingleString(pszTaskName, iTicks)
{
}


HRESULT CMSITSCAOpTaskDelete::Execute(CMSITSCASession *pSession)
{
    assert(pSession);
    HRESULT hr;
    CComPtr<ITaskService> pService;

    hr = pService.CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER);
    if (SUCCEEDED(hr)) {
        // Windows Vista or newer.
        CComVariant vEmpty;
        CComPtr<ITaskFolder> pTaskFolder;

        // Connect to local task service.
        hr = pService->Connect(vEmpty, vEmpty, vEmpty, vEmpty);
        if (FAILED(hr)) goto finish;

        // Get task folder.
        hr = pService->GetFolder(CComBSTR(L"\\"), &pTaskFolder);
        if (FAILED(hr)) goto finish;

        if (pSession->m_bRollbackEnabled) {
            CComPtr<IRegisteredTask> pTask, pTaskOrig;
            CComPtr<ITaskDefinition> pTaskDefinition;
            CComPtr<IPrincipal> pPrincipal;
            VARIANT_BOOL bEnabled;
            TASK_LOGON_TYPE logonType;
            CComBSTR sSSDL;
            CComVariant vSSDL;
            CStringW sDisplayNameOrig;
            UINT uiCount = 0;

            // Get the source task.
            hr = pTaskFolder->GetTask(CComBSTR(m_sValue), &pTask);
            if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) { hr = S_OK; goto finish; }
            else if (FAILED(hr)) goto finish;

            // Disable the task.
            hr = pTask->get_Enabled(&bEnabled);
            if (FAILED(hr)) goto finish;
            if (bEnabled) {
                // The task is enabled.

                // In case the task disabling fails halfway, try to re-enable it anyway on rollback.
                pSession->m_olRollback.AddHead(new CMSITSCAOpTaskEnable(m_sValue, TRUE));

                // Disable it.
                hr = pTask->put_Enabled(VARIANT_FALSE);
                if (FAILED(hr)) goto finish;
            }

            // Get task's definition.
            hr = pTask->get_Definition(&pTaskDefinition);
            if (FAILED(hr)) goto finish;

            // Get task principal.
            hr = pTaskDefinition->get_Principal(&pPrincipal);
            if (FAILED(hr)) goto finish;

            // Get task logon type.
            hr = pPrincipal->get_LogonType(&logonType);
            if (FAILED(hr)) goto finish;

            // Get task security descriptor.
            hr = pTask->GetSecurityDescriptor(DACL_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION | OWNER_SECURITY_INFORMATION | SACL_SECURITY_INFORMATION | LABEL_SECURITY_INFORMATION, &sSSDL);
            if (hr == HRESULT_FROM_WIN32(ERROR_PRIVILEGE_NOT_HELD)) vSSDL.Clear();
            else if (FAILED(hr)) goto finish;
            else {
                V_VT  (&vSSDL) = VT_BSTR;
                V_BSTR(&vSSDL) = sSSDL.Detach();
            }

            // Register a backup copy of task.
            do {
                sDisplayNameOrig.Format(L"%ls (orig %u)", (LPCWSTR)m_sValue, ++uiCount);
                hr = pTaskFolder->RegisterTaskDefinition(
                    CComBSTR(sDisplayNameOrig), // path
                    pTaskDefinition,            // pDefinition
                    TASK_CREATE,                // flags
                    vEmpty,                     // userId
                    vEmpty,                     // password
                    logonType,                  // logonType
                    vSSDL,                      // sddl
                    &pTaskOrig);                // ppTask
            } while (hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS));
            // In case the backup copy creation failed halfway, try to delete its remains anyway on rollback.
            pSession->m_olRollback.AddHead(new CMSITSCAOpTaskDelete(sDisplayNameOrig));
            if (FAILED(hr)) goto finish;

            // Order rollback action to restore from backup copy.
            pSession->m_olRollback.AddHead(new CMSITSCAOpTaskCopy(sDisplayNameOrig, m_sValue));

            // Delete it.
            hr = pTaskFolder->DeleteTask(CComBSTR(m_sValue), 0);
            if (FAILED(hr)) goto finish;

            // Order commit action to delete backup copy.
            pSession->m_olCommit.AddTail(new CMSITSCAOpTaskDelete(sDisplayNameOrig));
        } else {
            // Delete the task.
            hr = pTaskFolder->DeleteTask(CComBSTR(m_sValue), 0);
            if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) hr = S_OK;
        }
    } else {
        // Windows XP or older.
        CComPtr<ITaskScheduler> pTaskScheduler;

        // Get task scheduler object.
        hr = pTaskScheduler.CoCreateInstance(CLSID_CTaskScheduler, NULL, CLSCTX_ALL);
        if (FAILED(hr)) goto finish;

        if (pSession->m_bRollbackEnabled) {
            CComPtr<ITask> pTask;
            DWORD dwFlags;
            CStringW sDisplayNameOrig;
            UINT uiCount = 0;

            // Load the task.
            hr = pTaskScheduler->Activate(m_sValue, IID_ITask, (IUnknown**)&pTask);
            if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) { hr = S_OK; goto finish; }
            else if (FAILED(hr)) goto finish;

            // Disable the task.
            hr = pTask->GetFlags(&dwFlags);
            if (FAILED(hr)) goto finish;
            if ((dwFlags & TASK_FLAG_DISABLED) == 0) {
                // The task is enabled.

                // In case the task disabling fails halfway, try to re-enable it anyway on rollback.
                pSession->m_olRollback.AddHead(new CMSITSCAOpTaskEnable(m_sValue, TRUE));

                // Disable it.
                dwFlags |= TASK_FLAG_DISABLED;
                hr = pTask->SetFlags(dwFlags);
                if (FAILED(hr)) goto finish;
            }

            // Prepare a backup copy of task.
            do {
                sDisplayNameOrig.Format(L"%ls (orig %u)", (LPCWSTR)m_sValue, ++uiCount);
                hr = pTaskScheduler->AddWorkItem(sDisplayNameOrig, pTask);
            } while (hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS));
            // In case the backup copy creation failed halfway, try to delete its remains anyway on rollback.
            pSession->m_olRollback.AddHead(new CMSITSCAOpTaskDelete(sDisplayNameOrig));
            if (FAILED(hr)) goto finish;

            // Save the backup copy.
            CComQIPtr<IPersistFile> pTaskFile(pTask);
            if (!pTaskFile) { hr = E_NOINTERFACE; goto finish; }
            hr = pTaskFile->Save(NULL, TRUE);
            if (FAILED(hr)) goto finish;

            // Order rollback action to restore from backup copy.
            pSession->m_olRollback.AddHead(new CMSITSCAOpTaskCopy(sDisplayNameOrig, m_sValue));

            // Delete it.
            hr = pTaskScheduler->Delete(m_sValue);
            if (FAILED(hr)) goto finish;

            // Order commit action to delete backup copy.
            pSession->m_olCommit.AddTail(new CMSITSCAOpTaskDelete(sDisplayNameOrig));
        } else {
            // Delete the task.
            hr = pTaskScheduler->Delete(m_sValue);
            if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) hr = S_OK;
        }
    }

finish:
    if (FAILED(hr)) {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        verify(::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_TASK_DELETE_FAILED) == ERROR_SUCCESS);
        verify(::MsiRecordSetStringW(hRecordProg, 2, m_sValue                        ) == ERROR_SUCCESS);
        verify(::MsiRecordSetInteger(hRecordProg, 3, hr                              ) == ERROR_SUCCESS);
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
    }
    return hr;
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpTaskEnable
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpTaskEnable::CMSITSCAOpTaskEnable(LPCWSTR pszTaskName, BOOL bEnable, int iTicks) :
    m_bEnable(bEnable),
    CMSITSCAOpTypeSingleString(pszTaskName, iTicks)
{
}


HRESULT CMSITSCAOpTaskEnable::Execute(CMSITSCASession *pSession)
{
    assert(pSession);
    HRESULT hr;
    CComPtr<ITaskService> pService;

    hr = pService.CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER);
    if (SUCCEEDED(hr)) {
        // Windows Vista or newer.
        CComVariant vEmpty;
        CComPtr<ITaskFolder> pTaskFolder;
        CComPtr<IRegisteredTask> pTask;
        VARIANT_BOOL bEnabled;

        // Connect to local task service.
        hr = pService->Connect(vEmpty, vEmpty, vEmpty, vEmpty);
        if (FAILED(hr)) goto finish;

        // Get task folder.
        hr = pService->GetFolder(CComBSTR(L"\\"), &pTaskFolder);
        if (FAILED(hr)) goto finish;

        // Get task.
        hr = pTaskFolder->GetTask(CComBSTR(m_sValue), &pTask);
        if (FAILED(hr)) goto finish;

        // Get currently enabled state.
        hr = pTask->get_Enabled(&bEnabled);
        if (FAILED(hr)) goto finish;

        // Modify enable state.
        if (m_bEnable) {
            if (pSession->m_bRollbackEnabled && !bEnabled) {
                // The task is disabled and supposed to be enabled.
                // However, we shall enable it at commit to prevent it from accidentally starting in the very installation phase.
                pSession->m_olCommit.AddTail(new CMSITSCAOpTaskEnable(m_sValue, TRUE));
                hr = S_OK; goto finish;
            } else
                bEnabled = VARIANT_TRUE;
        } else {
            if (pSession->m_bRollbackEnabled && bEnabled) {
                // The task is enabled and we will disable it now.
                // Order rollback to re-enable it.
                pSession->m_olRollback.AddHead(new CMSITSCAOpTaskEnable(m_sValue, TRUE));
            }
            bEnabled = VARIANT_FALSE;
        }

        // Set enable/disable.
        hr = pTask->put_Enabled(bEnabled);
        if (FAILED(hr)) goto finish;
    } else {
        // Windows XP or older.
        CComPtr<ITaskScheduler> pTaskScheduler;
        CComPtr<ITask> pTask;
        DWORD dwFlags;

        // Get task scheduler object.
        hr = pTaskScheduler.CoCreateInstance(CLSID_CTaskScheduler, NULL, CLSCTX_ALL);
        if (FAILED(hr)) goto finish;

        // Load the task.
        hr = pTaskScheduler->Activate(m_sValue, IID_ITask, (IUnknown**)&pTask);
        if (FAILED(hr)) goto finish;

        // Get task's current flags.
        hr = pTask->GetFlags(&dwFlags);
        if (FAILED(hr)) goto finish;

        // Modify flags.
        if (m_bEnable) {
            if (pSession->m_bRollbackEnabled && (dwFlags & TASK_FLAG_DISABLED)) {
                // The task is disabled and supposed to be enabled.
                // However, we shall enable it at commit to prevent it from accidentally starting in the very installation phase.
                pSession->m_olCommit.AddTail(new CMSITSCAOpTaskEnable(m_sValue, TRUE));
                hr = S_OK; goto finish;
            } else
                dwFlags &= ~TASK_FLAG_DISABLED;
        } else {
            if (pSession->m_bRollbackEnabled && !(dwFlags & TASK_FLAG_DISABLED)) {
                // The task is enabled and we will disable it now.
                // Order rollback to re-enable it.
                pSession->m_olRollback.AddHead(new CMSITSCAOpTaskEnable(m_sValue, TRUE));
            }
            dwFlags |=  TASK_FLAG_DISABLED;
        }

        // Set flags.
        hr = pTask->SetFlags(dwFlags);
        if (FAILED(hr)) goto finish;

        // Save the task.
        CComQIPtr<IPersistFile> pTaskFile(pTask);
        if (!pTaskFile) { hr = E_NOINTERFACE; goto finish; }
        hr = pTaskFile->Save(NULL, TRUE);
        if (FAILED(hr)) goto finish;
    }

finish:
    if (FAILED(hr)) {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        verify(::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_TASK_ENABLE_FAILED) == ERROR_SUCCESS);
        verify(::MsiRecordSetStringW(hRecordProg, 2, m_sValue                        ) == ERROR_SUCCESS);
        verify(::MsiRecordSetInteger(hRecordProg, 3, hr                              ) == ERROR_SUCCESS);
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
    }
    return hr;
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpTaskCopy
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpTaskCopy::CMSITSCAOpTaskCopy(LPCWSTR pszTaskSrc, LPCWSTR pszTaskDst, int iTicks) :
    CMSITSCAOpTypeSrcDstString(pszTaskSrc, pszTaskDst, iTicks)
{
}


HRESULT CMSITSCAOpTaskCopy::Execute(CMSITSCASession *pSession)
{
    assert(pSession);
    HRESULT hr;
    CComPtr<ITaskService> pService;

    hr = pService.CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER);
    if (SUCCEEDED(hr)) {
        // Windows Vista or newer.
        CComVariant vEmpty;
        CComPtr<ITaskFolder> pTaskFolder;
        CComPtr<IRegisteredTask> pTask, pTaskOrig;
        CComPtr<ITaskDefinition> pTaskDefinition;
        CComPtr<IPrincipal> pPrincipal;
        TASK_LOGON_TYPE logonType;
        CComBSTR sSSDL;

        // Connect to local task service.
        hr = pService->Connect(vEmpty, vEmpty, vEmpty, vEmpty);
        if (FAILED(hr)) goto finish;

        // Get task folder.
        hr = pService->GetFolder(CComBSTR(L"\\"), &pTaskFolder);
        if (FAILED(hr)) goto finish;

        // Get the source task.
        hr = pTaskFolder->GetTask(CComBSTR(m_sValue1), &pTask);
        if (FAILED(hr)) goto finish;

        // Get task's definition.
        hr = pTask->get_Definition(&pTaskDefinition);
        if (FAILED(hr)) goto finish;

        // Get task principal.
        hr = pTaskDefinition->get_Principal(&pPrincipal);
        if (FAILED(hr)) goto finish;

        // Get task logon type.
        hr = pPrincipal->get_LogonType(&logonType);
        if (FAILED(hr)) goto finish;

        // Get task security descriptor.
        hr = pTask->GetSecurityDescriptor(DACL_SECURITY_INFORMATION | GROUP_SECURITY_INFORMATION | OWNER_SECURITY_INFORMATION | SACL_SECURITY_INFORMATION | LABEL_SECURITY_INFORMATION, &sSSDL);
        if (FAILED(hr)) goto finish;

        // Register a new task.
        hr = pTaskFolder->RegisterTaskDefinition(
            CComBSTR(m_sValue2), // path
            pTaskDefinition,     // pDefinition
            TASK_CREATE,         // flags
            vEmpty,              // userId
            vEmpty,              // password
            logonType,           // logonType
            CComVariant(sSSDL),  // sddl
            &pTaskOrig);         // ppTask
        if (FAILED(hr)) goto finish;
    } else {
        // Windows XP or older.
        CComPtr<ITaskScheduler> pTaskScheduler;
        CComPtr<ITask> pTask;

        // Get task scheduler object.
        hr = pTaskScheduler.CoCreateInstance(CLSID_CTaskScheduler, NULL, CLSCTX_ALL);
        if (FAILED(hr)) goto finish;

        // Load the source task.
        hr = pTaskScheduler->Activate(m_sValue1, IID_ITask, (IUnknown**)&pTask);
        if (FAILED(hr)) goto finish;

        // Add with different name.
        hr = pTaskScheduler->AddWorkItem(m_sValue2, pTask);
        if (pSession->m_bRollbackEnabled) {
            // Order rollback action to delete the new copy. ITask::AddWorkItem() might made a blank task already.
            pSession->m_olRollback.AddHead(new CMSITSCAOpTaskDelete(m_sValue2));
        }
        if (FAILED(hr)) goto finish;

        // Save the task.
        CComQIPtr<IPersistFile> pTaskFile(pTask);
        if (!pTaskFile) { hr = E_NOINTERFACE; goto finish; }
        hr = pTaskFile->Save(NULL, TRUE);
        if (FAILED(hr)) goto finish;
    }

finish:
    if (FAILED(hr)) {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(4);
        verify(::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_TASK_COPY_FAILED) == ERROR_SUCCESS);
        verify(::MsiRecordSetStringW(hRecordProg, 2, m_sValue1                     ) == ERROR_SUCCESS);
        verify(::MsiRecordSetStringW(hRecordProg, 3, m_sValue2                     ) == ERROR_SUCCESS);
        verify(::MsiRecordSetInteger(hRecordProg, 4, hr                            ) == ERROR_SUCCESS);
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
    }
    return hr;
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
        if (!pSession->m_bContinueOnError && FAILED(hr)) {
            // Operation failed. Its Execute() method should have sent error message to Installer.
            // Therefore, just quit here.
            return hr;
        }

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
