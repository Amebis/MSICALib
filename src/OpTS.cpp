/*
    Copyright 1991-2018 Amebis

    This file is part of MSICA.

    MSICA is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    MSICA is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with MSICA. If not, see <http://www.gnu.org/licenses/>.
*/

#include "stdafx.h"

#pragma comment(lib, "mstask.lib")
#pragma comment(lib, "taskschd.lib")


namespace MSICA {

////////////////////////////////////////////////////////////////////////////
// COpTaskCreate
////////////////////////////////////////////////////////////////////////////

COpTaskCreate::COpTaskCreate(LPCWSTR pszTaskName, int iTicks) :
    m_dwFlags(0),
    m_dwPriority(NORMAL_PRIORITY_CLASS),
    m_wIdleMinutes(0),
    m_wDeadlineMinutes(0),
    m_dwMaxRuntimeMS(INFINITE),
    COpTypeSingleString(pszTaskName, iTicks)
{
}


HRESULT COpTaskCreate::Execute(CSession *pSession)
{
    HRESULT hr;
    PMSIHANDLE hRecordMsg = ::MsiCreateRecord(1);
    winstd::com_obj<ITaskService> pService;

    // Display our custom message in the progress bar.
    ::MsiRecordSetStringW(hRecordMsg, 1, m_sValue.c_str());
    if (MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ACTIONDATA, hRecordMsg) == IDCANCEL)
        return HRESULT_FROM_WIN32(ERROR_INSTALL_USEREXIT);

    {
        // Delete existing task first.
        // Since task deleting is a complicated job (when rollback/commit support is required), and we do have an operation just for that, we use it.
        // Don't worry, COpTaskDelete::Execute() returns S_OK if task doesn't exist.
        COpTaskDelete opDelete(m_sValue.c_str());
        hr = opDelete.Execute(pSession);
        if (FAILED(hr)) goto finish;
    }

    hr = pService.create(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER);
    if (SUCCEEDED(hr)) {
        // Windows Vista or newer.
        winstd::variant vEmpty;
        winstd::com_obj<ITaskDefinition> pTaskDefinition;
        winstd::com_obj<ITaskSettings> pTaskSettings;
        winstd::com_obj<IPrincipal> pPrincipal;
        winstd::com_obj<IActionCollection> pActionCollection;
        winstd::com_obj<IAction> pAction;
        winstd::com_obj<IIdleSettings> pIdleSettings;
        winstd::com_obj<IExecAction> pExecAction;
        winstd::com_obj<IRegistrationInfo> pRegististrationInfo;
        winstd::com_obj<ITriggerCollection> pTriggerCollection;
        winstd::com_obj<ITaskFolder> pTaskFolder;
        winstd::com_obj<IRegisteredTask> pTask;
        std::wstring str;
        UINT iTrigger;
        TASK_LOGON_TYPE logonType;
        winstd::bstr bstrContext(L"Author");

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
            pSession->m_olCommit.push_back(new COpTaskEnable(m_sValue.c_str(), TRUE));
            m_dwFlags |= TASK_FLAG_DISABLED;
        }
        hr = pTaskSettings->put_Enabled(m_dwFlags & TASK_FLAG_DISABLED ? VARIANT_FALSE : VARIANT_TRUE); if (FAILED(hr)) goto finish;

        // Get task actions.
        hr = pTaskDefinition->get_Actions(&pActionCollection);
        if (FAILED(hr)) goto finish;

        // Add execute action.
        hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
        if (FAILED(hr)) goto finish;
        hr = pAction.query_interface(&pExecAction);
        if (FAILED(hr)) goto finish;

        // Configure the action.
        hr = pExecAction->put_Path            (winstd::bstr(m_sApplicationName )); if (FAILED(hr)) goto finish;
        hr = pExecAction->put_Arguments       (winstd::bstr(m_sParameters      )); if (FAILED(hr)) goto finish;
        hr = pExecAction->put_WorkingDirectory(winstd::bstr(m_sWorkingDirectory)); if (FAILED(hr)) goto finish;

        // Set task description.
        hr = pTaskDefinition->get_RegistrationInfo(&pRegististrationInfo);    if (FAILED(hr)) goto finish;
        hr = pRegististrationInfo->put_Author(winstd::bstr(m_sAuthor));       if (FAILED(hr)) goto finish;
        hr = pRegististrationInfo->put_Description(winstd::bstr(m_sComment)); if (FAILED(hr)) goto finish;

        // Configure task "flags".
        if (m_dwFlags & TASK_FLAG_DELETE_WHEN_DONE) {
            hr = pTaskSettings->put_DeleteExpiredTaskAfter(winstd::bstr(L"PT24H"));
            if (FAILED(hr)) goto finish;
        }
        hr = pTaskSettings->put_Hidden(m_dwFlags & TASK_FLAG_HIDDEN ? VARIANT_TRUE : VARIANT_FALSE);                                         if (FAILED(hr)) goto finish;
        hr = pTaskSettings->put_WakeToRun(m_dwFlags & TASK_FLAG_SYSTEM_REQUIRED ? VARIANT_TRUE : VARIANT_FALSE);                             if (FAILED(hr)) goto finish;
        hr = pTaskSettings->put_DisallowStartIfOnBatteries(m_dwFlags & TASK_FLAG_DONT_START_IF_ON_BATTERIES ? VARIANT_TRUE : VARIANT_FALSE); if (FAILED(hr)) goto finish;
        hr = pTaskSettings->put_StopIfGoingOnBatteries(m_dwFlags & TASK_FLAG_KILL_IF_GOING_ON_BATTERIES ? VARIANT_TRUE : VARIANT_FALSE);     if (FAILED(hr)) goto finish;

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

        if (m_sAccountName.empty()) {
            logonType = TASK_LOGON_SERVICE_ACCOUNT;
            hr = pPrincipal->put_LogonType(logonType);                 if (FAILED(hr)) goto finish;
            hr = pPrincipal->put_UserId(winstd::bstr(L"S-1-5-18"));    if (FAILED(hr)) goto finish;
            hr = pPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);      if (FAILED(hr)) goto finish;
        } else if (m_dwFlags & TASK_FLAG_RUN_ONLY_IF_LOGGED_ON) {
            logonType = TASK_LOGON_INTERACTIVE_TOKEN;
            hr = pPrincipal->put_LogonType(logonType);                 if (FAILED(hr)) goto finish;
            hr = pPrincipal->put_RunLevel(TASK_RUNLEVEL_LUA);          if (FAILED(hr)) goto finish;
        } else {
            logonType = TASK_LOGON_PASSWORD;
            hr = pPrincipal->put_LogonType(logonType);                 if (FAILED(hr)) goto finish;
            hr = pPrincipal->put_UserId(winstd::bstr(m_sAccountName)); if (FAILED(hr)) goto finish;
            hr = pPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);      if (FAILED(hr)) goto finish;
        }

        // Connect principal and action collection.
        hr = pPrincipal->put_Id(bstrContext);             if (FAILED(hr)) goto finish;
        hr = pActionCollection->put_Context(bstrContext); if (FAILED(hr)) goto finish;

        // Configure idle settings.
        hr = pTaskSettings->put_RunOnlyIfIdle(m_dwFlags & TASK_FLAG_START_ONLY_IF_IDLE ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) goto finish;
        hr = pTaskSettings->get_IdleSettings(&pIdleSettings);
        if (FAILED(hr)) goto finish;
        sprintf(str, L"PT%uS", m_wIdleMinutes*60);
        hr = pIdleSettings->put_IdleDuration(winstd::bstr(str));
        if (FAILED(hr)) goto finish;
        sprintf(str, L"PT%uS", m_wDeadlineMinutes*60);
        hr = pIdleSettings->put_WaitTimeout(winstd::bstr(str));
        if (FAILED(hr)) goto finish;
        hr = pIdleSettings->put_RestartOnIdle(m_dwFlags & TASK_FLAG_RESTART_ON_IDLE_RESUME ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) goto finish;
        hr = pIdleSettings->put_StopOnIdleEnd(m_dwFlags & TASK_FLAG_KILL_ON_IDLE_END ? VARIANT_TRUE : VARIANT_FALSE);
        if (FAILED(hr)) goto finish;

        // Configure task runtime limit.
        sprintf(str, L"PT%uS", m_dwMaxRuntimeMS != INFINITE ? (m_dwMaxRuntimeMS + 500) / 1000 : 0);
        hr = pTaskSettings->put_ExecutionTimeLimit(winstd::bstr(str));
        if (FAILED(hr)) goto finish;

        // Get task trigger colection.
        hr = pTaskDefinition->get_Triggers(&pTriggerCollection);
        if (FAILED(hr)) goto finish;

        // Add triggers.
        iTrigger = 0;
        for (auto t = m_lTriggers.cbegin(), t_end = m_lTriggers.cend(); t != t_end; ++t, iTrigger++) {
            winstd::com_obj<ITrigger> pTrigger;
            const TASK_TRIGGER &ttData = *t;

            switch (ttData.TriggerType) {
                case TASK_TIME_TRIGGER_ONCE: {
                    winstd::com_obj<ITimeTrigger> pTriggerTime;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_TIME, &pTrigger); if (FAILED(hr)) goto finish;
                    hr = pTrigger.query_interface(&pTriggerTime);                  if (FAILED(hr)) goto finish;
                    sprintf(str, L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerTime->put_RandomDelay(winstd::bstr(str));         if (FAILED(hr)) goto finish;
                    break;
                }

                case TASK_TIME_TRIGGER_DAILY: {
                    winstd::com_obj<IDailyTrigger> pTriggerDaily;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_DAILY, &pTrigger);       if (FAILED(hr)) goto finish;
                    hr = pTrigger.query_interface(&pTriggerDaily);                        if (FAILED(hr)) goto finish;
                    hr = pTriggerDaily->put_DaysInterval(ttData.Type.Daily.DaysInterval); if (FAILED(hr)) goto finish;
                    sprintf(str, L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerDaily->put_RandomDelay(winstd::bstr(str));               if (FAILED(hr)) goto finish;
                    break;
                }

                case TASK_TIME_TRIGGER_WEEKLY: {
                    winstd::com_obj<IWeeklyTrigger> pTriggerWeekly;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_WEEKLY, &pTrigger);          if (FAILED(hr)) goto finish;
                    hr = pTrigger.query_interface(&pTriggerWeekly);                           if (FAILED(hr)) goto finish;
                    hr = pTriggerWeekly->put_WeeksInterval(ttData.Type.Weekly.WeeksInterval); if (FAILED(hr)) goto finish;
                    hr = pTriggerWeekly->put_DaysOfWeek(ttData.Type.Weekly.rgfDaysOfTheWeek); if (FAILED(hr)) goto finish;
                    sprintf(str, L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerWeekly->put_RandomDelay(winstd::bstr(str));                  if (FAILED(hr)) goto finish;
                    break;
                }

                case TASK_TIME_TRIGGER_MONTHLYDATE: {
                    winstd::com_obj<IMonthlyTrigger> pTriggerMonthly;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_MONTHLY, &pTrigger);          if (FAILED(hr)) goto finish;
                    hr = pTrigger.query_interface(&pTriggerMonthly);                           if (FAILED(hr)) goto finish;
                    hr = pTriggerMonthly->put_DaysOfMonth(ttData.Type.MonthlyDate.rgfDays);    if (FAILED(hr)) goto finish;
                    hr = pTriggerMonthly->put_MonthsOfYear(ttData.Type.MonthlyDate.rgfMonths); if (FAILED(hr)) goto finish;
                    sprintf(str, L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerMonthly->put_RandomDelay(winstd::bstr(str));                  if (FAILED(hr)) goto finish;
                    break;
                }

                case TASK_TIME_TRIGGER_MONTHLYDOW: {
                    winstd::com_obj<IMonthlyDOWTrigger> pTriggerMonthlyDOW;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_MONTHLYDOW, &pTrigger);              if (FAILED(hr)) goto finish;
                    hr = pTrigger.query_interface(&pTriggerMonthlyDOW);                               if (FAILED(hr)) goto finish;
                    hr = pTriggerMonthlyDOW->put_WeeksOfMonth(
                        ttData.Type.MonthlyDOW.wWhichWeek == TASK_FIRST_WEEK  ? 0x01 :
                        ttData.Type.MonthlyDOW.wWhichWeek == TASK_SECOND_WEEK ? 0x02 :
                        ttData.Type.MonthlyDOW.wWhichWeek == TASK_THIRD_WEEK  ? 0x04 :
                        ttData.Type.MonthlyDOW.wWhichWeek == TASK_FOURTH_WEEK ? 0x08 :
                        ttData.Type.MonthlyDOW.wWhichWeek == TASK_LAST_WEEK   ? 0x10 : 0x00);         if (FAILED(hr)) goto finish;
                    hr = pTriggerMonthlyDOW->put_DaysOfWeek(ttData.Type.MonthlyDOW.rgfDaysOfTheWeek); if (FAILED(hr)) goto finish;
                    hr = pTriggerMonthlyDOW->put_MonthsOfYear(ttData.Type.MonthlyDOW.rgfMonths);      if (FAILED(hr)) goto finish;
                    sprintf(str, L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerMonthlyDOW->put_RandomDelay(winstd::bstr(str));                      if (FAILED(hr)) goto finish;
                    break;
                }

                case TASK_EVENT_TRIGGER_ON_IDLE: {
                    hr = pTriggerCollection->Create(TASK_TRIGGER_IDLE, &pTrigger); if (FAILED(hr)) goto finish;
                    break;
                }

                case TASK_EVENT_TRIGGER_AT_SYSTEMSTART: {
                    winstd::com_obj<IBootTrigger> pTriggerBoot;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_BOOT, &pTrigger); if (FAILED(hr)) goto finish;
                    hr = pTrigger.query_interface(&pTriggerBoot);                  if (FAILED(hr)) goto finish;
                    sprintf(str, L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerBoot->put_Delay(winstd::bstr(str));               if (FAILED(hr)) goto finish;
                    break;
                }

                case TASK_EVENT_TRIGGER_AT_LOGON: {
                    winstd::com_obj<ILogonTrigger> pTriggerLogon;
                    hr = pTriggerCollection->Create(TASK_TRIGGER_LOGON, &pTrigger); if (FAILED(hr)) goto finish;
                    hr = pTrigger.query_interface(&pTriggerLogon);                  if (FAILED(hr)) goto finish;
                    sprintf(str, L"PT%uM", ttData.wRandomMinutesInterval);
                    hr = pTriggerLogon->put_Delay(winstd::bstr(str));               if (FAILED(hr)) goto finish;
                    break;
                }
            }

            // Set trigger ID.
            sprintf(str, L"%u", iTrigger);
            hr = pTrigger->put_Id(winstd::bstr(str));
            if (FAILED(hr)) goto finish;

            // Set trigger start date.
            sprintf(str, L"%04u-%02u-%02uT%02u:%02u:00", ttData.wBeginYear, ttData.wBeginMonth, ttData.wBeginDay, ttData.wStartHour, ttData.wStartMinute);
            hr = pTrigger->put_StartBoundary(winstd::bstr(str));
            if (FAILED(hr)) goto finish;

            if (ttData.rgFlags & TASK_TRIGGER_FLAG_HAS_END_DATE) {
                // Set trigger end date.
                sprintf(str, L"%04u-%02u-%02u", ttData.wEndYear, ttData.wEndMonth, ttData.wEndDay);
                hr = pTrigger->put_EndBoundary(winstd::bstr(str));
                if (FAILED(hr)) goto finish;
            }

            // Set trigger repetition duration and interval.
            if (ttData.MinutesDuration || ttData.MinutesInterval) {
                winstd::com_obj<IRepetitionPattern> pRepetitionPattern;

                hr = pTrigger->get_Repetition(&pRepetitionPattern);
                if (FAILED(hr)) goto finish;
                sprintf(str, L"PT%uM", ttData.MinutesDuration);
                hr = pRepetitionPattern->put_Duration(winstd::bstr(str));
                if (FAILED(hr)) goto finish;
                sprintf(str, L"PT%uM", ttData.MinutesInterval);
                hr = pRepetitionPattern->put_Interval(winstd::bstr(str));
                if (FAILED(hr)) goto finish;
                hr = pRepetitionPattern->put_StopAtDurationEnd(ttData.rgFlags & TASK_TRIGGER_FLAG_KILL_AT_DURATION_END ? VARIANT_TRUE : VARIANT_FALSE);
                if (FAILED(hr)) goto finish;
            }

            // Enable/disable trigger.
            hr = pTrigger->put_Enabled(ttData.rgFlags & TASK_TRIGGER_FLAG_DISABLED ? VARIANT_FALSE : VARIANT_TRUE);
            if (FAILED(hr)) goto finish;
        }

        // Get the task folder.
        hr = pService->GetFolder(winstd::bstr(L"\\"), &pTaskFolder);
        if (FAILED(hr)) goto finish;

#ifdef _DEBUG
        winstd::bstr xml;
        hr = pTaskDefinition->get_XmlText(&xml);
#endif

        // Register the task.
        hr = pTaskFolder->RegisterTaskDefinition(
            winstd::bstr(m_sValue), // path
            pTaskDefinition,        // pDefinition
            TASK_CREATE,            // flags
            vEmpty,                 // userId
            logonType != TASK_LOGON_SERVICE_ACCOUNT && !m_sPassword.empty() ? winstd::variant(m_sPassword.c_str()) : vEmpty, // password
            logonType,              // logonType
            vEmpty,                 // sddl
            &pTask);                // ppTask
    } else {
        // Windows XP or older.
        winstd::com_obj<ITaskScheduler> pTaskScheduler;
        winstd::com_obj<ITask> pTask;

        // Get task scheduler object.
        hr = pTaskScheduler.create(CLSID_CTaskScheduler, NULL, CLSCTX_ALL);
        if (FAILED(hr)) goto finish;

        // Create the new task.
        hr = pTaskScheduler->NewWorkItem(m_sValue.c_str(), CLSID_CTask, IID_ITask, (IUnknown**)&pTask);
        if (pSession->m_bRollbackEnabled) {
            // Order rollback action to delete the task. ITask::NewWorkItem() might made a blank task already.
            pSession->m_olRollback.push_front(new COpTaskDelete(m_sValue.c_str()));
        }
        if (FAILED(hr)) goto finish;

        // Set its properties.
        hr = pTask->SetApplicationName (m_sApplicationName.c_str() ); if (FAILED(hr)) goto finish;
        hr = pTask->SetParameters      (m_sParameters.c_str()      ); if (FAILED(hr)) goto finish;
        hr = pTask->SetWorkingDirectory(m_sWorkingDirectory.c_str()); if (FAILED(hr)) goto finish;
        hr = pTask->SetComment         (m_sComment.c_str()         ); if (FAILED(hr)) goto finish;
        if (pSession->m_bRollbackEnabled && (m_dwFlags & TASK_FLAG_DISABLED) == 0) {
            // The task is supposed to be enabled.
            // However, we shall enable it at commit to prevent it from accidentally starting in the very installation phase.
            pSession->m_olCommit.push_back(new COpTaskEnable(m_sValue.c_str(), TRUE));
            m_dwFlags |= TASK_FLAG_DISABLED;
        }
        hr = pTask->SetFlags           (m_dwFlags          ); if (FAILED(hr)) goto finish;
        hr = pTask->SetPriority        (m_dwPriority       ); if (FAILED(hr)) goto finish;

        // Set task credentials.
        hr = m_sAccountName.empty() ?
            pTask->SetAccountInformation(L""                   , NULL               ) :
            pTask->SetAccountInformation(m_sAccountName.c_str(), m_sPassword.c_str());
        if (FAILED(hr)) goto finish;

        hr = pTask->SetIdleWait  (m_wIdleMinutes, m_wDeadlineMinutes); if (FAILED(hr)) goto finish;
        hr = pTask->SetMaxRunTime(m_dwMaxRuntimeMS                  ); if (FAILED(hr)) goto finish;

        // Add triggers.
        for (auto t = m_lTriggers.cbegin(), t_end = m_lTriggers.cend(); t != t_end; ++t) {
            WORD wTriggerIdx;
            winstd::com_obj<ITaskTrigger> pTrigger;
            TASK_TRIGGER ttData = *t; // Don't use reference! We don't want to modify original trigger data by adding random startup delay.

            hr = pTask->CreateTrigger(&wTriggerIdx, &pTrigger);
            if (FAILED(hr)) goto finish;

            if (ttData.wRandomMinutesInterval) {
                // Windows XP doesn't support random startup delay. However, we can add fixed "random" delay when creating the trigger.
                WORD wStartTime = ttData.wStartHour * 60 + ttData.wStartMinute + (WORD)::MulDiv(rand(), ttData.wRandomMinutesInterval, RAND_MAX);
                FILETIME ftValue;
                SYSTEMTIME stValue = {
                    ttData.wBeginYear,
                    ttData.wBeginMonth,
                    0,
                    ttData.wBeginDay,
                };
                ULONGLONG ullValue;

                // Convert MDY date to numerical date (SYSTEMTIME -> FILETIME -> ULONGLONG).
                ::SystemTimeToFileTime(&stValue, &ftValue);
                ullValue = ((ULONGLONG)(ftValue.dwHighDateTime) << 32) | ftValue.dwLowDateTime;

                // Wrap days.
                while (wStartTime >= 1440) {
                    ullValue   += (ULONGLONG)864000000000;
                    wStartTime -= 1440;
                }

                // Convert numerical date to DMY (ULONGLONG -> FILETIME -> SYSTEMTIME).
                ftValue.dwHighDateTime = static_cast<DWORD>((ullValue >> 32) & 0xffffffff);
                ftValue.dwLowDateTime  = static_cast<DWORD>( ullValue        & 0xffffffff);
                ::FileTimeToSystemTime(&ftValue, &stValue);

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
        winstd::com_obj<IPersistFile> pTaskFile(pTask);
        if (!pTaskFile) { hr = E_NOINTERFACE; goto finish; }
        hr = pTaskFile->Save(NULL, TRUE);
    }

finish:
    if (FAILED(hr)) {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_TASK_CREATE);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue.c_str()         );
        ::MsiRecordSetInteger(hRecordProg, 3, hr                       );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
    }
    return hr;
}


UINT COpTaskCreate::SetFromRecord(MSIHANDLE hInstall, MSIHANDLE hRecord)
{
    UINT uiResult;
    int iValue;
    std::wstring sFolder;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord,  3, m_sApplicationName);
    if (uiResult != NO_ERROR) return uiResult;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord,  4, m_sParameters);
    if (uiResult != NO_ERROR) return uiResult;

    uiResult = ::MsiRecordGetStringW(hRecord, 5, sFolder);
    if (uiResult != NO_ERROR) return uiResult;
    uiResult = ::MsiGetTargetPathW(hInstall, sFolder.c_str(), m_sWorkingDirectory);
    if (uiResult != NO_ERROR) return uiResult;
    if (!m_sWorkingDirectory.empty() && m_sWorkingDirectory.back() == L'\\') {
        // Trim trailing backslash.
        m_sWorkingDirectory.resize(m_sWorkingDirectory.length() - 1);
    }

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord, 10, m_sAuthor);
    if (uiResult != NO_ERROR) return uiResult;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord, 11, m_sComment);
    if (uiResult != NO_ERROR) return uiResult;

    m_dwFlags = ::MsiRecordGetInteger(hRecord,  6);
    if (m_dwFlags == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;

    m_dwPriority = ::MsiRecordGetInteger(hRecord,  7);
    if (m_dwPriority == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord,  8, m_sAccountName);
    if (uiResult != NO_ERROR) return uiResult;

    uiResult = ::MsiRecordFormatStringW(hInstall, hRecord,  9, m_sPassword);
    if (uiResult != NO_ERROR) return uiResult;

    iValue = ::MsiRecordGetInteger(hRecord, 12);
    m_wIdleMinutes = iValue != MSI_NULL_INTEGER ? (WORD)iValue : 0;

    iValue = ::MsiRecordGetInteger(hRecord, 13);
    m_wDeadlineMinutes = iValue != MSI_NULL_INTEGER ? (WORD)iValue : 0;

    m_dwMaxRuntimeMS = ::MsiRecordGetInteger(hRecord, 14);
    if (m_dwMaxRuntimeMS == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;

    return NO_ERROR;
}


UINT COpTaskCreate::SetTriggersFromView(MSIHANDLE hView)
{
    for (;;) {
        UINT uiResult;
        PMSIHANDLE hRecord;
        TASK_TRIGGER ttData = { sizeof(TASK_TRIGGER) };
        ULONGLONG ullValue;
        FILETIME ftValue;
        SYSTEMTIME stValue;
        int iValue;

        // Fetch one record from the view.
        uiResult = ::MsiViewFetch(hView, &hRecord);
        if (uiResult == ERROR_NO_MORE_ITEMS) return NO_ERROR;
        else if (uiResult != NO_ERROR)  return uiResult;

        // Get StartDate.
        iValue = ::MsiRecordGetInteger(hRecord, 2);
        if (iValue == MSI_NULL_INTEGER) return ERROR_INVALID_FIELD;
        ullValue = ((ULONGLONG)iValue + 138426) * 864000000000;
        ftValue.dwHighDateTime = static_cast<DWORD>((ullValue >> 32) & 0xffffffff);
        ftValue.dwLowDateTime  = static_cast<DWORD>( ullValue        & 0xffffffff);
        if (!::FileTimeToSystemTime(&ftValue, &stValue))
            return ::GetLastError();
        ttData.wBeginYear  = stValue.wYear;
        ttData.wBeginMonth = stValue.wMonth;
        ttData.wBeginDay   = stValue.wDay;

        // Get EndDate.
        iValue = ::MsiRecordGetInteger(hRecord, 3);
        if (iValue != MSI_NULL_INTEGER) {
            ullValue = ((ULONGLONG)iValue + 138426) * 864000000000;
            ftValue.dwHighDateTime = static_cast<DWORD>((ullValue >> 32) & 0xffffffff);
            ftValue.dwLowDateTime  = static_cast<DWORD>( ullValue        & 0xffffffff);
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

        m_lTriggers.push_back(ttData);
    }
}


////////////////////////////////////////////////////////////////////////////
// COpTaskDelete
////////////////////////////////////////////////////////////////////////////

COpTaskDelete::COpTaskDelete(LPCWSTR pszTaskName, int iTicks) :
    COpTypeSingleString(pszTaskName, iTicks)
{
}


HRESULT COpTaskDelete::Execute(CSession *pSession)
{
    HRESULT hr;
    winstd::com_obj<ITaskService> pService;

    hr = pService.create(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER);
    if (SUCCEEDED(hr)) {
        // Windows Vista or newer.
        winstd::variant vEmpty;
        winstd::com_obj<ITaskFolder> pTaskFolder;

        // Connect to local task service.
        hr = pService->Connect(vEmpty, vEmpty, vEmpty, vEmpty);
        if (FAILED(hr)) goto finish;

        // Get task folder.
        hr = pService->GetFolder(winstd::bstr(L"\\"), &pTaskFolder);
        if (FAILED(hr)) goto finish;

        if (pSession->m_bRollbackEnabled) {
            winstd::com_obj<IRegisteredTask> pTask, pTaskOrig;
            winstd::com_obj<ITaskDefinition> pTaskDefinition;
            winstd::com_obj<IPrincipal> pPrincipal;
            VARIANT_BOOL bEnabled;
            TASK_LOGON_TYPE logonType;
            winstd::bstr sSSDL;
            winstd::variant vSSDL;
            std::wstring sDisplayNameOrig;
            UINT uiCount = 0;

            // Get the source task.
            hr = pTaskFolder->GetTask(winstd::bstr(m_sValue), &pTask);
            if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) { hr = S_OK; goto finish; }
            else if (FAILED(hr)) goto finish;

            // Disable the task.
            hr = pTask->get_Enabled(&bEnabled);
            if (FAILED(hr)) goto finish;
            if (bEnabled) {
                // The task is enabled.

                // In case the task disabling fails halfway, try to re-enable it anyway on rollback.
                pSession->m_olRollback.push_front(new COpTaskEnable(m_sValue.c_str(), TRUE));

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
            if (hr == HRESULT_FROM_WIN32(ERROR_PRIVILEGE_NOT_HELD)) ::VariantClear(&vSSDL);
            else if (FAILED(hr)) goto finish;
            else {
                V_VT  (&vSSDL) = VT_BSTR;
                V_BSTR(&vSSDL) = sSSDL.detach();
            }

            // Register a backup copy of task.
            do {
                sprintf(sDisplayNameOrig, L"%ls (orig %u)", m_sValue.c_str(), ++uiCount);
                hr = pTaskFolder->RegisterTaskDefinition(
                    winstd::bstr(sDisplayNameOrig), // path
                    pTaskDefinition,                // pDefinition
                    TASK_CREATE,                    // flags
                    vEmpty,                         // userId
                    vEmpty,                         // password
                    logonType,                      // logonType
                    vSSDL,                          // sddl
                    &pTaskOrig);                    // ppTask
            } while (hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS));
            // In case the backup copy creation failed halfway, try to delete its remains anyway on rollback.
            pSession->m_olRollback.push_front(new COpTaskDelete(sDisplayNameOrig.c_str()));
            if (FAILED(hr)) goto finish;

            // Order rollback action to restore from backup copy.
            pSession->m_olRollback.push_front(new COpTaskCopy(sDisplayNameOrig.c_str(), m_sValue.c_str()));

            // Delete it.
            hr = pTaskFolder->DeleteTask(winstd::bstr(m_sValue), 0);
            if (FAILED(hr)) goto finish;

            // Order commit action to delete backup copy.
            pSession->m_olCommit.push_back(new COpTaskDelete(sDisplayNameOrig.c_str()));
        } else {
            // Delete the task.
            hr = pTaskFolder->DeleteTask(winstd::bstr(m_sValue), 0);
            if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) hr = S_OK;
        }
    } else {
        // Windows XP or older.
        winstd::com_obj<ITaskScheduler> pTaskScheduler;

        // Get task scheduler object.
        hr = pTaskScheduler.create(CLSID_CTaskScheduler, NULL, CLSCTX_ALL);
        if (FAILED(hr)) goto finish;

        if (pSession->m_bRollbackEnabled) {
            winstd::com_obj<ITask> pTask;
            DWORD dwFlags;
            std::wstring sDisplayNameOrig;
            UINT uiCount = 0;

            // Load the task.
            hr = pTaskScheduler->Activate(m_sValue.c_str(), IID_ITask, (IUnknown**)&pTask);
            if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) { hr = S_OK; goto finish; }
            else if (FAILED(hr)) goto finish;

            // Disable the task.
            hr = pTask->GetFlags(&dwFlags);
            if (FAILED(hr)) goto finish;
            if ((dwFlags & TASK_FLAG_DISABLED) == 0) {
                // The task is enabled.

                // In case the task disabling fails halfway, try to re-enable it anyway on rollback.
                pSession->m_olRollback.push_front(new COpTaskEnable(m_sValue.c_str(), TRUE));

                // Disable it.
                dwFlags |= TASK_FLAG_DISABLED;
                hr = pTask->SetFlags(dwFlags);
                if (FAILED(hr)) goto finish;
            }

            // Prepare a backup copy of task.
            do {
                sprintf(sDisplayNameOrig, L"%ls (orig %u)", m_sValue.c_str(), ++uiCount);
                hr = pTaskScheduler->AddWorkItem(sDisplayNameOrig.c_str(), pTask);
            } while (hr == HRESULT_FROM_WIN32(ERROR_FILE_EXISTS));
            // In case the backup copy creation failed halfway, try to delete its remains anyway on rollback.
            pSession->m_olRollback.push_front(new COpTaskDelete(sDisplayNameOrig.c_str()));
            if (FAILED(hr)) goto finish;

            // Save the backup copy.
            winstd::com_obj<IPersistFile> pTaskFile(pTask);
            if (!pTaskFile) { hr = E_NOINTERFACE; goto finish; }
            hr = pTaskFile->Save(NULL, TRUE);
            if (FAILED(hr)) goto finish;

            // Order rollback action to restore from backup copy.
            pSession->m_olRollback.push_front(new COpTaskCopy(sDisplayNameOrig.c_str(), m_sValue.c_str()));

            // Delete it.
            hr = pTaskScheduler->Delete(m_sValue.c_str());
            if (FAILED(hr)) goto finish;

            // Order commit action to delete backup copy.
            pSession->m_olCommit.push_back(new COpTaskDelete(sDisplayNameOrig.c_str()));
        } else {
            // Delete the task.
            hr = pTaskScheduler->Delete(m_sValue.c_str());
            if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) hr = S_OK;
        }
    }

finish:
    if (FAILED(hr)) {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_TASK_DELETE);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue.c_str()         );
        ::MsiRecordSetInteger(hRecordProg, 3, hr                       );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
    }
    return hr;
}


////////////////////////////////////////////////////////////////////////////
// COpTaskEnable
////////////////////////////////////////////////////////////////////////////

COpTaskEnable::COpTaskEnable(LPCWSTR pszTaskName, BOOL bEnable, int iTicks) :
    m_bEnable(bEnable),
    COpTypeSingleString(pszTaskName, iTicks)
{
}


HRESULT COpTaskEnable::Execute(CSession *pSession)
{
    HRESULT hr;
    winstd::com_obj<ITaskService> pService;

    hr = pService.create(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER);
    if (SUCCEEDED(hr)) {
        // Windows Vista or newer.
        winstd::variant vEmpty;
        winstd::com_obj<ITaskFolder> pTaskFolder;
        winstd::com_obj<IRegisteredTask> pTask;
        VARIANT_BOOL bEnabled;

        // Connect to local task service.
        hr = pService->Connect(vEmpty, vEmpty, vEmpty, vEmpty);
        if (FAILED(hr)) goto finish;

        // Get task folder.
        hr = pService->GetFolder(winstd::bstr(L"\\"), &pTaskFolder);
        if (FAILED(hr)) goto finish;

        // Get task.
        hr = pTaskFolder->GetTask(winstd::bstr(m_sValue), &pTask);
        if (FAILED(hr)) goto finish;

        // Get currently enabled state.
        hr = pTask->get_Enabled(&bEnabled);
        if (FAILED(hr)) goto finish;

        // Modify enable state.
        if (m_bEnable) {
            if (pSession->m_bRollbackEnabled && !bEnabled) {
                // The task is disabled and supposed to be enabled.
                // However, we shall enable it at commit to prevent it from accidentally starting in the very installation phase.
                pSession->m_olCommit.push_back(new COpTaskEnable(m_sValue.c_str(), TRUE));
                hr = S_OK; goto finish;
            } else
                bEnabled = VARIANT_TRUE;
        } else {
            if (pSession->m_bRollbackEnabled && bEnabled) {
                // The task is enabled and we will disable it now.
                // Order rollback to re-enable it.
                pSession->m_olRollback.push_front(new COpTaskEnable(m_sValue.c_str(), TRUE));
            }
            bEnabled = VARIANT_FALSE;
        }

        // Set enable/disable.
        hr = pTask->put_Enabled(bEnabled);
        if (FAILED(hr)) goto finish;
    } else {
        // Windows XP or older.
        winstd::com_obj<ITaskScheduler> pTaskScheduler;
        winstd::com_obj<ITask> pTask;
        DWORD dwFlags;

        // Get task scheduler object.
        hr = pTaskScheduler.create(CLSID_CTaskScheduler, NULL, CLSCTX_ALL);
        if (FAILED(hr)) goto finish;

        // Load the task.
        hr = pTaskScheduler->Activate(m_sValue.c_str(), IID_ITask, (IUnknown**)&pTask);
        if (FAILED(hr)) goto finish;

        // Get task's current flags.
        hr = pTask->GetFlags(&dwFlags);
        if (FAILED(hr)) goto finish;

        // Modify flags.
        if (m_bEnable) {
            if (pSession->m_bRollbackEnabled && (dwFlags & TASK_FLAG_DISABLED)) {
                // The task is disabled and supposed to be enabled.
                // However, we shall enable it at commit to prevent it from accidentally starting in the very installation phase.
                pSession->m_olCommit.push_back(new COpTaskEnable(m_sValue.c_str(), TRUE));
                hr = S_OK; goto finish;
            } else
                dwFlags &= ~TASK_FLAG_DISABLED;
        } else {
            if (pSession->m_bRollbackEnabled && !(dwFlags & TASK_FLAG_DISABLED)) {
                // The task is enabled and we will disable it now.
                // Order rollback to re-enable it.
                pSession->m_olRollback.push_front(new COpTaskEnable(m_sValue.c_str(), TRUE));
            }
            dwFlags |=  TASK_FLAG_DISABLED;
        }

        // Set flags.
        hr = pTask->SetFlags(dwFlags);
        if (FAILED(hr)) goto finish;

        // Save the task.
        winstd::com_obj<IPersistFile> pTaskFile(pTask);
        if (!pTaskFile) { hr = E_NOINTERFACE; goto finish; }
        hr = pTaskFile->Save(NULL, TRUE);
        if (FAILED(hr)) goto finish;
    }

finish:
    if (FAILED(hr)) {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_TASK_ENABLE);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue.c_str()         );
        ::MsiRecordSetInteger(hRecordProg, 3, hr                       );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
    }
    return hr;
}


////////////////////////////////////////////////////////////////////////////
// COpTaskCopy
////////////////////////////////////////////////////////////////////////////

COpTaskCopy::COpTaskCopy(LPCWSTR pszTaskSrc, LPCWSTR pszTaskDst, int iTicks) :
    COpTypeSrcDstString(pszTaskSrc, pszTaskDst, iTicks)
{
}


HRESULT COpTaskCopy::Execute(CSession *pSession)
{
    HRESULT hr;
    winstd::com_obj<ITaskService> pService;

    hr = pService.create(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER);
    if (SUCCEEDED(hr)) {
        // Windows Vista or newer.
        winstd::variant vEmpty;
        winstd::com_obj<ITaskFolder> pTaskFolder;
        winstd::com_obj<IRegisteredTask> pTask, pTaskOrig;
        winstd::com_obj<ITaskDefinition> pTaskDefinition;
        winstd::com_obj<IPrincipal> pPrincipal;
        TASK_LOGON_TYPE logonType;
        winstd::bstr sSSDL;

        // Connect to local task service.
        hr = pService->Connect(vEmpty, vEmpty, vEmpty, vEmpty);
        if (FAILED(hr)) goto finish;

        // Get task folder.
        hr = pService->GetFolder(winstd::bstr(L"\\"), &pTaskFolder);
        if (FAILED(hr)) goto finish;

        // Get the source task.
        hr = pTaskFolder->GetTask(winstd::bstr(m_sValue1), &pTask);
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
            winstd::bstr(m_sValue2), // path
            pTaskDefinition,         // pDefinition
            TASK_CREATE,             // flags
            vEmpty,                  // userId
            vEmpty,                  // password
            logonType,               // logonType
            winstd::variant(sSSDL),      // sddl
            &pTaskOrig);             // ppTask
        if (FAILED(hr)) goto finish;
    } else {
        // Windows XP or older.
        winstd::com_obj<ITaskScheduler> pTaskScheduler;
        winstd::com_obj<ITask> pTask;

        // Get task scheduler object.
        hr = pTaskScheduler.create(CLSID_CTaskScheduler, NULL, CLSCTX_ALL);
        if (FAILED(hr)) goto finish;

        // Load the source task.
        hr = pTaskScheduler->Activate(m_sValue1.c_str(), IID_ITask, (IUnknown**)&pTask);
        if (FAILED(hr)) goto finish;

        // Add with different name.
        hr = pTaskScheduler->AddWorkItem(m_sValue2.c_str(), pTask);
        if (pSession->m_bRollbackEnabled) {
            // Order rollback action to delete the new copy. ITask::AddWorkItem() might made a blank task already.
            pSession->m_olRollback.push_front(new COpTaskDelete(m_sValue2.c_str()));
        }
        if (FAILED(hr)) goto finish;

        // Save the task.
        winstd::com_obj<IPersistFile> pTaskFile(pTask);
        if (!pTaskFile) { hr = E_NOINTERFACE; goto finish; }
        hr = pTaskFile->Save(NULL, TRUE);
        if (FAILED(hr)) goto finish;
    }

finish:
    if (FAILED(hr)) {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(4);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_TASK_COPY);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue1.c_str()      );
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue2.c_str()      );
        ::MsiRecordSetInteger(hRecordProg, 4, hr                     );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
    }
    return hr;
}

} // namespace MSICA
