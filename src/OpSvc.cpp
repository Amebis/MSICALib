﻿/*
    SPDX-License-Identifier: GPL-3.0-or-later
    Copyright © 1991-2022 Amebis
*/

#include "pch.h"

#pragma comment(lib, "advapi32.lib")


namespace MSICA {

////////////////////////////////////////////////////////////////////////////
// COpSvcSetStart
////////////////////////////////////////////////////////////////////////////

COpSvcSetStart::COpSvcSetStart(LPCWSTR pszService, DWORD dwStartType, int iTicks) :
    COpTypeSingleString(pszService, iTicks),
    m_dwStartType(dwStartType)
{
}


HRESULT COpSvcSetStart::Execute(CSession *pSession)
{
    DWORD dwError;
    SC_HANDLE hSCM;

    // Open Service Control Manager.
    hSCM = ::OpenSCManager(NULL, SERVICES_ACTIVE_DATABASE, SC_MANAGER_CONNECT);
    if (hSCM) {
        SC_HANDLE hService;

        // Open the specified service.
        hService = ::OpenServiceW(hSCM, m_sValue.c_str(), SERVICE_QUERY_CONFIG | SERVICE_CHANGE_CONFIG);
        if (hService) {
            QUERY_SERVICE_CONFIG *sc;
            DWORD dwSize;
            HANDLE hHeap = ::GetProcessHeap();

            ::QueryServiceConfig(hService, NULL, 0, &dwSize);
            sc = (QUERY_SERVICE_CONFIG*)::HeapAlloc(hHeap, 0, dwSize);
            if (sc) {
                // Query current service config.
                if (::QueryServiceConfig(hService, sc, dwSize, &dwSize)) {
                    if (sc->dwStartType != m_dwStartType) {
                        // Set requested service start.
                        if (::ChangeServiceConfig(hService, SERVICE_NO_CHANGE, m_dwStartType, SERVICE_NO_CHANGE, NULL, NULL, NULL, NULL, NULL, NULL, NULL)) {
                            if (pSession->m_bRollbackEnabled) {
                                // Order rollback action to revert the service start change.
                                pSession->m_olRollback.push_front(new COpSvcSetStart(m_sValue.c_str(), sc->dwStartType));
                            }
                            dwError = NO_ERROR;
                        } else
                            dwError = ::GetLastError();
                    } else {
                        // Service is already configured to start the requested way.
                        dwError = NO_ERROR;
                    }
                } else
                    dwError = ::GetLastError();

                ::HeapFree(hHeap, 0, sc);
            } else
                dwError = ERROR_OUTOFMEMORY;
        } else
            dwError = ::GetLastError();
    } else
        dwError = ::GetLastError();

    if (dwError == NO_ERROR)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_SVC_SET_START);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue.c_str()           );
        ::MsiRecordSetInteger(hRecordProg, 3, dwError                    );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return HRESULT_FROM_WIN32(dwError);
    }
}


////////////////////////////////////////////////////////////////////////////
// COpSvcControl
////////////////////////////////////////////////////////////////////////////

COpSvcControl::COpSvcControl(LPCWSTR pszService, BOOL bWait, int iTicks) :
    COpTypeSingleString(pszService, iTicks),
    m_bWait(bWait)
{
}


DWORD COpSvcControl::WaitForState(CSession *pSession, SC_HANDLE hService, DWORD dwPendingState, DWORD dwFinalState)
{
    PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);

    // Prepare hRecordProg for progress messages.
    ::MsiRecordSetInteger(hRecordProg, 1, 2);
    ::MsiRecordSetInteger(hRecordProg, 3, 0);

    // Wait for the service to start.
    for (;;) {
        SERVICE_STATUS_PROCESS ssp;
        DWORD dwSize;

        // Check service status.
        if (::QueryServiceStatusEx(hService, SC_STATUS_PROCESS_INFO, (LPBYTE)&ssp, sizeof(SERVICE_STATUS_PROCESS), &dwSize)) {
            if (ssp.dwCurrentState == dwPendingState) {
                // Service is pending. Wait some more ...
                ::Sleep(ssp.dwWaitHint < 1000 ? ssp.dwWaitHint : 1000);
            } else if (ssp.dwCurrentState == dwFinalState) {
                // Service is in expected state.
                return NO_ERROR;
            } else {
                // Service is in unexpected state.
                return ERROR_ASSERTION_FAILURE;
            }
        } else {
            // Service query failed.
            return ::GetLastError();
        }

        // Check if user cancelled installation.
        ::MsiRecordSetInteger(hRecordProg, 2, 0);
        if (::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_PROGRESS, hRecordProg) == IDCANCEL)
            return ERROR_INSTALL_USEREXIT;
    }
}


////////////////////////////////////////////////////////////////////////////
// COpSvcStart
////////////////////////////////////////////////////////////////////////////

COpSvcStart::COpSvcStart(LPCWSTR pszService, BOOL bWait, int iTicks) : COpSvcControl(pszService, bWait, iTicks)
{
}


HRESULT COpSvcStart::Execute(CSession *pSession)
{
    DWORD dwError;
    SC_HANDLE hSCM;

    // Open Service Control Manager.
    hSCM = ::OpenSCManager(NULL, SERVICES_ACTIVE_DATABASE, SC_MANAGER_CONNECT);
    if (hSCM) {
        SC_HANDLE hService;

        // Open the specified service.
        hService = ::OpenServiceW(hSCM, m_sValue.c_str(), SERVICE_START | (m_bWait ? SERVICE_QUERY_STATUS : 0));
        if (hService) {
            // Start the service.
            if (::StartService(hService, 0, NULL)) {
                dwError = m_bWait ? WaitForState(pSession, hService, SERVICE_START_PENDING, SERVICE_RUNNING) : NO_ERROR;
                if (dwError == NO_ERROR && pSession->m_bRollbackEnabled) {
                    // Order rollback action to stop the service.
                    pSession->m_olRollback.push_front(new COpSvcStop(m_sValue.c_str()));
                }
            } else {
                dwError = ::GetLastError();
                if (dwError == ERROR_SERVICE_ALREADY_RUNNING) {
                    // Service is already running. Not an error.
                    dwError = NO_ERROR;
                }
            }
        } else
            dwError = ::GetLastError();
    } else
        dwError = ::GetLastError();

    if (dwError == NO_ERROR)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_SVC_START);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue.c_str()       );
        ::MsiRecordSetInteger(hRecordProg, 3, dwError                );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return HRESULT_FROM_WIN32(dwError);
    }
}


////////////////////////////////////////////////////////////////////////////
// COpSvcStop
////////////////////////////////////////////////////////////////////////////

COpSvcStop::COpSvcStop(LPCWSTR pszService, BOOL bWait, int iTicks) : COpSvcControl(pszService, bWait, iTicks)
{
}


HRESULT COpSvcStop::Execute(CSession *pSession)
{
    DWORD dwError;
    SC_HANDLE hSCM;

    // Open Service Control Manager.
    hSCM = ::OpenSCManager(NULL, SERVICES_ACTIVE_DATABASE, SC_MANAGER_CONNECT);
    if (hSCM) {
        SC_HANDLE hService;

        // Open the specified service.
        hService = ::OpenServiceW(hSCM, m_sValue.c_str(), SERVICE_STOP | (m_bWait ? SERVICE_QUERY_STATUS : 0));
        if (hService) {
            SERVICE_STATUS ss;
            // Stop the service.
            if (::ControlService(hService, SERVICE_CONTROL_STOP, &ss)) {
                dwError = m_bWait ? WaitForState(pSession, hService, SERVICE_STOP_PENDING, SERVICE_STOPPED) : NO_ERROR;
                if (dwError == NO_ERROR && pSession->m_bRollbackEnabled) {
                    // Order rollback action to start the service.
                    pSession->m_olRollback.push_front(new COpSvcStart(m_sValue.c_str()));
                }
            } else {
                dwError = ::GetLastError();
                if (dwError == ERROR_SERVICE_NOT_ACTIVE) {
                    // Service is already stopped. Not an error.
                    dwError = NO_ERROR;
                }
            }
        } else
            dwError = ::GetLastError();
    } else
        dwError = ::GetLastError();

    if (dwError == NO_ERROR)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_SVC_STOP);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue.c_str()      );
        ::MsiRecordSetInteger(hRecordProg, 3, dwError               );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return HRESULT_FROM_WIN32(dwError);
    }
}

} // namespace MSICA
