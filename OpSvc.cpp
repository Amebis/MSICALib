#include "stdafx.h"

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
        hService = ::OpenServiceW(hSCM, m_sValue, SERVICE_QUERY_CONFIG | SERVICE_CHANGE_CONFIG);
        if (hService) {
            QUERY_SERVICE_CONFIG sc;
            DWORD dwSize;

            // Query current service config.
            if (::QueryServiceConfig(hService, &sc, sizeof(sc), &dwSize)) {
                if (sc.dwStartType != m_dwStartType) {
                    // Set requested service start.
                    if (::ChangeServiceConfig(hService, SERVICE_NO_CHANGE, m_dwStartType, SERVICE_NO_CHANGE, NULL, NULL, NULL, NULL, NULL, NULL, NULL)) {
                        if (pSession->m_bRollbackEnabled) {
                            // Order rollback action to revert the service start change.
                            pSession->m_olRollback.AddHead(new COpSvcSetStart(m_sValue, sc.dwStartType));
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
        } else
            dwError = ::GetLastError();
    } else
        dwError = ::GetLastError();

    if (dwError == NO_ERROR)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_SVC_SET_START_FAILED);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue                          );
        ::MsiRecordSetInteger(hRecordProg, 3, dwError                           );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(dwError);
    }
}


////////////////////////////////////////////////////////////////////////////
// COpSvcStart
////////////////////////////////////////////////////////////////////////////

COpSvcStart::COpSvcStart(LPCWSTR pszService, int iTicks) : COpTypeSingleString(pszService, iTicks)
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
        hService = ::OpenServiceW(hSCM, m_sValue, SERVICE_START);
        if (hService) {
            // Start the service.
            if (::StartService(hService, 0, NULL)) {
                if (pSession->m_bRollbackEnabled) {
                    // Order rollback action to stop the service.
                    pSession->m_olRollback.AddHead(new COpSvcStop(m_sValue));
                }
                dwError = NO_ERROR;
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
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_SVC_START_FAILED);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue                      );
        ::MsiRecordSetInteger(hRecordProg, 3, dwError                       );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(dwError);
    }
}


////////////////////////////////////////////////////////////////////////////
// COpSvcStop
////////////////////////////////////////////////////////////////////////////

COpSvcStop::COpSvcStop(LPCWSTR pszService, int iTicks) : COpTypeSingleString(pszService, iTicks)
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
        hService = ::OpenServiceW(hSCM, m_sValue, SERVICE_STOP);
        if (hService) {
            SERVICE_STATUS ss;
            // Stop the service.
            if (::ControlService(hService, SERVICE_CONTROL_STOP, &ss)) {
                if (pSession->m_bRollbackEnabled) {
                    // Order rollback action to start the service.
                    pSession->m_olRollback.AddHead(new COpSvcStart(m_sValue));
                }
                dwError = NO_ERROR;
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
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_SVC_STOP_FAILED);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue                     );
        ::MsiRecordSetInteger(hRecordProg, 3, dwError                      );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(dwError);
    }
}

} // namespace MSICA
