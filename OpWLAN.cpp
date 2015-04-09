/*
    Copyright 1991-2015 Amebis

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

#pragma comment(lib, "wlanapi.lib")


namespace MSICA {

////////////////////////////////////////////////////////////////////////////
// COpWLANProfile
////////////////////////////////////////////////////////////////////////////

COpWLANProfile::COpWLANProfile(const GUID &guidInterface, LPCWSTR pszProfileName, int iTicks) :
    COpTypeSingleString(pszProfileName, iTicks),
    m_guidInterface(guidInterface)
{
}


////////////////////////////////////////////////////////////////////////////
// COpWLANProfileDelete
////////////////////////////////////////////////////////////////////////////

COpWLANProfileDelete::COpWLANProfileDelete(const GUID &guidInterface, LPCWSTR pszProfileName, int iTicks) : COpWLANProfile(guidInterface, pszProfileName, iTicks)
{
}


HRESULT COpWLANProfileDelete::Execute(CSession *pSession)
{
    DWORD dwError, dwNegotiatedVersion;
    HANDLE hClientHandle;

    dwError = ::WlanOpenHandle(2, NULL, &dwNegotiatedVersion, &hClientHandle);
    if (dwError == NO_ERROR) {
        if (pSession->m_bRollbackEnabled) {
            LPWSTR pszProfileXML = NULL;
            DWORD dwFlags = 0, dwGrantedAccess = 0;

            // Get profile settings as XML first.
            dwError = ::WlanGetProfile(hClientHandle, &m_guidInterface, m_sValue, NULL, &pszProfileXML, &dwFlags, &dwGrantedAccess);
            if (dwError == NO_ERROR) {
                // Delete the profile.
                dwError = ::WlanDeleteProfile(hClientHandle, &m_guidInterface, m_sValue, NULL);
                if (dwError == NO_ERROR) {
                    // Order rollback action to recreate it.
                    pSession->m_olRollback.AddHead(new COpWLANProfileSet(m_guidInterface, dwFlags, m_sValue, pszProfileXML));
                }
                ::WlanFreeMemory(pszProfileXML);
            }
        } else {
            // Delete the profile.
            dwError = ::WlanDeleteProfile(hClientHandle, &m_guidInterface, m_sValue, NULL);
        }
        ::WlanCloseHandle(hClientHandle, NULL);
    }

    if (dwError == NO_ERROR || dwError == ERROR_NOT_FOUND)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(4);
        ATL::CAtlStringW sGUID;
        GuidToString(&m_guidInterface, sGUID);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_WLAN_PROFILE_DELETE);
        ::MsiRecordSetStringW(hRecordProg, 2, sGUID                            );
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue                         );
        ::MsiRecordSetInteger(hRecordProg, 4, dwError                          );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(dwError);
    }
}


////////////////////////////////////////////////////////////////////////////
// COpWLANProfileSet
////////////////////////////////////////////////////////////////////////////

COpWLANProfileSet::COpWLANProfileSet(const GUID &guidInterface, DWORD dwFlags, LPCWSTR pszProfileName, LPCWSTR pszProfileXML, int iTicks) :
    COpWLANProfile(guidInterface, pszProfileName, iTicks),
    m_dwFlags(dwFlags),
    m_sProfileXML(pszProfileXML)
{
}


HRESULT COpWLANProfileSet::Execute(CSession *pSession)
{
    DWORD dwError, dwNegotiatedVersion;
    HANDLE hClientHandle;
    WLAN_REASON_CODE wlrc = 0;

    {
        // Delete existing profile first.
        // Since deleting a profile is a complicated job (when rollback/commit support is required), and we do have an operation just for that, we use it.
        // Don't worry, COpWLANProfileDelete::Execute() returns S_OK if profile doesn't exist.
        COpWLANProfileDelete opDelete(m_guidInterface, m_sValue);
        HRESULT hr = opDelete.Execute(pSession);
        if (FAILED(hr)) return hr;
    }

    dwError = ::WlanOpenHandle(2, NULL, &dwNegotiatedVersion, &hClientHandle);
    if (dwError == NO_ERROR) {
        // Set the profile.
        dwError = ::WlanSetProfile(hClientHandle, &m_guidInterface, m_dwFlags, m_sProfileXML, NULL, TRUE, NULL, &wlrc);
        if (dwError == NO_ERROR && pSession->m_bRollbackEnabled) {
            // Order rollback action to delete it.
            pSession->m_olRollback.AddHead(new COpWLANProfileDelete(m_guidInterface, m_sValue));
        }
        ::WlanCloseHandle(hClientHandle, NULL);
    }

    if (dwError == NO_ERROR)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(5);
        ATL::CAtlStringW sGUID, sReason;
        DWORD dwSize = 1024, dwResult;
        LPWSTR szBuffer = sReason.GetBuffer(dwSize);

        GuidToString(&m_guidInterface, sGUID);
        dwResult = ::WlanReasonCodeToString(wlrc, dwSize, szBuffer, NULL);
        sReason.ReleaseBuffer(dwSize);
        if (dwResult != NO_ERROR) sReason.Format(L"0x%x", wlrc);

        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_WLAN_PROFILE_SET);
        ::MsiRecordSetStringW(hRecordProg, 2, sGUID                         );
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue                      );
        ::MsiRecordSetStringW(hRecordProg, 4, sReason                       );
        ::MsiRecordSetInteger(hRecordProg, 5, dwError                       );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(dwError);
    }
}

} // namespace MSICA
