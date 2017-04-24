/*
    Copyright 1991-2017 Amebis

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
    if (::pfnWlanOpenHandle && ::pfnWlanCloseHandle && ::pfnWlanGetProfile && ::pfnWlanDeleteProfile && ::pfnWlanFreeMemory) {
        DWORD dwError, dwNegotiatedVersion;
        HANDLE hClientHandle;

        dwError = ::pfnWlanOpenHandle(2, NULL, &dwNegotiatedVersion, &hClientHandle);
        if (dwError == NO_ERROR) {
            if (pSession->m_bRollbackEnabled) {
                LPWSTR pszProfileXML = NULL;
                DWORD dwFlags = 0, dwGrantedAccess = 0;

                // Get profile settings as XML first.
                dwError = ::pfnWlanGetProfile(hClientHandle, &m_guidInterface, m_sValue.c_str(), NULL, &pszProfileXML, &dwFlags, &dwGrantedAccess);
                if (dwError == NO_ERROR) {
                    // Delete the profile.
                    dwError = ::pfnWlanDeleteProfile(hClientHandle, &m_guidInterface, m_sValue.c_str(), NULL);
                    if (dwError == NO_ERROR) {
                        // Order rollback action to recreate it.
                        pSession->m_olRollback.push_front(new COpWLANProfileSet(m_guidInterface, dwFlags, m_sValue.c_str(), pszProfileXML));
                    }
                    ::pfnWlanFreeMemory(pszProfileXML);
                }
            } else {
                // Delete the profile.
                dwError = ::pfnWlanDeleteProfile(hClientHandle, &m_guidInterface, m_sValue.c_str(), NULL);
            }
            ::pfnWlanCloseHandle(hClientHandle, NULL);
        }

        if (dwError == NO_ERROR || dwError == ERROR_NOT_FOUND)
            return S_OK;
        else {
            PMSIHANDLE hRecordProg = ::MsiCreateRecord(4);
            ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_WLAN_PROFILE_DELETE            );
            ::MsiRecordSetStringW(hRecordProg, 2, winstd::wstring_guid(m_guidInterface).c_str());
            ::MsiRecordSetStringW(hRecordProg, 3, m_sValue.c_str()                             );
            ::MsiRecordSetInteger(hRecordProg, 4, dwError                                      );
            ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
            return HRESULT_FROM_WIN32(dwError);
        }
    } else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(1);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_WLAN_NOT_INSTALLED);
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return E_NOTIMPL;
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
    if (::pfnWlanOpenHandle && ::pfnWlanCloseHandle && ::pfnWlanSetProfile && ::pfnWlanReasonCodeToString) {
        DWORD dwError, dwNegotiatedVersion;
        HANDLE hClientHandle;
        WLAN_REASON_CODE wlrc = 0;

        {
            // Delete existing profile first.
            // Since deleting a profile is a complicated job (when rollback/commit support is required), and we do have an operation just for that, we use it.
            // Don't worry, COpWLANProfileDelete::Execute() returns S_OK if profile doesn't exist.
            COpWLANProfileDelete opDelete(m_guidInterface, m_sValue.c_str());
            HRESULT hr = opDelete.Execute(pSession);
            if (FAILED(hr)) return hr;
        }

        dwError = ::pfnWlanOpenHandle(2, NULL, &dwNegotiatedVersion, &hClientHandle);
        if (dwError == NO_ERROR) {
            // Set the profile.
            dwError = ::pfnWlanSetProfile(hClientHandle, &m_guidInterface, m_dwFlags, m_sProfileXML.c_str(), NULL, TRUE, NULL, &wlrc);
            if (dwError == NO_ERROR && pSession->m_bRollbackEnabled) {
                // Order rollback action to delete it.
                pSession->m_olRollback.push_front(new COpWLANProfileDelete(m_guidInterface, m_sValue.c_str()));
            }
            ::pfnWlanCloseHandle(hClientHandle, NULL);
        }

        if (dwError == NO_ERROR)
            return S_OK;
        else {
            PMSIHANDLE hRecordProg = ::MsiCreateRecord(5);
            std::wstring sReason;
            std::unique_ptr<WCHAR[]> szBuffer(new WCHAR[1024]);
            if (szBuffer && ::pfnWlanReasonCodeToString(wlrc, 1024, szBuffer.get(), NULL) == NO_ERROR)
                sReason.assign(szBuffer.get(), wcsnlen(szBuffer.get(), 1024));
            else
                sprintf(sReason, L"0x%x", wlrc);

            ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_WLAN_PROFILE_SET);
            ::MsiRecordSetStringW(hRecordProg, 2, winstd::wstring_guid(m_guidInterface).c_str());
            ::MsiRecordSetStringW(hRecordProg, 3, m_sValue.c_str()              );
            ::MsiRecordSetStringW(hRecordProg, 4, sReason.c_str()               );
            ::MsiRecordSetInteger(hRecordProg, 5, dwError                       );
            ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
            return HRESULT_FROM_WIN32(dwError);
        }
    } else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(1);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_WLAN_NOT_INSTALLED);
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return E_NOTIMPL;
    }
}

} // namespace MSICA
