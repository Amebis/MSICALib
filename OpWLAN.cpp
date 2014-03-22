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
            ATL::CAtlStringW sBackupName;
            UINT uiCount = 0;

            do {
                // Rename the profile to make a backup.
                sBackupName.Format(L"%ls (orig %u)", (LPCWSTR)m_sValue, ++uiCount);
                dwError = ::WlanRenameProfile(hClientHandle, &m_guidInterface, m_sValue, sBackupName, NULL);
            } while (dwError == ERROR_ALREADY_EXISTS);
            if (dwError == NO_ERROR) {
                // Order rollback action to restore from backup copy.
                pSession->m_olRollback.AddHead(new COpWLANProfileRename(m_guidInterface, sBackupName, m_sValue));

                // Order commit action to delete backup copy.
                pSession->m_olCommit.AddTail(new COpWLANProfileDelete(m_guidInterface, sBackupName));
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
        GuidToString(m_guidInterface, sGUID);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_WLAN_PROFILE_DELETE);
        ::MsiRecordSetStringW(hRecordProg, 2, sGUID                            );
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue                         );
        ::MsiRecordSetInteger(hRecordProg, 4, dwError                          );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(dwError);
    }
}


////////////////////////////////////////////////////////////////////////////
// COpWLANProfileRename
////////////////////////////////////////////////////////////////////////////

COpWLANProfileRename::COpWLANProfileRename(const GUID &guidInterface, LPCWSTR pszProfileNameSrc, LPCWSTR pszProfileNameDst, int iTicks) :
    COpWLANProfile(guidInterface, pszProfileNameSrc, iTicks),
    m_sProfileNameDst(pszProfileNameDst)
{
}


HRESULT COpWLANProfileRename::Execute(CSession *pSession)
{
    DWORD dwError, dwNegotiatedVersion;
    HANDLE hClientHandle;

    dwError = ::WlanOpenHandle(2, NULL, &dwNegotiatedVersion, &hClientHandle);
    if (dwError == NO_ERROR) {
        // Rename the profile.
        dwError = ::WlanRenameProfile(hClientHandle, &m_guidInterface, m_sValue, m_sProfileNameDst, NULL);
        if (dwError == NO_ERROR && pSession->m_bRollbackEnabled) {
            // Order rollback action to rename it back.
            pSession->m_olRollback.AddHead(new COpWLANProfileRename(m_guidInterface, m_sProfileNameDst, m_sValue));
        }
        ::WlanCloseHandle(hClientHandle, NULL);
    }

    if (dwError == NO_ERROR)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(5);
        ATL::CAtlStringW sGUID;
        GuidToString(m_guidInterface, sGUID);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_WLAN_PROFILE_RENAME);
        ::MsiRecordSetStringW(hRecordProg, 2, sGUID                            );
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue                         );
        ::MsiRecordSetStringW(hRecordProg, 4, m_sProfileNameDst                );
        ::MsiRecordSetInteger(hRecordProg, 5, dwError                          );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(dwError);
    }
}


////////////////////////////////////////////////////////////////////////////
// COpWLANProfileSet
////////////////////////////////////////////////////////////////////////////

COpWLANProfileSet::COpWLANProfileSet(const GUID &guidInterface, LPCWSTR pszProfileName, LPCWSTR pszProfileXML, int iTicks) :
    COpWLANProfile(guidInterface, pszProfileName, iTicks),
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
        dwError = ::WlanSetProfile(hClientHandle, &m_guidInterface, 0, m_sProfileXML, NULL, TRUE, NULL, &wlrc);
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
        DWORD dwSize = 1024;
        LPWSTR szBuffer = sReason.GetBuffer(dwSize);

        GuidToString(m_guidInterface, sGUID);
        if (::WlanReasonCodeToString(wlrc, dwSize, szBuffer, NULL) == NO_ERROR) {
            sReason.ReleaseBuffer(dwSize);
            ::MsiRecordSetStringW(hRecordProg, 4, sReason);
        } else {
            sReason.ReleaseBuffer(dwSize);
            sReason.Format(L"0x%x", wlrc);
        }

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
