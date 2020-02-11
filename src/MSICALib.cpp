/*
    Copyright 1991-2020 Amebis

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

#include "pch.h"

#pragma comment(lib, "msi.lib")
#pragma comment(lib, "shlwapi.lib")


namespace MSICA {

////////////////////////////////////////////////////////////////////////////
// COperation
////////////////////////////////////////////////////////////////////////////

COperation::COperation(int iTicks) :
    m_iTicks(iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// COpTypeSingleString
////////////////////////////////////////////////////////////////////////////

COpTypeSingleString::COpTypeSingleString(LPCWSTR pszValue, int iTicks) :
    m_sValue(pszValue),
    COperation(iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// COpTypeSrcDstString
////////////////////////////////////////////////////////////////////////////

COpTypeSrcDstString::COpTypeSrcDstString(LPCWSTR pszValue1, LPCWSTR pszValue2, int iTicks) :
    m_sValue1(pszValue1),
    m_sValue2(pszValue2),
    COperation(iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// COpTypeBoolean
////////////////////////////////////////////////////////////////////////////

COpTypeBoolean::COpTypeBoolean(BOOL bValue, int iTicks) :
    m_bValue(bValue),
    COperation(iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// COpRollbackEnable
////////////////////////////////////////////////////////////////////////////

COpRollbackEnable::COpRollbackEnable(BOOL bEnable, int iTicks) :
    COpTypeBoolean(bEnable, iTicks)
{
}


HRESULT COpRollbackEnable::Execute(CSession *pSession)
{
    pSession->m_bRollbackEnabled = m_bValue;
    return S_OK;
}


////////////////////////////////////////////////////////////////////////////
// COpList
////////////////////////////////////////////////////////////////////////////

COpList::COpList(int iTicks) : COperation(iTicks)
{
}


DWORD COpList::LoadFromFile(LPCTSTR pszFileName)
{
    CStream fSequence;
    if (!fSequence.create(pszFileName, GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN))
        return GetLastError();

    // Load operation sequence.
    if (!(fSequence >> *this))
        return GetLastError();

    return NO_ERROR;
}


DWORD COpList::SaveToFile(LPCTSTR pszFileName) const
{
    CStream fSequence;
    if (!fSequence.create(pszFileName, GENERIC_WRITE, FILE_SHARE_READ, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN))
        return GetLastError();

    // Save execute sequence to file.
    DWORD dwResult = (fSequence << *this) ? NO_ERROR : GetLastError();
    fSequence.free();

    if (dwResult != NO_ERROR) ::DeleteFile(pszFileName);
    return dwResult;
}


HRESULT COpList::Execute(CSession *pSession)
{
    HRESULT hr;
    PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);

    // Tell the installer to use explicit progress messages.
    ::MsiRecordSetInteger(hRecordProg, 1, 1);
    ::MsiRecordSetInteger(hRecordProg, 2, 1);
    ::MsiRecordSetInteger(hRecordProg, 3, 0);
    ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_PROGRESS, hRecordProg);

    // Prepare hRecordProg for progress messages.
    ::MsiRecordSetInteger(hRecordProg, 1, 2);
    ::MsiRecordSetInteger(hRecordProg, 3, 0);

    for (auto op = cbegin(), op_end = cend(); op != op_end; ++op) {
        COperation *pOp = op->get();

        hr = pOp->Execute(pSession);
        if (!pSession->m_bContinueOnError && FAILED(hr)) {
            // Operation failed. Its Execute() method should have sent error message to Installer.
            // Therefore, just quit here.
            return hr;
        }

        ::MsiRecordSetInteger(hRecordProg, 2, pOp->m_iTicks);
        if (::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_PROGRESS, hRecordProg) == IDCANCEL)
            return HRESULT_FROM_WIN32(ERROR_INSTALL_USEREXIT);
    }

    ::MsiRecordSetInteger(hRecordProg, 2, m_iTicks);
    ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_PROGRESS, hRecordProg);

    return S_OK;
}


////////////////////////////////////////////////////////////////////////////
// CSession
////////////////////////////////////////////////////////////////////////////

CSession::CSession() :
    m_hInstall(NULL),
    m_bContinueOnError(FALSE),
    m_bRollbackEnabled(FALSE)
{
}


////////////////////////////////////////////////////////////////////////////
// Helper functions
////////////////////////////////////////////////////////////////////////////

UINT SaveSequence(MSIHANDLE hInstall, LPCTSTR szActionExecute, LPCTSTR szActionCommit, LPCTSTR szActionRollback, const COpList &olExecute)
{
    UINT uiResult;
    DWORD dwResult;
    winstd::tstring sSequenceFilename;
    PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);

    // Prepare our own sequence script file.
    // The InstallCertificates is a deferred custom action, thus all this information will be unavailable to it.
    // Therefore save all required info to file now.
    {
        std::unique_ptr<TCHAR> szBuffer(new TCHAR[MAX_PATH]);
        if (!szBuffer) return ERROR_OUTOFMEMORY;
        ::GetTempPath(MAX_PATH, szBuffer.get());
        ::GetTempFileName(szBuffer.get(), _T("MSICA"), 0, szBuffer.get());
        sSequenceFilename.assign(szBuffer.get(), wcsnlen(szBuffer.get(), MAX_PATH));
    }
    // Save execute sequence to file.
    dwResult = olExecute.SaveToFile(sSequenceFilename.c_str());
    if (dwResult == NO_ERROR) {
        // Store sequence script file names to properties for deferred custiom actions.
        uiResult = ::MsiSetProperty(hInstall, szActionExecute, sSequenceFilename.c_str());
        if (uiResult == NO_ERROR) {
            LPCTSTR pszExtension = ::PathFindExtension(sSequenceFilename.c_str());
            winstd::tstring sSequenceFilename2;

            sprintf(sSequenceFilename2, _T("%.*ls-rb%ls"), pszExtension - sSequenceFilename.c_str(), sSequenceFilename.c_str(), pszExtension);
            uiResult = ::MsiSetProperty(hInstall, szActionRollback, sSequenceFilename2.c_str());
            if (uiResult == NO_ERROR) {
                sprintf(sSequenceFilename2, _T("%.*ls-cm%ls"), pszExtension - sSequenceFilename.c_str(), sSequenceFilename.c_str(), pszExtension);
                uiResult = ::MsiSetProperty(hInstall, szActionCommit, sSequenceFilename2.c_str());
                if (uiResult != NO_ERROR) {
                    ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_PROPERTY_SET);
                    ::MsiRecordSetString (hRecordProg, 2, szActionCommit            );
                    ::MsiRecordSetInteger(hRecordProg, 3, uiResult                  );
                    ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                }
            } else {
                ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_PROPERTY_SET);
                ::MsiRecordSetString (hRecordProg, 2, szActionRollback          );
                ::MsiRecordSetInteger(hRecordProg, 3, uiResult                  );
                ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
            }
        } else {
            ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_PROPERTY_SET);
            ::MsiRecordSetString (hRecordProg, 2, szActionExecute           );
            ::MsiRecordSetInteger(hRecordProg, 3, uiResult                  );
            ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        }
        if (uiResult != NO_ERROR) ::DeleteFile(sSequenceFilename.c_str());
    } else {
        uiResult = ERROR_INSTALL_SCRIPT_WRITE;
        ::MsiRecordSetInteger(hRecordProg, 1, uiResult                 );
        ::MsiRecordSetString (hRecordProg, 2, sSequenceFilename.c_str());
        ::MsiRecordSetInteger(hRecordProg, 3, dwResult                 );
        ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
    }

    return uiResult;
}


UINT ExecuteSequence(MSIHANDLE hInstall)
{
    UINT uiResult;
    DWORD dwResult;
    HRESULT hr;
    winstd::com_initializer com_init(NULL);
    winstd::tstring sSequenceFilename;

    uiResult = ::MsiGetProperty(hInstall, _T("CustomActionData"), sSequenceFilename);
    if (uiResult == NO_ERROR) {
        MSICA::COpList lstOperations;
        BOOL bIsCleanup = ::MsiGetMode(hInstall, MSIRUNMODE_COMMIT) || ::MsiGetMode(hInstall, MSIRUNMODE_ROLLBACK);

        // Load operation sequence.
        dwResult = lstOperations.LoadFromFile(sSequenceFilename.c_str());
        if (dwResult == NO_ERROR) {
            MSICA::CSession session;

            session.m_hInstall = hInstall;

            // In case of commit/rollback, continue sequence on error, to do as much cleanup as possible.
            session.m_bContinueOnError = bIsCleanup;

            // Execute the operations.
            hr = lstOperations.Execute(&session);
            if (!bIsCleanup) {
                // Save cleanup scripts of delayed action regardless of action's execution status.
                // Rollback action MUST be scheduled in InstallExecuteSequence before this action! Otherwise cleanup won't be performed in case this action execution failed.
                LPCTSTR pszExtension = ::PathFindExtension(sSequenceFilename.c_str());
                winstd::tstring sSequenceFilenameCM, sSequenceFilenameRB;

                sprintf(sSequenceFilenameRB, _T("%.*ls-rb%ls"), pszExtension - sSequenceFilename.c_str(), sSequenceFilename.c_str(), pszExtension);
                sprintf(sSequenceFilenameCM, _T("%.*ls-cm%ls"), pszExtension - sSequenceFilename.c_str(), sSequenceFilename.c_str(), pszExtension);

                // After commit, delete rollback file. After rollback, delete commit file.
                session.m_olCommit.push_back(new MSICA::COpFileDelete(
#ifdef _UNICODE
                    sSequenceFilenameRB.c_str()
#else
                    std::wstring(sSequenceFilenameRB)
#endif
                    ));
                session.m_olRollback.push_back(new MSICA::COpFileDelete(
#ifdef _UNICODE
                    sSequenceFilenameCM.c_str()
#else
                    std::wstring(sSequenceFilenameCM)
#endif
                    ));

                // Save commit file first.
                dwResult = session.m_olCommit.SaveToFile(sSequenceFilenameCM.c_str());
                if (dwResult == NO_ERROR) {
                    // Save rollback file next.
                    dwResult = session.m_olRollback.SaveToFile(sSequenceFilenameRB.c_str());
                    if (dwResult == NO_ERROR) {
                        uiResult = NO_ERROR;
                    } else {
                        // Saving rollback file failed.
                        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
                        uiResult = ERROR_INSTALL_SCRIPT_WRITE;
                        ::MsiRecordSetInteger(hRecordProg, 1, uiResult                   );
                        ::MsiRecordSetString (hRecordProg, 2, sSequenceFilenameRB.c_str());
                        ::MsiRecordSetInteger(hRecordProg, 3, dwResult                   );
                        ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                    }
                } else {
                    // Saving commit file failed.
                    PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
                    uiResult = ERROR_INSTALL_SCRIPT_WRITE;
                    ::MsiRecordSetInteger(hRecordProg, 1, uiResult                   );
                    ::MsiRecordSetString (hRecordProg, 2, sSequenceFilenameCM.c_str());
                    ::MsiRecordSetInteger(hRecordProg, 3, dwResult                   );
                    ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                }

                if (uiResult != NO_ERROR) {
                    // The commit and/or rollback scripts were not written to file successfully. Perform the cleanup immediately.
                    session.m_bContinueOnError = TRUE;
                    session.m_bRollbackEnabled = FALSE;
                    session.m_olRollback.Execute(&session);
                    ::DeleteFile(sSequenceFilenameRB.c_str());
                }
            } else {
                // No cleanup after cleanup support.
                uiResult = NO_ERROR;
            }

            if (FAILED(hr)) {
                // Execution of the action failed.
                uiResult = HRESULT_CODE(hr);
            }

            ::DeleteFile(sSequenceFilename.c_str());
        } else if (dwResult == ERROR_FILE_NOT_FOUND && bIsCleanup) {
            // Sequence file not found and this is rollback/commit action. Either of the following scenarios are possible:
            // - The delayed action failed to save the rollback/commit file. The delayed action performed cleanup itself. No further action is required.
            // - Somebody removed the rollback/commit file between delayed action and rollback/commit action. No further action is possible.
            uiResult = NO_ERROR;
        } else {
            // Sequence loading failed. Probably, LOCAL SYSTEM doesn't have read access to user's temp directory.
            PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
            uiResult = ERROR_INSTALL_SCRIPT_READ;
            ::MsiRecordSetInteger(hRecordProg, 1, uiResult                 );
            ::MsiRecordSetString (hRecordProg, 2, sSequenceFilename.c_str());
            ::MsiRecordSetInteger(hRecordProg, 3, dwResult                 );
            ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        }
    } else {
        // Couldn't get CustomActionData property. uiResult has the error code.
    }

    return uiResult;
}

} // namespace MSICA
