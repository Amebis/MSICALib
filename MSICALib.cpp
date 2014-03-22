#include "stdafx.h"

#pragma comment(lib, "msi.lib")


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

COpList::COpList(int iTicks) :
    COperation(iTicks),
    ATL::CAtlList<COperation*>(sizeof(COperation*))
{
}


void COpList::Free()
{
    POSITION pos;

    for (pos = GetHeadPosition(); pos;) {
        COperation *pOp = GetNext(pos);
        COpList *pOpList = dynamic_cast<COpList*>(pOp);

        if (pOpList) {
            // Recursivelly free sublists.
            pOpList->Free();
        }
        delete pOp;
    }

    RemoveAll();
}


HRESULT COpList::LoadFromFile(LPCTSTR pszFileName)
{
    HRESULT hr;
    ATL::CAtlFile fSequence;

    hr = fSequence.Create(pszFileName, GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN);
    if (FAILED(hr)) return hr;

    // Load operation sequence.
    return fSequence >> *this;
}


HRESULT COpList::SaveToFile(LPCTSTR pszFileName) const
{
    HRESULT hr;
    ATL::CAtlFile fSequence;

    hr = fSequence.Create(pszFileName, GENERIC_WRITE, FILE_SHARE_READ, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN);
    if (FAILED(hr)) return hr;

    // Save execute sequence to file.
    hr = fSequence << *this;
    fSequence.Close();

    if (FAILED(hr)) ::DeleteFile(pszFileName);
    return hr;
}


HRESULT COpList::Execute(CSession *pSession)
{
    POSITION pos;
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

    for (pos = GetHeadPosition(); pos;) {
        COperation *pOp = GetNext(pos);

        hr = pOp->Execute(pSession);
        if (!pSession->m_bContinueOnError && FAILED(hr)) {
            // Operation failed. Its Execute() method should have sent error message to Installer.
            // Therefore, just quit here.
            return hr;
        }

        ::MsiRecordSetInteger(hRecordProg, 2, pOp->m_iTicks);
        if (::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_PROGRESS, hRecordProg) == IDCANCEL)
            return AtlHresultFromWin32(ERROR_INSTALL_USEREXIT);
    }

    ::MsiRecordSetInteger(hRecordProg, 2, m_iTicks);
    ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_PROGRESS, hRecordProg);

    return S_OK;
}


////////////////////////////////////////////////////////////////////////////
// CSession
////////////////////////////////////////////////////////////////////////////

CSession::CSession() :
    m_bContinueOnError(FALSE),
    m_bRollbackEnabled(FALSE)
{
}


CSession::~CSession()
{
    m_olRollback.Free();
    m_olCommit.Free();
}


////////////////////////////////////////////////////////////////////////////
// Helper functions
////////////////////////////////////////////////////////////////////////////

UINT SaveSequence(MSIHANDLE hInstall, LPCTSTR szActionExecute, LPCTSTR szActionCommit, LPCTSTR szActionRollback, const COpList &olExecute)
{
    HRESULT hr;
    UINT uiResult;
    ATL::CAtlString sSequenceFilename;
    ATL::CAtlFile fSequence;
    PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);

    // Prepare our own sequence script file.
    // The InstallCertificates is a deferred custom action, thus all this information will be unavailable to it.
    // Therefore save all required info to file now.
    {
        LPTSTR szBuffer = sSequenceFilename.GetBuffer(MAX_PATH);
        ::GetTempPath(MAX_PATH, szBuffer);
        ::GetTempFileName(szBuffer, _T("MSICA"), 0, szBuffer);
        sSequenceFilename.ReleaseBuffer();
    }
    // Save execute sequence to file.
    hr = olExecute.SaveToFile(sSequenceFilename);
    if (SUCCEEDED(hr)) {
        // Store sequence script file names to properties for deferred custiom actions.
        uiResult = ::MsiSetProperty(hInstall, szActionExecute, sSequenceFilename);
        if (uiResult == NO_ERROR) {
            LPCTSTR pszExtension = ::PathFindExtension(sSequenceFilename);
            ATL::CAtlString sSequenceFilename2;

            sSequenceFilename2.Format(_T("%.*ls-rb%ls"), pszExtension - (LPCTSTR)sSequenceFilename, (LPCTSTR)sSequenceFilename, pszExtension);
            uiResult = ::MsiSetProperty(hInstall, szActionRollback, sSequenceFilename2);
            if (uiResult == NO_ERROR) {
                sSequenceFilename2.Format(_T("%.*ls-cm%ls"), pszExtension - (LPCTSTR)sSequenceFilename, (LPCTSTR)sSequenceFilename, pszExtension);
                uiResult = ::MsiSetProperty(hInstall, szActionCommit, sSequenceFilename2);
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
        if (uiResult != NO_ERROR) ::DeleteFile(sSequenceFilename);
    } else {
        uiResult = ERROR_INSTALL_SCRIPT_WRITE;
        ::MsiRecordSetInteger(hRecordProg, 1, uiResult         );
        ::MsiRecordSetString (hRecordProg, 2, sSequenceFilename);
        ::MsiRecordSetInteger(hRecordProg, 3, hr               );
        ::MsiProcessMessage(hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
    }

    return uiResult;
}


UINT ExecuteSequence(MSIHANDLE hInstall)
{
    UINT uiResult;
    HRESULT hr;
    BOOL bIsCoInitialized = SUCCEEDED(::CoInitialize(NULL));
    ATL::CAtlString sSequenceFilename;

    uiResult = ::MsiGetProperty(hInstall, _T("CustomActionData"), sSequenceFilename);
    if (uiResult == NO_ERROR) {
        MSICA::COpList lstOperations;
        BOOL bIsCleanup = ::MsiGetMode(hInstall, MSIRUNMODE_COMMIT) || ::MsiGetMode(hInstall, MSIRUNMODE_ROLLBACK);

        // Load operation sequence.
        hr = lstOperations.LoadFromFile(sSequenceFilename);
        if (SUCCEEDED(hr)) {
            MSICA::CSession session;

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
                session.m_olCommit.AddTail(new MSICA::COpFileDelete(
#ifdef _UNICODE
                    sSequenceFilenameRB
#else
                    ATL::CAtlStringW(sSequenceFilenameRB)
#endif
                    ));
                session.m_olRollback.AddTail(new MSICA::COpFileDelete(
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
                        uiResult = NO_ERROR;
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

                if (uiResult != NO_ERROR) {
                    // The commit and/or rollback scripts were not written to file successfully. Perform the cleanup immediately.
                    session.m_bContinueOnError = TRUE;
                    session.m_bRollbackEnabled = FALSE;
                    session.m_olRollback.Execute(&session);
                    ::DeleteFile(sSequenceFilenameRB);
                }
            } else {
                // No cleanup after cleanup support.
                uiResult = NO_ERROR;
            }

            if (FAILED(hr)) {
                // Execution of the action failed.
                uiResult = HRESULT_CODE(hr);
            }

            ::DeleteFile(sSequenceFilename);
        } else if (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && bIsCleanup) {
            // Sequence file not found and this is rollback/commit action. Either of the following scenarios are possible:
            // - The delayed action failed to save the rollback/commit file. The delayed action performed cleanup itself. No further action is required.
            // - Somebody removed the rollback/commit file between delayed action and rollback/commit action. No further action is possible.
            uiResult = NO_ERROR;
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

} // namespace MSICA
