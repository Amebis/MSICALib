﻿/*
    SPDX-License-Identifier: GPL-3.0-or-later
    Copyright © 1991-2022 Amebis
*/

#include "pch.h"


namespace MSICA {

////////////////////////////////////////////////////////////////////////////
// COpFileDelete
////////////////////////////////////////////////////////////////////////////

COpFileDelete::COpFileDelete(LPCWSTR pszFileName, int iTicks) :
    COpTypeSingleString(pszFileName, iTicks)
{
}


HRESULT COpFileDelete::Execute(CSession *pSession)
{
    DWORD dwError;

    if (pSession->m_bRollbackEnabled) {
        std::wstring sBackupName;
        UINT uiCount = 0;

        do {
            // Rename the file to make a backup.
            sprintf(sBackupName, L"%ls (orig %u)", m_sValue.c_str(), ++uiCount);
            dwError = ::MoveFileW(m_sValue.c_str(), sBackupName.c_str()) ? NO_ERROR : ::GetLastError();
        } while (dwError == ERROR_ALREADY_EXISTS);
        if (dwError == NO_ERROR) {
            // Order rollback action to restore from backup copy.
            pSession->m_olRollback.push_front(new COpFileMove(sBackupName.c_str(), m_sValue.c_str()));

            // Order commit action to delete backup copy.
            pSession->m_olCommit.push_back(new COpFileDelete(sBackupName.c_str()));
        }
    } else {
        // Delete the file.
        dwError = ::DeleteFileW(m_sValue.c_str()) ? NO_ERROR : ::GetLastError();
    }

    if (dwError == NO_ERROR || dwError == ERROR_FILE_NOT_FOUND)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_FILE_DELETE);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue.c_str()         );
        ::MsiRecordSetInteger(hRecordProg, 3, dwError                  );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return HRESULT_FROM_WIN32(dwError);
    }
}


////////////////////////////////////////////////////////////////////////////
// COpFileMove
////////////////////////////////////////////////////////////////////////////

COpFileMove::COpFileMove(LPCWSTR pszFileSrc, LPCWSTR pszFileDst, int iTicks) :
    COpTypeSrcDstString(pszFileSrc, pszFileDst, iTicks)
{
}


HRESULT COpFileMove::Execute(CSession *pSession)
{
    DWORD dwError;

    // Move the file.
    dwError = ::MoveFileW(m_sValue1.c_str(), m_sValue2.c_str()) ? NO_ERROR : ::GetLastError();
    if (dwError == NO_ERROR) {
        if (pSession->m_bRollbackEnabled) {
            // Order rollback action to move it back.
            pSession->m_olRollback.push_front(new COpFileMove(m_sValue2.c_str(), m_sValue1.c_str()));
        }

        return S_OK;
    } else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(4);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_FILE_MOVE);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue1.c_str()      );
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue2.c_str()      );
        ::MsiRecordSetInteger(hRecordProg, 4, dwError                );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return HRESULT_FROM_WIN32(dwError);
    }
}

} // namespace MSICA
