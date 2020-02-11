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

#pragma once

#include <WinStd\Common.h>
#include <WinStd\Win.h>

#include <list>
#include <string>

#include <msi.h>
#include <mstask.h>
#include <wincrypt.h>
#include <windows.h>


////////////////////////////////////////////////////////////////////
// Error codes (next unused 2581L)
////////////////////////////////////////////////////////////////////

#define ERROR_INSTALL_DATABASE_OPEN          2550L
#define ERROR_INSTALL_OPLIST_CREATE          2551L
#define ERROR_INSTALL_PROPERTY_SET           2553L
#define ERROR_INSTALL_SCRIPT_WRITE           2552L
#define ERROR_INSTALL_SCRIPT_READ            2560L
#define ERROR_INSTALL_FILE_DELETE            2554L
#define ERROR_INSTALL_FILE_MOVE              2555L
#define ERROR_INSTALL_REGKEY_CREATE          2561L
#define ERROR_INSTALL_REGKEY_COPY            2562L
#define ERROR_INSTALL_REGKEY_PROBING         2563L
#define ERROR_INSTALL_REGKEY_DELETE          2564L
#define ERROR_INSTALL_REGKEY_SETVALUE        2565L
#define ERROR_INSTALL_REGKEY_DELETEVALUE     2567L
#define ERROR_INSTALL_REGKEY_COPYVALUE       2568L
#define ERROR_INSTALL_REGKEY_PROBINGVAL      2566L
#define ERROR_INSTALL_TASK_CREATE            2556L
#define ERROR_INSTALL_TASK_DELETE            2557L
#define ERROR_INSTALL_TASK_ENABLE            2558L
#define ERROR_INSTALL_TASK_COPY              2559L
#define ERROR_INSTALL_CERT_INSTALL           2569L
#define ERROR_INSTALL_CERT_REMOVE            2570L
#define ERROR_INSTALL_SVC_SET_START          2571L
#define ERROR_INSTALL_SVC_START              2572L
#define ERROR_INSTALL_SVC_STOP               2573L
#define ERROR_INSTALL_WLAN_NOT_INSTALLED     2580L
#define ERROR_INSTALL_WLAN_SVC_NOT_STARTED   2579L
#define ERROR_INSTALL_WLAN_HANDLE_OPEN       2577L
#define ERROR_INSTALL_WLAN_PROFILE_NOT_UTF16 2578L
#define ERROR_INSTALL_WLAN_PROFILE_DELETE    2574L
#define ERROR_INSTALL_WLAN_PROFILE_RENAME    2575L
#define ERROR_INSTALL_WLAN_PROFILE_SET       2576L


namespace MSICA {


////////////////////////////////////////////////////////////////////////////
// Forward declarations
////////////////////////////////////////////////////////////////////////////

class CSession;


////////////////////////////////////////////////////////////////////////////
// CStream
////////////////////////////////////////////////////////////////////////////

class CStream : public winstd::file
{
};


////////////////////////////////////////////////////////////////////////////
// COperation
////////////////////////////////////////////////////////////////////////////

class COperation
{
public:
    COperation(int iTicks = 0);

    virtual HRESULT Execute(CSession *pSession) = 0;

    friend class COpList;
    friend inline BOOL operator <<(CStream &f, const COperation &op);
    friend inline BOOL operator >>(CStream &f, COperation &op);

protected:
    int m_iTicks;   // Number of ticks on a progress bar required for this action execution
};


////////////////////////////////////////////////////////////////////////////
// COpTypeSingleString
////////////////////////////////////////////////////////////////////////////

class COpTypeSingleString : public COperation
{
public:
    COpTypeSingleString(LPCWSTR pszValue = L"", int iTicks = 0);

    friend inline BOOL operator <<(CStream &f, const COpTypeSingleString &op);
    friend inline BOOL operator >>(CStream &f, COpTypeSingleString &op);

protected:
    std::wstring m_sValue;
};


////////////////////////////////////////////////////////////////////////////
// COpTypeSrcDstString
////////////////////////////////////////////////////////////////////////////

class COpTypeSrcDstString : public COperation
{
public:
    COpTypeSrcDstString(LPCWSTR pszValue1 = L"", LPCWSTR pszValue2 = L"", int iTicks = 0);

    friend inline BOOL operator <<(CStream &f, const COpTypeSrcDstString &op);
    friend inline BOOL operator >>(CStream &f, COpTypeSrcDstString &op);

protected:
    std::wstring m_sValue1;
    std::wstring m_sValue2;
};


////////////////////////////////////////////////////////////////////////////
// COpTypeBoolean
////////////////////////////////////////////////////////////////////////////

class COpTypeBoolean : public COperation
{
public:
    COpTypeBoolean(BOOL bValue = TRUE, int iTicks = 0);

    friend inline BOOL operator <<(CStream &f, const COpTypeBoolean &op);
    friend inline BOOL operator >>(CStream &f, COpTypeBoolean &op);

protected:
    BOOL m_bValue;
};


////////////////////////////////////////////////////////////////////////////
// COpRollbackEnable
////////////////////////////////////////////////////////////////////////////

class COpRollbackEnable : public COpTypeBoolean
{
public:
    COpRollbackEnable(BOOL bEnable = TRUE, int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// COpFileDelete
////////////////////////////////////////////////////////////////////////////

class COpFileDelete : public COpTypeSingleString
{
public:
    COpFileDelete(LPCWSTR pszFileName = L"", int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// COpFileMove
////////////////////////////////////////////////////////////////////////////

class COpFileMove : public COpTypeSrcDstString
{
public:
    COpFileMove(LPCWSTR pszFileSrc = L"", LPCWSTR pszFileDst = L"", int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// COpRegKeySingle
////////////////////////////////////////////////////////////////////////////

class COpRegKeySingle : public COpTypeSingleString
{
public:
    COpRegKeySingle(HKEY hKeyRoot = NULL, LPCWSTR pszKeyName = L"", int iTicks = 0);

    friend inline BOOL operator <<(CStream &f, const COpRegKeySingle &op);
    friend inline BOOL operator >>(CStream &f, COpRegKeySingle &op);

protected:
    HKEY m_hKeyRoot;
};


////////////////////////////////////////////////////////////////////////////
// COpRegKeySrcDst
////////////////////////////////////////////////////////////////////////////

class COpRegKeySrcDst : public COpTypeSrcDstString
{
public:
    COpRegKeySrcDst(HKEY hKeyRoot = NULL, LPCWSTR pszKeyNameSrc = L"", LPCWSTR pszKeyNameDst = L"", int iTicks = 0);

    friend inline BOOL operator <<(CStream &f, const COpRegKeySrcDst &op);
    friend inline BOOL operator >>(CStream &f, COpRegKeySrcDst &op);

protected:
    HKEY m_hKeyRoot;
};


////////////////////////////////////////////////////////////////////////////
// COpRegKeyCreate
////////////////////////////////////////////////////////////////////////////

class COpRegKeyCreate : public COpRegKeySingle
{
public:
    COpRegKeyCreate(HKEY hKeyRoot = NULL, LPCWSTR pszKeyName = L"", int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// COpRegKeyCopy
////////////////////////////////////////////////////////////////////////////

class COpRegKeyCopy : public COpRegKeySrcDst
{
public:
    COpRegKeyCopy(HKEY hKeyRoot = NULL, LPCWSTR pszKeyNameSrc = L"", LPCWSTR pszKeyNameDst = L"", int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);

private:
    static LONG CopyKeyRecursively(HKEY hKeyRoot, LPCWSTR pszKeyNameSrc, LPCWSTR pszKeyNameDst, REGSAM samAdditional);
    static LONG CopyKeyRecursively(HKEY hKeySrc, HKEY hKeyDst, REGSAM samAdditional);
};


////////////////////////////////////////////////////////////////////////////
// COpRegKeyDelete
////////////////////////////////////////////////////////////////////////////

class COpRegKeyDelete : public COpRegKeySingle
{
public:
    COpRegKeyDelete(HKEY hKeyRoot = NULL, LPCWSTR pszKeyName = L"", int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);

private:
    static LONG DeleteKeyRecursively(HKEY hKeyRoot, LPCWSTR pszKeyName, REGSAM samAdditional);
};


////////////////////////////////////////////////////////////////////////////
// COpRegValueSingle
////////////////////////////////////////////////////////////////////////////

class COpRegValueSingle : public COpRegKeySingle
{
public:
    COpRegValueSingle(HKEY hKeyRoot = NULL, LPCWSTR pszKeyName = L"", LPCWSTR pszValueName = L"", int iTicks = 0);

    friend inline BOOL operator <<(CStream &f, const COpRegValueSingle &op);
    friend inline BOOL operator >>(CStream &f, COpRegValueSingle &op);

protected:
    std::wstring m_sValueName;
};


////////////////////////////////////////////////////////////////////////////
// COpRegValueSrcDst
////////////////////////////////////////////////////////////////////////////

class COpRegValueSrcDst : public COpRegKeySingle
{
public:
    COpRegValueSrcDst(HKEY hKeyRoot = NULL, LPCWSTR pszKeyName = L"", LPCWSTR pszValueNameSrc = L"", LPCWSTR pszValueNameDst = L"", int iTicks = 0);

    friend inline BOOL operator <<(CStream &f, const COpRegValueSrcDst &op);
    friend inline BOOL operator >>(CStream &f, COpRegValueSrcDst &op);

protected:
    std::wstring m_sValueName1;
    std::wstring m_sValueName2;
};


////////////////////////////////////////////////////////////////////////////
// COpRegValueCreate
////////////////////////////////////////////////////////////////////////////

class COpRegValueCreate : public COpRegValueSingle
{
public:
    COpRegValueCreate(HKEY hKeyRoot = NULL, LPCWSTR pszKeyName = L"", LPCWSTR pszValueName = L"", int iTicks = 0);
    COpRegValueCreate(HKEY hKeyRoot, LPCWSTR pszKeyName, LPCWSTR pszValueName, DWORD dwData, int iTicks = 0);
    COpRegValueCreate(HKEY hKeyRoot, LPCWSTR pszKeyName, LPCWSTR pszValueName, LPCVOID lpData, SIZE_T nSize, int iTicks = 0);
    COpRegValueCreate(HKEY hKeyRoot, LPCWSTR pszKeyName, LPCWSTR pszValueName, LPCWSTR pszData, int iTicks = 0);
    COpRegValueCreate(HKEY hKeyRoot, LPCWSTR pszKeyName, LPCWSTR pszValueName, DWORDLONG qwData, int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);

    friend inline BOOL operator <<(CStream &f, const COpRegValueCreate &op);
    friend inline BOOL operator >>(CStream &f, COpRegValueCreate &op);

protected:
    DWORD              m_dwType;
    std::wstring       m_sData;
    std::vector<BYTE>  m_binData;
    DWORD              m_dwData;
    std::vector<WCHAR> m_szData;
    DWORDLONG          m_qwData;
};


////////////////////////////////////////////////////////////////////////////
// COpRegValueCopy
////////////////////////////////////////////////////////////////////////////

class COpRegValueCopy : public COpRegValueSrcDst
{
public:
    COpRegValueCopy(HKEY hKeyRoot = NULL, LPCWSTR pszKeyName = L"", LPCWSTR pszValueNameSrc = L"", LPCWSTR pszValueNameDst = L"", int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// COpRegValueDelete
////////////////////////////////////////////////////////////////////////////

class COpRegValueDelete : public COpRegValueSingle
{
public:
    COpRegValueDelete(HKEY hKeyRoot = NULL, LPCWSTR pszKeyName = L"", LPCWSTR pszValueName = L"", int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// COpTaskCreate
////////////////////////////////////////////////////////////////////////////

class COpTaskCreate : public COpTypeSingleString
{
public:
    COpTaskCreate(LPCWSTR pszTaskName = L"", int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);

    UINT SetFromRecord(MSIHANDLE hInstall, MSIHANDLE hRecord);
    UINT SetTriggersFromView(MSIHANDLE hView);

    friend inline BOOL operator <<(CStream &f, const COpTaskCreate &op);
    friend inline BOOL operator >>(CStream &f, COpTaskCreate &op);

protected:
    std::wstring               m_sApplicationName;
    std::wstring               m_sParameters;
    std::wstring               m_sWorkingDirectory;
    std::wstring               m_sAuthor;
    std::wstring               m_sComment;
    DWORD                      m_dwFlags;
    DWORD                      m_dwPriority;
    std::wstring               m_sAccountName;
    winstd::sanitizing_wstring m_sPassword;
    WORD                       m_wIdleMinutes;
    WORD                       m_wDeadlineMinutes;
    DWORD                      m_dwMaxRuntimeMS;

    std::list<TASK_TRIGGER>    m_lTriggers;
};


////////////////////////////////////////////////////////////////////////////
// COpTaskDelete
////////////////////////////////////////////////////////////////////////////

class COpTaskDelete : public COpTypeSingleString
{
public:
    COpTaskDelete(LPCWSTR pszTaskName = L"", int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// COpTaskEnable
////////////////////////////////////////////////////////////////////////////

class COpTaskEnable : public COpTypeSingleString
{
public:
    COpTaskEnable(LPCWSTR pszTaskName = L"", BOOL bEnable = TRUE, int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);

    friend inline BOOL operator <<(CStream &f, const COpTaskEnable &op);
    friend inline BOOL operator >>(CStream &f, COpTaskEnable &op);

protected:
    BOOL m_bEnable;
};


////////////////////////////////////////////////////////////////////////////
// COpTaskCopy
////////////////////////////////////////////////////////////////////////////

class COpTaskCopy : public COpTypeSrcDstString
{
public:
    COpTaskCopy(LPCWSTR pszTaskSrc = L"", LPCWSTR pszTaskDst = L"", int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// COpCertStore
////////////////////////////////////////////////////////////////////////////

class COpCertStore : public COpTypeSingleString
{
public:
    COpCertStore(LPCWSTR pszStore = L"", DWORD dwEncodingType = 0, DWORD dwFlags = 0, int iTicks = 0);

    friend inline BOOL operator <<(CStream &f, const COpCertStore &op);
    friend inline BOOL operator >>(CStream &f, COpCertStore &op);

protected:
    DWORD m_dwEncodingType;
    DWORD m_dwFlags;
};


////////////////////////////////////////////////////////////////////////////
// COpCert
////////////////////////////////////////////////////////////////////////////

class COpCert : public COpCertStore
{
public:
    COpCert(LPCVOID lpCert = NULL, SIZE_T nSize = 0, LPCWSTR pszStore = L"", DWORD dwEncodingType = 0, DWORD dwFlags = 0, int iTicks = 0);

    friend inline BOOL operator <<(CStream &f, const COpCert &op);
    friend inline BOOL operator >>(CStream &f, COpCert &op);

protected:
    std::vector<BYTE> m_binCert;
};


////////////////////////////////////////////////////////////////////////////
// COpCertInstall
////////////////////////////////////////////////////////////////////////////

class COpCertInstall : public COpCert
{
public:
    COpCertInstall(LPCVOID lpCert = NULL, SIZE_T nSize = 0, LPCWSTR pszStore = L"", DWORD dwEncodingType = 0, DWORD dwFlags = 0, int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// COpCertRemove
////////////////////////////////////////////////////////////////////////////

class COpCertRemove : public COpCert
{
public:
    COpCertRemove(LPCVOID lpCert = NULL, SIZE_T nSize = 0, LPCWSTR pszStore = L"", DWORD dwEncodingType = 0, DWORD dwFlags = 0, int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// COpSvcSetStart
////////////////////////////////////////////////////////////////////////////

class COpSvcSetStart : public COpTypeSingleString
{
public:
    COpSvcSetStart(LPCWSTR pszService = L"", DWORD dwStartType = SERVICE_DEMAND_START, int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);

    friend inline BOOL operator <<(CStream &f, const COpSvcSetStart &op);
    friend inline BOOL operator >>(CStream &f, COpSvcSetStart &op);

protected:
    DWORD m_dwStartType;
};


////////////////////////////////////////////////////////////////////////////
// COpSvcControl
////////////////////////////////////////////////////////////////////////////

class COpSvcControl : public COpTypeSingleString
{
public:
    COpSvcControl(LPCWSTR pszService = L"", BOOL bWait = FALSE, int iTicks = 0);

    static DWORD WaitForState(CSession *pSession, SC_HANDLE hService, DWORD dwPendingState, DWORD dwFinalState);

    friend inline BOOL operator <<(CStream &f, const COpSvcControl &op);
    friend inline BOOL operator >>(CStream &f, COpSvcControl &op);

protected:
    BOOL m_bWait;
};


////////////////////////////////////////////////////////////////////////////
// COpSvcStart
////////////////////////////////////////////////////////////////////////////

class COpSvcStart : public COpSvcControl
{
public:
    COpSvcStart(LPCWSTR pszService = L"", BOOL bWait = FALSE, int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// COpSvcStop
////////////////////////////////////////////////////////////////////////////

class COpSvcStop : public COpSvcControl
{
public:
    COpSvcStop(LPCWSTR pszService = L"", BOOL bWait = FALSE, int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// COpWLANProfile
////////////////////////////////////////////////////////////////////////////

class COpWLANProfile : public COpTypeSingleString
{
public:
    COpWLANProfile(const GUID &guidInterface = GUID_NULL, LPCWSTR pszProfileName = L"", int iTicks = 0);

    friend inline BOOL operator <<(CStream &f, const COpWLANProfile &op);
    friend inline BOOL operator >>(CStream &f, COpWLANProfile &op);

protected:
    GUID m_guidInterface;
};


////////////////////////////////////////////////////////////////////////////
// COpWLANProfileDelete
////////////////////////////////////////////////////////////////////////////

class COpWLANProfileDelete : public COpWLANProfile
{
public:
    COpWLANProfileDelete(const GUID &guidInterface = GUID_NULL, LPCWSTR pszProfileName = L"", int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// COpWLANProfileSet
////////////////////////////////////////////////////////////////////////////

class COpWLANProfileSet : public COpWLANProfile
{
public:
    COpWLANProfileSet(const GUID &guidInterface = GUID_NULL, DWORD dwFlags = 0, LPCWSTR pszProfileName = L"", LPCWSTR pszProfileXML = L"", int iTicks = 0);
    virtual HRESULT Execute(CSession *pSession);

    friend inline BOOL operator <<(CStream &f, const COpWLANProfileSet &op);
    friend inline BOOL operator >>(CStream &f, COpWLANProfileSet &op);

protected:
    DWORD m_dwFlags;
    std::wstring m_sProfileXML;
};


////////////////////////////////////////////////////////////////////////////
// COpList
////////////////////////////////////////////////////////////////////////////

class COpList : public COperation, public std::list<std::unique_ptr<COperation> >
{
public:
    COpList(int iTicks = 0);

    inline void push_front(COperation *pOp);
    inline void push_back(COperation *pOp);
    DWORD LoadFromFile(LPCTSTR pszFileName);
    DWORD SaveToFile(LPCTSTR pszFileName) const;

    virtual HRESULT Execute(CSession *pSession);

    friend inline BOOL operator <<(CStream &f, const COpList &list);
    friend inline BOOL operator >>(CStream &f, COpList &list);

public:
    enum OPERATION {
        OP_ROLLBACK_ENABLE = 1,
        OP_FILE_DELETE,
        OP_FILE_MOVE,
        OP_REG_KEY_CREATE,
        OP_REG_KEY_COPY,
        OP_REG_KEY_DELETE,
        OP_REG_VALUE_CREATE,
        OP_REG_VALUE_COPY,
        OP_REG_VALUE_DELETE,
        OP_TASK_CREATE,
        OP_TASK_DELETE,
        OP_TASK_ENABLE,
        OP_TASK_COPY,
        OP_CERT_INSTALL,
        OP_CERT_REMOVE,
        OP_SVC_SET_START,
        OP_SVC_START,
        OP_SVC_STOP,
        OP_WLAN_PROFILE_DELETE,
        OP_WLAN_PROFILE_SET,
        OP_SUBLIST
    };

protected:
    template <class T, enum OPERATION ID> inline static BOOL write(CStream &f, const COperation *p);
    template <class T> inline BOOL read_back(CStream &f);
};


////////////////////////////////////////////////////////////////////////////
// CSession
////////////////////////////////////////////////////////////////////////////

class CSession
{
public:
    CSession();

    MSIHANDLE m_hInstall;        // Installer handle
    BOOL m_bContinueOnError;     // Continue execution on operation error?
    BOOL m_bRollbackEnabled;     // Is rollback enabled?
    COpList m_olRollback;        // Rollback operation list
    COpList m_olCommit;          // Commit operation list
};


////////////////////////////////////////////////////////////////////////////
// Helper functions
////////////////////////////////////////////////////////////////////////////

UINT SaveSequence(MSIHANDLE hInstall, LPCTSTR szActionExecute, LPCTSTR szActionCommit, LPCTSTR szActionRollback, const COpList &olExecute);
UINT ExecuteSequence(MSIHANDLE hInstall);

} // namespace MSICA


////////////////////////////////////////////////////////////////////
// Local includes
////////////////////////////////////////////////////////////////////

#include <WinStd/MSI.h>

#include <memory>

#include <msiquery.h>
#include <mstask.h>
#include <tchar.h>


////////////////////////////////////////////////////////////////////
// Inline helper functions
////////////////////////////////////////////////////////////////////

template<class _Elem, class _Traits, class _Ax>
inline UINT MsiRecordFormatStringA(MSIHANDLE hInstall, MSIHANDLE hRecord, unsigned int iField, std::basic_string<_Elem, _Traits, _Ax> &sValue)
{
    UINT uiResult;
    PMSIHANDLE hRecordEx;

    // Read string to format.
    uiResult = ::MsiRecordGetStringA(hRecord, iField, sValue);
    if (uiResult != NO_ERROR) return uiResult;

    // If the string is empty, there's nothing left to do.
    if (sValue.empty()) return NO_ERROR;

    // Create a record.
    hRecordEx = ::MsiCreateRecord(1);
    if (!hRecordEx) return ERROR_INVALID_HANDLE;

    // Populate record with data.
    uiResult = ::MsiRecordSetStringA(hRecordEx, 0, sValue.c_str());
    if (uiResult != NO_ERROR) return uiResult;

    // Do the formatting.
    return ::MsiFormatRecordA(hInstall, hRecordEx, sValue);
}


template<class _Elem, class _Traits, class _Ax>
inline UINT MsiRecordFormatStringW(MSIHANDLE hInstall, MSIHANDLE hRecord, unsigned int iField, std::basic_string<_Elem, _Traits, _Ax> &sValue)
{
    UINT uiResult;
    PMSIHANDLE hRecordEx;

    // Read string to format.
    uiResult = ::MsiRecordGetStringW(hRecord, iField, sValue);
    if (uiResult != NO_ERROR) return uiResult;

    // If the string is empty, there's nothing left to do.
    if (sValue.empty()) return NO_ERROR;

    // Create a record.
    hRecordEx = ::MsiCreateRecord(1);
    if (!hRecordEx) return ERROR_INVALID_HANDLE;

    // Populate record with data.
    uiResult = ::MsiRecordSetStringW(hRecordEx, 0, sValue.c_str());
    if (uiResult != NO_ERROR) return uiResult;

    // Do the formatting.
    return ::MsiFormatRecordW(hInstall, hRecordEx, sValue);
}

#ifdef UNICODE
#define MsiRecordFormatString  MsiRecordFormatStringW
#else
#define MsiRecordFormatString  MsiRecordFormatStringA
#endif // !UNICODE


namespace MSICA {

////////////////////////////////////////////////////////////////////////////
// Inline operators
////////////////////////////////////////////////////////////////////////////

#define MSICALIB_STREAM_DATA(T)                                           \
inline BOOL operator <<(CStream &f, const T &val)                         \
{                                                                         \
    DWORD dwWritten;                                                      \
                                                                          \
    if (!::WriteFile(f, &val, sizeof(T), &dwWritten, NULL)) return FALSE; \
    if (dwWritten != sizeof(T)) {                                         \
        ::SetLastError(ERROR_WRITE_FAULT);                                \
        return FALSE;                                                     \
    }                                                                     \
                                                                          \
    return TRUE;                                                          \
}                                                                         \
                                                                          \
inline BOOL operator >>(CStream &f, T &val)                               \
{                                                                         \
    DWORD dwRead;                                                         \
                                                                          \
    if (!::ReadFile(f, &val, sizeof(T), &dwRead, NULL)) return FALSE;     \
    if (dwRead != sizeof(T)) {                                            \
         ::SetLastError(ERROR_READ_FAULT);                                \
         return FALSE;                                                    \
    }                                                                     \
                                                                          \
    return TRUE;                                                          \
}

MSICALIB_STREAM_DATA(short)
MSICALIB_STREAM_DATA(unsigned short)
MSICALIB_STREAM_DATA(long)
MSICALIB_STREAM_DATA(unsigned long)
MSICALIB_STREAM_DATA(int)
MSICALIB_STREAM_DATA(size_t)
#ifndef _WIN64
MSICALIB_STREAM_DATA(DWORDLONG)
#endif
MSICALIB_STREAM_DATA(HKEY)
MSICALIB_STREAM_DATA(GUID)
MSICALIB_STREAM_DATA(TASK_TRIGGER)
MSICALIB_STREAM_DATA(COpList::OPERATION)

#undef MSICALIB_STREAM_DATA

template <class _Ty, class _Ax>
inline BOOL operator <<(CStream &f, const std::vector<_Ty, _Ax> &val)
{
    // Write element count.
    size_t nCount = val.size();
    if (!(f << nCount)) return FALSE;

    // Write data (opaque).
    DWORD dwWritten;
    if (!::WriteFile(f, val.data(), static_cast<DWORD>(sizeof(_Ty) * nCount), &dwWritten, NULL)) return FALSE;
    if (dwWritten != sizeof(_Ty) * nCount) {
        ::SetLastError(ERROR_WRITE_FAULT);
        return FALSE;
    }

    return TRUE;
}


template <class _Ty, class _Ax>
inline BOOL operator >>(CStream &f, std::vector<_Ty, _Ax> &val)
{
    // Read element count.
    size_t nCount;
    if (!(f >> nCount)) return FALSE;

    // Allocate the buffer.
    val.resize(nCount);

    // Read data (opaque).
    DWORD dwRead;
    if (!::ReadFile(f, val.data(), static_cast<DWORD>(sizeof(_Ty) * nCount), &dwRead, NULL)) return FALSE;
    if (dwRead != sizeof(_Ty) * nCount) {
         ::SetLastError(ERROR_READ_FAULT);
         return FALSE;
    }

    return TRUE;
}


template<class _Elem, class _Traits, class _Ax>
inline BOOL operator <<(CStream &f, const std::basic_string<_Elem, _Traits, _Ax> &val)
{
    // Write string length (in characters).
    size_t iLength = val.length();
    if (!(f << iLength)) return FALSE;

    // Write string data (without terminator).
    DWORD dwWritten;
    if (!::WriteFile(f, val.c_str(), static_cast<DWORD>(sizeof(_Elem) * iLength), &dwWritten, NULL)) return FALSE;
    if (dwWritten != sizeof(_Elem) * iLength) {
        ::SetLastError(ERROR_WRITE_FAULT);
        return FALSE;
    }

    return TRUE;
}


template<class _Elem, class _Traits, class _Ax>
inline BOOL operator >>(CStream &f, std::basic_string<_Elem, _Traits, _Ax> &val)
{
    // Read string length (in characters).
    size_t iLength;
    if (!(f >> iLength)) return FALSE;

    // Allocate the buffer.
    std::unique_ptr<_Elem[]> buf(new _Elem[iLength]);
    if (!buf) {
        ::SetLastError(ERROR_OUTOFMEMORY);
        return FALSE;
    }

    // Read string data (without terminator).
    DWORD dwRead;
    if (!::ReadFile(f, buf.get(), static_cast<DWORD>(sizeof(_Elem) * iLength), &dwRead, NULL)) return FALSE;
    if (dwRead != sizeof(_Elem) * iLength) {
         ::SetLastError(ERROR_READ_FAULT);
         return FALSE;
    }
    val.assign(buf.get(), iLength);

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COperation &op)
{
    return f << op.m_iTicks;
}


inline BOOL operator >>(CStream &f, COperation &op)
{
    return f >> op.m_iTicks;
}


inline BOOL operator <<(CStream &f, const COpTypeSingleString &op)
{
    if (!(f << (const COperation &)op)) return FALSE;
    if (!(f << op.m_sValue           )) return FALSE;

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpTypeSingleString &op)
{
    if (!(f >> (COperation &)op)) return FALSE;
    if (!(f >> op.m_sValue     )) return FALSE;

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpTypeSrcDstString &op)
{
    if (!(f << (const COperation &)op)) return FALSE;
    if (!(f << op.m_sValue1          )) return FALSE;
    if (!(f << op.m_sValue2          )) return FALSE;

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpTypeSrcDstString &op)
{
    if (!(f >> (COperation &)op)) return FALSE;
    if (!(f >> op.m_sValue1    )) return FALSE;
    if (!(f >> op.m_sValue2    )) return FALSE;

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpTypeBoolean &op)
{
    if (!(f << (const COperation &)op)) return FALSE;
    if (!(f << op.m_bValue           )) return FALSE;

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpTypeBoolean &op)
{
    if (!(f >> (COperation &)op)) return FALSE;
    if (!(f >> op.m_bValue     )) return FALSE;

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpRegKeySingle &op)
{
    if (!(f << (const COpTypeSingleString &)op)) return FALSE;
    if (!(f << op.m_hKeyRoot                  )) return FALSE;

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpRegKeySingle &op)
{
    if (!(f >> (COpTypeSingleString &)op)) return FALSE;
    if (!(f >> op.m_hKeyRoot            )) return FALSE;

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpRegKeySrcDst &op)
{
    if (!(f << (const COpTypeSrcDstString &)op)) return FALSE;
    if (!(f << op.m_hKeyRoot                  )) return FALSE;

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpRegKeySrcDst &op)
{
    if (!(f >> (COpTypeSrcDstString &)op)) return FALSE;
    if (!(f >> op.m_hKeyRoot            )) return FALSE;

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpRegValueSingle &op)
{
    if (!(f << (const COpRegKeySingle &)op)) return FALSE;
    if (!(f << op.m_sValueName            )) return FALSE;

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpRegValueSingle &op)
{
    if (!(f >> (COpRegKeySingle &)op)) return FALSE;
    if (!(f >> op.m_sValueName      )) return FALSE;

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpRegValueSrcDst &op)
{
    if (!(f << (const COpRegKeySingle &)op)) return FALSE;
    if (!(f << op.m_sValueName1           )) return FALSE;
    if (!(f << op.m_sValueName2           )) return FALSE;

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpRegValueSrcDst &op)
{
    if (!(f >> (COpRegKeySingle &)op)) return FALSE;
    if (!(f >> op.m_sValueName1     )) return FALSE;
    if (!(f >> op.m_sValueName2     )) return FALSE;

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpRegValueCreate &op)
{
    if (!(f << (const COpRegValueSingle &)op)) return FALSE;
    if (!(f << op.m_dwType                  )) return FALSE;
    switch (op.m_dwType) {
    case REG_SZ:
    case REG_EXPAND_SZ:
    case REG_LINK:                if (!(f << op.m_sData  )) return FALSE; break;
    case REG_BINARY:              if (!(f << op.m_binData)) return FALSE; break;
    case REG_DWORD_LITTLE_ENDIAN:
    case REG_DWORD_BIG_ENDIAN:    if (!(f << op.m_dwData )) return FALSE; break;
    case REG_MULTI_SZ:            if (!(f << op.m_szData )) return FALSE; break;
    case REG_QWORD_LITTLE_ENDIAN: if (!(f << op.m_qwData )) return FALSE; break;
    default: ::SetLastError(ERROR_INVALID_DATA); return FALSE;
    }

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpRegValueCreate &op)
{
    if (!(f >> (COpRegValueSingle &)op)) return FALSE;
    if (!(f >> op.m_dwType            )) return FALSE;
    switch (op.m_dwType) {
    case REG_SZ:
    case REG_EXPAND_SZ:
    case REG_LINK:                if (!(f >> op.m_sData  )) return FALSE; break;
    case REG_BINARY:              if (!(f >> op.m_binData)) return FALSE; break;
    case REG_DWORD_LITTLE_ENDIAN:
    case REG_DWORD_BIG_ENDIAN:    if (!(f >> op.m_dwData )) return FALSE; break;
    case REG_MULTI_SZ:            if (!(f >> op.m_szData )) return FALSE; break;
    case REG_QWORD_LITTLE_ENDIAN: if (!(f >> op.m_qwData )) return FALSE; break;
    default:                      ::SetLastError(ERROR_INVALID_DATA); return FALSE;
    }

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpTaskCreate &op)
{
    if (!(f << (const COpTypeSingleString&)op)) return FALSE;
    if (!(f << op.m_sApplicationName         )) return FALSE;
    if (!(f << op.m_sParameters              )) return FALSE;
    if (!(f << op.m_sWorkingDirectory        )) return FALSE;
    if (!(f << op.m_sAuthor                  )) return FALSE;
    if (!(f << op.m_sComment                 )) return FALSE;
    if (!(f << op.m_dwFlags                  )) return FALSE;
    if (!(f << op.m_dwPriority               )) return FALSE;
    if (!(f << op.m_sAccountName             )) return FALSE;
    if (!(f << op.m_sPassword                )) return FALSE;
    if (!(f << op.m_wDeadlineMinutes         )) return FALSE;
    if (!(f << op.m_wIdleMinutes             )) return FALSE;
    if (!(f << op.m_dwMaxRuntimeMS           )) return FALSE;
    if (!(f << op.m_lTriggers.size()         )) return FALSE;
    for (auto t = op.m_lTriggers.cbegin(), t_end = op.m_lTriggers.cend(); t != t_end; ++t)
        if (!(f << *t)) return FALSE;

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpTaskCreate &op)
{
    size_t nCount;

    if (!(f >> (COpTypeSingleString&)op)) return FALSE;
    if (!(f >> op.m_sApplicationName   )) return FALSE;
    if (!(f >> op.m_sParameters        )) return FALSE;
    if (!(f >> op.m_sWorkingDirectory  )) return FALSE;
    if (!(f >> op.m_sAuthor            )) return FALSE;
    if (!(f >> op.m_sComment           )) return FALSE;
    if (!(f >> op.m_dwFlags            )) return FALSE;
    if (!(f >> op.m_dwPriority         )) return FALSE;
    if (!(f >> op.m_sAccountName       )) return FALSE;
    if (!(f >> op.m_sPassword          )) return FALSE;
    if (!(f >> op.m_wDeadlineMinutes   )) return FALSE;
    if (!(f >> op.m_wIdleMinutes       )) return FALSE;
    if (!(f >> op.m_dwMaxRuntimeMS     )) return FALSE;
    if (!(f >> nCount                  )) return FALSE;
    op.m_lTriggers.clear();
    while (nCount--) {
        TASK_TRIGGER ttData;
        if (!(f >> ttData)) return FALSE;
        op.m_lTriggers.push_back(std::move(ttData));
    }

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpTaskEnable &op)
{
    if (!(f << (const COpTypeSingleString&)op)) return FALSE;
    if (!(f << op.m_bEnable                  )) return FALSE;

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpTaskEnable &op)
{
    if (!(f >> (COpTypeSingleString&)op)) return FALSE;
    if (!(f >> op.m_bEnable            )) return FALSE;

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpCertStore &op)
{
    if (!(f << (const COpTypeSingleString&)op)) return FALSE;
    if (!(f << op.m_dwEncodingType           )) return FALSE;
    if (!(f << op.m_dwFlags                  )) return FALSE;

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpCertStore &op)
{
    if (!(f >> (COpTypeSingleString&)op)) return FALSE;
    if (!(f >> op.m_dwEncodingType     )) return FALSE;
    if (!(f >> op.m_dwFlags            )) return FALSE;

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpCert &op)
{
    if (!(f << (const COpCertStore&)op)) return FALSE;
    if (!(f << op.m_binCert           )) return FALSE;

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpCert &op)
{
    if (!(f >> (COpCertStore&)op)) return FALSE;
    if (!(f >> op.m_binCert     )) return FALSE;

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpSvcSetStart &op)
{
    if (!(f << (const COpTypeSingleString&)op)) return FALSE;
    if (!(f << op.m_dwStartType              )) return FALSE;

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpSvcSetStart &op)
{
    if (!(f >> (COpTypeSingleString&)op)) return FALSE;
    if (!(f >> op.m_dwStartType        )) return FALSE;

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpSvcControl &op)
{
    if (!(f << (const COpTypeSingleString&)op)) return FALSE;
    if (!(f << op.m_bWait                    )) return FALSE;

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpSvcControl &op)
{
    if (!(f >> (COpTypeSingleString&)op)) return FALSE;
    if (!(f >> op.m_bWait              )) return FALSE;

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpWLANProfile &op)
{
    if (!(f << (const COpTypeSingleString&)op)) return FALSE;
    if (!(f << op.m_guidInterface            )) return FALSE;

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpWLANProfile &op)
{
    if (!(f >> (COpTypeSingleString&)op)) return FALSE;
    if (!(f >> op.m_guidInterface      )) return FALSE;

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpWLANProfileSet &op)
{
    if (!(f << (const COpWLANProfile&)op)) return FALSE;
    if (!(f << op.m_dwFlags             )) return FALSE;
    if (!(f << op.m_sProfileXML         )) return FALSE;

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpWLANProfileSet &op)
{
    if (!(f >> (COpWLANProfile&)op)) return FALSE;
    if (!(f >> op.m_dwFlags       )) return FALSE;
    if (!(f >> op.m_sProfileXML   )) return FALSE;

    return TRUE;
}


inline BOOL operator <<(CStream &f, const COpList &list)
{
    if (!(f << (const COperation &)list)) return FALSE;
    if (!(f << list.size()             )) return FALSE;
    for (auto o = list.cbegin(), o_end = list.cend(); o != o_end; ++o) {
        BOOL bResult;
        const COperation *pOp = o->get();
             if (dynamic_cast<const COpRollbackEnable*   >(pOp)) bResult = list.write<COpRollbackEnable,    COpList::OP_ROLLBACK_ENABLE    >(f, pOp);
        else if (dynamic_cast<const COpFileDelete*       >(pOp)) bResult = list.write<COpFileDelete,        COpList::OP_FILE_DELETE        >(f, pOp);
        else if (dynamic_cast<const COpFileMove*         >(pOp)) bResult = list.write<COpFileMove,          COpList::OP_FILE_MOVE          >(f, pOp);
        else if (dynamic_cast<const COpRegKeyCreate*     >(pOp)) bResult = list.write<COpRegKeyCreate,      COpList::OP_REG_KEY_CREATE     >(f, pOp);
        else if (dynamic_cast<const COpRegKeyCopy*       >(pOp)) bResult = list.write<COpRegKeyCopy,        COpList::OP_REG_KEY_COPY       >(f, pOp);
        else if (dynamic_cast<const COpRegKeyDelete*     >(pOp)) bResult = list.write<COpRegKeyDelete,      COpList::OP_REG_KEY_DELETE     >(f, pOp);
        else if (dynamic_cast<const COpRegValueCreate*   >(pOp)) bResult = list.write<COpRegValueCreate,    COpList::OP_REG_VALUE_CREATE   >(f, pOp);
        else if (dynamic_cast<const COpRegValueCopy*     >(pOp)) bResult = list.write<COpRegValueCopy,      COpList::OP_REG_VALUE_COPY     >(f, pOp);
        else if (dynamic_cast<const COpRegValueDelete*   >(pOp)) bResult = list.write<COpRegValueDelete,    COpList::OP_REG_VALUE_DELETE   >(f, pOp);
        else if (dynamic_cast<const COpTaskCreate*       >(pOp)) bResult = list.write<COpTaskCreate,        COpList::OP_TASK_CREATE        >(f, pOp);
        else if (dynamic_cast<const COpTaskDelete*       >(pOp)) bResult = list.write<COpTaskDelete,        COpList::OP_TASK_DELETE        >(f, pOp);
        else if (dynamic_cast<const COpTaskEnable*       >(pOp)) bResult = list.write<COpTaskEnable,        COpList::OP_TASK_ENABLE        >(f, pOp);
        else if (dynamic_cast<const COpTaskCopy*         >(pOp)) bResult = list.write<COpTaskCopy,          COpList::OP_TASK_COPY          >(f, pOp);
        else if (dynamic_cast<const COpCertInstall*      >(pOp)) bResult = list.write<COpCertInstall,       COpList::OP_CERT_INSTALL       >(f, pOp);
        else if (dynamic_cast<const COpCertRemove*       >(pOp)) bResult = list.write<COpCertRemove,        COpList::OP_CERT_REMOVE        >(f, pOp);
        else if (dynamic_cast<const COpSvcSetStart*      >(pOp)) bResult = list.write<COpSvcSetStart,       COpList::OP_SVC_SET_START      >(f, pOp);
        else if (dynamic_cast<const COpSvcStart*         >(pOp)) bResult = list.write<COpSvcStart,          COpList::OP_SVC_START          >(f, pOp);
        else if (dynamic_cast<const COpSvcStop*          >(pOp)) bResult = list.write<COpSvcStop,           COpList::OP_SVC_STOP           >(f, pOp);
        else if (dynamic_cast<const COpWLANProfileDelete*>(pOp)) bResult = list.write<COpWLANProfileDelete, COpList::OP_WLAN_PROFILE_DELETE>(f, pOp);
        else if (dynamic_cast<const COpWLANProfileSet*   >(pOp)) bResult = list.write<COpWLANProfileSet,    COpList::OP_WLAN_PROFILE_SET   >(f, pOp);
        else if (dynamic_cast<const COpList*             >(pOp)) bResult = list.write<COpList,              COpList::OP_SUBLIST            >(f, pOp);
        else {
            // Unsupported type of operation.
            ::SetLastError(ERROR_INVALID_DATA);
            return FALSE;
        }

        if (!bResult) return FALSE;
    }

    return TRUE;
}


inline BOOL operator >>(CStream &f, COpList &list)
{
    size_t nCount;

    if (!(f >> (COperation &)list)) return FALSE;
    if (!(f >> nCount            )) return FALSE;
    list.clear();
    while (nCount--) {
        BOOL bResult;
        COpList::OPERATION opCode;

        if (!(f >> opCode)) return FALSE;

        switch (opCode) {
        case COpList::OP_ROLLBACK_ENABLE:     bResult = list.read_back<COpRollbackEnable   >(f); break;
        case COpList::OP_FILE_DELETE:         bResult = list.read_back<COpFileDelete       >(f); break;
        case COpList::OP_FILE_MOVE:           bResult = list.read_back<COpFileMove         >(f); break;
        case COpList::OP_REG_KEY_CREATE:      bResult = list.read_back<COpRegKeyCreate     >(f); break;
        case COpList::OP_REG_KEY_COPY:        bResult = list.read_back<COpRegKeyCopy       >(f); break;
        case COpList::OP_REG_KEY_DELETE:      bResult = list.read_back<COpRegKeyDelete     >(f); break;
        case COpList::OP_REG_VALUE_CREATE:    bResult = list.read_back<COpRegValueCreate   >(f); break;
        case COpList::OP_REG_VALUE_COPY:      bResult = list.read_back<COpRegValueCopy     >(f); break;
        case COpList::OP_REG_VALUE_DELETE:    bResult = list.read_back<COpRegValueDelete   >(f); break;
        case COpList::OP_TASK_CREATE:         bResult = list.read_back<COpTaskCreate       >(f); break;
        case COpList::OP_TASK_DELETE:         bResult = list.read_back<COpTaskDelete       >(f); break;
        case COpList::OP_TASK_ENABLE:         bResult = list.read_back<COpTaskEnable       >(f); break;
        case COpList::OP_TASK_COPY:           bResult = list.read_back<COpTaskCopy         >(f); break;
        case COpList::OP_CERT_INSTALL:        bResult = list.read_back<COpCertInstall      >(f); break;
        case COpList::OP_CERT_REMOVE:         bResult = list.read_back<COpCertRemove       >(f); break;
        case COpList::OP_SVC_SET_START:       bResult = list.read_back<COpSvcSetStart      >(f); break;
        case COpList::OP_SVC_START:           bResult = list.read_back<COpSvcStart         >(f); break;
        case COpList::OP_SVC_STOP:            bResult = list.read_back<COpSvcStop          >(f); break;
        case COpList::OP_WLAN_PROFILE_DELETE: bResult = list.read_back<COpWLANProfileDelete>(f); break;
        case COpList::OP_WLAN_PROFILE_SET:    bResult = list.read_back<COpWLANProfileSet   >(f); break;
        case COpList::OP_SUBLIST:             bResult = list.read_back<COpList             >(f); break;
        default:
            // Unsupported type of operation.
            ::SetLastError(ERROR_INVALID_DATA);
            return FALSE;
        }

        if (!bResult) return FALSE;
    }

    return TRUE;
}


////////////////////////////////////////////////////////////////////
// Inline Functions
////////////////////////////////////////////////////////////////////


inline BOOL IsWow64Process()
{
#ifndef _WIN64
    // Find IsWow64Process() address in KERNEL32.DLL.
    BOOL (WINAPI *_IsWow64Process)(__in HANDLE hProcess, __out PBOOL Wow64Process) = (BOOL(WINAPI*)(__in HANDLE, __out PBOOL))::GetProcAddress(::GetModuleHandle(_T("KERNEL32.DLL")), "IsWow64Process");

    // See if our 32-bit process is running in 64-bit environment.
    if (_IsWow64Process) {
        BOOL bResult;

        // See, what IsWow64Process() says about current process.
        if (_IsWow64Process(::GetCurrentProcess(), &bResult)) {
            // Return result.
            return bResult;
        } else {
            // IsWow64Process() returned an error. Assume, the process is not WOW64.
            return FALSE;
        }
    } else {
        // This KERNEL32.DLL doesn't know IsWow64Process().Definitely not a WOW64 process.
        return FALSE;
    }
#else
    // 64-bit processes are never run as WOW64.
    return FALSE;
#endif
}


////////////////////////////////////////////////////////////////////////////
// Inline methods
////////////////////////////////////////////////////////////////////////////

inline void COpList::push_front(COperation *pOp)
{
    std::list<std::unique_ptr<COperation> >::push_front(std::move(std::unique_ptr<COperation>(pOp)));
}


inline void COpList::push_back(COperation *pOp)
{
    std::list<std::unique_ptr<COperation> >::push_back(std::move(std::unique_ptr<COperation>(pOp)));
}


template <class T, enum COpList::OPERATION ID> inline static BOOL COpList::write(CStream &f, const COperation *p)
{
    const T *pp = dynamic_cast<const T*>(p);
    if (!pp) {
        // Wrong type.
        ::SetLastError(ERROR_INVALID_DATA);
        return FALSE;
    }

    if (!(f << ID )) return FALSE;
    if (!(f << *pp)) return FALSE;

    return TRUE;
}


template <class T> inline BOOL COpList::read_back(CStream &f)
{
    // Create element.
    std::unique_ptr<T> p(new T());
    if (!p) {
        ::SetLastError(ERROR_OUTOFMEMORY);
        return FALSE;
    }

    // Load element from file.
    if (!(f >> *p)) return FALSE;

    // Add element.
    std::list<std::unique_ptr<COperation> >::push_back(std::move(p));
    return TRUE;
}

} // namespace MSICA
