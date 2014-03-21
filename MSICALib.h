#ifndef __MSICALib_H__
#define __MSICALib_H__

#include <atlbase.h>
#include <atlcoll.h>
#include <atlfile.h>
#include <atlstr.h>
#include <msi.h>
#include <mstask.h>
#include <wincrypt.h>
#include <windows.h>


////////////////////////////////////////////////////////////////////
// Error codes (next unused 2574L)
////////////////////////////////////////////////////////////////////

#define ERROR_INSTALL_DATABASE_OPEN                    2550L
#define ERROR_INSTALL_OPLIST_CREATE                    2551L
#define ERROR_INSTALL_PROPERTY_SET                     2553L
#define ERROR_INSTALL_SCRIPT_WRITE                     2552L
#define ERROR_INSTALL_SCRIPT_READ                      2560L
#define ERROR_INSTALL_FILE_DELETE_FAILED               2554L
#define ERROR_INSTALL_FILE_MOVE_FAILED                 2555L
#define ERROR_INSTALL_REGKEY_CREATE_FAILED             2561L
#define ERROR_INSTALL_REGKEY_COPY_FAILED               2562L
#define ERROR_INSTALL_REGKEY_PROBING_FAILED            2563L
#define ERROR_INSTALL_REGKEY_DELETE_FAILED             2564L
#define ERROR_INSTALL_REGKEY_SETVALUE_FAILED           2565L
#define ERROR_INSTALL_REGKEY_DELETEVALUE_FAILED        2567L
#define ERROR_INSTALL_REGKEY_COPYVALUE_FAILED          2568L
#define ERROR_INSTALL_REGKEY_PROBINGVAL_FAILED         2566L
#define ERROR_INSTALL_TASK_CREATE_FAILED               2556L
#define ERROR_INSTALL_TASK_DELETE_FAILED               2557L
#define ERROR_INSTALL_TASK_ENABLE_FAILED               2558L
#define ERROR_INSTALL_TASK_COPY_FAILED                 2559L
#define ERROR_INSTALL_CERT_INSTALL_FAILED              2569L
#define ERROR_INSTALL_CERT_REMOVE_FAILED               2570L
#define ERROR_INSTALL_SVC_SET_START_FAILED             2571L
#define ERROR_INSTALL_SVC_START_FAILED                 2572L
#define ERROR_INSTALL_SVC_STOP_FAILED                  2573L


namespace MSICA {


////////////////////////////////////////////////////////////////////////////
// Forward declarations
////////////////////////////////////////////////////////////////////////////

class CSession;


////////////////////////////////////////////////////////////////////////////
// COperation
////////////////////////////////////////////////////////////////////////////

class COperation
{
public:
    COperation(int iTicks = 0);

    virtual HRESULT Execute(CSession *pSession) = 0;

    friend class COpList;
    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COperation &op);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COperation &op);

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

    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COpTypeSingleString &op);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COpTypeSingleString &op);

protected:
    ATL::CAtlStringW m_sValue;
};


////////////////////////////////////////////////////////////////////////////
// COpTypeSrcDstString
////////////////////////////////////////////////////////////////////////////

class COpTypeSrcDstString : public COperation
{
public:
    COpTypeSrcDstString(LPCWSTR pszValue1 = L"", LPCWSTR pszValue2 = L"", int iTicks = 0);

    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COpTypeSrcDstString &op);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COpTypeSrcDstString &op);

protected:
    ATL::CAtlStringW m_sValue1;
    ATL::CAtlStringW m_sValue2;
};


////////////////////////////////////////////////////////////////////////////
// COpTypeBoolean
////////////////////////////////////////////////////////////////////////////

class COpTypeBoolean : public COperation
{
public:
    COpTypeBoolean(BOOL bValue = TRUE, int iTicks = 0);

    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COpTypeBoolean &op);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COpTypeBoolean &op);

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

    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COpRegKeySingle &op);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COpRegKeySingle &op);

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

    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COpRegKeySrcDst &op);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COpRegKeySrcDst &op);

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

    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COpRegValueSingle &op);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COpRegValueSingle &op);

protected:
    ATL::CAtlStringW m_sValueName;
};


////////////////////////////////////////////////////////////////////////////
// COpRegValueSrcDst
////////////////////////////////////////////////////////////////////////////

class COpRegValueSrcDst : public COpRegKeySingle
{
public:
    COpRegValueSrcDst(HKEY hKeyRoot = NULL, LPCWSTR pszKeyName = L"", LPCWSTR pszValueNameSrc = L"", LPCWSTR pszValueNameDst = L"", int iTicks = 0);

    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COpRegValueSrcDst &op);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COpRegValueSrcDst &op);

protected:
    ATL::CAtlStringW m_sValueName1;
    ATL::CAtlStringW m_sValueName2;
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

    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COpRegValueCreate &op);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COpRegValueCreate &op);

protected:
    DWORD m_dwType;
    ATL::CAtlStringW m_sData;
    ATL::CAtlArray<BYTE> m_binData;
    DWORD m_dwData;
    ATL::CAtlArray<WCHAR> m_szData;
    DWORDLONG m_qwData;
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
    virtual ~COpTaskCreate();
    virtual HRESULT Execute(CSession *pSession);

    UINT SetFromRecord(MSIHANDLE hInstall, MSIHANDLE hRecord);
    UINT SetTriggersFromView(MSIHANDLE hView);

    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COpTaskCreate &op);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COpTaskCreate &op);

protected:
    ATL::CAtlStringW m_sApplicationName;
    ATL::CAtlStringW m_sParameters;
    ATL::CAtlStringW m_sWorkingDirectory;
    ATL::CAtlStringW m_sAuthor;
    ATL::CAtlStringW m_sComment;
    DWORD            m_dwFlags;
    DWORD            m_dwPriority;
    ATL::CAtlStringW m_sAccountName;
    ATL::CAtlStringW m_sPassword;
    WORD             m_wIdleMinutes;
    WORD             m_wDeadlineMinutes;
    DWORD            m_dwMaxRuntimeMS;

    ATL::CAtlList<TASK_TRIGGER> m_lTriggers;
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

    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COpTaskEnable &op);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COpTaskEnable &op);

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

    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COpCertStore &op);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COpCertStore &op);

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

    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COpCert &op);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COpCert &op);

protected:
    ATL::CAtlArray<BYTE> m_binCert;
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

    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COpSvcSetStart &op);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COpSvcSetStart &op);

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

    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COpSvcControl &op);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COpSvcControl &op);

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
// COpList
////////////////////////////////////////////////////////////////////////////

class COpList : public COperation, public ATL::CAtlList<COperation*>
{
public:
    COpList(int iTicks = 0);

    void Free();
    HRESULT LoadFromFile(LPCTSTR pszFileName);
    HRESULT SaveToFile(LPCTSTR pszFileName) const;

    virtual HRESULT Execute(CSession *pSession);

    friend inline HRESULT operator <<(ATL::CAtlFile &f, const COpList &list);
    friend inline HRESULT operator >>(ATL::CAtlFile &f, COpList &list);

protected:
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
        OP_SUBLIST
    };

protected:
    template <class T, enum OPERATION ID> inline static HRESULT Save(ATL::CAtlFile &f, const COperation *p);
    template <class T> inline HRESULT LoadAndAddTail(ATL::CAtlFile &f);
};


////////////////////////////////////////////////////////////////////////////
// CSession
////////////////////////////////////////////////////////////////////////////

class CSession
{
public:
    CSession();
    virtual ~CSession();

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

#include <atlfile.h>
#include <atlstr.h>
#include <msiquery.h>
#include <mstask.h>


////////////////////////////////////////////////////////////////////
// Inline helper functions
////////////////////////////////////////////////////////////////////

inline UINT MsiGetPropertyA(MSIHANDLE hInstall, LPCSTR szName, ATL::CAtlStringA &sValue)
{
    DWORD dwSize = 0;
    UINT uiResult;

    // Query the actual string length first.
    uiResult = ::MsiGetPropertyA(hInstall, szName, "", &dwSize);
    if (uiResult == ERROR_MORE_DATA) {
        // Prepare the buffer to read the string data into and read it.
        LPSTR szBuffer = sValue.GetBuffer(dwSize++);
        if (!szBuffer) return ERROR_OUTOFMEMORY;
        uiResult = ::MsiGetPropertyA(hInstall, szName, szBuffer, &dwSize);
        sValue.ReleaseBuffer(uiResult == NO_ERROR ? dwSize : 0);
        return uiResult;
    } else if (uiResult == NO_ERROR) {
        // The string in database is empty.
        sValue.Empty();
        return NO_ERROR;
    } else {
        // Return error code.
        return uiResult;
    }
}


inline UINT MsiGetPropertyW(MSIHANDLE hInstall, LPCWSTR szName, ATL::CAtlStringW &sValue)
{
    DWORD dwSize = 0;
    UINT uiResult;

    // Query the actual string length first.
    uiResult = ::MsiGetPropertyW(hInstall, szName, L"", &dwSize);
    if (uiResult == ERROR_MORE_DATA) {
        // Prepare the buffer to read the string data into and read it.
        LPWSTR szBuffer = sValue.GetBuffer(dwSize++);
        if (!szBuffer) return ERROR_OUTOFMEMORY;
        uiResult = ::MsiGetPropertyW(hInstall, szName, szBuffer, &dwSize);
        sValue.ReleaseBuffer(uiResult == NO_ERROR ? dwSize : 0);
        return uiResult;
    } else if (uiResult == NO_ERROR) {
        // The string in database is empty.
        sValue.Empty();
        return NO_ERROR;
    } else {
        // Return error code.
        return uiResult;
    }
}


inline UINT MsiRecordGetStringA(MSIHANDLE hRecord, unsigned int iField, ATL::CAtlStringA &sValue)
{
    DWORD dwSize = 0;
    UINT uiResult;

    // Query the actual string length first.
    uiResult = ::MsiRecordGetStringA(hRecord, iField, "", &dwSize);
    if (uiResult == ERROR_MORE_DATA) {
        // Prepare the buffer to read the string data into and read it.
        LPSTR szBuffer = sValue.GetBuffer(dwSize++);
        if (!szBuffer) return ERROR_OUTOFMEMORY;
        uiResult = ::MsiRecordGetStringA(hRecord, iField, szBuffer, &dwSize);
        sValue.ReleaseBuffer(uiResult == NO_ERROR ? dwSize : 0);
        return uiResult;
    } else if (uiResult == NO_ERROR) {
        // The string in database is empty.
        sValue.Empty();
        return NO_ERROR;
    } else {
        // Return error code.
        return uiResult;
    }
}


inline UINT MsiRecordGetStringW(MSIHANDLE hRecord, unsigned int iField, ATL::CAtlStringW &sValue)
{
    DWORD dwSize = 0;
    UINT uiResult;

    // Query the actual string length first.
    uiResult = ::MsiRecordGetStringW(hRecord, iField, L"", &dwSize);
    if (uiResult == ERROR_MORE_DATA) {
        // Prepare the buffer to read the string data into and read it.
        LPWSTR szBuffer = sValue.GetBuffer(dwSize++);
        if (!szBuffer) return ERROR_OUTOFMEMORY;
        uiResult = ::MsiRecordGetStringW(hRecord, iField, szBuffer, &dwSize);
        sValue.ReleaseBuffer(uiResult == NO_ERROR ? dwSize : 0);
        return uiResult;
    } else if (uiResult == NO_ERROR) {
        // The string in database is empty.
        sValue.Empty();
        return NO_ERROR;
    } else {
        // Return error code.
        return uiResult;
    }
}


inline UINT MsiRecordGetStream(MSIHANDLE hRecord, unsigned int iField, ATL::CAtlArray<BYTE> &binData)
{
    DWORD dwSize = 0;
    UINT uiResult;

    // Query the actual data length first.
    uiResult = ::MsiRecordReadStream(hRecord, iField, NULL, &dwSize);
    if (uiResult == NO_ERROR) {
        if (!binData.SetCount(dwSize)) return ERROR_OUTOFMEMORY;
        return ::MsiRecordReadStream(hRecord, iField, (char*)binData.GetData(), &dwSize);
    } else {
        // Return error code.
        return uiResult;
    }
}


inline UINT MsiFormatRecordA(MSIHANDLE hInstall, MSIHANDLE hRecord, ATL::CAtlStringA &sValue)
{
    DWORD dwSize = 0;
    UINT uiResult;

    // Query the final string length first.
    uiResult = ::MsiFormatRecordA(hInstall, hRecord, "", &dwSize);
    if (uiResult == ERROR_MORE_DATA) {
        // Prepare the buffer to format the string data into and read it.
        LPSTR szBuffer = sValue.GetBuffer(dwSize++);
        if (!szBuffer) return ERROR_OUTOFMEMORY;
        uiResult = ::MsiFormatRecordA(hInstall, hRecord, szBuffer, &dwSize);
        sValue.ReleaseBuffer(uiResult == NO_ERROR ? dwSize : 0);
        return uiResult;
    } else if (uiResult == NO_ERROR) {
        // The result is empty.
        sValue.Empty();
        return NO_ERROR;
    } else {
        // Return error code.
        return uiResult;
    }
}


inline UINT MsiFormatRecordW(MSIHANDLE hInstall, MSIHANDLE hRecord, ATL::CAtlStringW &sValue)
{
    DWORD dwSize = 0;
    UINT uiResult;

    // Query the final string length first.
    uiResult = ::MsiFormatRecordW(hInstall, hRecord, L"", &dwSize);
    if (uiResult == ERROR_MORE_DATA) {
        // Prepare the buffer to format the string data into and read it.
        LPWSTR szBuffer = sValue.GetBuffer(dwSize++);
        if (!szBuffer) return ERROR_OUTOFMEMORY;
        uiResult = ::MsiFormatRecordW(hInstall, hRecord, szBuffer, &dwSize);
        sValue.ReleaseBuffer(uiResult == NO_ERROR ? dwSize : 0);
        return uiResult;
    } else if (uiResult == NO_ERROR) {
        // The result is empty.
        sValue.Empty();
        return NO_ERROR;
    } else {
        // Return error code.
        return uiResult;
    }
}


inline UINT MsiRecordFormatStringA(MSIHANDLE hInstall, MSIHANDLE hRecord, unsigned int iField, ATL::CAtlStringA &sValue)
{
    UINT uiResult;
    PMSIHANDLE hRecordEx;

    // Read string to format.
    uiResult = ::MsiRecordGetStringA(hRecord, iField, sValue);
    if (uiResult != NO_ERROR) return uiResult;

    // If the string is empty, there's nothing left to do.
    if (sValue.IsEmpty()) return NO_ERROR;

    // Create a record.
    hRecordEx = ::MsiCreateRecord(1);
    if (!hRecordEx) return ERROR_INVALID_HANDLE;

    // Populate record with data.
    uiResult = ::MsiRecordSetStringA(hRecordEx, 0, sValue);
    if (uiResult != NO_ERROR) return uiResult;

    // Do the formatting.
    return ::MsiFormatRecordA(hInstall, hRecordEx, sValue);
}


inline UINT MsiRecordFormatStringW(MSIHANDLE hInstall, MSIHANDLE hRecord, unsigned int iField, ATL::CAtlStringW &sValue)
{
    UINT uiResult;
    PMSIHANDLE hRecordEx;

    // Read string to format.
    uiResult = ::MsiRecordGetStringW(hRecord, iField, sValue);
    if (uiResult != NO_ERROR) return uiResult;

    // If the string is empty, there's nothing left to do.
    if (sValue.IsEmpty()) return NO_ERROR;

    // Create a record.
    hRecordEx = ::MsiCreateRecord(1);
    if (!hRecordEx) return ERROR_INVALID_HANDLE;

    // Populate record with data.
    uiResult = ::MsiRecordSetStringW(hRecordEx, 0, sValue);
    if (uiResult != NO_ERROR) return uiResult;

    // Do the formatting.
    return ::MsiFormatRecordW(hInstall, hRecordEx, sValue);
}

#ifdef UNICODE
#define MsiRecordFormatString  MsiRecordFormatStringW
#else
#define MsiRecordFormatString  MsiRecordFormatStringA
#endif // !UNICODE


inline UINT MsiGetTargetPathA(MSIHANDLE hInstall, LPCSTR szFolder, ATL::CAtlStringA &sValue)
{
    DWORD dwSize = 0;
    UINT uiResult;

    // Query the final string length first.
    uiResult = ::MsiGetTargetPathA(hInstall, szFolder, "", &dwSize);
    if (uiResult == ERROR_MORE_DATA) {
        // Prepare the buffer to format the string data into and read it.
        LPSTR szBuffer = sValue.GetBuffer(dwSize++);
        if (!szBuffer) return ERROR_OUTOFMEMORY;
        uiResult = ::MsiGetTargetPathA(hInstall, szFolder, szBuffer, &dwSize);
        sValue.ReleaseBuffer(uiResult == NO_ERROR ? dwSize : 0);
        return uiResult;
    } else if (uiResult == NO_ERROR) {
        // The result is empty.
        sValue.Empty();
        return NO_ERROR;
    } else {
        // Return error code.
        return uiResult;
    }
}


inline UINT MsiGetTargetPathW(MSIHANDLE hInstall, LPCWSTR szFolder, ATL::CAtlStringW &sValue)
{
    DWORD dwSize = 0;
    UINT uiResult;

    // Query the final string length first.
    uiResult = ::MsiGetTargetPathW(hInstall, szFolder, L"", &dwSize);
    if (uiResult == ERROR_MORE_DATA) {
        // Prepare the buffer to format the string data into and read it.
        LPWSTR szBuffer = sValue.GetBuffer(dwSize++);
        if (!szBuffer) return ERROR_OUTOFMEMORY;
        uiResult = ::MsiGetTargetPathW(hInstall, szFolder, szBuffer, &dwSize);
        sValue.ReleaseBuffer(uiResult == NO_ERROR ? dwSize : 0);
        return uiResult;
    } else if (uiResult == NO_ERROR) {
        // The result is empty.
        sValue.Empty();
        return NO_ERROR;
    } else {
        // Return error code.
        return uiResult;
    }
}


inline DWORD CertGetNameStringA(PCCERT_CONTEXT pCertContext, DWORD dwType, DWORD dwFlags, void *pvTypePara, ATL::CAtlStringA &sNameString)
{
    // Query the final string length first.
    DWORD dwSize = ::CertGetNameStringA(pCertContext, dwType, dwFlags, pvTypePara, NULL, 0);

    // Prepare the buffer to format the string data into and read it.
    LPSTR szBuffer = sNameString.GetBuffer(dwSize);
    if (!szBuffer) return ERROR_OUTOFMEMORY;
    dwSize = ::CertGetNameStringA(pCertContext, dwType, dwFlags, pvTypePara, szBuffer, dwSize);
    sNameString.ReleaseBuffer(dwSize);
    return dwSize;
}


inline DWORD CertGetNameStringW(PCCERT_CONTEXT pCertContext, DWORD dwType, DWORD dwFlags, void *pvTypePara, ATL::CAtlStringW &sNameString)
{
    // Query the final string length first.
    DWORD dwSize = ::CertGetNameStringW(pCertContext, dwType, dwFlags, pvTypePara, NULL, 0);

    // Prepare the buffer to format the string data into and read it.
    LPWSTR szBuffer = sNameString.GetBuffer(dwSize);
    if (!szBuffer) return ERROR_OUTOFMEMORY;
    dwSize = ::CertGetNameStringW(pCertContext, dwType, dwFlags, pvTypePara, szBuffer, dwSize);
    sNameString.ReleaseBuffer(dwSize);
    return dwSize;
}


namespace MSICA {

////////////////////////////////////////////////////////////////////////////
// Inline operators
////////////////////////////////////////////////////////////////////////////

inline HRESULT operator <<(ATL::CAtlFile &f, int i)
{
    HRESULT hr;
    DWORD dwWritten;

    hr = f.Write(&i, sizeof(int), &dwWritten);
    return SUCCEEDED(hr) ? dwWritten == sizeof(int) ? hr : E_FAIL : hr;
}


inline HRESULT operator >>(ATL::CAtlFile &f, int &i)
{
    HRESULT hr;
    DWORD dwRead;

    hr = f.Read(&i, sizeof(int), dwRead);
    return SUCCEEDED(hr) ? dwRead == sizeof(int) ? hr : E_FAIL : hr;
}


inline HRESULT operator <<(ATL::CAtlFile &f, DWORDLONG i)
{
    HRESULT hr;
    DWORD dwWritten;

    hr = f.Write(&i, sizeof(DWORDLONG), &dwWritten);
    return SUCCEEDED(hr) ? dwWritten == sizeof(DWORDLONG) ? hr : E_FAIL : hr;
}


inline HRESULT operator >>(ATL::CAtlFile &f, DWORDLONG &i)
{
    HRESULT hr;
    DWORD dwRead;

    hr = f.Read(&i, sizeof(DWORDLONG), dwRead);
    return SUCCEEDED(hr) ? dwRead == sizeof(DWORDLONG) ? hr : E_FAIL : hr;
}


template <class E>
inline HRESULT operator <<(ATL::CAtlFile &f, const ATL::CAtlArray<E> &a)
{
    HRESULT hr;
    DWORD dwCount = (DWORD)a.GetCount(), dwWritten;

    // Write element count.
    hr =  f.Write(&dwCount, sizeof(DWORD), &dwWritten);
    if (FAILED(hr)) return hr;
    if (dwWritten < sizeof(DWORD)) return E_FAIL;

    // Write data.
    hr = f.Write(a.GetData(), sizeof(E) * dwCount, &dwWritten);
    return SUCCEEDED(hr) ? dwWritten == sizeof(E) * dwCount ? hr : E_FAIL : hr;
}


template <class E>
inline HRESULT operator >>(ATL::CAtlFile &f, ATL::CAtlArray<E> &a)
{
    HRESULT hr;
    DWORD dwCount, dwRead;

    // Read element count as 32-bit integer.
    hr =  f.Read(&dwCount, sizeof(DWORD), dwRead);
    if (FAILED(hr)) return hr;
    if (dwRead < sizeof(DWORD)) return E_FAIL;

    // Allocate the buffer.
    if (!a.SetCount(dwCount)) return E_OUTOFMEMORY;

    // Read data.
    hr = f.Read(a.GetData(), sizeof(E) * dwCount, dwRead);
    if (SUCCEEDED(hr)) a.SetCount(dwRead / sizeof(E));
    return hr;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const ATL::CAtlStringA &str)
{
    HRESULT hr;
    int iLength = str.GetLength();
    DWORD dwWritten;

    // Write string length (in characters) as 32-bit integer.
    hr =  f.Write(&iLength, sizeof(int), &dwWritten);
    if (FAILED(hr)) return hr;
    if (dwWritten < sizeof(int)) return E_FAIL;

    // Write string data (without terminator).
    hr = f.Write((LPCSTR)str, sizeof(CHAR) * iLength, &dwWritten);
    return SUCCEEDED(hr) ? dwWritten == sizeof(CHAR) * iLength ? hr : E_FAIL : hr;
}


inline HRESULT operator >>(ATL::CAtlFile &f, ATL::CAtlStringA &str)
{
    HRESULT hr;
    int iLength;
    LPSTR buf;
    DWORD dwRead;

    // Read string length (in characters) as 32-bit integer.
    hr =  f.Read(&iLength, sizeof(int), dwRead);
    if (FAILED(hr)) return hr;
    if (dwRead < sizeof(int)) return E_FAIL;

    // Allocate the buffer.
    buf = str.GetBuffer(iLength);
    if (!buf) return E_OUTOFMEMORY;

    // Read string data (without terminator).
    hr = f.Read(buf, sizeof(CHAR) * iLength, dwRead);
    str.ReleaseBuffer(SUCCEEDED(hr) ? dwRead / sizeof(CHAR) : 0);
    return hr;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const ATL::CAtlStringW &str)
{
    HRESULT hr;
    int iLength = str.GetLength();
    DWORD dwWritten;

    // Write string length (in characters) as 32-bit integer.
    hr = f.Write(&iLength, sizeof(int), &dwWritten);
    if (FAILED(hr)) return hr;
    if (dwWritten < sizeof(int)) return E_FAIL;

    // Write string data (without terminator).
    hr = f.Write((LPCWSTR)str, sizeof(WCHAR) * iLength, &dwWritten);
    return SUCCEEDED(hr) ? dwWritten == sizeof(WCHAR) * iLength ? hr : E_FAIL : hr;
}


inline HRESULT operator >>(ATL::CAtlFile &f, ATL::CAtlStringW &str)
{
    HRESULT hr;
    int iLength;
    LPWSTR buf;
    DWORD dwRead;

    // Read string length (in characters) as 32-bit integer.
    hr =  f.Read(&iLength, sizeof(int), dwRead);
    if (FAILED(hr)) return hr;
    if (dwRead < sizeof(int)) return E_FAIL;

    // Allocate the buffer.
    buf = str.GetBuffer(iLength);
    if (!buf) return E_OUTOFMEMORY;

    // Read string data (without terminator).
    hr = f.Read(buf, sizeof(WCHAR) * iLength, dwRead);
    str.ReleaseBuffer(SUCCEEDED(hr) ? dwRead / sizeof(WCHAR) : 0);
    return hr;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const TASK_TRIGGER &ttData)
{
    HRESULT hr;
    DWORD dwWritten;

    hr = f.Write(&ttData, sizeof(TASK_TRIGGER), &dwWritten);
    return SUCCEEDED(hr) ? dwWritten == sizeof(TASK_TRIGGER) ? hr : E_FAIL : hr;
}


inline HRESULT operator >>(ATL::CAtlFile &f, TASK_TRIGGER &ttData)
{
    HRESULT hr;
    DWORD dwRead;

    hr = f.Read(&ttData, sizeof(TASK_TRIGGER), dwRead);
    return SUCCEEDED(hr) ? dwRead == sizeof(TASK_TRIGGER) ? hr : E_FAIL : hr;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COperation &op)
{
    return f << op.m_iTicks;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COperation &op)
{
    return f >> op.m_iTicks;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COpTypeSingleString &op)
{
    HRESULT hr;

    hr = f << (const COperation &)op; if (FAILED(hr)) return hr;
    hr = f << op.m_sValue;            if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COpTypeSingleString &op)
{
    HRESULT hr;

    hr = f >> (COperation &)op; if (FAILED(hr)) return hr;
    hr = f >> op.m_sValue;      if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COpTypeSrcDstString &op)
{
    HRESULT hr;

    hr = f << (const COperation &)op; if (FAILED(hr)) return hr;
    hr = f << op.m_sValue1;           if (FAILED(hr)) return hr;
    hr = f << op.m_sValue2;           if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COpTypeSrcDstString &op)
{
    HRESULT hr;

    hr = f >> (COperation &)op; if (FAILED(hr)) return hr;
    hr = f >> op.m_sValue1;     if (FAILED(hr)) return hr;
    hr = f >> op.m_sValue2;     if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COpTypeBoolean &op)
{
    HRESULT hr;

    hr = f << (const COperation &)op; if (FAILED(hr)) return hr;
    hr = f << (int)(op.m_bValue);     if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COpTypeBoolean &op)
{
    HRESULT hr;
    int iValue;

    hr = f >> (COperation &)op; if (FAILED(hr)) return hr;
    hr = f >> iValue;           if (FAILED(hr)) return hr; op.m_bValue = iValue ? TRUE : FALSE;

    return S_OK;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COpRegKeySingle &op)
{
    HRESULT hr;

    hr = f << (const COpTypeSingleString &)op; if (FAILED(hr)) return hr;
    hr = f << (int)(op.m_hKeyRoot);            if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COpRegKeySingle &op)
{
    HRESULT hr;
    int iValue;

    hr = f >> (COpTypeSingleString &)op; if (FAILED(hr)) return hr;
    hr = f >> iValue;                    if (FAILED(hr)) return hr; op.m_hKeyRoot = (HKEY)iValue;

    return S_OK;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COpRegKeySrcDst &op)
{
    HRESULT hr;

    hr = f << (const COpTypeSrcDstString &)op; if (FAILED(hr)) return hr;
    hr = f << (int)(op.m_hKeyRoot);            if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COpRegKeySrcDst &op)
{
    HRESULT hr;
    int iValue;

    hr = f >> (COpTypeSrcDstString &)op; if (FAILED(hr)) return hr;
    hr = f >> iValue;                    if (FAILED(hr)) return hr; op.m_hKeyRoot = (HKEY)iValue;

    return S_OK;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COpRegValueSingle &op)
{
    HRESULT hr;

    hr = f << (const COpRegKeySingle &)op; if (FAILED(hr)) return hr;
    hr = f << op.m_sValueName;             if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COpRegValueSingle &op)
{
    HRESULT hr;

    hr = f >> (COpRegKeySingle &)op; if (FAILED(hr)) return hr;
    hr = f >> op.m_sValueName;       if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COpRegValueSrcDst &op)
{
    HRESULT hr;

    hr = f << (const COpRegKeySingle &)op; if (FAILED(hr)) return hr;
    hr = f << op.m_sValueName1;            if (FAILED(hr)) return hr;
    hr = f << op.m_sValueName2;            if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COpRegValueSrcDst &op)
{
    HRESULT hr;

    hr = f >> (COpRegKeySingle &)op; if (FAILED(hr)) return hr;
    hr = f >> op.m_sValueName1;      if (FAILED(hr)) return hr;
    hr = f >> op.m_sValueName2;      if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COpRegValueCreate &op)
{
    HRESULT hr;

    hr = f << (const COpRegValueSingle &)op; if (FAILED(hr)) return hr;
    hr = f << (int)(op.m_dwType);            if (FAILED(hr)) return hr;
    switch (op.m_dwType) {
    case REG_SZ:
    case REG_EXPAND_SZ:
    case REG_LINK:                hr = f << op.m_sData;         if (FAILED(hr)) return hr; break;
    case REG_BINARY:              hr = f << op.m_binData;       if (FAILED(hr)) return hr; break;
    case REG_DWORD_LITTLE_ENDIAN:
    case REG_DWORD_BIG_ENDIAN:    hr = f << (int)(op.m_dwData); if (FAILED(hr)) return hr; break;
    case REG_MULTI_SZ:            hr = f << op.m_szData;        if (FAILED(hr)) return hr; break;
    case REG_QWORD_LITTLE_ENDIAN: hr = f << op.m_qwData;        if (FAILED(hr)) return hr; break;
    default:                      return E_UNEXPECTED;
    }

    return S_OK;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COpRegValueCreate &op)
{
    HRESULT hr;

    hr = f >> (COpRegValueSingle &)op; if (FAILED(hr)) return hr;
    hr = f >> (int&)(op.m_dwType);     if (FAILED(hr)) return hr;
    switch (op.m_dwType) {
    case REG_SZ:
    case REG_EXPAND_SZ:
    case REG_LINK:                hr = f >> op.m_sData;          if (FAILED(hr)) return hr; break;
    case REG_BINARY:              hr = f >> op.m_binData;        if (FAILED(hr)) return hr; break;
    case REG_DWORD_LITTLE_ENDIAN:
    case REG_DWORD_BIG_ENDIAN:    hr = f >> (int&)(op.m_dwData); if (FAILED(hr)) return hr; break;
    case REG_MULTI_SZ:            hr = f >> op.m_szData;         if (FAILED(hr)) return hr; break;
    case REG_QWORD_LITTLE_ENDIAN: hr = f >> op.m_qwData;         if (FAILED(hr)) return hr; break;
    default:                      return E_UNEXPECTED;
    }

    return S_OK;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COpTaskCreate &op)
{
    HRESULT hr;
    POSITION pos;

    hr = f << (const COpTypeSingleString&)op;                          if (FAILED(hr)) return hr;
    hr = f << op.m_sApplicationName;                                   if (FAILED(hr)) return hr;
    hr = f << op.m_sParameters;                                        if (FAILED(hr)) return hr;
    hr = f << op.m_sWorkingDirectory;                                  if (FAILED(hr)) return hr;
    hr = f << op.m_sAuthor;                                            if (FAILED(hr)) return hr;
    hr = f << op.m_sComment;                                           if (FAILED(hr)) return hr;
    hr = f << (int)(op.m_dwFlags);                                     if (FAILED(hr)) return hr;
    hr = f << (int)(op.m_dwPriority);                                  if (FAILED(hr)) return hr;
    hr = f << op.m_sAccountName;                                       if (FAILED(hr)) return hr;
    hr = f << op.m_sPassword;                                          if (FAILED(hr)) return hr;
    hr = f << (int)MAKELONG(op.m_wDeadlineMinutes, op.m_wIdleMinutes); if (FAILED(hr)) return hr;
    hr = f << (int)(op.m_dwMaxRuntimeMS);                              if (FAILED(hr)) return hr;
    hr = f << (int)(op.m_lTriggers.GetCount());                        if (FAILED(hr)) return hr;
    for (pos = op.m_lTriggers.GetHeadPosition(); pos;) {
        hr = f << op.m_lTriggers.GetNext(pos);
        if (FAILED(hr)) return hr;
    }

    return S_OK;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COpTaskCreate &op)
{
    HRESULT hr;
    DWORD dwValue;

    hr = f >> (COpTypeSingleString&)op;             if (FAILED(hr)) return hr;
    hr = f >> op.m_sApplicationName;                if (FAILED(hr)) return hr;
    hr = f >> op.m_sParameters;                     if (FAILED(hr)) return hr;
    hr = f >> op.m_sWorkingDirectory;               if (FAILED(hr)) return hr;
    hr = f >> op.m_sAuthor;                         if (FAILED(hr)) return hr;
    hr = f >> op.m_sComment;                        if (FAILED(hr)) return hr;
    hr = f >> (int&)(op.m_dwFlags);                 if (FAILED(hr)) return hr;
    hr = f >> (int&)(op.m_dwPriority);              if (FAILED(hr)) return hr;
    hr = f >> op.m_sAccountName;                    if (FAILED(hr)) return hr;
    hr = f >> op.m_sPassword;                       if (FAILED(hr)) return hr;
    hr = f >> (int&)dwValue;                        if (FAILED(hr)) return hr; op.m_wIdleMinutes = HIWORD(dwValue); op.m_wDeadlineMinutes = LOWORD(dwValue);
    hr = f >> (int&)(op.m_dwMaxRuntimeMS);          if (FAILED(hr)) return hr;
    hr = f >> (int&)dwValue;                        if (FAILED(hr)) return hr;
    while (dwValue--) {
        TASK_TRIGGER ttData;
        hr = f >> ttData;
        if (FAILED(hr)) return hr;
        op.m_lTriggers.AddTail(ttData);
    }

    return S_OK;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COpTaskEnable &op)
{
    HRESULT hr;

    hr = f << (const COpTypeSingleString&)op; if (FAILED(hr)) return hr;
    hr = f << (int)(op.m_bEnable);            if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COpTaskEnable &op)
{
    HRESULT hr;
    int iTemp;

    hr = f >> (COpTypeSingleString&)op; if (FAILED(hr)) return hr;
    hr = f >> iTemp;                    if (FAILED(hr)) return hr; op.m_bEnable = iTemp ? TRUE : FALSE;

    return S_OK;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COpCertStore &op)
{
    HRESULT hr;

    hr = f << (const COpTypeSingleString&)op; if (FAILED(hr)) return hr;
    hr = f << (int)(op.m_dwEncodingType);     if (FAILED(hr)) return hr;
    hr = f << (int)(op.m_dwFlags);            if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COpCertStore &op)
{
    HRESULT hr;

    hr = f >> (COpTypeSingleString&)op;    if (FAILED(hr)) return hr;
    hr = f >> (int&)(op.m_dwEncodingType); if (FAILED(hr)) return hr;
    hr = f >> (int&)(op.m_dwFlags);        if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COpCert &op)
{
    HRESULT hr;

    hr = f << (const COpCertStore&)op; if (FAILED(hr)) return hr;
    hr = f << op.m_binCert;            if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COpCert &op)
{
    HRESULT hr;

    hr = f >> (COpCertStore&)op; if (FAILED(hr)) return hr;
    hr = f >> op.m_binCert;      if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COpSvcSetStart &op)
{
    HRESULT hr;

    hr = f << (const COpTypeSingleString&)op; if (FAILED(hr)) return hr;
    hr = f << (int)(op.m_dwStartType);        if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COpSvcSetStart &op)
{
    HRESULT hr;

    hr = f >> (COpTypeSingleString&)op; if (FAILED(hr)) return hr;
    hr = f >> (int&)(op.m_dwStartType); if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COpSvcControl &op)
{
    HRESULT hr;

    hr = f << (const COpTypeSingleString&)op; if (FAILED(hr)) return hr;
    hr = f << (int)(op.m_bWait);              if (FAILED(hr)) return hr;

    return S_OK;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COpSvcControl &op)
{
    HRESULT hr;
    int iValue;

    hr = f >> (COpTypeSingleString&)op; if (FAILED(hr)) return hr;
    hr = f >> iValue;                   if (FAILED(hr)) return hr; op.m_bWait = iValue ? TRUE : FALSE;

    return S_OK;
}


inline HRESULT operator <<(ATL::CAtlFile &f, const COpList &list)
{
    POSITION pos;
    HRESULT hr;

    hr = f << (const COperation &)list;
    if (FAILED(hr)) return hr;

    hr = f << (int)(list.GetCount());
    if (FAILED(hr)) return hr;

    for (pos = list.GetHeadPosition(); pos;) {
        const COperation *pOp = list.GetNext(pos);
        if (dynamic_cast<const COpRollbackEnable*>(pOp))
            hr = list.Save<COpRollbackEnable, COpList::OP_ROLLBACK_ENABLE>(f, pOp);
        else if (dynamic_cast<const COpFileDelete*>(pOp))
            hr = list.Save<COpFileDelete, COpList::OP_FILE_DELETE>(f, pOp);
        else if (dynamic_cast<const COpFileMove*>(pOp))
            hr = list.Save<COpFileMove, COpList::OP_FILE_MOVE>(f, pOp);
        else if (dynamic_cast<const COpRegKeyCreate*>(pOp))
            hr = list.Save<COpRegKeyCreate, COpList::OP_REG_KEY_CREATE>(f, pOp);
        else if (dynamic_cast<const COpRegKeyCopy*>(pOp))
            hr = list.Save<COpRegKeyCopy, COpList::OP_REG_KEY_COPY>(f, pOp);
        else if (dynamic_cast<const COpRegKeyDelete*>(pOp))
            hr = list.Save<COpRegKeyDelete, COpList::OP_REG_KEY_DELETE>(f, pOp);
        else if (dynamic_cast<const COpRegValueCreate*>(pOp))
            hr = list.Save<COpRegValueCreate, COpList::OP_REG_VALUE_CREATE>(f, pOp);
        else if (dynamic_cast<const COpRegValueCopy*>(pOp))
            hr = list.Save<COpRegValueCopy, COpList::OP_REG_VALUE_COPY>(f, pOp);
        else if (dynamic_cast<const COpRegValueDelete*>(pOp))
            hr = list.Save<COpRegValueDelete, COpList::OP_REG_VALUE_DELETE>(f, pOp);
        else if (dynamic_cast<const COpTaskCreate*>(pOp))
            hr = list.Save<COpTaskCreate, COpList::OP_TASK_CREATE>(f, pOp);
        else if (dynamic_cast<const COpTaskDelete*>(pOp))
            hr = list.Save<COpTaskDelete, COpList::OP_TASK_DELETE>(f, pOp);
        else if (dynamic_cast<const COpTaskEnable*>(pOp))
            hr = list.Save<COpTaskEnable, COpList::OP_TASK_ENABLE>(f, pOp);
        else if (dynamic_cast<const COpTaskCopy*>(pOp))
            hr = list.Save<COpTaskCopy, COpList::OP_TASK_COPY>(f, pOp);
        else if (dynamic_cast<const COpCertInstall*>(pOp))
            hr = list.Save<COpCertInstall, COpList::OP_CERT_INSTALL>(f, pOp);
        else if (dynamic_cast<const COpCertRemove*>(pOp))
            hr = list.Save<COpCertRemove, COpList::OP_CERT_REMOVE>(f, pOp);
        else if (dynamic_cast<const COpSvcSetStart*>(pOp))
            hr = list.Save<COpSvcSetStart, COpList::OP_SVC_SET_START>(f, pOp);
        else if (dynamic_cast<const COpSvcStart*>(pOp))
            hr = list.Save<COpSvcStart, COpList::OP_SVC_START>(f, pOp);
        else if (dynamic_cast<const COpSvcStop*>(pOp))
            hr = list.Save<COpSvcStop, COpList::OP_SVC_STOP>(f, pOp);
        else if (dynamic_cast<const COpList*>(pOp))
            hr = list.Save<COpList, COpList::OP_SUBLIST>(f, pOp);
        else {
            // Unsupported type of operation.
            hr = E_UNEXPECTED;
        }

        if (FAILED(hr)) return hr;
    }

    return S_OK;
}


inline HRESULT operator >>(ATL::CAtlFile &f, COpList &list)
{
    HRESULT hr;
    int iCount;

    hr = f >> (COperation &)list;
    if (FAILED(hr)) return hr;

    hr = f >> iCount;
    if (FAILED(hr)) return hr;

    while (iCount--) {
        int iTemp;

        hr = f >> iTemp;
        if (FAILED(hr)) return hr;

        switch ((COpList::OPERATION)iTemp) {
        case COpList::OP_ROLLBACK_ENABLE:  hr = list.LoadAndAddTail<COpRollbackEnable>(f); break;
        case COpList::OP_FILE_DELETE:      hr = list.LoadAndAddTail<COpFileDelete    >(f); break;
        case COpList::OP_FILE_MOVE:        hr = list.LoadAndAddTail<COpFileMove      >(f); break;
        case COpList::OP_REG_KEY_CREATE:   hr = list.LoadAndAddTail<COpRegKeyCreate  >(f); break;
        case COpList::OP_REG_KEY_COPY:     hr = list.LoadAndAddTail<COpRegKeyCopy    >(f); break;
        case COpList::OP_REG_KEY_DELETE:   hr = list.LoadAndAddTail<COpRegKeyDelete  >(f); break;
        case COpList::OP_REG_VALUE_CREATE: hr = list.LoadAndAddTail<COpRegValueCreate>(f); break;
        case COpList::OP_REG_VALUE_COPY:   hr = list.LoadAndAddTail<COpRegValueCopy  >(f); break;
        case COpList::OP_REG_VALUE_DELETE: hr = list.LoadAndAddTail<COpRegValueDelete>(f); break;
        case COpList::OP_TASK_CREATE:      hr = list.LoadAndAddTail<COpTaskCreate    >(f); break;
        case COpList::OP_TASK_DELETE:      hr = list.LoadAndAddTail<COpTaskDelete    >(f); break;
        case COpList::OP_TASK_ENABLE:      hr = list.LoadAndAddTail<COpTaskEnable    >(f); break;
        case COpList::OP_TASK_COPY:        hr = list.LoadAndAddTail<COpTaskCopy      >(f); break;
        case COpList::OP_CERT_INSTALL:     hr = list.LoadAndAddTail<COpCertInstall   >(f); break;
        case COpList::OP_CERT_REMOVE:      hr = list.LoadAndAddTail<COpCertRemove    >(f); break;
        case COpList::OP_SVC_SET_START:    hr = list.LoadAndAddTail<COpSvcSetStart   >(f); break;
        case COpList::OP_SVC_START:        hr = list.LoadAndAddTail<COpSvcStart      >(f); break;
        case COpList::OP_SVC_STOP:         hr = list.LoadAndAddTail<COpSvcStop       >(f); break;
        case COpList::OP_SUBLIST:          hr = list.LoadAndAddTail<COpList          >(f); break;
        default:
            // Unsupported type of operation.
            hr = E_UNEXPECTED;
        }

        if (FAILED(hr)) return hr;
    }

    return S_OK;
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

#ifndef _WIN64
#endif



////////////////////////////////////////////////////////////////////////////
// Inline methods
////////////////////////////////////////////////////////////////////////////

template <class T, enum COpList::OPERATION ID> inline static HRESULT COpList::Save(ATL::CAtlFile &f, const COperation *p)
{
    HRESULT hr;
    const T *pp = dynamic_cast<const T*>(p);
    if (!pp) return E_UNEXPECTED;

    hr = f << (int)ID;
    if (FAILED(hr)) return hr;

    return f << *pp;
}


template <class T> inline HRESULT COpList::LoadAndAddTail(ATL::CAtlFile &f)
{
    HRESULT hr;

    // Create element.
    T *p = new T();
    if (!p) return E_OUTOFMEMORY;

    // Load element from file.
    hr = f >> *p;
    if (FAILED(hr)) {
        delete p;
        return hr;
    }

    // Add element.
    AddTail(p);
    return S_OK;
}

} // namespace MSICA

#endif // __MSICALib_H__
