#ifndef __MSITSCA_H__
#define __MSITSCA_H__


////////////////////////////////////////////////////////////////////////////
// Constants
////////////////////////////////////////////////////////////////////////////

#define MSITSCA_VERSION       0x01000000

#define MSITSCA_VERSION_MAJ   1
#define MSITSCA_VERSION_MIN   0
#define MSITSCA_VERSION_REV   0

#define MSITSCA_VERSION_STR   "1.0"
#define MSITSCA_VERSION_INST  "1.0.0.0"


////////////////////////////////////////////////////////////////////
// Resource IDs
////////////////////////////////////////////////////////////////////

#define IDR_MAINFRAME 1

#if !defined(RC_INVOKED) && !defined(MIDL_PASS)

#include <msi.h>


////////////////////////////////////////////////////////////////////
// Calling declaration
////////////////////////////////////////////////////////////////////

#if defined(MSITSCA_DLL)
#define MSITSCA_API __declspec(dllexport)
#elif defined(MSITSCA_DLLIMP)
#define MSITSCA_API __declspec(dllimport)
#else
#define MSITSCA_API
#endif


////////////////////////////////////////////////////////////////////
// Exported functions
////////////////////////////////////////////////////////////////////

#ifdef __cplusplus
extern "C" {
#endif

    UINT MSITSCA_API EvaluateScheduledTasks(MSIHANDLE hInstall);
    UINT MSITSCA_API InstallScheduledTasks(MSIHANDLE hInstall);
    UINT MSITSCA_API FinalizeScheduledTasks(MSIHANDLE hInstall);

#ifdef __cplusplus
}
#endif


////////////////////////////////////////////////////////////////////
// Local includes
////////////////////////////////////////////////////////////////////

#include <atlfile.h>
#include <atlstr.h>
#include <assert.h>
#include <msiquery.h>
#include <mstask.h>


////////////////////////////////////////////////////////////////////
// Inline Functions
////////////////////////////////////////////////////////////////////

inline UINT MsiGetPropertyA(MSIHANDLE hInstall, LPCSTR szName, CStringA &sValue)
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
        sValue.ReleaseBuffer(uiResult == ERROR_SUCCESS ? dwSize : 0);
        return uiResult;
    } else if (uiResult == ERROR_SUCCESS) {
        // The string in database is empty.
        sValue.Empty();
        return ERROR_SUCCESS;
    } else {
        // Return error code.
        return uiResult;
    }
}


inline UINT MsiGetPropertyW(MSIHANDLE hInstall, LPCWSTR szName, CStringW &sValue)
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
        sValue.ReleaseBuffer(uiResult == ERROR_SUCCESS ? dwSize : 0);
        return uiResult;
    } else if (uiResult == ERROR_SUCCESS) {
        // The string in database is empty.
        sValue.Empty();
        return ERROR_SUCCESS;
    } else {
        // Return error code.
        return uiResult;
    }
}


inline UINT MsiRecordGetStringA(MSIHANDLE hRecord, unsigned int iField, CStringA &sValue)
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
        sValue.ReleaseBuffer(uiResult == ERROR_SUCCESS ? dwSize : 0);
        return uiResult;
    } else if (uiResult == ERROR_SUCCESS) {
        // The string in database is empty.
        sValue.Empty();
        return ERROR_SUCCESS;
    } else {
        // Return error code.
        return uiResult;
    }
}


inline UINT MsiRecordGetStringW(MSIHANDLE hRecord, unsigned int iField, CStringW &sValue)
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
        sValue.ReleaseBuffer(uiResult == ERROR_SUCCESS ? dwSize : 0);
        return uiResult;
    } else if (uiResult == ERROR_SUCCESS) {
        // The string in database is empty.
        sValue.Empty();
        return ERROR_SUCCESS;
    } else {
        // Return error code.
        return uiResult;
    }
}


inline UINT MsiFormatRecordA(MSIHANDLE hInstall, MSIHANDLE hRecord, CStringA &sValue)
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
        sValue.ReleaseBuffer(uiResult == ERROR_SUCCESS ? dwSize : 0);
        return uiResult;
    } else if (uiResult == ERROR_SUCCESS) {
        // The result is empty.
        sValue.Empty();
        return ERROR_SUCCESS;
    } else {
        // Return error code.
        return uiResult;
    }
}


inline UINT MsiFormatRecordW(MSIHANDLE hInstall, MSIHANDLE hRecord, CStringW &sValue)
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
        sValue.ReleaseBuffer(uiResult == ERROR_SUCCESS ? dwSize : 0);
        return uiResult;
    } else if (uiResult == ERROR_SUCCESS) {
        // The result is empty.
        sValue.Empty();
        return ERROR_SUCCESS;
    } else {
        // Return error code.
        return uiResult;
    }
}


inline UINT MsiRecordFormatStringA(MSIHANDLE hInstall, MSIHANDLE hRecord, unsigned int iField, CStringA &sValue)
{
    UINT uiResult;
    PMSIHANDLE hRecordEx;

    // Read string to format.
    uiResult = ::MsiRecordGetStringA(hRecord, iField, sValue);
    if (uiResult != ERROR_SUCCESS) return uiResult;

    // If the string is empty, there's nothing left to do.
    if (sValue.IsEmpty()) return ERROR_SUCCESS;

    // Create a record.
    hRecordEx = ::MsiCreateRecord(1);
    if (!hRecordEx) return ERROR_INVALID_HANDLE;

    // Populate record with data.
    uiResult = ::MsiRecordSetStringA(hRecordEx, 0, sValue);
    if (uiResult != ERROR_SUCCESS) return uiResult;

    // Do the formatting.
    return ::MsiFormatRecordA(hInstall, hRecordEx, sValue);
}


inline UINT MsiRecordFormatStringW(MSIHANDLE hInstall, MSIHANDLE hRecord, unsigned int iField, CStringW &sValue)
{
    UINT uiResult;
    PMSIHANDLE hRecordEx;

    // Read string to format.
    uiResult = ::MsiRecordGetStringW(hRecord, iField, sValue);
    if (uiResult != ERROR_SUCCESS) return uiResult;

    // If the string is empty, there's nothing left to do.
    if (sValue.IsEmpty()) return ERROR_SUCCESS;

    // Create a record.
    hRecordEx = ::MsiCreateRecord(1);
    if (!hRecordEx) return ERROR_INVALID_HANDLE;

    // Populate record with data.
    uiResult = ::MsiRecordSetStringW(hRecordEx, 0, sValue);
    if (uiResult != ERROR_SUCCESS) return uiResult;

    // Do the formatting.
    return ::MsiFormatRecordW(hInstall, hRecordEx, sValue);
}

#ifdef UNICODE
#define MsiRecordFormatString  MsiRecordFormatStringW
#else
#define MsiRecordFormatString  MsiRecordFormatStringA
#endif // !UNICODE


////////////////////////////////////////////////////////////////////
// Inline operators
////////////////////////////////////////////////////////////////////

inline HRESULT operator <<(CAtlFile &f, int i)
{
    return f.Write(&i, sizeof(int));
}


inline HRESULT operator >>(CAtlFile &f, int &i)
{
    return f.Read(&i, sizeof(int));
}


inline HRESULT operator <<(CAtlFile &f, const CStringA &str)
{
    HRESULT hr;
    int iLength = str.GetLength();

    // Write string length (in characters) as 32-bit integer.
    hr =  f.Write(&iLength, sizeof(int));
    if (FAILED(hr)) return hr;

    // Write string data (without terminator).
    return f.Write((LPCSTR)str, sizeof(CHAR) * iLength);
}


inline HRESULT operator >>(CAtlFile &f, CStringA &str)
{
    HRESULT hr;
    int iLength;
    LPSTR buf;

    // Read string length (in characters) as 32-bit integer.
    hr =  f.Read(&iLength, sizeof(int));
    if (FAILED(hr)) return hr;

    // Allocate the buffer.
    buf = str.GetBuffer(iLength);
    if (!buf) return E_OUTOFMEMORY;

    // Read string data (without terminator).
    hr = f.Read(buf, sizeof(CHAR) * iLength);
    str.ReleaseBuffer(SUCCEEDED(hr) ? iLength : 0);
    return hr;
}


inline HRESULT operator <<(CAtlFile &f, const CStringW &str)
{
    HRESULT hr;
    int iLength = str.GetLength();

    // Write string length (in characters) as 32-bit integer.
    hr = f.Write(&iLength, sizeof(int));
    if (FAILED(hr)) return hr;

    // Write string data (without terminator).
    return f.Write((LPCWSTR)str, sizeof(WCHAR) * iLength);
}


inline HRESULT operator >>(CAtlFile &f, CStringW &str)
{
    HRESULT hr;
    int iLength;
    LPWSTR buf;

    // Read string length (in characters) as 32-bit integer.
    hr =  f.Read(&iLength, sizeof(int));
    if (FAILED(hr)) return hr;

    // Allocate the buffer.
    buf = str.GetBuffer(iLength);
    if (!buf) return E_OUTOFMEMORY;

    // Read string data (without terminator).
    hr = f.Read(buf, sizeof(WCHAR) * iLength);
    str.ReleaseBuffer(SUCCEEDED(hr) ? iLength : 0);
    return hr;
}


inline HRESULT operator >>(CAtlFile &f, ITask *pTask)
{
    assert(pTask);

    HRESULT hr;
    CStringW sValue, sValue2;
    int iValue, iValue2;
    UINT nTriggers;

    hr = f >> sValue;
    if (FAILED(hr)) return hr;
    hr = pTask->SetApplicationName(sValue);
    if (FAILED(hr)) return hr;

    hr = f >> sValue;
    if (FAILED(hr)) return hr;
    hr = pTask->SetParameters(sValue);
    if (FAILED(hr)) return hr;

    hr = f >> sValue;
    if (FAILED(hr)) return hr;
    hr = pTask->SetWorkingDirectory(sValue);
    if (FAILED(hr)) return hr;

    hr = f >> iValue;
    if (FAILED(hr)) return hr;
    hr = pTask->SetFlags(iValue);
    if (FAILED(hr)) return hr;

    hr = f >> iValue;
    if (FAILED(hr)) return hr;
    hr = pTask->SetPriority(iValue);
    if (FAILED(hr)) return hr;

    hr = f >> sValue;
    if (FAILED(hr)) return hr;
    hr = f >> sValue2;
    if (FAILED(hr)) return hr;
    hr = !sValue.IsEmpty() ? pTask->SetAccountInformation(sValue, sValue2.IsEmpty() ? NULL : sValue2) : S_OK;
    {
        // Clear the password in memory before proceeding.
        int iLength = sValue2.GetLength();
        LPWSTR pszValue2 = sValue2.GetBuffer(iLength);
        ::SecureZeroMemory(pszValue2, sizeof(WCHAR) * iLength);
        sValue2.ReleaseBuffer(0);
    }
    if (FAILED(hr)) return hr;

    hr = f >> sValue;
    if (FAILED(hr)) return hr;
    hr = pTask->SetComment(sValue);
    if (FAILED(hr)) return hr;

    hr = f >> iValue;
    if (FAILED(hr)) return hr;
    hr = f >> iValue2;
    if (FAILED(hr)) return hr;
    hr = pTask->SetIdleWait((WORD)iValue, (WORD)iValue2);
    if (FAILED(hr)) return hr;

    hr = f >> iValue;
    if (FAILED(hr)) return hr;
    hr = pTask->SetMaxRunTime(iValue);
    if (FAILED(hr)) return hr;

    // Read and add triggers.
    hr = f >> (int&)nTriggers;
    if (FAILED(hr)) return hr;

    while (nTriggers--) {
        CComPtr<ITaskTrigger> pTrigger;
        TASK_TRIGGER ttData;
        WORD wTriggerIdx;
        ULONGLONG ullValue;
        FILETIME ftValue;
        SYSTEMTIME stValue;
        int iDaysInterval, iWeeksInterval, iDaysOfTheWeek, iDaysOfMonth, iWeekOfMonth, iMonthsOfYear;

        ZeroMemory(&ttData, sizeof(TASK_TRIGGER));
        ttData.cbTriggerSize = sizeof(TASK_TRIGGER);

        hr = pTask->CreateTrigger(&wTriggerIdx, &pTrigger);
        if (FAILED(hr)) return hr;

        hr = f >> iValue;
        if (FAILED(hr)) return hr;
        ullValue = ((ULONGLONG)iValue + 138426) * 864000000000;
        ftValue.dwHighDateTime = ullValue >> 32;
        ftValue.dwLowDateTime  = ullValue & 0xffffffff;
        if (!::FileTimeToSystemTime(&ftValue, &stValue))
            return AtlHresultFromLastError();
        ttData.wBeginYear  = stValue.wYear;
        ttData.wBeginMonth = stValue.wMonth;
        ttData.wBeginDay   = stValue.wDay;

        hr = f >> iValue;
        if (FAILED(hr)) return hr;
        if (iValue != MSI_NULL_INTEGER) {
            ullValue = ((ULONGLONG)iValue + 138426) * 864000000000;
            ftValue.dwHighDateTime = ullValue >> 32;
            ftValue.dwLowDateTime  = ullValue & 0xffffffff;
            if (!::FileTimeToSystemTime(&ftValue, &stValue))
                return AtlHresultFromLastError();
            ttData.wEndYear  = stValue.wYear;
            ttData.wEndMonth = stValue.wMonth;
            ttData.wEndDay   = stValue.wDay;
            ttData.rgFlags  |= TASK_TRIGGER_FLAG_HAS_END_DATE;
        }

        hr = f >> iValue;
        if (FAILED(hr)) return hr;
        ttData.wStartHour   = (WORD)(iValue / 60);
        ttData.wStartMinute = (WORD)(iValue % 60);

        hr = f >> iValue;
        if (FAILED(hr)) return hr;
        ttData.MinutesDuration = iValue;

        hr = f >> iValue;
        if (FAILED(hr)) return hr;
        ttData.MinutesInterval = iValue;

        hr = f >> iValue;
        if (FAILED(hr)) return hr;
        ttData.rgFlags |= iValue & ~TASK_TRIGGER_FLAG_HAS_END_DATE;

        hr = f >> iValue;
        if (FAILED(hr)) return hr;
        ttData.TriggerType = (TASK_TRIGGER_TYPE)iValue;

        hr = f >> iDaysInterval;
        if (FAILED(hr)) return hr;
        hr = f >> iWeeksInterval;
        if (FAILED(hr)) return hr;
        hr = f >> iDaysOfTheWeek;
        if (FAILED(hr)) return hr;
        hr = f >> iDaysOfMonth;
        if (FAILED(hr)) return hr;
        hr = f >> iWeekOfMonth;
        if (FAILED(hr)) return hr;
        hr = f >> iMonthsOfYear;
        if (FAILED(hr)) return hr;
        switch (ttData.TriggerType) {
        case TASK_TIME_TRIGGER_DAILY:
            ttData.Type.Daily.DaysInterval          = (WORD)iDaysInterval;
            break;
        case TASK_TIME_TRIGGER_WEEKLY:
            ttData.Type.Weekly.WeeksInterval        = (WORD)iWeeksInterval;
            ttData.Type.Weekly.rgfDaysOfTheWeek     = (WORD)iDaysOfTheWeek;
            break;
        case TASK_TIME_TRIGGER_MONTHLYDATE:
            ttData.Type.MonthlyDate.rgfDays         = iDaysOfMonth;
            ttData.Type.MonthlyDate.rgfMonths       = (WORD)iMonthsOfYear;
            break;
        case TASK_TIME_TRIGGER_MONTHLYDOW:
            ttData.Type.MonthlyDOW.wWhichWeek       = (WORD)iWeekOfMonth;
            ttData.Type.MonthlyDOW.rgfDaysOfTheWeek = (WORD)iDaysOfTheWeek;
            ttData.Type.MonthlyDOW.rgfMonths        = (WORD)iMonthsOfYear;
            break;
        }

        hr = pTrigger->SetTrigger(&ttData);
        if (FAILED(hr)) return hr;
    }

    return S_OK;
}

#endif // !defined(RC_INVOKED) && !defined(MIDL_PASS)

#endif // __MSITSCA_H__
