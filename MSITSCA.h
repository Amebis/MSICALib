#ifndef __MSITSCA_H__
#define __MSITSCA_H__


////////////////////////////////////////////////////////////////////////////
// Konstante
////////////////////////////////////////////////////////////////////////////

#define MSITSCA_VERSION       0x01000000

#define MSITSCA_VERSION_MAJ   1
#define MSITSCA_VERSION_MIN   0
#define MSITSCA_VERSION_REV   0

#define MSITSCA_VERSION_STR   "1.0"
#define MSITSCA_VERSION_INST  "1.0.0.0"


////////////////////////////////////////////////////////////////////
// Kode virov
////////////////////////////////////////////////////////////////////

#define IDR_MAINFRAME 1
// TODO: Dodaj definicije konstant virov tukaj.

#if !defined(RC_INVOKED) && !defined(MIDL_PASS)

#include <msi.h>


////////////////////////////////////////////////////////////////////
// Naèin klicanja funkcij
////////////////////////////////////////////////////////////////////

#if defined(MSITSCA_DLL)
#define MSITSCA_API __declspec(dllexport)
#elif defined(MSITSCA_DLLIMP)
#define MSITSCA_API __declspec(dllimport)
#else
#define MSITSCA_API
#endif


////////////////////////////////////////////////////////////////////
// Javne funkcije
////////////////////////////////////////////////////////////////////

#ifdef __cplusplus
extern "C" {
#endif

    UINT MSITSCA_API EvaluateScheduledTasks(MSIHANDLE hInstall);
    UINT MSITSCA_API InstallScheduledTasks(MSIHANDLE hInstall);
    UINT MSITSCA_API CommitScheduledTasks(MSIHANDLE hInstall);
    UINT MSITSCA_API RollbackScheduledTasks(MSIHANDLE hInstall);

#ifdef __cplusplus
}
#endif


////////////////////////////////////////////////////////////////////
// Globalne funkcije in spremenljivke
////////////////////////////////////////////////////////////////////

namespace MSITSCA {
    extern HINSTANCE hInstance; // roèica modula
}


////////////////////////////////////////////////////////////////////
// Lokalni include
////////////////////////////////////////////////////////////////////

#include <afx.h>
#include <atlstr.h>
#include <assert.h>
#include <msiquery.h>


////////////////////////////////////////////////////////////////////
// Funkcije inline
////////////////////////////////////////////////////////////////////

inline UINT MsiRecordGetStringA(MSIHANDLE hRecord, unsigned int iField, CStringA &sValue)
{
    DWORD dwSize = 0;
    UINT uiResult;

    // Query the actual string length first.
    uiResult = ::MsiRecordGetStringA(hRecord, iField, "", &dwSize);
    if (uiResult == ERROR_MORE_DATA) {
        // Prepare the buffer to read the string data into and read it.
        LPSTR szBuffer = sValue.GetBuffer(dwSize++);
        assert(szBuffer);
        uiResult = ::MsiRecordGetStringA(hRecord, iField, szBuffer, &dwSize);
        if (uiResult == ERROR_SUCCESS) {
            // Read succeeded.
            sValue.ReleaseBuffer(dwSize);
            return ERROR_SUCCESS;
        } else {
            // Read failed. Empty the string and return error code.
            sValue.ReleaseBuffer(0);
            return uiResult;
        }
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
        assert(szBuffer);
        uiResult = ::MsiRecordGetStringW(hRecord, iField, szBuffer, &dwSize);
        if (uiResult == ERROR_SUCCESS) {
            // Read succeeded.
            sValue.ReleaseBuffer(dwSize);
            return ERROR_SUCCESS;
        } else {
            // Read failed. Empty the string and return error code.
            sValue.ReleaseBuffer(0);
            return uiResult;
        }
    } else if (uiResult == ERROR_SUCCESS) {
        // The string in database is empty.
        sValue.Empty();
        return ERROR_SUCCESS;
    } else {
        // Return error code.
        return uiResult;
    }
}


////////////////////////////////////////////////////////////////////
// Operatorji inline
////////////////////////////////////////////////////////////////////

inline CFile& operator <<(CFile &f, int i)
{
    f.Write(&i, sizeof(int));
    return f;
}


inline CFile& operator <<(CFile &f, const CStringA &str)
{
    int iLength = str.GetLength();
    f.Write(&iLength, sizeof(int));
    f.Write((LPCSTR)str, sizeof(CHAR) * iLength);
    return f;
}


inline CFile& operator <<(CFile &f, const CStringW &str)
{
    int iLength = str.GetLength();
    f.Write(&iLength, sizeof(int));
    f.Write((LPCWSTR)str, sizeof(WCHAR) * iLength);
    return f;
}

#endif // !defined(RC_INVOKED) && !defined(MIDL_PASS)

#endif // __MSITSCA_H__
