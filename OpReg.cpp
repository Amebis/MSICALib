/*
    MSICA, Copyright (C) Amebis 2015

    This program is free software; you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation; either version 2 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.

    See the GNU General Public License for more details, included in the file
    LICENSE.rtf which you should have received along with this program.

    If you did not receive a copy of the GNU General Public License,
    write to the Free Software Foundation, Inc., 675 Mass Ave,
    Cambridge, MA 02139, USA.

    Amebis can be contacted at http://www.amebis.si
*/

#include "stdafx.h"


namespace MSICA {


////////////////////////////////////////////////////////////////////////////
// COpRegKeySingle
////////////////////////////////////////////////////////////////////////////

COpRegKeySingle::COpRegKeySingle(HKEY hKeyRoot, LPCWSTR pszKeyName, int iTicks) :
    m_hKeyRoot(hKeyRoot),
    COpTypeSingleString(pszKeyName, iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// COpRegKeySrcDst
////////////////////////////////////////////////////////////////////////////

COpRegKeySrcDst::COpRegKeySrcDst(HKEY hKeyRoot, LPCWSTR pszKeyNameSrc, LPCWSTR pszKeyNameDst, int iTicks) :
    m_hKeyRoot(hKeyRoot),
    COpTypeSrcDstString(pszKeyNameSrc, pszKeyNameDst, iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// COpRegKeyCreate
////////////////////////////////////////////////////////////////////////////

COpRegKeyCreate::COpRegKeyCreate(HKEY hKeyRoot, LPCWSTR pszKeyName, int iTicks) : COpRegKeySingle(hKeyRoot, pszKeyName, iTicks)
{
}


HRESULT COpRegKeyCreate::Execute(CSession *pSession)
{
    LONG lResult;
    REGSAM samAdditional = 0;
    ATL::CAtlStringW sPartialName;
    int iStart = 0;

#ifndef _WIN64
    if (IsWow64Process()) {
        // 32-bit processes run as WOW64 should use 64-bit registry too.
        samAdditional |= KEY_WOW64_64KEY;
    }
#endif

    for (;;) {
        HKEY hKey;

        int iStartNext = m_sValue.Find(L'\\', iStart);
        if (iStartNext >= 0)
            sPartialName.SetString(m_sValue, iStartNext);
        else
            sPartialName = m_sValue;

        // Try to open the key, to see if it exists.
        lResult = ::RegOpenKeyExW(m_hKeyRoot, sPartialName, 0, KEY_ENUMERATE_SUB_KEYS | samAdditional, &hKey);
        if (lResult == ERROR_FILE_NOT_FOUND) {
            // The key doesn't exist yet. Create it.

            if (pSession->m_bRollbackEnabled) {
                // Order rollback action to delete the key. ::RegCreateEx() might create a key but return failure.
                pSession->m_olRollback.AddHead(new COpRegKeyDelete(m_hKeyRoot, sPartialName));
            }

            // Create the key.
            lResult = ::RegCreateKeyExW(m_hKeyRoot, sPartialName, NULL, NULL, REG_OPTION_NON_VOLATILE, KEY_ENUMERATE_SUB_KEYS | samAdditional, NULL, &hKey, NULL);
            if (lResult != NO_ERROR) break;
            ::RegCloseKey(hKey);
        } else if (lResult == NO_ERROR) {
            // This key already exists. Release its handle and continue.
            ::RegCloseKey(hKey);
        } else
            break;

        if (iStartNext < 0) break;
        iStart = iStartNext + 1;
    }

    if (lResult == NO_ERROR)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(4);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_CREATE  );
        ::MsiRecordSetInteger(hRecordProg, 2, (UINT)m_hKeyRoot & 0x7fffffff);
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue                     );
        ::MsiRecordSetInteger(hRecordProg, 4, lResult                      );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(lResult);
    }
}


////////////////////////////////////////////////////////////////////////////
// COpRegKeyCopy
////////////////////////////////////////////////////////////////////////////

COpRegKeyCopy::COpRegKeyCopy(HKEY hKeyRoot, LPCWSTR pszKeyNameSrc, LPCWSTR pszKeyNameDst, int iTicks) : COpRegKeySrcDst(hKeyRoot, pszKeyNameSrc, pszKeyNameDst, iTicks)
{
}


HRESULT COpRegKeyCopy::Execute(CSession *pSession)
{
    LONG lResult;
    REGSAM samAdditional = 0;

    {
        // Delete existing destination key first.
        // Since deleting registry key is a complicated job (when rollback/commit support is required), and we do have an operation just for that, we use it.
        // Don't worry, COpRegKeyDelete::Execute() returns S_OK if key doesn't exist.
        COpRegKeyDelete opDelete(m_hKeyRoot, m_sValue2);
        HRESULT hr = opDelete.Execute(pSession);
        if (FAILED(hr)) return hr;
    }

#ifndef _WIN64
    if (IsWow64Process()) {
        // 32-bit processes run as WOW64 should use 64-bit registry too.
        samAdditional |= KEY_WOW64_64KEY;
    }
#endif

    if (pSession->m_bRollbackEnabled) {
        // Order rollback action to delete the destination key.
        pSession->m_olRollback.AddHead(new COpRegKeyDelete(m_hKeyRoot, m_sValue2));
    }

    // Copy the registry key.
    lResult = CopyKeyRecursively(m_hKeyRoot, m_sValue1, m_sValue2, samAdditional);
    if (lResult == NO_ERROR)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(5);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_COPY    );
        ::MsiRecordSetInteger(hRecordProg, 2, (UINT)m_hKeyRoot & 0x7fffffff);
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue1                    );
        ::MsiRecordSetStringW(hRecordProg, 4, m_sValue2                    );
        ::MsiRecordSetInteger(hRecordProg, 5, lResult                      );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(lResult);
    }
}


LONG COpRegKeyCopy::CopyKeyRecursively(HKEY hKeyRoot, LPCWSTR pszKeyNameSrc, LPCWSTR pszKeyNameDst, REGSAM samAdditional)
{
    LONG lResult;
    HKEY hKeySrc, hKeyDst;

    // Open source key.
    lResult = ::RegOpenKeyExW(hKeyRoot, pszKeyNameSrc, 0, READ_CONTROL | KEY_READ | samAdditional, &hKeySrc);
    if (lResult != NO_ERROR) return lResult;

    {
        DWORD dwSecurityDescriptorSize, dwClassLen = MAX_PATH;
        SECURITY_ATTRIBUTES sa = { sizeof(SECURITY_ATTRIBUTES) };
        LPWSTR pszClass = new WCHAR[dwClassLen];

        // Get source key class length and security descriptor size.
        lResult = ::RegQueryInfoKeyW(hKeySrc, pszClass, &dwClassLen, NULL, NULL, NULL, NULL, NULL, NULL, NULL, &dwSecurityDescriptorSize, NULL);
        if (lResult != NO_ERROR) {
            delete [] pszClass;
            return lResult;
        }
        pszClass[dwClassLen] = 0;

        // Get source key security descriptor.
        sa.lpSecurityDescriptor = (PSECURITY_DESCRIPTOR)(new BYTE[dwSecurityDescriptorSize]);
        lResult = ::RegGetKeySecurity(hKeySrc, DACL_SECURITY_INFORMATION, sa.lpSecurityDescriptor, &dwSecurityDescriptorSize);
        if (lResult != NO_ERROR) {
            delete [] (LPBYTE)(sa.lpSecurityDescriptor);
            delete [] pszClass;
            return lResult;
        }

        // Create new destination key of the same class and security.
        lResult = ::RegCreateKeyExW(hKeyRoot, pszKeyNameDst, 0, pszClass, REG_OPTION_NON_VOLATILE, KEY_WRITE | samAdditional, &sa, &hKeyDst, NULL);
        delete [] (LPBYTE)(sa.lpSecurityDescriptor);
        delete [] pszClass;
        if (lResult != NO_ERROR) return lResult;
    }

    // Copy subkey recursively.
    return CopyKeyRecursively(hKeySrc, hKeyDst, samAdditional);
}


LONG COpRegKeyCopy::CopyKeyRecursively(HKEY hKeySrc, HKEY hKeyDst, REGSAM samAdditional)
{
    LONG lResult;
    DWORD dwMaxSubKeyLen, dwMaxValueNameLen, dwMaxClassLen, dwMaxDataSize, dwIndex;
    LPWSTR pszName, pszClass;
    LPBYTE lpData;

    // Query the source key.
    lResult = ::RegQueryInfoKeyW(hKeySrc, NULL, NULL, NULL, NULL, &dwMaxSubKeyLen, &dwMaxClassLen, NULL, &dwMaxValueNameLen, &dwMaxDataSize, NULL, NULL);
    if (lResult != NO_ERROR) return lResult;

    // Copy values first.
    dwMaxValueNameLen++;
    pszName = new WCHAR[dwMaxValueNameLen];
    lpData = new BYTE[dwMaxDataSize];
    for (dwIndex = 0; ; dwIndex++) {
        DWORD dwNameLen = dwMaxValueNameLen, dwType, dwValueSize = dwMaxDataSize;

        // Read value.
        lResult = ::RegEnumValueW(hKeySrc, dwIndex, pszName, &dwNameLen, NULL, &dwType, lpData, &dwValueSize);
        if (lResult == ERROR_NO_MORE_ITEMS) {
            lResult = NO_ERROR;
            break;
        } else if (lResult != NO_ERROR)
            break;

        // Save value.
        lResult = ::RegSetValueExW(hKeyDst, pszName, 0, dwType, lpData, dwValueSize);
        if (lResult != NO_ERROR)
            break;
    }
    delete [] lpData;
    delete [] pszName;
    if (lResult != NO_ERROR) return lResult;

    // Iterate over all subkeys and copy them.
    dwMaxSubKeyLen++;
    pszName = new WCHAR[dwMaxSubKeyLen];
    dwMaxClassLen++;
    pszClass = new WCHAR[dwMaxClassLen];
    for (dwIndex = 0; ; dwIndex++) {
        DWORD dwNameLen = dwMaxSubKeyLen, dwClassLen = dwMaxClassLen;
        HKEY hKeySrcSub, hKeyDstSub;

        // Read subkey.
        lResult = ::RegEnumKeyExW(hKeySrc, dwIndex, pszName, &dwNameLen, NULL, pszClass, &dwClassLen, NULL);
        if (lResult == ERROR_NO_MORE_ITEMS) {
            lResult = NO_ERROR;
            break;
        } else if (lResult != NO_ERROR)
            break;

        // Open source subkey.
        lResult = ::RegOpenKeyExW(hKeySrc, pszName, 0, READ_CONTROL | KEY_READ | samAdditional, &hKeySrcSub);
        if (lResult != NO_ERROR) break;

        {
            DWORD dwSecurityDescriptorSize;
            SECURITY_ATTRIBUTES sa = { sizeof(SECURITY_ATTRIBUTES) };

            // Get source subkey security descriptor size.
            lResult = ::RegQueryInfoKeyW(hKeySrcSub, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, &dwSecurityDescriptorSize, NULL);
            if (lResult != NO_ERROR) break;

            // Get source subkey security descriptor.
            sa.lpSecurityDescriptor = (PSECURITY_DESCRIPTOR)(new BYTE[dwSecurityDescriptorSize]);
            lResult = ::RegGetKeySecurity(hKeySrc, DACL_SECURITY_INFORMATION, sa.lpSecurityDescriptor, &dwSecurityDescriptorSize);
            if (lResult != NO_ERROR) {
                delete [] (LPBYTE)(sa.lpSecurityDescriptor);
                break;
            }

            // Create new destination subkey of the same class and security.
            lResult = ::RegCreateKeyExW(hKeyDst, pszName, 0, pszClass, REG_OPTION_NON_VOLATILE, KEY_WRITE | samAdditional, &sa, &hKeyDstSub, NULL);
            delete [] (LPBYTE)(sa.lpSecurityDescriptor);
            if (lResult != NO_ERROR) break;
        }

        // Copy subkey recursively.
        lResult = CopyKeyRecursively(hKeySrcSub, hKeyDstSub, samAdditional);
        if (lResult != NO_ERROR) break;
    }
    delete [] pszClass;
    delete [] pszName;

    return lResult;
}


////////////////////////////////////////////////////////////////////////////
// COpRegKeyDelete
////////////////////////////////////////////////////////////////////////////

COpRegKeyDelete::COpRegKeyDelete(HKEY hKey, LPCWSTR pszKeyName, int iTicks) : COpRegKeySingle(hKey, pszKeyName, iTicks)
{
}


HRESULT COpRegKeyDelete::Execute(CSession *pSession)
{
    LONG lResult;
    HKEY hKey;
    REGSAM samAdditional = 0;

#ifndef _WIN64
    if (IsWow64Process()) {
        // 32-bit processes run as WOW64 should use 64-bit registry too.
        samAdditional |= KEY_WOW64_64KEY;
    }
#endif

    // Probe to see if the key exists.
    lResult = ::RegOpenKeyExW(m_hKeyRoot, m_sValue, 0, DELETE | samAdditional, &hKey);
    if (lResult == NO_ERROR) {
        ::RegCloseKey(hKey);

        if (pSession->m_bRollbackEnabled) {
            // Make a backup of the key first.
            ATL::CAtlStringW sBackupName;
            UINT uiCount = 0;
            int iLength = m_sValue.GetLength();

            // Trim trailing backslashes.
            while (iLength && m_sValue.GetAt(iLength - 1) == L'\\') iLength--;

            for (;;) {
                HKEY hKey;
                sBackupName.Format(L"%.*ls (orig %u)", iLength, (LPCWSTR)m_sValue, ++uiCount);
                lResult = ::RegOpenKeyExW(m_hKeyRoot, sBackupName, 0, KEY_ENUMERATE_SUB_KEYS | samAdditional, &hKey);
                if (lResult != NO_ERROR) break;
                ::RegCloseKey(hKey);
            }
            if (lResult == ERROR_FILE_NOT_FOUND) {
                // Since copying registry key is a complicated job (when rollback/commit support is required), and we do have an operation just for that, we use it.
                COpRegKeyCopy opCopy(m_hKeyRoot, m_sValue, sBackupName);
                HRESULT hr = opCopy.Execute(pSession);
                if (FAILED(hr)) return hr;

                // Order rollback action to restore the key from backup copy.
                pSession->m_olRollback.AddHead(new COpRegKeyCopy(m_hKeyRoot, sBackupName, m_sValue));

                // Order commit action to delete backup copy.
                pSession->m_olCommit.AddTail(new COpRegKeyDelete(m_hKeyRoot, sBackupName));
            } else {
                PMSIHANDLE hRecordProg = ::MsiCreateRecord(4);
                ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_PROBING );
                ::MsiRecordSetInteger(hRecordProg, 2, (UINT)m_hKeyRoot & 0x7fffffff);
                ::MsiRecordSetStringW(hRecordProg, 3, sBackupName                  );
                ::MsiRecordSetInteger(hRecordProg, 4, lResult                      );
                ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                return AtlHresultFromWin32(lResult);
            }
        }

        // Delete the registry key.
        lResult = DeleteKeyRecursively(m_hKeyRoot, m_sValue, samAdditional);
    }

    if (lResult == NO_ERROR || lResult == ERROR_FILE_NOT_FOUND)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(4);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_DELETE  );
        ::MsiRecordSetInteger(hRecordProg, 2, (UINT)m_hKeyRoot & 0x7fffffff);
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue                     );
        ::MsiRecordSetInteger(hRecordProg, 4, lResult                      );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(lResult);
    }
}


LONG COpRegKeyDelete::DeleteKeyRecursively(HKEY hKeyRoot, LPCWSTR pszKeyName, REGSAM samAdditional)
{
    HKEY hKey;
    LONG lResult;

    // Open the key.
    lResult = ::RegOpenKeyExW(hKeyRoot, pszKeyName, 0, DELETE | KEY_READ | samAdditional, &hKey);
    if (lResult == NO_ERROR) {
        DWORD dwMaxSubKeyLen;

        // Determine the largest subkey name.
        lResult = ::RegQueryInfoKeyW(hKey, NULL, NULL, NULL, NULL, &dwMaxSubKeyLen, NULL, NULL, NULL, NULL, NULL, NULL);
        if (lResult == NO_ERROR) {
            LPWSTR pszSubKeyName;

            // Prepare buffer to hold the subkey names (including zero terminator).
            dwMaxSubKeyLen++;
            pszSubKeyName = new WCHAR[dwMaxSubKeyLen];
            if (pszSubKeyName) {
                DWORD dwIndex;

                // Iterate over all subkeys and delete them. Skip failed.
                for (dwIndex = 0; ;) {
                    DWORD dwNameLen = dwMaxSubKeyLen;
                    lResult = ::RegEnumKeyExW(hKey, dwIndex, pszSubKeyName, &dwNameLen, NULL, NULL, NULL, NULL);
                    if (lResult == NO_ERROR) {
                        lResult = DeleteKeyRecursively(hKey, pszSubKeyName, samAdditional);
                        if (lResult != NO_ERROR)
                            dwIndex++;
                    } else if (lResult == ERROR_NO_MORE_ITEMS) {
                        lResult = NO_ERROR;
                        break;
                    } else
                        dwIndex++;
                }

                delete [] pszSubKeyName;
            } else
                lResult = ERROR_OUTOFMEMORY;
        }

        ::RegCloseKey(hKey);

        // Finally try to delete the key.
        lResult = ::RegDeleteKeyW(hKeyRoot, pszKeyName);
    } else if (lResult == ERROR_FILE_NOT_FOUND) {
        // The key doesn't exist. Not really an error in this case.
        lResult = NO_ERROR;
    }

    return lResult;
}


////////////////////////////////////////////////////////////////////////////
// COpRegValueSingle
////////////////////////////////////////////////////////////////////////////

COpRegValueSingle::COpRegValueSingle(HKEY hKeyRoot, LPCWSTR pszKeyName, LPCWSTR pszValueName, int iTicks) :
    m_sValueName(pszValueName),
    COpRegKeySingle(hKeyRoot, pszKeyName, iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// COpRegValueSrcDst
////////////////////////////////////////////////////////////////////////////

COpRegValueSrcDst::COpRegValueSrcDst(HKEY hKeyRoot, LPCWSTR pszKeyName, LPCWSTR pszValueNameSrc, LPCWSTR pszValueNameDst, int iTicks) :
    m_sValueName1(pszValueNameSrc),
    m_sValueName2(pszValueNameDst),
    COpRegKeySingle(hKeyRoot, pszKeyName, iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// COpRegValueCreate
////////////////////////////////////////////////////////////////////////////

COpRegValueCreate::COpRegValueCreate(HKEY hKeyRoot, LPCWSTR pszKeyName, LPCWSTR pszValueName, int iTicks) :
    m_dwType(REG_NONE),
    COpRegValueSingle(hKeyRoot, pszKeyName, pszValueName, iTicks)
{
}


COpRegValueCreate::COpRegValueCreate(HKEY hKeyRoot, LPCWSTR pszKeyName, LPCWSTR pszValueName, DWORD dwData, int iTicks) :
    m_dwType(REG_DWORD),
    m_dwData(dwData),
    COpRegValueSingle(hKeyRoot, pszKeyName, pszValueName, iTicks)
{
}


COpRegValueCreate::COpRegValueCreate(HKEY hKeyRoot, LPCWSTR pszKeyName, LPCWSTR pszValueName, LPCVOID lpData, SIZE_T nSize, int iTicks) :
    m_dwType(REG_BINARY),
    COpRegValueSingle(hKeyRoot, pszKeyName, pszValueName, iTicks)
{
    m_binData.SetCount(nSize);
    memcpy(m_binData.GetData(), lpData, nSize);
}


COpRegValueCreate::COpRegValueCreate(HKEY hKeyRoot, LPCWSTR pszKeyName, LPCWSTR pszValueName, LPCWSTR pszData, int iTicks) :
    m_dwType(REG_SZ),
    m_sData(pszData),
    COpRegValueSingle(hKeyRoot, pszKeyName, pszValueName, iTicks)
{
}


COpRegValueCreate::COpRegValueCreate(HKEY hKeyRoot, LPCWSTR pszKeyName, LPCWSTR pszValueName, DWORDLONG qwData, int iTicks) :
    m_dwType(REG_QWORD),
    m_qwData(qwData),
    COpRegValueSingle(hKeyRoot, pszKeyName, pszValueName, iTicks)
{
}


HRESULT COpRegValueCreate::Execute(CSession *pSession)
{
    LONG lResult;
    REGSAM sam = KEY_QUERY_VALUE | STANDARD_RIGHTS_WRITE | KEY_SET_VALUE;
    HKEY hKey;

    {
        // Delete existing value first.
        // Since deleting registry value is a complicated job (when rollback/commit support is required), and we do have an operation just for that, we use it.
        // Don't worry, COpRegValueDelete::Execute() returns S_OK if key doesn't exist.
        COpRegValueDelete opDelete(m_hKeyRoot, m_sValue, m_sValueName);
        HRESULT hr = opDelete.Execute(pSession);
        if (FAILED(hr)) return hr;
    }

#ifndef _WIN64
    if (IsWow64Process()) {
        // 32-bit processes run as WOW64 should use 64-bit registry too.
        sam |= KEY_WOW64_64KEY;
    }
#endif

    // Open the key.
    lResult = ::RegOpenKeyExW(m_hKeyRoot, m_sValue, 0, sam, &hKey);
    if (lResult == NO_ERROR) {
        if (pSession->m_bRollbackEnabled) {
            // Order rollback action to delete the value.
            pSession->m_olRollback.AddHead(new COpRegValueDelete(m_hKeyRoot, m_sValue, m_sValueName));
        }

        // Set the registry value.
        switch (m_dwType) {
        case REG_SZ:
        case REG_EXPAND_SZ:
        case REG_LINK:
            lResult = ::RegSetValueExW(hKey, m_sValueName, 0, m_dwType, (const BYTE*)(LPCWSTR)m_sData, (m_sData.GetLength() + 1) * sizeof(WCHAR)); break;
            break;

        case REG_BINARY:
            lResult = ::RegSetValueExW(hKey, m_sValueName, 0, m_dwType, m_binData.GetData(), (DWORD)m_binData.GetCount() * sizeof(BYTE)); break;

        case REG_DWORD_LITTLE_ENDIAN:
        case REG_DWORD_BIG_ENDIAN:
            lResult = ::RegSetValueExW(hKey, m_sValueName, 0, m_dwType, (const BYTE*)&m_dwData, sizeof(DWORD)); break;
            break;

        case REG_MULTI_SZ:
            lResult = ::RegSetValueExW(hKey, m_sValueName, 0, m_dwType, (const BYTE*)m_szData.GetData(), (DWORD)m_szData.GetCount() * sizeof(WCHAR)); break;
            break;

        case REG_QWORD_LITTLE_ENDIAN:
            lResult = ::RegSetValueExW(hKey, m_sValueName, 0, m_dwType, (const BYTE*)&m_qwData, sizeof(DWORDLONG)); break;
            break;

        default:
            lResult = ERROR_UNSUPPORTED_TYPE;
        }

        ::RegCloseKey(hKey);
    }

    if (lResult == NO_ERROR)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(5);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_SETVALUE);
        ::MsiRecordSetInteger(hRecordProg, 2, (UINT)m_hKeyRoot & 0x7fffffff);
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue                     );
        ::MsiRecordSetStringW(hRecordProg, 4, m_sValueName                 );
        ::MsiRecordSetInteger(hRecordProg, 5, lResult                      );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(lResult);
    }
}


////////////////////////////////////////////////////////////////////////////
// COpRegValueCopy
////////////////////////////////////////////////////////////////////////////

COpRegValueCopy::COpRegValueCopy(HKEY hKeyRoot, LPCWSTR pszKeyName, LPCWSTR pszValueNameSrc, LPCWSTR pszValueNameDst, int iTicks) : COpRegValueSrcDst(hKeyRoot, pszKeyName, pszValueNameSrc, pszValueNameDst, iTicks)
{
}


HRESULT COpRegValueCopy::Execute(CSession *pSession)
{
    LONG lResult;
    REGSAM sam = KEY_QUERY_VALUE | KEY_SET_VALUE;
    HKEY hKey;

    {
        // Delete existing destination value first.
        // Since deleting registry value is a complicated job (when rollback/commit support is required), and we do have an operation just for that, we use it.
        // Don't worry, COpRegValueDelete::Execute() returns S_OK if key doesn't exist.
        COpRegValueDelete opDelete(m_hKeyRoot, m_sValue, m_sValueName2);
        HRESULT hr = opDelete.Execute(pSession);
        if (FAILED(hr)) return hr;
    }

#ifndef _WIN64
    if (IsWow64Process()) {
        // 32-bit processes run as WOW64 should use 64-bit registry too.
        sam |= KEY_WOW64_64KEY;
    }
#endif

    // Open the key.
    lResult = ::RegOpenKeyExW(m_hKeyRoot, m_sValue, 0, sam, &hKey);
    if (lResult == NO_ERROR) {
        DWORD dwType, dwSize;

        // Query the source registry value size.
        lResult = ::RegQueryValueExW(hKey, m_sValueName1, 0, NULL, NULL, &dwSize);
        if (lResult == NO_ERROR) {
            LPBYTE lpData = new BYTE[dwSize];
            // Read the source registry value.
            lResult = ::RegQueryValueExW(hKey, m_sValueName1, 0, &dwType, lpData, &dwSize);
            if (lResult == NO_ERROR) {
                if (pSession->m_bRollbackEnabled) {
                    // Order rollback action to delete the destination copy.
                    pSession->m_olRollback.AddHead(new COpRegValueDelete(m_hKeyRoot, m_sValue, m_sValueName2));
                }

                // Store the value to destination.
                lResult = ::RegSetValueExW(hKey, m_sValueName2, 0, dwType, lpData, dwSize);
            }
            delete [] lpData;
        }

        ::RegCloseKey(hKey);
    }

    if (lResult == NO_ERROR)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(6);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_COPYVALUE);
        ::MsiRecordSetInteger(hRecordProg, 2, (UINT)m_hKeyRoot & 0x7fffffff );
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue                      );
        ::MsiRecordSetStringW(hRecordProg, 4, m_sValueName1                 );
        ::MsiRecordSetStringW(hRecordProg, 5, m_sValueName2                 );
        ::MsiRecordSetInteger(hRecordProg, 6, lResult                       );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(lResult);
    }
}


////////////////////////////////////////////////////////////////////////////
// COpRegValueDelete
////////////////////////////////////////////////////////////////////////////

COpRegValueDelete::COpRegValueDelete(HKEY hKeyRoot, LPCWSTR pszKeyName, LPCWSTR pszValueName, int iTicks) : COpRegValueSingle(hKeyRoot, pszKeyName, pszValueName, iTicks)
{
}


HRESULT COpRegValueDelete::Execute(CSession *pSession)
{
    LONG lResult;
    REGSAM sam = KEY_QUERY_VALUE | KEY_SET_VALUE;
    HKEY hKey;

#ifndef _WIN64
    if (IsWow64Process()) {
        // 32-bit processes run as WOW64 should use 64-bit registry too.
        sam |= KEY_WOW64_64KEY;
    }
#endif

    // Open the key.
    lResult = ::RegOpenKeyExW(m_hKeyRoot, m_sValue, 0, sam, &hKey);
    if (lResult == NO_ERROR) {
        DWORD dwType;

        // See if the value exists at all.
        lResult = ::RegQueryValueExW(hKey, m_sValueName, 0, &dwType, NULL, NULL);
        if (lResult == NO_ERROR) {
            if (pSession->m_bRollbackEnabled) {
                // Make a backup of the value first.
                ATL::CAtlStringW sBackupName;
                UINT uiCount = 0;

                for (;;) {
                    sBackupName.Format(L"%ls (orig %u)", (LPCWSTR)m_sValueName, ++uiCount);
                    lResult = ::RegQueryValueExW(hKey, sBackupName, 0, &dwType, NULL, NULL);
                    if (lResult != NO_ERROR) break;
                }
                if (lResult == ERROR_FILE_NOT_FOUND) {
                    // Since copying registry value is a complicated job (when rollback/commit support is required), and we do have an operation just for that, we use it.
                    COpRegValueCopy opCopy(m_hKeyRoot, m_sValue, m_sValueName, sBackupName);
                    HRESULT hr = opCopy.Execute(pSession);
                    if (FAILED(hr)) {
                        ::RegCloseKey(hKey);
                        return hr;
                    }

                    // Order rollback action to restore the key from backup copy.
                    pSession->m_olRollback.AddHead(new COpRegValueCopy(m_hKeyRoot, m_sValue, sBackupName, m_sValueName));

                    // Order commit action to delete backup copy.
                    pSession->m_olCommit.AddTail(new COpRegValueDelete(m_hKeyRoot, m_sValue, sBackupName));
                } else {
                    PMSIHANDLE hRecordProg = ::MsiCreateRecord(5);
                    ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_PROBINGVAL);
                    ::MsiRecordSetInteger(hRecordProg, 2, (UINT)m_hKeyRoot & 0x7fffffff  );
                    ::MsiRecordSetStringW(hRecordProg, 3, m_sValue                       );
                    ::MsiRecordSetStringW(hRecordProg, 3, sBackupName                    );
                    ::MsiRecordSetInteger(hRecordProg, 4, lResult                        );
                    ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                    ::RegCloseKey(hKey);
                    return AtlHresultFromWin32(lResult);
                }
            }

            // Delete the registry value.
            lResult = ::RegDeleteValueW(hKey, m_sValueName);
        }

        ::RegCloseKey(hKey);
    }

    if (lResult == NO_ERROR || lResult == ERROR_FILE_NOT_FOUND)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(5);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_DELETEVALUE);
        ::MsiRecordSetInteger(hRecordProg, 2, (UINT)m_hKeyRoot & 0x7fffffff   );
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue                        );
        ::MsiRecordSetStringW(hRecordProg, 4, m_sValueName                    );
        ::MsiRecordSetInteger(hRecordProg, 5, lResult                         );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(lResult);
    }
}

} // namespace MSICA
