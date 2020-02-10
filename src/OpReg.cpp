/*
    Copyright 1991-2018 Amebis

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
    std::wstring sPartialName;
    size_t iStart = 0;

#ifndef _WIN64
    if (IsWow64Process()) {
        // 32-bit processes run as WOW64 should use 64-bit registry too.
        samAdditional |= KEY_WOW64_64KEY;
    }
#endif

    for (;;) {
        HKEY hKey;

        size_t iStartNext = m_sValue.find(L'\\', iStart);
        if (iStartNext != std::wstring::npos)
            sPartialName.assign(m_sValue.c_str(), iStartNext);
        else
            sPartialName = m_sValue;

        // Try to open the key, to see if it exists.
        lResult = ::RegOpenKeyExW(m_hKeyRoot, sPartialName.c_str(), 0, KEY_ENUMERATE_SUB_KEYS | samAdditional, &hKey);
        if (lResult == ERROR_FILE_NOT_FOUND) {
            // The key doesn't exist yet. Create it.

            if (pSession->m_bRollbackEnabled) {
                // Order rollback action to delete the key. ::RegCreateEx() might create a key but return failure.
                pSession->m_olRollback.push_front(new COpRegKeyDelete(m_hKeyRoot, sPartialName.c_str()));
            }

            // Create the key.
            lResult = ::RegCreateKeyExW(m_hKeyRoot, sPartialName.c_str(), NULL, NULL, REG_OPTION_NON_VOLATILE, KEY_ENUMERATE_SUB_KEYS | samAdditional, NULL, &hKey, NULL);
            if (lResult != NO_ERROR) break;
            ::RegCloseKey(hKey);
        } else if (lResult == NO_ERROR) {
            // This key already exists. Release its handle and continue.
            ::RegCloseKey(hKey);
        } else
            break;

        if (iStartNext == std::wstring::npos) break;
        iStart = iStartNext + 1;
    }

    if (lResult == NO_ERROR)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(4);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_CREATE           );
        ::MsiRecordSetInteger(hRecordProg, 2, (int)(UINT_PTR)m_hKeyRoot & 0x7fffffff);
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue.c_str()                      );
        ::MsiRecordSetInteger(hRecordProg, 4, lResult                               );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return HRESULT_FROM_WIN32(lResult);
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
        COpRegKeyDelete opDelete(m_hKeyRoot, m_sValue2.c_str());
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
        pSession->m_olRollback.push_front(new COpRegKeyDelete(m_hKeyRoot, m_sValue2.c_str()));
    }

    // Copy the registry key.
    lResult = CopyKeyRecursively(m_hKeyRoot, m_sValue1.c_str(), m_sValue2.c_str(), samAdditional);
    if (lResult == NO_ERROR)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(5);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_COPY             );
        ::MsiRecordSetInteger(hRecordProg, 2, (int)(UINT_PTR)m_hKeyRoot & 0x7fffffff);
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue1.c_str()                     );
        ::MsiRecordSetStringW(hRecordProg, 4, m_sValue2.c_str()                     );
        ::MsiRecordSetInteger(hRecordProg, 5, lResult                               );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return HRESULT_FROM_WIN32(lResult);
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
        std::unique_ptr<WCHAR[]> pszClass(new WCHAR[dwClassLen]);
        if (!pszClass) return ERROR_OUTOFMEMORY;

        // Get source key class length and security descriptor size.
        lResult = ::RegQueryInfoKeyW(hKeySrc, pszClass.get(), &dwClassLen, NULL, NULL, NULL, NULL, NULL, NULL, NULL, &dwSecurityDescriptorSize, NULL);
        if (lResult != NO_ERROR) return lResult;
        pszClass.get()[dwClassLen] = 0;

        // Get source key security descriptor.
        std::unique_ptr<BYTE[]> sd(new BYTE[dwSecurityDescriptorSize]);
        if (!sd) return ERROR_OUTOFMEMORY;
        sa.lpSecurityDescriptor = (PSECURITY_DESCRIPTOR)sd.get();
        lResult = ::RegGetKeySecurity(hKeySrc, DACL_SECURITY_INFORMATION, sa.lpSecurityDescriptor, &dwSecurityDescriptorSize);
        if (lResult != NO_ERROR) return lResult;

        // Create new destination key of the same class and security.
        lResult = ::RegCreateKeyExW(hKeyRoot, pszKeyNameDst, 0, pszClass.get(), REG_OPTION_NON_VOLATILE, KEY_WRITE | samAdditional, &sa, &hKeyDst, NULL);
        if (lResult != NO_ERROR) return lResult;
    }

    // Copy subkey recursively.
    return CopyKeyRecursively(hKeySrc, hKeyDst, samAdditional);
}


LONG COpRegKeyCopy::CopyKeyRecursively(HKEY hKeySrc, HKEY hKeyDst, REGSAM samAdditional)
{
    LONG lResult;
    DWORD dwMaxSubKeyLen, dwMaxValueNameLen, dwMaxClassLen, dwMaxDataSize, dwIndex;

    // Query the source key.
    lResult = ::RegQueryInfoKeyW(hKeySrc, NULL, NULL, NULL, NULL, &dwMaxSubKeyLen, &dwMaxClassLen, NULL, &dwMaxValueNameLen, &dwMaxDataSize, NULL, NULL);
    if (lResult != NO_ERROR) return lResult;

    {
        // Copy values first.
        dwMaxValueNameLen++;
        std::unique_ptr<WCHAR[]> pszName(new WCHAR[dwMaxValueNameLen]);
        if (!pszName) return ERROR_OUTOFMEMORY;
        std::unique_ptr<BYTE[]> lpData(new BYTE[dwMaxDataSize]);
        if (!lpData) return ERROR_OUTOFMEMORY;
        for (dwIndex = 0; ; dwIndex++) {
            DWORD dwNameLen = dwMaxValueNameLen, dwType, dwValueSize = dwMaxDataSize;

            // Read value.
            lResult = ::RegEnumValueW(hKeySrc, dwIndex, pszName.get(), &dwNameLen, NULL, &dwType, lpData.get(), &dwValueSize);
                 if (lResult == ERROR_NO_MORE_ITEMS) break;
            else if (lResult != NO_ERROR           ) return lResult;

            // Save value.
            lResult = ::RegSetValueExW(hKeyDst, pszName.get(), 0, dwType, lpData.get(), dwValueSize);
            if (lResult != NO_ERROR) return lResult;
        }
    }

    {
        // Iterate over all subkeys and copy them.
        dwMaxSubKeyLen++;
        std::unique_ptr<WCHAR[]> pszName(new WCHAR[dwMaxSubKeyLen]);
        if (!pszName) return ERROR_OUTOFMEMORY;
        dwMaxClassLen++;
        std::unique_ptr<WCHAR[]> pszClass(new WCHAR[dwMaxClassLen]);
        if (!pszClass) return ERROR_OUTOFMEMORY;
        for (dwIndex = 0; ; dwIndex++) {
            DWORD dwNameLen = dwMaxSubKeyLen, dwClassLen = dwMaxClassLen;
            HKEY hKeySrcSub, hKeyDstSub;

            // Read subkey.
            lResult = ::RegEnumKeyExW(hKeySrc, dwIndex, pszName.get(), &dwNameLen, NULL, pszClass.get(), &dwClassLen, NULL);
                 if (lResult == ERROR_NO_MORE_ITEMS) break;
            else if (lResult != NO_ERROR           ) return lResult;

            // Open source subkey.
            lResult = ::RegOpenKeyExW(hKeySrc, pszName.get(), 0, READ_CONTROL | KEY_READ | samAdditional, &hKeySrcSub);
            if (lResult != NO_ERROR) return lResult;

            {
                DWORD dwSecurityDescriptorSize;
                SECURITY_ATTRIBUTES sa = { sizeof(SECURITY_ATTRIBUTES) };

                // Get source subkey security descriptor size.
                lResult = ::RegQueryInfoKeyW(hKeySrcSub, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, &dwSecurityDescriptorSize, NULL);
                if (lResult != NO_ERROR) return lResult;

                // Get source subkey security descriptor.
                std::unique_ptr<BYTE[]> sd(new BYTE[dwSecurityDescriptorSize]);
                if (!sd) return ERROR_OUTOFMEMORY;
                sa.lpSecurityDescriptor = (PSECURITY_DESCRIPTOR)sd.get();
                lResult = ::RegGetKeySecurity(hKeySrc, DACL_SECURITY_INFORMATION, sa.lpSecurityDescriptor, &dwSecurityDescriptorSize);
                if (lResult != NO_ERROR) return lResult;

                // Create new destination subkey of the same class and security.
                lResult = ::RegCreateKeyExW(hKeyDst, pszName.get(), 0, pszClass.get(), REG_OPTION_NON_VOLATILE, KEY_WRITE | samAdditional, &sa, &hKeyDstSub, NULL);
                if (lResult != NO_ERROR) return lResult;
            }

            // Copy subkey recursively.
            lResult = CopyKeyRecursively(hKeySrcSub, hKeyDstSub, samAdditional);
            if (lResult != NO_ERROR) return lResult;
        }
    }

    return NO_ERROR;
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
    lResult = ::RegOpenKeyExW(m_hKeyRoot, m_sValue.c_str(), 0, DELETE | samAdditional, &hKey);
    if (lResult == NO_ERROR) {
        ::RegCloseKey(hKey);

        if (pSession->m_bRollbackEnabled) {
            // Make a backup of the key first.
            std::wstring sBackupName;
            UINT uiCount = 0;
            auto iLength = m_sValue.length();

            // Trim trailing backslashes.
            while (iLength && m_sValue[iLength - 1] == L'\\') iLength--;

            for (;;) {
                HKEY hKeyTest;
                sprintf(sBackupName, L"%.*ls (orig %u)", iLength, m_sValue.c_str(), ++uiCount);
                lResult = ::RegOpenKeyExW(m_hKeyRoot, sBackupName.c_str(), 0, KEY_ENUMERATE_SUB_KEYS | samAdditional, &hKeyTest);
                if (lResult != NO_ERROR) break;
                ::RegCloseKey(hKeyTest);
            }
            if (lResult == ERROR_FILE_NOT_FOUND) {
                // Since copying registry key is a complicated job (when rollback/commit support is required), and we do have an operation just for that, we use it.
                COpRegKeyCopy opCopy(m_hKeyRoot, m_sValue.c_str(), sBackupName.c_str());
                HRESULT hr = opCopy.Execute(pSession);
                if (FAILED(hr)) return hr;

                // Order rollback action to restore the key from backup copy.
                pSession->m_olRollback.push_front(new COpRegKeyCopy(m_hKeyRoot, sBackupName.c_str(), m_sValue.c_str()));

                // Order commit action to delete backup copy.
                pSession->m_olCommit.push_back(new COpRegKeyDelete(m_hKeyRoot, sBackupName.c_str()));
            } else {
                PMSIHANDLE hRecordProg = ::MsiCreateRecord(4);
                ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_PROBING          );
                ::MsiRecordSetInteger(hRecordProg, 2, (int)(UINT_PTR)m_hKeyRoot & 0x7fffffff);
                ::MsiRecordSetStringW(hRecordProg, 3, sBackupName.c_str()                   );
                ::MsiRecordSetInteger(hRecordProg, 4, lResult                               );
                ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                return HRESULT_FROM_WIN32(lResult);
            }
        }

        // Delete the registry key.
        lResult = DeleteKeyRecursively(m_hKeyRoot, m_sValue.c_str(), samAdditional);
    }

    if (lResult == NO_ERROR || lResult == ERROR_FILE_NOT_FOUND)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(4);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_DELETE           );
        ::MsiRecordSetInteger(hRecordProg, 2, (int)(UINT_PTR)m_hKeyRoot & 0x7fffffff);
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue.c_str()                      );
        ::MsiRecordSetInteger(hRecordProg, 4, lResult                               );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return HRESULT_FROM_WIN32(lResult);
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
            // Prepare buffer to hold the subkey names (including zero terminator).
            dwMaxSubKeyLen++;
            std::unique_ptr<WCHAR[]> pszSubKeyName(new WCHAR[dwMaxSubKeyLen]);
            if (pszSubKeyName) {
                DWORD dwIndex;

                // Iterate over all subkeys and delete them. Skip failed.
                for (dwIndex = 0; ;) {
                    DWORD dwNameLen = dwMaxSubKeyLen;
                    lResult = ::RegEnumKeyExW(hKey, dwIndex, pszSubKeyName.get(), &dwNameLen, NULL, NULL, NULL, NULL);
                    if (lResult == NO_ERROR) {
                        lResult = DeleteKeyRecursively(hKey, pszSubKeyName.get(), samAdditional);
                        if (lResult != NO_ERROR)
                            dwIndex++;
                    } else if (lResult == ERROR_NO_MORE_ITEMS) {
                        lResult = NO_ERROR;
                        break;
                    } else
                        dwIndex++;
                }
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
    m_binData(reinterpret_cast<LPCBYTE>(lpData), reinterpret_cast<LPCBYTE>(lpData) + nSize),
    COpRegValueSingle(hKeyRoot, pszKeyName, pszValueName, iTicks)
{
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
        COpRegValueDelete opDelete(m_hKeyRoot, m_sValue.c_str(), m_sValueName.c_str());
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
    lResult = ::RegOpenKeyExW(m_hKeyRoot, m_sValue.c_str(), 0, sam, &hKey);
    if (lResult == NO_ERROR) {
        if (pSession->m_bRollbackEnabled) {
            // Order rollback action to delete the value.
            pSession->m_olRollback.push_front(new COpRegValueDelete(m_hKeyRoot, m_sValue.c_str(), m_sValueName.c_str()));
        }

        // Set the registry value.
        switch (m_dwType) {
        case REG_SZ:
        case REG_EXPAND_SZ:
        case REG_LINK:
            lResult = ::RegSetValueExW(hKey, m_sValueName.c_str(), 0, m_dwType, (const BYTE*)m_sData.c_str(), static_cast<DWORD>((m_sData.length() + 1) * sizeof(WCHAR))); break;
            break;

        case REG_BINARY:
            lResult = ::RegSetValueExW(hKey, m_sValueName.c_str(), 0, m_dwType, m_binData.data(), static_cast<DWORD>(m_binData.size() * sizeof(BYTE))); break;

        case REG_DWORD_LITTLE_ENDIAN:
        case REG_DWORD_BIG_ENDIAN:
            lResult = ::RegSetValueExW(hKey, m_sValueName.c_str(), 0, m_dwType, (const BYTE*)&m_dwData, sizeof(DWORD)); break;
            break;

        case REG_MULTI_SZ:
            lResult = ::RegSetValueExW(hKey, m_sValueName.c_str(), 0, m_dwType, (const BYTE*)m_szData.data(), static_cast<DWORD>(m_szData.size() * sizeof(WCHAR))); break;
            break;

        case REG_QWORD_LITTLE_ENDIAN:
            lResult = ::RegSetValueExW(hKey, m_sValueName.c_str(), 0, m_dwType, (const BYTE*)&m_qwData, sizeof(DWORDLONG)); break;
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
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_SETVALUE         );
        ::MsiRecordSetInteger(hRecordProg, 2, (int)(UINT_PTR)m_hKeyRoot & 0x7fffffff);
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue.c_str()                      );
        ::MsiRecordSetStringW(hRecordProg, 4, m_sValueName.c_str()                  );
        ::MsiRecordSetInteger(hRecordProg, 5, lResult                               );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return HRESULT_FROM_WIN32(lResult);
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
        COpRegValueDelete opDelete(m_hKeyRoot, m_sValue.c_str(), m_sValueName2.c_str());
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
    lResult = ::RegOpenKeyExW(m_hKeyRoot, m_sValue.c_str(), 0, sam, &hKey);
    if (lResult == NO_ERROR) {
        DWORD dwType, dwSize;

        // Query the source registry value size.
        lResult = ::RegQueryValueExW(hKey, m_sValueName1.c_str(), 0, NULL, NULL, &dwSize);
        if (lResult == NO_ERROR) {
            // Read the source registry value.
            std::unique_ptr<BYTE> lpData(new BYTE[dwSize]);
            if (lpData) {
                lResult = ::RegQueryValueExW(hKey, m_sValueName1.c_str(), 0, &dwType, lpData.get(), &dwSize);
                if (lResult == NO_ERROR) {
                    if (pSession->m_bRollbackEnabled) {
                        // Order rollback action to delete the destination copy.
                        pSession->m_olRollback.push_front(new COpRegValueDelete(m_hKeyRoot, m_sValue.c_str(), m_sValueName2.c_str()));
                    }

                    // Store the value to destination.
                    lResult = ::RegSetValueExW(hKey, m_sValueName2.c_str(), 0, dwType, lpData.get(), dwSize);
                }
            } else
                lResult = ERROR_OUTOFMEMORY;
        }

        ::RegCloseKey(hKey);
    }

    if (lResult == NO_ERROR)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(6);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_COPYVALUE        );
        ::MsiRecordSetInteger(hRecordProg, 2, (int)(UINT_PTR)m_hKeyRoot & 0x7fffffff);
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue.c_str()                      );
        ::MsiRecordSetStringW(hRecordProg, 4, m_sValueName1.c_str()                 );
        ::MsiRecordSetStringW(hRecordProg, 5, m_sValueName2.c_str()                 );
        ::MsiRecordSetInteger(hRecordProg, 6, lResult                               );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return HRESULT_FROM_WIN32(lResult);
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
    lResult = ::RegOpenKeyExW(m_hKeyRoot, m_sValue.c_str(), 0, sam, &hKey);
    if (lResult == NO_ERROR) {
        DWORD dwType;

        // See if the value exists at all.
        lResult = ::RegQueryValueExW(hKey, m_sValueName.c_str(), 0, &dwType, NULL, NULL);
        if (lResult == NO_ERROR) {
            if (pSession->m_bRollbackEnabled) {
                // Make a backup of the value first.
                std::wstring sBackupName;
                UINT uiCount = 0;

                for (;;) {
                    sprintf(sBackupName, L"%ls (orig %u)", m_sValueName.c_str(), ++uiCount);
                    lResult = ::RegQueryValueExW(hKey, sBackupName.c_str(), 0, &dwType, NULL, NULL);
                    if (lResult != NO_ERROR) break;
                }
                if (lResult == ERROR_FILE_NOT_FOUND) {
                    // Since copying registry value is a complicated job (when rollback/commit support is required), and we do have an operation just for that, we use it.
                    COpRegValueCopy opCopy(m_hKeyRoot, m_sValue.c_str(), m_sValueName.c_str(), sBackupName.c_str());
                    HRESULT hr = opCopy.Execute(pSession);
                    if (FAILED(hr)) {
                        ::RegCloseKey(hKey);
                        return hr;
                    }

                    // Order rollback action to restore the key from backup copy.
                    pSession->m_olRollback.push_front(new COpRegValueCopy(m_hKeyRoot, m_sValue.c_str(), sBackupName.c_str(), m_sValueName.c_str()));

                    // Order commit action to delete backup copy.
                    pSession->m_olCommit.push_back(new COpRegValueDelete(m_hKeyRoot, m_sValue.c_str(), sBackupName.c_str()));
                } else {
                    PMSIHANDLE hRecordProg = ::MsiCreateRecord(5);
                    ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_PROBINGVAL       );
                    ::MsiRecordSetInteger(hRecordProg, 2, (int)(UINT_PTR)m_hKeyRoot & 0x7fffffff);
                    ::MsiRecordSetStringW(hRecordProg, 3, m_sValue.c_str()                      );
                    ::MsiRecordSetStringW(hRecordProg, 3, sBackupName.c_str()                   );
                    ::MsiRecordSetInteger(hRecordProg, 4, lResult                               );
                    ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
                    ::RegCloseKey(hKey);
                    return HRESULT_FROM_WIN32(lResult);
                }
            }

            // Delete the registry value.
            lResult = ::RegDeleteValueW(hKey, m_sValueName.c_str());
        }

        ::RegCloseKey(hKey);
    }

    if (lResult == NO_ERROR || lResult == ERROR_FILE_NOT_FOUND)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(5);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_REGKEY_DELETEVALUE      );
        ::MsiRecordSetInteger(hRecordProg, 2, (int)(UINT_PTR)m_hKeyRoot & 0x7fffffff);
        ::MsiRecordSetStringW(hRecordProg, 3, m_sValue.c_str()                      );
        ::MsiRecordSetStringW(hRecordProg, 4, m_sValueName.c_str()                  );
        ::MsiRecordSetInteger(hRecordProg, 5, lResult                               );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return HRESULT_FROM_WIN32(lResult);
    }
}

} // namespace MSICA
