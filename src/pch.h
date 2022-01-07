/*
    SPDX-License-Identifier: GPL-3.0-or-later
    Copyright © 1991-2022 Amebis
*/

#pragma once

#include <MSICALib.h>

#include <WinStd/COM.h>
#include <WinStd/Crypt.h>

#include <msi.h>
#include <msiquery.h>
#include <mstask.h>
#include <shlwapi.h>
#include <taskschd.h>
#include <wlanapi.h>

// Must not statically link to Wlanapi.dll as it is not available on Windows
// without a WLAN interface.
extern VOID  (WINAPI *pfnWlanFreeMemory)(__in PVOID pMemory);
extern DWORD (WINAPI *pfnWlanOpenHandle)(__in DWORD dwClientVersion, __reserved PVOID pReserved, __out PDWORD pdwNegotiatedVersion, __out PHANDLE phClientHandle);
extern DWORD (WINAPI *pfnWlanCloseHandle)(__in HANDLE hClientHandle, __reserved PVOID pReserved);
extern DWORD (WINAPI *pfnWlanGetProfile)(__in HANDLE hClientHandle, __in CONST GUID *pInterfaceGuid, __in LPCWSTR strProfileName, __reserved PVOID pReserved, __deref_out LPWSTR *pstrProfileXml, __inout_opt DWORD *pdwFlags, __out_opt DWORD *pdwGrantedAccess);
extern DWORD (WINAPI *pfnWlanSetProfile)(__in HANDLE hClientHandle, __in CONST GUID *pInterfaceGuid, __in DWORD dwFlags, __in LPCWSTR strProfileXml, __in_opt LPCWSTR strAllUserProfileSecurity, __in BOOL bOverwrite, __reserved PVOID pReserved, __out DWORD *pdwReasonCode);
extern DWORD (WINAPI *pfnWlanDeleteProfile)(__in HANDLE hClientHandle, __in CONST GUID *pInterfaceGuid, __in LPCWSTR strProfileName, __reserved PVOID pReserved);
extern DWORD (WINAPI *pfnWlanReasonCodeToString)(__in DWORD dwReasonCode, __in DWORD dwBufferSize, __in_ecount(dwBufferSize) PWCHAR pStringBuffer, __reserved PVOID pReserved);
