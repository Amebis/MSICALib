/*
    Copyright © 1991-2021 Amebis

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

#pragma comment(lib, "crypt32.lib")


namespace MSICA {

////////////////////////////////////////////////////////////////////////////
// COpCertStore
////////////////////////////////////////////////////////////////////////////

COpCertStore::COpCertStore(LPCWSTR pszStore, DWORD dwEncodingType, DWORD dwFlags, int iTicks) :
    COpTypeSingleString(pszStore, iTicks),
    m_dwEncodingType(dwEncodingType),
    m_dwFlags(dwFlags)
{
}


////////////////////////////////////////////////////////////////////////////
// COpCert
////////////////////////////////////////////////////////////////////////////

COpCert::COpCert(LPCVOID lpCert, SIZE_T nSize, LPCWSTR pszStore, DWORD dwEncodingType, DWORD dwFlags, int iTicks) :
    m_binCert(reinterpret_cast<LPCBYTE>(lpCert), reinterpret_cast<LPCBYTE>(lpCert) + nSize),
    COpCertStore(pszStore, dwEncodingType, dwFlags, iTicks)
{
}


////////////////////////////////////////////////////////////////////////////
// COpCertInstall
////////////////////////////////////////////////////////////////////////////

COpCertInstall::COpCertInstall(LPCVOID lpCert, SIZE_T nSize, LPCWSTR pszStore, DWORD dwEncodingType, DWORD dwFlags, int iTicks) : COpCert(lpCert, nSize, pszStore, dwEncodingType, dwFlags, iTicks)
{
}


HRESULT COpCertInstall::Execute(CSession *pSession)
{
    DWORD dwError;
    HCERTSTORE hCertStore;

    // Open certificate store.
    hCertStore = ::CertOpenStore(CERT_STORE_PROV_SYSTEM_W, m_dwEncodingType, NULL, m_dwFlags, m_sValue.c_str());
    if (hCertStore) {
        // Create certificate context.
        PCCERT_CONTEXT pCertContext = ::CertCreateCertificateContext(m_dwEncodingType, m_binCert.data(), static_cast<DWORD>(m_binCert.size()));
        if (pCertContext) {
            PMSIHANDLE hRecordMsg = ::MsiCreateRecord(1);
            std::wstring sCertName;

            // Display our custom message in the progress bar.
            ::CertGetNameStringW(pCertContext, CERT_NAME_FRIENDLY_DISPLAY_TYPE, 0, NULL, sCertName);
            ::MsiRecordSetStringW(hRecordMsg, 1, sCertName.c_str());
            if (MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ACTIONDATA, hRecordMsg) == IDCANCEL)
                return HRESULT_FROM_WIN32(ERROR_INSTALL_USEREXIT);

            // Add certificate to certificate store.
            if (::CertAddCertificateContextToStore(hCertStore, pCertContext, CERT_STORE_ADD_NEW, NULL)) {
                if (pSession->m_bRollbackEnabled) {
                    // Order rollback action to delete the certificate.
                    pSession->m_olRollback.push_front(new COpCertRemove(m_binCert.data(), m_binCert.size(), m_sValue.c_str(), m_dwEncodingType, m_dwFlags));
                }
                dwError = NO_ERROR;
            } else {
                dwError = ::GetLastError();
                if (dwError == CRYPT_E_EXISTS) {
                    // Certificate store already contains given certificate. Nothing to install then.
                    dwError = NO_ERROR;
                }
            }
            ::CertFreeCertificateContext(pCertContext);
        } else
            dwError = ::GetLastError();
        ::CertCloseStore(hCertStore, 0);
    } else
        dwError = ::GetLastError();

    if (dwError == NO_ERROR)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_CERT_INSTALL);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue.c_str()          );
        ::MsiRecordSetInteger(hRecordProg, 3, dwError                   );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return HRESULT_FROM_WIN32(dwError);
    }
}


////////////////////////////////////////////////////////////////////////////
// COpCertRemove
////////////////////////////////////////////////////////////////////////////

COpCertRemove::COpCertRemove(LPCVOID lpCert, SIZE_T nSize, LPCWSTR pszStore, DWORD dwEncodingType, DWORD dwFlags, int iTicks) : COpCert(lpCert, nSize, pszStore, dwEncodingType, dwFlags, iTicks)
{
}


HRESULT COpCertRemove::Execute(CSession *pSession)
{
    DWORD dwError;
    HCERTSTORE hCertStore;

    // Open certificate store.
    hCertStore = ::CertOpenStore(CERT_STORE_PROV_SYSTEM_W, m_dwEncodingType, NULL, m_dwFlags, m_sValue.c_str());
    if (hCertStore) {
        // Create certificate context.
        PCCERT_CONTEXT pCertContext = ::CertCreateCertificateContext(m_dwEncodingType, m_binCert.data(), static_cast<DWORD>(m_binCert.size()));
        if (pCertContext) {
            PMSIHANDLE hRecordMsg = ::MsiCreateRecord(1);
            std::wstring sCertName;
            PCCERT_CONTEXT pCertContextExisting;

            // Display our custom message in the progress bar.
            ::CertGetNameStringW(pCertContext, CERT_NAME_FRIENDLY_DISPLAY_TYPE, 0, NULL, sCertName);
            ::MsiRecordSetStringW(hRecordMsg, 1, sCertName.c_str());
            if (MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ACTIONDATA, hRecordMsg) == IDCANCEL)
                return HRESULT_FROM_WIN32(ERROR_INSTALL_USEREXIT);

            pCertContextExisting = ::CertFindCertificateInStore(hCertStore, m_dwEncodingType, 0, CERT_FIND_EXISTING, pCertContext, NULL);
            if (pCertContextExisting) {
                // Delete certificate from certificate store.
                if (::CertDeleteCertificateFromStore(pCertContextExisting)) {
                    if (pSession->m_bRollbackEnabled) {
                        // Order rollback action to reinstall the certificate.
                        pSession->m_olRollback.push_front(new COpCertInstall(m_binCert.data(), m_binCert.size(), m_sValue.c_str(), m_dwEncodingType, m_dwFlags));
                    }
                    dwError = NO_ERROR;
                } else {
                    dwError = ::GetLastError();
                    ::CertFreeCertificateContext(pCertContextExisting);
                }
            } else {
                // We haven't found the certificate. Nothing to delete then.
                dwError = NO_ERROR;
            }
            ::CertFreeCertificateContext(pCertContext);
        } else
            dwError = ::GetLastError();
        ::CertCloseStore(hCertStore, 0);
    } else
        dwError = ::GetLastError();

    if (dwError == NO_ERROR)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_CERT_REMOVE);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue.c_str()         );
        ::MsiRecordSetInteger(hRecordProg, 3, dwError                  );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return HRESULT_FROM_WIN32(dwError);
    }
}

} // namespace MSICA
