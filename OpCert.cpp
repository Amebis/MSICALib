#include "stdafx.h"

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
    COpCertStore(pszStore, dwEncodingType, dwFlags, iTicks)
{
    m_binCert.SetCount(nSize);
    memcpy(m_binCert.GetData(), lpCert, nSize);
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
    hCertStore = ::CertOpenStore(CERT_STORE_PROV_SYSTEM_W, m_dwEncodingType, NULL, m_dwFlags, m_sValue);
    if (hCertStore) {
        // Create certificate context.
        PCCERT_CONTEXT pCertContext = ::CertCreateCertificateContext(m_dwEncodingType, m_binCert.GetData(), (DWORD)(m_binCert.GetCount()));
        if (pCertContext) {
            PMSIHANDLE hRecordMsg = ::MsiCreateRecord(1);
            ATL::CAtlStringW sCertName;

            // Display our custom message in the progress bar.
            ::CertGetNameStringW(pCertContext, CERT_NAME_FRIENDLY_DISPLAY_TYPE, 0, NULL, sCertName);
            ::MsiRecordSetStringW(hRecordMsg, 1, sCertName);
            if (MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ACTIONDATA, hRecordMsg) == IDCANCEL)
                return AtlHresultFromWin32(ERROR_INSTALL_USEREXIT);

            // Add certificate to certificate store.
            if (::CertAddCertificateContextToStore(hCertStore, pCertContext, CERT_STORE_ADD_NEW, NULL)) {
                if (pSession->m_bRollbackEnabled) {
                    // Order rollback action to delete the certificate.
                    pSession->m_olRollback.AddHead(new COpCertRemove(m_binCert.GetData(), m_binCert.GetCount(), m_sValue, m_dwEncodingType, m_dwFlags));
                }
                dwError = ERROR_SUCCESS;
            } else {
                dwError = ::GetLastError();
                if (dwError == CRYPT_E_EXISTS) {
                    // Certificate store already contains given certificate. Nothing to install then.
                    dwError = ERROR_SUCCESS;
                }
            }
            ::CertFreeCertificateContext(pCertContext);
        } else
            dwError = ::GetLastError();
        ::CertCloseStore(hCertStore, 0);
    } else
        dwError = ::GetLastError();

    if (dwError == ERROR_SUCCESS)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_CERT_INSTALL_FAILED);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue                         );
        ::MsiRecordSetInteger(hRecordProg, 3, dwError                          );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(dwError);
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
    hCertStore = ::CertOpenStore(CERT_STORE_PROV_SYSTEM_W, m_dwEncodingType, NULL, m_dwFlags, m_sValue);
    if (hCertStore) {
        // Create certificate context.
        PCCERT_CONTEXT pCertContext = ::CertCreateCertificateContext(m_dwEncodingType, m_binCert.GetData(), (DWORD)(m_binCert.GetCount()));
        if (pCertContext) {
            PMSIHANDLE hRecordMsg = ::MsiCreateRecord(1);
            ATL::CAtlStringW sCertName;
            PCCERT_CONTEXT pCertContextExisting;

            // Display our custom message in the progress bar.
            ::CertGetNameStringW(pCertContext, CERT_NAME_FRIENDLY_DISPLAY_TYPE, 0, NULL, sCertName);
            ::MsiRecordSetStringW(hRecordMsg, 1, sCertName);
            if (MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ACTIONDATA, hRecordMsg) == IDCANCEL)
                return AtlHresultFromWin32(ERROR_INSTALL_USEREXIT);

            pCertContextExisting = ::CertFindCertificateInStore(hCertStore, m_dwEncodingType, 0, CERT_FIND_EXISTING, pCertContext, NULL);
            if (pCertContextExisting) {
                // Delete certificate from certificate store.
                if (::CertDeleteCertificateFromStore(pCertContextExisting)) {
                    if (pSession->m_bRollbackEnabled) {
                        // Order rollback action to reinstall the certificate.
                        pSession->m_olRollback.AddHead(new COpCertInstall(m_binCert.GetData(), m_binCert.GetCount(), m_sValue, m_dwEncodingType, m_dwFlags));
                    }
                    dwError = ERROR_SUCCESS;
                } else {
                    dwError = ::GetLastError();
                    ::CertFreeCertificateContext(pCertContextExisting);
                }
            } else {
                // We haven't found the certificate. Nothing to delete then.
                dwError = ERROR_SUCCESS;
            }
            ::CertFreeCertificateContext(pCertContext);
        } else
            dwError = ::GetLastError();
        ::CertCloseStore(hCertStore, 0);
    } else
        dwError = ::GetLastError();

    if (dwError == ERROR_SUCCESS)
        return S_OK;
    else {
        PMSIHANDLE hRecordProg = ::MsiCreateRecord(3);
        ::MsiRecordSetInteger(hRecordProg, 1, ERROR_INSTALL_CERT_REMOVE_FAILED);
        ::MsiRecordSetStringW(hRecordProg, 2, m_sValue                        );
        ::MsiRecordSetInteger(hRecordProg, 3, dwError                         );
        ::MsiProcessMessage(pSession->m_hInstall, INSTALLMESSAGE_ERROR, hRecordProg);
        return AtlHresultFromWin32(dwError);
    }
}

} // namespace MSICA
