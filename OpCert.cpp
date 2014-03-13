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
    return E_NOTIMPL;
}

} // namespace MSICA
