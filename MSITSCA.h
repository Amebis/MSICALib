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

    UINT MSITSCA_API InstallScheduledTasks(MSIHANDLE hInstall);
    UINT MSITSCA_API CommitScheduledTasks(MSIHANDLE hInstall);
    UINT MSITSCA_API RollbackScheduledTasks(MSIHANDLE hInstall);
    UINT MSITSCA_API RemoveScheduledTasks(MSIHANDLE hInstall);

#ifdef __cplusplus
}
#endif


////////////////////////////////////////////////////////////////////
// Globalne funkcije in spremenljivke
////////////////////////////////////////////////////////////////////

namespace MSITSCA {
    extern HINSTANCE hInstance; // roèica modula
}

#endif // !defined(RC_INVOKED) && !defined(MIDL_PASS)

#endif // __MSITSCA_H__
