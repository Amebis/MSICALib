#ifndef __MSITSCA_H__
#define __MSITSCA_H__


////////////////////////////////////////////////////////////////////////////
// Constants
////////////////////////////////////////////////////////////////////////////

#define MSITSCA_VERSION       0x02000100

#define MSITSCA_VERSION_MAJ   2
#define MSITSCA_VERSION_MIN   0
#define MSITSCA_VERSION_REV   1

#define MSITSCA_VERSION_STR   "2.0.1"
#define MSITSCA_VERSION_INST  "2.0.1"


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

#ifdef __cplusplus
}
#endif

#endif // !defined(RC_INVOKED) && !defined(MIDL_PASS)

#endif // __MSITSCA_H__
