// StdAfx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#if defined(_UNICODE) && !defined(UNICODE)
#define UNICODE
#endif

#if defined(_WIN32) && !defined(WIN32)
#define WIN32
#endif

#ifndef STRICT
#define STRICT
#endif

#define VC_EXTRALEAN                       // Exclude rarely-used stuff from Windows headers
#define _WIN32_WINNT 0x0501                // Include Windows XP symbols
#define _WINSOCKAPI_                       // Prevent inclusion of winsock.h in windows.h
//#define ASSERT_TO_DEBUG                    // Enable false asserts to provide opportunity to attach debuger
#ifdef _WINDLL
#define MSITSCA_DLL                        // This is a DLL
#endif
#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE
#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS // Some CString constructors will be explicit

#include <atlbase.h>
#include <atlfile.h>
#include <atlstr.h>

using namespace ATL;

#include "BuildNum.h"

#include "MSITSCA.h"
#include "MSITSCAOp.h"

#include <assert.h>
#include <corerror.h>
#include <msi.h>
#include <msiquery.h>
#include <mstask.h>
#include <stdlib.h>
#include <taskschd.h>
#include <time.h>

#ifdef NDEBUG
#define verify(expr) ((void)(expr))
#else
#define verify(expr) assert(expr)
#endif
