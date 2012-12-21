#ifndef __MSITSCAOP_H__
#define __MSITSCAOP_H__

#include "MSITSCA.h"
#include <atlbase.h>
#include <atlcoll.h>
#include <atlfile.h>
#include <atlstr.h>
#include <mstask.h>
#include <windows.h>

class CMSITSCASession;


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOp
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOp
{
public:
    CMSITSCAOp();

    virtual HRESULT Execute(CMSITSCASession *pSession) = 0;
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpSingleStringOperation
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpSingleStringOperation : public CMSITSCAOp
{
public:
    CMSITSCAOpSingleStringOperation(LPCWSTR pszValue = L"");

    friend inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpSingleStringOperation &op);
    friend inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpSingleStringOperation &op);

protected:
    CStringW m_sValue;
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpDoubleStringOperation
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpSrcDstStringOperation : public CMSITSCAOp
{
public:
    CMSITSCAOpSrcDstStringOperation(LPCWSTR pszValue1 = L"", LPCWSTR pszValue2 = L"");

    friend inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpSrcDstStringOperation &op);
    friend inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpSrcDstStringOperation &op);

protected:
    CStringW m_sValue1;
    CStringW m_sValue2;
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpBooleanOperation
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpBooleanOperation : public CMSITSCAOp
{
public:
    CMSITSCAOpBooleanOperation(BOOL bValue = TRUE);

    friend inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpBooleanOperation &op);
    friend inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpBooleanOperation &op);

protected:
    BOOL m_bValue;
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpEnableRollback
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpEnableRollback : public CMSITSCAOpBooleanOperation
{
public:
    CMSITSCAOpEnableRollback(BOOL bEnable = TRUE);
    virtual HRESULT Execute(CMSITSCASession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpDeleteFile
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpDeleteFile : public CMSITSCAOpSingleStringOperation
{
public:
    CMSITSCAOpDeleteFile(LPCWSTR pszFileName = L"");
    virtual HRESULT Execute(CMSITSCASession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpMoveFile
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpMoveFile : public CMSITSCAOpSrcDstStringOperation
{
public:
    CMSITSCAOpMoveFile(LPCWSTR pszFileSrc = L"", LPCWSTR pszFileDst = L"");
    virtual HRESULT Execute(CMSITSCASession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpCreateTask
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpCreateTask : public CMSITSCAOpSingleStringOperation
{
public:
    CMSITSCAOpCreateTask(LPCWSTR pszTaskName = L"");
    virtual ~CMSITSCAOpCreateTask();

    virtual HRESULT Execute(CMSITSCASession *pSession);

    UINT SetFromRecord(MSIHANDLE hInstall, MSIHANDLE hRecord);
    UINT SetTriggersFromView(MSIHANDLE hView);

    friend inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpCreateTask &op);
    friend inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpCreateTask &op);

protected:
    BOOL     m_bForce;
    CStringW m_sApplicationName;
    CStringW m_sParameters;
    CStringW m_sWorkingDirectory;
    CStringW m_sComment;
    DWORD    m_dwFlags;
    DWORD    m_dwPriority;
    CStringW m_sAccountName;
    CStringW m_sPassword;
    WORD     m_wIdleMinutes;
    WORD     m_wDeadlineMinutes;
    DWORD    m_dwMaxRuntimeMS;

    CAtlList<TASK_TRIGGER> m_lTriggers;
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpDeleteTask
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpDeleteTask : public CMSITSCAOpSingleStringOperation
{
public:
    CMSITSCAOpDeleteTask(LPCWSTR pszTaskName = L"");
    virtual HRESULT Execute(CMSITSCASession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpEnableTask
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpEnableTask : public CMSITSCAOpSingleStringOperation
{
public:
    CMSITSCAOpEnableTask(LPCWSTR pszTaskName = L"", BOOL bEnable = TRUE);
    virtual HRESULT Execute(CMSITSCASession *pSession);

    friend inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpEnableTask &op);
    friend inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpEnableTask &op);

protected:
    BOOL m_bEnable;
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpCopyTask
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpCopyTask : public CMSITSCAOpSrcDstStringOperation
{
public:
    CMSITSCAOpCopyTask(LPCWSTR pszTaskSrc = L"", LPCWSTR pszTaskDst = L"");
    virtual HRESULT Execute(CMSITSCASession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpList
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpList : public CMSITSCAOp, public CAtlList<CMSITSCAOp*>
{
public:
    CMSITSCAOpList();

    void Free();
    HRESULT LoadFromFile(LPCTSTR pszFileName);
    HRESULT SaveToFile(LPCTSTR pszFileName) const;

    virtual HRESULT Execute(CMSITSCASession *pSession);

    friend inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpList &op);
    friend inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpList &op);

protected:
    enum OPERATION {
        OPERATION_ENABLE_ROLLBACK = 1,
        OPERATION_DELETE_FILE,
        OPERATION_MOVE_FILE,
        OPERATION_CREATE_TASK,
        OPERATION_DELETE_TASK,
        OPERATION_ENABLE_TASK,
        OPERATION_COPY_TASK,
        OPERATION_SUBLIST
    };

protected:
    template <class T, int ID> inline static HRESULT Save(CAtlFile &f, const CMSITSCAOp *p);
    template <class T> inline HRESULT LoadAndAddTail(CAtlFile &f);
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCASession
////////////////////////////////////////////////////////////////////////////

class CMSITSCASession
{
public:
    CMSITSCASession();
    virtual ~CMSITSCASession();

    HRESULT Initialize();

    CComPtr<ITaskScheduler> m_pTaskScheduler;   // Task scheduler interface
    BOOL m_bContinueOnError;                    // Continue execution on operation error?
    BOOL m_bRollbackEnabled;                    // Is rollback enabled?
    CMSITSCAOpList m_olRollback;                // Rollback operation list
    CMSITSCAOpList m_olCommit;                  // Commit operation list
};


////////////////////////////////////////////////////////////////////////////
// Inline operators
////////////////////////////////////////////////////////////////////////////

inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpSingleStringOperation &op)
{
    return f << op.m_sValue;
}


inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpSingleStringOperation &op)
{
    return f >> op.m_sValue;
}


inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpSrcDstStringOperation &op)
{
    HRESULT hr;

    hr = f << op.m_sValue1;
    if (FAILED(hr)) return hr;

    return f << op.m_sValue2;
}


inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpSrcDstStringOperation &op)
{
    HRESULT hr;

    hr = f >> op.m_sValue1;
    if (FAILED(hr)) return hr;

    return f >> op.m_sValue2;
}


inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpBooleanOperation &op)
{
    return f << (int)op.m_bValue;
}


inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpBooleanOperation &op)
{
    int iValue;
    HRESULT hr;

    hr = f >> iValue;
    if (FAILED(hr)) return hr;
    op.m_bValue = iValue ? TRUE : FALSE;

    return S_OK;
}


inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpCreateTask &op)
{
    HRESULT hr;
    POSITION pos;

    hr = f << (const CMSITSCAOpSingleStringOperation&)op;              if (FAILED(hr)) return hr;
    hr = f << (int)op.m_bForce;                                        if (FAILED(hr)) return hr;
    hr = f << op.m_sApplicationName;                                   if (FAILED(hr)) return hr;
    hr = f << op.m_sParameters;                                        if (FAILED(hr)) return hr;
    hr = f << op.m_sWorkingDirectory;                                  if (FAILED(hr)) return hr;
    hr = f << op.m_sComment;                                           if (FAILED(hr)) return hr;
    hr = f << (int)op.m_dwFlags;                                       if (FAILED(hr)) return hr;
    hr = f << (int)op.m_dwPriority;                                    if (FAILED(hr)) return hr;
    hr = f << op.m_sAccountName;                                       if (FAILED(hr)) return hr;
    hr = f << op.m_sPassword;                                          if (FAILED(hr)) return hr;
    hr = f << (int)MAKELONG(op.m_wDeadlineMinutes, op.m_wIdleMinutes); if (FAILED(hr)) return hr;
    hr = f << (int)op.m_dwMaxRuntimeMS;                                if (FAILED(hr)) return hr;
    hr = f << (int)op.m_lTriggers.GetCount();                          if (FAILED(hr)) return hr;
    for (pos = op.m_lTriggers.GetHeadPosition(); pos;) {
        hr = f << op.m_lTriggers.GetNext(pos);
        if (FAILED(hr)) return hr;
    }

    return S_OK;
}


inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpCreateTask &op)
{
    HRESULT hr;
    DWORD dwValue;

    hr = f >> (CMSITSCAOpSingleStringOperation&)op; if (FAILED(hr)) return hr;
    hr = f >> (int&)dwValue;                        if (FAILED(hr)) return hr; op.m_bForce = dwValue ? TRUE : FALSE;
    hr = f >> op.m_sApplicationName;                if (FAILED(hr)) return hr;
    hr = f >> op.m_sParameters;                     if (FAILED(hr)) return hr;
    hr = f >> op.m_sWorkingDirectory;               if (FAILED(hr)) return hr;
    hr = f >> op.m_sComment;                        if (FAILED(hr)) return hr;
    hr = f >> (int&)op.m_dwFlags;                   if (FAILED(hr)) return hr;
    hr = f >> (int&)op.m_dwPriority;                if (FAILED(hr)) return hr;
    hr = f >> op.m_sAccountName;                    if (FAILED(hr)) return hr;
    hr = f >> op.m_sPassword;                       if (FAILED(hr)) return hr;
    hr = f >> (int&)dwValue;                        if (FAILED(hr)) return hr; op.m_wIdleMinutes = HIWORD(dwValue); op.m_wDeadlineMinutes = LOWORD(dwValue);
    hr = f >> (int&)op.m_dwMaxRuntimeMS;            if (FAILED(hr)) return hr;
    hr = f >> (int&)dwValue;                        if (FAILED(hr)) return hr;
    while (dwValue--) {
        TASK_TRIGGER ttData;
        hr = f >> ttData;
        if (FAILED(hr)) return hr;
        op.m_lTriggers.AddTail(ttData);
    }

    return S_OK;
}


inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpEnableTask &op)
{
    HRESULT hr;

    hr = f << (const CMSITSCAOpSingleStringOperation&)op;
    if (FAILED(hr)) return hr;

    return f << (int)op.m_bEnable;
}


inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpEnableTask &op)
{
    HRESULT hr;
    int iTemp;

    hr = f >> (CMSITSCAOpSingleStringOperation&)op;
    if (FAILED(hr)) return hr;

    hr = f >> iTemp;
    if (FAILED(hr)) return hr;
    op.m_bEnable = iTemp ? TRUE : FALSE;

    return S_OK;
}


inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpList &list)
{
    POSITION pos;
    HRESULT hr;

    hr = f << (int)list.GetCount();
    if (FAILED(hr)) return hr;

    for (pos = list.GetHeadPosition(); pos;) {
        const CMSITSCAOp *pOp = list.GetNext(pos);
        if (dynamic_cast<const CMSITSCAOpEnableRollback*>(pOp))
            hr = list.Save<CMSITSCAOpEnableRollback, CMSITSCAOpList::OPERATION_ENABLE_ROLLBACK>(f, pOp);
        else if (dynamic_cast<const CMSITSCAOpDeleteFile*>(pOp))
            hr = list.Save<CMSITSCAOpDeleteFile, CMSITSCAOpList::OPERATION_DELETE_FILE>(f, pOp);
        else if (dynamic_cast<const CMSITSCAOpMoveFile*>(pOp))
            hr = list.Save<CMSITSCAOpMoveFile, CMSITSCAOpList::OPERATION_MOVE_FILE>(f, pOp);
        else if (dynamic_cast<const CMSITSCAOpCreateTask*>(pOp))
            hr = list.Save<CMSITSCAOpCreateTask, CMSITSCAOpList::OPERATION_CREATE_TASK>(f, pOp);
        else if (dynamic_cast<const CMSITSCAOpDeleteTask*>(pOp))
            hr = list.Save<CMSITSCAOpDeleteTask, CMSITSCAOpList::OPERATION_DELETE_TASK>(f, pOp);
        else if (dynamic_cast<const CMSITSCAOpEnableTask*>(pOp))
            hr = list.Save<CMSITSCAOpEnableTask, CMSITSCAOpList::OPERATION_ENABLE_TASK>(f, pOp);
        else if (dynamic_cast<const CMSITSCAOpCopyTask*>(pOp))
            hr = list.Save<CMSITSCAOpCopyTask, CMSITSCAOpList::OPERATION_COPY_TASK>(f, pOp);
        else if (dynamic_cast<const CMSITSCAOpList*>(pOp))
            hr = list.Save<CMSITSCAOpList, CMSITSCAOpList::OPERATION_SUBLIST>(f, pOp);
        else {
            // Unsupported type of operation.
            assert(0);
            hr = E_UNEXPECTED;
        }

        if (FAILED(hr)) return hr;
    }

    return S_OK;
}


inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpList &list)
{
    HRESULT hr;
    DWORD dwCount;

    hr = f >> (int&)dwCount;
    if (FAILED(hr)) return hr;

    while (dwCount--) {
        int iTemp;

        hr = f >> iTemp;
        if (FAILED(hr)) return hr;

        switch ((CMSITSCAOpList::OPERATION)iTemp) {
        case CMSITSCAOpList::OPERATION_ENABLE_ROLLBACK:
            hr = list.LoadAndAddTail<CMSITSCAOpEnableRollback>(f);
            break;
        case CMSITSCAOpList::OPERATION_DELETE_FILE:
            hr = list.LoadAndAddTail<CMSITSCAOpDeleteFile>(f);
            break;
        case CMSITSCAOpList::OPERATION_MOVE_FILE:
            hr = list.LoadAndAddTail<CMSITSCAOpMoveFile>(f);
            break;
        case CMSITSCAOpList::OPERATION_CREATE_TASK:
            hr = list.LoadAndAddTail<CMSITSCAOpCreateTask>(f);
            break;
        case CMSITSCAOpList::OPERATION_DELETE_TASK:
            hr = list.LoadAndAddTail<CMSITSCAOpDeleteTask>(f);
            break;
        case CMSITSCAOpList::OPERATION_ENABLE_TASK:
            hr = list.LoadAndAddTail<CMSITSCAOpEnableTask>(f);
            break;
        case CMSITSCAOpList::OPERATION_COPY_TASK:
            hr = list.LoadAndAddTail<CMSITSCAOpCopyTask>(f);
            break;
        case CMSITSCAOpList::OPERATION_SUBLIST:
            hr = list.LoadAndAddTail<CMSITSCAOpList>(f);
            break;
        default:
            // Unsupported type of operation.
            assert(0);
            hr = E_UNEXPECTED;
        }

        if (FAILED(hr)) return hr;
    }

    return S_OK;
}


////////////////////////////////////////////////////////////////////////////
// Inline methods
////////////////////////////////////////////////////////////////////////////

template <class T, int ID> inline static HRESULT CMSITSCAOpList::Save(CAtlFile &f, const CMSITSCAOp *p)
{
    assert(p);
    HRESULT hr;
    const T *pp = dynamic_cast<const T*>(p);
    assert(pp);

    hr = f << (int)ID;
    if (FAILED(hr)) return hr;

    return f << *pp;
}


template <class T> inline HRESULT CMSITSCAOpList::LoadAndAddTail(CAtlFile &f)
{
    HRESULT hr;

    // Create element.
    T *p = new T();
    if (!p) return E_OUTOFMEMORY;

    // Load element from file.
    hr = f >> *p;
    if (FAILED(hr)) {
        delete p;
        return hr;
    }

    // Add element.
    AddTail(p);
    return S_OK;
}

#endif // __MSITSCAOP_H__
