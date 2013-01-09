#ifndef __MSITSCAOP_H__
#define __MSITSCAOP_H__

#include "MSITSCA.h"
#include <atlbase.h>
#include <atlcoll.h>
#include <atlfile.h>
#include <atlstr.h>
#include <msi.h>
#include <mstask.h>
#include <windows.h>

class CMSITSCASession;


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOp
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOp
{
public:
    CMSITSCAOp(int iTicks = 0);

    virtual HRESULT Execute(CMSITSCASession *pSession) = 0;

    friend class CMSITSCAOpList;
    friend inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOp &op);
    friend inline HRESULT operator >>(CAtlFile &f, CMSITSCAOp &op);

protected:
    int m_iTicks;   // Number of ticks on a progress bar required for this action execution
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpTypeSingleString
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpTypeSingleString : public CMSITSCAOp
{
public:
    CMSITSCAOpTypeSingleString(LPCWSTR pszValue = L"", int iTicks = 0);

    friend inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpTypeSingleString &op);
    friend inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpTypeSingleString &op);

protected:
    CStringW m_sValue;
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpDoubleStringOperation
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpTypeSrcDstString : public CMSITSCAOp
{
public:
    CMSITSCAOpTypeSrcDstString(LPCWSTR pszValue1 = L"", LPCWSTR pszValue2 = L"", int iTicks = 0);

    friend inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpTypeSrcDstString &op);
    friend inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpTypeSrcDstString &op);

protected:
    CStringW m_sValue1;
    CStringW m_sValue2;
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpTypeBoolean
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpTypeBoolean : public CMSITSCAOp
{
public:
    CMSITSCAOpTypeBoolean(BOOL bValue = TRUE, int iTicks = 0);

    friend inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpTypeBoolean &op);
    friend inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpTypeBoolean &op);

protected:
    BOOL m_bValue;
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpRollbackEnable
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpRollbackEnable : public CMSITSCAOpTypeBoolean
{
public:
    CMSITSCAOpRollbackEnable(BOOL bEnable = TRUE, int iTicks = 0);
    virtual HRESULT Execute(CMSITSCASession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpFileDelete
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpFileDelete : public CMSITSCAOpTypeSingleString
{
public:
    CMSITSCAOpFileDelete(LPCWSTR pszFileName = L"", int iTicks = 0);
    virtual HRESULT Execute(CMSITSCASession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpFileMove
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpFileMove : public CMSITSCAOpTypeSrcDstString
{
public:
    CMSITSCAOpFileMove(LPCWSTR pszFileSrc = L"", LPCWSTR pszFileDst = L"", int iTicks = 0);
    virtual HRESULT Execute(CMSITSCASession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpTaskCreate
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpTaskCreate : public CMSITSCAOpTypeSingleString
{
public:
    CMSITSCAOpTaskCreate(LPCWSTR pszTaskName = L"", int iTicks = 0);
    virtual ~CMSITSCAOpTaskCreate();
    virtual HRESULT Execute(CMSITSCASession *pSession);

    UINT SetFromRecord(MSIHANDLE hInstall, MSIHANDLE hRecord);
    UINT SetTriggersFromView(MSIHANDLE hView);

    friend inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpTaskCreate &op);
    friend inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpTaskCreate &op);

protected:
    CStringW m_sApplicationName;
    CStringW m_sParameters;
    CStringW m_sWorkingDirectory;
    CStringW m_sAuthor;
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
// CMSITSCAOpTaskDelete
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpTaskDelete : public CMSITSCAOpTypeSingleString
{
public:
    CMSITSCAOpTaskDelete(LPCWSTR pszTaskName = L"", int iTicks = 0);
    virtual HRESULT Execute(CMSITSCASession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpTaskEnable
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpTaskEnable : public CMSITSCAOpTypeSingleString
{
public:
    CMSITSCAOpTaskEnable(LPCWSTR pszTaskName = L"", BOOL bEnable = TRUE, int iTicks = 0);
    virtual HRESULT Execute(CMSITSCASession *pSession);

    friend inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpTaskEnable &op);
    friend inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpTaskEnable &op);

protected:
    BOOL m_bEnable;
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpTaskCopy
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpTaskCopy : public CMSITSCAOpTypeSrcDstString
{
public:
    CMSITSCAOpTaskCopy(LPCWSTR pszTaskSrc = L"", LPCWSTR pszTaskDst = L"", int iTicks = 0);
    virtual HRESULT Execute(CMSITSCASession *pSession);
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpList
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpList : public CMSITSCAOp, public CAtlList<CMSITSCAOp*>
{
public:
    CMSITSCAOpList(int iTicks = 0);

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

    MSIHANDLE m_hInstall;        // Installer handle
    BOOL m_bContinueOnError;     // Continue execution on operation error?
    BOOL m_bRollbackEnabled;     // Is rollback enabled?
    CMSITSCAOpList m_olRollback; // Rollback operation list
    CMSITSCAOpList m_olCommit;   // Commit operation list
};


////////////////////////////////////////////////////////////////////////////
// Inline operators
////////////////////////////////////////////////////////////////////////////

inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOp &op)
{
    return f << op.m_iTicks;
}


inline HRESULT operator >>(CAtlFile &f, CMSITSCAOp &op)
{
    return f >> op.m_iTicks;
}


inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpTypeSingleString &op)
{
    HRESULT hr;

    hr = f << (const CMSITSCAOp &)op;
    if (FAILED(hr)) return hr;

    return f << op.m_sValue;
}


inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpTypeSingleString &op)
{
    HRESULT hr;

    hr = f >> (CMSITSCAOp &)op;
    if (FAILED(hr)) return hr;

    return f >> op.m_sValue;
}


inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpTypeSrcDstString &op)
{
    HRESULT hr;

    hr = f << (const CMSITSCAOp &)op;
    if (FAILED(hr)) return hr;

    hr = f << op.m_sValue1;
    if (FAILED(hr)) return hr;

    return f << op.m_sValue2;
}


inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpTypeSrcDstString &op)
{
    HRESULT hr;

    hr = f >> (CMSITSCAOp &)op;
    if (FAILED(hr)) return hr;

    hr = f >> op.m_sValue1;
    if (FAILED(hr)) return hr;

    return f >> op.m_sValue2;
}


inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpTypeBoolean &op)
{
    HRESULT hr;

    hr = f << (const CMSITSCAOp &)op;
    if (FAILED(hr)) return hr;

    return f << (int)op.m_bValue;
}


inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpTypeBoolean &op)
{
    int iValue;
    HRESULT hr;

    hr = f >> (CMSITSCAOp &)op;
    if (FAILED(hr)) return hr;

    hr = f >> iValue;
    if (FAILED(hr)) return hr;
    op.m_bValue = iValue ? TRUE : FALSE;

    return S_OK;
}


inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpTaskCreate &op)
{
    HRESULT hr;
    POSITION pos;

    hr = f << (const CMSITSCAOpTypeSingleString&)op;              if (FAILED(hr)) return hr;
    hr = f << op.m_sApplicationName;                                   if (FAILED(hr)) return hr;
    hr = f << op.m_sParameters;                                        if (FAILED(hr)) return hr;
    hr = f << op.m_sWorkingDirectory;                                  if (FAILED(hr)) return hr;
    hr = f << op.m_sAuthor;                                            if (FAILED(hr)) return hr;
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


inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpTaskCreate &op)
{
    HRESULT hr;
    DWORD dwValue;

    hr = f >> (CMSITSCAOpTypeSingleString&)op; if (FAILED(hr)) return hr;
    hr = f >> op.m_sApplicationName;                if (FAILED(hr)) return hr;
    hr = f >> op.m_sParameters;                     if (FAILED(hr)) return hr;
    hr = f >> op.m_sWorkingDirectory;               if (FAILED(hr)) return hr;
    hr = f >> op.m_sAuthor;                         if (FAILED(hr)) return hr;
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


inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpTaskEnable &op)
{
    HRESULT hr;

    hr = f << (const CMSITSCAOpTypeSingleString&)op;
    if (FAILED(hr)) return hr;

    return f << (int)op.m_bEnable;
}


inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpTaskEnable &op)
{
    HRESULT hr;
    int iTemp;

    hr = f >> (CMSITSCAOpTypeSingleString&)op;
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

    hr = f << (const CMSITSCAOp &)list;
    if (FAILED(hr)) return hr;

    hr = f << (int)list.GetCount();
    if (FAILED(hr)) return hr;

    for (pos = list.GetHeadPosition(); pos;) {
        const CMSITSCAOp *pOp = list.GetNext(pos);
        if (dynamic_cast<const CMSITSCAOpRollbackEnable*>(pOp))
            hr = list.Save<CMSITSCAOpRollbackEnable, CMSITSCAOpList::OPERATION_ENABLE_ROLLBACK>(f, pOp);
        else if (dynamic_cast<const CMSITSCAOpFileDelete*>(pOp))
            hr = list.Save<CMSITSCAOpFileDelete, CMSITSCAOpList::OPERATION_DELETE_FILE>(f, pOp);
        else if (dynamic_cast<const CMSITSCAOpFileMove*>(pOp))
            hr = list.Save<CMSITSCAOpFileMove, CMSITSCAOpList::OPERATION_MOVE_FILE>(f, pOp);
        else if (dynamic_cast<const CMSITSCAOpTaskCreate*>(pOp))
            hr = list.Save<CMSITSCAOpTaskCreate, CMSITSCAOpList::OPERATION_CREATE_TASK>(f, pOp);
        else if (dynamic_cast<const CMSITSCAOpTaskDelete*>(pOp))
            hr = list.Save<CMSITSCAOpTaskDelete, CMSITSCAOpList::OPERATION_DELETE_TASK>(f, pOp);
        else if (dynamic_cast<const CMSITSCAOpTaskEnable*>(pOp))
            hr = list.Save<CMSITSCAOpTaskEnable, CMSITSCAOpList::OPERATION_ENABLE_TASK>(f, pOp);
        else if (dynamic_cast<const CMSITSCAOpTaskCopy*>(pOp))
            hr = list.Save<CMSITSCAOpTaskCopy, CMSITSCAOpList::OPERATION_COPY_TASK>(f, pOp);
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

    hr = f >> (CMSITSCAOp &)list;
    if (FAILED(hr)) return hr;

    hr = f >> (int&)dwCount;
    if (FAILED(hr)) return hr;

    while (dwCount--) {
        int iTemp;

        hr = f >> iTemp;
        if (FAILED(hr)) return hr;

        switch ((CMSITSCAOpList::OPERATION)iTemp) {
        case CMSITSCAOpList::OPERATION_ENABLE_ROLLBACK:
            hr = list.LoadAndAddTail<CMSITSCAOpRollbackEnable>(f);
            break;
        case CMSITSCAOpList::OPERATION_DELETE_FILE:
            hr = list.LoadAndAddTail<CMSITSCAOpFileDelete>(f);
            break;
        case CMSITSCAOpList::OPERATION_MOVE_FILE:
            hr = list.LoadAndAddTail<CMSITSCAOpFileMove>(f);
            break;
        case CMSITSCAOpList::OPERATION_CREATE_TASK:
            hr = list.LoadAndAddTail<CMSITSCAOpTaskCreate>(f);
            break;
        case CMSITSCAOpList::OPERATION_DELETE_TASK:
            hr = list.LoadAndAddTail<CMSITSCAOpTaskDelete>(f);
            break;
        case CMSITSCAOpList::OPERATION_ENABLE_TASK:
            hr = list.LoadAndAddTail<CMSITSCAOpTaskEnable>(f);
            break;
        case CMSITSCAOpList::OPERATION_COPY_TASK:
            hr = list.LoadAndAddTail<CMSITSCAOpTaskCopy>(f);
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
