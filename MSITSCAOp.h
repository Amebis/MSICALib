#ifndef __MSITSCAOP_H__
#define __MSITSCAOP_H__

#include "MSITSCA.h"
#include <atlcoll.h>
#include <atlfile.h>
#include <atlstr.h>
#include <mstask.h>
#include <windows.h>


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOp
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOp
{
public:
    CMSITSCAOp();

    virtual HRESULT Execute(ITaskScheduler *pTaskScheduler) = 0;
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpSingleTaskOperation
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpSingleTaskOperation : public CMSITSCAOp
{
public:
    CMSITSCAOpSingleTaskOperation(LPCWSTR pszTaskName = L"");

    friend inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpSingleTaskOperation &op);
    friend inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpSingleTaskOperation &op);

protected:
    CStringW m_sTaskName;
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpsrcDstTaskOperation
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpSrcDstTaskOperation : public CMSITSCAOp
{
public:
    CMSITSCAOpSrcDstTaskOperation(LPCWSTR pszSourceTaskName = L"", LPCWSTR pszDestinationTaskName = L"");

    friend inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpSrcDstTaskOperation &op);
    friend inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpSrcDstTaskOperation &op);

protected:
    CStringW m_sSourceTaskName;
    CStringW m_sDestinationTaskName;
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpDeleteTask
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpDeleteTask : public CMSITSCAOpSingleTaskOperation
{
public:
    CMSITSCAOpDeleteTask(LPCWSTR pszTaskName = L"");
    virtual HRESULT Execute(ITaskScheduler *pTaskScheduler);
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpEnableTask
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpEnableTask : public CMSITSCAOpSingleTaskOperation
{
public:
    CMSITSCAOpEnableTask(LPCWSTR pszTaskName = L"", BOOL bEnable = TRUE);
    virtual HRESULT Execute(ITaskScheduler *pTaskScheduler);

    friend inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpEnableTask &op);
    friend inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpEnableTask &op);

protected:
    BOOL m_bEnable;
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpCopyTask
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOpCopyTask : public CMSITSCAOpSrcDstTaskOperation
{
public:
    CMSITSCAOpCopyTask(LPCWSTR pszSourceTaskName = L"", LPCWSTR pszDestinationTaskName = L"");
    virtual HRESULT Execute(ITaskScheduler *pTaskScheduler);
};


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOperationList
////////////////////////////////////////////////////////////////////////////

class CMSITSCAOperationList : public CAtlList<CMSITSCAOp*>
{
public:
    CMSITSCAOperationList();

    HRESULT Save(CAtlFile &f) const;
    HRESULT Load(CAtlFile &f);
    void Free();

    HRESULT Execute(ITaskScheduler *pTaskScheduler, BOOL bContinueOnError);

protected:
    enum OPERATION {
        OPERATION_DELETE_TASK = 1,
        OPERATION_ENABLE_TASK,
        OPERATION_COPY_TASK
    };

protected:
    template <class T, int ID> inline static HRESULT Save(CAtlFile &f, const CMSITSCAOp *p);
    template <class T> inline HRESULT LoadAndAddTail(CAtlFile &f);
};


////////////////////////////////////////////////////////////////////////////
// Inline methods
////////////////////////////////////////////////////////////////////////////

template <class T, int ID> inline static HRESULT CMSITSCAOperationList::Save(CAtlFile &f, const CMSITSCAOp *p)
{
    assert(p);
    HRESULT hr;
    const T *pp = dynamic_cast<const T*>(p);
    assert(pp);

    hr = f << (int)ID;
    if (FAILED(hr)) return hr;

    return f << *pp;
}


template <class T> inline HRESULT CMSITSCAOperationList::LoadAndAddTail(CAtlFile &f)
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


////////////////////////////////////////////////////////////////////////////
// Inline operators
////////////////////////////////////////////////////////////////////////////

inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpSingleTaskOperation &op)
{
    return f << op.m_sTaskName;
}


inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpSingleTaskOperation &op)
{
    return f >> op.m_sTaskName;
}


inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpSrcDstTaskOperation &op)
{
    HRESULT hr;

    hr = f << op.m_sSourceTaskName;
    if (FAILED(hr)) return hr;

    return f << op.m_sDestinationTaskName;
}


inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpSrcDstTaskOperation &op)
{
    HRESULT hr;

    hr = f >> op.m_sSourceTaskName;
    if (FAILED(hr)) return hr;

    return f >> op.m_sDestinationTaskName;
}


inline HRESULT operator <<(CAtlFile &f, const CMSITSCAOpEnableTask &op)
{
    HRESULT hr;

    hr = f << (const CMSITSCAOpSingleTaskOperation&)op;
    if (FAILED(hr)) return hr;

    return f << (int)op.m_bEnable;
}


inline HRESULT operator >>(CAtlFile &f, CMSITSCAOpEnableTask &op)
{
    HRESULT hr;
    int iTemp;

    hr = f >> (CMSITSCAOpSingleTaskOperation&)op;
    if (FAILED(hr)) return hr;

    hr = f >> iTemp;
    if (FAILED(hr)) return hr;
    op.m_bEnable = iTemp ? TRUE : FALSE;

    return S_OK;
}

#endif // __MSITSCAOP_H__
