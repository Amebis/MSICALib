#include "StdAfx.h"


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOp
////////////////////////////////////////////////////////////////////////////

CMSITSCAOp::CMSITSCAOp()
{
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpSingleTaskOperation
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpSingleTaskOperation::CMSITSCAOpSingleTaskOperation(LPCWSTR pszTaskName) :
    m_sTaskName(pszTaskName),
    CMSITSCAOp()
{
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpsrcDstTaskOperation
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpSrcDstTaskOperation::CMSITSCAOpSrcDstTaskOperation(LPCWSTR pszSourceTaskName, LPCWSTR pszDestinationTaskName) :
    m_sSourceTaskName(pszSourceTaskName),
    m_sDestinationTaskName(pszDestinationTaskName),
    CMSITSCAOp()
{
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpDeleteTask
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpDeleteTask::CMSITSCAOpDeleteTask(LPCWSTR pszTaskName) :
    CMSITSCAOpSingleTaskOperation(pszTaskName)
{
}


HRESULT CMSITSCAOpDeleteTask::Execute(ITaskScheduler *pTaskScheduler)
{
    assert(pTaskScheduler);

    // Delete the task.
    return pTaskScheduler->Delete(m_sTaskName);
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpEnableTask
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpEnableTask::CMSITSCAOpEnableTask(LPCWSTR pszTaskName, BOOL bEnable) :
    m_bEnable(bEnable),
    CMSITSCAOpSingleTaskOperation(pszTaskName)
{
}


HRESULT CMSITSCAOpEnableTask::Execute(ITaskScheduler *pTaskScheduler)
{
    assert(pTaskScheduler);
    HRESULT hr;
    CComPtr<ITask> pTask;

    // Load the task.
    hr = pTaskScheduler->Activate(m_sTaskName, IID_ITask, (IUnknown**)&pTask);
    if (SUCCEEDED(hr)) {
        DWORD dwFlags;

        // Get task's current flags.
        hr = pTask->GetFlags(&dwFlags);
        if (SUCCEEDED(hr)) {
            // Modify flags.
            if (m_bEnable)
                dwFlags &= ~TASK_FLAG_DISABLED;
            else
                dwFlags |=  TASK_FLAG_DISABLED;

            // Set flags.
            hr = pTask->SetFlags(dwFlags);
            if (SUCCEEDED(hr)) {
                // Save the task.
                CComQIPtr<IPersistFile> pTaskFile(pTask);
                hr = pTaskFile->Save(NULL, TRUE);
            }
        }
    }

    return hr;
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOpCopyTask
////////////////////////////////////////////////////////////////////////////

CMSITSCAOpCopyTask::CMSITSCAOpCopyTask(LPCWSTR pszSourceTaskName, LPCWSTR pszDestinationTaskName) :
    CMSITSCAOpSrcDstTaskOperation(pszSourceTaskName, pszDestinationTaskName)
{
}


HRESULT CMSITSCAOpCopyTask::Execute(ITaskScheduler *pTaskScheduler)
{
    assert(pTaskScheduler);
    HRESULT hr;
    CComPtr<ITask> pTask;

    // Load the source task.
    hr = pTaskScheduler->Activate(m_sSourceTaskName, IID_ITask, (IUnknown**)&pTask);
    if (SUCCEEDED(hr)) {
        // Add with different name.
        hr = pTaskScheduler->AddWorkItem(m_sDestinationTaskName, pTask);
        if (SUCCEEDED(hr)) {
            // Save the task.
            CComQIPtr<IPersistFile> pTaskFile(pTask);
            hr = pTaskFile->Save(NULL, TRUE);
        }
    }

    return hr;
}


////////////////////////////////////////////////////////////////////////////
// CMSITSCAOperationList
////////////////////////////////////////////////////////////////////////////

CMSITSCAOperationList::CMSITSCAOperationList() :
    CAtlList<CMSITSCAOp*>(sizeof(CMSITSCAOp*))
{
}


HRESULT CMSITSCAOperationList::Save(CAtlFile &f) const
{
    POSITION pos;
    HRESULT hr;

    for (pos = GetHeadPosition(); pos;) {
        const CMSITSCAOp *pOp = GetNext(pos);
        if (dynamic_cast<const CMSITSCAOpDeleteTask*>(pOp))
            hr = Save<CMSITSCAOpDeleteTask, OPERATION_DELETE_TASK>(f, pOp);
        else if (dynamic_cast<const CMSITSCAOpEnableTask*>(pOp))
            hr = Save<CMSITSCAOpEnableTask, OPERATION_ENABLE_TASK>(f, pOp);
        else if (dynamic_cast<const CMSITSCAOpCopyTask*>(pOp))
            hr = Save<CMSITSCAOpCopyTask, OPERATION_COPY_TASK>(f, pOp);
        else {
            // Unsupported type of operation.
            assert(0);
            hr = E_UNEXPECTED;
        }

        if (FAILED(hr)) return hr;
    }

    return S_OK;
}


HRESULT CMSITSCAOperationList::Load(CAtlFile &f)
{
    int iTemp;
    HRESULT hr;

    for (;;) {
        hr = f >> iTemp;
        if (hr == HRESULT_FROM_WIN32(ERROR_HANDLE_EOF)) {
            hr = S_OK;
            break;
        } else if (FAILED(hr))
            return hr;

        switch ((OPERATION)iTemp) {
        case OPERATION_DELETE_TASK:
            hr = LoadAndAddTail<CMSITSCAOpDeleteTask>(f);
            break;
        case OPERATION_ENABLE_TASK:
            hr = LoadAndAddTail<CMSITSCAOpEnableTask>(f);
            break;
        case OPERATION_COPY_TASK:
            hr = LoadAndAddTail<CMSITSCAOpCopyTask>(f);
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


void CMSITSCAOperationList::Free()
{
    POSITION pos;

    for (pos = GetHeadPosition(); pos;)
        delete GetNext(pos);

    RemoveAll();
}


HRESULT CMSITSCAOperationList::Execute(ITaskScheduler *pTaskScheduler, BOOL bContinueOnError)
{
    POSITION pos;
    HRESULT hr;

    for (pos = GetHeadPosition(); pos;) {
        hr = GetNext(pos)->Execute(pTaskScheduler);
        if (!bContinueOnError && FAILED(hr)) return hr;
    }

    return S_OK;
}
