#include "stdafx.h"
#include "ProgressDialog.h"

CProgressDialog::CProgressDialog()
    : m_pOperation(NULL)
    , m_bCancelled(false)
    , m_bCompleted(false)
    , m_bHasError(false)
{
}

CProgressDialog::~CProgressDialog()
{
}

LRESULT CProgressDialog::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    m_staticProgress = GetDlgItem(IDC_STATIC_PROGRESS_TEXT);
    m_progressBar = GetDlgItem(IDC_PROGRESS_BAR);
    m_btnCancel = GetDlgItem(IDC_BTN_CANCEL_PROGRESS);

    m_progressBar.SetRange(0, 100);
    m_progressBar.SetPos(0);

    if (m_pOperation != NULL)
    {
        SetProgressText(m_pOperation->GetProgressMessage());
        SetProgressValue(m_pOperation->GetProgress());
    }
    else
    {
        SetProgressText(L"正在处理...");
    }

    SetTimer(1, 100, NULL);

    if (m_pOperation != NULL)
    {
        m_pOperation->Start();
    }

    return TRUE;
}

LRESULT CProgressDialog::OnTimer(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    if (m_pOperation == NULL)
    {
        KillTimer(1);
        EndDialog(IDOK);
        return 0;
    }

    if (m_pOperation->IsRunning())
    {
        SetProgressText(m_pOperation->GetProgressMessage());
        SetProgressValue(m_pOperation->GetProgress());
        return 0;
    }

    KillTimer(1);

    if (m_pOperation->HasError())
    {
        m_bHasError = true;
        m_errorMessage = m_pOperation->GetErrorMessage();
        EndDialog(IDABORT);
    }
    else if (m_pOperation->IsCancelled())
    {
        m_bCancelled = true;
        EndDialog(IDCANCEL);
    }
    else if (m_pOperation->IsCompleted())
    {
        m_bCompleted = true;
        EndDialog(IDOK);
    }
    else
    {
        EndDialog(IDOK);
    }

    return 0;
}

LRESULT CProgressDialog::OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    if (m_pOperation != NULL && m_pOperation->IsRunning())
    {
        m_pOperation->Cancel();
        m_bCancelled = true;
        m_btnCancel.EnableWindow(FALSE);
        SetProgressText(L"正在取消...");
    }
    else
    {
        KillTimer(1);
        EndDialog(IDCANCEL);
    }
    return 0;
}

LRESULT CProgressDialog::OnBtnCancel(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    if (m_pOperation != NULL && m_pOperation->IsRunning())
    {
        m_pOperation->Cancel();
        m_bCancelled = true;
    }

    m_btnCancel.EnableWindow(FALSE);
    SetProgressText(L"正在取消...");

    return 0;
}

void CProgressDialog::SetOperation(AsyncOperation* pOperation)
{
    m_pOperation = pOperation;
}

void CProgressDialog::SetProgressText(const std::wstring& text)
{
    m_staticProgress.SetWindowText(text.c_str());
}

void CProgressDialog::SetProgressValue(int value)
{
    if (value < 0) value = 0;
    if (value > 100) value = 100;
    m_progressBar.SetPos(value);
}
