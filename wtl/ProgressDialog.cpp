#include "stdafx.h"
#include "ProgressDialog.h"

CProgressDialog::CProgressDialog()
    : m_pOperation(NULL)
    , m_bCancelled(false)
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

    SetProgressText(L"正在处理...");

    SetTimer(1, 100, NULL);

    if (m_pOperation != NULL)
    {
        m_pOperation->Start();
    }

    return TRUE;
}

LRESULT CProgressDialog::OnTimer(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    if (m_pOperation != NULL && !m_pOperation->IsRunning())
    {
        KillTimer(1);
        EndDialog(IDOK);
    }
    return 0;
}

LRESULT CProgressDialog::OnAsyncProgress(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
{
    int progress = LOWORD(lParam);
    std::wstring* pMessage = (std::wstring*)HIWORD(lParam);

    if (pMessage != NULL)
    {
        SetProgressText(*pMessage);
        delete pMessage;
    }

    SetProgressValue(progress);
    return 0;
}

LRESULT CProgressDialog::OnAsyncComplete(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    KillTimer(1);
    EndDialog(IDOK);
    return 0;
}

LRESULT CProgressDialog::OnAsyncError(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
{
    std::wstring* pError = (std::wstring*)lParam;
    if (pError != NULL)
    {
        MessageBox(pError->c_str(), L"错误", MB_OK | MB_ICONERROR);
        delete pError;
    }

    KillTimer(1);
    EndDialog(IDABORT);
    return 0;
}

LRESULT CProgressDialog::OnAsyncCancelled(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    m_bCancelled = true;
    KillTimer(1);
    EndDialog(IDCANCEL);
    return 0;
}

LRESULT CProgressDialog::OnBtnCancel(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    if (m_pOperation != NULL)
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
