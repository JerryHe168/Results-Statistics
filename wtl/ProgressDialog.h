#pragma once

#include "stdafx.h"
#include "resource.h"
#include "AsyncOperation.h"
#include <string>

class CProgressDialog : public CDialogImpl<CProgressDialog>
{
public:
    enum { IDD = IDD_PROGRESS_DIALOG };

    CStatic m_staticProgress;
    CProgressBarCtrl m_progressBar;
    CButton m_btnCancel;

    AsyncOperation* m_pOperation;
    bool m_bCancelled;

    BEGIN_MSG_MAP(CProgressDialog)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
        MESSAGE_HANDLER(WM_TIMER, OnTimer)
        MESSAGE_HANDLER(WM_ASYNC_PROGRESS, OnAsyncProgress)
        MESSAGE_HANDLER(WM_ASYNC_COMPLETE, OnAsyncComplete)
        MESSAGE_HANDLER(WM_ASYNC_ERROR, OnAsyncError)
        MESSAGE_HANDLER(WM_ASYNC_CANCELLED, OnAsyncCancelled)
        COMMAND_ID_HANDLER(IDC_BTN_CANCEL_PROGRESS, OnBtnCancel)
    END_MSG_MAP()

    CProgressDialog();
    virtual ~CProgressDialog();

    void SetOperation(AsyncOperation* pOperation);
    AsyncOperation* GetOperation() const { return m_pOperation; }
    bool IsCancelled() const { return m_bCancelled; }

    LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnTimer(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnAsyncProgress(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/);
    LRESULT OnAsyncComplete(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnAsyncError(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/);
    LRESULT OnAsyncCancelled(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnBtnCancel(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

    void SetProgressText(const std::wstring& text);
    void SetProgressValue(int value);
};
