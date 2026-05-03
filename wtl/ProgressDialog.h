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
    bool m_bCompleted;
    bool m_bHasError;
    std::wstring m_errorMessage;

    BEGIN_MSG_MAP(CProgressDialog)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
        MESSAGE_HANDLER(WM_TIMER, OnTimer)
        MESSAGE_HANDLER(WM_CLOSE, OnClose)
        COMMAND_ID_HANDLER(IDC_BTN_CANCEL_PROGRESS, OnBtnCancel)
    END_MSG_MAP()

    CProgressDialog();
    virtual ~CProgressDialog();

    void SetOperation(AsyncOperation* pOperation);
    AsyncOperation* GetOperation() const { return m_pOperation; }
    bool IsCancelled() const { return m_bCancelled; }
    bool IsCompleted() const { return m_bCompleted; }
    bool HasError() const { return m_bHasError; }
    const std::wstring& GetErrorMessage() const { return m_errorMessage; }

    LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnTimer(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnBtnCancel(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

    void SetProgressText(const std::wstring& text);
    void SetProgressValue(int value);
};
