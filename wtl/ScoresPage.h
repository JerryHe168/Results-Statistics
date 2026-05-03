#pragma once

#include "stdafx.h"
#include "resource.h"

class CScoresPage : public CDialogImpl<CScoresPage>
{
public:
    enum { IDD = IDD_SCORES_PAGE };

    CEdit m_editPath;
    CButton m_btnImport;
    CListViewCtrl m_listView;

    BEGIN_MSG_MAP(CScoresPage)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
        MESSAGE_HANDLER(WM_SIZE, OnSize)
        COMMAND_ID_HANDLER(IDC_BTN_IMPORT_SCORES, OnBtnImport)
    END_MSG_MAP()

    CScoresPage();
    virtual ~CScoresPage();

    LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnBtnImport(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

    void LayoutControls();
    void InitializeListView();
};
