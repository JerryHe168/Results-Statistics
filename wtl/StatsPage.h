#pragma once

#include "stdafx.h"
#include "resource.h"

class CStatsPage : public CDialogImpl<CStatsPage>
{
public:
    enum { IDD = IDD_STATS_PAGE };

    CButton m_btnStatistics;
    CButton m_btnExport;
    CListViewCtrl m_listView;

    BEGIN_MSG_MAP(CStatsPage)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
        MESSAGE_HANDLER(WM_SIZE, OnSize)
        COMMAND_ID_HANDLER(IDC_BTN_STATISTICS, OnBtnStatistics)
        COMMAND_ID_HANDLER(IDC_BTN_EXPORT, OnBtnExport)
    END_MSG_MAP()

    CStatsPage();
    virtual ~CStatsPage();

    LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnBtnStatistics(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
    LRESULT OnBtnExport(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

    void LayoutControls();
    void InitializeListView();
};
