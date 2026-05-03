#pragma once

#include "stdafx.h"
#include "resource.h"

class CStatsPage : public CDialogImpl<CStatsPage>
{
public:
    enum { IDD = IDD_STATS_PAGE };

    CStatic m_staticTitle;

    BEGIN_MSG_MAP(CStatsPage)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
        MESSAGE_HANDLER(WM_SIZE, OnSize)
    END_MSG_MAP()

    CStatsPage();
    virtual ~CStatsPage();

    LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);

    void LayoutControls();
};
