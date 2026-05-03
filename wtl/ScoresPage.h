#pragma once

#include "stdafx.h"
#include "resource.h"

class CScoresPage : public CDialogImpl<CScoresPage>
{
public:
    enum { IDD = IDD_SCORES_PAGE };

    CStatic m_staticTitle;

    BEGIN_MSG_MAP(CScoresPage)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
        MESSAGE_HANDLER(WM_SIZE, OnSize)
    END_MSG_MAP()

    CScoresPage();
    virtual ~CScoresPage();

    LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);

    void LayoutControls();
};
