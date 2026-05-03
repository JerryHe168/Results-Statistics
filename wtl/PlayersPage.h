#pragma once

#include "stdafx.h"
#include "resource.h"

class CPlayersPage : public CDialogImpl<CPlayersPage>
{
public:
    enum { IDD = IDD_PLAYERS_PAGE };

    CStatic m_staticTitle;

    BEGIN_MSG_MAP(CPlayersPage)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
        MESSAGE_HANDLER(WM_SIZE, OnSize)
    END_MSG_MAP()

    CPlayersPage();
    virtual ~CPlayersPage();

    LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);

    void LayoutControls();
};
