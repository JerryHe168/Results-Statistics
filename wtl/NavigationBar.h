#pragma once

#include "stdafx.h"
#include "resource.h"

class CNavigationBar : public CDialogImpl<CNavigationBar>
{
public:
    enum { IDD = IDD_NAVIGATION_BAR };

    CButton m_btnPlayers;
    CButton m_btnScores;
    CButton m_btnStats;

    int m_nSelectedButton;

    BEGIN_MSG_MAP(CNavigationBar)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
        MESSAGE_HANDLER(WM_SIZE, OnSize)
        COMMAND_ID_HANDLER(IDC_BTN_PLAYERS, OnBtnPlayers)
        COMMAND_ID_HANDLER(IDC_BTN_SCORES, OnBtnScores)
        COMMAND_ID_HANDLER(IDC_BTN_STATS, OnBtnStats)
    END_MSG_MAP()

    CNavigationBar();
    virtual ~CNavigationBar();

    LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnBtnPlayers(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
    LRESULT OnBtnScores(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
    LRESULT OnBtnStats(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

    void UpdateButtonStates();
    void LayoutButtons();
};
