#include "stdafx.h"
#include "NavigationBar.h"

CNavigationBar::CNavigationBar() : m_nSelectedButton(0)
{
}

CNavigationBar::~CNavigationBar()
{
}

LRESULT CNavigationBar::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    m_btnPlayers = GetDlgItem(IDC_BTN_PLAYERS);
    m_btnScores = GetDlgItem(IDC_BTN_SCORES);
    m_btnStats = GetDlgItem(IDC_BTN_STATS);

    UpdateButtonStates();
    LayoutButtons();

    return TRUE;
}

LRESULT CNavigationBar::OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    LayoutButtons();
    return 0;
}

LRESULT CNavigationBar::OnBtnPlayers(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    m_nSelectedButton = 0;
    UpdateButtonStates();
    ::SendMessage(GetParent(), WM_SWITCH_PAGE, PAGE_PLAYERS, 0);
    return 0;
}

LRESULT CNavigationBar::OnBtnScores(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    m_nSelectedButton = 1;
    UpdateButtonStates();
    ::SendMessage(GetParent(), WM_SWITCH_PAGE, PAGE_SCORES, 0);
    return 0;
}

LRESULT CNavigationBar::OnBtnStats(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    m_nSelectedButton = 2;
    UpdateButtonStates();
    ::SendMessage(GetParent(), WM_SWITCH_PAGE, PAGE_STATS, 0);
    return 0;
}

void CNavigationBar::UpdateButtonStates()
{
    m_btnPlayers.EnableWindow(m_nSelectedButton != 0);
    m_btnScores.EnableWindow(m_nSelectedButton != 1);
    m_btnStats.EnableWindow(m_nSelectedButton != 2);
}

void CNavigationBar::LayoutButtons()
{
    RECT rcClient;
    GetClientRect(&rcClient);

    int nButtonWidth = rcClient.right - rcClient.left - 20;
    int nButtonHeight = 40;
    int nSpacing = 20;
    int nStartY = 20;

    m_btnPlayers.SetWindowPos(NULL, 10, nStartY, nButtonWidth, nButtonHeight, SWP_NOZORDER);
    m_btnScores.SetWindowPos(NULL, 10, nStartY + nButtonHeight + nSpacing, nButtonWidth, nButtonHeight, SWP_NOZORDER);
    m_btnStats.SetWindowPos(NULL, 10, nStartY + 2 * (nButtonHeight + nSpacing), nButtonWidth, nButtonHeight, SWP_NOZORDER);
}
