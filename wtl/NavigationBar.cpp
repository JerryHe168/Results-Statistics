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

    m_btnPlayers.SetWindowText(L"选手");
    m_btnScores.SetWindowText(L"成绩");
    m_btnStats.SetWindowText(L"统计");

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

    int nButtonSize = BUTTON_SIZE;
    int nSpacing = BUTTON_SPACING;
    int nStartX = (rcClient.right - rcClient.left - nButtonSize) / 2;
    int nStartY = 20;

    m_btnPlayers.SetWindowPos(NULL, nStartX, nStartY, nButtonSize, nButtonSize, SWP_NOZORDER);
    m_btnScores.SetWindowPos(NULL, nStartX, nStartY + nButtonSize + nSpacing, nButtonSize, nButtonSize, SWP_NOZORDER);
    m_btnStats.SetWindowPos(NULL, nStartX, nStartY + 2 * (nButtonSize + nSpacing), nButtonSize, nButtonSize, SWP_NOZORDER);
}