#include "stdafx.h"
#include "MainFrame.h"

CMainFrame::CMainFrame() : m_hCurrentPage(NULL)
{
}

CMainFrame::~CMainFrame()
{
}

BOOL CMainFrame::PreTranslateMessage(MSG* pMsg)
{
    if (CFrameWindowImpl<CMainFrame>::PreTranslateMessage(pMsg))
        return TRUE;

    return FALSE;
}

BOOL CMainFrame::OnIdle()
{
    UIUpdateToolBar();
    return FALSE;
}

LRESULT CMainFrame::OnCreate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    m_hWndClient = NULL;

    m_navBar.Create(m_hWnd, IDD_NAVIGATION_BAR);
    m_navBar.ShowWindow(SW_SHOW);

    m_playersPage.Create(m_hWnd);
    m_scoresPage.Create(m_hWnd);
    m_statsPage.Create(m_hWnd);

    SwitchPage(m_playersPage.m_hWnd);

    UpdateLayout();
    PostMessage(WM_SIZE);

    return 0;
}

LRESULT CMainFrame::OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    RECT rcClient;
    GetClientRect(&rcClient);

    int nNavBarWidth = 200;
    int nNavBarX = rcClient.right - nNavBarWidth;

    m_navBar.SetWindowPos(NULL, nNavBarX, 0, nNavBarWidth, rcClient.bottom, SWP_NOZORDER);

    LayoutPages();

    return 0;
}

LRESULT CMainFrame::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
    PostQuitMessage(0);
    bHandled = FALSE;
    return 1;
}

LRESULT CMainFrame::OnBtnPlayers(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    SwitchPage(m_playersPage.m_hWnd);
    return 0;
}

LRESULT CMainFrame::OnBtnScores(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    SwitchPage(m_scoresPage.m_hWnd);
    return 0;
}

LRESULT CMainFrame::OnBtnStats(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    SwitchPage(m_statsPage.m_hWnd);
    return 0;
}

void CMainFrame::SwitchPage(HWND hPage)
{
    if (m_hCurrentPage == hPage)
        return;

    if (m_hCurrentPage != NULL)
    {
        ::ShowWindow(m_hCurrentPage, SW_HIDE);
    }

    m_hCurrentPage = hPage;

    if (m_hCurrentPage != NULL)
    {
        ::ShowWindow(m_hCurrentPage, SW_SHOW);
        LayoutPages();
    }
}

void CMainFrame::LayoutPages()
{
    if (m_hCurrentPage == NULL)
        return;

    RECT rcClient;
    GetClientRect(&rcClient);

    int nNavBarWidth = 200;
    int nPageWidth = rcClient.right - nNavBarWidth;

    ::SetWindowPos(m_hCurrentPage, NULL, 0, 0, nPageWidth, rcClient.bottom, SWP_NOZORDER);
}
