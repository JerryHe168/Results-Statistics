#include "stdafx.h"
#include "MainFrame.h"
#include "DataProcessor.h"
#include <windows.h>
#include <objbase.h>

#pragma comment(lib, "ole32.lib")

CMainFrame::CMainFrame() : m_hCurrentPage(NULL), m_pAsyncStatistics(NULL)
{
}

CMainFrame::~CMainFrame()
{
    CleanupAsyncStatistics();
}

void CMainFrame::CleanupAsyncStatistics()
{
    if (m_pAsyncStatistics != NULL)
    {
        delete m_pAsyncStatistics;
        m_pAsyncStatistics = NULL;
    }
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
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr))
    {
        MessageBox(L"Failed to initialize COM library!", L"Error", MB_OK | MB_ICONERROR);
    }

    m_hWndClient = NULL;

    SetWindowText(L"成绩统计程序");

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

    int nNavBarWidth = NAVBAR_WIDTH;

    m_navBar.SetWindowPos(NULL, 0, 0, nNavBarWidth, rcClient.bottom, SWP_NOZORDER);

    LayoutPages();

    return 0;
}

LRESULT CMainFrame::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
    CoUninitialize();
    PostQuitMessage(0);
    bHandled = FALSE;
    return 1;
}

LRESULT CMainFrame::OnSwitchPage(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    switch (wParam)
    {
    case PAGE_PLAYERS:
        SwitchPage(m_playersPage.m_hWnd);
        break;
    case PAGE_SCORES:
        SwitchPage(m_scoresPage.m_hWnd);
        break;
    case PAGE_STATS:
        SwitchPage(m_statsPage.m_hWnd);
        break;
    }
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

    int nNavBarWidth = NAVBAR_WIDTH;
    int nPageX = nNavBarWidth;
    int nPageWidth = rcClient.right - nNavBarWidth;

    ::SetWindowPos(m_hCurrentPage, NULL, nPageX, 0, nPageWidth, rcClient.bottom, SWP_NOZORDER);
}

LRESULT CMainFrame::OnDoStatistics(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    if (!m_playersPage.HasData())
    {
        m_statsPage.MessageBox(L"请先导入选手数据！\n\n操作步骤：\n1. 点击\"选手\"按钮，导入选手数据\n2. 点击\"成绩\"按钮，导入成绩数据\n3. 然后再点击\"统计\"按钮进行统计",
                               L"提示", MB_OK | MB_ICONINFORMATION);
        return 0;
    }

    if (!m_scoresPage.HasData())
    {
        m_statsPage.MessageBox(L"请先导入成绩数据！\n\n操作步骤：\n1. 点击\"选手\"按钮，导入选手数据\n2. 点击\"成绩\"按钮，导入成绩数据\n3. 然后再点击\"统计\"按钮进行统计",
                               L"提示", MB_OK | MB_ICONINFORMATION);
        return 0;
    }

    CleanupAsyncStatistics();
    AsyncStatistics* pStatistics = new AsyncStatistics(m_hWnd, 
                                                         m_playersPage.GetParticipants(),
                                                         m_scoresPage.GetScoreEntries());
    m_pAsyncStatistics = pStatistics;

    CProgressDialog dlg;
    dlg.SetOperation(m_pAsyncStatistics);
    INT_PTR nResult = dlg.DoModal();

    if (nResult == IDOK)
    {
        m_statsPage.m_results = pStatistics->GetResults();
        m_statsPage.UpdateListViewWithResults();

        int nCount = (int)m_statsPage.m_results.size();
        std::wstring strMsg = L"统计完成！共 " + std::to_wstring(nCount) + L" 条记录。";
        m_statsPage.MessageBox(strMsg.c_str(), L"提示", MB_OK | MB_ICONINFORMATION);
    }
    else if (nResult == IDABORT)
    {
        std::wstring errorMsg = dlg.GetErrorMessage();
        if (errorMsg.empty())
        {
            errorMsg = L"统计失败！";
        }
        m_statsPage.MessageBox(errorMsg.c_str(), L"错误", MB_OK | MB_ICONERROR);
    }
    else if (nResult == IDCANCEL)
    {
        m_statsPage.MessageBox(L"统计已取消。", L"提示", MB_OK | MB_ICONINFORMATION);
    }

    CleanupAsyncStatistics();

    return 0;
}