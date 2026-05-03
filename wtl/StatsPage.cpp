#include "stdafx.h"
#include "StatsPage.h"

CStatsPage::CStatsPage()
{
}

CStatsPage::~CStatsPage()
{
}

LRESULT CStatsPage::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    m_staticTitle = GetDlgItem(IDC_STATIC_STATS);
    m_staticTitle.SetWindowText(L"统计页面");
    LayoutControls();
    return TRUE;
}

LRESULT CStatsPage::OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    LayoutControls();
    return 0;
}

void CStatsPage::LayoutControls()
{
    RECT rcClient;
    GetClientRect(&rcClient);

    int nTextWidth = 200;
    int nTextHeight = 30;
    int nX = (rcClient.right - rcClient.left - nTextWidth) / 2;
    int nY = (rcClient.bottom - rcClient.top - nTextHeight) / 2;

    m_staticTitle.SetWindowPos(NULL, nX, nY, nTextWidth, nTextHeight, SWP_NOZORDER);
}
