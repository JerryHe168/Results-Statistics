#include "stdafx.h"
#include "ScoresPage.h"

CScoresPage::CScoresPage()
{
}

CScoresPage::~CScoresPage()
{
}

LRESULT CScoresPage::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    m_staticTitle = GetDlgItem(IDC_STATIC_SCORES);
    LayoutControls();
    return TRUE;
}

LRESULT CScoresPage::OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    LayoutControls();
    return 0;
}

void CScoresPage::LayoutControls()
{
    RECT rcClient;
    GetClientRect(&rcClient);

    int nTextWidth = 200;
    int nTextHeight = 30;
    int nX = (rcClient.right - rcClient.left - nTextWidth) / 2;
    int nY = (rcClient.bottom - rcClient.top - nTextHeight) / 2;

    m_staticTitle.SetWindowPos(NULL, nX, nY, nTextWidth, nTextHeight, SWP_NOZORDER);
}
