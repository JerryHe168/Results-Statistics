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
    m_btnStatistics = GetDlgItem(IDC_BTN_STATISTICS);
    m_btnExport = GetDlgItem(IDC_BTN_EXPORT);
    m_listView = GetDlgItem(IDC_LIST_STATS);

    m_btnStatistics.SetWindowText(L"统计");
    m_btnExport.SetWindowText(L"导出");

    InitializeListView();
    LayoutControls();

    return TRUE;
}

LRESULT CStatsPage::OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    LayoutControls();
    return 0;
}

LRESULT CStatsPage::OnBtnStatistics(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    MessageBox(L"统计功能", L"提示", MB_OK | MB_ICONINFORMATION);
    return 0;
}

LRESULT CStatsPage::OnBtnExport(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    MessageBox(L"导出功能", L"提示", MB_OK | MB_ICONINFORMATION);
    return 0;
}

void CStatsPage::LayoutControls()
{
    RECT rcClient;
    GetClientRect(&rcClient);

    int nMargin = MARGIN;
    int nSpacing = BUTTON_SPACING;
    int nButtonWidth = IMPORT_BUTTON_WIDTH;
    int nButtonHeight = IMPORT_BUTTON_HEIGHT;
    int nToolbarHeight = TOOLBAR_HEIGHT;

    int nExportX = (rcClient.right - rcClient.left) - nMargin - nButtonWidth;
    int nExportY = nMargin;

    int nStatisticsX = nExportX - nButtonWidth - nSpacing;
    int nStatisticsY = nMargin;

    int nListX = nMargin;
    int nListY = nMargin + nToolbarHeight;
    int nListWidth = (rcClient.right - rcClient.left) - nMargin * 2;
    int nListHeight = (rcClient.bottom - rcClient.top) - nMargin - nToolbarHeight;

    m_btnStatistics.SetWindowPos(NULL, nStatisticsX, nStatisticsY, nButtonWidth, nButtonHeight, SWP_NOZORDER);
    m_btnExport.SetWindowPos(NULL, nExportX, nExportY, nButtonWidth, nButtonHeight, SWP_NOZORDER);
    m_listView.SetWindowPos(NULL, nListX, nListY, nListWidth, nListHeight, SWP_NOZORDER);
}

void CStatsPage::InitializeListView()
{
    m_listView.ModifyStyle(LVS_TYPEMASK, LVS_REPORT | LVS_SHOWSELALWAYS);
    m_listView.SetExtendedListViewStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

    m_listView.InsertColumn(0, L"选手ID", LVCFMT_LEFT, 100);
    m_listView.InsertColumn(1, L"选手姓名", LVCFMT_LEFT, 150);
    m_listView.InsertColumn(2, L"总成绩", LVCFMT_LEFT, 100);
    m_listView.InsertColumn(3, L"平均成绩", LVCFMT_LEFT, 100);
    m_listView.InsertColumn(4, L"排名", LVCFMT_LEFT, 80);
}
