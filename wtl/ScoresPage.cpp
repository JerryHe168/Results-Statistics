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
    m_editPath = GetDlgItem(IDC_EDIT_PATH_SCORES);
    m_btnImport = GetDlgItem(IDC_BTN_IMPORT_SCORES);
    m_listView = GetDlgItem(IDC_LIST_SCORES);

    m_btnImport.SetWindowText(L"导入");

    InitializeListView();
    LayoutControls();

    return TRUE;
}

LRESULT CScoresPage::OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    LayoutControls();
    return 0;
}

LRESULT CScoresPage::OnBtnImport(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    MessageBox(L"成绩导入功能", L"提示", MB_OK | MB_ICONINFORMATION);
    return 0;
}

void CScoresPage::LayoutControls()
{
    RECT rcClient;
    GetClientRect(&rcClient);

    int nMargin = MARGIN;
    int nSpacing = CONTROL_SPACING;
    int nButtonWidth = IMPORT_BUTTON_WIDTH;
    int nButtonHeight = IMPORT_BUTTON_HEIGHT;
    int nToolbarHeight = TOOLBAR_HEIGHT;

    int nEditWidth = (rcClient.right - rcClient.left) - nMargin * 2 - nButtonWidth - nSpacing;
    int nEditX = nMargin;
    int nEditY = nMargin;

    int nButtonX = nEditX + nEditWidth + nSpacing;
    int nButtonY = nMargin;

    int nListX = nMargin;
    int nListY = nMargin + nToolbarHeight;
    int nListWidth = (rcClient.right - rcClient.left) - nMargin * 2;
    int nListHeight = (rcClient.bottom - rcClient.top) - nMargin * 2 - nToolbarHeight;

    m_editPath.SetWindowPos(NULL, nEditX, nEditY, nEditWidth, nButtonHeight, SWP_NOZORDER);
    m_btnImport.SetWindowPos(NULL, nButtonX, nButtonY, nButtonWidth, nButtonHeight, SWP_NOZORDER);
    m_listView.SetWindowPos(NULL, nListX, nListY, nListWidth, nListHeight, SWP_NOZORDER);
}

void CScoresPage::InitializeListView()
{
    m_listView.ModifyStyle(LVS_TYPEMASK, LVS_REPORT | LVS_SHOWSELALWAYS);
    m_listView.SetExtendedListViewStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

    m_listView.InsertColumn(0, L"选手ID", LVCFMT_LEFT, 100);
    m_listView.InsertColumn(1, L"选手姓名", LVCFMT_LEFT, 150);
    m_listView.InsertColumn(2, L"成绩", LVCFMT_LEFT, 100);
    m_listView.InsertColumn(3, L"日期", LVCFMT_LEFT, 150);
}
