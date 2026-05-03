#include "stdafx.h"
#include "PlayersPage.h"
#include <windows.h>
#include <commdlg.h>
#include <algorithm>

#pragma comment(lib, "comdlg32.lib")

CPlayersPage::CPlayersPage()
{
}

CPlayersPage::~CPlayersPage()
{
}

LRESULT CPlayersPage::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    m_editPath = GetDlgItem(IDC_EDIT_PATH_PLAYERS);
    m_btnImport = GetDlgItem(IDC_BTN_IMPORT_PLAYERS);
    m_listView = GetDlgItem(IDC_LIST_PLAYERS);

    m_btnImport.SetWindowText(L"导入");

    InitializeListView();
    LayoutControls();

    return TRUE;
}

LRESULT CPlayersPage::OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    LayoutControls();
    return 0;
}

LRESULT CPlayersPage::OnBtnImport(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    if (ShowFileDialog())
    {
        if (ImportFile(m_strFilePath))
        {
            m_editPath.SetWindowText(m_strFilePath.c_str());
            UpdateListView();
        }
        else
        {
            MessageBox(L"导入文件失败！", L"错误", MB_OK | MB_ICONERROR);
        }
    }
    return 0;
}

void CPlayersPage::LayoutControls()
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

void CPlayersPage::InitializeListView()
{
    m_listView.ModifyStyle(LVS_TYPEMASK, LVS_REPORT | LVS_SHOWSELALWAYS);
    m_listView.SetExtendedListViewStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

    ClearListView();
}

void CPlayersPage::ClearListView()
{
    int nColumnCount = m_listView.GetHeader().GetItemCount();
    for (int i = nColumnCount - 1; i >= 0; i--)
    {
        m_listView.DeleteColumn(i);
    }
    m_listView.DeleteAllItems();
}

void CPlayersPage::UpdateListView()
{
    ClearListView();

    if (m_headers.empty() && m_data.empty())
    {
        return;
    }

    if (m_headers.empty())
    {
        if (!m_data.empty())
        {
            for (size_t i = 0; i < m_data[0].size(); i++)
            {
                std::wstring strCol = L"列" + std::to_wstring(i + 1);
                m_listView.InsertColumn((int)i, strCol.c_str(), LVCFMT_LEFT, 100);
            }
        }
    }
    else
    {
        for (size_t i = 0; i < m_headers.size(); i++)
        {
            m_listView.InsertColumn((int)i, m_headers[i].c_str(), LVCFMT_LEFT, 100);
        }
    }

    for (size_t i = 0; i < m_data.size(); i++)
    {
        if (m_data[i].size() > 0)
        {
            m_listView.InsertItem((int)i, m_data[i][0].c_str());
            for (size_t j = 1; j < m_data[i].size(); j++)
            {
                m_listView.SetItemText((int)i, (int)j, m_data[i][j].c_str());
            }
        }
    }
}

bool CPlayersPage::ShowFileDialog()
{
    OPENFILENAME ofn;
    wchar_t szFile[MAX_PATH] = { 0 };

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = m_hWnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFilter = L"所有支持的文件\0*.xls;*.xlsx;*.csv\0Excel 文件 (*.xls;*.xlsx)\0*.xls;*.xlsx\0CSV 文件 (*.csv)\0*.csv\0所有文件 (*.*)\0*.*\0";
    ofn.nFilterIndex = 1;
    ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.lpstrInitialDir = NULL;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;

    if (GetOpenFileName(&ofn) == TRUE)
    {
        m_strFilePath = szFile;
        return true;
    }
    return false;
}

ImportFileFormatPlayers CPlayersPage::DetectFileFormat(const std::wstring& filePath)
{
    std::wstring lowerPath = filePath;
    std::transform(lowerPath.begin(), lowerPath.end(), lowerPath.begin(), ::towlower);

    if (lowerPath.length() >= 4)
    {
        std::wstring ext = lowerPath.substr(lowerPath.length() - 4);
        if (ext == L".csv")
        {
            return ImportFileFormatPlayers::Csv;
        }
        if (ext == L".xls")
        {
            return ImportFileFormatPlayers::Excel;
        }
    }

    if (lowerPath.length() >= 5)
    {
        std::wstring ext = lowerPath.substr(lowerPath.length() - 5);
        if (ext == L".xlsx")
        {
            return ImportFileFormatPlayers::Excel;
        }
    }

    return ImportFileFormatPlayers::Unknown;
}

bool CPlayersPage::ImportFile(const std::wstring& filePath)
{
    m_headers.clear();
    m_data.clear();

    ImportFileFormatPlayers format = DetectFileFormat(filePath);

    switch (format)
    {
    case ImportFileFormatPlayers::Excel:
        {
            ExcelReader reader;
            return reader.ReadRawData(filePath, m_headers, m_data);
        }
    case ImportFileFormatPlayers::Csv:
        {
            CsvReader reader;
            return reader.ReadRawData(filePath, m_headers, m_data);
        }
    default:
        MessageBox(L"不支持的文件格式！", L"错误", MB_OK | MB_ICONERROR);
        return false;
    }
}
