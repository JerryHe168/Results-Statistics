#include "stdafx.h"
#include "ScoresPage.h"
#include <windows.h>
#include <commdlg.h>
#include <algorithm>

#pragma comment(lib, "comdlg32.lib")

CScoresPage::CScoresPage()
    : m_pAsyncImport(NULL)
{
}

CScoresPage::~CScoresPage()
{
    CleanupAsyncImport();
}

void CScoresPage::CleanupAsyncImport()
{
    if (m_pAsyncImport != NULL)
    {
        delete m_pAsyncImport;
        m_pAsyncImport = NULL;
    }
}

void CScoresPage::StartAsyncImport(const std::wstring& filePath)
{
    CleanupAsyncImport();
    m_pAsyncImport = new AsyncImportScores(m_hWnd, filePath);

    CProgressDialog dlg;
    dlg.SetOperation(m_pAsyncImport);
    dlg.DoModal();
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
    if (ShowFileDialog())
    {
        StartAsyncImport(m_strFilePath);
    }
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

    ClearListView();
}

void CScoresPage::ClearListView()
{
    int nColumnCount = m_listView.GetHeader().GetItemCount();
    for (int i = nColumnCount - 1; i >= 0; i--)
    {
        m_listView.DeleteColumn(i);
    }
    m_listView.DeleteAllItems();
}

void CScoresPage::UpdateListView()
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

bool CScoresPage::ShowFileDialog()
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

bool CScoresPage::ImportFile(const std::wstring& filePath)
{
    m_headers.clear();
    m_data.clear();

    DataProcessor processor;
    FileFormat format = processor.DetectFileFormat(filePath);

    switch (format)
    {
    case FileFormat::Excel:
        {
            ExcelReader reader;
            return reader.ReadRawData(filePath, m_headers, m_data);
        }
    case FileFormat::Csv:
        {
            CsvReader reader;
            return reader.ReadRawData(filePath, m_headers, m_data);
        }
    default:
        MessageBox(L"不支持的文件格式！", L"错误", MB_OK | MB_ICONERROR);
        return false;
    }
}

void CScoresPage::ParseScoreEntries()
{
    m_scoreEntries.clear();

    if (m_data.empty())
    {
        return;
    }

    for (size_t i = 0; i < m_data.size(); i++)
    {
        const auto& row = m_data[i];
        ScoreEntry scoreEntry;

        if (row.size() > 0)
        {
            scoreEntry.rank = StringToInt(row[0]);
        }
        else
        {
            scoreEntry.rank = 0;
        }

        if (row.size() > 1)
        {
            scoreEntry.group = row[1];
            scoreEntry.groupNumber = ExtractGroupNumberFromGroup(row[1]);
        }
        else
        {
            scoreEntry.group = L"";
            scoreEntry.groupNumber = -1;
        }

        if (row.size() > 2)
        {
            scoreEntry.time = row[2];
        }
        else
        {
            scoreEntry.time = L"";
        }

        m_scoreEntries.push_back(scoreEntry);
    }
}

int CScoresPage::ExtractGroupNumberFromGroup(const std::wstring& group)
{
    if (group.empty())
    {
        return -1;
    }

    int groupNumber = 0;
    for (wchar_t c : group)
    {
        if (c >= L'0' && c <= L'9')
        {
            groupNumber = groupNumber * 10 + (c - L'0');
        }
        else if (groupNumber > 0)
        {
            break;
        }
    }

    if (groupNumber == 0)
    {
        return -1;
    }

    return groupNumber;
}

int CScoresPage::StringToInt(const std::wstring& str)
{
    if (str.empty())
    {
        return 0;
    }

    int result = 0;
    for (wchar_t c : str)
    {
        if (c >= L'0' && c <= L'9')
        {
            result = result * 10 + (c - L'0');
        }
        else
        {
            break;
        }
    }

    return result;
}

const std::vector<ScoreEntry>& CScoresPage::GetScoreEntries() const
{
    return m_scoreEntries;
}

bool CScoresPage::HasData() const
{
    return !m_scoreEntries.empty();
}

LRESULT CScoresPage::OnAsyncComplete(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    if (wParam != ASYNC_OP_IMPORT_SCORES)
    {
        return 0;
    }

    if (m_pAsyncImport == NULL)
    {
        return 0;
    }

    m_headers = m_pAsyncImport->GetHeaders();
    m_data = m_pAsyncImport->GetData();
    m_strFilePath = m_pAsyncImport->GetFilePath();

    ParseScoreEntries();

    m_editPath.SetWindowText(m_strFilePath.c_str());
    UpdateListView();

    CleanupAsyncImport();

    return 0;
}

LRESULT CScoresPage::OnAsyncError(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
{
    if (wParam != ASYNC_OP_IMPORT_SCORES)
    {
        return 0;
    }

    std::wstring* pError = (std::wstring*)lParam;
    if (pError != NULL)
    {
        MessageBox(pError->c_str(), L"错误", MB_OK | MB_ICONERROR);
        delete pError;
    }

    CleanupAsyncImport();

    return 0;
}

LRESULT CScoresPage::OnAsyncCancelled(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    if (wParam != ASYNC_OP_IMPORT_SCORES)
    {
        return 0;
    }

    CleanupAsyncImport();

    return 0;
}
