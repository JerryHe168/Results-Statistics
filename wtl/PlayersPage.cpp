#include "stdafx.h"
#include "PlayersPage.h"
#include <windows.h>
#include <commdlg.h>
#include <regex>

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

ImportFileFormat CPlayersPage::DetectFileFormat(const std::wstring& filePath)
{
    std::wstring lowerPath = filePath;
    std::transform(lowerPath.begin(), lowerPath.end(), lowerPath.begin(), ::towlower);

    if (lowerPath.length() >= 4)
    {
        std::wstring ext = lowerPath.substr(lowerPath.length() - 4);
        if (ext == L".csv")
        {
            return ImportFileFormat::Csv;
        }
        if (ext == L".xls")
        {
            return ImportFileFormat::Excel;
        }
    }

    if (lowerPath.length() >= 5)
    {
        std::wstring ext = lowerPath.substr(lowerPath.length() - 5);
        if (ext == L".xlsx")
        {
            return ImportFileFormat::Excel;
        }
    }

    return ImportFileFormat::Unknown;
}

bool CPlayersPage::ImportFile(const std::wstring& filePath)
{
    m_headers.clear();
    m_data.clear();

    ImportFileFormat format = DetectFileFormat(filePath);

    switch (format)
    {
    case ImportFileFormat::Csv:
        return ImportCsvFile(filePath);
    case ImportFileFormat::Excel:
        MessageBox(L"Excel 文件导入需要 COM 自动化支持。\n请先将 Excel 文件另存为 CSV 格式，然后导入 CSV 文件。", L"提示", MB_OK | MB_ICONINFORMATION);
        return false;
    default:
        MessageBox(L"不支持的文件格式！", L"错误", MB_OK | MB_ICONERROR);
        return false;
    }
}

bool CPlayersPage::ImportCsvFile(const std::wstring& filePath)
{
    FILE* file = NULL;
    errno_t err = _wfopen_s(&file, filePath.c_str(), L"rb");
    if (err != 0 || file == NULL)
    {
        return false;
    }

    fseek(file, 0, SEEK_END);
    long fileSize = ftell(file);
    fseek(file, 0, SEEK_SET);

    std::string content(fileSize, 0);
    fread(&content[0], 1, fileSize, file);
    fclose(file);

    if (content.length() >= 3 &&
        (unsigned char)content[0] == 0xEF &&
        (unsigned char)content[1] == 0xBB &&
        (unsigned char)content[2] == 0xBF)
    {
        content = content.substr(3);
    }

    std::wstring wcontent = StringToWString(content);
    std::wistringstream iss(wcontent);
    std::wstring line;

    bool isFirstLine = true;
    while (std::getline(iss, line))
    {
        if (!line.empty() && line.back() == L'\r')
        {
            line.pop_back();
        }

        if (line.empty())
        {
            continue;
        }

        std::vector<std::wstring> columns = SplitCsvLine(line);

        if (isFirstLine)
        {
            isFirstLine = false;
            m_headers = columns;
        }
        else
        {
            m_data.push_back(columns);
        }
    }

    return true;
}

std::vector<std::wstring> CPlayersPage::SplitCsvLine(const std::wstring& line)
{
    std::vector<std::wstring> result;
    std::wstring current;
    bool inQuotes = false;

    for (size_t i = 0; i < line.length(); i++)
    {
        wchar_t c = line[i];

        if (c == L'"')
        {
            if (inQuotes && i + 1 < line.length() && line[i + 1] == L'"')
            {
                current += L'"';
                i++;
            }
            else
            {
                inQuotes = !inQuotes;
            }
        }
        else if (c == L',' && !inQuotes)
        {
            result.push_back(Trim(current));
            current.clear();
        }
        else
        {
            current += c;
        }
    }

    result.push_back(Trim(current));
    return result;
}

std::wstring CPlayersPage::Trim(const std::wstring& str)
{
    size_t start = str.find_first_not_of(L" \t\"");
    if (start == std::wstring::npos)
    {
        return L"";
    }

    size_t end = str.find_last_not_of(L" \t\"");
    return str.substr(start, end - start + 1);
}

std::wstring CPlayersPage::StringToWString(const std::string& str)
{
    if (str.empty())
    {
        return L"";
    }

    int size = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, NULL, 0);
    if (size <= 0)
    {
        return L"";
    }

    std::wstring result(size - 1, 0);
    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, &result[0], size);

    return result;
}
