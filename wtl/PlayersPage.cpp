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
            ParseParticipants();
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

bool CPlayersPage::ImportFile(const std::wstring& filePath)
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

void CPlayersPage::ParseParticipants()
{
    m_participants.clear();

    if (m_data.empty())
    {
        return;
    }

    for (size_t i = 0; i < m_data.size(); i++)
    {
        const auto& row = m_data[i];
        Participant participant;

        if (row.size() > 0)
        {
            participant.maleId = row[0];
            participant.maleGroupNumber = ExtractGroupNumber(row[0]);
        }
        else
        {
            participant.maleId = L"";
            participant.maleGroupNumber = -1;
        }

        if (row.size() > 1)
        {
            participant.maleName = row[1];
        }
        else
        {
            participant.maleName = L"";
        }

        if (row.size() > 2)
        {
            participant.femaleId = row[2];
            participant.femaleGroupNumber = ExtractGroupNumber(row[2]);
        }
        else
        {
            participant.femaleId = L"";
            participant.femaleGroupNumber = -1;
        }

        if (row.size() > 3)
        {
            participant.femaleName = row[3];
        }
        else
        {
            participant.femaleName = L"";
        }

        m_participants.push_back(participant);
    }
}

int CPlayersPage::ExtractGroupNumber(const std::wstring& id)
{
    if (id.empty())
    {
        return -1;
    }

    int groupNumber = 0;
    for (wchar_t c : id)
    {
        if (c >= L'0' && c <= L'9')
        {
            groupNumber = groupNumber * 10 + (c - L'0');
        }
        else
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

const std::vector<Participant>& CPlayersPage::GetParticipants() const
{
    return m_participants;
}

bool CPlayersPage::HasData() const
{
    return !m_participants.empty();
}
