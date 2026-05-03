#include "stdafx.h"
#include "StatsPage.h"
#include "PlayersPage.h"
#include "ScoresPage.h"
#include "DataProcessor.h"
#include <windows.h>
#include <commdlg.h>
#include <fstream>
#include <sstream>

#pragma comment(lib, "comdlg32.lib")

CStatsPage::CStatsPage()
{
}

CStatsPage::~CStatsPage()
{
}

LRESULT CStatsPage::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
    m_btnTemplate = GetDlgItem(IDC_BTN_TEMPLATE);
    m_btnStatistics = GetDlgItem(IDC_BTN_STATISTICS);
    m_btnExport = GetDlgItem(IDC_BTN_EXPORT);
    m_listView = GetDlgItem(IDC_LIST_STATS);

    m_btnTemplate.SetWindowText(L"模板");
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

LRESULT CStatsPage::OnBtnTemplate(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    std::wstring filePath;
    if (ShowFileDialogForImport(filePath))
    {
        if (ImportTemplate(filePath))
        {
            UpdateListViewWithTemplate();
            MessageBox(L"模板导入成功！", L"提示", MB_OK | MB_ICONINFORMATION);
        }
        else
        {
            MessageBox(L"模板导入失败！", L"错误", MB_OK | MB_ICONERROR);
        }
    }
    return 0;
}

LRESULT CStatsPage::OnBtnStatistics(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    HWND hWndMain = GetParent();
    if (hWndMain != NULL)
    {
        ::SendMessage(hWndMain, WM_DO_STATISTICS, 0, 0);
    }
    return 0;
}

LRESULT CStatsPage::OnBtnExport(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
    if (m_results.empty())
    {
        MessageBox(L"没有可导出的数据！请先点击\"统计\"按钮。", L"提示", MB_OK | MB_ICONWARNING);
        return 0;
    }

    std::wstring filePath;
    if (ShowFileDialogForExport(filePath))
    {
        ImportFileFormatStats format = DetectFileFormat(filePath);

        bool success = false;
        if (format == ImportFileFormatStats::Excel)
        {
            success = ExportResults(filePath);
        }
        else if (format == ImportFileFormatStats::Csv)
        {
            success = ExportResultsToCsv(filePath);
        }
        else
        {
            MessageBox(L"不支持的文件格式！", L"错误", MB_OK | MB_ICONERROR);
            return 0;
        }

        if (success)
        {
            MessageBox(L"导出成功！", L"提示", MB_OK | MB_ICONINFORMATION);
        }
        else
        {
            MessageBox(L"导出失败！", L"错误", MB_OK | MB_ICONERROR);
        }
    }
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

    int nTemplateX = nStatisticsX - nButtonWidth - nSpacing;
    int nTemplateY = nMargin;

    int nListX = nMargin;
    int nListY = nMargin + nToolbarHeight;
    int nListWidth = (rcClient.right - rcClient.left) - nMargin * 2;
    int nListHeight = (rcClient.bottom - rcClient.top) - nMargin * 2 - nToolbarHeight;

    m_btnTemplate.SetWindowPos(NULL, nTemplateX, nTemplateY, nButtonWidth, nButtonHeight, SWP_NOZORDER);
    m_btnStatistics.SetWindowPos(NULL, nStatisticsX, nStatisticsY, nButtonWidth, nButtonHeight, SWP_NOZORDER);
    m_btnExport.SetWindowPos(NULL, nExportX, nExportY, nButtonWidth, nButtonHeight, SWP_NOZORDER);
    m_listView.SetWindowPos(NULL, nListX, nListY, nListWidth, nListHeight, SWP_NOZORDER);
}

void CStatsPage::InitializeListView()
{
    m_listView.ModifyStyle(LVS_TYPEMASK, LVS_REPORT | LVS_SHOWSELALWAYS);
    m_listView.SetExtendedListViewStyle(LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

    ClearListView();
}

void CStatsPage::ClearListView()
{
    int nColumnCount = m_listView.GetHeader().GetItemCount();
    for (int i = nColumnCount - 1; i >= 0; i--)
    {
        m_listView.DeleteColumn(i);
    }
    m_listView.DeleteAllItems();
}

void CStatsPage::UpdateListViewWithTemplate()
{
    ClearListView();

    if (m_templateHeaders.empty())
    {
        return;
    }

    for (size_t i = 0; i < m_templateHeaders.size(); i++)
    {
        m_listView.InsertColumn((int)i, m_templateHeaders[i].c_str(), LVCFMT_LEFT, 100);
    }
}

void CStatsPage::UpdateListViewWithResults()
{
    ClearListView();

    if (!m_templateHeaders.empty())
    {
        for (size_t i = 0; i < m_templateHeaders.size(); i++)
        {
            m_listView.InsertColumn((int)i, m_templateHeaders[i].c_str(), LVCFMT_LEFT, 100);
        }
    }
    else
    {
        m_listView.InsertColumn(0, L"名次", LVCFMT_LEFT, 80);
        m_listView.InsertColumn(1, L"组别", LVCFMT_LEFT, 80);
        m_listView.InsertColumn(2, L"姓名", LVCFMT_LEFT, 150);
        m_listView.InsertColumn(3, L"成绩", LVCFMT_LEFT, 100);
    }

    for (size_t i = 0; i < m_results.size(); i++)
    {
        int nItem = m_listView.InsertItem((int)i, std::to_wstring(m_results[i].rank).c_str());
        m_listView.SetItemText(nItem, 1, m_results[i].group.c_str());
        m_listView.SetItemText(nItem, 2, m_results[i].names.c_str());
        m_listView.SetItemText(nItem, 3, m_results[i].time.c_str());
    }
}

bool CStatsPage::ShowFileDialogForImport(std::wstring& filePath)
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
        filePath = szFile;
        return true;
    }
    return false;
}

bool CStatsPage::ShowFileDialogForExport(std::wstring& filePath)
{
    OPENFILENAME ofn;
    wchar_t szFile[MAX_PATH] = { 0 };

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = m_hWnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFilter = L"Excel 文件 (*.xlsx)\0*.xlsx\0Excel 97-2003 文件 (*.xls)\0*.xls\0CSV 文件 (*.csv)\0*.csv\0所有文件 (*.*)\0*.*\0";
    ofn.nFilterIndex = 1;
    ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.lpstrInitialDir = NULL;
    ofn.Flags = OFN_OVERWRITEPROMPT | OFN_HIDEREADONLY | OFN_NOCHANGEDIR;

    if (GetSaveFileName(&ofn) == TRUE)
    {
        filePath = szFile;

        std::wstring lowerPath = filePath;
        std::transform(lowerPath.begin(), lowerPath.end(), lowerPath.begin(), ::towlower);

        bool hasExtension = false;
        if (lowerPath.length() >= 5)
        {
            std::wstring ext = lowerPath.substr(lowerPath.length() - 5);
            if (ext == L".xlsx")
            {
                hasExtension = true;
            }
        }
        if (!hasExtension && lowerPath.length() >= 4)
        {
            std::wstring ext = lowerPath.substr(lowerPath.length() - 4);
            if (ext == L".xls" || ext == L".csv")
            {
                hasExtension = true;
            }
        }

        if (!hasExtension)
        {
            switch (ofn.nFilterIndex)
            {
            case 1:
                filePath += L".xlsx";
                break;
            case 2:
                filePath += L".xls";
                break;
            case 3:
                filePath += L".csv";
                break;
            default:
                break;
            }
        }

        return true;
    }
    return false;
}

ImportFileFormatStats CStatsPage::DetectFileFormat(const std::wstring& filePath)
{
    std::wstring lowerPath = filePath;
    std::transform(lowerPath.begin(), lowerPath.end(), lowerPath.begin(), ::towlower);

    if (lowerPath.length() >= 4)
    {
        std::wstring ext = lowerPath.substr(lowerPath.length() - 4);
        if (ext == L".csv")
        {
            return ImportFileFormatStats::Csv;
        }
        if (ext == L".xls")
        {
            return ImportFileFormatStats::Excel;
        }
    }

    if (lowerPath.length() >= 5)
    {
        std::wstring ext = lowerPath.substr(lowerPath.length() - 5);
        if (ext == L".xlsx")
        {
            return ImportFileFormatStats::Excel;
        }
    }

    return ImportFileFormatStats::Unknown;
}

bool CStatsPage::ImportTemplate(const std::wstring& filePath)
{
    m_templateHeaders.clear();

    ImportFileFormatStats format = DetectFileFormat(filePath);

    std::vector<std::wstring> headers;
    std::vector<std::vector<std::wstring>> data;

    switch (format)
    {
    case ImportFileFormatStats::Excel:
        {
            ExcelReader reader;
            if (!reader.ReadRawData(filePath, headers, data))
            {
                return false;
            }
        }
        break;
    case ImportFileFormatStats::Csv:
        {
            CsvReader reader;
            if (!reader.ReadRawData(filePath, headers, data))
            {
                return false;
            }
        }
        break;
    default:
        MessageBox(L"不支持的文件格式！", L"错误", MB_OK | MB_ICONERROR);
        return false;
    }

    m_templateHeaders = headers;
    return true;
}

bool CStatsPage::ExportResults(const std::wstring& filePath)
{
    if (m_results.empty())
    {
        return false;
    }

    DataProcessor processor;
    if (!m_templateHeaders.empty())
    {
        return processor.ExportResults(filePath, m_results, m_templateHeaders);
    }
    else
    {
        return processor.ExportResults(filePath, m_results);
    }
}

bool CStatsPage::ExportResultsToCsv(const std::wstring& filePath)
{
    if (m_results.empty())
    {
        return false;
    }

    DataProcessor processor;
    if (!m_templateHeaders.empty())
    {
        return processor.ExportResultsToCsv(filePath, m_results, m_templateHeaders);
    }
    else
    {
        return processor.ExportResultsToCsv(filePath, m_results);
    }
}
