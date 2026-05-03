#pragma once

#include "stdafx.h"
#include "resource.h"
#include "DataTypes.h"
#include "DataProcessor.h"
#include "ExcelReader.h"
#include "ExcelWriter.h"
#include "CsvReader.h"
#include <string>
#include <vector>
#include <algorithm>

enum class ImportFileFormatStats {
    Excel,
    Csv,
    Unknown
};

class CStatsPage : public CDialogImpl<CStatsPage>
{
public:
    enum { IDD = IDD_STATS_PAGE };

    CButton m_btnTemplate;
    CButton m_btnStatistics;
    CButton m_btnExport;
    CListViewCtrl m_listView;

    std::vector<std::wstring> m_templateHeaders;
    std::vector<Participant> m_participants;
    std::vector<ScoreEntry> m_scoreEntries;
    std::vector<ResultEntry> m_results;

    BEGIN_MSG_MAP(CStatsPage)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
        MESSAGE_HANDLER(WM_SIZE, OnSize)
        COMMAND_ID_HANDLER(IDC_BTN_TEMPLATE, OnBtnTemplate)
        COMMAND_ID_HANDLER(IDC_BTN_STATISTICS, OnBtnStatistics)
        COMMAND_ID_HANDLER(IDC_BTN_EXPORT, OnBtnExport)
    END_MSG_MAP()

    CStatsPage();
    virtual ~CStatsPage();

    LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnBtnTemplate(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
    LRESULT OnBtnStatistics(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
    LRESULT OnBtnExport(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

    void LayoutControls();
    void InitializeListView();
    void ClearListView();
    void UpdateListViewWithTemplate();
    void UpdateListViewWithResults();

    bool ShowFileDialogForImport(std::wstring& filePath);
    bool ShowFileDialogForExport(std::wstring& filePath);

    ImportFileFormatStats DetectFileFormat(const std::wstring& filePath);
    bool ImportTemplate(const std::wstring& filePath);

    bool ExportResults(const std::wstring& filePath);
    bool ExportResultsToCsv(const std::wstring& filePath);

    std::string WStringToString(const std::wstring& wstr) const;
    std::string EscapeCsvField(const std::wstring& field) const;
};
