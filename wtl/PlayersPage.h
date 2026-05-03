#pragma once

#include "stdafx.h"
#include "resource.h"
#include <string>
#include <vector>
#include <algorithm>
#include <fstream>
#include <sstream>

enum class ImportFileFormat {
    Excel,
    Csv,
    Unknown
};

class CPlayersPage : public CDialogImpl<CPlayersPage>
{
public:
    enum { IDD = IDD_PLAYERS_PAGE };

    CEdit m_editPath;
    CButton m_btnImport;
    CListViewCtrl m_listView;
    std::wstring m_strFilePath;
    std::vector<std::wstring> m_headers;
    std::vector<std::vector<std::wstring>> m_data;

    BEGIN_MSG_MAP(CPlayersPage)
        MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)
        MESSAGE_HANDLER(WM_SIZE, OnSize)
        COMMAND_ID_HANDLER(IDC_BTN_IMPORT_PLAYERS, OnBtnImport)
    END_MSG_MAP()

    CPlayersPage();
    virtual ~CPlayersPage();

    LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnBtnImport(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

    void LayoutControls();
    void InitializeListView();
    void UpdateListView();
    void ClearListView();
    bool ShowFileDialog();
    bool ImportFile(const std::wstring& filePath);
    bool ImportCsvFile(const std::wstring& filePath);
    ImportFileFormat DetectFileFormat(const std::wstring& filePath);
    std::vector<std::wstring> SplitCsvLine(const std::wstring& line);
    std::wstring Trim(const std::wstring& str);
    std::wstring StringToWString(const std::string& str);
};
