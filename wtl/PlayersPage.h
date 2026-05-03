#pragma once

#include "stdafx.h"
#include "resource.h"
#include "DataTypes.h"
#include "DataProcessor.h"
#include "ExcelReader.h"
#include "CsvReader.h"
#include "AsyncOperation.h"
#include "ProgressDialog.h"
#include <string>
#include <vector>
#include <algorithm>

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
    std::vector<Participant> m_participants;

    AsyncImportPlayers* m_pAsyncImport;

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

    void ParseParticipants();
    int ExtractGroupNumber(const std::wstring& id);
    const std::vector<Participant>& GetParticipants() const;
    bool HasData() const;

    void StartAsyncImport(const std::wstring& filePath);
    void CleanupAsyncImport();
};
