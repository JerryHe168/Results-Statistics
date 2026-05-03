#pragma once

#include "stdafx.h"
#include "resource.h"
#include "NavigationBar.h"
#include "PlayersPage.h"
#include "ScoresPage.h"
#include "StatsPage.h"
#include "AsyncOperation.h"
#include "ProgressDialog.h"

class CMainFrame : 
    public CFrameWindowImpl<CMainFrame>,
    public CUpdateUI<CMainFrame>,
    public CMessageFilter,
    public CIdleHandler
{
public:
    DECLARE_FRAME_WND_CLASS(NULL, IDR_MAINFRAME)

    CNavigationBar m_navBar;
    CPlayersPage m_playersPage;
    CScoresPage m_scoresPage;
    CStatsPage m_statsPage;

    HWND m_hCurrentPage;
    AsyncStatistics* m_pAsyncStatistics;

    BEGIN_MSG_MAP(CMainFrame)
        MESSAGE_HANDLER(WM_CREATE, OnCreate)
        MESSAGE_HANDLER(WM_SIZE, OnSize)
        MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
        MESSAGE_HANDLER(WM_SWITCH_PAGE, OnSwitchPage)
        MESSAGE_HANDLER(WM_DO_STATISTICS, OnDoStatistics)
        MESSAGE_HANDLER(WM_ASYNC_COMPLETE, OnAsyncComplete)
        MESSAGE_HANDLER(WM_ASYNC_ERROR, OnAsyncError)
        MESSAGE_HANDLER(WM_ASYNC_CANCELLED, OnAsyncCancelled)
        CHAIN_MSG_MAP(CUpdateUI<CMainFrame>)
        CHAIN_MSG_MAP(CFrameWindowImpl<CMainFrame>)
    END_MSG_MAP()

    BEGIN_UPDATE_UI_MAP(CMainFrame)
    END_UPDATE_UI_MAP()

    CMainFrame();
    virtual ~CMainFrame();

    virtual BOOL PreTranslateMessage(MSG* pMsg);
    virtual BOOL OnIdle();

    LRESULT OnCreate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnSize(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
    LRESULT OnSwitchPage(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnDoStatistics(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);

    LRESULT OnAsyncComplete(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/);
    LRESULT OnAsyncError(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/);
    LRESULT OnAsyncCancelled(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/);

    void SwitchPage(HWND hPage);
    void LayoutPages();
    void CleanupAsyncStatistics();
};
