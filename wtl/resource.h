#pragma once

#define IDD_NAVIGATION_BAR            101
#define IDD_PLAYERS_PAGE               102
#define IDD_SCORES_PAGE                103
#define IDD_STATS_PAGE                 104

#define IDC_BTN_PLAYERS                201
#define IDC_BTN_SCORES                 202
#define IDC_BTN_STATS                  203

#define IDC_EDIT_PATH_PLAYERS          301
#define IDC_BTN_IMPORT_PLAYERS         302
#define IDC_LIST_PLAYERS               303

#define IDC_EDIT_PATH_SCORES           401
#define IDC_BTN_IMPORT_SCORES          402
#define IDC_LIST_SCORES                403

#define IDC_BTN_TEMPLATE               500
#define IDC_BTN_STATISTICS             501
#define IDC_BTN_EXPORT                 502
#define IDC_LIST_STATS                 503

#define IDI_APP_ICON                   601

#define IDR_MAINFRAME                  701

#define WM_SWITCH_PAGE                 (WM_USER + 100)
#define WM_DO_STATISTICS               (WM_USER + 101)
#define WM_ASYNC_START                 (WM_USER + 110)
#define WM_ASYNC_PROGRESS              (WM_USER + 111)
#define WM_ASYNC_COMPLETE              (WM_USER + 112)
#define WM_ASYNC_ERROR                 (WM_USER + 113)
#define WM_ASYNC_CANCELLED             (WM_USER + 114)
#define PAGE_PLAYERS                   0
#define PAGE_SCORES                    1
#define PAGE_STATS                     2

#define IDD_PROGRESS_DIALOG            110
#define IDC_STATIC_PROGRESS_TEXT       800
#define IDC_PROGRESS_BAR               801
#define IDC_BTN_CANCEL_PROGRESS        802

#define ASYNC_OP_IMPORT_PLAYERS        1
#define ASYNC_OP_IMPORT_SCORES         2
#define ASYNC_OP_IMPORT_TEMPLATE       3
#define ASYNC_OP_STATISTICS            4
#define ASYNC_OP_EXPORT                5

#define NAVBAR_WIDTH                   80
#define TOOLBAR_HEIGHT                 40
#define BUTTON_SIZE                    60
#define BUTTON_SPACING                 10
#define CONTROL_SPACING                5
#define IMPORT_BUTTON_WIDTH            80
#define IMPORT_BUTTON_HEIGHT           30
#define MARGIN                         10
#define NAVBAR_START_Y                 20
