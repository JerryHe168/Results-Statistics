#include "stdafx.h"
#include "MainFrame.h"

CAppModule _Module;

int WINAPI _tWinMain(HINSTANCE hInstance, HINSTANCE /*hPrevInstance*/, LPTSTR /*lpstrCmdLine*/, int nCmdShow)
{
    HRESULT hRes = ::CoInitialize(NULL);
    ATLASSERT(SUCCEEDED(hRes));

    AtlInitCommonControls(ICC_COOL_CLASSES | ICC_BAR_CLASSES);

    hRes = _Module.Init(NULL, hInstance);
    ATLASSERT(SUCCEEDED(hRes));

    int nRet = 0;
    {
        CMessageLoop theLoop;
        _Module.AddMessageLoop(&theLoop);

        CMainFrame wndMain;

        if (wndMain.CreateEx() == NULL)
        {
            ATLTRACE(_T("Main window creation failed!\n"));
            return 0;
        }

        wndMain.ShowWindow(nCmdShow);
        wndMain.UpdateWindow();

        nRet = theLoop.Run();

        _Module.RemoveMessageLoop();
    }

    _Module.Term();
    ::CoUninitialize();

    return nRet;
}
