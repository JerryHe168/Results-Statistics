#include "stdafx.h"
#include "AsyncOperation.h"
#include "ExcelReader.h"
#include "CsvReader.h"
#include "DataProcessor.h"
#include <windows.h>
#include <process.h>
#include <objbase.h>

#pragma comment(lib, "ole32.lib")

AsyncOperation::AsyncOperation(int operationType, HWND hNotifyWnd)
    : m_operationType(operationType)
    , m_hNotifyWnd(hNotifyWnd)
    , m_bCancelled(false)
    , m_bRunning(false)
    , m_hThread(NULL)
{
}

AsyncOperation::~AsyncOperation()
{
    if (m_hThread != NULL)
    {
        Cancel();
        WaitForSingleObject(m_hThread, INFINITE);
        CloseHandle(m_hThread);
        m_hThread = NULL;
    }
}

unsigned int __stdcall AsyncOperation::ThreadProc(void* pParam)
{
    AsyncOperation* pThis = static_cast<AsyncOperation*>(pParam);
    if (pThis == NULL)
    {
        return 0;
    }

    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr))
    {
        pThis->NotifyError(L"Failed to initialize COM");
        return 1;
    }

    pThis->m_bRunning = true;
    pThis->Run();
    pThis->m_bRunning = false;

    CoUninitialize();
    return 0;
}

void AsyncOperation::Start()
{
    if (m_hThread != NULL)
    {
        return;
    }

    m_bCancelled = false;
    m_bRunning = false;

    unsigned int threadId;
    m_hThread = (HANDLE)_beginthreadex(NULL, 0, ThreadProc, this, 0, &threadId);

    if (m_hThread == NULL)
    {
        NotifyError(L"Failed to create thread");
    }
}

void AsyncOperation::Cancel()
{
    m_bCancelled = true;
}

void AsyncOperation::SetProgress(int current, int total, const std::wstring& message)
{
    if (m_hNotifyWnd == NULL)
    {
        return;
    }

    int progress = 0;
    if (total > 0)
    {
        progress = (current * 100) / total;
    }

    std::wstring* pMessage = new std::wstring(message);
    ::PostMessage(m_hNotifyWnd, WM_ASYNC_PROGRESS, 
                   (WPARAM)m_operationType, 
                   (LPARAM)MAKELONG(progress, (INT_PTR)pMessage));
}

void AsyncOperation::NotifyComplete()
{
    if (m_hNotifyWnd != NULL)
    {
        ::PostMessage(m_hNotifyWnd, WM_ASYNC_COMPLETE, (WPARAM)m_operationType, 0);
    }
}

void AsyncOperation::NotifyError(const std::wstring& errorMessage)
{
    if (m_hNotifyWnd != NULL)
    {
        std::wstring* pError = new std::wstring(errorMessage);
        ::PostMessage(m_hNotifyWnd, WM_ASYNC_ERROR, (WPARAM)m_operationType, (LPARAM)pError);
    }
}

void AsyncOperation::NotifyCancelled()
{
    if (m_hNotifyWnd != NULL)
    {
        ::PostMessage(m_hNotifyWnd, WM_ASYNC_CANCELLED, (WPARAM)m_operationType, 0);
    }
}

AsyncImportPlayers::AsyncImportPlayers(HWND hNotifyWnd, const std::wstring& filePath)
    : AsyncOperation(ASYNC_OP_IMPORT_PLAYERS, hNotifyWnd)
    , m_filePath(filePath)
{
}

AsyncImportPlayers::~AsyncImportPlayers()
{
}

void AsyncImportPlayers::Run()
{
    if (IsCancelled())
    {
        NotifyCancelled();
        return;
    }

    SetProgress(0, 100, L"正在导入选手数据...");

    DataProcessor processor;
    FileFormat format = processor.DetectFileFormat(m_filePath);

    bool success = false;

    if (format == FileFormat::Excel)
    {
        SetProgress(10, 100, L"正在读取 Excel 文件...");
        if (IsCancelled())
        {
            NotifyCancelled();
            return;
        }

        ExcelReader reader;
        success = reader.ReadRawData(m_filePath, m_headers, m_data);
    }
    else if (format == FileFormat::Csv)
    {
        SetProgress(10, 100, L"正在读取 CSV 文件...");
        if (IsCancelled())
        {
            NotifyCancelled();
            return;
        }

        CsvReader reader;
        success = reader.ReadRawData(m_filePath, m_headers, m_data);
    }
    else
    {
        NotifyError(L"不支持的文件格式！");
        return;
    }

    if (IsCancelled())
    {
        NotifyCancelled();
        return;
    }

    SetProgress(100, 100, L"导入完成");

    if (success)
    {
        NotifyComplete();
    }
    else
    {
        NotifyError(L"导入文件失败！");
    }
}

AsyncImportScores::AsyncImportScores(HWND hNotifyWnd, const std::wstring& filePath)
    : AsyncOperation(ASYNC_OP_IMPORT_SCORES, hNotifyWnd)
    , m_filePath(filePath)
{
}

AsyncImportScores::~AsyncImportScores()
{
}

void AsyncImportScores::Run()
{
    if (IsCancelled())
    {
        NotifyCancelled();
        return;
    }

    SetProgress(0, 100, L"正在导入成绩数据...");

    DataProcessor processor;
    FileFormat format = processor.DetectFileFormat(m_filePath);

    bool success = false;

    if (format == FileFormat::Excel)
    {
        SetProgress(10, 100, L"正在读取 Excel 文件...");
        if (IsCancelled())
        {
            NotifyCancelled();
            return;
        }

        ExcelReader reader;
        success = reader.ReadRawData(m_filePath, m_headers, m_data);
    }
    else if (format == FileFormat::Csv)
    {
        SetProgress(10, 100, L"正在读取 CSV 文件...");
        if (IsCancelled())
        {
            NotifyCancelled();
            return;
        }

        CsvReader reader;
        success = reader.ReadRawData(m_filePath, m_headers, m_data);
    }
    else
    {
        NotifyError(L"不支持的文件格式！");
        return;
    }

    if (IsCancelled())
    {
        NotifyCancelled();
        return;
    }

    SetProgress(100, 100, L"导入完成");

    if (success)
    {
        NotifyComplete();
    }
    else
    {
        NotifyError(L"导入文件失败！");
    }
}

AsyncImportTemplate::AsyncImportTemplate(HWND hNotifyWnd, const std::wstring& filePath)
    : AsyncOperation(ASYNC_OP_IMPORT_TEMPLATE, hNotifyWnd)
    , m_filePath(filePath)
{
}

AsyncImportTemplate::~AsyncImportTemplate()
{
}

void AsyncImportTemplate::Run()
{
    if (IsCancelled())
    {
        NotifyCancelled();
        return;
    }

    SetProgress(0, 100, L"正在导入模板...");

    DataProcessor processor;
    FileFormat format = processor.DetectFileFormat(m_filePath);

    std::vector<std::wstring> headers;
    std::vector<std::vector<std::wstring>> data;
    bool success = false;

    if (format == FileFormat::Excel)
    {
        SetProgress(10, 100, L"正在读取 Excel 文件...");
        if (IsCancelled())
        {
            NotifyCancelled();
            return;
        }

        ExcelReader reader;
        success = reader.ReadRawData(m_filePath, headers, data);
    }
    else if (format == FileFormat::Csv)
    {
        SetProgress(10, 100, L"正在读取 CSV 文件...");
        if (IsCancelled())
        {
            NotifyCancelled();
            return;
        }

        CsvReader reader;
        success = reader.ReadRawData(m_filePath, headers, data);
    }
    else
    {
        NotifyError(L"不支持的文件格式！");
        return;
    }

    if (IsCancelled())
    {
        NotifyCancelled();
        return;
    }

    m_headers = headers;
    SetProgress(100, 100, L"导入完成");

    if (success)
    {
        NotifyComplete();
    }
    else
    {
        NotifyError(L"导入文件失败！");
    }
}

AsyncStatistics::AsyncStatistics(HWND hNotifyWnd, 
                                 const std::vector<Participant>& participants,
                                 const std::vector<ScoreEntry>& scoreEntries)
    : AsyncOperation(ASYNC_OP_STATISTICS, hNotifyWnd)
    , m_participants(participants)
    , m_scoreEntries(scoreEntries)
{
}

AsyncStatistics::~AsyncStatistics()
{
}

void AsyncStatistics::Run()
{
    if (IsCancelled())
    {
        NotifyCancelled();
        return;
    }

    SetProgress(0, 100, L"正在进行统计...");

    if (m_scoreEntries.empty())
    {
        NotifyError(L"没有成绩数据！");
        return;
    }

    DataProcessor processor;

    SetProgress(50, 100, L"正在匹配数据...");
    if (IsCancelled())
    {
        NotifyCancelled();
        return;
    }

    processor.ProcessData(m_participants, m_scoreEntries, m_results);

    if (IsCancelled())
    {
        NotifyCancelled();
        return;
    }

    SetProgress(100, 100, L"统计完成");
    NotifyComplete();
}

AsyncExport::AsyncExport(HWND hNotifyWnd, 
                          const std::wstring& filePath,
                          const std::vector<ResultEntry>& results,
                          const std::vector<std::wstring>& headers)
    : AsyncOperation(ASYNC_OP_EXPORT, hNotifyWnd)
    , m_filePath(filePath)
    , m_results(results)
    , m_headers(headers)
{
}

AsyncExport::~AsyncExport()
{
}

void AsyncExport::Run()
{
    if (IsCancelled())
    {
        NotifyCancelled();
        return;
    }

    SetProgress(0, 100, L"正在导出数据...");

    DataProcessor processor;
    FileFormat format = processor.DetectFileFormat(m_filePath);

    bool success = false;

    if (format == FileFormat::Excel)
    {
        SetProgress(30, 100, L"正在创建 Excel 文件...");
        if (IsCancelled())
        {
            NotifyCancelled();
            return;
        }

        if (!m_headers.empty())
        {
            success = processor.ExportResults(m_filePath, m_results, m_headers);
        }
        else
        {
            success = processor.ExportResults(m_filePath, m_results);
        }
    }
    else if (format == FileFormat::Csv)
    {
        SetProgress(30, 100, L"正在创建 CSV 文件...");
        if (IsCancelled())
        {
            NotifyCancelled();
            return;
        }

        if (!m_headers.empty())
        {
            success = processor.ExportResultsToCsv(m_filePath, m_results, m_headers);
        }
        else
        {
            success = processor.ExportResultsToCsv(m_filePath, m_results);
        }
    }
    else
    {
        NotifyError(L"不支持的文件格式！");
        return;
    }

    if (IsCancelled())
    {
        NotifyCancelled();
        return;
    }

    SetProgress(100, 100, L"导出完成");

    if (success)
    {
        NotifyComplete();
    }
    else
    {
        NotifyError(L"导出文件失败！");
    }
}
