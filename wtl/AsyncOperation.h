#pragma once

#include "stdafx.h"
#include "resource.h"
#include "DataTypes.h"
#include <string>
#include <vector>
#include <functional>

class AsyncOperation
{
public:
    AsyncOperation(int operationType, HWND hNotifyWnd);
    virtual ~AsyncOperation();

    int GetOperationType() const { return m_operationType; }
    HWND GetNotifyWnd() const { return m_hNotifyWnd; }

    void Start();
    void Cancel();
    bool IsCancelled() const { return m_bCancelled; }
    bool IsRunning() const { return m_bRunning; }

    void SetProgress(int current, int total, const std::wstring& message = L"");
    void NotifyComplete();
    void NotifyError(const std::wstring& errorMessage);
    void NotifyCancelled();

    virtual void Run() = 0;

protected:
    int m_operationType;
    HWND m_hNotifyWnd;
    volatile bool m_bCancelled;
    volatile bool m_bRunning;
    HANDLE m_hThread;

    static unsigned int __stdcall ThreadProc(void* pParam);
};

class AsyncImportPlayers : public AsyncOperation
{
public:
    AsyncImportPlayers(HWND hNotifyWnd, const std::wstring& filePath);
    ~AsyncImportPlayers();

    const std::vector<std::wstring>& GetHeaders() const { return m_headers; }
    const std::vector<std::vector<std::wstring>>& GetData() const { return m_data; }
    const std::wstring& GetFilePath() const { return m_filePath; }

    virtual void Run();

private:
    std::wstring m_filePath;
    std::vector<std::wstring> m_headers;
    std::vector<std::vector<std::wstring>> m_data;
};

class AsyncImportScores : public AsyncOperation
{
public:
    AsyncImportScores(HWND hNotifyWnd, const std::wstring& filePath);
    ~AsyncImportScores();

    const std::vector<std::wstring>& GetHeaders() const { return m_headers; }
    const std::vector<std::vector<std::wstring>>& GetData() const { return m_data; }
    const std::wstring& GetFilePath() const { return m_filePath; }

    virtual void Run();

private:
    std::wstring m_filePath;
    std::vector<std::wstring> m_headers;
    std::vector<std::vector<std::wstring>> m_data;
};

class AsyncImportTemplate : public AsyncOperation
{
public:
    AsyncImportTemplate(HWND hNotifyWnd, const std::wstring& filePath);
    ~AsyncImportTemplate();

    const std::vector<std::wstring>& GetHeaders() const { return m_headers; }
    const std::wstring& GetFilePath() const { return m_filePath; }

    virtual void Run();

private:
    std::wstring m_filePath;
    std::vector<std::wstring> m_headers;
};

class AsyncStatistics : public AsyncOperation
{
public:
    AsyncStatistics(HWND hNotifyWnd, 
                    const std::vector<Participant>& participants,
                    const std::vector<ScoreEntry>& scoreEntries);
    ~AsyncStatistics();

    const std::vector<ResultEntry>& GetResults() const { return m_results; }

    virtual void Run();

private:
    std::vector<Participant> m_participants;
    std::vector<ScoreEntry> m_scoreEntries;
    std::vector<ResultEntry> m_results;
};

class AsyncExport : public AsyncOperation
{
public:
    AsyncExport(HWND hNotifyWnd, 
                const std::wstring& filePath,
                const std::vector<ResultEntry>& results,
                const std::vector<std::wstring>& headers);
    ~AsyncExport();

    const std::wstring& GetFilePath() const { return m_filePath; }

    virtual void Run();

private:
    std::wstring m_filePath;
    std::vector<ResultEntry> m_results;
    std::vector<std::wstring> m_headers;
};
