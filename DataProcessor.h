#pragma once
#include "DataTypes.h"
#include <vector>
#include <string>

class DataProcessor {
public:
    DataProcessor();
    ~DataProcessor();

    bool ProcessData(const std::vector<Participant>& participants,
                     const std::vector<ScoreEntry>& scoreEntries,
                     std::vector<ResultEntry>& results);

    bool ExportResults(const std::wstring& filePath, const std::vector<ResultEntry>& results);
    bool ExportResultsToCsv(const std::wstring& filePath, const std::vector<ResultEntry>& results);

private:
    std::string WStringToString(const std::wstring& str);
    std::string EscapeCsvField(const std::wstring& field);
};
