#pragma once
#include "DataTypes.h"
#include <vector>
#include <string>

class CsvReader {
public:
    CsvReader();
    ~CsvReader();

    bool ReadRegistrationInfo(const std::wstring& filePath, std::vector<Participant>& participants);
    bool ReadScoreList(const std::wstring& filePath, std::vector<ScoreEntry>& scoreEntries);

private:
    int ExtractGroupNumber(const std::wstring& id);
    std::vector<std::wstring> SplitCsvLine(const std::wstring& line);
    std::wstring Trim(const std::wstring& str);
    std::wstring StringToWString(const std::string& str);
};
