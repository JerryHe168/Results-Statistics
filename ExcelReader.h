#pragma once
#include "DataTypes.h"
#include <vector>
#include <string>

class ExcelReader {
public:
    ExcelReader();
    ~ExcelReader();

    bool ReadRegistrationInfo(const std::wstring& filePath, std::vector<Participant>& participants);
    bool ReadScoreList(const std::wstring& filePath, std::vector<ScoreEntry>& scoreEntries);

private:
    int ExtractGroupNumber(const std::wstring& id);
};
