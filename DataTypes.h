#pragma once
#include <string>
#include <vector>
#include <unordered_map>

struct Participant {
    std::wstring maleId;
    std::wstring maleName;
    int maleGroupNumber;
    std::wstring femaleId;
    std::wstring femaleName;
    int femaleGroupNumber;
};

struct ScoreEntry {
    int rank;
    std::wstring group;
    std::wstring time;
    int groupNumber;
};

struct ResultEntry {
    int rank;
    std::wstring group;
    std::wstring names;
    std::wstring time;
};
