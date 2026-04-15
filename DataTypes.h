#pragma once
#include <string>
#include <vector>
#include <unordered_map>

struct Participant {
    std::wstring maleId;
    std::wstring maleName;
    std::wstring femaleId;
    std::wstring femaleName;
    int groupNumber;
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
