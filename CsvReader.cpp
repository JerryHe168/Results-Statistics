#pragma execution_character_set("utf-8")

#include "CsvReader.h"
#include <windows.h>
#include <iostream>
#include <fstream>
#include <sstream>
#include <regex>
#include <algorithm>
#include <codecvt>

CsvReader::CsvReader() {
}

CsvReader::~CsvReader() {
}

std::wstring CsvReader::StringToWString(const std::string& str) {
    if (str.empty()) {
        return L"";
    }

    int size = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, NULL, 0);
    if (size <= 0) {
        return L"";
    }

    std::wstring result(size - 1, 0);
    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, &result[0], size);
    return result;
}

std::wstring CsvReader::Trim(const std::wstring& str) {
    size_t start = str.find_first_not_of(L" \t\"");
    if (start == std::wstring::npos) {
        return L"";
    }
    size_t end = str.find_last_not_of(L" \t\"");
    return str.substr(start, end - start + 1);
}

std::vector<std::wstring> CsvReader::SplitCsvLine(const std::wstring& line) {
    std::vector<std::wstring> result;
    std::wstring current;
    bool inQuotes = false;

    for (size_t i = 0; i < line.length(); i++) {
        wchar_t c = line[i];

        if (c == L'"') {
            if (inQuotes && i + 1 < line.length() && line[i + 1] == L'"') {
                current += L'"';
                i++;
            }
            else {
                inQuotes = !inQuotes;
            }
        }
        else if (c == L',' && !inQuotes) {
            result.push_back(Trim(current));
            current.clear();
        }
        else {
            current += c;
        }
    }

    result.push_back(Trim(current));
    return result;
}

int CsvReader::ExtractGroupNumber(const std::wstring& id) {
    std::wregex regex(L"(\\d+)");
    std::wsmatch match;
    if (std::regex_search(id, match, regex)) {
        return std::stoi(match[1].str());
    }
    return -1;
}

bool CsvReader::ReadRegistrationInfo(const std::wstring& filePath, std::vector<Participant>& participants) {
    participants.clear();

    std::string narrowPath;
    int size = WideCharToMultiByte(CP_UTF8, 0, filePath.c_str(), -1, NULL, 0, NULL, NULL);
    if (size > 0) {
        narrowPath.resize(size - 1);
        WideCharToMultiByte(CP_UTF8, 0, filePath.c_str(), -1, &narrowPath[0], size, NULL, NULL);
    }

    std::ifstream file(narrowPath, std::ios::binary);
    if (!file.is_open()) {
        std::wcerr << L"Failed to open CSV file: " << filePath << std::endl;
        return false;
    }

    std::string content((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
    file.close();

    if (content.length() >= 3 && 
        (unsigned char)content[0] == 0xEF && 
        (unsigned char)content[1] == 0xBB && 
        (unsigned char)content[2] == 0xBF) {
        content = content.substr(3);
    }

    std::wstring wcontent = StringToWString(content);
    std::wistringstream iss(wcontent);
    std::wstring line;

    bool isFirstLine = true;
    while (std::getline(iss, line)) {
        if (!line.empty() && line.back() == L'\r') {
            line.pop_back();
        }

        if (line.empty()) {
            continue;
        }

        if (isFirstLine) {
            std::vector<std::wstring> header = SplitCsvLine(line);
            if (header.size() >= 2) {
                std::wstring lowerHeader = header[0];
                std::transform(lowerHeader.begin(), lowerHeader.end(), lowerHeader.begin(), ::towlower);
                if (lowerHeader.find(L"男生") != std::wstring::npos || 
                    lowerHeader.find(L"male") != std::wstring::npos ||
                    lowerHeader.find(L"编号") != std::wstring::npos) {
                    isFirstLine = false;
                    continue;
                }
            }
            isFirstLine = false;
        }

        std::vector<std::wstring> columns = SplitCsvLine(line);
        if (columns.size() < 4) {
            continue;
        }

        Participant participant;
        participant.maleId = columns[0];
        participant.maleName = columns[1];
        participant.femaleId = columns[2];
        participant.femaleName = columns[3];
        participant.maleGroupNumber = ExtractGroupNumber(participant.maleId);
        participant.femaleGroupNumber = ExtractGroupNumber(participant.femaleId);

        if (!participant.maleName.empty() || !participant.femaleName.empty()) {
            participants.push_back(participant);
        }
    }

    return true;
}

bool CsvReader::ReadScoreList(const std::wstring& filePath, std::vector<ScoreEntry>& scoreEntries) {
    scoreEntries.clear();

    std::string narrowPath;
    int size = WideCharToMultiByte(CP_UTF8, 0, filePath.c_str(), -1, NULL, 0, NULL, NULL);
    if (size > 0) {
        narrowPath.resize(size - 1);
        WideCharToMultiByte(CP_UTF8, 0, filePath.c_str(), -1, &narrowPath[0], size, NULL, NULL);
    }

    std::ifstream file(narrowPath, std::ios::binary);
    if (!file.is_open()) {
        std::wcerr << L"Failed to open CSV file: " << filePath << std::endl;
        return false;
    }

    std::string content((std::istreambuf_iterator<char>(file)), std::istreambuf_iterator<char>());
    file.close();

    if (content.length() >= 3 && 
        (unsigned char)content[0] == 0xEF && 
        (unsigned char)content[1] == 0xBB && 
        (unsigned char)content[2] == 0xBF) {
        content = content.substr(3);
    }

    std::wstring wcontent = StringToWString(content);
    std::wistringstream iss(wcontent);
    std::wstring line;

    bool isFirstLine = true;
    while (std::getline(iss, line)) {
        if (!line.empty() && line.back() == L'\r') {
            line.pop_back();
        }

        if (line.empty()) {
            continue;
        }

        if (isFirstLine) {
            std::vector<std::wstring> header = SplitCsvLine(line);
            if (header.size() >= 2) {
                std::wstring lowerHeader = header[0];
                std::transform(lowerHeader.begin(), lowerHeader.end(), lowerHeader.begin(), ::towlower);
                if (lowerHeader.find(L"名次") != std::wstring::npos || 
                    lowerHeader.find(L"rank") != std::wstring::npos ||
                    lowerHeader.find(L"排名") != std::wstring::npos) {
                    isFirstLine = false;
                    continue;
                }
            }
            isFirstLine = false;
        }

        std::vector<std::wstring> columns = SplitCsvLine(line);
        if (columns.size() < 3) {
            continue;
        }

        ScoreEntry entry;

        try {
            entry.rank = std::stoi(columns[0]);
        }
        catch (...) {
            entry.rank = 0;
        }

        entry.group = columns[1];
        entry.time = columns[2];
        entry.groupNumber = ExtractGroupNumber(entry.group);

        if (entry.rank > 0) {
            scoreEntries.push_back(entry);
        }
    }

    return true;
}
