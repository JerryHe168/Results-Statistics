#pragma execution_character_set("utf-8")

#define NOMINMAX
#include <windows.h>
#include <iostream>
#include <string>
#include <vector>
#include <algorithm>
#include "ExcelReader.h"
#include "CsvReader.h"
#include "DataProcessor.h"
#include "DataTypes.h"

enum class FileFormat {
    Excel,
    Csv,
    Unknown
};

FileFormat DetectFileFormat(const std::wstring& filePath) {
    std::wstring lowerPath = filePath;
    std::transform(lowerPath.begin(), lowerPath.end(), lowerPath.begin(), ::towlower);

    if (lowerPath.length() >= 4) {
        std::wstring ext = lowerPath.substr(lowerPath.length() - 4);
        if (ext == L".csv") {
            return FileFormat::Csv;
        }
        if (ext == L".xls") {
            return FileFormat::Excel;
        }
    }

    if (lowerPath.length() >= 5) {
        std::wstring ext = lowerPath.substr(lowerPath.length() - 5);
        if (ext == L".xlsx") {
            return FileFormat::Excel;
        }
    }

    return FileFormat::Unknown;
}

std::wstring GetFileExtension(const std::wstring& filePath) {
    size_t dotPos = filePath.find_last_of(L'.');
    if (dotPos != std::wstring::npos) {
        return filePath.substr(dotPos);
    }
    return L"";
}

int wmain(int argc, wchar_t* argv[]) {
    std::wcout << L"========================================" << std::endl;
    std::wcout << L"       Results Statistics Program" << std::endl;
    std::wcout << L"========================================" << std::endl;
    std::wcout << std::endl;

    std::wcout << L"Supported formats: .xls, .xlsx, .csv" << std::endl;
    std::wcout << std::endl;

    std::wstring registrationFile;
    std::wstring scoreFile;
    std::wstring outputFile;

    if (argc >= 4) {
        registrationFile = argv[1];
        scoreFile = argv[2];
        outputFile = argv[3];
    }
    else {
        std::wcout << L"Please enter registration info file path: ";
        std::getline(std::wcin, registrationFile);

        std::wcout << L"Please enter score list file path: ";
        std::getline(std::wcin, scoreFile);

        std::wcout << L"Please enter output result file path: ";
        std::getline(std::wcin, outputFile);
    }

    FileFormat regFormat = DetectFileFormat(registrationFile);
    FileFormat scoreFormat = DetectFileFormat(scoreFile);
    FileFormat outputFormat = DetectFileFormat(outputFile);

    if (regFormat == FileFormat::Unknown) {
        std::wcerr << L"Error: Unsupported registration file format" << std::endl;
        std::wcerr << L"       Supported formats: .xls, .xlsx, .csv" << std::endl;
        return 1;
    }

    if (scoreFormat == FileFormat::Unknown) {
        std::wcerr << L"Error: Unsupported score file format" << std::endl;
        std::wcerr << L"       Supported formats: .xls, .xlsx, .csv" << std::endl;
        return 1;
    }

    if (outputFormat == FileFormat::Unknown) {
        std::wcerr << L"Error: Unsupported output file format" << std::endl;
        std::wcerr << L"       Supported formats: .xls, .xlsx, .csv" << std::endl;
        return 1;
    }

    std::wcout << std::endl;
    std::wcout << L"Processing..." << std::endl;
    std::wcout << std::endl;

    ExcelReader excelReader;
    CsvReader csvReader;
    DataProcessor dataProcessor;

    std::vector<Participant> participants;
    std::vector<ScoreEntry> scoreEntries;
    std::vector<ResultEntry> results;

    std::wcout << L"1. Reading registration info..." << std::endl;
    bool regSuccess = false;
    if (regFormat == FileFormat::Excel) {
        regSuccess = excelReader.ReadRegistrationInfo(registrationFile, participants);
    }
    else {
        regSuccess = csvReader.ReadRegistrationInfo(registrationFile, participants);
    }

    if (!regSuccess) {
        std::wcerr << L"Error: Failed to read registration info file" << std::endl;
        return 1;
    }
    std::wcout << L"   Successfully read " << participants.size() << L" registration entries" << std::endl;

    std::wcout << L"2. Reading score list..." << std::endl;
    bool scoreSuccess = false;
    if (scoreFormat == FileFormat::Excel) {
        scoreSuccess = excelReader.ReadScoreList(scoreFile, scoreEntries);
    }
    else {
        scoreSuccess = csvReader.ReadScoreList(scoreFile, scoreEntries);
    }

    if (!scoreSuccess) {
        std::wcerr << L"Error: Failed to read score list file" << std::endl;
        return 1;
    }
    std::wcout << L"   Successfully read " << scoreEntries.size() << L" score entries" << std::endl;

    std::wcout << L"3. Processing data matching..." << std::endl;
    if (!dataProcessor.ProcessData(participants, scoreEntries, results)) {
        std::wcerr << L"Error: Data processing failed" << std::endl;
        return 1;
    }
    std::wcout << L"   Successfully processed " << results.size() << L" result entries" << std::endl;

    std::wcout << L"4. Exporting results..." << std::endl;
    bool exportSuccess = false;
    if (outputFormat == FileFormat::Excel) {
        exportSuccess = dataProcessor.ExportResults(outputFile, results);
    }
    else {
        exportSuccess = dataProcessor.ExportResultsToCsv(outputFile, results);
    }

    if (!exportSuccess) {
        std::wcerr << L"Error: Failed to export result file" << std::endl;
        return 1;
    }
    std::wcout << L"   Successfully exported to: " << outputFile << std::endl;

    std::wcout << std::endl;
    std::wcout << L"========================================" << std::endl;
    std::wcout << L"       Processing Complete!" << std::endl;
    std::wcout << L"========================================" << std::endl;
    std::wcout << std::endl;

    std::wcout << L"Processing Summary:" << std::endl;
    std::wcout << L"  - Registration Info: " << participants.size() << L" entries" << std::endl;
    std::wcout << L"  - Score Records: " << scoreEntries.size() << L" entries" << std::endl;
    std::wcout << L"  - Output Results: " << results.size() << L" entries" << std::endl;
    std::wcout << std::endl;

    if (results.size() > 0) {
        std::wcout << L"First 5 results preview:" << std::endl;
        std::wcout << L"----------------------------------------" << std::endl;
        std::wcout << L"Rank\tGroup\tNames\t\tScore" << std::endl;
        std::wcout << L"----------------------------------------" << std::endl;
        size_t previewCount = std::min(results.size(), (size_t)5);
        for (size_t i = 0; i < previewCount; i++) {
            std::wcout << results[i].rank << L"\t"
                       << results[i].group << L"\t"
                       << results[i].names << L"\t"
                       << results[i].time << std::endl;
        }
        std::wcout << L"----------------------------------------" << std::endl;
    }

    std::wcout << std::endl;
    std::wcout << L"Press any key to exit..." << std::endl;
    std::wcin.get();

    return 0;
}
