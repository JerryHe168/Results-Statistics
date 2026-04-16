#pragma execution_character_set("utf-8")

/**
 * @file main.cpp
 * @brief 成绩统计程序 - 主程序入口
 * 
 * 程序功能：
 * 1. 读取报名信息（Excel或CSV格式）
 * 2. 读取成绩清单（Excel或CSV格式）
 * 3. 根据组别匹配男生和女生姓名
 * 4. 导出结果到Excel或CSV文件
 * 
 * 支持的格式：.xls, .xlsx, .csv
 */

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

/**
 * @brief 文件格式枚举
 * 
 * 用于标识检测到的文件格式类型。
 */
enum class FileFormat {
    Excel,      ///< Excel格式（.xls 或 .xlsx）
    Csv,        ///< CSV格式（.csv）
    Unknown     ///< 未知或不支持的格式
};

/**
 * @brief 检测文件格式
 * 
 * 根据文件扩展名检测文件格式。
 * 
 * @param filePath 文件路径
 * @return FileFormat 文件格式枚举值
 */
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

/**
 * @brief 获取文件扩展名
 * 
 * @param filePath 文件路径
 * @return std::wstring 文件扩展名（包含点号）
 */
std::wstring GetFileExtension(const std::wstring& filePath) {
    size_t dotPos = filePath.find_last_of(L'.');
    if (dotPos != std::wstring::npos) {
        return filePath.substr(dotPos);
    }
    return L"";
}

/**
 * @brief 主程序入口
 * 
 * 执行流程：
 * 1. 显示欢迎信息和支持的格式
 * 2. 获取三个文件路径（命令行参数或交互式输入）
 * 3. 检测每个文件的格式
 * 4. 验证格式是否支持
 * 5. 读取报名信息
 * 6. 读取成绩清单
 * 7. 处理数据匹配
 * 8. 导出结果
 * 9. 显示处理结果和预览
 * 10. 等待用户按键退出
 * 
 * @param argc 命令行参数个数
 * @param argv 命令行参数数组
 * @return int 程序退出码（0=成功，1=失败）
 */
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
