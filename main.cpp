#pragma execution_character_set("utf-8")

// ============================================================
// 成绩统计程序 - 主程序入口
// ============================================================
// 程序功能：
// 1. 读取报名信息（Excel或CSV格式）
// 2. 读取成绩清单（Excel或CSV格式）
// 3. 根据组别匹配男生和女生姓名
// 4. 导出结果到Excel或CSV文件
// 
// 支持的格式：
// - 输入：.xls, .xlsx, .csv
// - 输出：.xls, .xlsx, .csv
// 
// 使用方式：
//   方式1：命令行参数
//     ResultsStatistics.exe 报名信息.xls 成绩清单.xls 输出结果.xlsx
//   
//   方式2：交互式输入
//     ResultsStatistics.exe
//     然后按提示输入三个文件路径
// ============================================================

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

// ------------------------------------------------------------
// 文件格式枚举
// ------------------------------------------------------------
// 用于标识检测到的文件格式类型
// ------------------------------------------------------------
enum class FileFormat {
    Excel,      // Excel格式（.xls 或 .xlsx）
    Csv,        // CSV格式（.csv）
    Unknown     // 未知或不支持的格式
};

// ------------------------------------------------------------
// 检测文件格式
// ------------------------------------------------------------
// 参数：
//   filePath - 文件路径
// 
// 返回值：
//   FileFormat枚举值
// 
// 工作原理：
//   1. 将文件路径转换为小写
//   2. 检查文件扩展名
//      - .csv  -> CSV格式
//      - .xls  -> Excel格式
//      - .xlsx -> Excel格式
// 
// 注意：
//   扩展名检测顺序很重要，因为.xls是.xlsx的子串
//   需要先检查4字符扩展名（.csv, .xls），再检查5字符扩展名（.xlsx）
// ------------------------------------------------------------
FileFormat DetectFileFormat(const std::wstring& filePath) {
    // 转换为小写以便不区分大小写比较
    std::wstring lowerPath = filePath;
    std::transform(lowerPath.begin(), lowerPath.end(), lowerPath.begin(), ::towlower);

    // 检查4字符扩展名：.csv, .xls
    if (lowerPath.length() >= 4) {
        std::wstring ext = lowerPath.substr(lowerPath.length() - 4);
        if (ext == L".csv") {
            return FileFormat::Csv;
        }
        if (ext == L".xls") {
            return FileFormat::Excel;
        }
    }

    // 检查5字符扩展名：.xlsx
    if (lowerPath.length() >= 5) {
        std::wstring ext = lowerPath.substr(lowerPath.length() - 5);
        if (ext == L".xlsx") {
            return FileFormat::Excel;
        }
    }

    // 不支持的格式
    return FileFormat::Unknown;
}

// ------------------------------------------------------------
// 获取文件扩展名（备用函数，当前未使用）
// ------------------------------------------------------------
std::wstring GetFileExtension(const std::wstring& filePath) {
    size_t dotPos = filePath.find_last_of(L'.');
    if (dotPos != std::wstring::npos) {
        return filePath.substr(dotPos);
    }
    return L"";
}

// ============================================================
// 主程序入口
// ============================================================
// 执行流程：
// 1. 显示欢迎信息和支持的格式
// 2. 获取三个文件路径（命令行参数或交互式输入）
// 3. 检测每个文件的格式
// 4. 验证格式是否支持
// 5. 读取报名信息
// 6. 读取成绩清单
// 7. 处理数据匹配
// 8. 导出结果
// 9. 显示处理结果和预览
// 10. 等待用户按键退出
// ============================================================
int wmain(int argc, wchar_t* argv[]) {
    // --------------------------------------------------------
    // 显示欢迎信息
    // --------------------------------------------------------
    std::wcout << L"========================================" << std::endl;
    std::wcout << L"       Results Statistics Program" << std::endl;
    std::wcout << L"========================================" << std::endl;
    std::wcout << std::endl;

    std::wcout << L"Supported formats: .xls, .xlsx, .csv" << std::endl;
    std::wcout << std::endl;

    // --------------------------------------------------------
    // 获取文件路径
    // --------------------------------------------------------
    std::wstring registrationFile;  // 报名信息文件路径
    std::wstring scoreFile;         // 成绩清单文件路径
    std::wstring outputFile;        // 输出结果文件路径

    if (argc >= 4) {
        // 方式1：从命令行参数获取
        registrationFile = argv[1];
        scoreFile = argv[2];
        outputFile = argv[3];
    }
    else {
        // 方式2：交互式输入
        std::wcout << L"Please enter registration info file path: ";
        std::getline(std::wcin, registrationFile);

        std::wcout << L"Please enter score list file path: ";
        std::getline(std::wcin, scoreFile);

        std::wcout << L"Please enter output result file path: ";
        std::getline(std::wcin, outputFile);
    }

    // --------------------------------------------------------
    // 检测文件格式
    // --------------------------------------------------------
    FileFormat regFormat = DetectFileFormat(registrationFile);
    FileFormat scoreFormat = DetectFileFormat(scoreFile);
    FileFormat outputFormat = DetectFileFormat(outputFile);

    // --------------------------------------------------------
    // 验证格式是否支持
    // --------------------------------------------------------
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

    // --------------------------------------------------------
    // 开始处理
    // --------------------------------------------------------
    std::wcout << std::endl;
    std::wcout << L"Processing..." << std::endl;
    std::wcout << std::endl;

    // 创建读取器和处理器实例
    ExcelReader excelReader;      // Excel文件读取器
    CsvReader csvReader;          // CSV文件读取器
    DataProcessor dataProcessor;  // 数据处理器

    // 数据容器
    std::vector<Participant> participants;  // 报名信息列表
    std::vector<ScoreEntry> scoreEntries;   // 成绩条目列表
    std::vector<ResultEntry> results;       // 结果条目列表

    // --------------------------------------------------------
    // 步骤1：读取报名信息
    // --------------------------------------------------------
    std::wcout << L"1. Reading registration info..." << std::endl;
    bool regSuccess = false;
    if (regFormat == FileFormat::Excel) {
        // 使用Excel读取器
        regSuccess = excelReader.ReadRegistrationInfo(registrationFile, participants);
    }
    else {
        // 使用CSV读取器
        regSuccess = csvReader.ReadRegistrationInfo(registrationFile, participants);
    }

    if (!regSuccess) {
        std::wcerr << L"Error: Failed to read registration info file" << std::endl;
        return 1;
    }
    std::wcout << L"   Successfully read " << participants.size() << L" registration entries" << std::endl;

    // --------------------------------------------------------
    // 步骤2：读取成绩清单
    // --------------------------------------------------------
    std::wcout << L"2. Reading score list..." << std::endl;
    bool scoreSuccess = false;
    if (scoreFormat == FileFormat::Excel) {
        // 使用Excel读取器
        scoreSuccess = excelReader.ReadScoreList(scoreFile, scoreEntries);
    }
    else {
        // 使用CSV读取器
        scoreSuccess = csvReader.ReadScoreList(scoreFile, scoreEntries);
    }

    if (!scoreSuccess) {
        std::wcerr << L"Error: Failed to read score list file" << std::endl;
        return 1;
    }
    std::wcout << L"   Successfully read " << scoreEntries.size() << L" score entries" << std::endl;

    // --------------------------------------------------------
    // 步骤3：处理数据匹配
    // --------------------------------------------------------
    // 核心算法：
    // 1. 建立两个映射表：男生组号->姓名，女生组号->姓名
    // 2. 遍历成绩条目，根据组号分别查找男生和女生姓名
    // 3. 拼接姓名后生成结果
    // --------------------------------------------------------
    std::wcout << L"3. Processing data matching..." << std::endl;
    if (!dataProcessor.ProcessData(participants, scoreEntries, results)) {
        std::wcerr << L"Error: Data processing failed" << std::endl;
        return 1;
    }
    std::wcout << L"   Successfully processed " << results.size() << L" result entries" << std::endl;

    // --------------------------------------------------------
    // 步骤4：导出结果
    // --------------------------------------------------------
    std::wcout << L"4. Exporting results..." << std::endl;
    bool exportSuccess = false;
    if (outputFormat == FileFormat::Excel) {
        // 导出到Excel
        exportSuccess = dataProcessor.ExportResults(outputFile, results);
    }
    else {
        // 导出到CSV
        exportSuccess = dataProcessor.ExportResultsToCsv(outputFile, results);
    }

    if (!exportSuccess) {
        std::wcerr << L"Error: Failed to export result file" << std::endl;
        return 1;
    }
    std::wcout << L"   Successfully exported to: " << outputFile << std::endl;

    // --------------------------------------------------------
    // 显示处理结果
    // --------------------------------------------------------
    std::wcout << std::endl;
    std::wcout << L"========================================" << std::endl;
    std::wcout << L"       Processing Complete!" << std::endl;
    std::wcout << L"========================================" << std::endl;
    std::wcout << std::endl;

    // 显示处理摘要
    std::wcout << L"Processing Summary:" << std::endl;
    std::wcout << L"  - Registration Info: " << participants.size() << L" entries" << std::endl;
    std::wcout << L"  - Score Records: " << scoreEntries.size() << L" entries" << std::endl;
    std::wcout << L"  - Output Results: " << results.size() << L" entries" << std::endl;
    std::wcout << std::endl;

    // --------------------------------------------------------
    // 显示前5条结果预览
    // --------------------------------------------------------
    if (results.size() > 0) {
        std::wcout << L"First 5 results preview:" << std::endl;
        std::wcout << L"----------------------------------------" << std::endl;
        std::wcout << L"Rank\tGroup\tNames\t\tScore" << std::endl;
        std::wcout << L"----------------------------------------" << std::endl;
        
        // 最多显示5条
        size_t previewCount = std::min(results.size(), (size_t)5);
        for (size_t i = 0; i < previewCount; i++) {
            std::wcout << results[i].rank << L"\t"
                       << results[i].group << L"\t"
                       << results[i].names << L"\t"
                       << results[i].time << std::endl;
        }
        std::wcout << L"----------------------------------------" << std::endl;
    }

    // --------------------------------------------------------
    // 等待用户按键退出
    // --------------------------------------------------------
    std::wcout << std::endl;
    std::wcout << L"Press any key to exit..." << std::endl;
    std::wcin.get();

    return 0;
}
