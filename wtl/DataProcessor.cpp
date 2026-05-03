#include "stdafx.h"
#pragma execution_character_set("utf-8")

/**
 * @class DataProcessor
 * @brief 数据处理器类
 * 
 * 负责数据匹配和结果导出功能。
 */

#include "DataProcessor.h"
#include "ExcelWriter.h"
#include <windows.h>
#include <comutil.h>
#include <iostream>
#include <fstream>
#include <sstream>
#include <iomanip>
#include <algorithm>

#pragma comment(lib, "comsuppw.lib")

/**
 * @brief 构造函数
 */
DataProcessor::DataProcessor() {
}

/**
 * @brief 析构函数
 */
DataProcessor::~DataProcessor() {
}

/**
 * @brief 数据匹配处理
 * 
 * 根据组别匹配男生和女生姓名，生成结果条目。
 * 
 * @param participants 报名信息列表
 * @param scoreEntries 成绩条目列表
 * @param results 结果列表
 */
void DataProcessor::ProcessData(const std::vector<Participant>& participants,
                                 const std::vector<ScoreEntry>& scoreEntries,
                                 std::vector<ResultEntry>& results) {
    results.clear();

    if (participants.empty()) {
        std::wcerr << L"Warning: No participants data available" << std::endl;
    }

    if (scoreEntries.empty()) {
        std::wcerr << L"Error: No score entries available" << std::endl;
        return;
    }

    // 双映射表设计：使用两个独立的映射表分别存储男生和女生信息
    // 
    // 为什么使用两个映射表：
    // 1. 同一组的男生和女生可能在报名信息的不同行
    //    例如：成绩清单23组：罗晓东 梁馨尹
    //         罗晓东（男生23A）可能在第3行
    //         梁馨尹（女生23B）可能在第5行
    // 2. 如果假设每一行的男生和女生是同一组的，会导致匹配错误
    // 3. 双映射表允许男生和女生独立映射到组号
    //
    // 映射表结构：key = 组号，value = 姓名
    std::unordered_map<int, std::wstring> maleMap;
    std::unordered_map<int, std::wstring> femaleMap;

    // 建立映射表：遍历所有报名信息
    for (const auto& participant : participants) {
        // 如果男生组号有效且姓名非空，添加到男生映射表
        if (participant.maleGroupNumber >= 0 && !participant.maleName.empty()) {
            maleMap[participant.maleGroupNumber] = participant.maleName;
        }
        // 如果女生组号有效且姓名非空，添加到女生映射表
        if (participant.femaleGroupNumber >= 0 && !participant.femaleName.empty()) {
            femaleMap[participant.femaleGroupNumber] = participant.femaleName;
        }
    }

    // 遍历成绩条目，匹配姓名
    for (const auto& scoreEntry : scoreEntries) {
        ResultEntry result;
        result.rank = scoreEntry.rank;
        result.group = scoreEntry.group;
        result.time = scoreEntry.time;

        // 组号无效的情况
        if (scoreEntry.groupNumber < 0) {
            std::wcerr << L"Warning: Invalid group number for rank " << scoreEntry.rank << std::endl;
            result.names = L"Invalid Group";
        }
        else {
            std::wstring maleName;
            std::wstring femaleName;

            // 从男生映射表查找当前组的男生姓名
            auto maleIt = maleMap.find(scoreEntry.groupNumber);
            if (maleIt != maleMap.end()) {
                maleName = maleIt->second;
            }

            // 从女生映射表查找当前组的女生姓名
            auto femaleIt = femaleMap.find(scoreEntry.groupNumber);
            if (femaleIt != femaleMap.end()) {
                femaleName = femaleIt->second;
            }

            // 姓名组合逻辑：
            // - 男生和女生都没找到：标记为 Unknown
            // - 只有男生：使用男生姓名
            // - 只有女生：使用女生姓名
            // - 都有：男生姓名 + 空格 + 女生姓名
            if (maleName.empty() && femaleName.empty()) {
                std::wcerr << L"Warning: Participant info not found for group number " << scoreEntry.groupNumber << std::endl;
                result.names = L"Unknown";
            }
            else if (maleName.empty()) {
                result.names = femaleName;
            }
            else if (femaleName.empty()) {
                result.names = maleName;
            }
            else {
                result.names = maleName + L" " + femaleName;
            }
        }

        results.push_back(result);
    }
}

/**
 * @brief 导出结果到Excel文件
 * 
 * 使用 ExcelWriter 类创建Excel文件并写入结果数据。
 * 
 * @param filePath 输出文件路径
 * @param results 结果列表
 * @return true-导出成功，false-导出失败
 */
bool DataProcessor::ExportResults(const std::wstring& filePath, const std::vector<ResultEntry>& results) {
    std::vector<std::wstring> defaultHeaders;
    defaultHeaders.push_back(L"名次");
    defaultHeaders.push_back(L"组别");
    defaultHeaders.push_back(L"姓名");
    defaultHeaders.push_back(L"成绩");
    return ExportResults(filePath, results, defaultHeaders);
}

/**
 * @brief 导出结果到Excel文件（使用自定义表头）
 */
bool DataProcessor::ExportResults(const std::wstring& filePath, 
                                   const std::vector<ResultEntry>& results,
                                   const std::vector<std::wstring>& headers) {
    if (results.empty()) {
        std::wcerr << L"Warning: No results to export" << std::endl;
    }

    ExcelWriter writer;

    // 创建新工作簿
    if (!writer.CreateNewWorkbook()) {
        return false;
    }

    // 写入表头
    for (size_t i = 0; i < headers.size(); i++) {
        writer.WriteCell(1, (long)(i + 1), headers[i]);
    }

    // 写入数据行
    for (size_t i = 0; i < results.size(); i++) {
        long row = (long)(i + 2);
        writer.WriteCell(row, 1, results[i].rank);
        writer.WriteCell(row, 2, results[i].group);
        writer.WriteCell(row, 3, results[i].names);
        writer.WriteCell(row, 4, results[i].time);
    }

    // 保存并关闭
    return writer.SaveAndClose(filePath);
}

/**
 * @brief 宽字符字符串转换为UTF-8字符串
 * 
 * 使用Windows API WideCharToMultiByte进行编码转换。
 * 
 * @param wstr 宽字符字符串（UTF-16）
 * @return UTF-8编码的字符串
 */
std::string DataProcessor::WStringToString(const std::wstring& wstr) const {
    if (wstr.empty()) {
        return std::string();
    }

    int size = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, NULL, 0, NULL, NULL);
    if (size <= 0) {
        return std::string();
    }

    std::string result(size - 1, 0);
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &result[0], size, NULL, NULL);

    return result;
}

/**
 * @brief CSV字段转义
 * 
 * 处理包含逗号、双引号的CSV字段。
 * 
 * @param field 原始字段值
 * @return 转义后的CSV字段字符串
 */
std::string DataProcessor::EscapeCsvField(const std::wstring& field) const {
    std::string str = WStringToString(field);
    
    bool needsQuotes = false;
    if (str.find(',') != std::string::npos ||
        str.find('"') != std::string::npos ||
        str.find('\n') != std::string::npos ||
        str.find('\r') != std::string::npos) {
        needsQuotes = true;
    }

    if (!needsQuotes) {
        return str;
    }

    std::string escaped;
    escaped += '"';

    for (char c : str) {
        if (c == '"') {
            escaped += "\"\"";
        }
        else {
            escaped += c;
        }
    }

    escaped += '"';
    return escaped;
}

/**
 * @brief 导出结果到CSV文件
 * 
 * 使用UTF-8编码。
 * 
 * @param filePath 输出文件路径
 * @param results 结果列表
 * @return true-导出成功，false-导出失败
 */
bool DataProcessor::ExportResultsToCsv(const std::wstring& filePath, const std::vector<ResultEntry>& results) {
    std::vector<std::wstring> defaultHeaders;
    defaultHeaders.push_back(L"名次");
    defaultHeaders.push_back(L"组别");
    defaultHeaders.push_back(L"姓名");
    defaultHeaders.push_back(L"成绩");
    return ExportResultsToCsv(filePath, results, defaultHeaders);
}

/**
 * @brief 导出结果到CSV文件（使用自定义表头）
 */
bool DataProcessor::ExportResultsToCsv(const std::wstring& filePath, 
                                        const std::vector<ResultEntry>& results,
                                        const std::vector<std::wstring>& headers) {
    if (results.empty()) {
        std::wcerr << L"Warning: No results to export" << std::endl;
    }

    FILE* file = NULL;
    // _wfopen_s 使用宽字符路径，支持中文路径
    // "wb" 表示以二进制写入模式打开
    errno_t err = _wfopen_s(&file, filePath.c_str(), L"wb");
    if (err != 0 || file == NULL) {
        std::wcerr << L"Failed to create CSV file: " << filePath << std::endl;
        return false;
    }

    // UTF-8 BOM（字节顺序标记）：0xEF 0xBB 0xBF
    unsigned char bom[] = { 0xEF, 0xBB, 0xBF };
    fwrite(bom, 1, sizeof(bom), file);

    // 写入表头
    for (size_t i = 0; i < headers.size(); i++) {
        if (i > 0) {
            fprintf(file, ",");
        }
        fprintf(file, "%s", EscapeCsvField(headers[i]).c_str());
    }
    fprintf(file, "\r\n");

    for (const auto& result : results) {
        fprintf(file, "%d,", result.rank);
        // 使用 EscapeCsvField 函数转义字段，处理包含逗号、双引号的情况
        fprintf(file, "%s,", EscapeCsvField(result.group).c_str());
        fprintf(file, "%s,", EscapeCsvField(result.names).c_str());
        fprintf(file, "%s\r\n", EscapeCsvField(result.time).c_str());
    }

    fclose(file);

    return true;
}

/**
 * @brief 检测文件格式
 */
FileFormat DataProcessor::DetectFileFormat(const std::wstring& filePath) {
    std::wstring lowerPath = filePath;
    std::transform(lowerPath.begin(), lowerPath.end(), lowerPath.begin(), ::towlower);

    if (lowerPath.length() >= 5) {
        std::wstring ext = lowerPath.substr(lowerPath.length() - 5);
        if (ext == L".xlsx") {
            return FileFormat::Excel;
        }
    }

    if (lowerPath.length() >= 4) {
        std::wstring ext = lowerPath.substr(lowerPath.length() - 4);
        if (ext == L".xls") {
            return FileFormat::Excel;
        }
        if (ext == L".csv") {
            return FileFormat::Csv;
        }
    }

    return FileFormat::Unknown;
}
