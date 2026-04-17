#pragma execution_character_set("utf-8")

/**
 * @class CsvReader
 * @brief CSV文件读取器类
 * 
 * 负责读取CSV格式的报名信息和成绩清单文件。
 */

#include "CsvReader.h"
#include <windows.h>
#include <iostream>
#include <fstream>
#include <sstream>
#include <regex>
#include <algorithm>
#include <codecvt>

/**
 * @brief 构造函数
 */
CsvReader::CsvReader() {
}

/**
 * @brief 析构函数
 */
CsvReader::~CsvReader() {
}

/**
 * @brief UTF-8字符串转换为宽字符字符串
 * 
 * 使用Windows API MultiByteToWideChar进行编码转换。
 * 
 * @param str UTF-8编码的字符串
 * @return 宽字符字符串（UTF-16）
 */
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

/**
 * @brief 去除字符串首尾的空白和引号
 * 
 * @param str 原始字符串
 * @return 处理后的字符串
 */
std::wstring CsvReader::Trim(const std::wstring& str) {
    size_t start = str.find_first_not_of(L" \t\"");
    if (start == std::wstring::npos) {
        return L"";
    }

    size_t end = str.find_last_not_of(L" \t\"");
    return str.substr(start, end - start + 1);
}

/**
 * @brief 解析CSV格式的一行数据
 * 
 * 遵循RFC 4180标准解析CSV行。
 * 
 * @param line CSV格式的一行字符串
 * @return 解析后的字段列表
 */
std::vector<std::wstring> CsvReader::SplitCsvLine(const std::wstring& line) {
    std::vector<std::wstring> result;
    std::wstring current;

    // 状态机变量：标记是否处于引号包裹的字段中
    // 这是解析CSV的核心逻辑，用于处理包含逗号或双引号的字段
    bool inQuotes = false;

    for (size_t i = 0; i < line.length(); i++) {
        wchar_t c = line[i];

        if (c == L'"') {
            // 双引号有两种含义：字段边界 或 转义的双引号
            
            // RFC 4180 标准：字段内部的双引号需要用两个双引号表示（转义）
            // 例如：字段值 "Hello,World" 在CSV中表示为 ""Hello,World""
            if (inQuotes && i + 1 < line.length() && line[i + 1] == L'"') {
                current += L'"';  // 转义的双引号，添加一个双引号
                i++;              // 跳过第二个双引号
            }
            else {
                // 切换引号状态：字段开始或结束
                // 进入或退出引号包裹的字段
                inQuotes = !inQuotes;
            }
        }
        else if (c == L',' && !inQuotes) {
            // 逗号分隔符：只有在引号外部才作为字段分隔符
            // 如果在引号内部，逗号是字段内容的一部分
            result.push_back(Trim(current));
            current.clear();
        }
        else {
            // 普通字符：添加到当前字段
            current += c;
        }
    }

    // 处理最后一个字段
    result.push_back(Trim(current));
    return result;
}

/**
 * @brief 从编号中提取组号
 * 
 * 使用正则表达式匹配字符串中的第一个连续数字序列。
 * 
 * @param id 编号字符串
 * @return 提取的组号，无法提取则返回-1
 */
int CsvReader::ExtractGroupNumber(const std::wstring& id) {
    // 正则表达式 L"(\\d+)"：匹配编号中的连续数字
    // 例如："23A" → 23，"17B" → 17
    std::wregex regex(L"(\\d+)");
    std::wsmatch match;

    if (std::regex_search(id, match, regex)) {
        // match[1] 是第一个捕获组的内容
        return std::stoi(match[1].str());
    }

    return -1;
}

/**
 * @brief 读取报名信息CSV文件
 * 
 * 读取报名信息CSV文件，解析男生编号、男生姓名、女生编号、女生姓名。
 * 
 * @param filePath CSV文件路径
 * @param participants 输出参数，存储读取到的报名信息列表
 * @return true-读取成功，false-读取失败
 */
bool CsvReader::ReadRegistrationInfo(const std::wstring& filePath, std::vector<Participant>& participants) {
    participants.clear();

    FILE* file = NULL;
    errno_t err = _wfopen_s(&file, filePath.c_str(), L"rb");
    if (err != 0 || file == NULL) {
        std::wcerr << L"Failed to open CSV file: " << filePath << std::endl;
        return false;
    }

    fseek(file, 0, SEEK_END);
    long fileSize = ftell(file);
    fseek(file, 0, SEEK_SET);

    std::string content(fileSize, 0);
    fread(&content[0], 1, fileSize, file);
    fclose(file);

    if (content.length() >= 3 &&
        (unsigned char)content[0] == 0xEF &&
        (unsigned char)content[1] == 0xBB &&
        (unsigned char)content[2] == 0xBF) {
        content = content.substr(3);
    }

    std::wstring wcontent = StringToWString(content);
    std::wistringstream iss(wcontent);
    std::wstring line;

    // 表头检测逻辑：标记是否为第一行，用于检测并跳过表头
    bool isFirstLine = true;
    while (std::getline(iss, line)) {
        // 处理 Windows 换行符：Windows 用 \r\n，去除行末的 \r
        if (!line.empty() && line.back() == L'\r') {
            line.pop_back();
        }

        // 跳过空行
        if (line.empty()) {
            continue;
        }

        // 表头检测：第一行可能是表头
        // 检测标准：第一列包含"男生"、"male"或"编号"等关键词
        if (isFirstLine) {
            std::vector<std::wstring> header = SplitCsvLine(line);
            if (header.size() >= 2) {
                std::wstring lowerHeader = header[0];
                std::transform(lowerHeader.begin(), lowerHeader.end(), lowerHeader.begin(), ::towlower);

                // 检测表头关键词：中文"男生"、"编号" 或英文"male"
                if (lowerHeader.find(L"男生") != std::wstring::npos ||
                    lowerHeader.find(L"male") != std::wstring::npos ||
                    lowerHeader.find(L"编号") != std::wstring::npos) {
                    isFirstLine = false;
                    continue;  // 跳过表头行
                }
            }
            isFirstLine = false;  // 无论是否检测到表头，第一行只检查一次
        }

        // 报名信息需要至少4列：男生编号、男生姓名、女生编号、女生姓名
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

/**
 * @brief 读取成绩清单CSV文件
 * 
 * 读取成绩清单CSV文件，解析名次、组别、成绩时间。
 * 
 * @param filePath CSV文件路径
 * @param scoreEntries 输出参数，存储读取到的成绩条目列表
 * @return true-读取成功，false-读取失败
 */
bool CsvReader::ReadScoreList(const std::wstring& filePath, std::vector<ScoreEntry>& scoreEntries) {
    scoreEntries.clear();

    FILE* file = NULL;
    errno_t err = _wfopen_s(&file, filePath.c_str(), L"rb");
    if (err != 0 || file == NULL) {
        std::wcerr << L"Failed to open CSV file: " << filePath << std::endl;
        return false;
    }

    fseek(file, 0, SEEK_END);
    long fileSize = ftell(file);
    fseek(file, 0, SEEK_SET);

    std::string content(fileSize, 0);
    fread(&content[0], 1, fileSize, file);
    fclose(file);

    // UTF-8 BOM 检测和去除：0xEF 0xBB 0xBF 是 UTF-8 编码的字节顺序标记
    if (content.length() >= 3 &&
        (unsigned char)content[0] == 0xEF &&
        (unsigned char)content[1] == 0xBB &&
        (unsigned char)content[2] == 0xBF) {
        content = content.substr(3);
    }

    std::wstring wcontent = StringToWString(content);
    std::wistringstream iss(wcontent);
    std::wstring line;

    // 表头检测逻辑：标记是否为第一行，用于检测并跳过表头
    bool isFirstLine = true;
    while (std::getline(iss, line)) {
        // 处理 Windows 换行符：Windows 用 \r\n，去除行末的 \r
        if (!line.empty() && line.back() == L'\r') {
            line.pop_back();
        }

        // 跳过空行
        if (line.empty()) {
            continue;
        }

        // 表头检测：第一行可能是表头
        // 检测标准：第一列包含"名次"、"rank"或"排名"等关键词
        if (isFirstLine) {
            std::vector<std::wstring> header = SplitCsvLine(line);
            if (header.size() >= 2) {
                std::wstring lowerHeader = header[0];
                std::transform(lowerHeader.begin(), lowerHeader.end(), lowerHeader.begin(), ::towlower);

                // 检测表头关键词：中文"名次"、"排名" 或英文"rank"
                if (lowerHeader.find(L"名次") != std::wstring::npos ||
                    lowerHeader.find(L"rank") != std::wstring::npos ||
                    lowerHeader.find(L"排名") != std::wstring::npos) {
                    isFirstLine = false;
                    continue;  // 跳过表头行
                }
            }
            isFirstLine = false;  // 无论是否检测到表头，第一行只检查一次
        }

        // 成绩清单需要至少3列：名次、组别、成绩时间
        std::vector<std::wstring> columns = SplitCsvLine(line);
        if (columns.size() < 3) {
            continue;
        }

        ScoreEntry entry;

        // 名次转换：使用 try-catch 处理可能的转换异常
        try {
            entry.rank = std::stoi(columns[0]);
        }
        catch (...) {
            entry.rank = 0;
        }

        entry.group = columns[1];
        entry.time = columns[2];

        entry.groupNumber = ExtractGroupNumber(entry.group);

        // 名次过滤：只保留有效名次（名次 > 0）
        if (entry.rank > 0) {
            scoreEntries.push_back(entry);
        }
    }

    return true;
}
