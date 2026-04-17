#pragma execution_character_set("utf-8")

#include "ExcelReader.h"
#include <windows.h>
#include <comutil.h>
#include <iostream>
#include <regex>
#include <algorithm>

#pragma comment(lib, "comsuppw.lib")

/**
 * @brief 构造函数
 */
ExcelReader::ExcelReader() {
}

/**
 * @brief 析构函数
 */
ExcelReader::~ExcelReader() {
}

/**
 * @brief 从编号中提取组号
 * 
 * 使用正则表达式匹配字符串中的第一个连续数字序列。
 * 
 * @param id 编号字符串
 * @return 提取的组号，无法提取则返回-1
 */
int ExcelReader::ExtractGroupNumber(const std::wstring& id) const {
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
 * @brief 读取报名信息Excel文件
 * 
 * 使用COM自动化技术，解析男生编号、男生姓名、女生编号、女生姓名。
 * 
 * @param filePath Excel文件路径
 * @param participants 输出参数，存储读取到的报名信息列表
 * @return true-读取成功，false-读取失败
 */
bool ExcelReader::ReadRegistrationInfo(const std::wstring& filePath, std::vector<Participant>& participants) {
    participants.clear();

    ExcelSession session;
    if (!session.OpenFile(filePath)) {
        return false;
    }

    // 报名信息的列结构：
    // 列 1：男生编号（如 "23A", "18A"）
    // 列 2：男生姓名
    // 列 3：女生编号（如 "16B", "13B"）
    // 列 4：女生姓名

    // 遍历数据行（从 lBound + 1 开始是为了跳过表头行）
    long lBound = session.GetRowLowerBound();
    long uBound = session.GetRowUpperBound();

    for (long row = lBound + 1; row <= uBound; row++) {
        Participant participant;

        // 读取第 1 列：男生编号
        participant.maleId = session.GetCellString(row, 1);

        // 读取第 2 列：男生姓名
        participant.maleName = session.GetCellString(row, 2);

        // 读取第 3 列：女生编号
        participant.femaleId = session.GetCellString(row, 3);

        // 读取第 4 列：女生姓名
        participant.femaleName = session.GetCellString(row, 4);

        // 从编号中提取组号
        participant.maleGroupNumber = ExtractGroupNumber(participant.maleId);
        participant.femaleGroupNumber = ExtractGroupNumber(participant.femaleId);

        // 只添加有姓名的记录
        if (!participant.maleName.empty() || !participant.femaleName.empty()) {
            participants.push_back(participant);
        }
    }

    return true;
}

/**
 * @brief 读取成绩清单Excel文件
 * 
 * 使用COM自动化技术，解析名次、组别、成绩时间。
 * 
 * @param filePath Excel文件路径
 * @param scoreEntries 输出参数，存储读取到的成绩条目列表
 * @return true-读取成功，false-读取失败
 */
bool ExcelReader::ReadScoreList(const std::wstring& filePath, std::vector<ScoreEntry>& scoreEntries) {
    scoreEntries.clear();

    ExcelSession session;
    if (!session.OpenFile(filePath)) {
        return false;
    }

    // 成绩清单的列结构：
    // 列 1：名次
    // 列 2：组别（如 "23组"）
    // 列 3：成绩时间

    // 遍历数据行（从 lBound + 1 开始是为了跳过表头行）
    long lBound = session.GetRowLowerBound();
    long uBound = session.GetRowUpperBound();

    for (long row = lBound + 1; row <= uBound; row++) {
        ScoreEntry entry;

        // 读取第 1 列：名次
        entry.rank = session.GetCellLong(row, 1, 0);

        // 读取第 2 列：组别
        // 组别可能是字符串（如 "23组"）或数字
        VARIANT cellValue;
        if (session.GetCellValue(row, 2, cellValue)) {
            if (cellValue.vt == VT_BSTR) {
                entry.group = cellValue.bstrVal;
            }
            else if (cellValue.vt == VT_I4) {
                // 数字格式的组别，添加 "组" 后缀
                entry.group = std::to_wstring(cellValue.lVal) + L"组";
            }
            else if (cellValue.vt == VT_R8) {
                entry.group = std::to_wstring((long)cellValue.dblVal) + L"组";
            }
            VariantClear(&cellValue);
        }

        // 读取第 3 列：成绩时间
        entry.time = session.GetCellTime(row, 3);

        // 从组别中提取组号
        entry.groupNumber = ExtractGroupNumber(entry.group);

        // 只添加有效名次的记录
        if (entry.rank > 0) {
            scoreEntries.push_back(entry);
        }
    }

    return true;
}
