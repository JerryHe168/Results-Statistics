#pragma once
#include "DataTypes.h"
#include <vector>
#include <string>

/**
 * @class CsvReader
 * @brief CSV文件读取器类
 * 
 * 负责读取CSV格式的报名信息和成绩清单文件。
 */
class CsvReader {
public:
    /**
     * @brief 构造函数
     */
    CsvReader();

    /**
     * @brief 析构函数
     */
    ~CsvReader();

    /**
     * @brief 读取报名信息CSV文件
     * 
     * 解析男生编号、男生姓名、女生编号、女生姓名。
     * 
     * @param filePath CSV文件路径
     * @param participants 输出参数，存储读取到的报名信息列表
     * @return true-读取成功，false-读取失败
     */
    bool ReadRegistrationInfo(const std::wstring& filePath, std::vector<Participant>& participants);

    /**
     * @brief 读取成绩清单CSV文件
     * 
     * 解析名次、组别、成绩时间。
     * 
     * @param filePath CSV文件路径
     * @param scoreEntries 输出参数，存储读取到的成绩条目列表
     * @return true-读取成功，false-读取失败
     */
    bool ReadScoreList(const std::wstring& filePath, std::vector<ScoreEntry>& scoreEntries);

private:
    /**
     * @brief 从编号中提取组号
     * 
     * 使用正则表达式匹配字符串中的第一个连续数字序列。
     * 
     * @param id 编号字符串
     * @return 提取的组号，无法提取则返回-1
     */
    int ExtractGroupNumber(const std::wstring& id) const;

    /**
     * @brief 解析CSV格式的一行数据
     * 
     * 支持双引号包裹字段和双引号转义。
     * 
     * @param line CSV格式的一行字符串
     * @return 解析后的字段列表
     */
    std::vector<std::wstring> SplitCsvLine(const std::wstring& line) const;

    /**
     * @brief 去除字符串首尾的空白和引号
     * 
     * @param str 原始字符串
     * @return 处理后的字符串
     */
    std::wstring Trim(const std::wstring& str) const;

    /**
     * @brief UTF-8字符串转换为宽字符字符串
     * 
     * 使用Windows API MultiByteToWideChar进行编码转换。
     * 
     * @param str UTF-8编码的字符串
     * @return 宽字符字符串（UTF-16）
     */
    std::wstring StringToWString(const std::string& str) const;
};
