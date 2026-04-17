#pragma once
#include "DataTypes.h"
#include <vector>
#include <string>

/**
 * @class ExcelReader
 * @brief Excel文件读取器类
 * 
 * 负责使用COM自动化技术读取Excel格式的报名信息和成绩清单文件。
 */
class ExcelReader {
public:
    /**
     * @brief 构造函数
     */
    ExcelReader();

    /**
     * @brief 析构函数
     */
    ~ExcelReader();

    /**
     * @brief 读取报名信息Excel文件
     * 
     * 使用COM自动化技术，解析男生编号、男生姓名、女生编号、女生姓名。
     * 
     * @param filePath Excel文件路径
     * @param participants 输出参数，存储读取到的报名信息列表
     * @return true-读取成功，false-读取失败
     */
    bool ReadRegistrationInfo(const std::wstring& filePath, std::vector<Participant>& participants);

    /**
     * @brief 读取成绩清单Excel文件
     * 
     * 使用COM自动化技术，解析名次、组别、成绩时间。
     * 
     * @param filePath Excel文件路径
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
};
