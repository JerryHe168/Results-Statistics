#pragma once
#include "DataTypes.h"
#include <vector>
#include <string>

/**
 * @class DataProcessor
 * @brief 数据处理器类
 * 
 * 负责数据匹配和结果导出功能。
 */
class DataProcessor {
public:
    /**
     * @brief 构造函数
     */
    DataProcessor();

    /**
     * @brief 析构函数
     */
    ~DataProcessor();

    /**
     * @brief 数据匹配处理
     * 
     * 根据组别匹配男生和女生姓名，生成结果条目。
     * 
     * @param participants 报名信息列表
     * @param scoreEntries 成绩条目列表
     * @param results 结果列表
     * @return true-处理成功，false-处理失败
     */
    bool ProcessData(const std::vector<Participant>& participants,
                     const std::vector<ScoreEntry>& scoreEntries,
                     std::vector<ResultEntry>& results);

    /**
     * @brief 导出结果到Excel文件
     * 
     * 使用COM自动化技术创建Excel文件并写入结果数据。
     * 
     * @param filePath 输出文件路径
     * @param results 结果列表
     * @return true-导出成功，false-导出失败
     */
    bool ExportResults(const std::wstring& filePath, const std::vector<ResultEntry>& results);

    /**
     * @brief 导出结果到CSV文件
     * 
     * 使用UTF-8编码。
     * 
     * @param filePath 输出文件路径
     * @param results 结果列表
     * @return true-导出成功，false-导出失败
     */
    bool ExportResultsToCsv(const std::wstring& filePath, const std::vector<ResultEntry>& results);

private:
    /**
     * @brief 宽字符字符串转换为UTF-8字符串
     * 
     * 使用Windows API WideCharToMultiByte进行编码转换。
     * 
     * @param wstr 宽字符字符串（UTF-16）
     * @return UTF-8编码的字符串
     */
    std::string WStringToString(const std::wstring& wstr) const;

    /**
     * @brief CSV字段转义
     * 
     * 处理包含逗号、双引号的CSV字段。
     * 
     * @param field 原始字段值
     * @return 转义后的CSV字段字符串
     */
    std::string EscapeCsvField(const std::wstring& field) const;
};
