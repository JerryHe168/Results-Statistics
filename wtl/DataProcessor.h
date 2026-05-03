#pragma once
#include "DataTypes.h"
#include <vector>
#include <string>

enum class FileFormat {
    Excel,
    Csv,
    Unknown
};

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
     */
    void ProcessData(const std::vector<Participant>& participants,
                     const std::vector<ScoreEntry>& scoreEntries,
                     std::vector<ResultEntry>& results);

    /**
     * @brief 导出结果到Excel文件（使用默认表头）
     */
    bool ExportResults(const std::wstring& filePath, const std::vector<ResultEntry>& results);

    /**
     * @brief 导出结果到Excel文件（使用自定义表头）
     * 
     * @param headers 表头列表
     */
    bool ExportResults(const std::wstring& filePath, 
                       const std::vector<ResultEntry>& results,
                       const std::vector<std::wstring>& headers);

    /**
     * @brief 导出结果到CSV文件（使用默认表头）
     */
    bool ExportResultsToCsv(const std::wstring& filePath, const std::vector<ResultEntry>& results);

    /**
     * @brief 导出结果到CSV文件（使用自定义表头）
     * 
     * @param headers 表头列表
     */
    bool ExportResultsToCsv(const std::wstring& filePath, 
                            const std::vector<ResultEntry>& results,
                            const std::vector<std::wstring>& headers);

    /**
     * @brief 检测文件格式
     * 
     * @param filePath 文件路径
     * @return 文件格式枚举值
     */
    FileFormat DetectFileFormat(const std::wstring& filePath);

private:
    /**
     * @brief 宽字符字符串转换为UTF-8字符串
     */
    std::string WStringToString(const std::wstring& wstr) const;

    /**
     * @brief CSV字段转义
     */
    std::string EscapeCsvField(const std::wstring& field) const;
};
