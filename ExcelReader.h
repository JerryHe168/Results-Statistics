#pragma once
#include "DataTypes.h"
#include <vector>
#include <string>
#include <windows.h>

/**
 * @class ExcelSession
 * @brief Excel COM 会话封装类
 * 
 * 使用 RAII 模式封装 Excel COM 会话的整个生命周期，
 * 包括创建 Excel 实例、打开文件、获取数据、自动清理资源。
 */
class ExcelSession {
public:
    /**
     * @brief 构造函数
     */
    ExcelSession();

    /**
     * @brief 析构函数
     */
    ~ExcelSession();

    /**
     * @brief 打开 Excel 文件
     * 
     * 创建 Excel Application 实例，打开指定的工作簿，
     * 并获取第一个工作表的 UsedRange 数据。
     * 
     * @param filePath Excel 文件路径
     * @return true-打开成功，false-打开失败
     */
    bool OpenFile(const std::wstring& filePath);

    /**
     * @brief 获取 SAFEARRAY 数据
     * @return SAFEARRAY 指针，失败返回 NULL
     */
    SAFEARRAY* GetSafeArray() const { return m_pSafeArray; }

    /**
     * @brief 获取行下界
     * @return 行下界（通常为 1）
     */
    long GetRowLowerBound() const { return m_lBound; }

    /**
     * @brief 获取行上界
     * @return 行上界
     */
    long GetRowUpperBound() const { return m_uBound; }

    /**
     * @brief 获取单元格值（字符串）
     * 
     * @param row 行号（从 1 开始）
     * @param col 列号（从 1 开始）
     * @param defaultVal 默认值
     * @return 单元格字符串值
     */
    std::wstring GetCellString(long row, long col, const std::wstring& defaultVal = L"") const;

    /**
     * @brief 获取单元格值（整数）
     * 
     * @param row 行号（从 1 开始）
     * @param col 列号（从 1 开始）
     * @param defaultVal 默认值
     * @return 单元格整数值
     */
    long GetCellLong(long row, long col, long defaultVal = 0) const;

    /**
     * @brief 获取单元格值（浮点数）
     * 
     * @param row 行号（从 1 开始）
     * @param col 列号（从 1 开始）
     * @param defaultVal 默认值
     * @return 单元格浮点数值
     */
    double GetCellDouble(long row, long col, double defaultVal = 0.0) const;

    /**
     * @brief 获取时间单元格值
     * 
     * 处理三种时间格式：VT_BSTR（字符串）、VT_DATE（Variant时间）、VT_R8（浮点数）。
     * 
     * @param row 行号（从 1 开始）
     * @param col 列号（从 1 开始）
     * @return 格式化的时间字符串（HH:MM:SS）
     */
    std::wstring GetCellTime(long row, long col) const;

    /**
     * @brief 获取单元格值
     * 
     * @param row 行号（从 1 开始）
     * @param col 列号（从 1 开始）
     * @param cellValue 输出参数，存储单元格值
     * @return true-获取成功，false-获取失败
     */
    bool GetCellValue(long row, long col, VARIANT& cellValue) const;

private:
    IDispatch* m_pExcelApp;
    IDispatch* m_pWorkbooks;
    IDispatch* m_pWorkbook;
    IDispatch* m_pWorksheets;
    IDispatch* m_pWorksheet;
    IDispatch* m_pRange;
    VARIANT m_varResult;
    SAFEARRAY* m_pSafeArray;
    long m_lBound;
    long m_uBound;

    /**
     * @brief 释放所有 COM 对象
     */
    void Release();
};

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
