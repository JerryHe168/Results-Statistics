#pragma once
#include <windows.h>
#include <string>
#include <vector>

/**
 * @class ExcelWriter
 * @brief Excel COM 写入封装类
 * 
 * 使用 RAII 模式封装 Excel COM 写入操作的整个生命周期，
 * 包括创建新工作簿、写入单元格、保存文件、自动清理资源。
 */
class ExcelWriter {
public:
    /**
     * @brief 构造函数
     */
    ExcelWriter();

    /**
     * @brief 析构函数
     */
    ~ExcelWriter();

    /**
     * @brief 创建新的 Excel 工作簿
     * 
     * 创建 Excel Application 实例，新建一个工作簿，
     * 并获取第一个工作表。
     * 
     * @return true-创建成功，false-创建失败
     */
    bool CreateNewWorkbook();

    /**
     * @brief 写入单元格值（字符串）
     * 
     * @param row 行号（从 1 开始）
     * @param col 列号（从 1 开始）
     * @param value 字符串值
     * @return true-写入成功，false-写入失败
     */
    bool WriteCell(long row, long col, const std::wstring& value);

    /**
     * @brief 写入单元格值（整数）
     * 
     * @param row 行号（从 1 开始）
     * @param col 列号（从 1 开始）
     * @param value 整数值
     * @return true-写入成功，false-写入失败
     */
    bool WriteCell(long row, long col, int value);

    /**
     * @brief 保存并关闭文件
     * 
     * 根据文件扩展名选择格式：
     * - .xlsx: 使用格式代码 51
     * - .xls: 使用格式代码 56
     * - 默认: 使用格式代码 56 (.xls)
     * 
     * @param filePath 输出文件路径
     * @return true-保存成功，false-保存失败
     */
    bool SaveAndClose(const std::wstring& filePath);

private:
    IDispatch* m_pExcelApp;
    IDispatch* m_pWorkbooks;
    IDispatch* m_pWorkbook;
    IDispatch* m_pWorksheets;
    IDispatch* m_pWorksheet;

    /**
     * @brief 释放所有 COM 对象
     */
    void Release();
};
