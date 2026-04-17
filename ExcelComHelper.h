#pragma once
#include <windows.h>
#include <string>

/**
 * @class ExcelComHelper
 * @brief Excel COM 通用辅助类
 * 
 * 封装常用的 COM 操作模式，减少重复代码：
 * - GetProperty: 获取属性值
 * - SetProperty: 设置属性值
 * - InvokeMethod: 调用方法
 * - GetItem: 获取集合项
 * - SafeRelease: 安全释放 COM 对象
 */
class ExcelComHelper {
public:
    /**
     * @brief 获取属性值
     * @param pDispatch COM 对象指针
     * @param propName 属性名
     * @param pResult 输出参数，存储结果
     * @return true-成功，false-失败
     */
    static bool GetProperty(IDispatch* pDispatch, const std::wstring& propName, VARIANT* pResult);

    /**
     * @brief 获取 IDispatch 类型的属性
     * @param pDispatch COM 对象指针
     * @param propName 属性名
     * @return IDispatch* 指针，失败返回 NULL
     */
    static IDispatch* GetPropertyDispatch(IDispatch* pDispatch, const std::wstring& propName);

    /**
     * @brief 设置属性值
     * @param pDispatch COM 对象指针
     * @param propName 属性名
     * @param value 属性值
     * @return true-成功，false-失败
     */
    static bool SetProperty(IDispatch* pDispatch, const std::wstring& propName, const VARIANT& value);

    /**
     * @brief 调用无参数方法
     * @param pDispatch COM 对象指针
     * @param methodName 方法名
     * @param pResult 输出参数，存储结果（可为 NULL）
     * @return true-成功，false-失败
     */
    static bool InvokeMethod(IDispatch* pDispatch, const std::wstring& methodName, VARIANT* pResult = NULL);

    /**
     * @brief 调用单参数方法
     * @param pDispatch COM 对象指针
     * @param methodName 方法名
     * @param arg1 参数1
     * @param pResult 输出参数，存储结果（可为 NULL）
     * @return true-成功，false-失败
     */
    static bool InvokeMethod(IDispatch* pDispatch, const std::wstring& methodName, 
                              const VARIANT& arg1, VARIANT* pResult = NULL);

    /**
     * @brief 调用双参数方法
     * @param pDispatch COM 对象指针
     * @param methodName 方法名
     * @param arg1 参数1
     * @param arg2 参数2
     * @param pResult 输出参数，存储结果（可为 NULL）
     * @return true-成功，false-失败
     */
    static bool InvokeMethod(IDispatch* pDispatch, const std::wstring& methodName,
                              const VARIANT& arg1, const VARIANT& arg2, VARIANT* pResult = NULL);

    /**
     * @brief 获取集合项（单参数，如 Worksheets.Item(1)）
     * @param pDispatch 集合对象指针
     * @param index 索引值
     * @return IDispatch* 指针，失败返回 NULL
     */
    static IDispatch* GetItem(IDispatch* pDispatch, long index);

    /**
     * @brief 获取集合项（双参数，如 Cells(row, col)）
     * @param pDispatch 集合对象指针
     * @param row 行号
     * @param col 列号
     * @return IDispatch* 指针，失败返回 NULL
     */
    static IDispatch* GetItem(IDispatch* pDispatch, long row, long col);

    /**
     * @brief 安全释放 COM 对象
     * @param pDispatch COM 对象指针引用
     */
    static void SafeRelease(IDispatch*& pDispatch);
};
