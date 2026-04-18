#pragma once
#include <windows.h>
#include <string>

/**
 * @class ExcelComHelper
 * @brief Excel COM 操作辅助类
 * 
 * 提供通用的 COM 调用静态方法和单元格值转换方法，
 * 供 ExcelSession 内部使用。
 */
class ExcelComHelper {
public:
    /**
     * @brief 获取 COM 对象的属性（带错误输出）
     * @param pDispatch COM 对象指针
     * @param propertyName 属性名
     * @param result 输出结果
     * @return true-成功，false-失败
     */
    static bool GetProperty(IDispatch* pDispatch, const wchar_t* propertyName, VARIANT& result);

    /**
     * @brief 设置 COM 对象的属性（不带错误输出，用于 Visible 这类非关键属性）
     * @param pDispatch COM 对象指针
     * @param propertyName 属性名
     * @param value 属性值
     */
    static void SetPropertyNoFail(IDispatch* pDispatch, const wchar_t* propertyName, VARIANT& value);

    /**
     * @brief 调用 COM 对象的方法（带错误输出）
     * @param pDispatch COM 对象指针
     * @param methodName 方法名
     * @param args 参数数组
     * @param argCount 参数个数
     * @param result 输出结果
     * @return true-成功，false-失败
     */
    static bool InvokeMethod(IDispatch* pDispatch, const wchar_t* methodName, VARIANT* args, int argCount, VARIANT& result);

    /**
     * @brief 将 VARIANT 转换为字符串
     * @param var 源 VARIANT
     * @param defaultVal 默认值
     * @return 转换后的字符串
     */
    static std::wstring VariantToString(const VARIANT& var, const std::wstring& defaultVal = L"");

    /**
     * @brief 将 VARIANT 转换为整数
     * @param var 源 VARIANT
     * @param defaultVal 默认值
     * @return 转换后的整数
     */
    static long VariantToLong(const VARIANT& var, long defaultVal = 0);

    /**
     * @brief 将 VARIANT 转换为浮点数
     * @param var 源 VARIANT
     * @param defaultVal 默认值
     * @return 转换后的浮点数
     */
    static double VariantToDouble(const VARIANT& var, double defaultVal = 0.0);

    /**
     * @brief 将 VARIANT 转换为时间字符串
     * @param var 源 VARIANT
     * @return 格式化的时间字符串（HH:MM:SS）
     */
    static std::wstring VariantToTime(const VARIANT& var);
};
