#pragma once
#include <windows.h>
#include <string>

/**
 * @class ExcelComHelper
 * @brief Excel COM 操作辅助类
 * 
 * 提供通用的 COM 调用静态方法和单元格值转换方法，
 * 供 ExcelSession 和 ExcelWriter 使用。
 */
class ExcelComHelper {
public:
    ExcelComHelper(const ExcelComHelper&) = delete;
    ExcelComHelper& operator=(const ExcelComHelper&) = delete;
    /**
     * @brief 安全释放 COM 对象
     * @param pDispatch COM 对象指针引用，释放后置为 NULL
     */
    static void SafeRelease(IDispatch*& pDispatch);

    /**
     * @brief 获取 COM 对象的属性（带错误输出）
     * @param pDispatch COM 对象指针
     * @param propertyName 属性名
     * @param result 输出结果
     * @return true-成功，false-失败
     */
    static bool GetProperty(IDispatch* pDispatch, const wchar_t* propertyName, VARIANT& result);

    /**
     * @brief 获取 COM 对象的属性（返回 IDispatch*）
     * @param pDispatch COM 对象指针
     * @param propertyName 属性名
     * @return IDispatch* 指针，失败返回 NULL
     */
    static IDispatch* GetPropertyDispatch(IDispatch* pDispatch, const wchar_t* propertyName);

    /**
     * @brief 设置 COM 对象的属性（带错误输出）
     * @param pDispatch COM 对象指针
     * @param propertyName 属性名
     * @param value 属性值
     * @return true-成功，false-失败
     */
    static bool SetProperty(IDispatch* pDispatch, const wchar_t* propertyName, const VARIANT& value);

    /**
     * @brief 设置 COM 对象的属性（不带错误输出，用于 Visible 这类非关键属性）
     * @param pDispatch COM 对象指针
     * @param propertyName 属性名
     * @param value 属性值
     */
    static void SetPropertyNoFail(IDispatch* pDispatch, const wchar_t* propertyName, VARIANT& value);

    /**
     * @brief 获取集合中的项目（单个整数索引）
     * @param pDispatch COM 对象指针
     * @param index 索引
     * @return IDispatch* 指针，失败返回 NULL
     */
    static IDispatch* GetItem(IDispatch* pDispatch, long index);

    /**
     * @brief 获取集合中的项目（两个整数索引，用于 Cells）
     * @param pDispatch COM 对象指针
     * @param index1 第一个索引
     * @param index2 第二个索引
     * @return IDispatch* 指针，失败返回 NULL
     */
    static IDispatch* GetItem(IDispatch* pDispatch, long index1, long index2);

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
     * @brief 调用 COM 对象的方法（无参数，无返回值）
     * @param pDispatch COM 对象指针
     * @param methodName 方法名
     * @return true-成功，false-失败
     */
    static bool InvokeMethod(IDispatch* pDispatch, const wchar_t* methodName);

    /**
     * @brief 调用 COM 对象的方法（无参数，有返回值）
     * @param pDispatch COM 对象指针
     * @param methodName 方法名
     * @param result 输出结果
     * @return true-成功，false-失败
     */
    static bool InvokeMethod(IDispatch* pDispatch, const wchar_t* methodName, VARIANT* result);

    /**
     * @brief 调用 COM 对象的方法（1个参数，无返回值）
     * @param pDispatch COM 对象指针
     * @param methodName 方法名
     * @param arg1 第一个参数
     * @return true-成功，false-失败
     */
    static bool InvokeMethod(IDispatch* pDispatch, const wchar_t* methodName, const VARIANT& arg1);

    /**
     * @brief 调用 COM 对象的方法（2个参数，无返回值）
     * @param pDispatch COM 对象指针
     * @param methodName 方法名
     * @param arg1 第一个参数
     * @param arg2 第二个参数
     * @param result 输出结果（可为 NULL）
     * @return true-成功，false-失败
     */
    static bool InvokeMethod(IDispatch* pDispatch, const wchar_t* methodName, const VARIANT& arg1, const VARIANT& arg2, VARIANT* result = NULL);

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
