#pragma execution_character_set("utf-8")

#include "ExcelComHelper.h"
#include <windows.h>
#include <comutil.h>
#include <iostream>
#include <string>

#pragma comment(lib, "comsuppw.lib")

/**
 * @brief 安全释放 COM 对象
 */
void ExcelComHelper::SafeRelease(IDispatch*& pDispatch) {
    if (pDispatch) {
        pDispatch->Release();
        pDispatch = NULL;
    }
}

/**
 * @brief 获取 COM 对象的属性（带错误输出）
 */
bool ExcelComHelper::GetProperty(IDispatch* pDispatch, const wchar_t* propertyName, VARIANT& result) {
    if (!pDispatch) {
        return false;
    }

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(propertyName);
    HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get " << propertyName << L" property. HRESULT: " << hr << std::endl;
        return false;
    }

    VariantInit(&result);
    DISPPARAMS dpNoArgs = { NULL, NULL, 0, 0 };
    hr = pDispatch->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get " << propertyName << L". HRESULT: " << hr << std::endl;
        return false;
    }
    return true;
}

/**
 * @brief 获取 COM 对象的属性（返回 IDispatch*）
 */
IDispatch* ExcelComHelper::GetPropertyDispatch(IDispatch* pDispatch, const wchar_t* propertyName) {
    VARIANT result;
    if (!GetProperty(pDispatch, propertyName, result)) {
        return NULL;
    }
    if (result.vt == VT_DISPATCH) {
        return result.pdispVal;
    }
    VariantClear(&result);
    return NULL;
}

/**
 * @brief 设置 COM 对象的属性（带错误输出）
 */
bool ExcelComHelper::SetProperty(IDispatch* pDispatch, const wchar_t* propertyName, const VARIANT& value) {
    if (!pDispatch) {
        return false;
    }

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(propertyName);
    HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get " << propertyName << L" property. HRESULT: " << hr << std::endl;
        return false;
    }

    VARIANT arg = value;
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPPARAMS dp;
    dp.cArgs = 1;
    dp.rgvarg = &arg;
    dp.cNamedArgs = 1;
    dp.rgdispidNamedArgs = &dispidNamed;

    hr = pDispatch->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dp, NULL, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to set " << propertyName << L". HRESULT: " << hr << std::endl;
        return false;
    }
    return true;
}

/**
 * @brief 设置 COM 对象的属性（不带错误输出，用于 Visible 这类非关键属性）
 */
void ExcelComHelper::SetPropertyNoFail(IDispatch* pDispatch, const wchar_t* propertyName, VARIANT& value) {
    if (!pDispatch) {
        return;
    }

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(propertyName);
    HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (SUCCEEDED(hr)) {
        DISPPARAMS dp;
        dp.cArgs = 1;
        dp.rgvarg = &value;
        dp.cNamedArgs = 0;
        dp.rgdispidNamedArgs = NULL;
        pDispatch->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dp, NULL, NULL, NULL);
    }
}

/**
 * @brief 获取集合中的项目（单个整数索引）
 */
IDispatch* ExcelComHelper::GetItem(IDispatch* pDispatch, long index) {
    if (!pDispatch) {
        return NULL;
    }

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(L"Item");
    HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Item method. HRESULT: " << hr << std::endl;
        return NULL;
    }

    VARIANT arg;
    VariantInit(&arg);
    arg.vt = VT_I4;
    arg.lVal = index;

    DISPPARAMS dp;
    dp.cArgs = 1;
    dp.rgvarg = &arg;
    dp.cNamedArgs = 0;
    dp.rgdispidNamedArgs = NULL;

    VARIANT result;
    VariantInit(&result);
    hr = pDispatch->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to invoke Item method. HRESULT: " << hr << std::endl;
        return NULL;
    }

    if (result.vt == VT_DISPATCH) {
        return result.pdispVal;
    }
    std::wcerr << L"Item method returned non-dispatch type. Type: " << result.vt << std::endl;
    VariantClear(&result);
    return NULL;
}

/**
 * @brief 获取集合中的项目（两个整数索引，用于 Cells）
 */
IDispatch* ExcelComHelper::GetItem(IDispatch* pDispatch, long index1, long index2) {
    if (!pDispatch) {
        return NULL;
    }

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(L"Item");
    HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Item method. HRESULT: " << hr << std::endl;
        return NULL;
    }

    VARIANT args[2];
    VariantInit(&args[0]);
    args[0].vt = VT_I4;
    args[0].lVal = index2;

    VariantInit(&args[1]);
    args[1].vt = VT_I4;
    args[1].lVal = index1;

    DISPPARAMS dp;
    dp.cArgs = 2;
    dp.rgvarg = args;
    dp.cNamedArgs = 0;
    dp.rgdispidNamedArgs = NULL;

    VARIANT result;
    VariantInit(&result);
    hr = pDispatch->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to invoke Item method. HRESULT: " << hr << std::endl;
        return NULL;
    }

    if (result.vt == VT_DISPATCH) {
        return result.pdispVal;
    }
    std::wcerr << L"Item method returned non-dispatch type. Type: " << result.vt << std::endl;
    VariantClear(&result);
    return NULL;
}

/**
 * @brief 调用 COM 对象的方法（带错误输出）
 */
bool ExcelComHelper::InvokeMethod(IDispatch* pDispatch, const wchar_t* methodName, VARIANT* args, int argCount, VARIANT& result) {
    if (!pDispatch) {
        return false;
    }

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(methodName);
    HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get " << methodName << L" method. HRESULT: " << hr << std::endl;
        return false;
    }

    DISPPARAMS dp;
    dp.cArgs = argCount;
    dp.rgvarg = args;
    dp.cNamedArgs = 0;
    dp.rgdispidNamedArgs = NULL;

    VariantInit(&result);
    hr = pDispatch->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL);
    if (FAILED(hr)) {
        return false;
    }
    return true;
}

/**
 * @brief 调用 COM 对象的方法（无参数，无返回值）
 */
bool ExcelComHelper::InvokeMethod(IDispatch* pDispatch, const wchar_t* methodName) {
    VARIANT result;
    VariantInit(&result);
    bool success = InvokeMethod(pDispatch, methodName, NULL, 0, result);
    VariantClear(&result);
    return success;
}

/**
 * @brief 调用 COM 对象的方法（无参数，有返回值）
 */
bool ExcelComHelper::InvokeMethod(IDispatch* pDispatch, const wchar_t* methodName, VARIANT* result) {
    if (!result) {
        return InvokeMethod(pDispatch, methodName);
    }
    return InvokeMethod(pDispatch, methodName, NULL, 0, *result);
}

/**
 * @brief 调用 COM 对象的方法（1个参数，无返回值）
 */
bool ExcelComHelper::InvokeMethod(IDispatch* pDispatch, const wchar_t* methodName, const VARIANT& arg1) {
    VARIANT args[1];
    args[0] = arg1;
    VARIANT result;
    VariantInit(&result);
    bool success = InvokeMethod(pDispatch, methodName, args, 1, result);
    VariantClear(&result);
    return success;
}

/**
 * @brief 调用 COM 对象的方法（2个参数，可选返回值）
 */
bool ExcelComHelper::InvokeMethod(IDispatch* pDispatch, const wchar_t* methodName, const VARIANT& arg1, const VARIANT& arg2, VARIANT* result) {
    VARIANT args[2];
    args[0] = arg2;
    args[1] = arg1;
    if (result) {
        return InvokeMethod(pDispatch, methodName, args, 2, *result);
    }
    VARIANT dummyResult;
    VariantInit(&dummyResult);
    bool success = InvokeMethod(pDispatch, methodName, args, 2, dummyResult);
    VariantClear(&dummyResult);
    return success;
}

/**
 * @brief 将 VARIANT 转换为字符串
 */
std::wstring ExcelComHelper::VariantToString(const VARIANT& var, const std::wstring& defaultVal) {
    std::wstring result = defaultVal;

    if (var.vt == VT_BSTR) {
        result = var.bstrVal;
    }
    else if (var.vt == VT_I4) {
        result = std::to_wstring(var.lVal);
    }
    else if (var.vt == VT_R8) {
        result = std::to_wstring((long long)var.dblVal);
    }

    return result;
}

/**
 * @brief 将 VARIANT 转换为整数
 */
long ExcelComHelper::VariantToLong(const VARIANT& var, long defaultVal) {
    long result = defaultVal;

    if (var.vt == VT_I4) {
        result = var.lVal;
    }
    else if (var.vt == VT_R8) {
        result = (long)var.dblVal;
    }
    else if (var.vt == VT_BSTR) {
        try {
            result = std::stoi(var.bstrVal);
        }
        catch (...) {
            result = defaultVal;
        }
    }

    return result;
}

/**
 * @brief 将 VARIANT 转换为浮点数
 */
double ExcelComHelper::VariantToDouble(const VARIANT& var, double defaultVal) {
    double result = defaultVal;

    if (var.vt == VT_R8) {
        result = var.dblVal;
    }
    else if (var.vt == VT_I4) {
        result = (double)var.lVal;
    }

    return result;
}

/**
 * @brief 将 VARIANT 转换为时间字符串
 */
std::wstring ExcelComHelper::VariantToTime(const VARIANT& var) {
    std::wstring result;

    if (var.vt == VT_BSTR) {
        result = var.bstrVal;
    }
    else if (var.vt == VT_DATE) {
        SYSTEMTIME st;
        VariantTimeToSystemTime(var.date, &st);
        wchar_t buffer[32];
        swprintf_s(buffer, L"%d:%02d:%02d", st.wHour, st.wMinute, st.wSecond);
        result = buffer;
    }
    else if (var.vt == VT_R8) {
        double timeVal = var.dblVal;
        int hours = (int)(timeVal * 24);
        int minutes = (int)((timeVal * 24 - hours) * 60);
        int seconds = (int)(((timeVal * 24 - hours) * 60 - minutes) * 60);
        wchar_t buffer[32];
        swprintf_s(buffer, L"%d:%02d:%02d", hours, minutes, seconds);
        result = buffer;
    }

    return result;
}
