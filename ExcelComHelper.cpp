#pragma execution_character_set("utf-8")

#include "ExcelComHelper.h"
#include <windows.h>
#include <comutil.h>
#include <iostream>
#include <string>

#pragma comment(lib, "comsuppw.lib")

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
