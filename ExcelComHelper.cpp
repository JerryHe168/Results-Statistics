#pragma execution_character_set("utf-8")

#include "ExcelComHelper.h"
#include <windows.h>
#include <comutil.h>
#include <iostream>

#pragma comment(lib, "comsuppw.lib")

/**
 * @brief 获取属性值
 */
bool ExcelComHelper::GetProperty(IDispatch* pDispatch, const std::wstring& propName, VARIANT* pResult) {
    if (!pDispatch || !pResult) {
        return false;
    }

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(propName.c_str());
    HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        return false;
    }

    DISPPARAMS dpNoArgs = { NULL, NULL, 0, 0 };
    VariantInit(pResult);
    hr = pDispatch->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, pResult, NULL, NULL);

    return SUCCEEDED(hr);
}

/**
 * @brief 获取 IDispatch 类型的属性
 */
IDispatch* ExcelComHelper::GetPropertyDispatch(IDispatch* pDispatch, const std::wstring& propName) {
    VARIANT result;
    VariantInit(&result);

    if (!GetProperty(pDispatch, propName, &result)) {
        VariantClear(&result);
        return NULL;
    }

    if (result.vt != VT_DISPATCH) {
        VariantClear(&result);
        return NULL;
    }

    return result.pdispVal;
}

/**
 * @brief 设置属性值
 */
bool ExcelComHelper::SetProperty(IDispatch* pDispatch, const std::wstring& propName, const VARIANT& value) {
    if (!pDispatch) {
        return false;
    }

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(propName.c_str());
    HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        return false;
    }

    DISPPARAMS dp;
    DISPID dispidPut = DISPID_PROPERTYPUT;
    dp.cArgs = 1;
    dp.rgvarg = const_cast<VARIANT*>(&value);
    dp.cNamedArgs = 1;
    dp.rgdispidNamedArgs = &dispidPut;

    VARIANT varResult;
    VariantInit(&varResult);
    hr = pDispatch->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dp, &varResult, NULL, NULL);
    VariantClear(&varResult);

    return SUCCEEDED(hr);
}

/**
 * @brief 调用无参数方法
 */
bool ExcelComHelper::InvokeMethod(IDispatch* pDispatch, const std::wstring& methodName, VARIANT* pResult) {
    if (!pDispatch) {
        return false;
    }

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(methodName.c_str());
    HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        return false;
    }

    DISPPARAMS dpNoArgs = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = pDispatch->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpNoArgs, &result, NULL, NULL);

    if (pResult) {
        *pResult = result;
    }
    else {
        VariantClear(&result);
    }

    return SUCCEEDED(hr);
}

/**
 * @brief 调用单参数方法
 */
bool ExcelComHelper::InvokeMethod(IDispatch* pDispatch, const std::wstring& methodName,
                                    const VARIANT& arg1, VARIANT* pResult) {
    if (!pDispatch) {
        return false;
    }

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(methodName.c_str());
    HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        return false;
    }

    VARIANT args[1];
    args[0] = arg1;

    DISPPARAMS dp;
    dp.cArgs = 1;
    dp.rgvarg = args;
    dp.cNamedArgs = 0;
    dp.rgdispidNamedArgs = NULL;

    VARIANT result;
    VariantInit(&result);

    hr = pDispatch->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL);

    if (pResult) {
        *pResult = result;
    }
    else {
        VariantClear(&result);
    }

    return SUCCEEDED(hr);
}

/**
 * @brief 调用双参数方法
 */
bool ExcelComHelper::InvokeMethod(IDispatch* pDispatch, const std::wstring& methodName,
                                    const VARIANT& arg1, const VARIANT& arg2, VARIANT* pResult) {
    if (!pDispatch) {
        return false;
    }

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(methodName.c_str());
    HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        return false;
    }

    VARIANT args[2];
    args[1] = arg1;
    args[0] = arg2;

    DISPPARAMS dp;
    dp.cArgs = 2;
    dp.rgvarg = args;
    dp.cNamedArgs = 0;
    dp.rgdispidNamedArgs = NULL;

    VARIANT result;
    VariantInit(&result);

    hr = pDispatch->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL);

    if (pResult) {
        *pResult = result;
    }
    else {
        VariantClear(&result);
    }

    return SUCCEEDED(hr);
}

/**
 * @brief 获取集合项（单参数）
 */
IDispatch* ExcelComHelper::GetItem(IDispatch* pDispatch, long index) {
    if (!pDispatch) {
        return NULL;
    }

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(L"Item");
    HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        return NULL;
    }

    VARIANT arg;
    VariantInit(&arg);
    arg.vt = VT_I4;
    arg.lVal = index;

    VARIANT args[1];
    args[0] = arg;

    DISPPARAMS dp;
    dp.cArgs = 1;
    dp.rgvarg = args;
    dp.cNamedArgs = 0;
    dp.rgdispidNamedArgs = NULL;

    VARIANT result;
    VariantInit(&result);

    hr = pDispatch->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);

    if (FAILED(hr) || result.vt != VT_DISPATCH) {
        VariantClear(&result);
        return NULL;
    }

    return result.pdispVal;
}

/**
 * @brief 获取集合项（双参数）
 */
IDispatch* ExcelComHelper::GetItem(IDispatch* pDispatch, long row, long col) {
    if (!pDispatch) {
        return NULL;
    }

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(L"Cells");
    HRESULT hr = pDispatch->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        return NULL;
    }

    VARIANT args[2];
    args[1].vt = VT_I4;
    args[1].lVal = row;
    args[0].vt = VT_I4;
    args[0].lVal = col;

    DISPPARAMS dp;
    dp.cArgs = 2;
    dp.rgvarg = args;
    dp.cNamedArgs = 0;
    dp.rgdispidNamedArgs = NULL;

    VARIANT result;
    VariantInit(&result);

    hr = pDispatch->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);

    if (FAILED(hr) || result.vt != VT_DISPATCH) {
        VariantClear(&result);
        return NULL;
    }

    return result.pdispVal;
}

/**
 * @brief 安全释放 COM 对象
 */
void ExcelComHelper::SafeRelease(IDispatch*& pDispatch) {
    if (pDispatch) {
        pDispatch->Release();
        pDispatch = NULL;
    }
}
