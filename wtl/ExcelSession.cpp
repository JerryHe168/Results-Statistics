#include "stdafx.h"
#pragma execution_character_set("utf-8")

#include "ExcelSession.h"
#include <windows.h>
#include <comutil.h>
#include <iostream>

#pragma comment(lib, "comsuppw.lib")

/**
 * @brief 构造函数
 */
ExcelSession::ExcelSession()
    : m_pExcelApp(NULL), m_pWorkbooks(NULL), m_pWorkbook(NULL),
      m_pWorksheets(NULL), m_pWorksheet(NULL), m_pRange(NULL),
      m_pSafeArray(NULL), m_lBound(1), m_uBound(0) {
    VariantInit(&m_varResult);
}

/**
 * @brief 析构函数
 */
ExcelSession::~ExcelSession() {
    Release();
}

/**
 * @brief 释放所有 COM 对象
 */
void ExcelSession::Release() {
    if (m_pSafeArray) {
        VariantClear(&m_varResult);
        m_pSafeArray = NULL;
    }

    // 逆序释放 COM 对象
    // 创建顺序：pExcelApp -> pWorkbooks -> pWorkbook -> pWorksheets -> pWorksheet -> pRange
    // 释放顺序：pRange -> pWorksheet -> pWorksheets -> pWorkbook -> pWorkbooks -> pExcelApp

    if (m_pRange) {
        m_pRange->Release();
        m_pRange = NULL;
    }

    if (m_pWorksheet) {
        m_pWorksheet->Release();
        m_pWorksheet = NULL;
    }

    if (m_pWorksheets) {
        m_pWorksheets->Release();
        m_pWorksheets = NULL;
    }

    if (m_pWorkbook) {
        // 关闭工作簿（不保存）
        DISPID dispID;
        LPOLESTR ptName = const_cast<LPOLESTR>(L"Close");
        HRESULT hr = m_pWorkbook->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
        if (SUCCEEDED(hr)) {
            VARIANT saveChanges;
            VariantInit(&saveChanges);
            saveChanges.vt = VT_BOOL;
            saveChanges.boolVal = VARIANT_FALSE;

            VARIANT argsClose[1];
            argsClose[0] = saveChanges;

            DISPPARAMS dpClose;
            dpClose.cArgs = 1;
            dpClose.rgvarg = argsClose;
            dpClose.cNamedArgs = 0;
            dpClose.rgdispidNamedArgs = NULL;

            m_pWorkbook->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpClose, NULL, NULL, NULL);
        }

        m_pWorkbook->Release();
        m_pWorkbook = NULL;
    }

    if (m_pWorkbooks) {
        m_pWorkbooks->Release();
        m_pWorkbooks = NULL;
    }

    if (m_pExcelApp) {
        // 退出 Excel 应用程序
        DISPID dispID;
        LPOLESTR ptName = const_cast<LPOLESTR>(L"Quit");
        HRESULT hr = m_pExcelApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
        if (SUCCEEDED(hr)) {
            DISPPARAMS dpNoArgs = { NULL, NULL, 0, 0 };
            m_pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpNoArgs, NULL, NULL, NULL);
        }

        m_pExcelApp->Release();
        m_pExcelApp = NULL;
    }
}

/**
 * @brief 创建 Excel Application 实例
 */
bool ExcelSession::CreateExcelInstance() {
    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to create Excel application instance. HRESULT: " << hr << std::endl;
        return false;
    }

    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&m_pExcelApp);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to start Excel. HRESULT: " << hr << std::endl;
        return false;
    }
    return true;
}

/**
 * @brief 设置 Excel 不可见（失败不中断流程）
 */
void ExcelSession::SetExcelInvisible() {
    VARIANT visible;
    VariantInit(&visible);
    visible.vt = VT_BOOL;
    visible.boolVal = VARIANT_FALSE;
    ExcelComHelper::SetPropertyNoFail(m_pExcelApp, L"Visible", visible);
}

/**
 * @brief 获取 Workbooks 集合
 */
bool ExcelSession::GetWorkbooksCollection() {
    VARIANT result;
    VariantInit(&result);
    if (!ExcelComHelper::GetProperty(m_pExcelApp, L"Workbooks", result)) {
        return false;
    }
    m_pWorkbooks = result.pdispVal;
    return true;
}

/**
 * @brief 打开工作簿文件
 */
bool ExcelSession::OpenWorkbookFile(const std::wstring& filePath) {
    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(L"Open");
    HRESULT hr = m_pWorkbooks->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Open method. HRESULT: " << hr << std::endl;
        return false;
    }

    VARIANT filename;
    VariantInit(&filename);
    filename.vt = VT_BSTR;
    filename.bstrVal = SysAllocString(filePath.c_str());

    VARIANT args[1];
    args[0] = filename;

    DISPPARAMS dpOpen;
    dpOpen.cArgs = 1;
    dpOpen.rgvarg = args;
    dpOpen.cNamedArgs = 0;
    dpOpen.rgdispidNamedArgs = NULL;

    VARIANT result;
    VariantInit(&result);
    hr = m_pWorkbooks->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpOpen, &result, NULL, NULL);
    SysFreeString(filename.bstrVal);

    if (FAILED(hr)) {
        std::wcerr << L"Failed to open file. HRESULT: " << hr << std::endl;
        return false;
    }
    m_pWorkbook = result.pdispVal;
    return true;
}

/**
 * @brief 获取 Worksheets 集合
 */
bool ExcelSession::GetWorksheetsCollection() {
    VARIANT result;
    VariantInit(&result);
    if (!ExcelComHelper::GetProperty(m_pWorkbook, L"Worksheets", result)) {
        return false;
    }
    m_pWorksheets = result.pdispVal;
    return true;
}

/**
 * @brief 获取第一个工作表
 */
bool ExcelSession::GetFirstWorksheet() {
    VARIANT sheetIndex;
    VariantInit(&sheetIndex);
    sheetIndex.vt = VT_I4;
    sheetIndex.lVal = 1;

    VARIANT args[1];
    args[0] = sheetIndex;

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(L"Item");
    HRESULT hr = m_pWorksheets->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Item method. HRESULT: " << hr << std::endl;
        return false;
    }

    DISPPARAMS dpItem;
    dpItem.cArgs = 1;
    dpItem.rgvarg = args;
    dpItem.cNamedArgs = 0;
    dpItem.rgdispidNamedArgs = NULL;

    VARIANT result;
    VariantInit(&result);
    hr = m_pWorksheets->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpItem, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get worksheet. HRESULT: " << hr << std::endl;
        return false;
    }
    m_pWorksheet = result.pdispVal;
    return true;
}

/**
 * @brief 获取 UsedRange
 */
bool ExcelSession::GetUsedRange() {
    VARIANT result;
    VariantInit(&result);
    if (!ExcelComHelper::GetProperty(m_pWorksheet, L"UsedRange", result)) {
        return false;
    }
    m_pRange = result.pdispVal;
    return true;
}

/**
 * @brief 获取单元格数据并处理 SAFEARRAY
 */
bool ExcelSession::GetCellDataAndSafeArray() {
    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(L"Value");
    HRESULT hr = m_pRange->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Value property. HRESULT: " << hr << std::endl;
        return false;
    }

    VariantInit(&m_varResult);
    DISPPARAMS dpNoArgs = { NULL, NULL, 0, 0 };
    hr = m_pRange->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &m_varResult, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get cell data. HRESULT: " << hr << std::endl;
        VariantClear(&m_varResult);
        return false;
    }

    if ((m_varResult.vt & VT_ARRAY) != VT_ARRAY) {
        std::wcerr << L"Cell data is not an array. Type: " << m_varResult.vt << std::endl;
        VariantClear(&m_varResult);
        return false;
    }

    m_pSafeArray = m_varResult.parray;
    SafeArrayGetLBound(m_pSafeArray, 1, &m_lBound);
    SafeArrayGetUBound(m_pSafeArray, 1, &m_uBound);
    return true;
}

/**
 * @brief 打开 Excel 文件
 * 
 * 创建 Excel Application 实例，打开指定的工作簿，
 * 并获取第一个工作表的 UsedRange 数据。
 * 
 * @param filePath Excel 文件路径
 * @return true-打开成功，false-打开失败
 */
bool ExcelSession::OpenFile(const std::wstring& filePath) {
    Release();

    if (!CreateExcelInstance()) {
        return false;
    }

    SetExcelInvisible();

    if (!GetWorkbooksCollection()) {
        return false;
    }

    if (!OpenWorkbookFile(filePath)) {
        return false;
    }

    if (!GetWorksheetsCollection()) {
        return false;
    }

    if (!GetFirstWorksheet()) {
        return false;
    }

    if (!GetUsedRange()) {
        return false;
    }

    if (!GetCellDataAndSafeArray()) {
        return false;
    }

    return true;
}

/**
 * @brief 获取单元格值
 * 
 * @param row 行号（从 1 开始）
 * @param col 列号（从 1 开始）
 * @param cellValue 输出参数，存储单元格值
 * @return true-获取成功，false-获取失败
 */
bool ExcelSession::GetCellValue(long row, long col, VARIANT& cellValue) const {
    if (!m_pSafeArray || row < m_lBound || row > m_uBound) {
        return false;
    }

    // SAFEARRAY 索引从 1 开始
    long indices[2] = { row, col };
    VariantInit(&cellValue);
    SafeArrayGetElement(m_pSafeArray, indices, &cellValue);
    return true;
}

/**
 * @brief 获取单元格值（字符串）
 * 
 * @param row 行号（从 1 开始）
 * @param col 列号（从 1 开始）
 * @param defaultVal 默认值
 * @return 单元格字符串值
 */
std::wstring ExcelSession::GetCellString(long row, long col, const std::wstring& defaultVal) const {
    VARIANT cellValue = { 0 };
    if (!GetCellValue(row, col, cellValue)) {
        return defaultVal;
    }

    std::wstring result = ExcelComHelper::VariantToString(cellValue, defaultVal);
    VariantClear(&cellValue);
    return result;
}

/**
 * @brief 获取单元格值（整数）
 * 
 * @param row 行号（从 1 开始）
 * @param col 列号（从 1 开始）
 * @param defaultVal 默认值
 * @return 单元格整数值
 */
long ExcelSession::GetCellLong(long row, long col, long defaultVal) const {
    VARIANT cellValue = { 0 };
    if (!GetCellValue(row, col, cellValue)) {
        return defaultVal;
    }

    long result = ExcelComHelper::VariantToLong(cellValue, defaultVal);
    VariantClear(&cellValue);
    return result;
}

/**
 * @brief 获取单元格值（浮点数）
 * 
 * @param row 行号（从 1 开始）
 * @param col 列号（从 1 开始）
 * @param defaultVal 默认值
 * @return 单元格浮点数值
 */
double ExcelSession::GetCellDouble(long row, long col, double defaultVal) const {
    VARIANT cellValue = { 0 };
    if (!GetCellValue(row, col, cellValue)) {
        return defaultVal;
    }

    double result = ExcelComHelper::VariantToDouble(cellValue, defaultVal);
    VariantClear(&cellValue);
    return result;
}

/**
 * @brief 获取时间单元格值
 * 
 * 处理三种时间格式：VT_BSTR（字符串）、VT_DATE（Variant时间）、VT_R8（浮点数）。
 * 
 * @param row 行号（从 1 开始）
 * @param col 列号（从 1 开始）
 * @return 格式化的时间字符串（HH:MM:SS）
 */
std::wstring ExcelSession::GetCellTime(long row, long col) const {
    VARIANT cellValue = { 0 };
    if (!GetCellValue(row, col, cellValue)) {
        return L"";
    }

    std::wstring result = ExcelComHelper::VariantToTime(cellValue);
    VariantClear(&cellValue);
    return result;
}
