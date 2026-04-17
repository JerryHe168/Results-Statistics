#pragma execution_character_set("utf-8")

#include "ExcelWriter.h"
#include <windows.h>
#include <comutil.h>
#include <iostream>
#include <algorithm>

#pragma comment(lib, "comsuppw.lib")

/**
 * @brief 构造函数
 */
ExcelWriter::ExcelWriter()
    : m_pExcelApp(NULL), m_pWorkbooks(NULL), m_pWorkbook(NULL),
      m_pWorksheets(NULL), m_pWorksheet(NULL) {
}

/**
 * @brief 析构函数
 */
ExcelWriter::~ExcelWriter() {
    Release();
}

/**
 * @brief 释放所有 COM 对象
 */
void ExcelWriter::Release() {
    // 逆序释放 COM 对象
    // 创建顺序：pExcelApp -> pWorkbooks -> pWorkbook -> pWorksheets -> pWorksheet
    // 释放顺序：pWorksheet -> pWorksheets -> pWorkbook -> pWorkbooks -> pExcelApp

    if (m_pWorksheet) {
        m_pWorksheet->Release();
        m_pWorksheet = NULL;
    }

    if (m_pWorksheets) {
        m_pWorksheets->Release();
        m_pWorksheets = NULL;
    }

    if (m_pWorkbook) {
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
 * @brief 创建新的 Excel 工作簿
 * 
 * 创建 Excel Application 实例，新建一个工作簿，
 * 并获取第一个工作表。
 * 
 * @return true-创建成功，false-创建失败
 */
bool ExcelWriter::CreateNewWorkbook() {
    // 先释放之前的资源
    Release();

    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to create Excel application instance" << std::endl;
        return false;
    }

    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&m_pExcelApp);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to start Excel" << std::endl;
        return false;
    }

    // 设置 Excel 不可见
    VARIANT visible;
    VariantInit(&visible);
    visible.vt = VT_BOOL;
    visible.boolVal = VARIANT_FALSE;

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(L"Visible");
    hr = m_pExcelApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (SUCCEEDED(hr)) {
        DISPPARAMS dp = { NULL, NULL, 0, 0 };
        dp.cArgs = 1;
        dp.rgvarg = &visible;
        m_pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dp, NULL, NULL, NULL);
    }

    // 获取 Workbooks 集合
    ptName = const_cast<LPOLESTR>(L"Workbooks");
    hr = m_pExcelApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Workbooks property" << std::endl;
        Release();
        return false;
    }

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dpNoArgs = { NULL, NULL, 0, 0 };
    hr = m_pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Workbooks collection" << std::endl;
        Release();
        return false;
    }
    m_pWorkbooks = result.pdispVal;

    // 创建新工作簿
    ptName = const_cast<LPOLESTR>(L"Add");
    hr = m_pWorkbooks->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Add method" << std::endl;
        Release();
        return false;
    }

    VariantInit(&result);
    hr = m_pWorkbooks->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to create new workbook" << std::endl;
        Release();
        return false;
    }
    m_pWorkbook = result.pdispVal;

    // 获取 Worksheets 集合
    ptName = const_cast<LPOLESTR>(L"Worksheets");
    hr = m_pWorkbook->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Worksheets property" << std::endl;
        Release();
        return false;
    }

    VariantInit(&result);
    hr = m_pWorkbook->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Worksheets collection" << std::endl;
        Release();
        return false;
    }
    m_pWorksheets = result.pdispVal;

    // 获取第一个工作表（Item(1)）
    VARIANT sheetIndex;
    VariantInit(&sheetIndex);
    sheetIndex.vt = VT_I4;
    sheetIndex.lVal = 1;

    ptName = const_cast<LPOLESTR>(L"Item");
    hr = m_pWorksheets->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Item method" << std::endl;
        Release();
        return false;
    }

    VARIANT argsItem[1];
    argsItem[0] = sheetIndex;

    DISPPARAMS dpItem;
    dpItem.cArgs = 1;
    dpItem.rgvarg = argsItem;
    dpItem.cNamedArgs = 0;
    dpItem.rgdispidNamedArgs = NULL;

    VariantInit(&result);
    hr = m_pWorksheets->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpItem, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get worksheet" << std::endl;
        Release();
        return false;
    }
    m_pWorksheet = result.pdispVal;

    return true;
}

/**
 * @brief 写入单元格值（字符串）
 * 
 * @param row 行号（从 1 开始）
 * @param col 列号（从 1 开始）
 * @param value 字符串值
 * @return true-写入成功，false-写入失败
 */
bool ExcelWriter::WriteCell(long row, long col, const std::wstring& value) {
    if (!m_pWorksheet) {
        return false;
    }

    VARIANT cellValue;
    VariantInit(&cellValue);
    cellValue.vt = VT_BSTR;
    cellValue.bstrVal = SysAllocString(value.c_str());

    // 获取 Cells 属性
    LPOLESTR cellsName = const_cast<LPOLESTR>(L"Cells");
    DISPID cellsDispID;
    HRESULT hr = m_pWorksheet->GetIDsOfNames(IID_NULL, &cellsName, 1, LOCALE_USER_DEFAULT, &cellsDispID);
    if (FAILED(hr)) {
        SysFreeString(cellValue.bstrVal);
        return false;
    }

    // 参数顺序：DISPPARAMS 的参数是逆序的
    // rangeArgs[1] = row
    // rangeArgs[0] = col
    VARIANT rangeArgs[2];
    rangeArgs[1].vt = VT_I4;
    rangeArgs[1].lVal = row;
    rangeArgs[0].vt = VT_I4;
    rangeArgs[0].lVal = col;

    DISPPARAMS dpRange;
    dpRange.cArgs = 2;
    dpRange.rgvarg = rangeArgs;
    dpRange.cNamedArgs = 0;
    dpRange.rgdispidNamedArgs = NULL;

    VARIANT varRange;
    VariantInit(&varRange);

    hr = m_pWorksheet->Invoke(cellsDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpRange, &varRange, NULL, NULL);
    if (FAILED(hr) || varRange.vt != VT_DISPATCH) {
        VariantClear(&varRange);
        SysFreeString(cellValue.bstrVal);
        return false;
    }

    IDispatch* pCell = varRange.pdispVal;

    // 设置 Value 属性
    LPOLESTR valueName = const_cast<LPOLESTR>(L"Value");
    DISPID valueDispID;
    hr = pCell->GetIDsOfNames(IID_NULL, &valueName, 1, LOCALE_USER_DEFAULT, &valueDispID);
    if (FAILED(hr)) {
        pCell->Release();
        VariantClear(&varRange);
        SysFreeString(cellValue.bstrVal);
        return false;
    }

    // 关键修复：DISPPARAMS 命名参数配置
    // 设置属性时需要 cNamedArgs=1 且 rgdispidNamedArgs=&DISPID_PROPERTYPUT
    DISPPARAMS dpValue;
    DISPID dispidPut = DISPID_PROPERTYPUT;
    dpValue.cArgs = 1;
    dpValue.rgvarg = &cellValue;
    dpValue.cNamedArgs = 1;
    dpValue.rgdispidNamedArgs = &dispidPut;

    VARIANT varResult;
    VariantInit(&varResult);

    hr = pCell->Invoke(valueDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dpValue, &varResult, NULL, NULL);

    VariantClear(&varResult);
    pCell->Release();
    VariantClear(&varRange);
    SysFreeString(cellValue.bstrVal);

    return SUCCEEDED(hr);
}

/**
 * @brief 写入单元格值（整数）
 * 
 * @param row 行号（从 1 开始）
 * @param col 列号（从 1 开始）
 * @param value 整数值
 * @return true-写入成功，false-写入失败
 */
bool ExcelWriter::WriteCell(long row, long col, int value) {
    if (!m_pWorksheet) {
        return false;
    }

    VARIANT cellValue;
    VariantInit(&cellValue);
    cellValue.vt = VT_I4;
    cellValue.lVal = value;

    // 获取 Cells 属性
    LPOLESTR cellsName = const_cast<LPOLESTR>(L"Cells");
    DISPID cellsDispID;
    HRESULT hr = m_pWorksheet->GetIDsOfNames(IID_NULL, &cellsName, 1, LOCALE_USER_DEFAULT, &cellsDispID);
    if (FAILED(hr)) {
        return false;
    }

    // 参数顺序：DISPPARAMS 的参数是逆序的
    VARIANT rangeArgs[2];
    rangeArgs[1].vt = VT_I4;
    rangeArgs[1].lVal = row;
    rangeArgs[0].vt = VT_I4;
    rangeArgs[0].lVal = col;

    DISPPARAMS dpRange;
    dpRange.cArgs = 2;
    dpRange.rgvarg = rangeArgs;
    dpRange.cNamedArgs = 0;
    dpRange.rgdispidNamedArgs = NULL;

    VARIANT varRange;
    VariantInit(&varRange);

    hr = m_pWorksheet->Invoke(cellsDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpRange, &varRange, NULL, NULL);
    if (FAILED(hr) || varRange.vt != VT_DISPATCH) {
        VariantClear(&varRange);
        return false;
    }

    IDispatch* pCell = varRange.pdispVal;

    // 设置 Value 属性
    LPOLESTR valueName = const_cast<LPOLESTR>(L"Value");
    DISPID valueDispID;
    hr = pCell->GetIDsOfNames(IID_NULL, &valueName, 1, LOCALE_USER_DEFAULT, &valueDispID);
    if (FAILED(hr)) {
        pCell->Release();
        VariantClear(&varRange);
        return false;
    }

    // 关键修复：DISPPARAMS 命名参数配置
    DISPPARAMS dpValue;
    DISPID dispidPut = DISPID_PROPERTYPUT;
    dpValue.cArgs = 1;
    dpValue.rgvarg = &cellValue;
    dpValue.cNamedArgs = 1;
    dpValue.rgdispidNamedArgs = &dispidPut;

    VARIANT varResult;
    VariantInit(&varResult);

    hr = pCell->Invoke(valueDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dpValue, &varResult, NULL, NULL);

    VariantClear(&varResult);
    pCell->Release();
    VariantClear(&varRange);

    return SUCCEEDED(hr);
}

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
bool ExcelWriter::SaveAndClose(const std::wstring& filePath) {
    if (!m_pWorkbook || !m_pWorksheet) {
        return false;
    }

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(L"SaveAs");
    HRESULT hr = m_pWorkbook->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get SaveAs method" << std::endl;
        Release();
        return false;
    }

    VARIANT filename;
    VariantInit(&filename);
    filename.vt = VT_BSTR;
    filename.bstrVal = SysAllocString(filePath.c_str());

    // Excel 文件格式代码（XlFileFormat 枚举）：
    // 51 = xlOpenXMLWorkbook (Excel 2007+ .xlsx 格式)
    // 56 = xlExcel8 (Excel 97-2003 .xls 格式)
    int fileFormatValue = 56;
    std::wstring lowerPath = filePath;
    std::transform(lowerPath.begin(), lowerPath.end(), lowerPath.begin(), ::towlower);
    
    // 根据文件扩展名选择格式
    if (lowerPath.length() >= 5 && lowerPath.substr(lowerPath.length() - 5) == L".xlsx") {
        fileFormatValue = 51;  // .xlsx 格式
    }
    else if (lowerPath.length() >= 4 && lowerPath.substr(lowerPath.length() - 4) == L".xls") {
        fileFormatValue = 56;  // .xls 格式
    }

    VARIANT fileFormat;
    VariantInit(&fileFormat);
    fileFormat.vt = VT_I4;
    fileFormat.lVal = fileFormatValue;

    // SaveAs 方法参数：DISPPARAMS 的参数顺序是逆序的（栈结构）
    // argsSave[1] = 文件名
    // argsSave[0] = 格式代码
    VARIANT argsSave[2];
    argsSave[1] = filename;
    argsSave[0] = fileFormat;

    DISPPARAMS dpSave;
    dpSave.cArgs = 2;
    dpSave.rgvarg = argsSave;
    dpSave.cNamedArgs = 0;
    dpSave.rgdispidNamedArgs = NULL;

    hr = m_pWorkbook->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpSave, NULL, NULL, NULL);
    SysFreeString(filename.bstrVal);

    if (FAILED(hr)) {
        std::wcerr << L"Failed to save file. HRESULT: " << hr << std::endl;
        Release();
        return false;
    }

    // 关闭工作簿
    ptName = const_cast<LPOLESTR>(L"Close");
    hr = m_pWorkbook->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
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

    Release();
    return true;
}
