#include "stdafx.h"
#pragma execution_character_set("utf-8")

#include "ExcelWriter.h"
#include "ExcelComHelper.h"
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
    ExcelComHelper::SafeRelease(m_pWorksheet);
    ExcelComHelper::SafeRelease(m_pWorksheets);
    ExcelComHelper::SafeRelease(m_pWorkbook);
    ExcelComHelper::SafeRelease(m_pWorkbooks);

    if (m_pExcelApp) {
        ExcelComHelper::InvokeMethod(m_pExcelApp, L"Quit");
        ExcelComHelper::SafeRelease(m_pExcelApp);
    }
}

/**
 * @brief 写入单元格值（内部通用实现）
 */
bool ExcelWriter::WriteCellInternal(long row, long col, const VARIANT& value) {
    if (!m_pWorksheet) {
        return false;
    }

    IDispatch* pCells = ExcelComHelper::GetPropertyDispatch(m_pWorksheet, L"Cells");
    if (!pCells) {
        return false;
    }

    IDispatch* pCell = ExcelComHelper::GetItem(pCells, row, col);
    ExcelComHelper::SafeRelease(pCells);

    if (!pCell) {
        return false;
    }

    bool success = ExcelComHelper::SetProperty(pCell, L"Value", value);
    ExcelComHelper::SafeRelease(pCell);

    return success;
}

/**
 * @brief 创建新的 Excel 工作簿
 */
bool ExcelWriter::CreateNewWorkbook() {
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

    VARIANT visible;
    VariantInit(&visible);
    visible.vt = VT_BOOL;
    visible.boolVal = VARIANT_FALSE;
    ExcelComHelper::SetPropertyNoFail(m_pExcelApp, L"Visible", visible);

    m_pWorkbooks = ExcelComHelper::GetPropertyDispatch(m_pExcelApp, L"Workbooks");
    if (!m_pWorkbooks) {
        std::wcerr << L"Failed to get Workbooks collection" << std::endl;
        Release();
        return false;
    }

    VARIANT result;
    VariantInit(&result);
    if (!ExcelComHelper::InvokeMethod(m_pWorkbooks, L"Add", &result)) {
        std::wcerr << L"Failed to create new workbook" << std::endl;
        Release();
        return false;
    }

    if (result.vt != VT_DISPATCH) {
        VariantClear(&result);
        Release();
        return false;
    }
    m_pWorkbook = result.pdispVal;

    m_pWorksheets = ExcelComHelper::GetPropertyDispatch(m_pWorkbook, L"Worksheets");
    if (!m_pWorksheets) {
        std::wcerr << L"Failed to get Worksheets collection" << std::endl;
        Release();
        return false;
    }

    m_pWorksheet = ExcelComHelper::GetItem(m_pWorksheets, 1);
    if (!m_pWorksheet) {
        std::wcerr << L"Failed to get worksheet" << std::endl;
        Release();
        return false;
    }

    return true;
}

/**
 * @brief 写入单元格值（字符串）
 */
bool ExcelWriter::WriteCell(long row, long col, const std::wstring& value) {
    VARIANT cellValue;
    VariantInit(&cellValue);
    cellValue.vt = VT_BSTR;
    cellValue.bstrVal = SysAllocString(value.c_str());

    bool success = WriteCellInternal(row, col, cellValue);

    SysFreeString(cellValue.bstrVal);
    return success;
}

/**
 * @brief 写入单元格值（整数）
 */
bool ExcelWriter::WriteCell(long row, long col, int value) {
    VARIANT cellValue;
    VariantInit(&cellValue);
    cellValue.vt = VT_I4;
    cellValue.lVal = value;

    return WriteCellInternal(row, col, cellValue);
}

/**
 * @brief 保存并关闭文件
 */
bool ExcelWriter::SaveAndClose(const std::wstring& filePath) {
    if (!m_pWorkbook || !m_pWorksheet) {
        return false;
    }

    int fileFormatValue = 56;
    std::wstring lowerPath = filePath;
    std::transform(lowerPath.begin(), lowerPath.end(), lowerPath.begin(), ::towlower);
    
    if (lowerPath.length() >= 5 && lowerPath.substr(lowerPath.length() - 5) == L".xlsx") {
        fileFormatValue = 51;
    }
    else if (lowerPath.length() >= 4 && lowerPath.substr(lowerPath.length() - 4) == L".xls") {
        fileFormatValue = 56;
    }

    VARIANT filename;
    VariantInit(&filename);
    filename.vt = VT_BSTR;
    filename.bstrVal = SysAllocString(filePath.c_str());

    VARIANT fileFormat;
    VariantInit(&fileFormat);
    fileFormat.vt = VT_I4;
    fileFormat.lVal = fileFormatValue;

    bool success = ExcelComHelper::InvokeMethod(m_pWorkbook, L"SaveAs", filename, fileFormat, NULL);
    SysFreeString(filename.bstrVal);

    if (!success) {
        std::wcerr << L"Failed to save file" << std::endl;
        Release();
        return false;
    }

    VARIANT saveChanges;
    VariantInit(&saveChanges);
    saveChanges.vt = VT_BOOL;
    saveChanges.boolVal = VARIANT_FALSE;
    ExcelComHelper::InvokeMethod(m_pWorkbook, L"Close", saveChanges);

    Release();
    return true;
}
