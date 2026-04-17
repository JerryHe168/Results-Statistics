#pragma execution_character_set("utf-8")

/**
 * @class ExcelReader
 * @brief Excel文件读取器类
 * 
 * 负责使用COM自动化技术读取Excel格式的报名信息和成绩清单文件。
 */

#include "ExcelReader.h"
#include <windows.h>
#include <comutil.h>
#include <iostream>
#include <sstream>
#include <regex>
#include <algorithm>

#pragma comment(lib, "comsuppw.lib")

/**
 * @brief 构造函数
 */
ExcelReader::ExcelReader() {
    CoInitialize(NULL);
}

/**
 * @brief 析构函数
 */
ExcelReader::~ExcelReader() {
    CoUninitialize();
}

/**
 * @brief 从编号中提取组号
 * 
 * 使用正则表达式匹配字符串中的第一个连续数字序列。
 * 
 * @param id 编号字符串
 * @return 提取的组号，无法提取则返回-1
 */
int ExcelReader::ExtractGroupNumber(const std::wstring& id) {
    // 正则表达式 L"(\\d+)"：匹配编号中的连续数字
    // 例如："23A" → 23，"17B" → 17
    std::wregex regex(L"(\\d+)");
    std::wsmatch match;

    if (std::regex_search(id, match, regex)) {
        // match[1] 是第一个捕获组的内容
        return std::stoi(match[1].str());
    }

    return -1;
}

/**
 * @brief 读取报名信息Excel文件
 * 
 * 使用COM自动化技术，解析男生编号、男生姓名、女生编号、女生姓名。
 * 
 * @param filePath Excel文件路径
 * @param participants 输出参数，存储读取到的报名信息列表
 * @return true-读取成功，false-读取失败
 */
bool ExcelReader::ReadRegistrationInfo(const std::wstring& filePath, std::vector<Participant>& participants) {
    IDispatch* pExcelApp = NULL;
    IDispatch* pWorkbooks = NULL;
    IDispatch* pWorkbook = NULL;
    IDispatch* pWorksheets = NULL;
    IDispatch* pWorksheet = NULL;
    IDispatch* pRange = NULL;

    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to create Excel application instance. HRESULT: " << hr << std::endl;
        return false;
    }

    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pExcelApp);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to start Excel. HRESULT: " << hr << std::endl;
        return false;
    }

    VARIANT visible;
    VariantInit(&visible);
    visible.vt = VT_BOOL;
    visible.boolVal = VARIANT_FALSE;

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(L"Visible");
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (SUCCEEDED(hr)) {
        DISPPARAMS dp = { NULL, NULL, 0, 0 };
        dp.cArgs = 1;
        dp.rgvarg = &visible;
        pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dp, NULL, NULL, NULL);
    }

    ptName = const_cast<LPOLESTR>(L"Workbooks");
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Workbooks property. HRESULT: " << hr << std::endl;
        pExcelApp->Release();
        return false;
    }

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dpNoArgs = { NULL, NULL, 0, 0 };
    hr = pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Workbooks collection. HRESULT: " << hr << std::endl;
        pExcelApp->Release();
        return false;
    }
    pWorkbooks = result.pdispVal;

    ptName = const_cast<LPOLESTR>(L"Open");
    hr = pWorkbooks->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Open method. HRESULT: " << hr << std::endl;
        pWorkbooks->Release();
        pExcelApp->Release();
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

    VariantInit(&result);
    hr = pWorkbooks->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpOpen, &result, NULL, NULL);
    SysFreeString(filename.bstrVal);

    if (FAILED(hr)) {
        std::wcerr << L"Failed to open file. HRESULT: " << hr << std::endl;
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }
    pWorkbook = result.pdispVal;

    ptName = const_cast<LPOLESTR>(L"Worksheets");
    hr = pWorkbook->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Worksheets property. HRESULT: " << hr << std::endl;
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    VariantInit(&result);
    hr = pWorkbook->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Worksheets collection. HRESULT: " << hr << std::endl;
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }
    pWorksheets = result.pdispVal;

    VARIANT sheetIndex;
    VariantInit(&sheetIndex);
    sheetIndex.vt = VT_I4;
    sheetIndex.lVal = 1;

    ptName = const_cast<LPOLESTR>(L"Item");
    hr = pWorksheets->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Item method. HRESULT: " << hr << std::endl;
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
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
    hr = pWorksheets->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpItem, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get worksheet. HRESULT: " << hr << std::endl;
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }
    pWorksheet = result.pdispVal;

    ptName = const_cast<LPOLESTR>(L"UsedRange");
    hr = pWorksheet->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get UsedRange property. HRESULT: " << hr << std::endl;
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    VariantInit(&result);
    hr = pWorksheet->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get UsedRange. HRESULT: " << hr << std::endl;
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }
    pRange = result.pdispVal;

    ptName = const_cast<LPOLESTR>(L"Value");
    hr = pRange->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Value property. HRESULT: " << hr << std::endl;
        pRange->Release();
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    VARIANT varResult;
    VariantInit(&varResult);
    hr = pRange->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &varResult, NULL, NULL);
    
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get cell data. HRESULT: " << hr << std::endl;
        VariantClear(&varResult);
        pRange->Release();
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    // 关键修复：使用位运算检查VT_ARRAY类型
    // Excel返回 VT_ARRAY | VT_VARIANT (8204 = 0x200C)
    if ((varResult.vt & VT_ARRAY) != VT_ARRAY) {
        std::wcerr << L"Cell data is not an array. Type: " << varResult.vt << std::endl;
        VariantClear(&varResult);
        pRange->Release();
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    // 获取 SAFEARRAY 指针
    // SAFEARRAY 是 COM 中用于跨进程传递数组的安全类型
    // 注意：Excel 的 SAFEARRAY 索引从 1 开始（不是从 0 开始）
    SAFEARRAY* pSafeArray = varResult.parray;
    long lBound, uBound;
    SafeArrayGetLBound(pSafeArray, 1, &lBound);  // 获取下界（通常是1）
    SafeArrayGetUBound(pSafeArray, 1, &uBound);  // 获取上界

    // 遍历数据行
    // 从 lBound + 1 开始是为了跳过表头行
    // 报名信息的列结构：
    // 列 1：男生编号（如 "23A", "18A"）
    // 列 2：男生姓名
    // 列 3：女生编号（如 "16B", "13B"）
    // 列 4：女生姓名
    for (long row = lBound + 1; row <= uBound; row++) {
        Participant participant;
        VARIANT cellValue;
        
        // indices 数组：{行号, 列号}
        // 注意：SAFEARRAY 的索引顺序是 {行, 列}
        long indices[2] = { row, 1 };

        // 读取第 1 列：男生编号
        // VARIANT 类型处理：
        // - VT_BSTR: 字符串类型（最常见）
        // - VT_I4: 32位整数
        // - VT_R8: 双精度浮点数
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_BSTR) {
            participant.maleId = cellValue.bstrVal;
        }
        else if (cellValue.vt == VT_I4) {
            participant.maleId = std::to_wstring(cellValue.lVal);
        }
        else if (cellValue.vt == VT_R8) {
            participant.maleId = std::to_wstring((long long)cellValue.dblVal);
        }
        VariantClear(&cellValue);

        // 读取第 2 列：男生姓名
        indices[1] = 2;
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_BSTR) {
            participant.maleName = cellValue.bstrVal;
        }
        VariantClear(&cellValue);

        // 读取第 3 列：女生编号
        indices[1] = 3;
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_BSTR) {
            participant.femaleId = cellValue.bstrVal;
        }
        else if (cellValue.vt == VT_I4) {
            participant.femaleId = std::to_wstring(cellValue.lVal);
        }
        else if (cellValue.vt == VT_R8) {
            participant.femaleId = std::to_wstring((long long)cellValue.dblVal);
        }
        VariantClear(&cellValue);

        // 读取第 4 列：女生姓名
        indices[1] = 4;
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_BSTR) {
            participant.femaleName = cellValue.bstrVal;
        }
        VariantClear(&cellValue);

        // 从编号中提取组号
        participant.maleGroupNumber = ExtractGroupNumber(participant.maleId);
        participant.femaleGroupNumber = ExtractGroupNumber(participant.femaleId);

        // 只添加有姓名的记录
        if (!participant.maleName.empty() || !participant.femaleName.empty()) {
            participants.push_back(participant);
        }
    }

    // 逆序释放COM对象
    // 创建顺序：pExcelApp -> pWorkbooks -> pWorkbook -> pWorksheets -> pWorksheet -> pRange
    // 释放顺序：pRange -> pWorksheet -> pWorksheets -> pWorkbook -> pWorkbooks -> pExcelApp
    VariantClear(&varResult);
    pRange->Release();
    pWorksheet->Release();
    pWorksheets->Release();

    ptName = const_cast<LPOLESTR>(L"Close");
    hr = pWorkbook->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
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

        pWorkbook->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpClose, NULL, NULL, NULL);
    }

    pWorkbook->Release();
    pWorkbooks->Release();

    ptName = const_cast<LPOLESTR>(L"Quit");
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (SUCCEEDED(hr)) {
        pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpNoArgs, NULL, NULL, NULL);
    }

    pExcelApp->Release();

    return true;
}

/**
 * @brief 读取成绩清单Excel文件
 * 
 * 使用COM自动化技术，解析名次、组别、成绩时间。
 * 
 * @param filePath Excel文件路径
 * @param scoreEntries 输出参数，存储读取到的成绩条目列表
 * @return true-读取成功，false-读取失败
 */
bool ExcelReader::ReadScoreList(const std::wstring& filePath, std::vector<ScoreEntry>& scoreEntries) {
    IDispatch* pExcelApp = NULL;
    IDispatch* pWorkbooks = NULL;
    IDispatch* pWorkbook = NULL;
    IDispatch* pWorksheets = NULL;
    IDispatch* pWorksheet = NULL;
    IDispatch* pRange = NULL;

    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to create Excel application instance. HRESULT: " << hr << std::endl;
        return false;
    }

    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pExcelApp);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to start Excel. HRESULT: " << hr << std::endl;
        return false;
    }

    VARIANT visible;
    VariantInit(&visible);
    visible.vt = VT_BOOL;
    visible.boolVal = VARIANT_FALSE;

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(L"Visible");
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (SUCCEEDED(hr)) {
        DISPPARAMS dp = { NULL, NULL, 0, 0 };
        dp.cArgs = 1;
        dp.rgvarg = &visible;
        pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dp, NULL, NULL, NULL);
    }

    ptName = const_cast<LPOLESTR>(L"Workbooks");
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Workbooks property. HRESULT: " << hr << std::endl;
        pExcelApp->Release();
        return false;
    }

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dpNoArgs = { NULL, NULL, 0, 0 };
    hr = pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Workbooks collection. HRESULT: " << hr << std::endl;
        pExcelApp->Release();
        return false;
    }
    pWorkbooks = result.pdispVal;

    ptName = const_cast<LPOLESTR>(L"Open");
    hr = pWorkbooks->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Open method. HRESULT: " << hr << std::endl;
        pWorkbooks->Release();
        pExcelApp->Release();
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

    VariantInit(&result);
    hr = pWorkbooks->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpOpen, &result, NULL, NULL);
    SysFreeString(filename.bstrVal);

    if (FAILED(hr)) {
        std::wcerr << L"Failed to open file. HRESULT: " << hr << std::endl;
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }
    pWorkbook = result.pdispVal;

    ptName = const_cast<LPOLESTR>(L"Worksheets");
    hr = pWorkbook->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Worksheets property. HRESULT: " << hr << std::endl;
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    VariantInit(&result);
    hr = pWorkbook->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Worksheets collection. HRESULT: " << hr << std::endl;
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }
    pWorksheets = result.pdispVal;

    VARIANT sheetIndex;
    VariantInit(&sheetIndex);
    sheetIndex.vt = VT_I4;
    sheetIndex.lVal = 1;

    ptName = const_cast<LPOLESTR>(L"Item");
    hr = pWorksheets->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Item method. HRESULT: " << hr << std::endl;
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
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
    hr = pWorksheets->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpItem, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get worksheet. HRESULT: " << hr << std::endl;
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }
    pWorksheet = result.pdispVal;

    ptName = const_cast<LPOLESTR>(L"UsedRange");
    hr = pWorksheet->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get UsedRange property. HRESULT: " << hr << std::endl;
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    VariantInit(&result);
    hr = pWorksheet->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get UsedRange. HRESULT: " << hr << std::endl;
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }
    pRange = result.pdispVal;

    ptName = const_cast<LPOLESTR>(L"Value");
    hr = pRange->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Value property. HRESULT: " << hr << std::endl;
        pRange->Release();
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    VARIANT varResult;
    VariantInit(&varResult);
    hr = pRange->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &varResult, NULL, NULL);
    
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get cell data. HRESULT: " << hr << std::endl;
        VariantClear(&varResult);
        pRange->Release();
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    if ((varResult.vt & VT_ARRAY) != VT_ARRAY) {
        std::wcerr << L"Cell data is not an array. Type: " << varResult.vt << std::endl;
        VariantClear(&varResult);
        pRange->Release();
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    // 获取 SAFEARRAY 指针
    // 成绩清单的列结构：
    // 列 1：名次
    // 列 2：组别（如 "23组"）
    // 列 3：成绩时间
    SAFEARRAY* pSafeArray = varResult.parray;
    long lBound, uBound;
    SafeArrayGetLBound(pSafeArray, 1, &lBound);  // 获取下界（通常是1）
    SafeArrayGetUBound(pSafeArray, 1, &uBound);  // 获取上界

    // 遍历数据行
    // 从 lBound + 1 开始是为了跳过表头行
    for (long row = lBound + 1; row <= uBound; row++) {
        ScoreEntry entry;
        VARIANT cellValue;
        long indices[2] = { row, 1 };

        // 读取第 1 列：名次
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_I4) {
            entry.rank = cellValue.lVal;
        }
        else if (cellValue.vt == VT_R8) {
            entry.rank = (long)cellValue.dblVal;
        }
        else if (cellValue.vt == VT_BSTR) {
            // 字符串格式的名次，尝试转换为整数
            try {
                entry.rank = std::stoi(cellValue.bstrVal);
            }
            catch (...) {
                entry.rank = 0;  // 转换失败，标记为无效
            }
        }
        VariantClear(&cellValue);

        // 读取第 2 列：组别
        indices[1] = 2;
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_BSTR) {
            entry.group = cellValue.bstrVal;
        }
        else if (cellValue.vt == VT_I4) {
            // 数字格式的组别，添加 "组" 后缀
            entry.group = std::to_wstring(cellValue.lVal) + L"组";
        }
        else if (cellValue.vt == VT_R8) {
            entry.group = std::to_wstring((long)cellValue.dblVal) + L"组";
        }
        VariantClear(&cellValue);

        // 时间格式处理：
        // Excel 中的时间可能以三种形式存储：
        // - VT_BSTR: 字符串格式（如 "0:37:06"）
        // - VT_DATE: Variant 时间格式（使用 VariantTimeToSystemTime 转换）
        // - VT_R8: 浮点数格式（0.0 = 0:00:00, 1.0 = 24:00:00）
        //   例如：0.025 表示 0.6 小时 = 36 分钟
        //
        // 浮点数转时间公式：
        // - 小时 = timeVal * 24  （一天24小时）
        // - 分钟 = (小数部分) * 60  （一小时60分钟）
        // - 秒 = (小数部分) * 60    （一分钟60秒）
        indices[1] = 3;
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_BSTR) {
            entry.time = cellValue.bstrVal;
        }
        else if (cellValue.vt == VT_DATE) {
            SYSTEMTIME st;
            VariantTimeToSystemTime(cellValue.date, &st);
            wchar_t buffer[32];
            // 格式：小时:分钟:秒（分钟和秒补零）
            swprintf_s(buffer, L"%d:%02d:%02d", st.wHour, st.wMinute, st.wSecond);
            entry.time = buffer;
        }
        else if (cellValue.vt == VT_R8) {
            double timeVal = cellValue.dblVal;
            // 浮点数转时间：0.0 = 0:00:00, 1.0 = 24:00:00
            int hours = (int)(timeVal * 24);  // 乘以24得到小时数
            int minutes = (int)((timeVal * 24 - hours) * 60);  // 小数部分乘以60得到分钟
            int seconds = (int)(((timeVal * 24 - hours) * 60 - minutes) * 60);  // 小数部分乘以60得到秒
            wchar_t buffer[32];
            swprintf_s(buffer, L"%d:%02d:%02d", hours, minutes, seconds);
            entry.time = buffer;
        }
        VariantClear(&cellValue);

        entry.groupNumber = ExtractGroupNumber(entry.group);

        if (entry.rank > 0) {
            scoreEntries.push_back(entry);
        }
    }

    VariantClear(&varResult);
    pRange->Release();
    pWorksheet->Release();
    pWorksheets->Release();

    ptName = const_cast<LPOLESTR>(L"Close");
    hr = pWorkbook->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
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

        pWorkbook->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpClose, NULL, NULL, NULL);
    }

    pWorkbook->Release();
    pWorkbooks->Release();

    ptName = const_cast<LPOLESTR>(L"Quit");
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (SUCCEEDED(hr)) {
        pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpNoArgs, NULL, NULL, NULL);
    }

    pExcelApp->Release();

    return true;
}
