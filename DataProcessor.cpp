#pragma execution_character_set("utf-8")

// ============================================================
// DataProcessor.cpp - 数据处理器实现
// ============================================================
// 本文件实现了数据匹配和结果导出功能
// 
// 主要功能：
// 1. ProcessData：根据组别匹配男生和女生姓名
// 2. ExportResults：导出结果到Excel文件
// 3. ExportResultsToCsv：导出结果到CSV文件
// 
// 核心算法说明：
// - 建立两个独立的映射表：男生组号->姓名，女生组号->姓名
// - 遍历成绩条目，根据组号分别查找男生和女生姓名
// - 拼接姓名后生成结果
// 
// 为什么要使用两个独立的映射表？
// - 原来的逻辑假设每一行的男生和女生是同一组的
// - 但实际数据中，同一组的男生和女生可能在报名信息的不同行
// - 例如：成绩清单23组：罗晓东 梁馨尹
//   - 罗晓东（男生23A）在第3行
//   - 梁馨尹（女生23B）在第5行
// - 使用两个独立的映射表可以正确处理这种情况
// ============================================================

#include "DataProcessor.h"
#include <windows.h>
#include <comutil.h>
#include <iostream>
#include <fstream>
#include <sstream>
#include <iomanip>
#include <algorithm>

#pragma comment(lib, "comsuppw.lib")

// ============================================================
// 构造函数
// ============================================================
// 功能：初始化数据处理器
// 注意：本类不持有任何资源，构造函数和析构函数为空
// ============================================================
DataProcessor::DataProcessor() {
}

// ============================================================
// 析构函数
// ============================================================
DataProcessor::~DataProcessor() {
}

// ============================================================
// 数据匹配处理
// ============================================================
// 参数：
//   participants - 报名信息列表（输入）
//   scoreEntries - 成绩条目列表（输入）
//   results      - 结果列表（输出）
// 
// 返回值：
//   true  - 处理成功
//   false - 处理失败（当前总是返回true）
// 
// 执行流程：
//   1. 清空结果列表
//   2. 检查输入数据是否为空（输出警告但不中断）
//   3. 建立男生组号到姓名的映射表
//   4. 建立女生组号到姓名的映射表
//   5. 遍历每个成绩条目：
//      a. 复制名次、组别、成绩时间
//      b. 根据组号查找男生姓名
//      c. 根据组号查找女生姓名
//      d. 拼接姓名（处理各种情况：都有、只有男生、只有女生、都没有）
//      e. 添加到结果列表
// 
// 映射表建立规则：
//   - 只添加组号 >= 0 且姓名非空的条目
//   - 如果有重复组号，后添加的会覆盖先添加的
//   - 这符合预期：后面的行覆盖前面的行
// 
// 姓名拼接策略：
//   - 男生姓名 + 空格 + 女生姓名（两者都有）
//   - 只有男生姓名或只有女生姓名（其中一个为空）
//   - "Unknown"（两者都为空，输出警告）
//   - "Invalid Group"（组号无效）
// ============================================================
bool DataProcessor::ProcessData(const std::vector<Participant>& participants,
                                 const std::vector<ScoreEntry>& scoreEntries,
                                 std::vector<ResultEntry>& results) {
    results.clear();

    if (participants.empty()) {
        std::wcerr << L"Warning: No participants data available" << std::endl;
    }

    if (scoreEntries.empty()) {
        std::wcerr << L"Warning: No score entries available" << std::endl;
    }

    // 建立两个独立的映射表
    // maleMap：男生组号 -> 男生姓名
    // femaleMap：女生组号 -> 女生姓名
    // 使用 unordered_map 是因为哈希表的平均查找时间复杂度为 O(1)
    std::unordered_map<int, std::wstring> maleMap;
    std::unordered_map<int, std::wstring> femaleMap;

    // 遍历所有报名信息，填充映射表
    for (const auto& participant : participants) {
        // 填充男生映射表：组号有效且姓名非空
        if (participant.maleGroupNumber >= 0 && !participant.maleName.empty()) {
            maleMap[participant.maleGroupNumber] = participant.maleName;
        }

        // 填充女生映射表：组号有效且姓名非空
        if (participant.femaleGroupNumber >= 0 && !participant.femaleName.empty()) {
            femaleMap[participant.femaleGroupNumber] = participant.femaleName;
        }
    }

    // 遍历成绩条目，匹配姓名
    for (const auto& scoreEntry : scoreEntries) {
        ResultEntry result;

        // 复制基本信息
        result.rank = scoreEntry.rank;        // 名次
        result.group = scoreEntry.group;       // 组别（原始字符串）
        result.time = scoreEntry.time;         // 成绩时间

        // 处理无效组号的情况
        // groupNumber < 0 表示无法从组别字符串中提取组号
        if (scoreEntry.groupNumber < 0) {
            std::wcerr << L"Warning: Invalid group number for rank " << scoreEntry.rank << std::endl;
            result.names = L"Invalid Group";
        }
        else {
            std::wstring maleName;    // 男生姓名
            std::wstring femaleName;  // 女生姓名

            // 查找男生姓名
            // 使用 find() 而不是 operator[]，因为 operator[] 在键不存在时会插入默认值
            auto maleIt = maleMap.find(scoreEntry.groupNumber);
            if (maleIt != maleMap.end()) {
                maleName = maleIt->second;
            }

            // 查找女生姓名
            auto femaleIt = femaleMap.find(scoreEntry.groupNumber);
            if (femaleIt != femaleMap.end()) {
                femaleName = femaleIt->second;
            }

            // 处理各种姓名组合情况
            if (maleName.empty() && femaleName.empty()) {
                // 两者都为空：组别不存在于报名信息中
                std::wcerr << L"Warning: Participant info not found for group number " << scoreEntry.groupNumber << std::endl;
                result.names = L"Unknown";
            }
            else if (maleName.empty()) {
                // 只有女生姓名
                result.names = femaleName;
            }
            else if (femaleName.empty()) {
                // 只有男生姓名
                result.names = maleName;
            }
            else {
                // 两者都有，用空格拼接
                result.names = maleName + L" " + femaleName;
            }
        }

        results.push_back(result);
    }

    return true;
}

// ============================================================
// 导出结果到Excel文件
// ============================================================
// 参数：
//   filePath - 输出文件路径
//   results  - 结果列表
// 
// 返回值：
//   true  - 导出成功
//   false - 导出失败
// 
// 执行流程：
//   1. 检查结果是否为空（输出警告但继续）
//   2. 创建Excel.Application实例
//   3. 设置Excel不可见（后台运行）
//   4. 创建新工作簿（Workbooks.Add）
//   5. 获取第一个工作表
//   6. 定义lambda表达式简化单元格写入逻辑
//   7. 写入表头（Rank, Group, Names, Score）
//   8. 遍历结果列表，写入每一行数据
//   9. 根据文件扩展名选择保存格式
//   10. 保存文件（SaveAs）
//   11. 关闭工作簿，退出Excel
//   12. 释放所有COM对象
// 
// 文件格式选择：
//   - .xlsx：Excel 2007+ 格式，格式代码 51 (xlOpenXMLWorkbook)
//   - .xls：Excel 97-2003 格式，格式代码 56 (xlExcel8)
//   - 默认：使用 .xls 格式（代码 56）
// 
// 关键修复点：
//   - 必须正确设置DISPPARAMS的命名参数
//   - cNamedArgs = 1
//   - rgdispidNamedArgs = &DISPID_PROPERTYPUT
//   - 否则会返回 DISP_E_PARAMNOTFOUND (0x80020004) 错误
// ============================================================
bool DataProcessor::ExportResults(const std::wstring& filePath, const std::vector<ResultEntry>& results) {
    if (results.empty()) {
        std::wcerr << L"Warning: No results to export" << std::endl;
    }

    // 声明COM接口指针
    IDispatch* pExcelApp = NULL;      // Excel.Application对象
    IDispatch* pWorkbooks = NULL;      // Workbooks集合
    IDispatch* pWorkbook = NULL;       // 新建的Workbook
    IDispatch* pWorksheets = NULL;     // Worksheets集合
    IDispatch* pWorksheet = NULL;      // 第一个Worksheet
    IDispatch* pRange = NULL;          // 临时Range对象

    // 步骤1：创建Excel.Application实例
    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to create Excel application instance" << std::endl;
        return false;
    }

    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pExcelApp);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to start Excel" << std::endl;
        return false;
    }

    // 设置Excel不可见（后台运行）
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

    // 步骤2：获取Workbooks集合
    ptName = const_cast<LPOLESTR>(L"Workbooks");
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Workbooks property" << std::endl;
        pExcelApp->Release();
        return false;
    }

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dpNoArgs = { NULL, NULL, 0, 0 };
    hr = pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Workbooks collection" << std::endl;
        pExcelApp->Release();
        return false;
    }
    pWorkbooks = result.pdispVal;

    // 调用Add方法创建新工作簿
    ptName = const_cast<LPOLESTR>(L"Add");
    hr = pWorkbooks->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Add method" << std::endl;
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    VariantInit(&result);
    hr = pWorkbooks->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to create new workbook" << std::endl;
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }
    pWorkbook = result.pdispVal;

    // 步骤3：获取第一个工作表
    ptName = const_cast<LPOLESTR>(L"Worksheets");
    hr = pWorkbook->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Worksheets property" << std::endl;
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    VariantInit(&result);
    hr = pWorkbook->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Worksheets collection" << std::endl;
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }
    pWorksheets = result.pdispVal;

    // 获取第一个工作表（索引从1开始）
    VARIANT sheetIndex;
    VariantInit(&sheetIndex);
    sheetIndex.vt = VT_I4;
    sheetIndex.lVal = 1;

    ptName = const_cast<LPOLESTR>(L"Item");
    hr = pWorksheets->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Item method" << std::endl;
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
        std::wcerr << L"Failed to get worksheet" << std::endl;
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }
    pWorksheet = result.pdispVal;

    // 定义lambda表达式：写入字符串到单元格
    // 捕获列表 [&] 表示以引用方式捕获所有外部变量
    auto WriteCell = [&](long row, long col, const std::wstring& value) {
        // 准备单元格值（字符串类型，VT_BSTR）
        VARIANT cellValue;
        VariantInit(&cellValue);
        cellValue.vt = VT_BSTR;
        cellValue.bstrVal = SysAllocString(value.c_str());

        // 准备Cells属性的参数
        // 注意：参数顺序是反向的
        // rangeArgs[1] = 行号
        // rangeArgs[0] = 列号
        VARIANT rangeArgs[2];
        rangeArgs[1].vt = VT_I4;
        rangeArgs[1].lVal = row;
        rangeArgs[0].vt = VT_I4;
        rangeArgs[0].lVal = col;

        // 准备DISPPARAMS
        DISPPARAMS dpRange;
        dpRange.cArgs = 2;
        dpRange.rgvarg = rangeArgs;
        dpRange.cNamedArgs = 0;
        dpRange.rgdispidNamedArgs = NULL;

        VARIANT varRange;
        VariantInit(&varRange);

        // 获取Cells属性的DISPID
        LPOLESTR cellsName = const_cast<LPOLESTR>(L"Cells");
        DISPID cellsDispID;
        HRESULT hrCells = pWorksheet->GetIDsOfNames(IID_NULL, &cellsName, 1, LOCALE_USER_DEFAULT, &cellsDispID);
        if (FAILED(hrCells)) {
            std::wcerr << L"Failed to get Cells property. HRESULT: " << hrCells << std::endl;
            SysFreeString(cellValue.bstrVal);
            return;
        }

        // 调用Invoke获取单元格Range对象
        hrCells = pWorksheet->Invoke(cellsDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpRange, &varRange, NULL, NULL);
        if (FAILED(hrCells)) {
            std::wcerr << L"Failed to get cell (" << row << L"," << col << L"). HRESULT: " << hrCells << std::endl;
            VariantClear(&varRange);
            SysFreeString(cellValue.bstrVal);
            return;
        }

        // 检查返回值类型是否为IDispatch
        if (varRange.vt != VT_DISPATCH) {
            std::wcerr << L"Cell is not a dispatch object. Type: " << varRange.vt << std::endl;
            VariantClear(&varRange);
            SysFreeString(cellValue.bstrVal);
            return;
        }

        IDispatch* pCell = varRange.pdispVal;

        // 获取Value属性的DISPID
        LPOLESTR valueName = const_cast<LPOLESTR>(L"Value");
        DISPID valueDispID;
        HRESULT hrValue = pCell->GetIDsOfNames(IID_NULL, &valueName, 1, LOCALE_USER_DEFAULT, &valueDispID);
        if (FAILED(hrValue)) {
            std::wcerr << L"Failed to get Value property. HRESULT: " << hrValue << std::endl;
            pCell->Release();
            VariantClear(&varRange);
            SysFreeString(cellValue.bstrVal);
            return;
        }

        // 关键：准备设置属性的DISPPARAMS
        // 这里是之前出错的地方！
        // 必须设置 cNamedArgs = 1 和 rgdispidNamedArgs = &DISPID_PROPERTYPUT
        // 否则会返回 DISP_E_PARAMNOTFOUND (0x80020004) 错误
        DISPPARAMS dpValue;
        DISPID dispidPut = DISPID_PROPERTYPUT;
        dpValue.cArgs = 1;
        dpValue.rgvarg = &cellValue;
        dpValue.cNamedArgs = 1;                    // 必须设置为1
        dpValue.rgdispidNamedArgs = &dispidPut;   // 必须指向DISPID_PROPERTYPUT

        VARIANT varResult;
        VariantInit(&varResult);

        // 调用Invoke设置单元格值
        hrValue = pCell->Invoke(valueDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dpValue, &varResult, NULL, NULL);
        if (FAILED(hrValue)) {
            std::wcerr << L"Failed to set cell value (" << row << L"," << col << L"). HRESULT: " << hrValue << std::endl;
        }

        // 释放资源
        VariantClear(&varResult);
        pCell->Release();
        VariantClear(&varRange);
        SysFreeString(cellValue.bstrVal);
    };

    // 定义lambda表达式：写入整数到单元格
    // 与WriteCell类似，但值的类型是整数（VT_I4）
    // 用于写入名次（rank），这样在Excel中可以进行排序等操作
    auto WriteCellInt = [&](long row, long col, int value) {
        // 准备单元格值（整数类型，VT_I4）
        VARIANT cellValue;
        VariantInit(&cellValue);
        cellValue.vt = VT_I4;
        cellValue.lVal = value;

        // 准备Cells参数（与WriteCell相同）
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

        // 获取Cells属性
        LPOLESTR cellsName = const_cast<LPOLESTR>(L"Cells");
        DISPID cellsDispID;
        HRESULT hrCells = pWorksheet->GetIDsOfNames(IID_NULL, &cellsName, 1, LOCALE_USER_DEFAULT, &cellsDispID);
        if (FAILED(hrCells)) {
            std::wcerr << L"Failed to get Cells property. HRESULT: " << hrCells << std::endl;
            return;
        }

        hrCells = pWorksheet->Invoke(cellsDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpRange, &varRange, NULL, NULL);
        if (FAILED(hrCells)) {
            std::wcerr << L"Failed to get cell (" << row << L"," << col << L"). HRESULT: " << hrCells << std::endl;
            VariantClear(&varRange);
            return;
        }

        if (varRange.vt != VT_DISPATCH) {
            std::wcerr << L"Cell is not a dispatch object. Type: " << varRange.vt << std::endl;
            VariantClear(&varRange);
            return;
        }

        IDispatch* pCell = varRange.pdispVal;

        // 获取Value属性
        LPOLESTR valueName = const_cast<LPOLESTR>(L"Value");
        DISPID valueDispID;
        HRESULT hrValue = pCell->GetIDsOfNames(IID_NULL, &valueName, 1, LOCALE_USER_DEFAULT, &valueDispID);
        if (FAILED(hrValue)) {
            std::wcerr << L"Failed to get Value property. HRESULT: " << hrValue << std::endl;
            pCell->Release();
            VariantClear(&varRange);
            return;
        }

        // 准备DISPPARAMS（同样需要命名参数）
        DISPPARAMS dpValue;
        DISPID dispidPut = DISPID_PROPERTYPUT;
        dpValue.cArgs = 1;
        dpValue.rgvarg = &cellValue;
        dpValue.cNamedArgs = 1;
        dpValue.rgdispidNamedArgs = &dispidPut;

        VARIANT varResult;
        VariantInit(&varResult);

        // 设置单元格值
        hrValue = pCell->Invoke(valueDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dpValue, &varResult, NULL, NULL);
        if (FAILED(hrValue)) {
            std::wcerr << L"Failed to set cell value (" << row << L"," << col << L"). HRESULT: " << hrValue << std::endl;
        }

        // 释放资源
        VariantClear(&varResult);
        pCell->Release();
        VariantClear(&varRange);
    };

    // 步骤4：写入表头（第1行）
    WriteCell(1, 1, L"Rank");
    WriteCell(1, 2, L"Group");
    WriteCell(1, 3, L"Names");
    WriteCell(1, 4, L"Score");

    // 步骤5：写入数据行（从第2行开始）
    for (size_t i = 0; i < results.size(); i++) {
        long row = (long)(i + 2);  // 第2行开始
        WriteCellInt(row, 1, results[i].rank);   // 名次（整数格式）
        WriteCell(row, 2, results[i].group);     // 组别
        WriteCell(row, 3, results[i].names);     // 姓名
        WriteCell(row, 4, results[i].time);      // 成绩时间
    }

    // 步骤6：保存文件
    ptName = const_cast<LPOLESTR>(L"SaveAs");
    hr = pWorkbook->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get SaveAs method" << std::endl;
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    // 准备文件名参数
    VARIANT filename;
    VariantInit(&filename);
    filename.vt = VT_BSTR;
    filename.bstrVal = SysAllocString(filePath.c_str());

    // 根据文件扩展名选择格式
    // 51 = xlOpenXMLWorkbook (.xlsx)
    // 56 = xlExcel8 (.xls)
    int fileFormatValue = 56;  // 默认：xls格式
    std::wstring lowerPath = filePath;
    std::transform(lowerPath.begin(), lowerPath.end(), lowerPath.begin(), ::towlower);
    
    if (lowerPath.length() >= 5 && lowerPath.substr(lowerPath.length() - 5) == L".xlsx") {
        fileFormatValue = 51;  // xlsx格式
    }
    else if (lowerPath.length() >= 4 && lowerPath.substr(lowerPath.length() - 4) == L".xls") {
        fileFormatValue = 56;  // xls格式
    }
    else {
        fileFormatValue = 56;  // 其他扩展名，默认使用xls格式
    }

    // 准备文件格式参数
    VARIANT fileFormat;
    VariantInit(&fileFormat);
    fileFormat.vt = VT_I4;
    fileFormat.lVal = fileFormatValue;

    // 准备参数数组（注意顺序反向）
    // argsSave[1] = 第1个参数（FileName）
    // argsSave[0] = 第2个参数（FileFormat）
    VARIANT argsSave[2];
    argsSave[1] = filename;
    argsSave[0] = fileFormat;

    DISPPARAMS dpSave;
    dpSave.cArgs = 2;
    dpSave.rgvarg = argsSave;
    dpSave.cNamedArgs = 0;
    dpSave.rgdispidNamedArgs = NULL;

    // 调用SaveAs方法
    hr = pWorkbook->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpSave, NULL, NULL, NULL);
    SysFreeString(filename.bstrVal);

    if (FAILED(hr)) {
        std::wcerr << L"Failed to save file. HRESULT: " << hr << std::endl;
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    // 步骤7：清理资源
    // 释放工作表相关对象
    pWorksheet->Release();
    pWorksheets->Release();

    // 关闭工作簿
    ptName = const_cast<LPOLESTR>(L"Close");
    hr = pWorkbook->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (SUCCEEDED(hr)) {
        // 准备参数：不保存更改（已经用SaveAs保存过了）
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

        // 调用Close方法
        pWorkbook->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpClose, NULL, NULL, NULL);
    }

    // 释放工作簿相关对象
    pWorkbook->Release();
    pWorkbooks->Release();

    // 退出Excel
    ptName = const_cast<LPOLESTR>(L"Quit");
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (SUCCEEDED(hr)) {
        pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpNoArgs, NULL, NULL, NULL);
    }

    // 最后释放Application对象
    pExcelApp->Release();

    return true;
}

// ============================================================
// 宽字符字符串转换为UTF-8字符串
// ============================================================
// 参数：
//   wstr - 宽字符字符串（UTF-16，Windows原生格式）
// 
// 返回值：
//   UTF-8编码的字符串
// 
// 工作原理：
//   使用Windows API函数 WideCharToMultiByte 进行编码转换
// 
// 执行流程：
//   1. 检查输入是否为空
//   2. 第一次调用 WideCharToMultiByte，传入NULL缓冲区，获取所需大小
//   3. 根据返回大小分配缓冲区
//   4. 第二次调用 WideCharToMultiByte，执行实际转换
//   5. 返回转换后的字符串
// 
// 为什么要调用两次？
// - 第一次调用获取所需的缓冲区大小
// - 这样可以精确分配内存，避免缓冲区溢出
// - 这是Windows API的常见模式
// ============================================================
std::string DataProcessor::WStringToString(const std::wstring& wstr) {
    if (wstr.empty()) {
        return "";
    }

    // 第一次调用：获取所需的缓冲区大小
    int size = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, NULL, 0, NULL, NULL);
    if (size <= 0) {
        return "";
    }

    // 分配缓冲区
    // size - 1 是因为 WideCharToMultiByte 返回的大小包含null终止符
    // 而 std::string 不需要额外的null终止符空间
    std::string result(size - 1, 0);

    // 第二次调用：执行实际转换
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &result[0], size, NULL, NULL);

    return result;
}

// ============================================================
// CSV字段转义
// ============================================================
// 参数：
//   field - 原始字段值（宽字符字符串）
// 
// 返回值：
//   转义后的CSV字段字符串（UTF-8编码）
// 
// CSV格式规则（RFC 4180）：
//   1. 字段用逗号分隔
//   2. 如果字段包含逗号、双引号或换行符，必须用双引号包裹
//   3. 字段中的双引号必须用两个双引号表示（转义）
// 
// 转义规则：
//   情况1：字段不包含特殊字符 -> 直接返回原始值
//   情况2：字段包含逗号/双引号/换行符 -> 用双引号包裹，双引号转义为两个双引号
// ============================================================
std::string DataProcessor::EscapeCsvField(const std::wstring& field) {
    // 第一步：将宽字符转换为UTF-8字符串
    std::string str = WStringToString(field);
    
    // 第二步：检查是否需要转义
    // 需要转义的情况：包含逗号、双引号、换行符、回车符
    bool needsQuotes = false;
    if (str.find(',') != std::string::npos ||
        str.find('"') != std::string::npos ||
        str.find('\n') != std::string::npos ||
        str.find('\r') != std::string::npos) {
        needsQuotes = true;
    }

    // 不需要转义，直接返回
    if (!needsQuotes) {
        return str;
    }

    // 需要转义，构建转义后的字符串
    std::string escaped;
    escaped += '"';  // 开头双引号

    // 遍历每个字符
    for (char c : str) {
        if (c == '"') {
            escaped += "\"\"";  // 双引号转义为两个双引号
        }
        else {
            escaped += c;      // 其他字符直接输出
        }
    }

    escaped += '"';  // 结尾双引号
    return escaped;
}

// ============================================================
// 导出结果到CSV文件
// ============================================================
// 参数：
//   filePath - 输出文件路径
//   results  - 结果列表
// 
// 返回值：
//   true  - 导出成功
//   false - 导出失败
// 
// 执行流程：
//   1. 检查结果是否为空（输出警告但继续）
//   2. 使用 _wfopen_s 以二进制模式打开文件（支持中文路径）
//   3. 写入UTF-8 BOM（字节顺序标记）
//   4. 写入表头行
//   5. 遍历结果列表，写入每一行数据
//   6. 关闭文件
// 
// CSV文件格式：
//   - 编码：UTF-8 with BOM
//   - 分隔符：逗号
//   - 行尾：CRLF（\r\n）
//   - 字段转义：遵循RFC 4180标准
// 
// 为什么使用 _wfopen_s 而不是 std::ofstream？
// - std::ofstream 不支持UTF-8编码的中文路径
// - _wfopen_s 接受宽字符路径，可以正确处理中文字符
// 
// 为什么写入UTF-8 BOM？
// - BOM（Byte Order Mark）是文件开头的特殊字节序列
// - UTF-8 BOM 是 0xEF 0xBB 0xBF
// - 帮助Excel等应用程序识别文件编码
// - 如果没有BOM，Excel可能会使用错误的编码打开文件
// ============================================================
bool DataProcessor::ExportResultsToCsv(const std::wstring& filePath, const std::vector<ResultEntry>& results) {
    if (results.empty()) {
        std::wcerr << L"Warning: No results to export" << std::endl;
    }

    // 使用 _wfopen_s 打开文件（支持中文路径）
    // L"wb" 表示：w=写入，b=二进制模式
    FILE* file = NULL;
    errno_t err = _wfopen_s(&file, filePath.c_str(), L"wb");
    if (err != 0 || file == NULL) {
        std::wcerr << L"Failed to create CSV file: " << filePath << std::endl;
        return false;
    }

    // 写入UTF-8 BOM
    // 0xEF 0xBB 0xBF 是UTF-8的字节顺序标记
    unsigned char bom[] = { 0xEF, 0xBB, 0xBF };
    fwrite(bom, 1, sizeof(bom), file);

    // 写入表头行
    // 行尾使用 CRLF（\r\n），符合CSV标准
    fprintf(file, "Rank,Group,Names,Score\r\n");

    // 遍历结果列表，写入每一行数据
    for (const auto& result : results) {
        // 第1列：名次（整数，不需要转义）
        fprintf(file, "%d,", result.rank);

        // 第2列：组别（字符串，需要转义）
        fprintf(file, "%s,", EscapeCsvField(result.group).c_str());

        // 第3列：姓名（字符串，需要转义）
        fprintf(file, "%s,", EscapeCsvField(result.names).c_str());

        // 第4列：成绩时间（字符串，需要转义）
        // 注意：最后一个字段后面没有逗号，直接跟换行
        fprintf(file, "%s\r\n", EscapeCsvField(result.time).c_str());
    }

    // 关闭文件
    fclose(file);

    return true;
}
