#pragma execution_character_set("utf-8")

/**
 * @class DataProcessor
 * @brief 数据处理器类
 * 
 * 负责数据匹配和结果导出功能。
 */

#include "DataProcessor.h"
#include <windows.h>
#include <comutil.h>
#include <iostream>
#include <fstream>
#include <sstream>
#include <iomanip>
#include <algorithm>

#pragma comment(lib, "comsuppw.lib")

/**
 * @brief 构造函数
 */
DataProcessor::DataProcessor() {
}

/**
 * @brief 析构函数
 */
DataProcessor::~DataProcessor() {
}

/**
 * @brief 数据匹配处理
 * 
 * 根据组别匹配男生和女生姓名，生成结果条目。
 * 
 * @param participants 报名信息列表
 * @param scoreEntries 成绩条目列表
 * @param results 结果列表
 */
void DataProcessor::ProcessData(const std::vector<Participant>& participants,
                                 const std::vector<ScoreEntry>& scoreEntries,
                                 std::vector<ResultEntry>& results) {
    results.clear();

    if (participants.empty()) {
        std::wcerr << L"Warning: No participants data available" << std::endl;
    }

    if (scoreEntries.empty()) {
        std::wcerr << L"Error: No score entries available" << std::endl;
        return;
    }

    // 双映射表设计：使用两个独立的映射表分别存储男生和女生信息
    // 
    // 为什么使用两个映射表：
    // 1. 同一组的男生和女生可能在报名信息的不同行
    //    例如：成绩清单23组：罗晓东 梁馨尹
    //         罗晓东（男生23A）可能在第3行
    //         梁馨尹（女生23B）可能在第5行
    // 2. 如果假设每一行的男生和女生是同一组的，会导致匹配错误
    // 3. 双映射表允许男生和女生独立映射到组号
    //
    // 映射表结构：key = 组号，value = 姓名
    std::unordered_map<int, std::wstring> maleMap;
    std::unordered_map<int, std::wstring> femaleMap;

    // 建立映射表：遍历所有报名信息
    for (const auto& participant : participants) {
        // 如果男生组号有效且姓名非空，添加到男生映射表
        if (participant.maleGroupNumber >= 0 && !participant.maleName.empty()) {
            maleMap[participant.maleGroupNumber] = participant.maleName;
        }
        // 如果女生组号有效且姓名非空，添加到女生映射表
        if (participant.femaleGroupNumber >= 0 && !participant.femaleName.empty()) {
            femaleMap[participant.femaleGroupNumber] = participant.femaleName;
        }
    }

    // 遍历成绩条目，匹配姓名
    for (const auto& scoreEntry : scoreEntries) {
        ResultEntry result;
        result.rank = scoreEntry.rank;
        result.group = scoreEntry.group;
        result.time = scoreEntry.time;

        // 组号无效的情况
        if (scoreEntry.groupNumber < 0) {
            std::wcerr << L"Warning: Invalid group number for rank " << scoreEntry.rank << std::endl;
            result.names = L"Invalid Group";
        }
        else {
            std::wstring maleName;
            std::wstring femaleName;

            // 从男生映射表查找当前组的男生姓名
            auto maleIt = maleMap.find(scoreEntry.groupNumber);
            if (maleIt != maleMap.end()) {
                maleName = maleIt->second;
            }

            // 从女生映射表查找当前组的女生姓名
            auto femaleIt = femaleMap.find(scoreEntry.groupNumber);
            if (femaleIt != femaleMap.end()) {
                femaleName = femaleIt->second;
            }

            // 姓名组合逻辑：
            // - 男生和女生都没找到：标记为 Unknown
            // - 只有男生：使用男生姓名
            // - 只有女生：使用女生姓名
            // - 都有：男生姓名 + 空格 + 女生姓名
            if (maleName.empty() && femaleName.empty()) {
                std::wcerr << L"Warning: Participant info not found for group number " << scoreEntry.groupNumber << std::endl;
                result.names = L"Unknown";
            }
            else if (maleName.empty()) {
                result.names = femaleName;
            }
            else if (femaleName.empty()) {
                result.names = maleName;
            }
            else {
                result.names = maleName + L" " + femaleName;
            }
        }

        results.push_back(result);
    }
}

/**
 * @brief 导出结果到Excel文件
 * 
 * 使用COM自动化技术创建Excel文件并写入结果数据。
 * 
 * @param filePath 输出文件路径
 * @param results 结果列表
 * @return true-导出成功，false-导出失败
 */
bool DataProcessor::ExportResults(const std::wstring& filePath, const std::vector<ResultEntry>& results) {
    if (results.empty()) {
        std::wcerr << L"Warning: No results to export" << std::endl;
    }

    IDispatch* pExcelApp = NULL;
    IDispatch* pWorkbooks = NULL;
    IDispatch* pWorkbook = NULL;
    IDispatch* pWorksheets = NULL;
    IDispatch* pWorksheet = NULL;

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

    auto WriteCell = [&](long row, long col, const std::wstring& value) {
        VARIANT cellValue;
        VariantInit(&cellValue);
        cellValue.vt = VT_BSTR;
        cellValue.bstrVal = SysAllocString(value.c_str());

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

        LPOLESTR cellsName = const_cast<LPOLESTR>(L"Cells");
        DISPID cellsDispID;
        HRESULT hrCells = pWorksheet->GetIDsOfNames(IID_NULL, &cellsName, 1, LOCALE_USER_DEFAULT, &cellsDispID);
        if (FAILED(hrCells)) {
            std::wcerr << L"Failed to get Cells property. HRESULT: " << hrCells << std::endl;
            SysFreeString(cellValue.bstrVal);
            return;
        }

        hrCells = pWorksheet->Invoke(cellsDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpRange, &varRange, NULL, NULL);
        if (FAILED(hrCells)) {
            std::wcerr << L"Failed to get cell (" << row << L"," << col << L"). HRESULT: " << hrCells << std::endl;
            VariantClear(&varRange);
            SysFreeString(cellValue.bstrVal);
            return;
        }

        if (varRange.vt != VT_DISPATCH) {
            std::wcerr << L"Cell is not a dispatch object. Type: " << varRange.vt << std::endl;
            VariantClear(&varRange);
            SysFreeString(cellValue.bstrVal);
            return;
        }

        IDispatch* pCell = varRange.pdispVal;

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

        DISPPARAMS dpValue;
        DISPID dispidPut = DISPID_PROPERTYPUT;
        dpValue.cArgs = 1;
        dpValue.rgvarg = &cellValue;
        dpValue.cNamedArgs = 1;
        dpValue.rgdispidNamedArgs = &dispidPut;

        VARIANT varResult;
        VariantInit(&varResult);

        hrValue = pCell->Invoke(valueDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dpValue, &varResult, NULL, NULL);
        if (FAILED(hrValue)) {
            std::wcerr << L"Failed to set cell value (" << row << L"," << col << L"). HRESULT: " << hrValue << std::endl;
        }

        VariantClear(&varResult);
        pCell->Release();
        VariantClear(&varRange);
        SysFreeString(cellValue.bstrVal);
    };

    auto WriteCellInt = [&](long row, long col, int value) {
        VARIANT cellValue;
        VariantInit(&cellValue);
        cellValue.vt = VT_I4;
        cellValue.lVal = value;

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

        LPOLESTR valueName = const_cast<LPOLESTR>(L"Value");
        DISPID valueDispID;
        HRESULT hrValue = pCell->GetIDsOfNames(IID_NULL, &valueName, 1, LOCALE_USER_DEFAULT, &valueDispID);
        if (FAILED(hrValue)) {
            std::wcerr << L"Failed to get Value property. HRESULT: " << hrValue << std::endl;
            pCell->Release();
            VariantClear(&varRange);
            return;
        }

        DISPPARAMS dpValue;
        DISPID dispidPut = DISPID_PROPERTYPUT;
        dpValue.cArgs = 1;
        dpValue.rgvarg = &cellValue;
        dpValue.cNamedArgs = 1;
        dpValue.rgdispidNamedArgs = &dispidPut;

        VARIANT varResult;
        VariantInit(&varResult);

        hrValue = pCell->Invoke(valueDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dpValue, &varResult, NULL, NULL);
        if (FAILED(hrValue)) {
            std::wcerr << L"Failed to set cell value (" << row << L"," << col << L"). HRESULT: " << hrValue << std::endl;
        }

        VariantClear(&varResult);
        pCell->Release();
        VariantClear(&varRange);
    };

    WriteCell(1, 1, L"Rank");
    WriteCell(1, 2, L"Group");
    WriteCell(1, 3, L"Names");
    WriteCell(1, 4, L"Score");

    for (size_t i = 0; i < results.size(); i++) {
        long row = (long)(i + 2);
        WriteCellInt(row, 1, results[i].rank);
        WriteCell(row, 2, results[i].group);
        WriteCell(row, 3, results[i].names);
        WriteCell(row, 4, results[i].time);
    }

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

    VARIANT filename;
    VariantInit(&filename);
    filename.vt = VT_BSTR;
    filename.bstrVal = SysAllocString(filePath.c_str());

    // Excel 文件格式代码（XlFileFormat 枚举）：
    // 51 = xlOpenXMLWorkbook (Excel 2007+ .xlsx 格式)
    // 56 = xlExcel8 (Excel 97-2003 .xls 格式)
    // 默认使用 56 (.xls) 以保持兼容性
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
    else {
        fileFormatValue = 56;  // 默认使用 .xls 格式
    }

    VARIANT fileFormat;
    VariantInit(&fileFormat);
    fileFormat.vt = VT_I4;
    fileFormat.lVal = fileFormatValue;

    // SaveAs 方法参数说明：
    // 参数顺序（从右到左）：
    // argsSave[0] = 格式代码 (FileFormat)
    // argsSave[1] = 文件名 (Filename)
    // 注意：DISPPARAMS 的参数顺序是逆序的（栈结构）
    VARIANT argsSave[2];
    argsSave[1] = filename;
    argsSave[0] = fileFormat;

    DISPPARAMS dpSave;
    dpSave.cArgs = 2;
    dpSave.rgvarg = argsSave;
    dpSave.cNamedArgs = 0;
    dpSave.rgdispidNamedArgs = NULL;

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
 * @brief 宽字符字符串转换为UTF-8字符串
 * 
 * 使用Windows API WideCharToMultiByte进行编码转换。
 * 
 * @param wstr 宽字符字符串（UTF-16）
 * @return UTF-8编码的字符串
 */
std::string DataProcessor::WStringToString(const std::wstring& wstr) const {
    if (wstr.empty()) {
        return "";
    }

    int size = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, NULL, 0, NULL, NULL);
    if (size <= 0) {
        return "";
    }

    std::string result(size - 1, 0);
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &result[0], size, NULL, NULL);

    return result;
}

/**
 * @brief CSV字段转义
 * 
 * 处理包含逗号、双引号的CSV字段。
 * 
 * @param field 原始字段值
 * @return 转义后的CSV字段字符串
 */
std::string DataProcessor::EscapeCsvField(const std::wstring& field) const {
    std::string str = WStringToString(field);
    
    bool needsQuotes = false;
    if (str.find(',') != std::string::npos ||
        str.find('"') != std::string::npos ||
        str.find('\n') != std::string::npos ||
        str.find('\r') != std::string::npos) {
        needsQuotes = true;
    }

    if (!needsQuotes) {
        return str;
    }

    std::string escaped;
    escaped += '"';

    for (char c : str) {
        if (c == '"') {
            escaped += "\"\"";
        }
        else {
            escaped += c;
        }
    }

    escaped += '"';
    return escaped;
}

/**
 * @brief 导出结果到CSV文件
 * 
 * 使用UTF-8编码。
 * 
 * @param filePath 输出文件路径
 * @param results 结果列表
 * @return true-导出成功，false-导出失败
 */
bool DataProcessor::ExportResultsToCsv(const std::wstring& filePath, const std::vector<ResultEntry>& results) {
    if (results.empty()) {
        std::wcerr << L"Warning: No results to export" << std::endl;
    }

    FILE* file = NULL;
    // _wfopen_s 使用宽字符路径，支持中文路径
    // "wb" 表示以二进制写入模式打开
    errno_t err = _wfopen_s(&file, filePath.c_str(), L"wb");
    if (err != 0 || file == NULL) {
        std::wcerr << L"Failed to create CSV file: " << filePath << std::endl;
        return false;
    }

    // UTF-8 BOM（字节顺序标记）：0xEF 0xBB 0xBF
    // 
    // 为什么需要 BOM：
    // 1. 通知其他程序这是一个 UTF-8 编码的文件
    // 2. 特别是 Excel 打开 CSV 文件时，如果没有 BOM，
    //    它可能会错误地使用 ANSI 编码，导致中文乱码
    // 3. 虽然 BOM 不是 UTF-8 标准要求的，但在 Windows 平台上
    //    这是一个广泛使用的约定
    unsigned char bom[] = { 0xEF, 0xBB, 0xBF };
    fwrite(bom, 1, sizeof(bom), file);

    // 写入表头：Rank, Group, Names, Score
    // 使用 \r\n 作为换行符（Windows 标准）
    fprintf(file, "Rank,Group,Names,Score\r\n");

    for (const auto& result : results) {
        fprintf(file, "%d,", result.rank);
        // 使用 EscapeCsvField 函数转义字段，处理包含逗号、双引号的情况
        fprintf(file, "%s,", EscapeCsvField(result.group).c_str());
        fprintf(file, "%s,", EscapeCsvField(result.names).c_str());
        fprintf(file, "%s\r\n", EscapeCsvField(result.time).c_str());
    }

    fclose(file);

    return true;
}
