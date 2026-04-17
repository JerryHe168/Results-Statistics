#pragma execution_character_set("utf-8")

/**
 * @class ExcelSession
 * @brief Excel COM 会话封装类
 * 
 * 使用 RAII 模式封装 Excel COM 会话的整个生命周期，
 * 包括创建 Excel 实例、打开文件、获取数据、自动清理资源。
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
 * @brief 打开 Excel 文件
 * 
 * 创建 Excel Application 实例，打开指定的工作簿，
 * 并获取第一个工作表的 UsedRange 数据。
 * 
 * @param filePath Excel 文件路径
 * @return true-打开成功，false-打开失败
 */
bool ExcelSession::OpenFile(const std::wstring& filePath) {
    // 先释放之前的资源
    Release();

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
        std::wcerr << L"Failed to get Workbooks property. HRESULT: " << hr << std::endl;
        return false;
    }

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dpNoArgs = { NULL, NULL, 0, 0 };
    hr = m_pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Workbooks collection. HRESULT: " << hr << std::endl;
        return false;
    }
    m_pWorkbooks = result.pdispVal;

    // 打开文件
    ptName = const_cast<LPOLESTR>(L"Open");
    hr = m_pWorkbooks->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
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

    VariantInit(&result);
    hr = m_pWorkbooks->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpOpen, &result, NULL, NULL);
    SysFreeString(filename.bstrVal);

    if (FAILED(hr)) {
        std::wcerr << L"Failed to open file. HRESULT: " << hr << std::endl;
        return false;
    }
    m_pWorkbook = result.pdispVal;

    // 获取 Worksheets 集合
    ptName = const_cast<LPOLESTR>(L"Worksheets");
    hr = m_pWorkbook->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Worksheets property. HRESULT: " << hr << std::endl;
        return false;
    }

    VariantInit(&result);
    hr = m_pWorkbook->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Worksheets collection. HRESULT: " << hr << std::endl;
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
        std::wcerr << L"Failed to get Item method. HRESULT: " << hr << std::endl;
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
        std::wcerr << L"Failed to get worksheet. HRESULT: " << hr << std::endl;
        return false;
    }
    m_pWorksheet = result.pdispVal;

    // 获取 UsedRange
    ptName = const_cast<LPOLESTR>(L"UsedRange");
    hr = m_pWorksheet->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get UsedRange property. HRESULT: " << hr << std::endl;
        return false;
    }

    VariantInit(&result);
    hr = m_pWorksheet->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get UsedRange. HRESULT: " << hr << std::endl;
        return false;
    }
    m_pRange = result.pdispVal;

    // 获取 Value 属性（SAFEARRAY）
    ptName = const_cast<LPOLESTR>(L"Value");
    hr = m_pRange->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Value property. HRESULT: " << hr << std::endl;
        return false;
    }

    VariantInit(&m_varResult);
    hr = m_pRange->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &m_varResult, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get cell data. HRESULT: " << hr << std::endl;
        VariantClear(&m_varResult);
        return false;
    }

    // 关键修复：使用位运算检查 VT_ARRAY 类型
    // Excel 返回 VT_ARRAY | VT_VARIANT (8204 = 0x200C)
    if ((m_varResult.vt & VT_ARRAY) != VT_ARRAY) {
        std::wcerr << L"Cell data is not an array. Type: " << m_varResult.vt << std::endl;
        VariantClear(&m_varResult);
        return false;
    }

    // 获取 SAFEARRAY 指针
    m_pSafeArray = m_varResult.parray;
    SafeArrayGetLBound(m_pSafeArray, 1, &m_lBound);
    SafeArrayGetUBound(m_pSafeArray, 1, &m_uBound);

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
    VARIANT cellValue;
    if (!GetCellValue(row, col, cellValue)) {
        return defaultVal;
    }

    std::wstring result = defaultVal;

    // VARIANT 类型处理：
    // - VT_BSTR: 字符串类型（最常见）
    // - VT_I4: 32位整数
    // - VT_R8: 双精度浮点数
    if (cellValue.vt == VT_BSTR) {
        result = cellValue.bstrVal;
    }
    else if (cellValue.vt == VT_I4) {
        result = std::to_wstring(cellValue.lVal);
    }
    else if (cellValue.vt == VT_R8) {
        result = std::to_wstring((long long)cellValue.dblVal);
    }

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
    VARIANT cellValue;
    if (!GetCellValue(row, col, cellValue)) {
        return defaultVal;
    }

    long result = defaultVal;

    if (cellValue.vt == VT_I4) {
        result = cellValue.lVal;
    }
    else if (cellValue.vt == VT_R8) {
        result = (long)cellValue.dblVal;
    }
    else if (cellValue.vt == VT_BSTR) {
        // 字符串格式，尝试转换为整数
        try {
            result = std::stoi(cellValue.bstrVal);
        }
        catch (...) {
            result = defaultVal;
        }
    }

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
    VARIANT cellValue;
    if (!GetCellValue(row, col, cellValue)) {
        return defaultVal;
    }

    double result = defaultVal;

    if (cellValue.vt == VT_R8) {
        result = cellValue.dblVal;
    }
    else if (cellValue.vt == VT_I4) {
        result = (double)cellValue.lVal;
    }

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
    VARIANT cellValue;
    if (!GetCellValue(row, col, cellValue)) {
        return L"";
    }

    std::wstring result;

    // 时间格式处理：
    // Excel 中的时间可能以三种形式存储：
    // - VT_BSTR: 字符串格式（如 "0:37:06"）
    // - VT_DATE: Variant 时间格式（使用 VariantTimeToSystemTime 转换）
    // - VT_R8: 浮点数格式（0.0 = 0:00:00, 1.0 = 24:00:00）
    if (cellValue.vt == VT_BSTR) {
        result = cellValue.bstrVal;
    }
    else if (cellValue.vt == VT_DATE) {
        SYSTEMTIME st;
        VariantTimeToSystemTime(cellValue.date, &st);
        wchar_t buffer[32];
        // 格式：小时:分钟:秒（分钟和秒补零）
        swprintf_s(buffer, L"%d:%02d:%02d", st.wHour, st.wMinute, st.wSecond);
        result = buffer;
    }
    else if (cellValue.vt == VT_R8) {
        double timeVal = cellValue.dblVal;
        // 浮点数转时间：0.0 = 0:00:00, 1.0 = 24:00:00
        // 浮点数转时间公式：
        // - 小时 = timeVal * 24  （一天24小时）
        // - 分钟 = (小数部分) * 60  （一小时60分钟）
        // - 秒 = (小数部分) * 60    （一分钟60秒）
        int hours = (int)(timeVal * 24);  // 乘以24得到小时数
        int minutes = (int)((timeVal * 24 - hours) * 60);  // 小数部分乘以60得到分钟
        int seconds = (int)(((timeVal * 24 - hours) * 60 - minutes) * 60);  // 小数部分乘以60得到秒
        wchar_t buffer[32];
        swprintf_s(buffer, L"%d:%02d:%02d", hours, minutes, seconds);
        result = buffer;
    }

    VariantClear(&cellValue);
    return result;
}

/**
 * @class ExcelReader
 * @brief Excel文件读取器类
 * 
 * 负责使用COM自动化技术读取Excel格式的报名信息和成绩清单文件。
 */

/**
 * @brief 构造函数
 */
ExcelReader::ExcelReader() {
}

/**
 * @brief 析构函数
 */
ExcelReader::~ExcelReader() {
}

/**
 * @brief 从编号中提取组号
 * 
 * 使用正则表达式匹配字符串中的第一个连续数字序列。
 * 
 * @param id 编号字符串
 * @return 提取的组号，无法提取则返回-1
 */
int ExcelReader::ExtractGroupNumber(const std::wstring& id) const {
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
    participants.clear();

    ExcelSession session;
    if (!session.OpenFile(filePath)) {
        return false;
    }

    // 报名信息的列结构：
    // 列 1：男生编号（如 "23A", "18A"）
    // 列 2：男生姓名
    // 列 3：女生编号（如 "16B", "13B"）
    // 列 4：女生姓名

    // 遍历数据行（从 lBound + 1 开始是为了跳过表头行）
    long lBound = session.GetRowLowerBound();
    long uBound = session.GetRowUpperBound();

    for (long row = lBound + 1; row <= uBound; row++) {
        Participant participant;

        // 读取第 1 列：男生编号
        participant.maleId = session.GetCellString(row, 1);

        // 读取第 2 列：男生姓名
        participant.maleName = session.GetCellString(row, 2);

        // 读取第 3 列：女生编号
        participant.femaleId = session.GetCellString(row, 3);

        // 读取第 4 列：女生姓名
        participant.femaleName = session.GetCellString(row, 4);

        // 从编号中提取组号
        participant.maleGroupNumber = ExtractGroupNumber(participant.maleId);
        participant.femaleGroupNumber = ExtractGroupNumber(participant.femaleId);

        // 只添加有姓名的记录
        if (!participant.maleName.empty() || !participant.femaleName.empty()) {
            participants.push_back(participant);
        }
    }

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
    scoreEntries.clear();

    ExcelSession session;
    if (!session.OpenFile(filePath)) {
        return false;
    }

    // 成绩清单的列结构：
    // 列 1：名次
    // 列 2：组别（如 "23组"）
    // 列 3：成绩时间

    // 遍历数据行（从 lBound + 1 开始是为了跳过表头行）
    long lBound = session.GetRowLowerBound();
    long uBound = session.GetRowUpperBound();

    for (long row = lBound + 1; row <= uBound; row++) {
        ScoreEntry entry;

        // 读取第 1 列：名次
        entry.rank = session.GetCellLong(row, 1, 0);

        // 读取第 2 列：组别
        // 组别可能是字符串（如 "23组"）或数字
        VARIANT cellValue;
        if (session.GetCellValue(row, 2, cellValue)) {
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
        }

        // 读取第 3 列：成绩时间
        entry.time = session.GetCellTime(row, 3);

        // 从组别中提取组号
        entry.groupNumber = ExtractGroupNumber(entry.group);

        // 只添加有效名次的记录
        if (entry.rank > 0) {
            scoreEntries.push_back(entry);
        }
    }

    return true;
}
