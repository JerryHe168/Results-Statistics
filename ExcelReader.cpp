#pragma execution_character_set("utf-8")

// ============================================================
// ExcelReader.cpp - Excel文件读取器实现
// ============================================================
// 本文件实现了使用COM自动化技术读取Excel文件的功能
// 
// 主要功能：
// 1. ExcelReader() / ~ExcelReader()：初始化和释放COM库
// 2. ExtractGroupNumber：从编号中提取组号
// 3. ReadRegistrationInfo：读取报名信息Excel文件
// 4. ReadScoreList：读取成绩清单Excel文件
// 
// COM自动化技术说明：
// COM（Component Object Model）是微软的组件对象模型
// COM自动化允许应用程序通过IDispatch接口动态调用Excel对象的方法和属性
// 
// 关键技术点：
// 1. CoInitialize / CoUninitialize：初始化和释放COM库
// 2. CLSIDFromProgID：根据程序ID获取类ID
// 3. CoCreateInstance：创建COM对象实例
// 4. IDispatch::GetIDsOfNames：根据名称获取DISPID
// 5. IDispatch::Invoke：通过DISPID调用方法或访问属性
// 6. VARIANT：通用数据类型，支持多种数据类型
// 7. SAFEARRAY：安全数组，用于跨进程传递数组数据
// 8. DISPPARAMS：参数传递结构，支持命名参数
// 
// 关键修复点（历史问题）：
// 1. VT_ARRAY类型检查：使用位运算 (varResult.vt & VT_ARRAY) != VT_ARRAY
//    因为Excel返回的是 VT_ARRAY | VT_VARIANT (8204)，而不是单纯的 VT_ARRAY
// 
// 2. DISPPARAMS属性设置：
//    设置 cNamedArgs = 1 和 rgdispidNamedArgs = &DISPID_PROPERTYPUT
//    否则会返回 DISP_E_PARAMNOTFOUND (0x80020004) 错误
// ============================================================

#include "ExcelReader.h"
#include <windows.h>
#include <comutil.h>
#include <iostream>
#include <sstream>
#include <regex>
#include <algorithm>

#pragma comment(lib, "comsuppw.lib")

// ============================================================
// 构造函数
// ============================================================
// 功能：初始化COM库
// 
// 注意：
// - CoInitialize必须在使用任何COM功能之前调用
// - 参数NULL表示使用默认的并发模型
// - 每个线程都需要单独初始化COM库
// - 必须在析构函数中调用CoUninitialize进行配对
// 
// 为什么在构造函数中初始化COM？
// - ExcelReader类的主要功能就是操作Excel COM对象
// - 在对象创建时初始化COM，确保所有方法都能使用COM功能
// ============================================================
ExcelReader::ExcelReader() {
    // 初始化COM库
    // 参数NULL表示使用默认的并发模型（单线程单元）
    CoInitialize(NULL);
}

// ============================================================
// 析构函数
// ============================================================
// 功能：释放COM库
// 
// 注意：
// - CoUninitialize必须与CoInitialize配对调用
// - 在释放COM库之前，必须确保所有COM对象都已释放
// - 否则可能导致资源泄漏或程序崩溃
// 
// 为什么在析构函数中释放COM？
// - 与构造函数中的CoInitialize配对
// - 确保对象销毁时COM资源被正确释放
// ============================================================
ExcelReader::~ExcelReader() {
    // 释放COM库
    // 必须与CoInitialize配对调用
    CoUninitialize();
}

// ============================================================
// 从编号中提取组号
// ============================================================
// 参数：
//   id - 编号字符串，如 "12A", "18B", "13组"
// 
// 返回值：
//   提取的组号（整数），如果无法提取数字则返回 -1
// 
// 工作原理：
//   使用正则表达式匹配字符串中的第一个连续数字序列
// 
// 支持的格式：
//   - "12A"   -> 12  （男生编号格式：数字+字母A）
//   - "18B"   -> 18  （女生编号格式：数字+字母B）
//   - "13组"  -> 13  （组别格式：数字+中文"组"）
//   - "第5组" -> 5   （中文组别格式）
// 
// 正则表达式说明：
//   L"(\\d+)"  - 匹配一个或多个连续数字
//   \\d        - 匹配任意数字（等价于[0-9]）
//   +          - 匹配一个或多个
//   ()         - 捕获组，用于提取匹配的内容
// 
// 注意：
//   - 只提取第一个数字序列
//   - 例如 "12A34" 只提取 "12"，返回 12
//   - 如果没有数字，返回 -1
// ============================================================
int ExcelReader::ExtractGroupNumber(const std::wstring& id) {
    // 创建正则表达式：匹配一个或多个连续数字
    // L"(\\d+)" 中的 L 表示宽字符字符串
    // (\\d+) 中的 \\ 是转义的 \，在正则表达式中表示 \d
    std::wregex regex(L"(\\d+)");
    std::wsmatch match;  // 存储匹配结果

    // 在id中搜索匹配正则表达式的内容
    // regex_search 在字符串中搜索匹配项
    // 与 regex_match 不同，regex_match 要求整个字符串匹配
    if (std::regex_search(id, match, regex)) {
        // match[1] 是第一个捕获组的内容
        // match[0] 是整个匹配的内容
        // match[1], match[2]... 是各个捕获组的内容
        // 
        // 例如：id = "12A"
        // match[0] = "12"（整个匹配）
        // match[1] = "12"（第一个捕获组）
        // 
        // .str() 将匹配结果转换为 wstring
        // std::stoi 将 wstring 转换为整数
        return std::stoi(match[1].str());
    }

    // 如果没有找到数字，返回 -1 表示失败
    return -1;
}

// ============================================================
// 读取报名信息Excel文件
// ============================================================
// 参数：
//   filePath - Excel文件路径
//   participants - 输出参数，存储读取到的报名信息列表
// 
// 返回值：
//   true  - 读取成功
//   false - 读取失败
// 
// Excel文件格式要求：
//   第1列：男生编号（如 "12A", "18A"）
//   第2列：男生姓名
//   第3列：女生编号（如 "17B", "13B"）
//   第4列：女生姓名
//   第1行：表头（会自动跳过，从第2行开始读取数据）
// 
// 执行流程：
//   1. 声明COM接口指针
//   2. 创建Excel.Application实例
//   3. 设置Excel不可见（后台运行，不显示界面）
//   4. 获取Workbooks集合
//   5. 打开指定的Excel文件
//   6. 获取Worksheets集合
//   7. 获取第一个工作表
//   8. 获取UsedRange（已使用的单元格区域）
//   9. 获取单元格数据（返回SAFEARRAY）
//   10. 检查返回数据类型（关键修复：使用位运算检查VT_ARRAY）
//   11. 获取SAFEARRAY的边界信息
//   12. 遍历数组，从第2行开始读取数据（跳过表头）
//   13. 处理各种VARIANT数据类型（VT_BSTR, VT_I4, VT_R8等）
//   14. 提取组号
//   15. 过滤并添加到结果列表
//   16. 释放所有COM资源（顺序很重要！）
//   17. 关闭工作簿，退出Excel
// 
// 关键修复点：
// 1. VT_ARRAY类型检查：
//    错误方式：if (varResult.vt != VT_ARRAY)
//    正确方式：if ((varResult.vt & VT_ARRAY) != VT_ARRAY)
//    
//    原因：Excel返回的是 VT_ARRAY | VT_VARIANT (8204)
//    VT_ARRAY = 0x2000, VT_VARIANT = 0x000C
//    8204 = 0x200C = VT_ARRAY | VT_VARIANT
//    必须使用位运算来检查是否包含VT_ARRAY标志
// 
// 2. COM对象释放顺序：
//    必须按照创建的逆序释放
//    创建顺序：pExcelApp -> pWorkbooks -> pWorkbook -> pWorksheets -> pWorksheet -> pRange
//    释放顺序：pRange -> pWorksheet -> pWorksheets -> pWorkbook -> pWorkbooks -> pExcelApp
// 
// 3. SAFEARRAY索引：
//    Excel的SAFEARRAY索引从1开始，不是从0开始
//    第一行是表头，从第2行（lBound + 1）开始读取数据
// 
// 4. VARIANT类型处理：
//    Excel单元格可能返回不同类型的值：
//    - VT_BSTR：字符串（文本格式）
//    - VT_I4：32位整数
//    - VT_R8：64位浮点数（Excel默认数字格式）
//    需要分别处理这些类型
// ============================================================
bool ExcelReader::ReadRegistrationInfo(const std::wstring& filePath, std::vector<Participant>& participants) {
    // 声明COM接口指针
    // IDispatch是COM自动化的核心接口，支持动态调用
    IDispatch* pExcelApp = NULL;      // Excel.Application对象
    IDispatch* pWorkbooks = NULL;      // Workbooks集合
    IDispatch* pWorkbook = NULL;       // 打开的Workbook
    IDispatch* pWorksheets = NULL;     // Worksheets集合
    IDispatch* pWorksheet = NULL;      // 第一个Worksheet
    IDispatch* pRange = NULL;          // UsedRange区域

    // --------------------------------------------------------
    // 步骤1：创建Excel.Application实例
    // --------------------------------------------------------
    // 流程：
    // 1. CLSIDFromProgID：根据程序ID "Excel.Application" 获取类ID
    // 2. CoCreateInstance：根据类ID创建COM对象实例
    // 
    // 为什么需要这两步？
    // - COM系统使用CLSID（128位GUID）来标识组件
    // - 人类可读的程序ID（ProgID）如 "Excel.Application" 需要转换为CLSID
    // - CLSIDFromProgID 完成这个转换
    // --------------------------------------------------------

    // 获取Excel.Application的CLSID
    // 参数：
    //   L"Excel.Application" - 程序ID
    //   &clsid - 输出参数，接收CLSID
    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to create Excel application instance. HRESULT: " << hr << std::endl;
        return false;
    }

    // 创建Excel COM对象实例
    // 参数：
    //   clsid - 类ID
    //   NULL - 外部对象（用于聚合，NULL表示不聚合）
    //   CLSCTX_LOCAL_SERVER - 上下文（本地服务器，Excel是独立进程）
    //   IID_IDispatch - 接口ID（请求IDispatch接口）
    //   (void**)&pExcelApp - 输出参数，接收接口指针
    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pExcelApp);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to start Excel. HRESULT: " << hr << std::endl;
        return false;
    }

    // --------------------------------------------------------
    // 步骤2：设置Excel不可见（后台运行）
    // --------------------------------------------------------
    // 默认情况下，Excel会显示界面
    // 我们希望在后台运行，不显示界面
    // 需要设置 Application.Visible = false
    // 
    // 如何通过IDispatch设置属性？
    // 1. GetIDsOfNames：根据属性名 "Visible" 获取DISPID
    // 2. Invoke：使用DISPID调用，类型为 DISPATCH_PROPERTYPUT
    // 
    // DISPID是什么？
    // - DISPID（Dispatch Identifier）是一个整数
    // - 用于标识接口的方法或属性
    // - 比字符串名称更高效
    // --------------------------------------------------------

    // 准备VARIANT参数：false
    VARIANT visible;
    VariantInit(&visible);  // 初始化VARIANT
    visible.vt = VT_BOOL;   // 类型：布尔值
    visible.boolVal = VARIANT_FALSE;  // 值：false

    DISPID dispID;
    LPOLESTR ptName = const_cast<LPOLESTR>(L"Visible");  // 属性名

    // 获取Visible属性的DISPID
    // 参数：
    //   IID_NULL - 保留参数，必须为IID_NULL
    //   &ptName - 指向属性名字符串的指针数组
    //   1 - 名称数量
    //   LOCALE_USER_DEFAULT - 区域设置
    //   &dispID - 输出参数，接收DISPID
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (SUCCEEDED(hr)) {
        // 准备DISPPARAMS：参数传递结构
        DISPPARAMS dp = { NULL, NULL, 0, 0 };
        dp.cArgs = 1;           // 参数数量
        dp.rgvarg = &visible;   // 参数数组（VARIANT数组）

        // 调用Invoke设置属性
        // 参数：
        //   dispID - 方法/属性的DISPID
        //   IID_NULL - 保留参数
        //   LOCALE_USER_DEFAULT - 区域设置
        //   DISPATCH_PROPERTYPUT - 调用类型（设置属性）
        //   &dp - 参数结构
        //   NULL - 输出参数（设置属性不需要返回值）
        //   NULL - 异常信息
        //   NULL - 参数错误
        pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dp, NULL, NULL, NULL);
    }

    // --------------------------------------------------------
    // 步骤3：获取Workbooks集合
    // --------------------------------------------------------
    // Excel对象模型：
    // Application -> Workbooks -> Workbook -> Worksheets -> Worksheet -> Range
    // 
    // 我们需要获取 Application.Workbooks 属性
    // 
    // 如何通过IDispatch获取属性？
    // 1. GetIDsOfNames：根据属性名 "Workbooks" 获取DISPID
    // 2. Invoke：使用DISPID调用，类型为 DISPATCH_PROPERTYGET
    // 
    // 调用类型：
    // - DISPATCH_METHOD：调用方法
    // - DISPATCH_PROPERTYGET：获取属性
    // - DISPATCH_PROPERTYPUT：设置属性
    // --------------------------------------------------------

    ptName = const_cast<LPOLESTR>(L"Workbooks");  // 属性名
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Workbooks property. HRESULT: " << hr << std::endl;
        pExcelApp->Release();  // 释放已创建的对象
        return false;
    }

    VARIANT result;
    VariantInit(&result);

    // 准备DISPPARAMS：无参数
    DISPPARAMS dpNoArgs = { NULL, NULL, 0, 0 };

    // 调用Invoke获取Workbooks属性
    // 注意：调用类型是 DISPATCH_PROPERTYGET
    hr = pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &result, NULL, NULL);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Workbooks collection. HRESULT: " << hr << std::endl;
        pExcelApp->Release();
        return false;
    }
    // result.pdispVal 包含返回的IDispatch指针
    pWorkbooks = result.pdispVal;

    // --------------------------------------------------------
    // 步骤4：打开指定的Excel文件
    // --------------------------------------------------------
    // 需要调用 Workbooks.Open 方法
    // 
    // 如何通过IDispatch调用方法？
    // 1. GetIDsOfNames：根据方法名 "Open" 获取DISPID
    // 2. 准备参数（VARIANT数组）
    // 3. 准备DISPPARAMS
    // 4. Invoke：使用DISPID调用，类型为 DISPATCH_METHOD
    // 
    // 参数顺序注意：
    // COM方法的参数在rgvarg数组中是反向的
    // 例如：方法 Open(FileName)
    // rgvarg[0] = FileName（第一个参数在数组的最后位置）
    // 
    // 为什么要反向？
    // - 这是IDispatch接口的设计
    // - 支持可选参数和命名参数
    // - 命名参数放在数组前面
    // --------------------------------------------------------

    ptName = const_cast<LPOLESTR>(L"Open");  // 方法名
    hr = pWorkbooks->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Open method. HRESULT: " << hr << std::endl;
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    // 准备参数：文件名
    VARIANT filename;
    VariantInit(&filename);
    filename.vt = VT_BSTR;  // 类型：字符串
    filename.bstrVal = SysAllocString(filePath.c_str());  // 分配BSTR字符串

    // 准备参数数组
    // 注意：参数顺序是反向的
    // args[0] = 第一个参数（文件名）
    VARIANT args[1];
    args[0] = filename;

    // 准备DISPPARAMS
    DISPPARAMS dpOpen;
    dpOpen.cArgs = 1;           // 参数数量
    dpOpen.rgvarg = args;       // 参数数组
    dpOpen.cNamedArgs = 0;      // 命名参数数量
    dpOpen.rgdispidNamedArgs = NULL;  // 命名参数的DISPID数组

    // 调用Invoke执行Open方法
    // 注意：调用类型是 DISPATCH_METHOD
    VariantInit(&result);
    hr = pWorkbooks->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpOpen, &result, NULL, NULL);
    
    // 释放BSTR字符串
    // SysAllocString分配的内存必须用SysFreeString释放
    SysFreeString(filename.bstrVal);

    if (FAILED(hr)) {
        std::wcerr << L"Failed to open file. HRESULT: " << hr << std::endl;
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }
    // result.pdispVal 包含返回的Workbook指针
    pWorkbook = result.pdispVal;

    // --------------------------------------------------------
    // 步骤5：获取Worksheets集合
    // --------------------------------------------------------
    // 类似步骤3，获取 Workbook.Worksheets 属性
    // --------------------------------------------------------

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

    // --------------------------------------------------------
    // 步骤6：获取第一个工作表
    // --------------------------------------------------------
    // 需要调用 Worksheets.Item(1)
    // Excel的索引从1开始，不是从0开始！
    // 
    // Item是参数化属性
    // 调用方式类似方法，但使用 DISPATCH_PROPERTYGET
    // --------------------------------------------------------

    // 准备参数：工作表索引（1）
    VARIANT sheetIndex;
    VariantInit(&sheetIndex);
    sheetIndex.vt = VT_I4;  // 类型：32位整数
    sheetIndex.lVal = 1;     // 值：1（第一个工作表）

    ptName = const_cast<LPOLESTR>(L"Item");  // 属性名
    hr = pWorksheets->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Item method. HRESULT: " << hr << std::endl;
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    // 准备参数数组
    VARIANT argsItem[1];
    argsItem[0] = sheetIndex;

    // 准备DISPPARAMS
    DISPPARAMS dpItem;
    dpItem.cArgs = 1;
    dpItem.rgvarg = argsItem;
    dpItem.cNamedArgs = 0;
    dpItem.rgdispidNamedArgs = NULL;

    // 调用Invoke获取Item属性
    // 注意：参数化属性使用 DISPATCH_PROPERTYGET
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

    // --------------------------------------------------------
    // 步骤7：获取UsedRange
    // --------------------------------------------------------
    // UsedRange表示工作表中已使用的单元格区域
    // 比直接获取所有单元格更高效
    // 
    // 获取 Worksheet.UsedRange 属性
    // --------------------------------------------------------

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

    // --------------------------------------------------------
    // 步骤8：获取单元格数据
    // --------------------------------------------------------
    // 获取 Range.Value 属性
    // 
    // 关键：Range.Value返回的是SAFEARRAY
    // - 如果是单个单元格，返回对应类型的VARIANT
    // - 如果是多个单元格，返回VT_ARRAY | VT_VARIANT类型的SAFEARRAY
    // 
    // SAFEARRAY是什么？
    // - SAFEARRAY是COM中的安全数组类型
    // - 包含边界信息（下界、上界）
    // - 支持多维数组
    // - 用于跨进程数据传递
    // --------------------------------------------------------

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

    // --------------------------------------------------------
    // 步骤9：检查返回数据类型（关键修复！）
    // --------------------------------------------------------
    // 关键修复点！
    // 
    // 错误方式（之前的bug）：
    //   if (varResult.vt != VT_ARRAY)
    //   
    // 错误原因：
    // Excel返回的不是单纯的 VT_ARRAY，而是 VT_ARRAY | VT_VARIANT
    // VT_ARRAY = 0x2000 (8192)
    // VT_VARIANT = 0x000C (12)
    // 组合值 = 0x200C (8204)
    // 
    // 为什么要用位运算？
    // - VARIANT的vt字段是一个位标志组合
    // - VT_ARRAY是其中一个标志位
    // - 需要用位运算来检查是否包含这个标志
    // 
    // 正确方式：
    //   if ((varResult.vt & VT_ARRAY) != VT_ARRAY)
    //   
    // 位运算说明：
    // & 是按位与运算
    // 如果 varResult.vt 包含 VT_ARRAY 标志
    // 那么 (varResult.vt & VT_ARRAY) == VT_ARRAY
    // --------------------------------------------------------

    // 正确的VT_ARRAY类型检查（使用位运算）
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

    // --------------------------------------------------------
    // 步骤10：获取SAFEARRAY的边界信息
    // --------------------------------------------------------
    // SAFEARRAY包含边界信息
    // - SafeArrayGetLBound：获取下界
    // - SafeArrayGetUBound：获取上界
    // 
    // 注意：
    // - Excel的SAFEARRAY是二维数组
    // - 第一维是行，第二维是列
    // - 索引从1开始，不是从0开始
    // --------------------------------------------------------

    // 获取SAFEARRAY指针
    SAFEARRAY* pSafeArray = varResult.parray;
    long lBound, uBound;

    // 获取第一维（行）的下界和上界
    // 参数：
    //   pSafeArray - SAFEARRAY指针
    //   1 - 维度（1表示第一维）
    //   &lBound - 输出参数，接收下界
    SafeArrayGetLBound(pSafeArray, 1, &lBound);
    SafeArrayGetUBound(pSafeArray, 1, &uBound);

    // 计算行数
    long rowCount = uBound - lBound + 1;

    // --------------------------------------------------------
    // 步骤11：遍历数组，读取数据
    // --------------------------------------------------------
    // 从第2行开始读取（跳过表头）
    // 循环条件：row = lBound + 1 到 uBound
    // 
    // SafeArrayGetElement：获取数组元素
    // 参数：
    //   pSafeArray - SAFEARRAY指针
    //   indices - 索引数组（行索引，列索引）
    //   &cellValue - 输出参数，接收元素值
    // 
    // 注意：
    // - indices[0] = 行索引
    // - indices[1] = 列索引
    // - 索引从1开始
    // --------------------------------------------------------

    // 从第2行开始读取（跳过表头）
    // lBound 是第一行（表头行）
    // lBound + 1 是第二行（第一行数据）
    for (long row = lBound + 1; row <= uBound; row++) {
        Participant participant;
        VARIANT cellValue;

        // ----------------------------------------------------
        // 读取第1列：男生编号
        // ----------------------------------------------------
        long indices[2] = { row, 1 };  // 行=row，列=1
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        
        // 处理不同类型的值
        if (cellValue.vt == VT_BSTR) {
            // 字符串类型（文本格式）
            participant.maleId = cellValue.bstrVal;
        }
        else if (cellValue.vt == VT_I4) {
            // 32位整数类型
            participant.maleId = std::to_wstring(cellValue.lVal);
        }
        else if (cellValue.vt == VT_R8) {
            // 64位浮点数类型（Excel默认数字格式）
            participant.maleId = std::to_wstring((long long)cellValue.dblVal);
        }
        VariantClear(&cellValue);  // 清理VARIANT

        // ----------------------------------------------------
        // 读取第2列：男生姓名
        // ----------------------------------------------------
        indices[1] = 2;  // 列=2
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_BSTR) {
            participant.maleName = cellValue.bstrVal;
        }
        VariantClear(&cellValue);

        // ----------------------------------------------------
        // 读取第3列：女生编号
        // ----------------------------------------------------
        indices[1] = 3;  // 列=3
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

        // ----------------------------------------------------
        // 读取第4列：女生姓名
        // ----------------------------------------------------
        indices[1] = 4;  // 列=4
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_BSTR) {
            participant.femaleName = cellValue.bstrVal;
        }
        VariantClear(&cellValue);

        // ----------------------------------------------------
        // 提取组号
        // ----------------------------------------------------
        participant.maleGroupNumber = ExtractGroupNumber(participant.maleId);
        participant.femaleGroupNumber = ExtractGroupNumber(participant.femaleId);

        // ----------------------------------------------------
        // 添加到结果列表
        // ----------------------------------------------------
        // 过滤条件：男生姓名或女生姓名至少有一个不为空
        // 这是为了跳过空行或无效行
        if (!participant.maleName.empty() || !participant.femaleName.empty()) {
            participants.push_back(participant);
        }
    }

    // --------------------------------------------------------
    // 步骤12：释放COM资源（顺序很重要！）
    // --------------------------------------------------------
    // COM对象的引用计数机制：
    // - 每个COM对象维护一个引用计数
    // - AddRef：增加引用计数
    // - Release：减少引用计数，计数为0时销毁对象
    // 
    // 释放顺序：
    // 必须按照创建的逆序释放
    // 创建顺序：pExcelApp -> pWorkbooks -> pWorkbook -> pWorksheets -> pWorksheet -> pRange
    // 释放顺序：pRange -> pWorksheet -> pWorksheets -> pWorkbook -> pWorkbooks -> pExcelApp
    // 
    // 为什么逆序释放？
    // - 子对象依赖父对象
    // - 父对象在子对象之前释放可能导致访问冲突
    // - 类似于栈的后进先出
    // --------------------------------------------------------

    // 清理VARIANT
    VariantClear(&varResult);

    // 按逆序释放COM对象
    pRange->Release();
    pWorksheet->Release();
    pWorksheets->Release();

    // --------------------------------------------------------
    // 步骤13：关闭工作簿
    // --------------------------------------------------------
    // 调用 Workbook.Close 方法
    // 
    // Close方法参数：
    //   SaveChanges - 是否保存更改
    //   我们只是读取数据，不需要保存，所以设置为 VARIANT_FALSE
    // --------------------------------------------------------

    ptName = const_cast<LPOLESTR>(L"Close");
    hr = pWorkbook->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (SUCCEEDED(hr)) {
        // 准备参数：不保存更改
        VARIANT saveChanges;
        VariantInit(&saveChanges);
        saveChanges.vt = VT_BOOL;
        saveChanges.boolVal = VARIANT_FALSE;  // false = 不保存

        // 准备参数数组
        VARIANT argsClose[1];
        argsClose[0] = saveChanges;

        // 准备DISPPARAMS
        DISPPARAMS dpClose;
        dpClose.cArgs = 1;
        dpClose.rgvarg = argsClose;
        dpClose.cNamedArgs = 0;
        dpClose.rgdispidNamedArgs = NULL;

        // 调用Close方法
        pWorkbook->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpClose, NULL, NULL, NULL);
    }

    // 释放Workbook和Workbooks
    pWorkbook->Release();
    pWorkbooks->Release();

    // --------------------------------------------------------
    // 步骤14：退出Excel
    // --------------------------------------------------------
    // 调用 Application.Quit 方法
    // 
    // 注意：
    // - Quit方法退出Excel进程
    // - 必须确保所有工作簿都已关闭
    // - 否则Excel可能不会真正退出
    // --------------------------------------------------------

    ptName = const_cast<LPOLESTR>(L"Quit");
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (SUCCEEDED(hr)) {
        // 调用Quit方法（无参数）
        pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpNoArgs, NULL, NULL, NULL);
    }

    // 最后释放Application对象
    pExcelApp->Release();

    return true;
}

// ============================================================
// 读取成绩清单Excel文件
// ============================================================
// 参数：
//   filePath - Excel文件路径
//   scoreEntries - 输出参数，存储读取到的成绩条目列表
// 
// 返回值：
//   true  - 读取成功
//   false - 读取失败
// 
// Excel文件格式要求：
//   第1列：名次（如 1, 2, 3）
//   第2列：组别（如 "13组", "22组"）
//   第3列：成绩时间（如 "0:37:06"）
//   第1行：表头（会自动跳过，从第2行开始读取数据）
// 
// 与 ReadRegistrationInfo 的主要区别：
// 1. 列映射不同：
//    - 第1列 -> 名次（rank，需要处理多种类型）
//    - 第2列 -> 组别（group）
//    - 第3列 -> 成绩时间（time，特殊处理：VT_DATE, VT_R8）
// 
// 2. 数据类型处理更复杂：
//    名次可能的类型：
//    - VT_I4：32位整数
//    - VT_R8：64位浮点数
//    - VT_BSTR：字符串（可能包含非数字字符）
//    
//    成绩时间可能的类型：
//    - VT_BSTR：字符串格式（如 "0:37:06"）
//    - VT_DATE：日期时间类型（Excel时间格式）
//    - VT_R8：浮点数（Excel时间序列值，0=0:00:00, 1=24:00:00）
// 
// 3. 过滤条件不同：
//    - 名次 > 0 才添加到结果列表
//    - 跳过无效行（如表头、空行、格式错误的行）
// 
// 特殊说明：Excel中的时间格式
// Excel使用浮点数表示时间：
// - 整数部分：日期（从1900/1/1开始的天数）
// - 小数部分：时间（一天中的比例）
// 
// 例如：
// - 0.5 = 12:00:00（中午）
// - 0.25 = 06:00:00（早上6点）
// - 0.025694... = 00:37:06（37分6秒）
// 
// 转换公式：
// - 小时 = timeVal * 24
// - 分钟 = (timeVal * 24 - 小时) * 60
// - 秒 = ((timeVal * 24 - 小时) * 60 - 分钟) * 60
// 
// 两种时间格式的处理：
// 1. VT_DATE：使用 VariantTimeToSystemTime 转换为 SYSTEMTIME
// 2. VT_R8：手动计算小时、分钟、秒
// ============================================================
bool ExcelReader::ReadScoreList(const std::wstring& filePath, std::vector<ScoreEntry>& scoreEntries) {
    // 声明COM接口指针
    IDispatch* pExcelApp = NULL;      // Excel.Application对象
    IDispatch* pWorkbooks = NULL;      // Workbooks集合
    IDispatch* pWorkbook = NULL;       // 打开的Workbook
    IDispatch* pWorksheets = NULL;     // Worksheets集合
    IDispatch* pWorksheet = NULL;      // 第一个Worksheet
    IDispatch* pRange = NULL;          // UsedRange区域

    // --------------------------------------------------------
    // 步骤1-8：与ReadRegistrationInfo相同
    // 创建Excel实例、设置不可见、打开文件、获取工作表、获取UsedRange、获取Value
    // --------------------------------------------------------

    // 创建Excel.Application实例
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

    // 设置Excel不可见
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

    // 获取Workbooks集合
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

    // 打开文件
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

    // 获取Worksheets集合
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

    // 获取第一个工作表
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

    // 获取UsedRange
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

    // 获取单元格数据（Range.Value）
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
    // 与ReadRegistrationInfo相同的修复
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

    // 获取SAFEARRAY边界信息
    SAFEARRAY* pSafeArray = varResult.parray;
    long lBound, uBound;
    SafeArrayGetLBound(pSafeArray, 1, &lBound);
    SafeArrayGetUBound(pSafeArray, 1, &uBound);

    // 从第2行开始读取（跳过表头）
    for (long row = lBound + 1; row <= uBound; row++) {
        ScoreEntry entry;
        VARIANT cellValue;

        // ----------------------------------------------------
        // 读取第1列：名次
        // ----------------------------------------------------
        // 名次可能的类型：
        // - VT_I4：32位整数
        // - VT_R8：64位浮点数
        // - VT_BSTR：字符串（需要尝试转换为整数）
        // ----------------------------------------------------
        long indices[2] = { row, 1 };
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        
        if (cellValue.vt == VT_I4) {
            // 32位整数类型
            entry.rank = cellValue.lVal;
        }
        else if (cellValue.vt == VT_R8) {
            // 64位浮点数类型
            entry.rank = (long)cellValue.dblVal;
        }
        else if (cellValue.vt == VT_BSTR) {
            // 字符串类型
            // 尝试将字符串转换为整数
            try {
                entry.rank = std::stoi(cellValue.bstrVal);
            }
            catch (...) {
                // 转换失败，设置为0
                // 后续会过滤掉 rank <= 0 的行
                entry.rank = 0;
            }
        }
        VariantClear(&cellValue);

        // ----------------------------------------------------
        // 读取第2列：组别
        // ----------------------------------------------------
        indices[1] = 2;
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_BSTR) {
            // 字符串类型（文本格式）
            entry.group = cellValue.bstrVal;
        }
        else if (cellValue.vt == VT_I4) {
            // 32位整数类型
            // 转换为字符串 + "组"
            entry.group = std::to_wstring(cellValue.lVal) + L"zu";
        }
        else if (cellValue.vt == VT_R8) {
            // 64位浮点数类型
            entry.group = std::to_wstring((long)cellValue.dblVal) + L"zu";
        }
        VariantClear(&cellValue);

        // ----------------------------------------------------
        // 读取第3列：成绩时间
        // ----------------------------------------------------
        // 成绩时间可能的类型：
        // - VT_BSTR：字符串格式（如 "0:37:06"）
        // - VT_DATE：日期时间类型
        // - VT_R8：浮点数（Excel时间序列值）
        // 
        // 特殊说明：Excel中的时间格式
        // Excel使用浮点数表示时间：
        // - 0 = 0:00:00 (午夜)
        // - 0.5 = 12:00:00 (中午)
        // - 1 = 24:00:00 (第二天午夜)
        // 
        // 例如：
        // 0.025694... = 00:37:06
        // 计算：
        // 0.025694 * 24 = 0.616656 小时
        // 0.616656 * 60 = 36.99936 分钟
        // 0.99936 * 60 = 59.9616 秒
        // 约等于 0小时37分0秒（四舍五入）
        // ----------------------------------------------------
        indices[1] = 3;
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        
        if (cellValue.vt == VT_BSTR) {
            // 字符串类型（文本格式）
            entry.time = cellValue.bstrVal;
        }
        else if (cellValue.vt == VT_DATE) {
            // 日期时间类型
            // 使用 VariantTimeToSystemTime 转换为 SYSTEMTIME
            SYSTEMTIME st;
            VariantTimeToSystemTime(cellValue.date, &st);
            
            // 格式化为 "小时:分钟:秒"
            wchar_t buffer[32];
            swprintf_s(buffer, L"%d:%02d:%02d", st.wHour, st.wMinute, st.wSecond);
            entry.time = buffer;
        }
        else if (cellValue.vt == VT_R8) {
            // 浮点数类型（Excel时间序列值）
            // 手动计算小时、分钟、秒
            double timeVal = cellValue.dblVal;
            
            // 计算小时：timeVal * 24
            int hours = (int)(timeVal * 24);
            
            // 计算分钟：(小数部分) * 60
            int minutes = (int)((timeVal * 24 - hours) * 60);
            
            // 计算秒：(小数部分) * 60
            int seconds = (int)(((timeVal * 24 - hours) * 60 - minutes) * 60);
            
            // 格式化为 "小时:分钟:秒"
            wchar_t buffer[32];
            swprintf_s(buffer, L"%d:%02d:%02d", hours, minutes, seconds);
            entry.time = buffer;
        }
        VariantClear(&cellValue);

        // ----------------------------------------------------
        // 从组别中提取组号
        // ----------------------------------------------------
        entry.groupNumber = ExtractGroupNumber(entry.group);

        // ----------------------------------------------------
        // 添加到结果列表
        // ----------------------------------------------------
        // 过滤条件：名次 > 0
        // 这是为了跳过无效行（如表头、空行、格式错误的行）
        if (entry.rank > 0) {
            scoreEntries.push_back(entry);
        }
    }

    // --------------------------------------------------------
    // 释放资源（与ReadRegistrationInfo相同）
    // --------------------------------------------------------
    VariantClear(&varResult);
    pRange->Release();
    pWorksheet->Release();
    pWorksheets->Release();

    // 关闭工作簿
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

    // 退出Excel
    ptName = const_cast<LPOLESTR>(L"Quit");
    hr = pExcelApp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (SUCCEEDED(hr)) {
        pExcelApp->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpNoArgs, NULL, NULL, NULL);
    }

    pExcelApp->Release();

    return true;
}
