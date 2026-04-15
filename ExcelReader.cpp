#pragma execution_character_set("utf-8")

#include "ExcelReader.h"
#include <windows.h>
#include <comutil.h>
#include <iostream>
#include <sstream>
#include <regex>
#include <algorithm>

#pragma comment(lib, "comsuppw.lib")

ExcelReader::ExcelReader() {
    CoInitialize(NULL);
}

ExcelReader::~ExcelReader() {
    CoUninitialize();
}

int ExcelReader::ExtractGroupNumber(const std::wstring& id) {
    std::wregex regex(L"(\\d+)");
    std::wsmatch match;
    if (std::regex_search(id, match, regex)) {
        return std::stoi(match[1].str());
    }
    return -1;
}

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
        std::wcerr << L"Failed to open file: " << filePath << L". HRESULT: " << hr << std::endl;
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

    ptName = const_cast<LPOLESTR>(L"Rows");
    hr = pRange->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (SUCCEEDED(hr)) {
        VARIANT rowsResult;
        VariantInit(&rowsResult);
        hr = pRange->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &rowsResult, NULL, NULL);
        if (SUCCEEDED(hr) && rowsResult.vt == VT_DISPATCH) {
            IDispatch* pRows = rowsResult.pdispVal;
            LPOLESTR countName = const_cast<LPOLESTR>(L"Count");
            DISPID countDispID;
            hr = pRows->GetIDsOfNames(IID_NULL, &countName, 1, LOCALE_USER_DEFAULT, &countDispID);
            if (SUCCEEDED(hr)) {
                VARIANT countResult;
                VariantInit(&countResult);
                hr = pRows->Invoke(countDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &countResult, NULL, NULL);
                if (SUCCEEDED(hr)) {
                    long rowCount = 0;
                    if (countResult.vt == VT_I4) {
                        rowCount = countResult.lVal;
                    }
                    else if (countResult.vt == VT_I2) {
                        rowCount = countResult.iVal;
                    }
                    std::wcout << L"   Worksheet has " << rowCount << L" rows" << std::endl;
                    if (rowCount <= 1) {
                        std::wcerr << L"Warning: Worksheet appears to be empty or has only header row" << std::endl;
                    }
                }
                VariantClear(&countResult);
            }
            pRows->Release();
        }
        VariantClear(&rowsResult);
    }

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
    
    std::wcout << L"   Value property HRESULT: " << hr << std::endl;
    std::wcout << L"   Value property type: " << varResult.vt << std::endl;
    
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
        std::wcerr << L"This may happen if the worksheet is empty or has only one cell." << std::endl;
        
        if (varResult.vt == VT_EMPTY || varResult.vt == VT_NULL) {
            std::wcerr << L"Worksheet appears to be empty." << std::endl;
        }
        else if (varResult.vt == VT_BSTR) {
            std::wcerr << L"Single cell value: " << varResult.bstrVal << std::endl;
        }
        
        VariantClear(&varResult);
        pRange->Release();
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    std::wcout << L"   Array type detected: " << varResult.vt << std::endl;
    SAFEARRAY* pSafeArray = varResult.parray;
    long lBound, uBound;
    SafeArrayGetLBound(pSafeArray, 1, &lBound);
    SafeArrayGetUBound(pSafeArray, 1, &uBound);

    long rowCount = uBound - lBound + 1;
    std::wcout << L"   Array row count: " << rowCount << std::endl;

    for (long row = lBound + 1; row <= uBound; row++) {
        Participant participant;
        VARIANT cellValue;

        long indices[2] = { row, 1 };
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

        indices[1] = 2;
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_BSTR) {
            participant.maleName = cellValue.bstrVal;
        }
        VariantClear(&cellValue);

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

        indices[1] = 4;
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_BSTR) {
            participant.femaleName = cellValue.bstrVal;
        }
        VariantClear(&cellValue);

        participant.groupNumber = ExtractGroupNumber(participant.maleId);

        if (!participant.maleName.empty() || !participant.femaleName.empty()) {
            participants.push_back(participant);
        }
    }

    std::wcout << L"   Successfully read " << participants.size() << L" registration entries" << std::endl;

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
        std::wcerr << L"Failed to open file: " << filePath << L". HRESULT: " << hr << std::endl;
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

    ptName = const_cast<LPOLESTR>(L"Rows");
    hr = pRange->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (SUCCEEDED(hr)) {
        VARIANT rowsResult;
        VariantInit(&rowsResult);
        hr = pRange->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &rowsResult, NULL, NULL);
        if (SUCCEEDED(hr) && rowsResult.vt == VT_DISPATCH) {
            IDispatch* pRows = rowsResult.pdispVal;
            LPOLESTR countName = const_cast<LPOLESTR>(L"Count");
            DISPID countDispID;
            hr = pRows->GetIDsOfNames(IID_NULL, &countName, 1, LOCALE_USER_DEFAULT, &countDispID);
            if (SUCCEEDED(hr)) {
                VARIANT countResult;
                VariantInit(&countResult);
                hr = pRows->Invoke(countDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpNoArgs, &countResult, NULL, NULL);
                if (SUCCEEDED(hr)) {
                    long rowCount = 0;
                    if (countResult.vt == VT_I4) {
                        rowCount = countResult.lVal;
                    }
                    else if (countResult.vt == VT_I2) {
                        rowCount = countResult.iVal;
                    }
                    std::wcout << L"   Worksheet has " << rowCount << L" rows" << std::endl;
                    if (rowCount <= 1) {
                        std::wcerr << L"Warning: Worksheet appears to be empty or has only header row" << std::endl;
                    }
                }
                VariantClear(&countResult);
            }
            pRows->Release();
        }
        VariantClear(&rowsResult);
    }

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
    
    std::wcout << L"   Value property HRESULT: " << hr << std::endl;
    std::wcout << L"   Value property type: " << varResult.vt << std::endl;
    
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
        std::wcerr << L"This may happen if the worksheet is empty or has only one cell." << std::endl;
        
        if (varResult.vt == VT_EMPTY || varResult.vt == VT_NULL) {
            std::wcerr << L"Worksheet appears to be empty." << std::endl;
        }
        else if (varResult.vt == VT_BSTR) {
            std::wcerr << L"Single cell value: " << varResult.bstrVal << std::endl;
        }
        
        VariantClear(&varResult);
        pRange->Release();
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    std::wcout << L"   Array type detected: " << varResult.vt << std::endl;
    SAFEARRAY* pSafeArray = varResult.parray;
    long lBound, uBound;
    SafeArrayGetLBound(pSafeArray, 1, &lBound);
    SafeArrayGetUBound(pSafeArray, 1, &uBound);

    long rowCount = uBound - lBound + 1;
    std::wcout << L"   Array row count: " << rowCount << std::endl;

    for (long row = lBound + 1; row <= uBound; row++) {
        ScoreEntry entry;
        VARIANT cellValue;

        long indices[2] = { row, 1 };
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_I4) {
            entry.rank = cellValue.lVal;
        }
        else if (cellValue.vt == VT_R8) {
            entry.rank = (long)cellValue.dblVal;
        }
        else if (cellValue.vt == VT_BSTR) {
            try {
                entry.rank = std::stoi(cellValue.bstrVal);
            }
            catch (...) {
                entry.rank = 0;
            }
        }
        VariantClear(&cellValue);

        indices[1] = 2;
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_BSTR) {
            entry.group = cellValue.bstrVal;
        }
        else if (cellValue.vt == VT_I4) {
            entry.group = std::to_wstring(cellValue.lVal) + L"zu";
        }
        else if (cellValue.vt == VT_R8) {
            entry.group = std::to_wstring((long)cellValue.dblVal) + L"zu";
        }
        VariantClear(&cellValue);

        indices[1] = 3;
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_BSTR) {
            entry.time = cellValue.bstrVal;
        }
        else if (cellValue.vt == VT_DATE) {
            SYSTEMTIME st;
            VariantTimeToSystemTime(cellValue.date, &st);
            wchar_t buffer[32];
            swprintf_s(buffer, L"%d:%02d:%02d", st.wHour, st.wMinute, st.wSecond);
            entry.time = buffer;
        }
        else if (cellValue.vt == VT_R8) {
            double timeVal = cellValue.dblVal;
            int hours = (int)(timeVal * 24);
            int minutes = (int)((timeVal * 24 - hours) * 60);
            int seconds = (int)(((timeVal * 24 - hours) * 60 - minutes) * 60);
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

    std::wcout << L"   Successfully read " << scoreEntries.size() << L" score entries" << std::endl;

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
