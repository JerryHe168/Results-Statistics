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

    ptName = const_cast<LPOLESTR>(L"Open");
    hr = pWorkbooks->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Open method" << std::endl;
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
        std::wcerr << L"Failed to open file: " << filePath << std::endl;
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

    ptName = const_cast<LPOLESTR>(L"UsedRange");
    hr = pWorksheet->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get UsedRange property" << std::endl;
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
        std::wcerr << L"Failed to get UsedRange" << std::endl;
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
        std::wcerr << L"Failed to get Value property" << std::endl;
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
    if (FAILED(hr) || varResult.vt != VT_ARRAY) {
        std::wcerr << L"Failed to get cell data" << std::endl;
        VariantClear(&varResult);
        pRange->Release();
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    SAFEARRAY* pSafeArray = varResult.parray;
    long lBound, uBound;
    SafeArrayGetLBound(pSafeArray, 1, &lBound);
    SafeArrayGetUBound(pSafeArray, 1, &uBound);

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

    ptName = const_cast<LPOLESTR>(L"Open");
    hr = pWorkbooks->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get Open method" << std::endl;
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
        std::wcerr << L"Failed to open file: " << filePath << std::endl;
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

    ptName = const_cast<LPOLESTR>(L"UsedRange");
    hr = pWorksheet->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get UsedRange property" << std::endl;
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
        std::wcerr << L"Failed to get UsedRange" << std::endl;
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
        std::wcerr << L"Failed to get Value property" << std::endl;
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
    if (FAILED(hr) || varResult.vt != VT_ARRAY) {
        std::wcerr << L"Failed to get cell data" << std::endl;
        VariantClear(&varResult);
        pRange->Release();
        pWorksheet->Release();
        pWorksheets->Release();
        pWorkbook->Release();
        pWorkbooks->Release();
        pExcelApp->Release();
        return false;
    }

    SAFEARRAY* pSafeArray = varResult.parray;
    long lBound, uBound;
    SafeArrayGetLBound(pSafeArray, 1, &lBound);
    SafeArrayGetUBound(pSafeArray, 1, &uBound);

    for (long row = lBound + 1; row <= uBound; row++) {
        ScoreEntry entry;
        VARIANT cellValue;

        long indices[2] = { row, 1 };
        SafeArrayGetElement(pSafeArray, indices, &cellValue);
        if (cellValue.vt == VT_I4) {
            entry.rank = cellValue.lVal;
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
