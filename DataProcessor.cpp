#pragma execution_character_set("utf-8")

#include "DataProcessor.h"
#include <windows.h>
#include <comutil.h>
#include <iostream>
#include <sstream>
#include <iomanip>
#include <algorithm>

#pragma comment(lib, "comsuppw.lib")

DataProcessor::DataProcessor() {
}

DataProcessor::~DataProcessor() {
}

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

    std::unordered_map<int, std::wstring> maleMap;
    std::unordered_map<int, std::wstring> femaleMap;

    for (const auto& participant : participants) {
        if (participant.maleGroupNumber >= 0 && !participant.maleName.empty()) {
            maleMap[participant.maleGroupNumber] = participant.maleName;
        }
        if (participant.femaleGroupNumber >= 0 && !participant.femaleName.empty()) {
            femaleMap[participant.femaleGroupNumber] = participant.femaleName;
        }
    }

    for (const auto& scoreEntry : scoreEntries) {
        ResultEntry result;
        result.rank = scoreEntry.rank;
        result.group = scoreEntry.group;
        result.time = scoreEntry.time;

        if (scoreEntry.groupNumber < 0) {
            std::wcerr << L"Warning: Invalid group number for rank " << scoreEntry.rank << std::endl;
            result.names = L"Invalid Group";
        }
        else {
            std::wstring maleName;
            std::wstring femaleName;

            auto maleIt = maleMap.find(scoreEntry.groupNumber);
            if (maleIt != maleMap.end()) {
                maleName = maleIt->second;
            }

            auto femaleIt = femaleMap.find(scoreEntry.groupNumber);
            if (femaleIt != femaleMap.end()) {
                femaleName = femaleIt->second;
            }

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

    return true;
}

bool DataProcessor::ExportResults(const std::wstring& filePath, const std::vector<ResultEntry>& results) {
    if (results.empty()) {
        std::wcerr << L"Warning: No results to export" << std::endl;
    }

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

    int fileFormatValue = 56;
    std::wstring lowerPath = filePath;
    std::transform(lowerPath.begin(), lowerPath.end(), lowerPath.begin(), ::towlower);
    
    if (lowerPath.length() >= 5 && lowerPath.substr(lowerPath.length() - 5) == L".xlsx") {
        fileFormatValue = 51;
    }
    else if (lowerPath.length() >= 4 && lowerPath.substr(lowerPath.length() - 4) == L".xls") {
        fileFormatValue = 56;
    }
    else {
        fileFormatValue = 56;
    }

    VARIANT fileFormat;
    VariantInit(&fileFormat);
    fileFormat.vt = VT_I4;
    fileFormat.lVal = fileFormatValue;

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
