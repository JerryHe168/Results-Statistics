#pragma once

#include <windows.h>
#include <string>
#include <vector>
#include <algorithm>
#include <fstream>
#include <sstream>

class FileImporter {
public:
    FileImporter();
    ~FileImporter();

    enum class FileFormat {
        Excel,
        Csv,
        Unknown
    };

    static FileFormat DetectFileFormat(const std::wstring& filePath);

    bool ImportFile(const std::wstring& filePath);

    const std::vector<std::wstring>& GetHeaders() const { return m_headers; }
    const std::vector<std::vector<std::wstring>>& GetData() const { return m_data; }

private:
    std::vector<std::wstring> m_headers;
    std::vector<std::vector<std::wstring>> m_data;

    bool ImportExcelFile(const std::wstring& filePath);
    bool ImportCsvFile(const std::wstring& filePath);

    std::vector<std::wstring> SplitCsvLine(const std::wstring& line) const;
    std::wstring Trim(const std::wstring& str) const;
    std::wstring StringToWString(const std::string& str) const;
};
