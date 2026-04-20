#include <windows.h>
#include <shellapi.h>
#include <commdlg.h>
#include <sqlext.h>

#include <algorithm>
#include <cctype>
#include <cwctype>
#include <cstring>
#include <cstdlib>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <map>
#include <regex>
#include <sstream>
#include <stdexcept>
#include <string>
#include <vector>

namespace fs = std::filesystem;

static std::string wideToUtf8(const std::wstring& value) {
    if (value.empty()) return {};
    int size = WideCharToMultiByte(CP_UTF8, 0, value.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (size <= 0) return {};
    std::string result(static_cast<size_t>(size - 1), '\0');
    WideCharToMultiByte(CP_UTF8, 0, value.c_str(), -1, result.data(), size, nullptr, nullptr);
    return result;
}

static std::wstring utf8ToWide(const std::string& value) {
    if (value.empty()) return {};
    int size = MultiByteToWideChar(CP_UTF8, 0, value.c_str(), -1, nullptr, 0);
    if (size <= 0) {
        size = MultiByteToWideChar(CP_ACP, 0, value.c_str(), -1, nullptr, 0);
        if (size <= 0) return {};
        std::wstring fallback(static_cast<size_t>(size - 1), L'\0');
        MultiByteToWideChar(CP_ACP, 0, value.c_str(), -1, fallback.data(), size);
        return fallback;
    }
    std::wstring result(static_cast<size_t>(size - 1), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, value.c_str(), -1, result.data(), size);
    return result;
}

static fs::path makePath(const std::string& value) {
    return fs::path(utf8ToWide(value));
}

static std::string pathToUtf8(const fs::path& path) {
    return wideToUtf8(path.wstring());
}

static std::vector<std::string> commandLineArgsUtf8() {
    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (!argv) return {};
    std::vector<std::string> result;
    for (int i = 0; i < argc; ++i) {
        result.push_back(wideToUtf8(argv[i]));
    }
    LocalFree(argv);
    return result;
}

struct Options {
    std::string input;
    std::string db = "postgres";
    std::string host = "localhost";
    std::string port = "5432";
    std::string dbname = "sociology_survey";
    std::string user = "postgres";
    std::string password;
    std::string table;
    std::string driver;
    std::string connstr;
    bool dryRun = false;
    bool dropExisting = false;
    bool interactive = false;
};

struct TableData {
    std::vector<std::string> columns;
    std::vector<std::vector<std::string>> rows;
};

struct NamedTable {
    std::string sourceName;
    std::string tableName;
    TableData data;
};

enum class DataType { Integer, Real, Boolean, Date, Timestamp, Text };

static std::ofstream g_log;

static void logLine(const std::string& level, const std::string& message) {
    std::ostringstream line;
    line << "[" << level << "] " << message;
    std::cout << line.str() << "\n";
    if (g_log.is_open()) {
        g_log << line.str() << "\n";
    }
}

static std::string trim(const std::string& s) {
    size_t start = 0;
    while (start < s.size() && std::isspace(static_cast<unsigned char>(s[start]))) start++;
    size_t end = s.size();
    while (end > start && std::isspace(static_cast<unsigned char>(s[end - 1]))) end--;
    return s.substr(start, end - start);
}

static std::string lower(std::string s) {
    std::transform(s.begin(), s.end(), s.begin(), [](unsigned char c) {
        return static_cast<char>(std::tolower(c));
    });
    return s;
}

static bool endsWithLower(const std::string& s, const std::string& suffix) {
    std::string a = lower(s);
    return a.size() >= suffix.size() && a.substr(a.size() - suffix.size()) == suffix;
}

static std::string stripOuterQuotes(std::string s) {
    s = trim(s);
    if (s.size() >= 2 && ((s.front() == '"' && s.back() == '"') || (s.front() == '\'' && s.back() == '\''))) {
        return s.substr(1, s.size() - 2);
    }
    return s;
}

static std::string prompt(const std::string& label, const std::string& defaultValue = "") {
    std::cout << label;
    if (!defaultValue.empty()) std::cout << " [" << defaultValue << "]";
    std::cout << ": ";
    std::string value;
    std::getline(std::cin, value);
    value = stripOuterQuotes(value);
    if (value.empty()) return defaultValue;
    return value;
}

static bool promptYesNo(const std::string& label, bool defaultValue) {
    std::string suffix = defaultValue ? "Д/н" : "д/Н";
    while (true) {
        std::string answer = lower(trim(prompt(label + " (" + suffix + ")")));
        if (answer.empty()) return defaultValue;
        if (answer == "y" || answer == "yes" || answer == "д" || answer == "да" || answer == "Д" || answer == "Да" || answer == "ДА") return true;
        if (answer == "n" || answer == "no" || answer == "н" || answer == "нет" || answer == "Н" || answer == "Нет" || answer == "НЕТ") return false;
        std::cout << "Введите y/yes/да или n/no/нет.\n";
    }
}

static std::string chooseInputFileDialog() {
    std::vector<wchar_t> fileName(32768, L'\0');
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = nullptr;
    ofn.lpstrFilter =
        L"CSV и Excel (*.csv;*.xlsx)\0*.csv;*.xlsx\0"
        L"CSV (*.csv)\0*.csv\0"
        L"Excel (*.xlsx)\0*.xlsx\0"
        L"Все файлы (*.*)\0*.*\0";
    ofn.lpstrFile = fileName.data();
    ofn.nMaxFile = static_cast<DWORD>(fileName.size());
    ofn.lpstrTitle = L"Выберите CSV или XLSX файл для загрузки";
    ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;

    if (!GetOpenFileNameW(&ofn)) return {};
    return wideToUtf8(fileName.data());
}

static std::string shellQuote(const std::string& s) {
    std::string out = "'";
    for (char c : s) {
        if (c == '\'') out += "''";
        else out += c;
    }
    out += "'";
    return out;
}

static std::wstring shellQuoteW(const std::wstring& s) {
    std::wstring out = L"'";
    for (wchar_t c : s) {
        if (c == L'\'') out += L"''";
        else out += c;
    }
    out += L"'";
    return out;
}

static std::string sanitizeIdentifier(const std::string& raw, const std::string& fallback) {
    std::wstring s = utf8ToWide(trim(raw));
    if (s.empty()) s = utf8ToWide(fallback);
    std::wstring out;
    bool lastUnderscore = false;
    for (wchar_t c : s) {
        bool asciiAlnum = (c >= L'0' && c <= L'9') || (c >= L'A' && c <= L'Z') || (c >= L'a' && c <= L'z');
        bool cyrillic = c >= 0x0400 && c <= 0x04FF;
        if (asciiAlnum || cyrillic) {
            out += static_cast<wchar_t>(std::towlower(c));
            lastUnderscore = false;
        } else if (!lastUnderscore) {
            out += L'_';
            lastUnderscore = true;
        }
    }
    while (!out.empty() && out.front() == L'_') out.erase(out.begin());
    while (!out.empty() && out.back() == L'_') out.pop_back();
    if (out.empty()) out = utf8ToWide(fallback);
    if (!out.empty() && out.front() >= L'0' && out.front() <= L'9') out = L"c_" + out;
    return wideToUtf8(out);
}

static std::vector<std::string> uniqueTableNames(const std::vector<std::string>& rawNames) {
    std::vector<std::string> result;
    std::map<std::string, int> seen;
    for (size_t i = 0; i < rawNames.size(); ++i) {
        std::string base = sanitizeIdentifier(rawNames[i], "table_" + std::to_string(i + 1));
        std::string name = base;
        int suffix = 2;
        while (seen[name] > 0) {
            name = base + "_" + std::to_string(suffix++);
        }
        seen[name]++;
        result.push_back(name);
    }
    return result;
}

static std::vector<std::string> normalizeColumns(const std::vector<std::string>& raw) {
    std::vector<std::string> result;
    std::map<std::string, int> seen;
    for (size_t i = 0; i < raw.size(); ++i) {
        std::string base = sanitizeIdentifier(raw[i], "column_" + std::to_string(i + 1));
        std::string name = base;
        int suffix = 2;
        while (seen[name] > 0) {
            name = base + "_" + std::to_string(suffix++);
        }
        seen[name]++;
        result.push_back(name);
    }
    return result;
}

static std::vector<std::string> parseCsvLine(const std::string& line, char delimiter) {
    std::vector<std::string> fields;
    std::string field;
    bool inQuotes = false;
    for (size_t i = 0; i < line.size(); ++i) {
        char c = line[i];
        if (c == '"') {
            if (inQuotes && i + 1 < line.size() && line[i + 1] == '"') {
                field += '"';
                ++i;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (c == delimiter && !inQuotes) {
            fields.push_back(field);
            field.clear();
        } else {
            field += c;
        }
    }
    fields.push_back(field);
    return fields;
}

static std::string stripTrailingSemicolons(std::string s) {
    while (!s.empty() && (s.back() == ';' || s.back() == '\r')) s.pop_back();
    return s;
}

static void replaceAll(std::string& s, const std::string& from, const std::string& to) {
    size_t pos = 0;
    while ((pos = s.find(from, pos)) != std::string::npos) {
        s.replace(pos, from.size(), to);
        pos += to.size();
    }
}

static std::string normalizeWrappedCsvRecord(const std::vector<std::string>& physicalLines) {
    std::vector<std::string> lines = physicalLines;
    for (auto& line : lines) line = stripTrailingSemicolons(line);

    // Some spreadsheet exports wrap the whole row in quotes and split quoted
    // fields across physical lines. Reconnect those accidental line-edge quotes.
    if (lines.size() > 1) {
        for (size_t i = 0; i + 1 < lines.size(); ++i) {
            if (!lines[i].empty() && lines[i].back() == '"') lines[i].pop_back();
        }
        for (size_t i = 1; i < lines.size(); ++i) {
            if (!lines[i].empty() && lines[i].front() == '"') lines[i].erase(lines[i].begin());
        }
    }

    std::string record;
    for (size_t i = 0; i < lines.size(); ++i) {
        if (i) record += '\n';
        record += lines[i];
    }
    if (record.size() >= 2 && record.front() == '"' && record.back() == '"') {
        record = record.substr(1, record.size() - 2);
        replaceAll(record, "\"\"", "\"");
    }
    return record;
}

static char detectDelimiter(const std::string& header) {
    std::vector<char> candidates = {',', ';', '\t'};
    char best = ',';
    size_t bestCount = 0;
    for (char c : candidates) {
        size_t count = static_cast<size_t>(std::count(header.begin(), header.end(), c));
        if (count > bestCount) {
            best = c;
            bestCount = count;
        }
    }
    return best;
}

static TableData readCsv(const fs::path& path) {
    std::ifstream in(path, std::ios::binary);
    if (!in) throw std::runtime_error("Не удалось открыть CSV файл: " + pathToUtf8(path));

    std::string header;
    if (!std::getline(in, header)) throw std::runtime_error("CSV файл пустой");
    if (header.size() >= 3 &&
        static_cast<unsigned char>(header[0]) == 0xEF &&
        static_cast<unsigned char>(header[1]) == 0xBB &&
        static_cast<unsigned char>(header[2]) == 0xBF) {
        header = header.substr(3);
    }
    if (!header.empty() && header.back() == '\r') header.pop_back();

    char delimiter = detectDelimiter(header);
    TableData data;
    data.columns = normalizeColumns(parseCsvLine(stripTrailingSemicolons(header), delimiter));

    std::string line;
    std::vector<std::string> pending;
    while (std::getline(in, line)) {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        pending.push_back(line);

        auto candidate = normalizeWrappedCsvRecord(pending);
        auto row = parseCsvLine(candidate, delimiter);
        if (row.size() < data.columns.size()) continue;
        if (row.size() > data.columns.size()) {
            std::vector<std::string> compact;
            compact.reserve(data.columns.size());
            for (size_t i = 0; i + 1 < data.columns.size(); ++i) compact.push_back(row[i]);
            std::string tail;
            for (size_t i = data.columns.size() - 1; i < row.size(); ++i) {
                if (!tail.empty()) tail += delimiter;
                tail += row[i];
            }
            compact.push_back(tail);
            row = compact;
        }
        row.resize(data.columns.size());
        bool any = false;
        for (const auto& v : row) {
            if (!trim(v).empty()) {
                any = true;
                break;
            }
        }
        if (any) data.rows.push_back(row);
        pending.clear();
    }
    if (!pending.empty()) {
        auto candidate = normalizeWrappedCsvRecord(pending);
        auto row = parseCsvLine(candidate, delimiter);
        row.resize(data.columns.size());
        bool any = false;
        for (const auto& v : row) {
            if (!trim(v).empty()) {
                any = true;
                break;
            }
        }
        if (any) data.rows.push_back(row);
    }

    logLine("ИНФО", "Разделитель CSV: " + std::string(delimiter == '\t' ? "\\t" : std::string(1, delimiter)));
    return data;
}

static std::string readFile(const fs::path& path) {
    std::ifstream in(path, std::ios::binary);
    if (!in) throw std::runtime_error("Не удалось прочитать файл: " + pathToUtf8(path));
    std::ostringstream ss;
    ss << in.rdbuf();
    return ss.str();
}

static std::string xmlDecode(std::string s) {
    struct Pair { const char* from; const char* to; };
    const Pair pairs[] = {
        {"&amp;", "&"}, {"&lt;", "<"}, {"&gt;", ">"},
        {"&quot;", "\""}, {"&apos;", "'"}
    };
    for (const auto& p : pairs) {
        size_t pos = 0;
        while ((pos = s.find(p.from, pos)) != std::string::npos) {
            s.replace(pos, std::strlen(p.from), p.to);
            pos += std::strlen(p.to);
        }
    }
    return s;
}

static std::string xmlAttr(const std::string& attrs, const std::string& name) {
    std::regex re(name + "=\"([^\"]*)\"", std::regex::icase);
    std::smatch m;
    if (std::regex_search(attrs, m, re)) return xmlDecode(m[1].str());
    return {};
}

static int excelColumnIndex(const std::string& cellRef) {
    int n = 0;
    for (char c : cellRef) {
        if (!std::isalpha(static_cast<unsigned char>(c))) break;
        n = n * 26 + (std::toupper(static_cast<unsigned char>(c)) - 'A' + 1);
    }
    return n - 1;
}

static std::vector<std::string> readSharedStrings(const fs::path& root) {
    fs::path file = root / "xl" / "sharedStrings.xml";
    if (!fs::exists(file)) return {};
    std::string xml = readFile(file);
    std::vector<std::string> strings;
    std::regex siRe("<si[^>]*>([\\s\\S]*?)</si>", std::regex::icase);
    std::regex tRe("<t[^>]*>([\\s\\S]*?)</t>", std::regex::icase);
    auto begin = std::sregex_iterator(xml.begin(), xml.end(), siRe);
    auto end = std::sregex_iterator();
    for (auto it = begin; it != end; ++it) {
        std::string si = (*it)[1].str();
        std::string value;
        auto tb = std::sregex_iterator(si.begin(), si.end(), tRe);
        for (auto jt = tb; jt != end; ++jt) value += xmlDecode((*jt)[1].str());
        strings.push_back(value);
    }
    return strings;
}

struct XlsxSheetInfo {
    std::string name;
    std::string relId;
    fs::path path;
};

static TableData parseXlsxSheet(const fs::path& sheetPath, const std::vector<std::string>& shared) {
    std::string xml = readFile(sheetPath);
    std::vector<std::vector<std::string>> rawRows;
    std::regex rowRe("<row[^>]*>([\\s\\S]*?)</row>", std::regex::icase);
    std::regex cellRe("<c\\s+([^>]*)>([\\s\\S]*?)</c>", std::regex::icase);
    std::regex refRe("r=\"([A-Z]+[0-9]+)\"", std::regex::icase);
    std::regex valueRe("<v[^>]*>([\\s\\S]*?)</v>", std::regex::icase);
    std::regex inlineRe("<t[^>]*>([\\s\\S]*?)</t>", std::regex::icase);

    auto rowsBegin = std::sregex_iterator(xml.begin(), xml.end(), rowRe);
    auto rowsEnd = std::sregex_iterator();
    for (auto rit = rowsBegin; rit != rowsEnd; ++rit) {
        std::string rowXml = (*rit)[1].str();
        std::vector<std::string> row;
        auto cellsBegin = std::sregex_iterator(rowXml.begin(), rowXml.end(), cellRe);
        for (auto cit = cellsBegin; cit != rowsEnd; ++cit) {
            std::string attrs = (*cit)[1].str();
            std::string body = (*cit)[2].str();
            std::smatch m;
            int index = static_cast<int>(row.size());
            if (std::regex_search(attrs, m, refRe)) index = excelColumnIndex(m[1].str());
            if (index < 0) continue;
            if (static_cast<int>(row.size()) <= index) row.resize(static_cast<size_t>(index + 1));

            std::string type = xmlAttr(attrs, "t");
            std::string value;
            if (type == "inlineStr") {
                if (std::regex_search(body, m, inlineRe)) value = xmlDecode(m[1].str());
            } else if (std::regex_search(body, m, valueRe)) {
                value = xmlDecode(m[1].str());
                if (type == "s") {
                    int si = std::atoi(value.c_str());
                    if (si >= 0 && static_cast<size_t>(si) < shared.size()) value = shared[static_cast<size_t>(si)];
                }
            }
            row[static_cast<size_t>(index)] = value;
        }
        bool any = false;
        for (const auto& v : row) {
            if (!trim(v).empty()) {
                any = true;
                break;
            }
        }
        if (any) rawRows.push_back(row);
    }

    if (rawRows.empty()) return {};
    TableData data;
    data.columns = normalizeColumns(rawRows.front());
    for (size_t i = 1; i < rawRows.size(); ++i) {
        rawRows[i].resize(data.columns.size());
        data.rows.push_back(rawRows[i]);
    }
    return data;
}

static fs::path resolveXlsxTarget(const fs::path& root, const std::string& target) {
    std::string t = target;
    std::replace(t.begin(), t.end(), '\\', '/');
    if (!t.empty() && t.front() == '/') {
        t.erase(t.begin());
        return root / makePath(t);
    }
    if (t.rfind("xl/", 0) == 0) return root / makePath(t);
    return root / "xl" / makePath(t);
}

static std::vector<XlsxSheetInfo> readXlsxSheetInfos(const fs::path& root) {
    fs::path workbook = root / "xl" / "workbook.xml";
    fs::path relsFile = root / "xl" / "_rels" / "workbook.xml.rels";
    if (!fs::exists(workbook)) throw std::runtime_error("В XLSX не найден xl/workbook.xml");
    if (!fs::exists(relsFile)) throw std::runtime_error("В XLSX не найден xl/_rels/workbook.xml.rels");

    std::string workbookXml = readFile(workbook);
    std::string relsXml = readFile(relsFile);

    std::map<std::string, fs::path> relTargets;
    std::regex relRe("<Relationship\\s+([^>]*)/?\\s*>", std::regex::icase);
    auto relBegin = std::sregex_iterator(relsXml.begin(), relsXml.end(), relRe);
    auto relEnd = std::sregex_iterator();
    for (auto it = relBegin; it != relEnd; ++it) {
        std::string attrs = (*it)[1].str();
        std::string id = xmlAttr(attrs, "Id");
        std::string target = xmlAttr(attrs, "Target");
        if (!id.empty() && !target.empty()) relTargets[id] = resolveXlsxTarget(root, target);
    }

    std::vector<XlsxSheetInfo> sheets;
    std::regex sheetRe("<sheet\\s+([^>]*)/?\\s*>", std::regex::icase);
    auto sheetBegin = std::sregex_iterator(workbookXml.begin(), workbookXml.end(), sheetRe);
    auto sheetEnd = std::sregex_iterator();
    for (auto it = sheetBegin; it != sheetEnd; ++it) {
        std::string attrs = (*it)[1].str();
        XlsxSheetInfo info;
        info.name = xmlAttr(attrs, "name");
        info.relId = xmlAttr(attrs, "r:id");
        auto found = relTargets.find(info.relId);
        if (!info.name.empty() && found != relTargets.end() && fs::exists(found->second)) {
            info.path = found->second;
            sheets.push_back(info);
        }
    }

    if (sheets.empty()) throw std::runtime_error("В XLSX не найдено ни одного листа с данными");
    return sheets;
}

static std::vector<NamedTable> readXlsxTables(const fs::path& path) {
    fs::path temp = fs::temp_directory_path() / ("sql_loader_xlsx_" + std::to_string(GetCurrentProcessId()));
    if (fs::exists(temp)) fs::remove_all(temp);
    fs::create_directories(temp);
    fs::path zipCopy = temp / "workbook.zip";
    fs::copy_file(path, zipCopy, fs::copy_options::overwrite_existing);

    std::wstring command = L"powershell -NoProfile -ExecutionPolicy Bypass -Command "
        L"\"Expand-Archive -LiteralPath " + shellQuoteW(zipCopy.wstring()) +
        L" -DestinationPath " + shellQuoteW(temp.wstring()) + L" -Force\"";
    int code = _wsystem(command.c_str());
    if (code != 0) {
        fs::remove_all(temp);
        throw std::runtime_error("Не удалось распаковать XLSX. Убедитесь, что это настоящий .xlsx файл.");
    }

    try {
        std::vector<std::string> shared = readSharedStrings(temp);
        auto sheets = readXlsxSheetInfos(temp);
        std::vector<std::string> rawNames;
        for (const auto& sheet : sheets) rawNames.push_back(sheet.name);
        auto tableNames = uniqueTableNames(rawNames);

        std::vector<NamedTable> tables;
        for (size_t i = 0; i < sheets.size(); ++i) {
            TableData data = parseXlsxSheet(sheets[i].path, shared);
            if (data.columns.empty()) {
                logLine("ИНФО", "Лист '" + sheets[i].name + "' пустой, пропускаю");
                continue;
            }
            tables.push_back({sheets[i].name, tableNames[i], data});
        }
        if (tables.empty()) throw std::runtime_error("В XLSX нет листов с непустыми данными");
        fs::remove_all(temp);
        return tables;
    } catch (...) {
        fs::remove_all(temp);
        throw;
    }
}

static bool isNullish(const std::string& s) {
    std::string v = lower(trim(s));
    return v.empty() || v == "null" || v == "n/a" || v == "na";
}

static bool matches(const std::string& s, const std::regex& re) {
    return std::regex_match(trim(s), re);
}

static std::vector<DataType> inferTypes(const TableData& data) {
    std::regex intRe("[-+]?[0-9]+");
    std::regex realRe("[-+]?(?:[0-9]+[\\.,][0-9]+|[0-9]+)");
    std::regex boolRe("(true|false|yes|no|y|n|0|1)", std::regex::icase);
    std::regex dateRe("[0-9]{4}-[0-9]{2}-[0-9]{2}");
    std::regex tsRe("[0-9]{4}-[0-9]{2}-[0-9]{2}[ T][0-9]{2}:[0-9]{2}:[0-9]{2}.*");

    std::vector<DataType> types(data.columns.size(), DataType::Integer);
    for (size_t col = 0; col < data.columns.size(); ++col) {
        bool canInt = true, canReal = true, canBool = true, canDate = true, canTs = true;
        size_t seen = 0;
        for (const auto& row : data.rows) {
            if (col >= row.size() || isNullish(row[col])) continue;
            std::string v = trim(row[col]);
            seen++;
            if (!matches(v, intRe)) canInt = false;
            if (!matches(v, realRe)) canReal = false;
            if (!matches(v, boolRe)) canBool = false;
            if (!matches(v, dateRe)) canDate = false;
            if (!matches(v, tsRe)) canTs = false;
        }
        if (seen == 0) types[col] = DataType::Text;
        else if (canBool) types[col] = DataType::Boolean;
        else if (canInt) types[col] = DataType::Integer;
        else if (canReal) types[col] = DataType::Real;
        else if (canTs) types[col] = DataType::Timestamp;
        else if (canDate) types[col] = DataType::Date;
        else types[col] = DataType::Text;
    }
    return types;
}

static std::string typeName(DataType t, const std::string& db) {
    if (db == "mysql") {
        switch (t) {
            case DataType::Integer: return "BIGINT";
            case DataType::Real: return "DOUBLE";
            case DataType::Boolean: return "BOOLEAN";
            case DataType::Date: return "DATE";
            case DataType::Timestamp: return "DATETIME";
            case DataType::Text: return "TEXT";
        }
    }
    if (db == "sqlserver") {
        switch (t) {
            case DataType::Integer: return "BIGINT";
            case DataType::Real: return "FLOAT";
            case DataType::Boolean: return "BIT";
            case DataType::Date: return "DATE";
            case DataType::Timestamp: return "DATETIME2";
            case DataType::Text: return "NVARCHAR(MAX)";
        }
    }
    switch (t) {
        case DataType::Integer: return "BIGINT";
        case DataType::Real: return "DOUBLE PRECISION";
        case DataType::Boolean: return "BOOLEAN";
        case DataType::Date: return "DATE";
        case DataType::Timestamp: return "TIMESTAMP";
        case DataType::Text: return "TEXT";
    }
    return "TEXT";
}

static std::string quoteIdent(const std::string& id, const std::string& db) {
    std::string escaped = id;
    if (db == "mysql") {
        replaceAll(escaped, "`", "``");
        return "`" + escaped + "`";
    }
    if (db == "sqlserver") {
        replaceAll(escaped, "]", "]]");
        return "[" + escaped + "]";
    }
    replaceAll(escaped, "\"", "\"\"");
    return "\"" + escaped + "\"";
}

static std::string sqlLiteral(const std::string& value, DataType type) {
    if (isNullish(value)) return "NULL";
    std::string v = trim(value);
    if (type == DataType::Real) {
        std::replace(v.begin(), v.end(), ',', '.');
        return v;
    }
    if (type == DataType::Integer) return v;
    if (type == DataType::Boolean) {
        std::string l = lower(v);
        if (l == "true" || l == "yes" || l == "y" || l == "1") return "TRUE";
        return "FALSE";
    }
    std::string out = "'";
    for (char c : v) {
        if (c == '\'') out += "''";
        else out += c;
    }
    out += "'";
    return out;
}

class OdbcConnection {
public:
    explicit OdbcConnection(const std::string& connstr) {
        SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &env_);
        SQLSetEnvAttr(env_, SQL_ATTR_ODBC_VERSION, reinterpret_cast<void*>(SQL_OV_ODBC3), 0);
        SQLAllocHandle(SQL_HANDLE_DBC, env_, &dbc_);
        std::wstring connstrW = utf8ToWide(connstr);
        SQLWCHAR out[2048];
        SQLSMALLINT outLen = 0;
        SQLRETURN ret = SQLDriverConnectW(
            dbc_, nullptr,
            reinterpret_cast<SQLWCHAR*>(connstrW.data()),
            SQL_NTS, out, sizeof(out), &outLen, SQL_DRIVER_NOPROMPT);
        if (!SQL_SUCCEEDED(ret)) {
            throw std::runtime_error("Не удалось подключиться через ODBC: " + diagnostics(SQL_HANDLE_DBC, dbc_));
        }
    }

    ~OdbcConnection() {
        if (dbc_) {
            SQLDisconnect(dbc_);
            SQLFreeHandle(SQL_HANDLE_DBC, dbc_);
        }
        if (env_) SQLFreeHandle(SQL_HANDLE_ENV, env_);
    }

    void exec(const std::string& sql) {
        SQLHSTMT stmt = nullptr;
        SQLAllocHandle(SQL_HANDLE_STMT, dbc_, &stmt);
        std::wstring sqlW = utf8ToWide(sql);
        SQLRETURN ret = SQLExecDirectW(stmt, reinterpret_cast<SQLWCHAR*>(sqlW.data()), SQL_NTS);
        if (!SQL_SUCCEEDED(ret)) {
            std::string err = diagnostics(SQL_HANDLE_STMT, stmt);
            SQLFreeHandle(SQL_HANDLE_STMT, stmt);
            throw std::runtime_error("Ошибка SQL: " + err + "\nSQL: " + sql);
        }
        SQLFreeHandle(SQL_HANDLE_STMT, stmt);
    }

private:
    static std::string diagnostics(SQLSMALLINT type, SQLHANDLE handle) {
        std::wostringstream ss;
        SQLWCHAR state[6];
        SQLWCHAR text[1024];
        SQLINTEGER nativeError = 0;
        SQLSMALLINT textLen = 0;
        for (SQLSMALLINT i = 1; ; ++i) {
            SQLRETURN ret = SQLGetDiagRecW(type, handle, i, state, &nativeError, text, sizeof(text) / sizeof(SQLWCHAR), &textLen);
            if (!SQL_SUCCEEDED(ret)) break;
            if (i > 1) ss << L" | ";
            ss << reinterpret_cast<wchar_t*>(state) << L": " << reinterpret_cast<wchar_t*>(text);
        }
        std::string result = wideToUtf8(ss.str());
        return result.empty() ? "диагностика ODBC недоступна" : result;
    }

    SQLHENV env_ = nullptr;
    SQLHDBC dbc_ = nullptr;
};

static std::string buildConnStr(const Options& opt) {
    if (!opt.connstr.empty()) return opt.connstr;
    std::string driver = opt.driver;
    if (driver.empty()) {
        if (opt.db == "postgres") driver = "PostgreSQL ODBC Driver(UNICODE)";
        else if (opt.db == "mysql") driver = "MySQL ODBC 8.0 Unicode Driver";
        else if (opt.db == "sqlserver") driver = "ODBC Driver 18 for SQL Server";
        else driver = opt.db;
    }
    if (opt.db == "sqlserver") {
        return "DRIVER={" + driver + "};SERVER=" + opt.host + "," + opt.port +
            ";DATABASE=" + opt.dbname + ";UID=" + opt.user + ";PWD=" + opt.password +
            ";TrustServerCertificate=yes;";
    }
    return "DRIVER={" + driver + "};SERVER=" + opt.host + ";PORT=" + opt.port +
        ";DATABASE=" + opt.dbname + ";UID=" + opt.user + ";PWD=" + opt.password + ";";
}

static std::string createTableSql(const Options& opt, const TableData& data, const std::vector<DataType>& types) {
    std::ostringstream sql;
    sql << "CREATE TABLE " << quoteIdent(opt.table, opt.db) << " (";
    for (size_t i = 0; i < data.columns.size(); ++i) {
        if (i) sql << ", ";
        sql << quoteIdent(data.columns[i], opt.db) << " " << typeName(types[i], opt.db);
    }
    sql << ")";
    return sql.str();
}

static std::string insertSql(const Options& opt, const TableData& data, const std::vector<DataType>& types, const std::vector<std::string>& row) {
    std::ostringstream sql;
    sql << "INSERT INTO " << quoteIdent(opt.table, opt.db) << " (";
    for (size_t i = 0; i < data.columns.size(); ++i) {
        if (i) sql << ", ";
        sql << quoteIdent(data.columns[i], opt.db);
    }
    sql << ") VALUES (";
    for (size_t i = 0; i < data.columns.size(); ++i) {
        if (i) sql << ", ";
        sql << sqlLiteral(i < row.size() ? row[i] : "", types[i]);
    }
    sql << ")";
    return sql.str();
}

static void printUsage() {
    std::cout <<
        "Загрузчик SQL\n\n"
        "Использование:\n"
        "  sql_loader.exe\n"
        "  sql_loader.exe --input data.csv [--table table_name] [--dry-run]\n"
        "  sql_loader.exe --input data.xlsx --db postgres --host localhost --port 5432 --dbname sociology_survey --user postgres --password YOUR_PASSWORD\n\n"
        "Базы данных через ODBC:\n"
        "  postgres, mysql, sqlserver или своя строка --connstr \"DRIVER={...};...\"\n\n";
}

static Options interactiveOptions() {
    Options opt;
    opt.interactive = true;

    std::cout << "Загрузчик SQL - интерфейс в терминале\n";
    std::cout << "Поддерживаемые файлы: .csv и .xlsx\n\n";

    std::cout << "Сначала укажите данные SQL-сервера.\n";
    std::cout << "\nТип базы данных:\n";
    std::cout << "  1 - PostgreSQL\n";
    std::cout << "  2 - MySQL\n";
    std::cout << "  3 - SQL Server\n";
    std::cout << "  4 - Своя строка подключения ODBC\n";
    std::string dbChoice = trim(prompt("Выберите базу данных", "1"));
    if (dbChoice == "2") opt.db = "mysql";
    else if (dbChoice == "3") opt.db = "sqlserver";
    else if (dbChoice == "4") {
        opt.db = "custom";
        opt.connstr = prompt("Строка подключения ODBC");
    } else {
        opt.db = "postgres";
    }

    if (opt.connstr.empty()) {
        if (opt.db == "postgres") {
            opt.host = prompt("Хост", "localhost");
            opt.port = prompt("Порт", "5432");
            opt.dbname = prompt("Имя базы данных", "sociology_survey");
            opt.user = prompt("Пользователь", "postgres");
            opt.password = prompt("Пароль");
        } else if (opt.db == "mysql") {
            opt.host = prompt("Хост", "localhost");
            opt.port = prompt("Порт", "3306");
            opt.dbname = prompt("Имя базы данных");
            opt.user = prompt("Пользователь", "root");
            opt.password = prompt("Пароль");
        } else if (opt.db == "sqlserver") {
            opt.host = prompt("Хост", "localhost");
            opt.port = prompt("Порт", "1433");
            opt.dbname = prompt("Имя базы данных");
            opt.user = prompt("Пользователь", "sa");
            opt.password = prompt("Пароль");
        }
    }

    std::cout << "\nТеперь выберите файл в окне Проводника.\n";
    while (true) {
        opt.input = chooseInputFileDialog();
        if (opt.input.empty()) {
            opt.input = prompt("Файл не выбран. Вставьте путь к CSV/XLSX файлу вручную или нажмите Enter для повторного выбора");
            if (opt.input.empty()) continue;
        }
        if (!fs::exists(makePath(opt.input))) {
            std::cout << "Файл не найден: " << opt.input << "\n";
            continue;
        }
        if (!endsWithLower(opt.input, ".csv") && !endsWithLower(opt.input, ".xlsx")) {
            std::cout << "Неподдерживаемый файл. Используйте .csv или .xlsx.\n";
            continue;
        }
        break;
    }

    if (endsWithLower(opt.input, ".csv")) {
        std::string defaultTable = sanitizeIdentifier(pathToUtf8(makePath(opt.input).stem()), "imported_data");
        opt.table = sanitizeIdentifier(prompt("Имя таблицы", defaultTable), defaultTable);
    } else {
        std::cout << "Для XLSX каждый лист будет загружен в отдельную SQL-таблицу.\n";
        std::cout << "Названия таблиц будут взяты из названий листов Excel.\n";
    }

    std::cout << "\nПеред записью в SQL можно выполнить безопасную проверку.\n";
    opt.dryRun = promptYesNo("Только проверить, без изменений в базе", true);
    if (!opt.dryRun) {
        opt.dropExisting = promptYesNo("Удалить существующую таблицу с таким же именем перед загрузкой", false);
    }

    std::cout << "\n";
    return opt;
}

static Options parseArgs(const std::vector<std::string>& args) {
    Options opt;
    if (args.size() <= 1) return interactiveOptions();
    for (size_t i = 1; i < args.size(); ++i) {
        std::string a = args[i];
        auto need = [&](const std::string& name) -> std::string {
            if (i + 1 >= args.size()) throw std::runtime_error("Не указано значение для " + name);
            return args[++i];
        };
        if (a == "--input") opt.input = need(a);
        else if (a == "--db") opt.db = lower(need(a));
        else if (a == "--host") opt.host = need(a);
        else if (a == "--port") opt.port = need(a);
        else if (a == "--dbname") opt.dbname = need(a);
        else if (a == "--user") opt.user = need(a);
        else if (a == "--password") opt.password = need(a);
        else if (a == "--table") opt.table = sanitizeIdentifier(need(a), "imported_data");
        else if (a == "--driver") opt.driver = need(a);
        else if (a == "--connstr") opt.connstr = need(a);
        else if (a == "--dry-run") opt.dryRun = true;
        else if (a == "--drop-existing") opt.dropExisting = true;
        else if (a == "--help" || a == "-h") {
            printUsage();
            std::exit(0);
        } else {
            throw std::runtime_error("Неизвестный аргумент: " + a);
        }
    }
    if (opt.input.empty()) throw std::runtime_error("Нужно указать --input");
    if (opt.table.empty()) {
        opt.table = sanitizeIdentifier(pathToUtf8(makePath(opt.input).stem()), "imported_data");
    }
    return opt;
}

int main() {
    try {
        SetConsoleOutputCP(CP_UTF8);
        SetConsoleCP(CP_UTF8);
        g_log.open("sql_loader.log", std::ios::app);
        Options opt = parseArgs(commandLineArgsUtf8());
        logLine("ИНФО", "Файл: " + opt.input);
        logLine("ИНФО", "Целевая СУБД: " + opt.db + ", база: " + opt.dbname);

        fs::path inputPath = makePath(opt.input);
        std::vector<NamedTable> tables;
        if (endsWithLower(opt.input, ".csv")) {
            TableData data = readCsv(inputPath);
            if (data.columns.empty()) throw std::runtime_error("Не найдено ни одного столбца");
            tables.push_back({pathToUtf8(inputPath.filename()), opt.table, data});
        } else if (endsWithLower(opt.input, ".xlsx")) {
            tables = readXlsxTables(inputPath);
        } else {
            throw std::runtime_error("Неподдерживаемый формат файла. Используйте .csv или .xlsx");
        }

        if (opt.dryRun) {
            logLine("ИНФО", "Режим проверки. Предпросмотр SQL:");
            for (const auto& table : tables) {
                Options tableOpt = opt;
                tableOpt.table = table.tableName;
                auto types = inferTypes(table.data);
                logLine("ИНФО", "Источник: " + table.sourceName + " -> таблица: " + table.tableName);
                logLine("ИНФО", "Найдено столбцов: " + std::to_string(table.data.columns.size()));
                logLine("ИНФО", "Найдено строк: " + std::to_string(table.data.rows.size()));
                for (size_t i = 0; i < table.data.columns.size(); ++i) {
                    logLine("ИНФО", "Столбец " + std::to_string(i + 1) + ": " + table.data.columns[i] + " -> " + typeName(types[i], opt.db));
                }
                logLine("SQL", createTableSql(tableOpt, table.data, types));
                if (!table.data.rows.empty()) logLine("SQL", insertSql(tableOpt, table.data, types, table.data.rows.front()));
            }
            logLine("ИНФО", "Итог: файл успешно прочитан, база данных не изменялась");
            return 0;
        }

        OdbcConnection db(buildConnStr(opt));
        logLine("ИНФО", "Подключение через ODBC выполнено");

        size_t totalOk = 0;
        size_t totalRows = 0;
        for (const auto& table : tables) {
            Options tableOpt = opt;
            tableOpt.table = table.tableName;
            auto types = inferTypes(table.data);
            totalRows += table.data.rows.size();

            logLine("ИНФО", "Источник: " + table.sourceName + " -> таблица: " + table.tableName);
            logLine("ИНФО", "Найдено столбцов: " + std::to_string(table.data.columns.size()));
            logLine("ИНФО", "Найдено строк: " + std::to_string(table.data.rows.size()));
            for (size_t i = 0; i < table.data.columns.size(); ++i) {
                logLine("ИНФО", "Столбец " + std::to_string(i + 1) + ": " + table.data.columns[i] + " -> " + typeName(types[i], opt.db));
            }

            if (opt.dropExisting) {
                db.exec("DROP TABLE IF EXISTS " + quoteIdent(tableOpt.table, opt.db));
                logLine("ИНФО", "Таблица '" + tableOpt.table + "' удалена, если она была");
            }
            db.exec(createTableSql(tableOpt, table.data, types));
            logLine("ИНФО", "Таблица '" + tableOpt.table + "' создана");

            size_t ok = 0;
            for (size_t i = 0; i < table.data.rows.size(); ++i) {
                try {
                    db.exec(insertSql(tableOpt, table.data, types, table.data.rows[i]));
                    ok++;
                    totalOk++;
                    if (ok % 100 == 0) logLine("ИНФО", "Таблица '" + tableOpt.table + "': загружено строк: " + std::to_string(ok));
                } catch (const std::exception& e) {
                    logLine("ОШИБКА", "Таблица '" + tableOpt.table + "', строка " + std::to_string(i + 2) + " пропущена: " + e.what());
                }
            }
            logLine("ИНФО", "Таблица '" + tableOpt.table + "': загружено строк: " + std::to_string(ok) + "/" + std::to_string(table.data.rows.size()));
        }

        logLine("ИНФО", "Готово. Загружено таблиц: " + std::to_string(tables.size()));
        logLine("ИНФО", "Готово. Загружено строк: " + std::to_string(totalOk) + "/" + std::to_string(totalRows));
        logLine("ИНФО", "Итог: файл '" + opt.input + "' загружен в SQL");
        return 0;
    } catch (const std::exception& e) {
        logLine("ОШИБКА", e.what());
        printUsage();
        return 1;
    }
}
