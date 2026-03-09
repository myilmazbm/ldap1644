// LDAPStatCollector - Hızlı 1644 EVTX analiz aracı (C++)
// Derleme (Visual Studio Developer Command Prompt):
//   cl /EHsc ldapstat.cpp wevtapi.lib /Fe:LDAPStatCollector.exe

#include <windows.h>
#include <winevt.h>

#include <iostream>
#include <string>
#include <vector>
#include <unordered_map>
#include <algorithm>
#include <memory>

#pragma comment(lib, "wevtapi.lib")

struct StatItem {
    uint64_t count = 0;
    double totalMs = 0.0;
};

struct GlobalStats {
    uint64_t totalQueries = 0;
    double totalTimeMs = 0.0;
};

struct Parsed1644 {
    std::wstring ip;
    double timeMs = 0.0;
    std::wstring filter;
    std::wstring baseDn;
    std::wstring user;
};

static void PrintError(const char* where, DWORD err) {
    LPWSTR msg = nullptr;
    FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                   NULL, err, 0, (LPWSTR)&msg, 0, NULL);
    std::wcerr << L"[HATA] " << where << L" (0x" << std::hex << err << L"): "
               << (msg ? msg : L"bilinmeyen hata") << std::endl;
    if (msg) LocalFree(msg);
}

static std::wstring ToLower(const std::wstring& s) {
    std::wstring r = s;
    std::transform(r.begin(), r.end(), r.begin(), [](wchar_t c) { return (wchar_t)towlower(c); });
    return r;
}

static std::wstring Trim(const std::wstring& s) {
    size_t start = 0;
    while (start < s.size() && iswspace(s[start])) ++start;
    size_t end = s.size();
    while (end > start && iswspace(s[end - 1])) --end;
    return s.substr(start, end - start);
}

// Basit substring arama: "Search Time:" satırından sayı çek
static bool TryParseDoubleAfter(const std::wstring& text, const std::wstring& token, double& out) {
    size_t pos = ToLower(text).find(ToLower(token));
    if (pos == std::wstring::npos) return false;
    pos += token.size();
    while (pos < text.size() && iswspace(text[pos])) ++pos;

    std::wstring num;
    while (pos < text.size() && (iswdigit(text[pos]) || text[pos] == L'.' || text[pos] == L',')) {
        wchar_t c = text[pos++];
        if (c == L',') c = L'.'; // ondalık için
        num.push_back(c);
    }
    if (num.empty()) return false;
    try {
        out = std::stod(num);
        return true;
    } catch (...) {
        return false;
    }
}

static bool TryExtractAfterLabel(const std::wstring& text, const std::wstring& label, std::wstring& out) {
    size_t pos = ToLower(text).find(ToLower(label));
    if (pos == std::wstring::npos) return false;
    pos += label.size();
    while (pos < text.size() && (text[pos] == L' ' || text[pos] == L'\t' || text[pos] == L':')) ++pos;

    size_t end = text.find(L'\n', pos);
    if (end == std::wstring::npos) end = text.size();
    out = Trim(text.substr(pos, end - pos));
    if (!out.empty() && out.front() == L'\'' && out.back() == L'\'') {
        out = out.substr(1, out.size() - 2);
    }
    return !out.empty();
}

// Çok basit XML içinden <Data Name="xxx">value</Data> çekme
static bool TryExtractDataField(const std::wstring& xml, const std::wstring& name, std::wstring& out) {
    std::wstring pattern = L"<Data Name=\"" + name + L"\">";
    size_t pos = xml.find(pattern);
    if (pos == std::wstring::npos) return false;
    pos += pattern.size();
    size_t end = xml.find(L"</Data>", pos);
    if (end == std::wstring::npos) return false;
    out = xml.substr(pos, end - pos);
    return !out.empty();
}

static bool Parse1644FromXml(const std::wstring& xml, Parsed1644& e) {
    bool ok = false;

    // Önce XML Data alanları
    std::wstring tmp;
    if (TryExtractDataField(xml, L"ClientIP", tmp) || TryExtractDataField(xml, L"Client Ip", tmp)) {
        e.ip = Trim(tmp);
        ok = true;
    } else {
        // Mesaj metninden ara
        TryExtractAfterLabel(xml, L"Client IP", e.ip);
    }

    double t = 0.0;
    if (TryExtractDataField(xml, L"SearchTime", tmp)) {
        try {
            t = std::stod(tmp);
            e.timeMs = t;
            ok = true;
        } catch (...) {}
    }
    if (e.timeMs <= 0.0) {
        if (TryParseDoubleAfter(xml, L"Search Time", t) || TryParseDoubleAfter(xml, L"Duration_MS", t)) {
            e.timeMs = t;
            ok = true;
        }
    }

    if (TryExtractDataField(xml, L"Filter", e.filter) == false) {
        TryExtractAfterLabel(xml, L"Search Filter", e.filter);
    }
    e.filter = Trim(e.filter);

    if (TryExtractDataField(xml, L"BaseDN", e.baseDn) == false) {
        TryExtractAfterLabel(xml, L"Search Scope Base", e.baseDn);
    }
    e.baseDn = Trim(e.baseDn);

    if (TryExtractDataField(xml, L"User", e.user) == false &&
        TryExtractDataField(xml, L"Client", e.user) == false) {
        TryExtractAfterLabel(xml, L"User", e.user);
    }
    e.user = Trim(e.user);

    return ok;
}

class StatsAggregator {
public:
    void Add(const Parsed1644& e) {
        if (e.timeMs <= 0.0) return;
        globals.totalQueries++;
        globals.totalTimeMs += e.timeMs;

        ipStats[e.ip].count++;
        ipStats[e.ip].totalMs += e.timeMs;

        std::wstring normFilter = Trim(ToLower(e.filter));
        filterStats[normFilter].count++;
        filterStats[normFilter].totalMs += e.timeMs;

        std::wstring normBase = Trim(ToLower(e.baseDn));
        baseStats[normBase].count++;
        baseStats[normBase].totalMs += e.timeMs;
    }

    void PrintIpStats() const {
        std::wcout << L"\n=== IP Gruplama ===\n";
        std::vector<std::pair<std::wstring, StatItem>> v(ipStats.begin(), ipStats.end());
        std::sort(v.begin(), v.end(), [](auto& a, auto& b) {
            double avgA = a.second.count ? a.second.totalMs / a.second.count : 0.0;
            double avgB = b.second.count ? b.second.totalMs / b.second.count : 0.0;
            return avgA > avgB;
        });

        double totalQ = (double)globals.totalQueries;
        double running = 0.0;

        std::wcout << L"IP\tCount\tAvgMs\t%All\t%Running\n";
        for (auto& kv : v) {
            const auto& ip = kv.first;
            const auto& st = kv.second;
            double avg = st.count ? st.totalMs / st.count : 0.0;
            double pct = totalQ > 0 ? (st.count * 100.0 / totalQ) : 0.0;
            running += pct;
            std::wcout << ip << L"\t" << st.count << L"\t" << avg << L"\t"
                       << pct << L"\t" << running << L"\n";
        }
    }

    void PrintTopFilters(size_t topN = 20) const {
        std::wcout << L"\n=== En Pahalı Filtreler (Top " << topN << L") ===\n";
        std::vector<std::pair<std::wstring, StatItem>> v(filterStats.begin(), filterStats.end());
        std::sort(v.begin(), v.end(), [](auto& a, auto& b) {
            return a.second.totalMs > b.second.totalMs;
        });
        if (v.size() > topN) v.resize(topN);

        std::wcout << L"Count\tAvgMs\tTotalMs\tFilter\n";
        for (auto& kv : v) {
            const auto& f = kv.first;
            const auto& st = kv.second;
            double avg = st.count ? st.totalMs / st.count : 0.0;
            std::wcout << st.count << L"\t" << avg << L"\t" << st.totalMs << L"\t" << f << L"\n";
        }
    }

    void PrintTopBases(size_t topN = 20) const {
        std::wcout << L"\n=== En Yoğun BaseDN (Top " << topN << L") ===\n";
        std::vector<std::pair<std::wstring, StatItem>> v(baseStats.begin(), baseStats.end());
        std::sort(v.begin(), v.end(), [](auto& a, auto& b) {
            return a.second.totalMs > b.second.totalMs;
        });
        if (v.size() > topN) v.resize(topN);

        std::wcout << L"Count\tAvgMs\tTotalMs\tBaseDN\n";
        for (auto& kv : v) {
            const auto& dn = kv.first;
            const auto& st = kv.second;
            double avg = st.count ? st.totalMs / st.count : 0.0;
            std::wcout << st.count << L"\t" << avg << L"\t" << st.totalMs << L"\t" << dn << L"\n";
        }
    }

    void PrintSummary() const {
        std::wcout << L"\n=== Özet ===\n";
        std::wcout << L"Toplam 1644 olay sayısı: " << globals.totalQueries << L"\n";
        double avg = globals.totalQueries ? globals.totalTimeMs / globals.totalQueries : 0.0;
        std::wcout << L"Ortalama arama süresi (ms): " << avg << L"\n";
    }

private:
    GlobalStats globals;
    std::unordered_map<std::wstring, StatItem> ipStats;
    std::unordered_map<std::wstring, StatItem> filterStats;
    std::unordered_map<std::wstring, StatItem> baseStats;
};

int wmain(int argc, wchar_t* argv[]) {
    std::wcout.sync_with_stdio(false);

    const wchar_t* channel = L"Directory Service";

    std::wcout << L"LDAPStatCollector - EVTX 1644 analiz\n";
    std::wcout << L"Kanal: " << channel << L"\n";

    EVT_HANDLE hQuery = EvtQuery(
        NULL,
        channel,
        L"*[(System/EventID=1644)]",
        EvtQueryChannelPath);

    if (!hQuery) {
        DWORD err = GetLastError();
        PrintError("EvtQuery", err);
        return 1;
    }

    StatsAggregator agg;

    const DWORD batchSize = 128;
    EVT_HANDLE events[batchSize];
    DWORD returned = 0;

    while (true) {
        if (!EvtNext(hQuery, batchSize, events, 10000, 0, &returned)) {
            DWORD err = GetLastError();
            if (err == ERROR_NO_MORE_ITEMS) break;
            if (err == ERROR_TIMEOUT) continue;
            PrintError("EvtNext", err);
            break;
        }

        for (DWORD i = 0; i < returned; ++i) {
            EVT_HANDLE hEvent = events[i];

            // XML boyutunu öğren
            DWORD bufferSize = 0;
            DWORD bufferUsed = 0;
            DWORD propCount = 0;

            if (!EvtRender(NULL, hEvent, EvtRenderEventXml,
                           0, NULL, &bufferUsed, &propCount)) {
                DWORD err = GetLastError();
                if (err != ERROR_INSUFFICIENT_BUFFER) {
                    PrintError("EvtRender(size)", err);
                    EvtClose(hEvent);
                    continue;
                }
            }

            std::unique_ptr<wchar_t[]> buffer(new wchar_t[bufferUsed / sizeof(wchar_t) + 1]);
            if (!EvtRender(NULL, hEvent, EvtRenderEventXml,
                           bufferUsed, buffer.get(), &bufferUsed, &propCount)) {
                PrintError("EvtRender", GetLastError());
                EvtClose(hEvent);
                continue;
            }

            std::wstring xml(buffer.get());
            Parsed1644 e{};
            if (Parse1644FromXml(xml, e)) {
                agg.Add(e);
            }

            EvtClose(hEvent);
        }
    }

    EvtClose(hQuery);

    agg.PrintSummary();
    agg.PrintIpStats();
    agg.PrintTopFilters();
    agg.PrintTopBases();

    std::wcout << L"\nBitti.\n";
    return 0;
}

