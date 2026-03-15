// Event1644Reader - C++ Win32 port of Event1644Reader.ps1 v1.04 by Ming Chen
// Scans .evtx files for Event 1644 (LDAP search events), extracts fields to CSV,
// then optionally imports into Excel via COM automation with pivot tables.
//
// Build: cl /EHsc /W4 /DUNICODE /D_UNICODE Event1644Reader.cpp ole32.lib oleaut32.lib wevtapi.lib shlwapi.lib

#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif
#define WIN32_LEAN_AND_MEAN

#include <windows.h>
#include <winevt.h>
#include <shlwapi.h>
#include <oleauto.h>

#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <cwchar>
#include <string>
#include <vector>

#pragma comment(lib, "wevtapi.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "shlwapi.lib")

// ============================================================================
// Utility helpers
// ============================================================================

static std::wstring GetExeDirectory()
{
    WCHAR buf[MAX_PATH];
    GetModuleFileNameW(nullptr, buf, MAX_PATH);
    PathRemoveFileSpecW(buf);
    return buf;
}

static std::wstring ReadLine(const wchar_t* prompt)
{
    wprintf(L"%s", prompt);
    WCHAR buf[1024] = {};
    if (fgetws(buf, _countof(buf), stdin))
    {
        size_t len = wcslen(buf);
        while (len > 0 && (buf[len - 1] == L'\n' || buf[len - 1] == L'\r'))
            buf[--len] = L'\0';
    }
    return buf;
}

// Parse client string that may look like:
//   IPv4  "10.0.0.1:12345"  or  "10.0.0.1"
//   IPv6  "[fe80::1%25]:12345"
//   Named "SomeMachine"
struct ClientInfo
{
    std::wstring ip;
    std::wstring port;
};

static ClientInfo ParseClient(const std::wstring& client)
{
    ClientInfo ci;
    ci.ip   = L"Unknown";
    ci.port = L"Unknown";

    if (client.empty())
        return ci;

    // IPv6 bracket form: [addr]:port
    if (client[0] == L'[')
    {
        size_t close = client.find(L']');
        if (close != std::wstring::npos)
        {
            ci.ip = client.substr(0, close + 1); // include brackets
            if (close + 2 < client.size() && client[close + 1] == L':')
                ci.port = client.substr(close + 2);
        }
        return ci;
    }

    // IPv4: digits and dots, optionally :port
    bool looksLikeIPv4 = false;
    for (size_t i = 0; i < client.size(); ++i)
    {
        wchar_t c = client[i];
        if (c == L'.')
        {
            looksLikeIPv4 = true;
            break;
        }
    }

    if (looksLikeIPv4)
    {
        size_t colon = client.rfind(L':');
        if (colon != std::wstring::npos)
        {
            ci.ip   = client.substr(0, colon);
            ci.port = client.substr(colon + 1);
        }
        else
        {
            ci.ip = client;
        }
        return ci;
    }

    // Named client (hostname etc.)
    ci.ip = client;
    return ci;
}

// ============================================================================
// Event 1644 record
// ============================================================================

struct Event1644
{
    std::wstring ldapServer;
    std::wstring timeGenerated;
    std::wstring clientIP;
    std::wstring clientPort;
    std::wstring startingNode;       // Property[0]
    std::wstring filter;             // Property[1]
    std::wstring searchScope;        // Property[5]
    std::wstring attributeSelection; // Property[6]
    std::wstring serverControls;     // Property[7]
    std::wstring visitedEntries;     // Property[2]
    std::wstring returnedEntries;    // Property[3]
    // KB2800945+ extended fields
    std::wstring usedIndexes;        // Property[8]
    std::wstring pagesReferenced;    // Property[9]
    std::wstring pagesReadFromDisk;  // Property[10]
    std::wstring pagesPreReadFromDisk; // Property[11]
    std::wstring cleanPagesModified; // Property[12]
    std::wstring dirtyPagesModified; // Property[13]
    std::wstring searchTimeMS;       // Property[14]
    std::wstring attrPreventingOpt;  // Property[15]
};

// CSV-escape: double-quote any field that contains comma, quote, or newline
static std::wstring CsvEscape(const std::wstring& s)
{
    bool needQuote = false;
    for (auto c : s)
    {
        if (c == L',' || c == L'"' || c == L'\n' || c == L'\r')
        {
            needQuote = true;
            break;
        }
    }
    if (!needQuote)
        return s;

    std::wstring out = L"\"";
    for (auto c : s)
    {
        if (c == L'"')
            out += L"\"\"";
        else
            out += c;
    }
    out += L'"';
    return out;
}

static const wchar_t* CSV_HEADER =
    L"LDAPServer,TimeGenerated,ClientIP,ClientPort,StartingNode,Filter,"
    L"SearchScope,AttributeSelection,ServerControls,VisitedEntries,ReturnedEntries,"
    L"UsedIndexes,PagesReferenced,PagesReadFromDisk,PagesPreReadFromDisk,"
    L"CleanPagesModified,DirtyPagesModified,SearchTimeMS,AttributesPreventingOptimization";

static std::wstring Event1644ToCsvLine(const Event1644& e)
{
    std::wstring line;
    line += CsvEscape(e.ldapServer)        + L",";
    line += CsvEscape(e.timeGenerated)     + L",";
    line += CsvEscape(e.clientIP)          + L",";
    line += CsvEscape(e.clientPort)        + L",";
    line += CsvEscape(e.startingNode)      + L",";
    line += CsvEscape(e.filter)            + L",";
    line += CsvEscape(e.searchScope)       + L",";
    line += CsvEscape(e.attributeSelection)+ L",";
    line += CsvEscape(e.serverControls)    + L",";
    line += CsvEscape(e.visitedEntries)    + L",";
    line += CsvEscape(e.returnedEntries)   + L",";
    line += CsvEscape(e.usedIndexes)       + L",";
    line += CsvEscape(e.pagesReferenced)   + L",";
    line += CsvEscape(e.pagesReadFromDisk) + L",";
    line += CsvEscape(e.pagesPreReadFromDisk)+ L",";
    line += CsvEscape(e.cleanPagesModified)+ L",";
    line += CsvEscape(e.dirtyPagesModified)+ L",";
    line += CsvEscape(e.searchTimeMS)      + L",";
    line += CsvEscape(e.attrPreventingOpt);
    return line;
}

// ============================================================================
// Windows Event Log helpers
// ============================================================================

// Render a system value from an event as a wide string
static std::wstring RenderSystemValue(EVT_HANDLE hEvent, EVT_HANDLE hContext)
{
    DWORD bufSize = 0, propCount = 0;
    EvtRender(hContext, hEvent, EvtRenderEventValues, 0, nullptr, &bufSize, &propCount);
    if (GetLastError() != ERROR_INSUFFICIENT_BUFFER)
        return {};

    std::vector<BYTE> buf(bufSize);
    if (!EvtRender(hContext, hEvent, EvtRenderEventValues, bufSize,
                   buf.data(), &bufSize, &propCount))
        return {};

    auto* pVals = reinterpret_cast<PEVT_VARIANT>(buf.data());
    if (propCount == 0)
        return {};

    switch (pVals[0].Type)
    {
    case EvtVarTypeString:
        return pVals[0].StringVal ? pVals[0].StringVal : L"";
    case EvtVarTypeFileTime:
    {
        FILETIME ft = { pVals[0].FileTimeVal & 0xFFFFFFFF, (DWORD)(pVals[0].FileTimeVal >> 32) };
        SYSTEMTIME st;
        FileTimeToSystemTime(&ft, &st);
        WCHAR timeBuf[64];
        _snwprintf_s(timeBuf, _countof(timeBuf), _TRUNCATE,
                     L"%04d-%02d-%02d %02d:%02d:%02d",
                     st.wYear, st.wMonth, st.wDay,
                     st.wHour, st.wMinute, st.wSecond);
        return timeBuf;
    }
    default:
        return {};
    }
}

// Render event XML for extracting EventData properties
static std::wstring RenderEventXml(EVT_HANDLE hEvent)
{
    DWORD bufSize = 0, propCount = 0;
    EvtRender(nullptr, hEvent, EvtRenderEventXml, 0, nullptr, &bufSize, &propCount);
    if (GetLastError() != ERROR_INSUFFICIENT_BUFFER)
        return {};

    std::vector<WCHAR> buf(bufSize / sizeof(WCHAR) + 1);
    if (!EvtRender(nullptr, hEvent, EvtRenderEventXml, bufSize,
                   buf.data(), &bufSize, &propCount))
        return {};

    return buf.data();
}

// Render all EventData values from an event handle
static bool RenderEventData(EVT_HANDLE hEvent, std::vector<std::wstring>& values)
{
    values.clear();

    // Use UserData path to extract all EventData/Data elements
    DWORD bufSize = 0, propCount = 0;
    EvtRender(nullptr, hEvent, EvtRenderEventValues, 0, nullptr, &bufSize, &propCount);

    // Instead, render as XML and parse EventData manually for reliability
    std::wstring xml = RenderEventXml(hEvent);
    if (xml.empty())
        return false;

    // Find <EventData> ... </EventData>  and extract <Data ...>value</Data>
    const wchar_t* edStart = wcsstr(xml.c_str(), L"<EventData>");
    const wchar_t* edEnd   = wcsstr(xml.c_str(), L"</EventData>");
    if (!edStart || !edEnd)
        return false;

    const wchar_t* p = edStart;
    while (p < edEnd)
    {
        // Find next <Data
        const wchar_t* dataTag = wcsstr(p, L"<Data");
        if (!dataTag || dataTag >= edEnd)
            break;

        // Find closing > of <Data ...>
        const wchar_t* gt = wcschr(dataTag, L'>');
        if (!gt || gt >= edEnd)
            break;

        gt++; // skip '>'

        // Find </Data>
        const wchar_t* closeTag = wcsstr(gt, L"</Data>");
        if (!closeTag || closeTag >= edEnd)
        {
            // Self-closing or empty
            values.push_back(L"");
            p = gt;
            continue;
        }

        std::wstring val(gt, closeTag);

        // Unescape basic XML entities
        std::wstring unescaped;
        for (size_t i = 0; i < val.size(); ++i)
        {
            if (val[i] == L'&')
            {
                if (val.compare(i, 4, L"&lt;") == 0)       { unescaped += L'<'; i += 3; }
                else if (val.compare(i, 4, L"&gt;") == 0)  { unescaped += L'>'; i += 3; }
                else if (val.compare(i, 5, L"&amp;") == 0) { unescaped += L'&'; i += 4; }
                else if (val.compare(i, 6, L"&apos;") == 0){ unescaped += L'\''; i += 5; }
                else if (val.compare(i, 6, L"&quot;") == 0){ unescaped += L'"'; i += 5; }
                else unescaped += val[i];
            }
            else
            {
                unescaped += val[i];
            }
        }

        values.push_back(unescaped);
        p = closeTag + 7; // skip </Data>
    }

    return true;
}

static std::wstring SafeGet(const std::vector<std::wstring>& v, size_t idx)
{
    return idx < v.size() ? v[idx] : L"";
}

// ============================================================================
// Process a single .evtx file
// ============================================================================

static int ProcessEvtxFile(const std::wstring& evtxPath, const std::wstring& csvPath)
{
    // Query Event 1644 from the evtx file
    // XPath: Event[System[EventID=1644]]
    const wchar_t* query = L"Event/System[EventID=1644]";

    EVT_HANDLE hQuery = EvtQuery(nullptr, evtxPath.c_str(), query,
                                 EvtQueryFilePath | EvtQueryForwardDirection);
    if (!hQuery)
    {
        // Try alternate: some evtx may need channel-based query
        DWORD err = GetLastError();
        if (err == ERROR_EVT_CHANNEL_NOT_FOUND || err == ERROR_EVT_INVALID_QUERY)
        {
            // Try with wildcard query
            hQuery = EvtQuery(nullptr, evtxPath.c_str(), L"*[System[EventID=1644]]",
                              EvtQueryFilePath | EvtQueryForwardDirection);
        }
        if (!hQuery)
            return 0;
    }

    // Create render contexts for system properties
    LPCWSTR computerPath[] = { L"Event/System/Computer" };
    EVT_HANDLE hCtxComputer = EvtCreateRenderContext(1, computerPath, EvtRenderContextValues);

    LPCWSTR timePath[] = { L"Event/System/TimeCreated/@SystemTime" };
    EVT_HANDLE hCtxTime = EvtCreateRenderContext(1, timePath, EvtRenderContextValues);

    FILE* csvFile = nullptr;
    int eventCount = 0;

    EVT_HANDLE hEvents[100];
    DWORD dwReturned = 0;

    while (EvtNext(hQuery, _countof(hEvents), hEvents, INFINITE, 0, &dwReturned))
    {
        for (DWORD i = 0; i < dwReturned; ++i)
        {
            EVT_HANDLE hEvent = hEvents[i];

            // Get computer name and time
            std::wstring computerName = RenderSystemValue(hEvent, hCtxComputer);
            std::wstring timeCreated  = RenderSystemValue(hEvent, hCtxTime);

            // Get EventData properties
            std::vector<std::wstring> props;
            if (!RenderEventData(hEvent, props))
            {
                EvtClose(hEvent);
                continue;
            }

            // Build Event1644 record (matching PS script property indices)
            Event1644 rec;
            rec.ldapServer        = computerName;
            rec.timeGenerated     = timeCreated;

            std::wstring clientStr = SafeGet(props, 4);
            ClientInfo ci = ParseClient(clientStr);
            rec.clientIP          = ci.ip;
            rec.clientPort        = ci.port;

            rec.startingNode      = SafeGet(props, 0);
            rec.filter            = SafeGet(props, 1);
            rec.searchScope       = SafeGet(props, 5);
            rec.attributeSelection= SafeGet(props, 6);
            rec.serverControls    = SafeGet(props, 7);
            rec.visitedEntries    = SafeGet(props, 2);
            rec.returnedEntries   = SafeGet(props, 3);
            // Extended fields (KB2800945+)
            rec.usedIndexes       = SafeGet(props, 8);
            rec.pagesReferenced   = SafeGet(props, 9);
            rec.pagesReadFromDisk = SafeGet(props, 10);
            rec.pagesPreReadFromDisk = SafeGet(props, 11);
            rec.cleanPagesModified= SafeGet(props, 12);
            rec.dirtyPagesModified= SafeGet(props, 13);
            rec.searchTimeMS      = SafeGet(props, 14);
            rec.attrPreventingOpt = SafeGet(props, 15);

            // Open CSV on first record
            if (!csvFile)
            {
                _wfopen_s(&csvFile, csvPath.c_str(), L"w, ccs=UTF-8");
                if (!csvFile)
                {
                    wprintf(L"  Error: cannot create %s\n", csvPath.c_str());
                    EvtClose(hEvent);
                    goto cleanup;
                }
                fwprintf(csvFile, L"%s\n", CSV_HEADER);
            }

            fwprintf(csvFile, L"%s\n", Event1644ToCsvLine(rec).c_str());
            eventCount++;

            EvtClose(hEvent);
        }
    }

cleanup:
    if (csvFile)
        fclose(csvFile);
    if (hCtxComputer) EvtClose(hCtxComputer);
    if (hCtxTime)     EvtClose(hCtxTime);
    EvtClose(hQuery);

    return eventCount;
}

// ============================================================================
// Excel COM automation helpers
// ============================================================================

// Helper: call a COM method via IDispatch
static HRESULT AutoWrap(int autoType, VARIANT* pvResult, IDispatch* pDisp,
                        LPCOLESTR ptName, int cArgs, ...)
{
    if (!pDisp) return E_INVALIDARG;

    va_list marker;
    va_start(marker, cArgs);

    DISPID dispID;
    HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, const_cast<LPOLESTR*>(&ptName),
                                       1, LOCALE_USER_DEFAULT, &dispID);
    if (FAILED(hr))
    {
        va_end(marker);
        return hr;
    }

    // Allocate args in reverse order (COM convention)
    VARIANT* pArgs = new VARIANT[cArgs + 1];
    for (int i = cArgs - 1; i >= 0; i--)
    {
        pArgs[i] = va_arg(marker, VARIANT);
    }

    DISPPARAMS dp = { pArgs, nullptr, (UINT)cArgs, 0 };

    DISPID dispidNamed = DISPID_PROPERTYPUT;
    if (autoType & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF))
    {
        dp.cNamedArgs = 1;
        dp.rgdispidNamedArgs = &dispidNamed;
    }

    hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, (WORD)autoType,
                        &dp, pvResult, nullptr, nullptr);

    va_end(marker);
    delete[] pArgs;
    return hr;
}

// Convenience to get a dispatch property
static IDispatch* GetDispProp(IDispatch* pDisp, LPCOLESTR name)
{
    VARIANT result;
    VariantInit(&result);
    AutoWrap(DISPATCH_PROPERTYGET, &result, pDisp, name, 0);
    return (result.vt == VT_DISPATCH) ? result.pdispVal : nullptr;
}

// Convenience to call a method returning IDispatch
static IDispatch* CallMethod(IDispatch* pDisp, LPCOLESTR name, int cArgs, ...)
{
    VARIANT result;
    VariantInit(&result);

    if (cArgs == 0)
    {
        AutoWrap(DISPATCH_METHOD, &result, pDisp, name, 0);
    }
    else
    {
        va_list marker;
        va_start(marker, cArgs);

        VARIANT* pArgs = new VARIANT[cArgs];
        // Collect args (they'll be reversed in AutoWrap, so collect forward here
        // and pass as reversed array)
        VARIANT argsForward[16]; // max args we'll use
        for (int i = 0; i < cArgs && i < 16; i++)
            argsForward[i] = va_arg(marker, VARIANT);
        va_end(marker);

        // Build reversed
        for (int i = 0; i < cArgs; i++)
            pArgs[i] = argsForward[cArgs - 1 - i];

        DISPPARAMS dp = { pArgs, nullptr, (UINT)cArgs, 0 };
        DISPID dispID;
        HRESULT hr = pDisp->GetIDsOfNames(IID_NULL, const_cast<LPOLESTR*>(&name),
                                           1, LOCALE_USER_DEFAULT, &dispID);
        if (SUCCEEDED(hr))
        {
            pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD,
                          &dp, &result, nullptr, nullptr);
        }
        delete[] pArgs;
    }

    return (result.vt == VT_DISPATCH) ? result.pdispVal : nullptr;
}

static void PutProp(IDispatch* pDisp, LPCOLESTR name, VARIANT val)
{
    AutoWrap(DISPATCH_PROPERTYPUT, nullptr, pDisp, name, 1, val);
}

static void PutPropInt(IDispatch* pDisp, LPCOLESTR name, int val)
{
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_I4;
    v.lVal = val;
    PutProp(pDisp, name, v);
}

static void PutPropBool(IDispatch* pDisp, LPCOLESTR name, bool val)
{
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_BOOL;
    v.boolVal = val ? VARIANT_TRUE : VARIANT_FALSE;
    PutProp(pDisp, name, v);
}

static void PutPropStr(IDispatch* pDisp, LPCOLESTR name, LPCOLESTR val)
{
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_BSTR;
    v.bstrVal = SysAllocString(val);
    PutProp(pDisp, name, v);
    SysFreeString(v.bstrVal);
}

static VARIANT MakeInt(int v)
{
    VARIANT r;
    VariantInit(&r);
    r.vt = VT_I4;
    r.lVal = v;
    return r;
}

static VARIANT MakeBstr(LPCOLESTR s)
{
    VARIANT r;
    VariantInit(&r);
    r.vt = VT_BSTR;
    r.bstrVal = SysAllocString(s);
    return r;
}

static VARIANT MakeDispatch(IDispatch* d)
{
    VARIANT r;
    VariantInit(&r);
    r.vt = VT_DISPATCH;
    r.pdispVal = d;
    return r;
}

// Get Item(index) from a collection
static IDispatch* GetItem(IDispatch* pCollection, int index)
{
    VARIANT result, vIdx;
    VariantInit(&result);
    vIdx = MakeInt(index);
    AutoWrap(DISPATCH_PROPERTYGET, &result, pCollection, L"Item", 1, vIdx);
    return (result.vt == VT_DISPATCH) ? result.pdispVal : nullptr;
}

static IDispatch* GetItemStr(IDispatch* pCollection, LPCOLESTR name)
{
    VARIANT result, vName;
    VariantInit(&result);
    vName = MakeBstr(name);
    AutoWrap(DISPATCH_PROPERTYGET, &result, pCollection, L"Item", 1, vName);
    SysFreeString(vName.bstrVal);
    return (result.vt == VT_DISPATCH) ? result.pdispVal : nullptr;
}

// Get Cells(row, col)
static IDispatch* GetCells(IDispatch* pSheet, int row, int col)
{
    VARIANT result;
    VariantInit(&result);
    VARIANT vr = MakeInt(row), vc = MakeInt(col);
    AutoWrap(DISPATCH_PROPERTYGET, &result, pSheet, L"Cells", 2, vr, vc);
    return (result.vt == VT_DISPATCH) ? result.pdispVal : nullptr;
}

// Get Range("A1") or Range("A1:B2")
static IDispatch* GetRange(IDispatch* pSheet, LPCOLESTR addr)
{
    VARIANT result, vAddr;
    VariantInit(&result);
    vAddr = MakeBstr(addr);
    AutoWrap(DISPATCH_PROPERTYGET, &result, pSheet, L"Range", 1, vAddr);
    SysFreeString(vAddr.bstrVal);
    return (result.vt == VT_DISPATCH) ? result.pdispVal : nullptr;
}

static int GetUsedRowCount(IDispatch* pSheet)
{
    IDispatch* usedRange = GetDispProp(pSheet, L"UsedRange");
    if (!usedRange) return 0;
    IDispatch* rows = GetDispProp(usedRange, L"Rows");
    if (!rows) { usedRange->Release(); return 0; }
    VARIANT vCount;
    VariantInit(&vCount);
    AutoWrap(DISPATCH_PROPERTYGET, &vCount, rows, L"Count", 0);
    int count = (vCount.vt == VT_I4) ? vCount.lVal : 0;
    rows->Release();
    usedRange->Release();
    return count;
}

static int GetUsedColCount(IDispatch* pSheet)
{
    IDispatch* usedRange = GetDispProp(pSheet, L"UsedRange");
    if (!usedRange) return 0;
    IDispatch* cols = GetDispProp(usedRange, L"Columns");
    if (!cols) { usedRange->Release(); return 0; }
    VARIANT vCount;
    VariantInit(&vCount);
    AutoWrap(DISPATCH_PROPERTYGET, &vCount, cols, L"Count", 0);
    int count = (vCount.vt == VT_I4) ? vCount.lVal : 0;
    cols->Release();
    usedRange->Release();
    return count;
}

// ============================================================================
// Excel pivot-table setup (mirrors the PS script's sheets 2-8)
// ============================================================================

// Constants matching PS script
static const int xlRowField  = 1;
static const int xlPageField = 3;
static const int xlDataField = 4;
static const int xlAverage   = -4106;
static const int xlSum       = -4157;
static const int xlCount     = -4112;
static const int xlPercentOfTotal      = 8;
static const int xlPercentRunningTotal = 13;
static const LPCOLESTR mcNumberF  = L"###,###,###,###,###";
static const LPCOLESTR mcPercentF = L"#0.00%";

// Set pivot field properties:  orientation, numberFormat, function, calculation, baseField, name, position
static void SetPivotField(IDispatch* pField, int orientation, LPCOLESTR numFmt,
                          int function, int calculation, LPCOLESTR baseField,
                          LPCOLESTR name, int position = 0)
{
    if (!pField) return;
    if (orientation)    PutPropInt(pField, L"Orientation", orientation);
    if (numFmt)         PutPropStr(pField, L"NumberFormat", numFmt);
    if (function)       PutPropInt(pField, L"Function", function);
    if (calculation)    PutPropInt(pField, L"Calculation", calculation);
    if (baseField)      PutPropStr(pField, L"BaseField", baseField);
    if (name)           PutPropStr(pField, L"Name", name);
    if (position > 0)   PutPropInt(pField, L"Position", position);
}

// Get PivotField from sheet's pivot table
static IDispatch* GetPivotField(IDispatch* pSheet, LPCOLESTR ptName, LPCOLESTR fieldName)
{
    IDispatch* pivotTables = GetDispProp(pSheet, L"PivotTables");
    if (!pivotTables) return nullptr;
    IDispatch* pt = GetItemStr(pivotTables, ptName);
    if (!pt) { pivotTables->Release(); return nullptr; }
    VARIANT result, vName;
    VariantInit(&result);
    vName = MakeBstr(fieldName);
    AutoWrap(DISPATCH_METHOD, &result, pt, L"PivotFields", 1, vName);
    SysFreeString(vName.bstrVal);
    pt->Release();
    pivotTables->Release();
    return (result.vt == VT_DISPATCH) ? result.pdispVal : nullptr;
}

// Format a pivot table sheet (freeze panes, column widths, name)
static void FormatPivotSheet(IDispatch* pSheet, LPCOLESTR ptName,
                             int colWidths[], int colCount,
                             LPCOLESTR r3c1Text, LPCOLESTR sheetName)
{
    // Set column widths
    IDispatch* columns = GetDispProp(pSheet, L"Columns");
    if (columns)
    {
        for (int i = 0; i < colCount; i++)
        {
            if (colWidths[i] > 0)
            {
                IDispatch* col = GetItem(columns, i + 1);
                if (col)
                {
                    PutPropInt(col, L"ColumnWidth", colWidths[i]);
                    col->Release();
                }
            }
        }
        columns->Release();
    }

    // HasAutoFormat = false on pivot table
    IDispatch* pivotTables = GetDispProp(pSheet, L"PivotTables");
    if (pivotTables)
    {
        IDispatch* pt = GetItemStr(pivotTables, ptName);
        if (pt)
        {
            PutPropBool(pt, L"HasAutoFormat", false);
            pt->Release();
        }
        pivotTables->Release();
    }

    // Freeze panes at row 3, col 2
    IDispatch* app = GetDispProp(pSheet, L"Application");
    if (app)
    {
        IDispatch* wnd = GetDispProp(app, L"ActiveWindow");
        if (wnd)
        {
            PutPropInt(wnd, L"SplitRow", 3);
            PutPropInt(wnd, L"SplitColumn", 2);
            PutPropBool(wnd, L"FreezePanes", true);
            wnd->Release();
        }
        app->Release();
    }

    // Set cell text
    IDispatch* cell11 = GetCells(pSheet, 1, 1);
    if (cell11)
    {
        PutPropStr(cell11, L"Value", L"LDAPServer filter");
        cell11->Release();
    }
    IDispatch* cell31 = GetCells(pSheet, 3, 1);
    if (cell31)
    {
        PutPropStr(cell31, L"Value", r3c1Text);
        cell31->Release();
    }

    // Rename sheet
    PutPropStr(pSheet, L"Name", sheetName);
}

// ============================================================================
// Create Excel workbook with pivot tables
// ============================================================================

static void ImportCsvToExcel(const std::wstring& eventPath,
                             const std::vector<std::wstring>& csvFiles)
{
    wprintf(L"Import csv to excel.\n");

    CoInitialize(nullptr);

    // Create Excel Application
    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);
    if (FAILED(hr))
    {
        wprintf(L"  Error: Excel not found. CSV files are available for manual import.\n");
        CoUninitialize();
        return;
    }

    IDispatch* pExcel = nullptr;
    hr = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pExcel);
    if (FAILED(hr) || !pExcel)
    {
        wprintf(L"  Error: Cannot start Excel (0x%08X).\n", hr);
        CoUninitialize();
        return;
    }

    // Workbooks.Add()
    IDispatch* pWorkbooks = GetDispProp(pExcel, L"Workbooks");
    if (!pWorkbooks) { pExcel->Release(); CoUninitialize(); return; }

    IDispatch* pWb = CallMethod(pWorkbooks, L"Add", 0);
    if (!pWb) { pWorkbooks->Release(); pExcel->Release(); CoUninitialize(); return; }

    // Get Sheet1
    IDispatch* pSheets = GetDispProp(pWb, L"Worksheets");
    IDispatch* pSheet1 = GetItem(pSheets, 1);

    // Import each CSV via QueryTables
    int currentRow = 1;
    for (size_t fi = 0; fi < csvFiles.size(); fi++)
    {
        std::wstring connStr = L"TEXT;" + eventPath + L"\\" + csvFiles[fi];

        // Get range for insertion
        WCHAR cellAddr[32];
        _snwprintf_s(cellAddr, _countof(cellAddr), _TRUNCATE, L"A%d", currentRow);
        IDispatch* pRange = GetRange(pSheet1, cellAddr);
        if (!pRange) continue;

        // QueryTables.Add(connection, range)
        IDispatch* pQTs = GetDispProp(pSheet1, L"QueryTables");
        if (!pQTs) { pRange->Release(); continue; }

        VARIANT result;
        VariantInit(&result);
        VARIANT vConn = MakeBstr(connStr.c_str());
        VARIANT vRange = MakeDispatch(pRange);
        AutoWrap(DISPATCH_METHOD, &result, pQTs, L"Add", 2, vConn, vRange);
        SysFreeString(vConn.bstrVal);

        IDispatch* pConnector = (result.vt == VT_DISPATCH) ? result.pdispVal : nullptr;
        if (pConnector)
        {
            PutPropBool(pConnector, L"TextFileCommaDelimiter", true);
            PutPropInt(pConnector, L"TextFileParseType", 1);
            CallMethod(pConnector, L"Refresh", 0);
            pConnector->Release();
        }

        // Delete header row on 2nd+ files
        if (currentRow != 1)
        {
            IDispatch* pCell = GetCells(pSheet1, currentRow, 1);
            if (pCell)
            {
                IDispatch* pRow = GetDispProp(pCell, L"EntireRow");
                if (pRow)
                {
                    CallMethod(pRow, L"Delete", 0);
                    pRow->Release();
                }
                pCell->Release();
            }
        }

        currentRow = GetUsedRowCount(pSheet1) + 1;

        pQTs->Release();
        pRange->Release();
    }

    wprintf(L"Customizing XLS.\n");

    // ---- Sheet1: RawData - AutoFilter + freeze ----
    {
        IDispatch* rA1 = GetRange(pSheet1, L"A1");
        if (rA1)
        {
            CallMethod(rA1, L"AutoFilter", 0);
            rA1->Release();
        }
        IDispatch* app = GetDispProp(pSheet1, L"Application");
        if (app)
        {
            IDispatch* wnd = GetDispProp(app, L"ActiveWindow");
            if (wnd)
            {
                PutPropInt(wnd, L"SplitRow", 1);
                PutPropBool(wnd, L"FreezePanes", true);
                wnd->Release();
            }
            app->Release();
        }

        // Number format for numeric columns J,K,M,N,O,P,Q,R
        IDispatch* cols = GetDispProp(pSheet1, L"Columns");
        if (cols)
        {
            int numCols[] = { 10, 11, 13, 14, 15, 16, 17, 18 }; // J=10..R=18
            for (int c : numCols)
            {
                IDispatch* col = GetItem(cols, c);
                if (col)
                {
                    PutPropStr(col, L"NumberFormat", mcNumberF);
                    col->Release();
                }
            }
            cols->Release();
        }
    }

    int totalRows = GetUsedRowCount(pSheet1);
    int totalCols = GetUsedColCount(pSheet1);

    // Build data source reference
    WCHAR dataSource[128];
    _snwprintf_s(dataSource, _countof(dataSource), _TRUNCATE,
                 L"Sheet1!R1C1:R%dC%d", totalRows, totalCols);

    // ---- Sheet2: PivotTable1 - StartingNode grouping ----
    {
        IDispatch* pSheet2 = CallMethod(pSheets, L"Add", 0);
        if (pSheet2)
        {
            IDispatch* pivotCaches = GetDispProp(pWb, L"PivotCaches");
            if (pivotCaches)
            {
                VARIANT vSrc = MakeBstr(dataSource);
                VARIANT vType = MakeInt(1); // xlDatabase
                VARIANT vVer = MakeInt(5);  // xlPivotTableVersion15
                IDispatch* pCache = nullptr;
                VARIANT result;
                VariantInit(&result);
                AutoWrap(DISPATCH_METHOD, &result, pivotCaches, L"Create", 3, vType, vSrc, vVer);
                SysFreeString(vSrc.bstrVal);
                pCache = (result.vt == VT_DISPATCH) ? result.pdispVal : nullptr;

                if (pCache)
                {
                    VARIANT vDest = MakeBstr(L"Sheet2!R1C1");
                    CallMethod(pCache, L"CreatePivotTable", 1, vDest);
                    SysFreeString(vDest.bstrVal);

                    // Configure fields
                    IDispatch* pf;
                    pf = GetPivotField(pSheet2, L"PivotTable1", L"LDAPServer");
                    SetPivotField(pf, xlPageField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();

                    pf = GetPivotField(pSheet2, L"PivotTable1", L"StartingNode");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();

                    pf = GetPivotField(pSheet2, L"PivotTable1", L"Filter");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();

                    pf = GetPivotField(pSheet2, L"PivotTable1", L"ClientIP");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();

                    pf = GetPivotField(pSheet2, L"PivotTable1", L"TimeGenerated");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();

                    pf = GetPivotField(pSheet2, L"PivotTable1", L"ClientIP");
                    SetPivotField(pf, xlDataField, mcNumberF, 0, 0, nullptr, L"Search Count", 1); if(pf) pf->Release();

                    pf = GetPivotField(pSheet2, L"PivotTable1", L"SearchTimeMS");
                    SetPivotField(pf, xlDataField, mcNumberF, xlAverage, 0, nullptr, L"AvgSearchTime", 2); if(pf) pf->Release();

                    pf = GetPivotField(pSheet2, L"PivotTable1", L"ClientIP");
                    SetPivotField(pf, xlDataField, mcPercentF, 0, xlPercentOfTotal, nullptr, L"%GrandTotal", 3); if(pf) pf->Release();

                    pf = GetPivotField(pSheet2, L"PivotTable1", L"ClientIP");
                    SetPivotField(pf, xlDataField, mcPercentF, 0, xlPercentRunningTotal, L"StartingNode", L"%RunningTotal", 4); if(pf) pf->Release();

                    int widths[] = { 60, 12, 14, 12, 14, 0, 0 };
                    FormatPivotSheet(pSheet2, L"PivotTable1", widths, 7,
                                     L"StartingNode grouping", L"2.TopIP-StartingNode");

                    pCache->Release();
                }
                pivotCaches->Release();
            }
            pSheet2->Release();
        }
    }

    // ---- Sheet3: PivotTable2 - IP grouping ----
    {
        IDispatch* pSheet3 = CallMethod(pSheets, L"Add", 0);
        if (pSheet3)
        {
            IDispatch* pivotCaches = GetDispProp(pWb, L"PivotCaches");
            if (pivotCaches)
            {
                VARIANT vSrc = MakeBstr(dataSource);
                VARIANT vType = MakeInt(1), vVer = MakeInt(5);
                VARIANT result; VariantInit(&result);
                AutoWrap(DISPATCH_METHOD, &result, pivotCaches, L"Create", 3, vType, vSrc, vVer);
                SysFreeString(vSrc.bstrVal);
                IDispatch* pCache = (result.vt == VT_DISPATCH) ? result.pdispVal : nullptr;
                if (pCache)
                {
                    VARIANT vDest = MakeBstr(L"Sheet3!R1C1");
                    CallMethod(pCache, L"CreatePivotTable", 1, vDest);
                    SysFreeString(vDest.bstrVal);

                    IDispatch* pf;
                    pf = GetPivotField(pSheet3, L"PivotTable2", L"LDAPServer");
                    SetPivotField(pf, xlPageField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet3, L"PivotTable2", L"ClientIP");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet3, L"PivotTable2", L"Filter");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet3, L"PivotTable2", L"TimeGenerated");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();

                    pf = GetPivotField(pSheet3, L"PivotTable2", L"ClientIP");
                    SetPivotField(pf, xlDataField, mcNumberF, 0, 0, nullptr, L"Search Count", 1); if(pf) pf->Release();
                    pf = GetPivotField(pSheet3, L"PivotTable2", L"SearchTimeMS");
                    SetPivotField(pf, xlDataField, mcNumberF, xlAverage, 0, nullptr, L"AvgSearchTime (MS)", 2); if(pf) pf->Release();
                    pf = GetPivotField(pSheet3, L"PivotTable2", L"ClientIP");
                    SetPivotField(pf, xlDataField, mcPercentF, 0, xlPercentOfTotal, nullptr, L"%GrandTotal", 3); if(pf) pf->Release();
                    pf = GetPivotField(pSheet3, L"PivotTable2", L"ClientIP");
                    SetPivotField(pf, xlDataField, mcPercentF, 0, xlPercentRunningTotal, L"ClientIP", L"%RunningTotal", 4); if(pf) pf->Release();

                    int widths[] = { 60, 12, 19, 12, 14, 0, 0 };
                    FormatPivotSheet(pSheet3, L"PivotTable2", widths, 7,
                                     L"IP grouping", L"3.TopIP");
                    pCache->Release();
                }
                pivotCaches->Release();
            }
            pSheet3->Release();
        }
    }

    // ---- Sheet4: PivotTable3 - Filter grouping ----
    {
        IDispatch* pSheet4 = CallMethod(pSheets, L"Add", 0);
        if (pSheet4)
        {
            IDispatch* pivotCaches = GetDispProp(pWb, L"PivotCaches");
            if (pivotCaches)
            {
                VARIANT vSrc = MakeBstr(dataSource);
                VARIANT vType = MakeInt(1), vVer = MakeInt(5);
                VARIANT result; VariantInit(&result);
                AutoWrap(DISPATCH_METHOD, &result, pivotCaches, L"Create", 3, vType, vSrc, vVer);
                SysFreeString(vSrc.bstrVal);
                IDispatch* pCache = (result.vt == VT_DISPATCH) ? result.pdispVal : nullptr;
                if (pCache)
                {
                    VARIANT vDest = MakeBstr(L"Sheet4!R1C1");
                    CallMethod(pCache, L"CreatePivotTable", 1, vDest);
                    SysFreeString(vDest.bstrVal);

                    IDispatch* pf;
                    pf = GetPivotField(pSheet4, L"PivotTable3", L"LDAPServer");
                    SetPivotField(pf, xlPageField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet4, L"PivotTable3", L"Filter");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet4, L"PivotTable3", L"ClientIP");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet4, L"PivotTable3", L"TimeGenerated");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();

                    pf = GetPivotField(pSheet4, L"PivotTable3", L"ClientIP");
                    SetPivotField(pf, xlDataField, mcNumberF, 0, 0, nullptr, L"Search Count", 1); if(pf) pf->Release();
                    pf = GetPivotField(pSheet4, L"PivotTable3", L"SearchTimeMS");
                    SetPivotField(pf, xlDataField, mcNumberF, xlAverage, 0, nullptr, L"AvgSearchTime (MS)", 2); if(pf) pf->Release();
                    pf = GetPivotField(pSheet4, L"PivotTable3", L"ClientIP");
                    SetPivotField(pf, xlDataField, mcPercentF, 0, xlPercentOfTotal, nullptr, L"%GrandTotal", 3); if(pf) pf->Release();
                    pf = GetPivotField(pSheet4, L"PivotTable3", L"ClientIP");
                    SetPivotField(pf, xlDataField, mcPercentF, 0, xlPercentRunningTotal, L"Filter", L"%RunningTotal", 4); if(pf) pf->Release();

                    int widths[] = { 60, 12, 19, 12, 14, 0, 0 };
                    FormatPivotSheet(pSheet4, L"PivotTable3", widths, 7,
                                     L"Filter grouping", L"4.TopIP-Filters");
                    pCache->Release();
                }
                pivotCaches->Release();
            }
            pSheet4->Release();
        }
    }

    // ---- Sheet5: PivotTable4 - TopTime by IP ----
    {
        IDispatch* pSheet5 = CallMethod(pSheets, L"Add", 0);
        if (pSheet5)
        {
            IDispatch* pivotCaches = GetDispProp(pWb, L"PivotCaches");
            if (pivotCaches)
            {
                VARIANT vSrc = MakeBstr(dataSource);
                VARIANT vType = MakeInt(1), vVer = MakeInt(5);
                VARIANT result; VariantInit(&result);
                AutoWrap(DISPATCH_METHOD, &result, pivotCaches, L"Create", 3, vType, vSrc, vVer);
                SysFreeString(vSrc.bstrVal);
                IDispatch* pCache = (result.vt == VT_DISPATCH) ? result.pdispVal : nullptr;
                if (pCache)
                {
                    VARIANT vDest = MakeBstr(L"Sheet5!R1C1");
                    CallMethod(pCache, L"CreatePivotTable", 1, vDest);
                    SysFreeString(vDest.bstrVal);

                    IDispatch* pf;
                    pf = GetPivotField(pSheet5, L"PivotTable4", L"LDAPServer");
                    SetPivotField(pf, xlPageField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet5, L"PivotTable4", L"ClientIP");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet5, L"PivotTable4", L"Filter");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet5, L"PivotTable4", L"TimeGenerated");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();

                    pf = GetPivotField(pSheet5, L"PivotTable4", L"SearchTimeMS");
                    SetPivotField(pf, xlDataField, mcNumberF, xlSum, 0, nullptr, L"Total SearchTime (MS)"); if(pf) pf->Release();
                    pf = GetPivotField(pSheet5, L"PivotTable4", L"ClientIP");
                    SetPivotField(pf, xlDataField, mcNumberF, 0, 0, nullptr, L"Search Count"); if(pf) pf->Release();
                    pf = GetPivotField(pSheet5, L"PivotTable4", L"SearchTimeMS");
                    SetPivotField(pf, xlDataField, mcNumberF, xlAverage, 0, nullptr, L"AvgSearchTime (MS)"); if(pf) pf->Release();
                    pf = GetPivotField(pSheet5, L"PivotTable4", L"SearchTimeMS");
                    SetPivotField(pf, xlDataField, mcPercentF, 0, xlPercentOfTotal, nullptr, L"%GrandTotal (MS)"); if(pf) pf->Release();
                    pf = GetPivotField(pSheet5, L"PivotTable4", L"SearchTimeMS");
                    SetPivotField(pf, xlDataField, mcPercentF, 0, xlPercentRunningTotal, L"ClientIP", L"%RunningTotal (Ms)"); if(pf) pf->Release();

                    int widths[] = { 50, 21, 12, 19, 17, 19, 0 };
                    FormatPivotSheet(pSheet5, L"PivotTable4", widths, 7,
                                     L"IP grouping", L"5.TopTime-IP");
                    pCache->Release();
                }
                pivotCaches->Release();
            }
            pSheet5->Release();
        }
    }

    // ---- Sheet6: PivotTable5 - TopTime by Filters ----
    {
        IDispatch* pSheet6 = CallMethod(pSheets, L"Add", 0);
        if (pSheet6)
        {
            IDispatch* pivotCaches = GetDispProp(pWb, L"PivotCaches");
            if (pivotCaches)
            {
                VARIANT vSrc = MakeBstr(dataSource);
                VARIANT vType = MakeInt(1), vVer = MakeInt(5);
                VARIANT result; VariantInit(&result);
                AutoWrap(DISPATCH_METHOD, &result, pivotCaches, L"Create", 3, vType, vSrc, vVer);
                SysFreeString(vSrc.bstrVal);
                IDispatch* pCache = (result.vt == VT_DISPATCH) ? result.pdispVal : nullptr;
                if (pCache)
                {
                    VARIANT vDest = MakeBstr(L"Sheet6!R1C1");
                    CallMethod(pCache, L"CreatePivotTable", 1, vDest);
                    SysFreeString(vDest.bstrVal);

                    IDispatch* pf;
                    pf = GetPivotField(pSheet6, L"PivotTable5", L"LDAPServer");
                    SetPivotField(pf, xlPageField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet6, L"PivotTable5", L"Filter");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet6, L"PivotTable5", L"ClientIP");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet6, L"PivotTable5", L"TimeGenerated");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();

                    pf = GetPivotField(pSheet6, L"PivotTable5", L"SearchTimeMS");
                    SetPivotField(pf, xlDataField, mcNumberF, xlSum, 0, nullptr, L"Total SearchTime (MS)"); if(pf) pf->Release();
                    pf = GetPivotField(pSheet6, L"PivotTable5", L"ClientIP");
                    SetPivotField(pf, xlDataField, mcNumberF, 0, 0, nullptr, L"Search Count"); if(pf) pf->Release();
                    pf = GetPivotField(pSheet6, L"PivotTable5", L"SearchTimeMS");
                    SetPivotField(pf, xlDataField, mcNumberF, xlAverage, 0, nullptr, L"AvgSearchTime (MS)"); if(pf) pf->Release();
                    pf = GetPivotField(pSheet6, L"PivotTable5", L"SearchTimeMS");
                    SetPivotField(pf, xlDataField, mcPercentF, 0, xlPercentOfTotal, nullptr, L"%GrandTotal (MS)"); if(pf) pf->Release();
                    pf = GetPivotField(pSheet6, L"PivotTable5", L"SearchTimeMS");
                    SetPivotField(pf, xlDataField, mcPercentF, 0, xlPercentRunningTotal, L"Filter", L"%RunningTotal (MS)"); if(pf) pf->Release();

                    int widths[] = { 50, 21, 12, 19, 17, 19, 0 };
                    FormatPivotSheet(pSheet6, L"PivotTable5", widths, 7,
                                     L"Filter grouping", L"6.TopTime-Filters");
                    pCache->Release();
                }
                pivotCaches->Release();
            }
            pSheet6->Release();
        }
    }

    // ---- Sheet7: PivotTable6 - TimeRanks ----
    {
        IDispatch* pSheet7 = CallMethod(pSheets, L"Add", 0);
        if (pSheet7)
        {
            IDispatch* pivotCaches = GetDispProp(pWb, L"PivotCaches");
            if (pivotCaches)
            {
                VARIANT vSrc = MakeBstr(dataSource);
                VARIANT vType = MakeInt(1), vVer = MakeInt(5);
                VARIANT result; VariantInit(&result);
                AutoWrap(DISPATCH_METHOD, &result, pivotCaches, L"Create", 3, vType, vSrc, vVer);
                SysFreeString(vSrc.bstrVal);
                IDispatch* pCache = (result.vt == VT_DISPATCH) ? result.pdispVal : nullptr;
                if (pCache)
                {
                    VARIANT vDest = MakeBstr(L"Sheet7!R1C1");
                    CallMethod(pCache, L"CreatePivotTable", 1, vDest);
                    SysFreeString(vDest.bstrVal);

                    IDispatch* pf;
                    pf = GetPivotField(pSheet7, L"PivotTable6", L"LDAPServer");
                    SetPivotField(pf, xlPageField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet7, L"PivotTable6", L"SearchTimeMS");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet7, L"PivotTable6", L"Filter");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet7, L"PivotTable6", L"ClientIP");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();
                    pf = GetPivotField(pSheet7, L"PivotTable6", L"TimeGenerated");
                    SetPivotField(pf, xlRowField, nullptr, 0, 0, nullptr, nullptr); if(pf) pf->Release();

                    pf = GetPivotField(pSheet7, L"PivotTable6", L"ClientIP");
                    SetPivotField(pf, xlDataField, mcNumberF, 0, 0, nullptr, L"Search Count"); if(pf) pf->Release();
                    pf = GetPivotField(pSheet7, L"PivotTable6", L"SearchTimeMS");
                    SetPivotField(pf, xlDataField, mcPercentF, xlSum, xlPercentOfTotal, nullptr, L"%GrandTotal (MS)"); if(pf) pf->Release();
                    pf = GetPivotField(pSheet7, L"PivotTable6", L"SearchTimeMS");
                    SetPivotField(pf, xlDataField, mcPercentF, xlSum, xlPercentRunningTotal, L"SearchTimeMS", L"%RunningTotal (MS)"); if(pf) pf->Release();

                    int widths[] = { 60, 12, 17, 19, 0, 0, 0 };
                    FormatPivotSheet(pSheet7, L"PivotTable6", widths, 7,
                                     L"SearchTime (MS) grouping", L"7.TimeRanks");
                    pCache->Release();
                }
                pivotCaches->Release();
            }
            pSheet7->Release();
        }
    }

    // ---- Sheet8: Sandbox ----
    {
        IDispatch* pSheet8 = CallMethod(pSheets, L"Add", 0);
        if (pSheet8)
        {
            PutPropStr(pSheet8, L"Name", L"8.SandBox");
            pSheet8->Release();
        }
    }

    // Rename Sheet1
    PutPropStr(pSheet1, L"Name", L"1.RawData");

    // Set tab colors (35=light green, 36=light yellow)
    // Tab.ColorIndex = 35 for sheets 2-4, 36 for sheets 5-7
    // Tab.Color = 8109667 for sheet 8
    for (int si = 1; si <= 8; si++)
    {
        IDispatch* pSh = GetItem(pSheets, si);
        if (!pSh) continue;
        IDispatch* tab = GetDispProp(pSh, L"Tab");
        if (tab)
        {
            VARIANT vName;
            VariantInit(&vName);
            AutoWrap(DISPATCH_PROPERTYGET, &vName, pSh, L"Name", 0);
            if (vName.vt == VT_BSTR)
            {
                std::wstring nm = vName.bstrVal;
                if (nm[0] == L'2' || nm[0] == L'3' || nm[0] == L'4')
                    PutPropInt(tab, L"ColorIndex", 35);
                else if (nm[0] == L'5' || nm[0] == L'6' || nm[0] == L'7')
                    PutPropInt(tab, L"ColorIndex", 36);
                else if (nm[0] == L'8')
                    PutPropInt(tab, L"Color", 8109667);
                SysFreeString(vName.bstrVal);
            }
            tab->Release();
        }
        pSh->Release();
    }

    // Sort sheets by name (ascending)
    {
        VARIANT vCount;
        VariantInit(&vCount);
        AutoWrap(DISPATCH_PROPERTYGET, &vCount, pSheets, L"Count", 0);
        int sheetCount = (vCount.vt == VT_I4) ? vCount.lVal : 0;

        // Simple bubble sort of sheet names
        for (int pass = 0; pass < sheetCount - 1; pass++)
        {
            for (int si = 1; si < sheetCount; si++)
            {
                IDispatch* s1 = GetItem(pSheets, si);
                IDispatch* s2 = GetItem(pSheets, si + 1);
                if (s1 && s2)
                {
                    VARIANT n1, n2;
                    VariantInit(&n1); VariantInit(&n2);
                    AutoWrap(DISPATCH_PROPERTYGET, &n1, s1, L"Name", 0);
                    AutoWrap(DISPATCH_PROPERTYGET, &n2, s2, L"Name", 0);
                    if (n1.vt == VT_BSTR && n2.vt == VT_BSTR)
                    {
                        if (wcscmp(n1.bstrVal, n2.bstrVal) > 0)
                        {
                            // Move s1 after s2
                            VARIANT vAfter = MakeDispatch(s2);
                            AutoWrap(DISPATCH_METHOD, nullptr, s1, L"Move", 1, vAfter);
                        }
                    }
                    if (n1.vt == VT_BSTR) SysFreeString(n1.bstrVal);
                    if (n2.vt == VT_BSTR) SysFreeString(n2.bstrVal);
                }
                if (s1) s1->Release();
                if (s2) s2->Release();
            }
        }
    }

    // Activate Sheet1
    {
        IDispatch* pSh1 = GetItemStr(pSheets, L"1.RawData");
        if (pSh1)
        {
            CallMethod(pSh1, L"Activate", 0);
            pSh1->Release();
        }
    }

    // Prompt for file name
    std::wstring fileName = ReadLine(L"Enter a FileName to save extracted event 1644 xlsx:\n");
    if (!fileName.empty())
    {
        std::wstring fullPath = eventPath + L"\\" + fileName;
        wprintf(L"Saving file to %s.xlsx\n", fullPath.c_str());
        VARIANT vPath = MakeBstr(fullPath.c_str());
        AutoWrap(DISPATCH_METHOD, nullptr, pWb, L"SaveAs", 1, vPath);
        SysFreeString(vPath.bstrVal);
    }

    // Prompt to delete CSV
    std::wstring cleanup = ReadLine(L"Delete generated 1644-*.csv? ([Enter]/[Y] to delete, [N] to keep csv)\n");
    if (_wcsicmp(cleanup.c_str(), L"N") != 0)
    {
        for (auto& csv : csvFiles)
        {
            std::wstring csvFullPath = eventPath + L"\\" + csv;
            if (DeleteFileW(csvFullPath.c_str()))
                wprintf(L"\t%s deleted.\n", csv.c_str());
        }
    }

    // Make Excel visible
    PutPropBool(pExcel, L"Visible", true);

    // Cleanup COM
    if (pSheet1) pSheet1->Release();
    if (pSheets) pSheets->Release();
    if (pWb) pWb->Release();
    if (pWorkbooks) pWorkbooks->Release();
    pExcel->Release();
    CoUninitialize();
}

// ============================================================================
// Main
// ============================================================================

int wmain(int /*argc*/, wchar_t* /*argv*/[])
{
    wprintf(L"Event1644Reader: See https://support.microsoft.com/en-us/kb/3060643 "
            L"for sample walk through and pivotTable tips.\n");

    std::wstring scriptPath = GetExeDirectory();

    std::wstring eventPath = ReadLine(
        L"Enter local, mapped or UNC path to Evtx(s). Remove trailing blank.\n"
        L"For Example (c:\\CaseData)\n"
        L"Or press [Enter] if evtx is in the program folder.\n");

    if (eventPath.empty())
    {
        eventPath = scriptPath;
        wprintf(L"\tScanning event logs in %s\n", eventPath.c_str());
    }

    // Scan for .evtx files
    std::wstring searchPattern = eventPath + L"\\*.evtx";
    WIN32_FIND_DATAW fd;
    HANDLE hFind = FindFirstFileW(searchPattern.c_str(), &fd);

    std::vector<std::wstring> csvFiles;

    if (hFind == INVALID_HANDLE_VALUE)
    {
        wprintf(L"\tNo .evtx files found in %s\n", eventPath.c_str());
    }
    else
    {
        do
        {
            if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
                continue;

            std::wstring fileName = fd.cFileName;
            wprintf(L"Reading %s\n", fileName.c_str());

            std::wstring evtxPath = eventPath + L"\\" + fileName;

            // Get base name (without .evtx)
            std::wstring baseName = fileName;
            size_t dot = baseName.rfind(L'.');
            if (dot != std::wstring::npos)
                baseName = baseName.substr(0, dot);

            std::wstring csvName = L"1644-" + baseName + L".csv";
            std::wstring csvPath = eventPath + L"\\" + csvName;

            int count = ProcessEvtxFile(evtxPath, csvPath);
            if (count > 0)
            {
                wprintf(L"\tEvent 1644 found (%d events), generated %s\n", count, csvName.c_str());
                csvFiles.push_back(csvName);
            }
            else
            {
                wprintf(L"\tNo event 1644 found.\n");
            }
        } while (FindNextFileW(hFind, &fd));

        FindClose(hFind);
    }

    // Import to Excel if any CSV generated
    if (!csvFiles.empty())
    {
        ImportCsvToExcel(eventPath, csvFiles);
    }
    else
    {
        wprintf(L"\tNo event 1644 found in specified directory. %s\n", eventPath.c_str());
    }

    wprintf(L"Script completed.\n");
    return 0;
}

/*
cmd /c '"C:\Program Files\Microsoft Visual Studio\18\Community\VC\Auxiliary\Build\vcvarsall.bat" x64 && cl /EHsc /W4 /DUNICODE /D_UNICODE Event1644Reader.cpp ole32.lib oleaut32.lib wevtapi.lib shlwapi.lib'
cmd /c "\"C:\Program Files\Microsoft Visual Studio\18\Community\VC\Auxiliary\Build\vcvarsall.bat\" x64 && cl /EHsc /W4 /DUNICODE /D_UNICODE Event1644Reader.cpp ole32.lib oleaut32.lib wevtapi.lib shlwapi.lib"
*/