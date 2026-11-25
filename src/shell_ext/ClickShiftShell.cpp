// ClickShiftShell.cpp — Windows 11 IExplorerCommand + IEnumExplorerCommand shell extension
// Provides the expandable "Convert with ClickShift >" menu in the primary Windows 11 right-click menu.

#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN

#include <windows.h>
#include <shobjidl.h>
#include <shlwapi.h>
#include <shellapi.h>
#include <strsafe.h>
#include <vector>
#include <string>
#include <algorithm>
#include <cwctype>

static HMODULE g_hModule = nullptr;
static LONG g_cDllRef = 0;
static CRITICAL_SECTION g_cs;
static bool g_csInit = false;
static std::wstring g_lastExt;
static std::wstring g_cachedExe;
static std::wstring g_cachedIcon;
static ULONGLONG g_pathsTick = 0;
static const ULONGLONG kPathCacheTtlMs = 30000;

// {A7C2D814-3E5F-4B9C-9A12-8D4E6F0C3A5E}
static const CLSID CLSID_ClickShiftExplorerCommand = {
    0xa7c2d814, 0x3e5f, 0x4b9c, {0x9a, 0x12, 0x8d, 0x4e, 0x6f, 0x0c, 0x3a, 0x5e}
};

struct ExtRule {
    const wchar_t* ext;
    const wchar_t* defaultTarget;
    const wchar_t* targets[16];
};

static const ExtRule kRules[] = {
    {L"pdf",  L"docx", {L"docx", L"pptx", L"txt", L"md", L"html", L"csv", L"png", L"jpg", L"webp", L"odt", L"rtf", nullptr}},
    {L"docx", L"pdf",  {L"pdf", L"pptx", L"txt", L"md", L"html", L"rtf", L"odt", L"png", L"jpg", L"webp", nullptr}},
    {L"doc",  L"docx", {L"pdf", L"docx", L"pptx", L"txt", L"md", L"html", L"rtf", L"odt", L"png", L"jpg", L"webp", nullptr}},
    {L"odt",  L"docx", {L"docx", L"pdf", L"pptx", L"txt", L"md", L"html", L"rtf", L"png", L"jpg", L"webp", nullptr}},
    {L"rtf",  L"docx", {L"docx", L"pdf", L"pptx", L"txt", L"md", L"html", L"odt", L"png", L"jpg", L"webp", nullptr}},
    {L"xlsx", L"pdf",  {L"pdf", L"csv", L"ods", L"txt", L"html", L"png", L"jpg", L"webp", nullptr}},
    {L"xls",  L"xlsx", {L"xlsx", L"pdf", L"csv", L"ods", L"txt", L"html", L"png", L"jpg", L"webp", nullptr}},
    {L"ods",  L"xlsx", {L"xlsx", L"pdf", L"csv", L"txt", L"html", L"png", L"jpg", L"webp", nullptr}},
    {L"pptx", L"pdf",  {L"pdf", L"docx", L"odp", L"png", L"jpg", L"webp", nullptr}},
    {L"ppt",  L"pptx", {L"pptx", L"pdf", L"docx", L"odp", L"png", L"jpg", L"webp", nullptr}},
    {L"odp",  L"pptx", {L"pptx", L"pdf", L"docx", L"png", L"jpg", L"webp", nullptr}},
    {L"csv",  L"xlsx", {L"xlsx", L"pdf", L"ods", L"txt", L"html", L"docx", nullptr}},
    {L"tsv",  L"xlsx", {L"xlsx", L"pdf", L"csv", L"ods", L"txt", L"html", nullptr}},
    {L"txt",  L"docx", {L"docx", L"pdf", L"md", L"html", L"rtf", L"odt", nullptr}},
    {L"md",   L"docx", {L"docx", L"pdf", L"html", L"txt", L"rtf", L"odt", nullptr}},
    {L"markdown", L"docx", {L"docx", L"pdf", L"html", L"txt", L"rtf", L"odt", nullptr}},
    {L"rst",  L"docx", {L"docx", L"pdf", L"html", L"txt", L"md", nullptr}},
    {L"log",  L"txt",  {L"txt", L"docx", L"pdf", L"html", nullptr}},
    {L"html", L"docx", {L"docx", L"pdf", L"txt", L"md", L"odt", L"rtf", nullptr}},
    {L"htm",  L"docx", {L"docx", L"pdf", L"txt", L"md", L"odt", L"rtf", nullptr}},
    {L"json", L"txt",  {L"txt", L"docx", L"pdf", L"html", L"md", nullptr}},
    {L"xml",  L"txt",  {L"txt", L"docx", L"pdf", L"html", L"md", nullptr}},
    {L"yaml", L"txt",  {L"txt", L"docx", L"pdf", L"html", L"md", nullptr}},
    {L"yml",  L"txt",  {L"txt", L"docx", L"pdf", L"html", L"md", nullptr}},
    {L"ini",  L"txt",  {L"txt", L"docx", L"pdf", L"html", nullptr}},
    {L"cfg",  L"txt",  {L"txt", L"docx", L"pdf", L"html", nullptr}},
    {L"epub", L"docx", {L"docx", L"pdf", L"txt", L"html", L"md", L"png", L"jpg", L"webp", nullptr}},
    {L"png",  L"webp", {L"webp", L"jpg", L"pdf", L"docx", L"pptx", L"avif", L"jxl", L"tiff", L"bmp", L"ico", L"gif", nullptr}},
    {L"jpg",  L"webp", {L"webp", L"png", L"pdf", L"docx", L"pptx", L"avif", L"jxl", L"tiff", L"bmp", L"ico", L"gif", nullptr}},
    {L"jpeg", L"webp", {L"webp", L"png", L"pdf", L"docx", L"pptx", L"avif", L"jxl", L"tiff", L"bmp", L"ico", L"gif", nullptr}},
    {L"webp", L"png",  {L"png", L"jpg", L"pdf", L"docx", L"pptx", L"avif", L"jxl", L"tiff", L"bmp", L"ico", L"gif", nullptr}},
    {L"gif",  L"webp", {L"webp", L"png", L"jpg", L"pdf", L"docx", L"pptx", L"avif", L"jxl", L"tiff", L"bmp", L"ico", nullptr}},
    {L"bmp",  L"png",  {L"png", L"webp", L"jpg", L"pdf", L"docx", L"pptx", L"avif", L"jxl", L"tiff", L"ico", L"gif", nullptr}},
    {L"tif",  L"png",  {L"png", L"webp", L"jpg", L"pdf", L"docx", L"pptx", L"avif", L"jxl", L"bmp", L"ico", L"gif", nullptr}},
    {L"tiff", L"png",  {L"png", L"webp", L"jpg", L"pdf", L"docx", L"pptx", L"avif", L"jxl", L"bmp", L"ico", L"gif", nullptr}},
    {L"ico",  L"png",  {L"png", L"webp", L"jpg", L"pdf", L"docx", L"pptx", L"bmp", nullptr}},
    {L"avif", L"png",  {L"png", L"jpg", L"webp", L"pdf", L"docx", L"pptx", L"jxl", L"tiff", L"bmp", L"ico", L"gif", nullptr}},
    {L"heic", L"jpg",  {L"jpg", L"png", L"webp", L"pdf", L"docx", L"pptx", L"avif", L"jxl", L"tiff", L"bmp", nullptr}},
    {L"heif", L"jpg",  {L"jpg", L"png", L"webp", L"pdf", L"docx", L"pptx", L"avif", L"jxl", L"tiff", L"bmp", nullptr}},
    {L"jxl",  L"png",  {L"png", L"jpg", L"webp", L"pdf", L"docx", L"pptx", L"avif", L"tiff", L"bmp", L"ico", nullptr}},
    {L"jp2",  L"png",  {L"png", L"jpg", L"webp", L"pdf", L"docx", L"pptx", L"tiff", L"bmp", nullptr}},
    {L"ppm",  L"png",  {L"png", L"webp", L"jpg", L"pdf", L"docx", L"pptx", L"tiff", L"bmp", nullptr}},
    {L"pgm",  L"png",  {L"png", L"webp", L"jpg", L"pdf", L"docx", L"pptx", L"tiff", L"bmp", nullptr}},
    {L"pbm",  L"png",  {L"png", L"webp", L"jpg", L"pdf", L"docx", L"pptx", L"tiff", L"bmp", nullptr}},
    {L"wmf",  L"png",  {L"png", L"webp", L"jpg", L"pdf", L"docx", L"pptx", nullptr}},
    {L"emf",  L"png",  {L"png", L"webp", L"jpg", L"pdf", L"docx", L"pptx", nullptr}},
    {L"xps",  L"pdf",  {L"pdf", L"docx", L"png", L"jpg", nullptr}},
    {L"mht",  L"docx", {L"docx", L"pdf", L"html", L"txt", L"md", nullptr}},
    {L"wpd",  L"docx", {L"docx", L"pdf", L"txt", L"rtf", L"odt", nullptr}},
    {L"srt",  L"txt",  {L"txt", L"docx", L"pdf", L"html", L"md", nullptr}},
};

static const wchar_t* kAllTargets[] = {
    L"docx", L"pdf", L"xlsx", L"pptx", L"txt", L"md", L"html", L"csv",
    L"png", L"jpg", L"webp", L"odt", L"ods", L"odp", L"rtf",
    L"avif", L"jxl", L"tiff", L"bmp", L"ico", L"gif", nullptr
};

static std::wstring ToLower(const std::wstring& s) {
    std::wstring out = s;
    for (wchar_t& c : out) c = (wchar_t)towlower(c);
    return out;
}

static std::wstring ToUpper(const std::wstring& s) {
    std::wstring out = s;
    for (wchar_t& c : out) c = (wchar_t)towupper(c);
    return out;
}

static const ExtRule* FindRule(const std::wstring& ext) {
    if (ext.empty()) return nullptr;
    for (const auto& r : kRules) {
        if (ext == r.ext) return &r;
    }
    return nullptr;
}

static bool RuleHasTarget(const ExtRule* rule, const std::wstring& target) {
    if (!rule) return true;
    for (int i = 0; rule->targets[i] != nullptr; ++i) {
        if (target == rule->targets[i]) return true;
    }
    return false;
}

static std::wstring ExtractExtFromArray(IShellItemArray* psiArray) {
    if (!psiArray) return L"";
    IShellItem* psi = nullptr;
    if (FAILED(psiArray->GetItemAt(0, &psi)) || !psi) return L"";
    std::wstring ext;
    const SIGDN modes[] = {SIGDN_FILESYSPATH, SIGDN_PARENTRELATIVEPARSING, SIGDN_NORMALDISPLAY};
    for (SIGDN mode : modes) {
        LPWSTR pszPath = nullptr;
        if (SUCCEEDED(psi->GetDisplayName(mode, &pszPath)) && pszPath) {
            LPCWSTR dot = PathFindExtensionW(pszPath);
            if (dot && *dot == L'.' && dot[1] != 0) {
                ext = ToLower(dot + 1);
                CoTaskMemFree(pszPath);
                break;
            }
            CoTaskMemFree(pszPath);
        }
    }
    psi->Release();
    return ext;
}

static void SetCachedExt(const std::wstring& ext) {
    if (ext.empty()) return;
    if (g_csInit) {
        EnterCriticalSection(&g_cs);
        g_lastExt = ext;
        LeaveCriticalSection(&g_cs);
    }
}

static std::wstring GetCachedExt() {
    std::wstring out;
    if (g_csInit) {
        EnterCriticalSection(&g_cs);
        out = g_lastExt;
        LeaveCriticalSection(&g_cs);
    }
    return out;
}

static bool IsValidExeFile(const std::wstring& path) {
    if (path.size() < 5) return false;
    DWORD attrs = GetFileAttributesW(path.c_str());
    if (attrs == INVALID_FILE_ATTRIBUTES || (attrs & FILE_ATTRIBUTE_DIRECTORY)) return false;
    std::wstring tail = ToLower(path.substr(path.size() - 4));
    if (tail != L".exe") return false;
    LPCWSTR fname = PathFindFileNameW(path.c_str());
    if (fname) {
        std::wstring lowName = ToLower(fname);
        if (lowName == L"python.exe" || lowName == L"pythonw.exe" || lowName.rfind(L"python3", 0) == 0) {
            return false;
        }
    }
    return true;
}

static std::wstring FindLatestDownloadsExe() {
    wchar_t userProfile[MAX_PATH];
    userProfile[0] = 0;
    if (GetEnvironmentVariableW(L"USERPROFILE", userProfile, MAX_PATH) == 0) return L"";
    std::wstring dir = std::wstring(userProfile) + L"\\Downloads";
    std::wstring pattern = dir + L"\\ClickShift*.exe";
    WIN32_FIND_DATAW fd = {};
    HANDLE hFind = FindFirstFileW(pattern.c_str(), &fd);
    if (hFind == INVALID_HANDLE_VALUE) return L"";
    std::wstring bestPath;
    ULARGE_INTEGER bestTime = {};
    do {
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue;
        std::wstring candidate = dir + L"\\" + fd.cFileName;
        if (!IsValidExeFile(candidate)) continue;
        ULARGE_INTEGER cur = {};
        cur.LowPart = fd.ftLastWriteTime.dwLowDateTime;
        cur.HighPart = fd.ftLastWriteTime.dwHighDateTime;
        if (bestPath.empty() || cur.QuadPart >= bestTime.QuadPart) {
            bestTime = cur;
            bestPath = candidate;
        }
    } while (FindNextFileW(hFind, &fd));
    FindClose(hFind);
    return bestPath;
}

static std::wstring QuoteWindowsArg(const std::wstring& arg) {
    if (!arg.empty() && arg.find_first_of(L" \t\n\v\"") == std::wstring::npos && arg.back() != L'\\') {
        return arg;
    }
    std::wstring out = L"\"";
    size_t backslashes = 0;
    for (wchar_t ch : arg) {
        if (ch == L'\\') {
            backslashes++;
        } else if (ch == L'"') {
            out.append(backslashes * 2 + 1, L'\\');
            out.push_back(L'"');
            backslashes = 0;
        } else {
            out.append(backslashes, L'\\');
            out.push_back(ch);
            backslashes = 0;
        }
    }
    out.append(backslashes * 2, L'\\');
    out.push_back(L'"');
    return out;
}

static std::wstring WriteFilelistFallback(const std::vector<std::wstring>& files) {
    std::wstring dir;
    wchar_t localApp[MAX_PATH] = {0};
    DWORD nLocal = GetEnvironmentVariableW(L"LOCALAPPDATA", localApp, MAX_PATH);
    if (nLocal > 0 && nLocal < MAX_PATH) {
        std::wstring base = std::wstring(localApp) + L"\\ClickShift";
        CreateDirectoryW(base.c_str(), nullptr);
        dir = base + L"\\staging";
        CreateDirectoryW(dir.c_str(), nullptr);
    } else {
        wchar_t tmp[MAX_PATH] = {0};
        DWORD nTmp = GetTempPathW(MAX_PATH, tmp);
        if (nTmp == 0 || nTmp >= MAX_PATH) return L"";
        std::wstring baseTmp = tmp;
        while (!baseTmp.empty() && baseTmp.back() == L'\\') baseTmp.pop_back();
        dir = baseTmp + L"\\ClickShift";
        CreateDirectoryW(dir.c_str(), nullptr);
    }
    std::wstring path = dir + L"\\batch_" + std::to_wstring(GetCurrentProcessId()) +
                        L"_" + std::to_wstring(GetTickCount64()) + L".filelist";
    HANDLE hFile = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                               FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) return L"";
    std::string utf8Payload;
    for (const auto& f : files) {
        int needed = WideCharToMultiByte(CP_UTF8, 0, f.c_str(), -1, nullptr, 0, nullptr, nullptr);
        if (needed > 1) {
            std::string line(needed - 1, '\0');
            WideCharToMultiByte(CP_UTF8, 0, f.c_str(), -1, &line[0], needed, nullptr, nullptr);
            utf8Payload += line;
            utf8Payload.push_back('\n');
        }
    }
    DWORD written = 0;
    BOOL ok = utf8Payload.empty() || WriteFile(hFile, utf8Payload.data(), (DWORD)utf8Payload.size(), &written, nullptr);
    CloseHandle(hFile);
    if (!ok) {
        DeleteFileW(path.c_str());
        return L"";
    }
    return path;
}

static std::wstring ReadRegString(HKEY root, const wchar_t* subkey, const wchar_t* valueName) {
    wchar_t buf[1024];
    buf[0] = 0;
    DWORD cb = sizeof(buf);
    LONG rc = RegGetValueW(root, subkey, valueName, RRF_RT_REG_SZ, nullptr, buf, &cb);
    if (rc == ERROR_SUCCESS && buf[0]) {
        return std::wstring(buf);
    }
    if (rc == ERROR_MORE_DATA && cb > sizeof(buf)) {
        std::vector<wchar_t> large(cb / sizeof(wchar_t) + 1, 0);
        if (RegGetValueW(root, subkey, valueName, RRF_RT_REG_SZ, nullptr, large.data(), &cb) == ERROR_SUCCESS && large[0]) {
            return std::wstring(large.data());
        }
    }
    return L"";
}

static std::wstring ExtractExeFromCommand(const std::wstring& raw) {
    size_t start = raw.find_first_not_of(L" \t");
    if (start == std::wstring::npos) return L"";
    std::wstring cmd = raw.substr(start);
    if (cmd[0] == L'"') {
        size_t endQ = cmd.find(L'"', 1);
        if (endQ != std::wstring::npos) {
            std::wstring exe = cmd.substr(1, endQ - 1);
            if (IsValidExeFile(exe)) return exe;
        }
        return L"";
    }
    std::wstring low = ToLower(cmd);
    size_t pos = 0;
    while ((pos = low.find(L".exe", pos)) != std::wstring::npos) {
        size_t endExe = pos + 4;
        if (endExe == cmd.size() || cmd[endExe] == L' ' || cmd[endExe] == L'\t') {
            std::wstring exe = cmd.substr(0, endExe);
            if (IsValidExeFile(exe)) return exe;
        }
        pos = endExe;
    }
    return IsValidExeFile(cmd) ? cmd : L"";
}

static std::wstring ResolveClickShiftExeUncached() {
    wchar_t modPath[MAX_PATH];
    modPath[0] = 0;
    if (GetModuleFileNameW(g_hModule, modPath, MAX_PATH) > 0) {
        PathRemoveFileSpecW(modPath);
        std::wstring packagedExe = std::wstring(modPath) + L"\\ClickShift.exe";
        if (IsValidExeFile(packagedExe)) {
            return packagedExe;
        }
    }
    std::wstring exeReg = ReadRegString(HKEY_CURRENT_USER, L"Software\\ClickShift", L"ExePath");
    if (!exeReg.empty() && IsValidExeFile(exeReg)) {
        return exeReg;
    }
    wchar_t localApp[MAX_PATH];
    localApp[0] = 0;
    if (GetEnvironmentVariableW(L"LOCALAPPDATA", localApp, MAX_PATH) > 0) {
        std::wstring candidate = std::wstring(localApp) + L"\\ClickShift\\ClickShift.exe";
        if (IsValidExeFile(candidate)) {
            return candidate;
        }
    }
    std::wstring runCmd = ReadRegString(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Run", L"ClickShift");
    if (!runCmd.empty()) {
        std::wstring exe = ExtractExeFromCommand(runCmd);
        if (!exe.empty()) return exe;
    }
    std::wstring dlExe = FindLatestDownloadsExe();
    if (!dlExe.empty()) return dlExe;
    return L"";
}

static std::wstring ResolveIconStringUncached(const std::wstring& exe) {
    wchar_t modPath[MAX_PATH];
    modPath[0] = 0;
    if (GetModuleFileNameW(g_hModule, modPath, MAX_PATH) > 0) {
        PathRemoveFileSpecW(modPath);
        std::wstring ico = std::wstring(modPath) + L"\\clickshift.ico";
        if (GetFileAttributesW(ico.c_str()) != INVALID_FILE_ATTRIBUTES) {
            return ico;
        }
    }
    if (!exe.empty()) {
        return exe + L",0";
    }
    return L"shell32.dll,-154";
}

static void InvalidatePathCache() {
    if (g_csInit) {
        EnterCriticalSection(&g_cs);
        g_pathsTick = 0;
        g_cachedExe.clear();
        g_cachedIcon.clear();
        LeaveCriticalSection(&g_cs);
    }
}

static void EnsurePathsCached(std::wstring* outExe, std::wstring* outIcon) {
    ULONGLONG now = GetTickCount64();
    if (g_csInit) {
        EnterCriticalSection(&g_cs);
        bool fresh = (g_pathsTick != 0) && ((now - g_pathsTick) < kPathCacheTtlMs) && !g_cachedIcon.empty();
        if (fresh) {
            if (outExe) *outExe = g_cachedExe;
            if (outIcon) *outIcon = g_cachedIcon;
            LeaveCriticalSection(&g_cs);
            return;
        }
        LeaveCriticalSection(&g_cs);
    }
    std::wstring exe = ResolveClickShiftExeUncached();
    std::wstring ico = ResolveIconStringUncached(exe);
    if (g_csInit) {
        EnterCriticalSection(&g_cs);
        g_cachedExe = exe;
        g_cachedIcon = ico;
        g_pathsTick = GetTickCount64();
        LeaveCriticalSection(&g_cs);
    }
    if (outExe) *outExe = exe;
    if (outIcon) *outIcon = ico;
}

static std::wstring GetClickShiftExePath() {
    std::wstring exe;
    EnsurePathsCached(&exe, nullptr);
    return exe;
}

static std::wstring GetIconString() {
    std::wstring ico;
    EnsurePathsCached(nullptr, &ico);
    return ico;
}

static void LaunchClickShift(IShellItemArray* psiArray, const std::wstring& target, bool chooseMode) {
    if (!psiArray) return;
    std::wstring exe = GetClickShiftExePath();
    if (exe.empty()) {
        InvalidatePathCache();
        exe = GetClickShiftExePath();
    }
    if (exe.empty()) {
        MessageBoxW(nullptr, L"ClickShift executable could not be found. Try running ClickShift setup again.",
                    L"ClickShift", MB_OK | MB_ICONWARNING);
        return;
    }

    DWORD count = 0;
    if (FAILED(psiArray->GetCount(&count)) || count == 0) return;

    std::vector<std::wstring> files;
    files.reserve(count);
    for (DWORD i = 0; i < count; ++i) {
        IShellItem* psi = nullptr;
        if (SUCCEEDED(psiArray->GetItemAt(i, &psi)) && psi) {
            LPWSTR pszFile = nullptr;
            if (SUCCEEDED(psi->GetDisplayName(SIGDN_FILESYSPATH, &pszFile)) && pszFile) {
                files.emplace_back(pszFile);
                CoTaskMemFree(pszFile);
            }
            psi->Release();
        }
    }
    if (files.empty()) return;

    std::wstring baseCmd = QuoteWindowsArg(exe) + L" convert";
    if (chooseMode) {
        baseCmd += L" --choose";
    }
    std::wstring targetSuffix;
    if (!target.empty() && !chooseMode) {
        targetSuffix = L" -t " + QuoteWindowsArg(target);
    }

    std::wstring cmdLine = baseCmd;
    for (const auto& f : files) {
        cmdLine += L" " + QuoteWindowsArg(f);
    }
    cmdLine += targetSuffix;

    std::wstring flPath;
    if (cmdLine.size() > 16000) {
        flPath = WriteFilelistFallback(files);
        if (!flPath.empty()) {
            cmdLine = baseCmd + L" --filelist " + QuoteWindowsArg(flPath) + targetSuffix;
        }
    }

    STARTUPINFOW si = {};
    si.cb = sizeof(si);
    PROCESS_INFORMATION pi = {};
    std::vector<wchar_t> mutableCmd(cmdLine.begin(), cmdLine.end());
    mutableCmd.push_back(0);

    if (CreateProcessW(exe.c_str(), mutableCmd.data(), nullptr, nullptr, FALSE,
                       CREATE_NO_WINDOW, nullptr, nullptr, &si, &pi)) {
        CloseHandle(pi.hThread);
        CloseHandle(pi.hProcess);
    } else {
        if (!flPath.empty()) {
            DeleteFileW(flPath.c_str());
        }
        InvalidatePathCache();
        MessageBoxW(nullptr, L"ClickShift could not start the conversion process. Try reinstalling ClickShift.",
                    L"ClickShift", MB_OK | MB_ICONERROR);
    }
}

// ---------------------------------------------------------------------------
// Sub-command item (Auto, Convert to <EXT>, or Choose format...)
// ---------------------------------------------------------------------------
class ClickShiftSubCommand final : public IExplorerCommand, public IObjectWithSite {
    LONG m_cRef;
    IUnknown* m_pSite;
    std::wstring m_title;
    std::wstring m_target;
    bool m_isAuto;
    bool m_isChoose;
    const ExtRule* m_rule;
    bool m_resolved;

    void ResolveFallback(IShellItemArray* psiItemArray) {
        if (m_resolved) return;
        std::wstring ext = ExtractExtFromArray(psiItemArray);
        if (!ext.empty()) {
            SetCachedExt(ext);
        } else {
            ext = GetCachedExt();
        }
        if (!ext.empty()) {
            m_rule = FindRule(ext);
            if (m_isAuto && m_rule && m_rule->defaultTarget) {
                m_title = L"Auto - Convert to " + ToUpper(m_rule->defaultTarget) + L" (latest format)";
            }
            m_resolved = true;
        }
    }

public:
    ClickShiftSubCommand(const std::wstring& title, const std::wstring& target,
                         bool isAuto, bool isChoose,
                         const ExtRule* rule = nullptr, bool resolved = false)
        : m_cRef(1), m_pSite(nullptr), m_title(title), m_target(target),
          m_isAuto(isAuto), m_isChoose(isChoose), m_rule(rule), m_resolved(resolved) {
        InterlockedIncrement(&g_cDllRef);
    }

    ~ClickShiftSubCommand() {
        if (m_pSite) m_pSite->Release();
        InterlockedDecrement(&g_cDllRef);
    }

    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IExplorerCommand) {
            *ppv = static_cast<IExplorerCommand*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == IID_IObjectWithSite) {
            *ppv = static_cast<IObjectWithSite*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    IFACEMETHODIMP_(ULONG) AddRef() override {
        return InterlockedIncrement(&m_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release() override {
        LONG c = InterlockedDecrement(&m_cRef);
        if (c == 0) delete this;
        return c;
    }

    // IObjectWithSite
    IFACEMETHODIMP SetSite(IUnknown* pUnkSite) override {
        if (m_pSite) m_pSite->Release();
        m_pSite = pUnkSite;
        if (m_pSite) m_pSite->AddRef();
        return S_OK;
    }

    IFACEMETHODIMP GetSite(REFIID riid, void** ppvSite) override {
        if (!ppvSite) return E_POINTER;
        *ppvSite = nullptr;
        if (!m_pSite) return E_FAIL;
        return m_pSite->QueryInterface(riid, ppvSite);
    }

    // IExplorerCommand
    IFACEMETHODIMP GetTitle(IShellItemArray* psiItemArray, LPWSTR* ppszName) override {
        if (!ppszName) return E_POINTER;
        if (!m_resolved && m_isAuto) {
            ResolveFallback(psiItemArray);
        }
        return SHStrDupW(m_title.c_str(), ppszName);
    }

    IFACEMETHODIMP GetIcon(IShellItemArray*, LPWSTR* ppszIcon) override {
        if (!ppszIcon) return E_POINTER;
        std::wstring ico = GetIconString();
        return SHStrDupW(ico.c_str(), ppszIcon);
    }

    IFACEMETHODIMP GetToolTip(IShellItemArray*, LPWSTR* ppszInfotip) override {
        if (ppszInfotip) *ppszInfotip = nullptr;
        return E_NOTIMPL;
    }

    IFACEMETHODIMP GetCanonicalName(GUID* pguidCommandName) override {
        if (pguidCommandName) ZeroMemory(pguidCommandName, sizeof(GUID));
        return E_NOTIMPL;
    }

    IFACEMETHODIMP GetState(IShellItemArray* psiItemArray, BOOL, EXPCMDSTATE* pCmdState) override {
        if (!pCmdState) return E_POINTER;
        if (!m_resolved) {
            ResolveFallback(psiItemArray);
        }
        if (!m_isAuto && !m_isChoose && m_rule && !RuleHasTarget(m_rule, m_target)) {
            *pCmdState = ECS_HIDDEN;
            return S_OK;
        }
        *pCmdState = ECS_ENABLED;
        return S_OK;
    }

    IFACEMETHODIMP Invoke(IShellItemArray* psiItemArray, IBindCtx*) override {
        LaunchClickShift(psiItemArray, m_target, m_isChoose);
        return S_OK;
    }

    IFACEMETHODIMP GetFlags(EXPCMDFLAGS* pFlags) override {
        if (!pFlags) return E_POINTER;
        *pFlags = ECF_DEFAULT;
        return S_OK;
    }

    IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** ppEnum) override {
        if (ppEnum) *ppEnum = nullptr;
        return E_NOTIMPL;
    }
};

// ---------------------------------------------------------------------------
// Sub-command enumerator
// ---------------------------------------------------------------------------
class ClickShiftEnumCommands final : public IEnumExplorerCommand {
    LONG m_cRef;
    std::vector<IExplorerCommand*> m_cmds;
    size_t m_idx;

public:
    ClickShiftEnumCommands(const std::wstring& ext) : m_cRef(1), m_idx(0) {
        InterlockedIncrement(&g_cDllRef);
        const ExtRule* rule = FindRule(ext);
        bool resolved = !ext.empty();
        m_cmds.reserve(rule ? 18 : 24);
        std::wstring autoTitle = L"Auto - Convert to latest format";
        if (rule && rule->defaultTarget) {
            autoTitle = L"Auto - Convert to " + ToUpper(rule->defaultTarget) + L" (latest format)";
        }
        m_cmds.push_back(new ClickShiftSubCommand(autoTitle, L"", true, false, rule, resolved));

        if (rule) {
            for (int i = 0; rule->targets[i] != nullptr; ++i) {
                std::wstring t = rule->targets[i];
                std::wstring label = L"Convert to " + ToUpper(t);
                m_cmds.push_back(new ClickShiftSubCommand(label, t, false, false, rule, resolved));
            }
        } else {
            for (int i = 0; kAllTargets[i] != nullptr; ++i) {
                std::wstring t = kAllTargets[i];
                std::wstring label = L"Convert to " + ToUpper(t);
                m_cmds.push_back(new ClickShiftSubCommand(label, t, false, false, nullptr, resolved));
            }
        }
        m_cmds.push_back(new ClickShiftSubCommand(L"Choose format...", L"", false, true, rule, resolved));
    }

    ~ClickShiftEnumCommands() {
        for (auto* c : m_cmds) c->Release();
        InterlockedDecrement(&g_cDllRef);
    }

    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IEnumExplorerCommand) {
            *ppv = static_cast<IEnumExplorerCommand*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    IFACEMETHODIMP_(ULONG) AddRef() override {
        return InterlockedIncrement(&m_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release() override {
        LONG c = InterlockedDecrement(&m_cRef);
        if (c == 0) delete this;
        return c;
    }

    IFACEMETHODIMP Next(ULONG celt, IExplorerCommand** pUICommand, ULONG* pceltFetched) override {
        if (!pUICommand) return E_POINTER;
        for (ULONG i = 0; i < celt; ++i) pUICommand[i] = nullptr;
        ULONG fetched = 0;
        while (fetched < celt && m_idx < m_cmds.size()) {
            pUICommand[fetched] = m_cmds[m_idx];
            pUICommand[fetched]->AddRef();
            fetched++;
            m_idx++;
        }
        if (pceltFetched) *pceltFetched = fetched;
        return (fetched == celt) ? S_OK : S_FALSE;
    }

    IFACEMETHODIMP Skip(ULONG celt) override {
        m_idx = (std::min)(m_idx + (size_t)celt, m_cmds.size());
        return (m_idx < m_cmds.size()) ? S_OK : S_FALSE;
    }

    IFACEMETHODIMP Reset() override {
        m_idx = 0;
        return S_OK;
    }

    IFACEMETHODIMP Clone(IEnumExplorerCommand** ppEnum) override {
        if (ppEnum) *ppEnum = nullptr;
        return E_NOTIMPL;
    }
};

// ---------------------------------------------------------------------------
// Root "Convert with ClickShift >" command
// ---------------------------------------------------------------------------
class ClickShiftRootCommand final : public IExplorerCommand, public IObjectWithSite {
    LONG m_cRef;
    IUnknown* m_pSite;
    std::wstring m_ext;
    IShellItemArray* m_lastArray;
    bool m_hasRule;

    void RefreshExtIfNeeded(IShellItemArray* psiItemArray) {
        if (!psiItemArray) return;
        if (psiItemArray == m_lastArray && !m_ext.empty()) return;
        m_lastArray = psiItemArray;
        std::wstring ext = ExtractExtFromArray(psiItemArray);
        if (!ext.empty()) {
            m_ext = ext;
            m_hasRule = (FindRule(ext) != nullptr);
            SetCachedExt(ext);
        }
    }

public:
    ClickShiftRootCommand() : m_cRef(1), m_pSite(nullptr), m_lastArray(nullptr), m_hasRule(true) {
        InterlockedIncrement(&g_cDllRef);
    }

    ~ClickShiftRootCommand() {
        if (m_pSite) m_pSite->Release();
        InterlockedDecrement(&g_cDllRef);
    }

    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IExplorerCommand) {
            *ppv = static_cast<IExplorerCommand*>(this);
            AddRef();
            return S_OK;
        }
        if (riid == IID_IObjectWithSite) {
            *ppv = static_cast<IObjectWithSite*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    IFACEMETHODIMP_(ULONG) AddRef() override {
        return InterlockedIncrement(&m_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release() override {
        LONG c = InterlockedDecrement(&m_cRef);
        if (c == 0) delete this;
        return c;
    }

    // IObjectWithSite
    IFACEMETHODIMP SetSite(IUnknown* pUnkSite) override {
        if (m_pSite) m_pSite->Release();
        m_pSite = pUnkSite;
        if (m_pSite) m_pSite->AddRef();
        return S_OK;
    }

    IFACEMETHODIMP GetSite(REFIID riid, void** ppvSite) override {
        if (!ppvSite) return E_POINTER;
        *ppvSite = nullptr;
        if (!m_pSite) return E_FAIL;
        return m_pSite->QueryInterface(riid, ppvSite);
    }

    // IExplorerCommand
    IFACEMETHODIMP GetTitle(IShellItemArray* psiItemArray, LPWSTR* ppszName) override {
        if (!ppszName) return E_POINTER;
        RefreshExtIfNeeded(psiItemArray);
        return SHStrDupW(L"Convert with ClickShift", ppszName);
    }

    IFACEMETHODIMP GetIcon(IShellItemArray* psiItemArray, LPWSTR* ppszIcon) override {
        if (!ppszIcon) return E_POINTER;
        RefreshExtIfNeeded(psiItemArray);
        std::wstring ico = GetIconString();
        return SHStrDupW(ico.c_str(), ppszIcon);
    }

    IFACEMETHODIMP GetToolTip(IShellItemArray*, LPWSTR* ppszInfotip) override {
        if (ppszInfotip) *ppszInfotip = nullptr;
        return E_NOTIMPL;
    }

    IFACEMETHODIMP GetCanonicalName(GUID* pguidCommandName) override {
        if (!pguidCommandName) return E_POINTER;
        *pguidCommandName = CLSID_ClickShiftExplorerCommand;
        return S_OK;
    }

    IFACEMETHODIMP GetState(IShellItemArray* psiItemArray, BOOL, EXPCMDSTATE* pCmdState) override {
        if (!pCmdState) return E_POINTER;
        RefreshExtIfNeeded(psiItemArray);
        if (!m_ext.empty() && !m_hasRule) {
            *pCmdState = ECS_HIDDEN;
            return S_OK;
        }
        *pCmdState = ECS_ENABLED;
        return S_OK;
    }

    IFACEMETHODIMP Invoke(IShellItemArray* psiItemArray, IBindCtx*) override {
        LaunchClickShift(psiItemArray, L"", false);
        return S_OK;
    }

    IFACEMETHODIMP GetFlags(EXPCMDFLAGS* pFlags) override {
        if (!pFlags) return E_POINTER;
        *pFlags = ECF_HASSUBCOMMANDS;
        return S_OK;
    }

    IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** ppEnum) override {
        if (!ppEnum) return E_POINTER;
        std::wstring ext = m_ext.empty() ? GetCachedExt() : m_ext;
        *ppEnum = new ClickShiftEnumCommands(ext);
        return S_OK;
    }
};

// ---------------------------------------------------------------------------
// ClassFactory & DLL exports
// ---------------------------------------------------------------------------
class ClickShiftClassFactory final : public IClassFactory {
    LONG m_cRef;
public:
    ClickShiftClassFactory() : m_cRef(1) {
        InterlockedIncrement(&g_cDllRef);
    }
    ~ClickShiftClassFactory() {
        InterlockedDecrement(&g_cDllRef);
    }

    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IClassFactory) {
            *ppv = static_cast<IClassFactory*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    IFACEMETHODIMP_(ULONG) AddRef() override {
        return InterlockedIncrement(&m_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release() override {
        LONG c = InterlockedDecrement(&m_cRef);
        if (c == 0) delete this;
        return c;
    }

    IFACEMETHODIMP CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv) override {
        if (pUnkOuter) return CLASS_E_NOAGGREGATION;
        ClickShiftRootCommand* cmd = new ClickShiftRootCommand();
        HRESULT hr = cmd->QueryInterface(riid, ppv);
        cmd->Release();
        return hr;
    }

    IFACEMETHODIMP LockServer(BOOL fLock) override {
        if (fLock) InterlockedIncrement(&g_cDllRef);
        else InterlockedDecrement(&g_cDllRef);
        return S_OK;
    }
};

extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID) {
    if (dwReason == DLL_PROCESS_ATTACH) {
        g_hModule = hInstance;
        DisableThreadLibraryCalls(hInstance);
        InitializeCriticalSection(&g_cs);
        g_csInit = true;
    } else if (dwReason == DLL_PROCESS_DETACH) {
        if (g_csInit) {
            DeleteCriticalSection(&g_cs);
            g_csInit = false;
        }
    }
    return TRUE;
}

extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    if (rclsid == CLSID_ClickShiftExplorerCommand) {
        ClickShiftClassFactory* cf = new ClickShiftClassFactory();
        HRESULT hr = cf->QueryInterface(riid, ppv);
        cf->Release();
        return hr;
    }
    *ppv = nullptr;
    return CLASS_E_CLASSNOTAVAILABLE;
}

extern "C" HRESULT __stdcall DllCanUnloadNow(void) {
    return (g_cDllRef == 0) ? S_OK : S_FALSE;
}
