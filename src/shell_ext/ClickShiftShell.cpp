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
    {L"jxl",  L"png",  {L"png", L"jpg", L"webp", L"pdf", L"docx", L"pptx", L"avif", L"tiff", L"bmp", L"ico", nullptr}},
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

static std::wstring GetClickShiftExePath() {
    wchar_t buf[MAX_PATH * 2] = {0};
    DWORD cb = sizeof(buf);
    if (RegGetValueW(HKEY_CURRENT_USER, L"Software\\ClickShift", L"ExePath",
                     RRF_RT_REG_SZ, nullptr, buf, &cb) == ERROR_SUCCESS && buf[0]) {
        if (GetFileAttributesW(buf) != INVALID_FILE_ATTRIBUTES) {
            return std::wstring(buf);
        }
    }
    wchar_t localApp[MAX_PATH] = {0};
    if (GetEnvironmentVariableW(L"LOCALAPPDATA", localApp, MAX_PATH) > 0) {
        std::wstring candidate = std::wstring(localApp) + L"\\ClickShift\\ClickShift.exe";
        if (GetFileAttributesW(candidate.c_str()) != INVALID_FILE_ATTRIBUTES) {
            return candidate;
        }
    }
    cb = sizeof(buf);
    if (RegGetValueW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Run",
                     L"ClickShift", RRF_RT_REG_SZ, nullptr, buf, &cb) == ERROR_SUCCESS && buf[0]) {
        std::wstring cmd(buf);
        if (!cmd.empty() && cmd[0] == L'"') {
            size_t endQ = cmd.find(L'"', 1);
            if (endQ != std::wstring::npos) {
                std::wstring exe = cmd.substr(1, endQ - 1);
                if (GetFileAttributesW(exe.c_str()) != INVALID_FILE_ATTRIBUTES) {
                    return exe;
                }
            }
        }
    }
    return L"";
}

static std::wstring GetIconString() {
    wchar_t modPath[MAX_PATH] = {0};
    if (GetModuleFileNameW(g_hModule, modPath, MAX_PATH) > 0) {
        PathRemoveFileSpecW(modPath);
        std::wstring ico = std::wstring(modPath) + L"\\clickshift.ico";
        if (GetFileAttributesW(ico.c_str()) != INVALID_FILE_ATTRIBUTES) {
            return ico;
        }
    }
    std::wstring exe = GetClickShiftExePath();
    if (!exe.empty()) {
        return exe + L",0";
    }
    return L"shell32.dll,-154";
}

static void LaunchClickShift(IShellItemArray* psiArray, const std::wstring& target, bool chooseMode) {
    if (!psiArray) return;
    std::wstring exe = GetClickShiftExePath();
    if (exe.empty()) return;

    DWORD count = 0;
    if (FAILED(psiArray->GetCount(&count)) || count == 0) return;

    std::wstring cmdLine = L"\"" + exe + L"\" convert";
    if (chooseMode) {
        cmdLine += L" --choose";
    }
    DWORD added = 0;
    for (DWORD i = 0; i < count; ++i) {
        IShellItem* psi = nullptr;
        if (SUCCEEDED(psiArray->GetItemAt(i, &psi)) && psi) {
            LPWSTR pszFile = nullptr;
            if (SUCCEEDED(psi->GetDisplayName(SIGDN_FILESYSPATH, &pszFile)) && pszFile) {
                cmdLine += L" \"";
                cmdLine += pszFile;
                cmdLine += L"\"";
                added++;
                CoTaskMemFree(pszFile);
            }
            psi->Release();
        }
    }
    if (added == 0) return;
    if (!target.empty() && !chooseMode) {
        cmdLine += L" -t " + target;
    }

    STARTUPINFOW si = {sizeof(si)};
    PROCESS_INFORMATION pi = {0};
    std::vector<wchar_t> mutableCmd(cmdLine.begin(), cmdLine.end());
    mutableCmd.push_back(0);

    if (CreateProcessW(exe.c_str(), mutableCmd.data(), nullptr, nullptr, FALSE,
                       CREATE_NO_WINDOW, nullptr, nullptr, &si, &pi)) {
        CloseHandle(pi.hThread);
        CloseHandle(pi.hProcess);
    }
}

// ---------------------------------------------------------------------------
// Sub-command item (Auto, Convert to <EXT>, or Choose format...)
// ---------------------------------------------------------------------------
class ClickShiftSubCommand : public IExplorerCommand, public IObjectWithSite {
    LONG m_cRef;
    IUnknown* m_pSite;
    std::wstring m_title;
    std::wstring m_target;
    bool m_isAuto;
    bool m_isChoose;

public:
    ClickShiftSubCommand(const std::wstring& title, const std::wstring& target, bool isAuto, bool isChoose)
        : m_cRef(1), m_pSite(nullptr), m_title(title), m_target(target), m_isAuto(isAuto), m_isChoose(isChoose) {
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
        if (m_isAuto) {
            std::wstring ext = ExtractExtFromArray(psiItemArray);
            if (!ext.empty()) SetCachedExt(ext);
            else ext = GetCachedExt();
            const ExtRule* rule = FindRule(ext);
            if (rule && rule->defaultTarget) {
                std::wstring dyn = L"Auto - Convert to " + ToUpper(rule->defaultTarget) + L" (latest format)";
                return SHStrDupW(dyn.c_str(), ppszName);
            }
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
        std::wstring ext = ExtractExtFromArray(psiItemArray);
        if (!ext.empty()) SetCachedExt(ext);
        else ext = GetCachedExt();
        if (!m_isAuto && !m_isChoose && !ext.empty()) {
            const ExtRule* rule = FindRule(ext);
            if (rule && !RuleHasTarget(rule, m_target)) {
                *pCmdState = ECS_HIDDEN;
                return S_OK;
            }
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
class ClickShiftEnumCommands : public IEnumExplorerCommand {
    LONG m_cRef;
    std::vector<IExplorerCommand*> m_cmds;
    size_t m_idx;

public:
    ClickShiftEnumCommands(const std::wstring& ext) : m_cRef(1), m_idx(0) {
        InterlockedIncrement(&g_cDllRef);
        const ExtRule* rule = FindRule(ext);
        std::wstring autoTitle = L"Auto - Convert to latest format";
        if (rule && rule->defaultTarget) {
            autoTitle = L"Auto - Convert to " + ToUpper(rule->defaultTarget) + L" (latest format)";
        }
        m_cmds.push_back(new ClickShiftSubCommand(autoTitle, L"", true, false));

        if (rule) {
            for (int i = 0; rule->targets[i] != nullptr; ++i) {
                std::wstring t = rule->targets[i];
                std::wstring label = L"Convert to " + ToUpper(t);
                m_cmds.push_back(new ClickShiftSubCommand(label, t, false, false));
            }
        } else {
            for (int i = 0; kAllTargets[i] != nullptr; ++i) {
                std::wstring t = kAllTargets[i];
                std::wstring label = L"Convert to " + ToUpper(t);
                m_cmds.push_back(new ClickShiftSubCommand(label, t, false, false));
            }
        }
        m_cmds.push_back(new ClickShiftSubCommand(L"Choose format...", L"", false, true));
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
class ClickShiftRootCommand : public IExplorerCommand, public IObjectWithSite {
    LONG m_cRef;
    IUnknown* m_pSite;
    std::wstring m_ext;

public:
    ClickShiftRootCommand() : m_cRef(1), m_pSite(nullptr) {
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
        std::wstring ext = ExtractExtFromArray(psiItemArray);
        if (!ext.empty()) {
            m_ext = ext;
            SetCachedExt(ext);
        }
        return SHStrDupW(L"Convert with ClickShift", ppszName);
    }

    IFACEMETHODIMP GetIcon(IShellItemArray* psiItemArray, LPWSTR* ppszIcon) override {
        if (!ppszIcon) return E_POINTER;
        std::wstring ext = ExtractExtFromArray(psiItemArray);
        if (!ext.empty()) {
            m_ext = ext;
            SetCachedExt(ext);
        }
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
        std::wstring ext = ExtractExtFromArray(psiItemArray);
        if (!ext.empty()) {
            m_ext = ext;
            SetCachedExt(ext);
            if (!FindRule(ext)) {
                *pCmdState = ECS_HIDDEN;
                return S_OK;
            }
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
class ClickShiftClassFactory : public IClassFactory {
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
