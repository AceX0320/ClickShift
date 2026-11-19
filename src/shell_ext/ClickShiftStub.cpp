// ClickShiftStub.cpp — Launcher for ClickShift from Windows Start Menu / Search.
// When Windows activates the packaged "ClickShift" entry, this finds the active
// ClickShift executable and starts it cleanly.

#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif
#define WIN32_LEAN_AND_MEAN

#include <windows.h>
#include <shlwapi.h>
#include <string>
#include <vector>
#include <cwctype>

static std::wstring ToLower(const std::wstring& s) {
    std::wstring out = s;
    for (wchar_t& c : out) c = (wchar_t)towlower(c);
    return out;
}

static bool IsValidExeFile(const std::wstring& path) {
    if (path.size() < 5) return false;
    DWORD attrs = GetFileAttributesW(path.c_str());
    if (attrs == INVALID_FILE_ATTRIBUTES || (attrs & FILE_ATTRIBUTE_DIRECTORY)) return false;
    if (ToLower(path.substr(path.size() - 4)) != L".exe") return false;
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

static std::wstring FindClickShiftExe() {
    // 1. Bundled Store MSIX: ClickShift.exe next to ClickShiftStub.exe.
    wchar_t selfDir[MAX_PATH];
    selfDir[0] = 0;
    if (GetModuleFileNameW(nullptr, selfDir, MAX_PATH) > 0) {
        PathRemoveFileSpecW(selfDir);
        std::wstring bundled = std::wstring(selfDir) + L"\\ClickShift.exe";
        if (IsValidExeFile(bundled)) return bundled;
    }
    // 2. Active EXE registered in HKCU\Software\ClickShift\ExePath.
    std::wstring regExe = ReadRegString(HKEY_CURRENT_USER, L"Software\\ClickShift", L"ExePath");
    if (!regExe.empty() && IsValidExeFile(regExe)) return regExe;

    // 3. Startup Run key entry.
    std::wstring runCmd = ReadRegString(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Run", L"ClickShift");
    if (!runCmd.empty()) {
        std::wstring exe = ExtractExeFromCommand(runCmd);
        if (!exe.empty()) return exe;
    }
    // 4. Default install path in %LOCALAPPDATA%\ClickShift.
    wchar_t localApp[MAX_PATH];
    localApp[0] = 0;
    if (GetEnvironmentVariableW(L"LOCALAPPDATA", localApp, MAX_PATH) > 0) {
        std::wstring fallback = std::wstring(localApp) + L"\\ClickShift\\ClickShift.exe";
        if (IsValidExeFile(fallback)) return fallback;
    }
    // 5. Latest ClickShift*.exe in %USERPROFILE%\Downloads.
    std::wstring dlExe = FindLatestDownloadsExe();
    if (!dlExe.empty()) return dlExe;
    return L"";
}

int WINAPI wWinMain(HINSTANCE, HINSTANCE, PWSTR lpCmdLine, int) {
    std::wstring exe = FindClickShiftExe();
    if (exe.empty()) {
        MessageBoxW(nullptr,
                    L"ClickShift executable could not be found. Run ClickShift Setup again.",
                    L"ClickShift", MB_OK | MB_ICONWARNING);
        return 1;
    }

    std::wstring cmdLine = QuoteWindowsArg(exe);
    if (lpCmdLine && lpCmdLine[0]) {
        cmdLine += L" ";
        cmdLine += lpCmdLine;
    }

    STARTUPINFOW si = {};
    si.cb = sizeof(si);
    PROCESS_INFORMATION pi = {};
    std::vector<wchar_t> mutableCmd(cmdLine.begin(), cmdLine.end());
    mutableCmd.push_back(0);

    if (CreateProcessW(exe.c_str(), mutableCmd.data(), nullptr, nullptr, FALSE,
                       CREATE_NEW_PROCESS_GROUP, nullptr, nullptr, &si, &pi)) {
        CloseHandle(pi.hThread);
        CloseHandle(pi.hProcess);
        return 0;
    }

    MessageBoxW(nullptr,
                L"ClickShift could not be started. Try opening ClickShift from your Downloads or install folder.",
                L"ClickShift", MB_OK | MB_ICONERROR);
    return 1;
}
