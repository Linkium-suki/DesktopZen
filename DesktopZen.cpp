/*
 * DesktopZen Ultimate - 桌面极简主义终极版 (Windows 7 兼容修复版)
 */

 // 必须放在最顶端：指定目标系统为 Windows 7
#ifndef WINVER
#define WINVER 0x0601
#endif
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0601
#endif

#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <shellapi.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <commctrl.h>
#include <filesystem>
#include <vector>
#include <string>
#include <algorithm>
#include <mutex>
#include <atomic>
#include <thread>
#include <fstream>
#include <chrono>
#include <map>

#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "User32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Advapi32.lib")

namespace fs = std::filesystem;

// ================= 常量定义 =================

#define WM_TRAYICON         (WM_USER + 101)
#define ID_TRAY_EXIT        9001
#define ID_TRAY_OPEN_CACHE  9002
#define ID_TRAY_PAUSE       9003
#define ID_TRAY_SETTINGS    9004
#define ID_TRAY_UNDO_LATEST 9005 
#define ID_TRAY_COEXIST     9006 
#define ID_TRAY_STATS       9007
#define ID_TRAY_HISTORY_BASE 10000
#define HOTKEY_ID           1001
#define MAX_CACHE_DAYS      3 

const wchar_t* CLASS_NAME = L"DesktopZen_Ultimate_Host";
const wchar_t* APP_TITLE = L"Desktop Zen Ultimate";
const wchar_t* REG_PATH = L"Software\\DesktopZen_Ultimate";
const wchar_t* REG_RUN = L"Software\\Microsoft\\Windows\\CurrentVersion\\Run";

// ================= 状态管理 =================

struct AppConfig {
    fs::path desktopPath;
    fs::path cachePath;
    bool enableHotKey = true;
    bool isFirstRun = false;
    unsigned long long totalCleaned = 0;
};

struct RuntimeState {
    HWND hWnd = nullptr;
    NOTIFYICONDATA nid = { 0 };
    std::mutex procMutex;
    UINT msgTaskbarCreated = 0;

    std::atomic<bool> isPaused{ false };
    std::atomic<bool> isInternalOp{ false };
    std::atomic<bool> isRunning{ true };
    std::atomic<bool> isWizardMode{ false };
};

AppConfig g_Config;
RuntimeState g_State;

// ================= 兼容性工具 =================

// 使用 Win32 API 替换 std::filesystem::last_write_time 以避免 CreateFile2 报错
bool SetFileLastWriteTimeNow(const fs::path& p) {
    HANDLE hFile = CreateFileW(p.c_str(), FILE_WRITE_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, NULL);
    if (hFile == INVALID_HANDLE_VALUE) return false;

    SYSTEMTIME st;
    FILETIME ft;
    GetSystemTime(&st);
    SystemTimeToFileTime(&st, &ft);
    bool success = SetFileTime(hFile, NULL, NULL, &ft);
    CloseHandle(hFile);
    return success;
}

fs::path GetKnownFolderPath(REFKNOWNFOLDERID rfid) {
    PWSTR path = nullptr;
    if (SUCCEEDED(SHGetKnownFolderPath(rfid, 0, NULL, &path))) {
        fs::path p(path);
        CoTaskMemFree(path);
        return p;
    }
    return fs::path();
}

bool IsSystemFile(const fs::path& p) {
    std::wstring name = p.filename().wstring();
    if (_wcsicmp(name.c_str(), L"desktop.ini") == 0) return true;
    if (p == g_Config.cachePath) return true;

    DWORD attr = GetFileAttributesW(p.c_str());
    if (attr == INVALID_FILE_ATTRIBUTES) return true;
    if (attr & FILE_ATTRIBUTE_SYSTEM) return true;
    if (attr & FILE_ATTRIBUTE_HIDDEN) return true;
    if (attr & FILE_ATTRIBUTE_TEMPORARY) return true;

    return false;
}

// ================= 配置管理 =================

void LoadConfig() {
    g_Config.desktopPath = GetKnownFolderPath(FOLDERID_Desktop);
    g_Config.cachePath = GetKnownFolderPath(FOLDERID_LocalAppData) / L"DesktopZenCache";

    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, REG_PATH, 0, NULL, 0, KEY_READ | KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS) {
        wchar_t buffer[MAX_PATH];
        DWORD size = sizeof(buffer);
        if (RegQueryValueExW(hKey, L"CachePath", NULL, NULL, (LPBYTE)buffer, &size) == ERROR_SUCCESS) {
            g_Config.cachePath = buffer;
        }
        else {
            g_Config.isFirstRun = true;
        }

        DWORD dwVal = 0; size = sizeof(dwVal);
        if (RegQueryValueExW(hKey, L"EnableHotKey", NULL, NULL, (LPBYTE)&dwVal, &size) == ERROR_SUCCESS) {
            g_Config.enableHotKey = (dwVal != 0);
        }

        unsigned long long stats = 0;
        size = sizeof(stats);
        if (RegQueryValueExW(hKey, L"TotalCleaned", NULL, NULL, (LPBYTE)&stats, &size) == ERROR_SUCCESS) {
            g_Config.totalCleaned = stats;
        }
        RegCloseKey(hKey);
    }
}

void SaveConfig() {
    HKEY hKey;
    if (RegCreateKeyExW(HKEY_CURRENT_USER, REG_PATH, 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS) {
        std::wstring ws = g_Config.cachePath.wstring();
        RegSetValueExW(hKey, L"CachePath", 0, REG_SZ, (const BYTE*)ws.c_str(), (DWORD)((ws.length() + 1) * sizeof(wchar_t)));
        DWORD dwVal = g_Config.enableHotKey ? 1 : 0;
        RegSetValueExW(hKey, L"EnableHotKey", 0, REG_DWORD, (const BYTE*)&dwVal, sizeof(dwVal));
        RegSetValueExW(hKey, L"TotalCleaned", 0, REG_QWORD, (const BYTE*)&g_Config.totalCleaned, sizeof(g_Config.totalCleaned));
        RegCloseKey(hKey);
    }
}

void SetAutoStart(bool enable) {
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, REG_RUN, 0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS) {
        if (enable) {
            wchar_t exePath[MAX_PATH];
            GetModuleFileNameW(NULL, exePath, MAX_PATH);
            RegSetValueExW(hKey, L"DesktopZen", 0, REG_SZ, (BYTE*)exePath, (DWORD)((wcslen(exePath) + 1) * sizeof(wchar_t)));
        }
        else {
            RegDeleteValueW(hKey, L"DesktopZen");
        }
        RegCloseKey(hKey);
    }
}

// ================= 托盘 UI =================

void InitTray(HWND hwnd) {
    g_State.nid.cbSize = sizeof(NOTIFYICONDATA);
    g_State.nid.hWnd = hwnd;
    g_State.nid.uID = 1;
    g_State.nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_State.nid.uCallbackMessage = WM_TRAYICON;
    g_State.nid.hIcon = LoadIcon(NULL, IDI_SHIELD);
    wcscpy_s(g_State.nid.szTip, APP_TITLE);
    Shell_NotifyIcon(NIM_ADD, &g_State.nid);
}

void UpdateTrayIcon() {
    if (g_State.isPaused) {
        g_State.nid.hIcon = LoadIcon(NULL, IDI_WARNING);
        wcscpy_s(g_State.nid.szTip, L"Desktop Zen (已暂停)");
    }
    else {
        g_State.nid.hIcon = LoadIcon(NULL, IDI_SHIELD);
        wcscpy_s(g_State.nid.szTip, APP_TITLE);
    }
    Shell_NotifyIcon(NIM_MODIFY, &g_State.nid);
}

// ================= 业务逻辑 =================

void MoveToCache(const fs::path& source) {
    if (!fs::exists(source)) return;
    auto timestamp = std::chrono::duration_cast<std::chrono::milliseconds>(
        std::chrono::system_clock::now().time_since_epoch()).count();

    std::wstring safeName = std::to_wstring(timestamp) + L"_" + source.filename().wstring();
    fs::path dest = g_Config.cachePath / safeName;

    std::error_code ec;
    for (int i = 0; i < 5; ++i) {
        fs::rename(source, dest, ec);
        if (!ec) {
            g_Config.totalCleaned++;
            if (g_Config.totalCleaned % 5 == 0) SaveConfig();
            break;
        }
        std::this_thread::sleep_for(std::chrono::milliseconds(50));
    }
}

void EnforcePolicy(const std::wstring& vipFile = L"") {
    if (g_State.isInternalOp) return;
    if (g_State.isPaused && !g_State.isWizardMode) return;

    std::lock_guard<std::mutex> lock(g_State.procMutex);

    if (!fs::exists(g_Config.cachePath)) {
        std::error_code ec;
        fs::create_directories(g_Config.cachePath, ec);
    }

    std::vector<fs::directory_entry> files;
    try {
        for (const auto& entry : fs::directory_iterator(g_Config.desktopPath)) {
            if (IsSystemFile(entry.path())) continue;
            files.push_back(entry);
        }
    }
    catch (...) { return; }

    if (files.size() <= 1) return;

    fs::path keeper;
    if (!vipFile.empty()) {
        fs::path vipPath = g_Config.desktopPath / vipFile;
        for (auto& f : files) {
            if (f.path() == vipPath) {
                keeper = vipPath;
                break;
            }
        }
    }

    if (keeper.empty()) {
        // 兼容性优化：Windows 7 上直接使用系统时间比较，避开 filesystem 的 time_point 转换坑
        std::sort(files.begin(), files.end(), [](const fs::directory_entry& a, const fs::directory_entry& b) {
            WIN32_FILE_ATTRIBUTE_DATA dataA, dataB;
            GetFileAttributesExW(a.path().c_str(), GetFileExInfoStandard, &dataA);
            GetFileAttributesExW(b.path().c_str(), GetFileExInfoStandard, &dataB);
            return CompareFileTime(&dataA.ftLastWriteTime, &dataB.ftLastWriteTime) > 0;
            });
        keeper = files[0].path();
    }

    for (const auto& entry : files) {
        if (entry.path() != keeper) {
            MoveToCache(entry.path());
        }
    }
}

void PerformRestore(const fs::path& cachedFile, bool coexistMode) {
    if (!fs::exists(cachedFile)) return;
    g_State.isInternalOp = true;

    try {
        std::wstring rawName = cachedFile.filename().wstring();
        size_t u = rawName.find(L'_');
        std::wstring realName = (u != std::wstring::npos) ? rawName.substr(u + 1) : rawName;
        fs::path dest = g_Config.desktopPath / realName;

        if (fs::exists(dest)) {
            if (coexistMode) {
                std::wstring stem = dest.stem().wstring();
                std::wstring ext = dest.extension().wstring();
                dest = dest.parent_path() / (stem + L"_Restored" + ext);
            }
            else {
                MoveToCache(dest);
            }
        }

        std::error_code ec;
        fs::rename(cachedFile, dest, ec);

        // 修复：不使用 fs::last_write_time，使用 Win32 API 避免 CreateFile2 报错
        if (!ec) SetFileLastWriteTimeNow(dest);

        if (coexistMode) {
            g_State.isPaused = true;
            UpdateTrayIcon();
        }
    }
    catch (...) {}

    std::this_thread::sleep_for(std::chrono::milliseconds(200));
    g_State.isInternalOp = false;
    if (!coexistMode && !g_State.isPaused) EnforcePolicy();
}

void MaintenanceThread() {
    while (g_State.isRunning) {
        try {
            if (fs::exists(g_Config.cachePath)) {
                auto now = std::chrono::system_clock::now();
                for (const auto& entry : fs::directory_iterator(g_Config.cachePath)) {
                    // 修复：不再通过 entry.path() 获取 filesystem 时间戳，直接解析文件名
                    std::wstring fname = entry.path().filename().wstring();
                    size_t p = fname.find(L'_');
                    if (p != std::wstring::npos) {
                        try {
                            long long ts = std::stoll(fname.substr(0, p));
                            auto fileTime = std::chrono::time_point<std::chrono::system_clock>(std::chrono::milliseconds(ts));
                            auto hours = std::chrono::duration_cast<std::chrono::hours>(now - fileTime).count();
                            if (hours > MAX_CACHE_DAYS * 24) {
                                std::error_code ec;
                                fs::remove_all(entry.path(), ec);
                            }
                        }
                        catch (...) {}
                    }
                }
            }
        }
        catch (...) {}
        for (int i = 0; i < 6 * 60 && g_State.isRunning; ++i)
            std::this_thread::sleep_for(std::chrono::minutes(1));
    }
}

// 获取最近缓存列表
struct CacheItem { fs::path path; std::wstring name; long long time; };
std::vector<CacheItem> GetRecentCache(int limit = 10) {
    std::vector<CacheItem> items;
    try {
        if (fs::exists(g_Config.cachePath)) {
            for (const auto& entry : fs::directory_iterator(g_Config.cachePath)) {
                if (!entry.is_regular_file()) continue;
                std::wstring fname = entry.path().filename().wstring();
                size_t p = fname.find(L'_');
                if (p != std::wstring::npos) {
                    try {
                        long long ts = std::stoll(fname.substr(0, p));
                        items.push_back({ entry.path(), fname.substr(p + 1), ts });
                    }
                    catch (...) {}
                }
            }
        }
    }
    catch (...) {}
    std::sort(items.begin(), items.end(), [](auto& a, auto& b) { return a.time > b.time; });
    if (items.size() > (size_t)limit) items.resize(limit);
    return items;
}

// ================= 监控线程 =================

void MonitorThread() {
    // Windows 7 支持 CreateFileW 的备份语义标志
    HANDLE hDir = CreateFileW(g_Config.desktopPath.c_str(), FILE_LIST_DIRECTORY,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, NULL);
    if (hDir == INVALID_HANDLE_VALUE) return;

    std::vector<BYTE> buffer(64 * 1024);
    DWORD bytesReturned;

    while (g_State.isRunning) {
        if (ReadDirectoryChangesW(hDir, buffer.data(), (DWORD)buffer.size(), FALSE,
            FILE_NOTIFY_CHANGE_FILE_NAME | FILE_NOTIFY_CHANGE_DIR_NAME | FILE_NOTIFY_CHANGE_LAST_WRITE,
            &bytesReturned, NULL, NULL))
        {
            std::this_thread::sleep_for(std::chrono::milliseconds(50));
            if (g_State.isInternalOp) continue;

            std::wstring vipFile = L"";
            FILE_NOTIFY_INFORMATION* pNotify = (FILE_NOTIFY_INFORMATION*)buffer.data();
            while (true) {
                std::wstring fileName(pNotify->FileName, pNotify->FileNameLength / sizeof(wchar_t));
                if (pNotify->Action == FILE_ACTION_ADDED ||
                    pNotify->Action == FILE_ACTION_RENAMED_NEW_NAME ||
                    pNotify->Action == FILE_ACTION_MODIFIED) {
                    vipFile = fileName;
                }
                if (pNotify->NextEntryOffset == 0) break;
                pNotify = (FILE_NOTIFY_INFORMATION*)((BYTE*)pNotify + pNotify->NextEntryOffset);
            }
            EnforcePolicy(vipFile);
        }
        else {
            std::this_thread::sleep_for(std::chrono::seconds(2));
        }
    }
    CloseHandle(hDir);
}

// ================= 向导与主消息 =================

void RunWizard() {
    g_State.isWizardMode = true;
    g_State.isPaused = true;

    if (MessageBox(NULL, L"【步骤 1/3】检测到桌面文件较多，是否执行一键收纳？", L"初始化", MB_YESNO) == IDYES) {
        fs::path backup = g_Config.cachePath / L"_InitialBackup";
        fs::create_directories(backup);
        try {
            for (const auto& entry : fs::directory_iterator(g_Config.desktopPath)) {
                if (IsSystemFile(entry.path())) continue;
                fs::rename(entry.path(), backup / entry.path().filename());
            }
        }
        catch (...) {}
    }

    if (MessageBox(NULL, L"【步骤 2/3】是否允许开机自启？", L"设置", MB_YESNO) == IDYES) SetAutoStart(true);
    g_Config.enableHotKey = (MessageBox(NULL, L"是否启用 Alt+Z 快速撤销？", L"设置", MB_YESNO) == IDYES);

    g_State.isPaused = false;
    MessageBox(NULL, L"【步骤 3/3】配置完成！\n请尝试复制一个文件到桌面体验。", L"完成", MB_OK);

    SaveConfig();
    g_State.isWizardMode = false;
    EnforcePolicy();
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (msg == g_State.msgTaskbarCreated) { InitTray(hwnd); return 0; }

    switch (msg) {
    case WM_CREATE:
        g_State.msgTaskbarCreated = RegisterWindowMessageW(L"TaskbarCreated");
        InitTray(hwnd);
        if (g_Config.enableHotKey) RegisterHotKey(hwnd, HOTKEY_ID, MOD_ALT, 'Z');
        break;
    case WM_HOTKEY:
        if (wParam == HOTKEY_ID) {
            auto items = GetRecentCache(1);
            if (!items.empty()) { PerformRestore(items[0].path, false); MessageBeep(MB_OK); }
        }
        break;
    case WM_TRAYICON:
        if (lParam == WM_RBUTTONUP || lParam == WM_LBUTTONUP) {
            POINT pt; GetCursorPos(&pt);
            SetForegroundWindow(hwnd);
            HMENU hMenu = CreatePopupMenu();
            auto recent = GetRecentCache(10);

            std::wstring statsText = L"累计净化: " + std::to_wstring(g_Config.totalCleaned) + L" 个文件";
            AppendMenu(hMenu, MF_STRING | MF_DISABLED, ID_TRAY_STATS, statsText.c_str());
            AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);

            if (!recent.empty()) {
                std::wstring fName = recent[0].name;
                if (fName.length() > 20) fName = fName.substr(0, 20) + L"...";
                AppendMenu(hMenu, MF_STRING, ID_TRAY_UNDO_LATEST, (L"撤销: " + fName).c_str());
            }

            AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
            if (g_State.isPaused) AppendMenu(hMenu, MF_STRING | MF_CHECKED, ID_TRAY_PAUSE, L"恢复监控");
            else AppendMenu(hMenu, MF_STRING, ID_TRAY_PAUSE, L"暂停监控");

            AppendMenu(hMenu, MF_STRING, ID_TRAY_OPEN_CACHE, L"打开缓存文件夹");
            AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
            AppendMenu(hMenu, MF_STRING, ID_TRAY_EXIT, L"退出");

            int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_NONOTIFY, pt.x, pt.y, 0, hwnd, NULL);
            DestroyMenu(hMenu);

            if (cmd == ID_TRAY_EXIT) DestroyWindow(hwnd);
            else if (cmd == ID_TRAY_OPEN_CACHE) ShellExecute(NULL, L"open", g_Config.cachePath.c_str(), NULL, NULL, SW_SHOW);
            else if (cmd == ID_TRAY_UNDO_LATEST && !recent.empty()) PerformRestore(recent[0].path, false);
            else if (cmd == ID_TRAY_PAUSE) { g_State.isPaused = !g_State.isPaused; UpdateTrayIcon(); if (!g_State.isPaused) EnforcePolicy(); }
        }
        break;
    case WM_DESTROY:
        g_State.isRunning = false;
        SaveConfig();
        UnregisterHotKey(hwnd, HOTKEY_ID);
        Shell_NotifyIcon(NIM_DELETE, &g_State.nid);
        PostQuitMessage(0);
        break;
    default: return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int) {
    HANDLE hMutex = CreateMutex(NULL, TRUE, L"Global\\DesktopZenUltimate_Win7");
    if (GetLastError() == ERROR_ALREADY_EXISTS) return 0;

    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    LoadConfig();

    WNDCLASS wc = { 0 };
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    RegisterClass(&wc);

    HWND hwnd = CreateWindowEx(0, CLASS_NAME, APP_TITLE, 0, 0, 0, 0, 0, HWND_MESSAGE, NULL, hInstance, NULL);

    std::thread(MonitorThread).detach();
    std::thread(MaintenanceThread).detach();

    if (g_Config.isFirstRun) {
        std::thread(RunWizard).detach();
    }
    else {
        std::thread([] { std::this_thread::sleep_for(std::chrono::seconds(1)); EnforcePolicy(); }).detach();
    }

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    CoUninitialize();
    ReleaseMutex(hMutex);
    return 0;
}