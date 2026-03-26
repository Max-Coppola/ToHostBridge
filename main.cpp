// Dear ImGui: standalone example application for Windows API + DirectX 11
#include "imgui.h"
#include "imgui_impl_win32.h"
#include "imgui_impl_dx11.h"
#include <d3d11.h>
#include <tchar.h>

#include <iostream>
#include <string>
#include <vector>
#include <memory>
#include <mutex>
#include <deque>
#include <sstream>
#include <iomanip>
#include <chrono>

#include <winrt/Windows.Foundation.h>
#include <winmidi/init/Microsoft.Windows.Devices.Midi2.Initialization.hpp>

#include "SerialPort.h"
#include "VirtualMidiPort.h"

void SendToSerial(int portIdx, const std::vector<uint8_t>& data);

#include <shellapi.h>

#define WM_TRAYICON (WM_APP + 1)
#define IDI_ICON1 101
NOTIFYICONDATAW g_nid = {};

// Global config
bool g_reduceToTray  = false;
bool g_transmitSync  = true;

// Data
static ID3D11Device*            g_pd3dDevice = nullptr;
static ID3D11DeviceContext*     g_pd3dDeviceContext = nullptr;
static IDXGISwapChain*          g_pSwapChain = nullptr;
static bool                     g_SwapChainOccluded = false;
static UINT                     g_ResizeWidth = 0, g_ResizeHeight = 0;
static ID3D11RenderTargetView*  g_mainRenderTargetView = nullptr;

// Forward declarations of helper functions
bool CreateDeviceD3D(HWND hWnd);
void CleanupDeviceD3D();
void CreateRenderTarget();
void CleanupRenderTarget();
LRESULT WINAPI WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);

// Application State
std::unique_ptr<SerialPort> g_serialPort;
std::vector<std::unique_ptr<VirtualMidiPort>> g_virtualPortsOut;
std::unique_ptr<VirtualMidiPort> g_virtualPortIn;

// Serial write mutex + port running-status
// g_lastSentPort == 0 means "no port selected yet" (force prefix on first write)
std::mutex g_serialWriteMutex;
std::atomic<int> g_lastSentPort{0};

// Per-port enable flags (toggled via UI, read from callbacks)
std::atomic<bool> g_portEnabled[4];
// portInEnabled: controls COM -> virtual In (serial receive callback)
std::atomic<bool> g_portInEnabled{true};
// Bandwidth monitoring
std::atomic<uint64_t> g_bytesSent{0};
std::atomic<uint64_t> g_bytesReceived{0};
float g_chargePercent = 0.0f;

enum class LogSourceType {
    APP_INFO,
    COM_IN,
    COM_IN_SYNC,
    APP_OUT_1,
    APP_OUT_2,
    APP_OUT_3,
    APP_OUT_4
};
static const int LOG_SOURCE_COUNT = 7;

struct LogEntry {
    LogSourceType source;
    std::string timestamp; // "[HH:MM:SS.mmm] "
    std::string text;      // body without timestamp
    std::string cmdName;
    ImVec4 color;
    bool isSync;
    uint64_t seqNum;
};

std::mutex g_logMutex;
// Per-source ring buffers — each capped at g_maxLogLines independently
std::deque<LogEntry> g_logBySource[LOG_SOURCE_COUNT];
int g_maxLogLines = 500;
uint64_t g_logSeq = 0;    // global sequence counter
size_t g_logVersion = 0;  // incremented on every push

void AddLog(LogSourceType source, const std::string& msg, ImVec4 color, bool isSync = false, const std::string& cmdName = "") {
    auto now = std::chrono::system_clock::now();
    auto in_time_t = std::chrono::system_clock::to_time_t(now);
    auto ms = std::chrono::duration_cast<std::chrono::milliseconds>(now.time_since_epoch()) % 1000;
    
    std::stringstream time_ss;
    #pragma warning(disable: 4996)
    time_ss << std::put_time(std::localtime(&in_time_t), "%H:%M:%S") << "." << std::setfill('0') << std::setw(3) << ms.count();
    #pragma warning(default: 4996)
    
    std::string timeStr = "[" + time_ss.str() + "] ";

    std::lock_guard<std::mutex> lock(g_logMutex);
    int idx = (int)source;
    g_logBySource[idx].push_back({source, timeStr, msg, cmdName, color, isSync, ++g_logSeq});
    ++g_logVersion;
    if ((int)g_logBySource[idx].size() > g_maxLogLines) {
        g_logBySource[idx].pop_front();
    }
}

std::string BytesToHex(const std::vector<uint8_t>& data) {
    std::stringstream ss;
    for (auto b : data) {
        ss << std::hex << std::uppercase << std::setw(2) << std::setfill('0') << (int)b << " ";
    }
    return ss.str();
}

std::string GetMidiCommandName(const std::vector<uint8_t>& data) {
    if (data.empty()) return "";
    uint8_t st = data[0];
    if (st >= 0xF8) {
        switch(st) {
            case 0xF8: return "(Clock)";
            case 0xFA: return "(Start)";
            case 0xFB: return "(Continue)";
            case 0xFC: return "(Stop)";
            case 0xFE: return "(Active Sensing)";
            case 0xFF: return "(Reset)";
            default: return "(Realtime)";
        }
    }
    uint8_t cmd = st & 0xF0;
    switch(cmd) {
        case 0x80: return "(Note Off)";
        case 0x90: return (data.size() > 2 && data[2] == 0) ? "(Note Off)" : "(Note On)";
        case 0xA0: return "(Poly Aftertouch)";
        case 0xB0: {
            if (data.size() < 2) return "(CC)";
            uint8_t cc = data[1];
            switch(cc) {
                case 0:  return "(Bank Select MSB)";
                case 32: return "(Bank Select LSB)";
                case 1:  return "(Modulation)";
                case 7:  return "(Volume)";
                case 10: return "(Pan)";
                case 11: return "(Expression)";
                case 64: return "(Sustain)";
                case 120: return "(All Sound Off)";
                case 121: return "(Reset All CC)";
                case 123: return "(All Notes Off)";
                default: return "(CC " + std::to_string(cc) + ")";
            }
        }
        case 0xC0: return "(Program Change)";
        case 0xD0: return "(Channel Aftertouch)";
        case 0xE0: return "(Pitch Bend)";
        case 0xF0: {
            if (st == 0xF0) {
                // Check common SysEx
                if (data.size() >= 6 && data[1] == 0x7E && data[3] == 0x09 && data[4] == 0x01) return "(GM System On)";
                if (data.size() >= 9 && data[1] == 0x43 && data[3] == 0x4C && data[6] == 0x7E) return "(XG System On)";
                if (data.size() >= 9 && data[1] == 0x43 && data[3] == 0x4C && data[6] == 0x7D) return "(XG All Parameter Reset)";
                return "(SysEx)";
            }
            return "";
        }
    }
    return "";
}

std::vector<std::string> GetAvailableComPorts() {
    std::vector<std::string> ports;
    char path[5000];
    for (int i = 1; i <= 255; i++) {
        std::string comName = "COM" + std::to_string(i);
        DWORD res = QueryDosDeviceA(comName.c_str(), path, sizeof(path));
        if (res != 0) {
            ports.push_back(comName);
        }
    }
    return ports;
}

// Configuration persistence
std::string GetIniPathDir() {
    char path[MAX_PATH];
    GetModuleFileNameA(NULL, path, MAX_PATH);
    std::string s(path);
    return s.substr(0, s.find_last_of("\\/"));
}

std::string GetIniPath() {
    return GetIniPathDir() + "\\settings.ini";
}

// Encode imgui ini string (replace newlines with '|') for single-line INI storage
std::string EncodeIniData(const std::string& s) {
    std::string out; out.reserve(s.size());
    for (char c : s) out += (c == '\n') ? '|' : c;
    return out;
}
std::string DecodeIniData(const std::string& s) {
    std::string out; out.reserve(s.size());
    for (char c : s) out += (c == '|') ? '\n' : c;
    return out;
}

void SaveImguiIniToSettings() {
    size_t sz = 0;
    const char* raw = ImGui::SaveIniSettingsToMemory(&sz);
    if (!raw || sz == 0) return;
    std::string encoded = EncodeIniData(std::string(raw, sz));
    WritePrivateProfileStringA("ImGui", "Layout", encoded.c_str(), GetIniPath().c_str());
}

void LoadImguiIniFromSettings() {
    char buf[32768] = "";
    if (GetPrivateProfileStringA("ImGui", "Layout", "", buf, sizeof(buf), GetIniPath().c_str()) > 0) {
        std::string decoded = DecodeIniData(buf);
        ImGui::LoadIniSettingsFromMemory(decoded.c_str(), decoded.size());
    }
}

void SaveSettings(const std::string& comName, const std::string& baseName, bool autoStart, bool startWin, bool autoReconn, bool startTray, ImVec4 cSys, ImVec4 cIn, ImVec4* cOut, bool stayOnTop, bool startMinimized, bool lightUI) {
    std::string ini = GetIniPath();
    WritePrivateProfileStringA("General", "ComPortName", comName.c_str(), ini.c_str());
    WritePrivateProfileStringA("General", "BaseName", baseName.c_str(), ini.c_str());
    WritePrivateProfileStringA("General", "AutoStartVirtualMidi", autoStart ? "1" : "0", ini.c_str());
    WritePrivateProfileStringA("General", "StartWithWindows", startWin ? "1" : "0", ini.c_str());
    WritePrivateProfileStringA("General", "ReduceToTray", g_reduceToTray ? "1" : "0", ini.c_str());
    WritePrivateProfileStringA("General", "AutoReconnect", autoReconn ? "1" : "0", ini.c_str());
    WritePrivateProfileStringA("General", "StartToTray", startTray ? "1" : "0", ini.c_str());
    WritePrivateProfileStringA("General", "MaxLogLines", std::to_string(g_maxLogLines).c_str(), ini.c_str());
    WritePrivateProfileStringA("General", "TransmitSync", g_transmitSync ? "1" : "0", ini.c_str());
    WritePrivateProfileStringA("General", "StayOnTop", stayOnTop ? "1" : "0", ini.c_str());
    WritePrivateProfileStringA("General", "StartMinimized", startMinimized ? "1" : "0", ini.c_str());
    WritePrivateProfileStringA("General", "LightUI", lightUI ? "1" : "0", ini.c_str());
    WritePrivateProfileStringA("Ports", "PortInEnabled", g_portInEnabled.load() ? "1" : "0", ini.c_str());
    for (int p = 0; p < 4; ++p)
        WritePrivateProfileStringA("Ports", ("Port" + std::to_string(p+1) + "Enabled").c_str(), g_portEnabled[p].load() ? "1" : "0", ini.c_str());
    // Save imgui layout
    SaveImguiIniToSettings();

    auto saveColor = [&](const char* key, ImVec4 c) {
        char buf[64];
        snprintf(buf, sizeof(buf), "%.2f,%.2f,%.2f,%.2f", c.x, c.y, c.z, c.w);
        WritePrivateProfileStringA("Colors", key, buf, ini.c_str());
    };
    saveColor("ColSys", cSys);
    saveColor("ColIn", cIn);
    for (int i=0; i<4; i++) saveColor(("ColOut" + std::to_string(i+1)).c_str(), cOut[i]);
}

void LoadSettings(std::string& comName, char* baseName, bool& autoStart, bool& startWin, bool& autoReconn, bool& startTray, ImVec4& cSys, ImVec4& cIn, ImVec4* cOut, bool& stayOnTop, bool& startMinimized, bool& lightUI) {
    std::string ini = GetIniPath();
    char buf[128] = "";
    GetPrivateProfileStringA("General", "ComPortName", "", buf, 128, ini.c_str());
    comName = buf;
    GetPrivateProfileStringA("General", "BaseName", "CBX", baseName, 128, ini.c_str());
    autoStart = GetPrivateProfileIntA("General", "AutoStartVirtualMidi", 0, ini.c_str());
    startWin = GetPrivateProfileIntA("General", "StartWithWindows", 0, ini.c_str());
    g_reduceToTray = GetPrivateProfileIntA("General", "ReduceToTray", 0, ini.c_str());
    autoReconn = GetPrivateProfileIntA("General", "AutoReconnect", 0, ini.c_str());
    startTray = GetPrivateProfileIntA("General", "StartToTray", 0, ini.c_str());
    g_maxLogLines = GetPrivateProfileIntA("General", "MaxLogLines", 500, ini.c_str());
    if (g_maxLogLines < 100) g_maxLogLines = 100;
    if (g_maxLogLines > 1000) g_maxLogLines = 1000;
    g_transmitSync = GetPrivateProfileIntA("General", "TransmitSync", 1, ini.c_str());
    stayOnTop     = GetPrivateProfileIntA("General", "StayOnTop", 0, ini.c_str()) != 0;
    startMinimized = GetPrivateProfileIntA("General", "StartMinimized", 0, ini.c_str()) != 0;
    lightUI       = GetPrivateProfileIntA("General", "LightUI", 0, ini.c_str()) != 0;
    g_portInEnabled.store(GetPrivateProfileIntA("Ports", "PortInEnabled", 1, ini.c_str()) != 0);
    for (int p = 0; p < 4; ++p)
        g_portEnabled[p].store(GetPrivateProfileIntA("Ports", ("Port" + std::to_string(p+1) + "Enabled").c_str(), 1, ini.c_str()) != 0);

    auto loadColor = [&](const char* key, ImVec4 def) -> ImVec4 {
        char buf[64];
        if (GetPrivateProfileStringA("Colors", key, "", buf, 64, ini.c_str()) > 0) {
            float r, g, b, a;
            if (sscanf_s(buf, "%f,%f,%f,%f", &r, &g, &b, &a) == 4) return ImVec4(r,g,b,a);
        }
        return def;
    };
    cSys = loadColor("ColSys", ImVec4(0.6f, 0.6f, 0.6f, 1.0f));
    cIn  = loadColor("ColIn", ImVec4(0.4f, 0.8f, 1.0f, 1.0f));
    cOut[0] = loadColor("ColOut1", ImVec4(1.0f, 0.4f, 0.4f, 1.0f));
    cOut[1] = loadColor("ColOut2", ImVec4(1.0f, 0.8f, 0.4f, 1.0f));
    cOut[2] = loadColor("ColOut3", ImVec4(0.4f, 1.0f, 0.4f, 1.0f));
    cOut[3] = loadColor("ColOut4", ImVec4(0.8f, 0.4f, 1.0f, 1.0f));
}

void SetStartWithWindowsRegistry(bool enable) {
    HKEY hKey;
    if (RegOpenKeyExA(HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS) {
        if (enable) {
            char path[MAX_PATH];
            GetModuleFileNameA(NULL, path, MAX_PATH);
            RegSetValueExA(hKey, "ToHostBridge", 0, REG_SZ, (const BYTE*)path, strlen(path) + 1);
        } else {
            RegDeleteValueA(hKey, "ToHostBridge");
        }
        RegCloseKey(hKey);
    }
}

// Main code explicitly for Windows Subsystem
int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nCmdShow)
{
    HANDLE hMutex = CreateMutexW(NULL, TRUE, L"ToHostBridgeInstanceMutex");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        HWND hExisting = FindWindowW(L"ToHostBridge", nullptr);
        if (hExisting) {
            ShowWindow(hExisting, SW_RESTORE);
            SetForegroundWindow(hExisting);
        }
        return 0;
    }

    winrt::init_apartment();
    Microsoft::Windows::Devices::Midi2::Initialization::MidiDesktopAppSdkInitializer midiInit;
    bool midiAvailable = midiInit.InitializeSdkRuntime() && midiInit.EnsureServiceAvailable();

    ImGui_ImplWin32_EnableDpiAwareness();
    float main_scale = ImGui_ImplWin32_GetDpiScaleForMonitor(::MonitorFromPoint(POINT{ 0, 0 }, MONITOR_DEFAULTTOPRIMARY));

    WNDCLASSEXW wc = { sizeof(wc), CS_CLASSDC, WndProc, 0L, 0L, GetModuleHandle(nullptr), LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1)), nullptr, nullptr, nullptr, L"ToHostBridge", LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1)) };
    ::RegisterClassExW(&wc);
    // Adjusted window dimensions for bold headers and better spacing
    HWND hwnd = ::CreateWindowW(wc.lpszClassName, L"ToHost Bridge v1.0", WS_OVERLAPPEDWINDOW, 100, 100, (int)(530 * main_scale), (int)(430 * main_scale), nullptr, nullptr, wc.hInstance, nullptr);

    g_nid.cbSize = sizeof(NOTIFYICONDATAW);
    g_nid.hWnd = hwnd;
    g_nid.uID = 1;
    g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nid.uCallbackMessage = WM_TRAYICON;
    g_nid.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1));
    wcscpy_s(g_nid.szTip, L"ToHost Bridge");

    if (!CreateDeviceD3D(hwnd))
    {
        CleanupDeviceD3D();
        ::UnregisterClassW(wc.lpszClassName, wc.hInstance);
        return 1;
    }

    IMGUI_CHECKVERSION();
    ImGui::CreateContext();
    ImGuiIO& io = ImGui::GetIO(); (void)io;
    io.ConfigFlags |= ImGuiConfigFlags_NavEnableKeyboard;
    io.IniFilename = nullptr; // we manage imgui layout inside settings.ini ourselves
    
    // Load system font: Bold Segoe UI for the entire interface (DPI-aware baked size)
    float fontSize = 17.5f * main_scale;
    if (auto* fontBold = io.Fonts->AddFontFromFileTTF("C:\\Windows\\Fonts\\segoeuib.ttf", fontSize)) {
        // Default bold font loaded
    } else {
        io.Fonts->AddFontFromFileTTF("C:\\Windows\\Fonts\\arial.ttf", fontSize);
    }
    
    ImGui::StyleColorsDark();
    ImGuiStyle& style = ImGui::GetStyle();
    style.ScaleAllSizes(main_scale);
    // style.FontScaleDpi = main_scale; // Disabling this as we baked the scale into the high-quality font loading above

    ImGui_ImplWin32_Init(hwnd);
    ImGui_ImplDX11_Init(g_pd3dDevice, g_pd3dDeviceContext);

    ImVec4 clear_color = ImVec4(0.15f, 0.15f, 0.15f, 1.00f);

    // Dynamic COM Ports
    std::vector<std::string> comPortsVec = GetAvailableComPorts();
    std::vector<const char*> comPorts;
    for (const auto& p : comPortsVec) comPorts.push_back(p.c_str());
    std::string savedComName;
    char baseNameBuf[128] = "CBX";
    bool autoStartVirtualMidi = false;
    bool startWithWindows = false;
    bool autoReconnect = false;
    bool startToTray = false;
    bool stayOnTop = false;
    bool startMinimized = false;
    bool lightUI = false;
    ImVec4 colSys, colIn, colOut[4];

    // Init port-enable flags to true before LoadSettings (which may override)
    for (int i = 0; i < 4; ++i) g_portEnabled[i].store(true);

    LoadSettings(savedComName, baseNameBuf, autoStartVirtualMidi, startWithWindows, autoReconnect, startToTray, colSys, colIn, colOut, stayOnTop, startMinimized, lightUI);
    // Sync registry with loaded INI setting (especially important on first run after MSI install)
    SetStartWithWindowsRegistry(startWithWindows);

    // Load imgui layout AFTER LoadSettings so the INI file path is established
    LoadImguiIniFromSettings();
    if (lightUI)  ImGui::StyleColorsLight();
    if (stayOnTop) SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
    
    // Eagerly populate settings.ini on first boot
    if (GetFileAttributesA(GetIniPath().c_str()) == INVALID_FILE_ATTRIBUTES) {
        SaveSettings(savedComName, baseNameBuf, autoStartVirtualMidi, startWithWindows, autoReconnect, startToTray, colSys, colIn, colOut, stayOnTop, startMinimized, lightUI);
    }
    
    int selectedComPort = 0;
    bool foundSavedPort = false;
    if (!savedComName.empty()) {
        for (int i=0; i < (int)comPortsVec.size(); ++i) {
            if (comPortsVec[i] == savedComName) {
                selectedComPort = i;
                foundSavedPort = true;
                break;
            }
        }
    }
    
    // Auto start to system tray check
    if (startToTray) {
        ::Shell_NotifyIconW(NIM_ADD, &g_nid);
        ::ShowWindow(hwnd, SW_HIDE);
    } else if (startMinimized) {
        ::ShowWindow(hwnd, SW_SHOWMINIMIZED);
    } else {
        ::ShowWindow(hwnd, SW_SHOWDEFAULT);
    }
    ::UpdateWindow(hwnd);

    if (comPortsVec.size() > 0 && selectedComPort >= (int)comPortsVec.size()) selectedComPort = 0;

    bool isConnected = false;
    bool connectionLost = false;
    std::string activeComName = "";
    std::string connectionStatus = "";

    // Startup: if saved port not present yet, mark as waiting
    if (autoStartVirtualMidi && !savedComName.empty() && !foundSavedPort) {
        connectionLost = true;
        activeComName = savedComName;
        connectionStatus = "Waiting for " + savedComName + "...";
    }

    // Filters
    bool filterSys = true;
    bool filterIn = true;
    bool filterOut[4] = {true, true, true, true};
    bool filterSync = true;
    bool stopScroll = false;
    bool showNames = true;
    bool showTimestamp = true;
    size_t lastRenderedVersion = 0;

    // Cached visible snapshot
    std::vector<LogEntry> cachedSnapshot;
    size_t snapshotLogVersion = SIZE_MAX;
    bool stopScrollWasOn = false;
    std::vector<LogEntry> frozenSnapshot; // snapshot held while stop-scroll is active

    struct FilterState {
        bool sys, in_, sync;
        bool out[4];
        bool operator!=(const FilterState& o) const {
            return sys!=o.sys || in_!=o.in_ || sync!=o.sync ||
                   out[0]!=o.out[0] || out[1]!=o.out[1] || out[2]!=o.out[2] || out[3]!=o.out[3];
        }
    };
    FilterState lastFilterState = {true, true, true, {true,true,true,true}};

    // Lambda: build visible snapshot from per-source deques
    auto RebuildSnapshot = [&]() {
        std::lock_guard<std::mutex> lock(g_logMutex);
        // Merge all active sources, sorted by seqNum
        std::vector<LogEntry> merged;
        for (int s = 0; s < LOG_SOURCE_COUNT; ++s) {
            bool show = false;
            switch ((LogSourceType)s) {
                case LogSourceType::APP_INFO:    show = filterSys; break;
                case LogSourceType::COM_IN:      show = filterIn;  break;
                case LogSourceType::COM_IN_SYNC: show = filterIn && !filterSync; break;
                case LogSourceType::APP_OUT_1:   show = filterOut[0]; break;
                case LogSourceType::APP_OUT_2:   show = filterOut[1]; break;
                case LogSourceType::APP_OUT_3:   show = filterOut[2]; break;
                case LogSourceType::APP_OUT_4:   show = filterOut[3]; break;
            }
            if (!show) continue;
            for (const auto& e : g_logBySource[s]) {
                if (filterSync && e.isSync) continue;
                merged.push_back(e);
            }
        }
        // Sort by sequence number
        std::sort(merged.begin(), merged.end(), [](const LogEntry& a, const LogEntry& b){
            return a.seqNum < b.seqNum;
        });
        // Each source is individually capped — no merged trim needed
        cachedSnapshot = std::move(merged);
        snapshotLogVersion = g_logVersion;
        lastFilterState = {filterSys, filterIn, filterSync, {filterOut[0],filterOut[1],filterOut[2],filterOut[3]}};
    };

    // Pre-declare lambda for Connect
    auto ConnectPort = [&]() {
        if (isConnected || comPorts.empty() || !midiAvailable) return;
        std::string baseNameStr(baseNameBuf);
        g_serialPort.reset();

        std::string comName = comPorts[selectedComPort];
        activeComName = comName;

        // Restart virtual MIDI ports if they were stopped
        if (g_virtualPortsOut.empty() && !g_virtualPortIn) {
            std::wstring wBaseName(baseNameStr.begin(), baseNameStr.end());
            for (int i = 1; i <= 4; ++i) {
                if (!g_portEnabled[i-1].load(std::memory_order_relaxed)) continue;
                std::wstring portName = wBaseName + L" Out " + std::to_wstring(i);
                ImVec4 cOut = colOut[i-1];
                LogSourceType srcType = (LogSourceType)((int)LogSourceType::APP_OUT_1 + (i-1));
                auto vport = std::make_unique<VirtualMidiPort>(portName,
                    [i, cOut, srcType](const std::vector<uint8_t>& data) {
                        if (!g_transmitSync && !data.empty() && data[0] >= 0xF8) return;
                        bool isSync = (data.size() == 1 && data[0] >= 0xF8);
                        SendToSerial(i, data);
                        AddLog(srcType, "[App->COM] P" + std::to_string(i) + ": " + BytesToHex(data), cOut, isSync, GetMidiCommandName(data));
                    });
                g_virtualPortsOut.push_back(std::move(vport));
            }
            if (g_portInEnabled.load(std::memory_order_relaxed)) {
                std::wstring inPortName = wBaseName + L" In";
                g_virtualPortIn = std::make_unique<VirtualMidiPort>(inPortName, nullptr);
            }
            AddLog(LogSourceType::APP_INFO, "Virtual MIDI ports started.", colSys);
        }
        
        struct MidiParser {
            uint8_t runningStatus = 0;
            std::vector<uint8_t> buffer;
            bool isSysEx = false;
            std::vector<uint8_t> sysexBuffer;
        };
        auto parser = std::make_shared<MidiParser>();

        g_serialPort = std::make_unique<SerialPort>(comName, [colIn, parser](const std::vector<uint8_t>& data) {
            g_bytesReceived.fetch_add(data.size(), std::memory_order_relaxed);
            for (uint8_t b : data) {
                // Realtime bytes (High Priority)
                if (b >= 0xF8) {
                    std::vector<uint8_t> rt = { b };
                    if (g_transmitSync && g_portInEnabled.load(std::memory_order_relaxed)) {
                        if (g_virtualPortIn) g_virtualPortIn->SendMidi(rt);
                    }
                    AddLog(LogSourceType::COM_IN_SYNC, "[COM->App] " + BytesToHex(rt), colIn, true, GetMidiCommandName(rt));
                    continue;
                }

                // SysEx Handling
                if (b == 0xF0) {
                    parser->isSysEx = true;
                    parser->sysexBuffer = { 0xF0 };
                    continue;
                }
                if (parser->isSysEx) {
                    if (b == 0xF7) {
                        parser->sysexBuffer.push_back(0xF7);
                        if (g_portInEnabled.load(std::memory_order_relaxed) && g_virtualPortIn) 
                            g_virtualPortIn->SendMidi(parser->sysexBuffer);
                        AddLog(LogSourceType::COM_IN, "[COM->App] SysEx: " + BytesToHex(parser->sysexBuffer), colIn, false, "System Exclusive");
                        parser->isSysEx = false;
                    } else if (b >= 0x80) {
                        // Abort SysEx on any other status byte
                        parser->isSysEx = false;
                        // Process this byte normally below
                    } else {
                        parser->sysexBuffer.push_back(b);
                        continue;
                    }
                }

                // Standard MIDI Handling
                if (b >= 0x80 && b <= 0xEF) {
                    parser->runningStatus = b;
                    parser->buffer = { b };
                } else if (b < 0x80) {
                    if (parser->buffer.empty() && parser->runningStatus != 0) {
                        parser->buffer.push_back(parser->runningStatus);
                    }
                    parser->buffer.push_back(b);
                }
                
                if (parser->buffer.empty()) continue;
                uint8_t status = parser->buffer[0];
                uint8_t cmd = status & 0xF0;
                size_t expected = 0;
                if (cmd == 0xC0 || cmd == 0xD0) expected = 2;
                else if (cmd >= 0x80 && cmd <= 0xEF) expected = 3;
                
                if (expected > 0 && parser->buffer.size() == expected) {
                    if (g_portInEnabled.load(std::memory_order_relaxed) && g_virtualPortIn) 
                        g_virtualPortIn->SendMidi(parser->buffer);
                    AddLog(LogSourceType::COM_IN, "[COM->App] " + BytesToHex(parser->buffer), colIn, false, GetMidiCommandName(parser->buffer));
                    parser->buffer.clear();
                }
            }
        });

        // Reset port running-status so the receiver re-syncs on (re)connect
        g_lastSentPort.store(0, std::memory_order_relaxed);

        std::stringstream connMsg;
        if (g_serialPort->IsOpen()) {
            connMsg << "Opened " << comName << " successfully.\n";
            connMsg << "Virtual ports:\n";
            for (int i = 1; i <= 4; ++i) {
                if (g_portEnabled[i-1].load(std::memory_order_relaxed))
                    connMsg << "  " << baseNameStr << " Out " << i << "\n";
            }
            if (g_portInEnabled.load(std::memory_order_relaxed))
                connMsg << "  " << baseNameStr << " In\n";
            AddLog(LogSourceType::APP_INFO, "Connected to " + comName + " (" + baseNameStr + ")", colSys);
            isConnected = true;
            connectionLost = false;
            // Persist the newly-connected COM port name to settings
            SaveSettings(comName, baseNameStr, autoStartVirtualMidi, startWithWindows, autoReconnect, startToTray, colSys, colIn, colOut, stayOnTop, startMinimized, lightUI);
        } else {
            connMsg << "Failed to open " << comName << ". Is it in use?";
            AddLog(LogSourceType::APP_INFO, "Failed to open " + comName, colSys);
            g_serialPort.reset();
        }
        connectionStatus = connMsg.str();
    };

    bool firstFrame = true;
    bool done = false;
    while (!done)
    {
        MSG msg;
        while (::PeekMessage(&msg, nullptr, 0U, 0U, PM_REMOVE))
        {
            ::TranslateMessage(&msg);
            ::DispatchMessage(&msg);
            if (msg.message == WM_QUIT)
                done = true;
        }
        if (done) break;

        // First frame: start virtual ports in background, then optionally connect
        if (firstFrame) {
            firstFrame = false;
            // Capture everything needed by value so the detached thread is self-contained
            std::string savedNameCopy = savedComName;
            std::string baseNameCopy(baseNameBuf);
            bool autoStart = autoStartVirtualMidi;
            bool foundPort = foundSavedPort;
            ImVec4 colOutCopy[4] = { colOut[0], colOut[1], colOut[2], colOut[3] };
            ImVec4 colSysCopy = colSys;
            ImVec4 colInCopy  = colIn;
            std::vector<std::string> portsVecCopy = comPortsVec;
            bool midiOk = midiAvailable;

            std::thread([&, savedNameCopy, baseNameCopy, autoStart, foundPort,
                         colOutCopy, colSysCopy, colInCopy, portsVecCopy, midiOk]() mutable {
                // ── Create virtual MIDI ports (if auto-start is enabled) ──
                if (midiOk && autoStart) {
                    std::wstring wBaseName(baseNameCopy.begin(), baseNameCopy.end());
                    for (int i = 1; i <= 4; ++i) {
                        if (!g_portEnabled[i-1].load(std::memory_order_relaxed)) continue;
                        std::wstring portName = wBaseName + L" Out " + std::to_wstring(i);
                        ImVec4 cOut = colOutCopy[i-1];
                        LogSourceType srcType = (LogSourceType)((int)LogSourceType::APP_OUT_1 + (i-1));
                        auto vport = std::make_unique<VirtualMidiPort>(portName,
                            [i, cOut, srcType](const std::vector<uint8_t>& data) {
                                if (!g_transmitSync && !data.empty() && data[0] >= 0xF8) return;
                                bool isSync = (data.size() == 1 && data[0] >= 0xF8);
                                SendToSerial(i, data);
                                AddLog(srcType, "[App->COM] P" + std::to_string(i) + ": " + BytesToHex(data), cOut, isSync, GetMidiCommandName(data));
                            });
                        g_virtualPortsOut.push_back(std::move(vport));
                    }
                    if (g_portInEnabled.load(std::memory_order_relaxed)) {
                        std::wstring inPortName = wBaseName + L" In";
                        g_virtualPortIn = std::make_unique<VirtualMidiPort>(inPortName, nullptr);
                    }
                    AddLog(LogSourceType::APP_INFO, "Virtual MIDI ports started.", colSysCopy);
                }

                // ── Auto-connect to COM port ────────────────────────────
                if (autoStart && !portsVecCopy.empty()) {
                    if (!savedNameCopy.empty()) {
                        for (int i = 0; i < (int)portsVecCopy.size(); ++i) {
                            if (portsVecCopy[i] == savedNameCopy) {
                                selectedComPort = i;
                                ConnectPort();
                                return;
                            }
                        }
                        // Port not available yet — connectionLost already set in main
                    } else {
                        ConnectPort(); // no saved name, connect to first
                    }
                }
            }).detach();
        }

        if (g_SwapChainOccluded && g_pSwapChain->Present(0, DXGI_PRESENT_TEST) == DXGI_STATUS_OCCLUDED)
        {
            ::Sleep(10);
            continue;
        }
        g_SwapChainOccluded = false;

        if (g_ResizeWidth != 0 && g_ResizeHeight != 0)
        {
            CleanupRenderTarget();
            g_pSwapChain->ResizeBuffers(0, g_ResizeWidth, g_ResizeHeight, DXGI_FORMAT_UNKNOWN, 0);
            g_ResizeWidth = g_ResizeHeight = 0;
            CreateRenderTarget();
        }

        // ── Auto-Reconnect Polling + Bandwidth (every 1 second) ───────────────
        static auto lastCheck = std::chrono::steady_clock::now();
        static uint64_t prevSent = 0, prevRecv = 0;
        auto nowCheck = std::chrono::steady_clock::now();
        if (std::chrono::duration_cast<std::chrono::milliseconds>(nowCheck - lastCheck).count() > 1000) {
            lastCheck = nowCheck;

            // Bandwidth: 38400 baud, 8N1 → 10 bits/byte → 3840 bytes/sec max
            uint64_t curSent = g_bytesSent.load(std::memory_order_relaxed);
            uint64_t curRecv = g_bytesReceived.load(std::memory_order_relaxed);
            double sentRate = (double)(curSent - prevSent);
            double recvRate = (double)(curRecv - prevRecv);
            double maxRate = (sentRate > recvRate) ? sentRate : recvRate;
            // Calibration: While 38400 baud is 3840 B/s, practical synth processing 
            // and multiplexing overhead make the "struggle" point much lower.
            double saturationPoint = 800.0; 
            double chargeRaw = (maxRate / saturationPoint) * 100.0;
            prevSent = curSent; prevRecv = curRecv;
            
            float newCharge = isConnected ? ((float)chargeRaw > 100.0f ? 100.0f : (float)chargeRaw) : 0.0f;
            // Smooth moving average (alpha=0.4) to avoid instant spikes on transient bursts
            g_chargePercent = (g_chargePercent * 0.6f) + (newCharge * 0.4f);
            // Update Win32 title bar
            std::wstring wTitle = L"ToHost Bridge v1.0";

            if (!activeComName.empty()) {
                char dummyPath[1000];
                DWORD res = QueryDosDeviceA(activeComName.c_str(), dummyPath, sizeof(dummyPath));
                bool portExists = (res != 0);

                // Detect sudden disconnect
                if (isConnected && !portExists) {
                    isConnected = false;
                    connectionLost = true;
                    g_serialPort.reset();   // close the dead handle cleanly
                    AddLog(LogSourceType::APP_INFO, "COM port lost: " + activeComName, colSys);
                    connectionStatus = "COM port disconnected!";
                }

                // Auto-reconnect when port reappears
                if (!isConnected && connectionLost && portExists && autoReconnect) {
                    // Rebuild port list so dropdown + index are correct
                    comPortsVec = GetAvailableComPorts();
                    comPorts.clear();
                    for (const auto& p : comPortsVec) comPorts.push_back(p.c_str());
                    if (comPortsVec.size() > 0 && selectedComPort >= (int)comPortsVec.size()) selectedComPort = 0;
                    for (int i = 0; i < (int)comPortsVec.size(); ++i) {
                        if (comPortsVec[i] == activeComName) { selectedComPort = i; break; }
                    }
                    connectionLost = false;
                    ConnectPort();
                }
            }
        }

        ImGui_ImplDX11_NewFrame();
        ImGui_ImplWin32_NewFrame();
        ImGui::NewFrame();

        ImGui::SetNextWindowPos(ImVec2(0, 0));
        ImGui::SetNextWindowSize(io.DisplaySize);
        ImGui::Begin("Main", nullptr, ImGuiWindowFlags_NoTitleBar | ImGuiWindowFlags_NoResize | ImGuiWindowFlags_NoMove);

        // Header
        if (!midiAvailable) {
            ImGui::TextColored(ImVec4(1.0f, 0.2f, 0.2f, 1.0f), "ERROR: Windows MIDI Services Runtime is missing.");
            if (ImGui::Button("Download Windows MIDI Services (x64 Installer)")) {
                ShellExecuteA(NULL, "open", "https://microsoft.github.io/MIDI/get-latest/#:~:text=Download%20Latest%20x64%20Installer", NULL, NULL, SW_SHOWNORMAL);
            }
        }
        ImGui::Separator();

        if (ImGui::BeginTabBar("Tabs")) {

            // ── Connection Tab ─────────────────────────────────────────────
            if (ImGui::BeginTabItem("Connection")) {
                ImGui::Spacing();
                bool portsRunning = !g_virtualPortsOut.empty();
                if (portsRunning || isConnected) ImGui::BeginDisabled();

                if (ImGui::Button("Refresh COM Ports")) {
                    comPortsVec = GetAvailableComPorts();
                    comPorts.clear();
                    for (const auto& p : comPortsVec) comPorts.push_back(p.c_str());
                    selectedComPort = 0;
                }
                ImGui::SameLine();
                
                // Port Enables
                {
                    bool enIn = g_portInEnabled.load(std::memory_order_relaxed);
                    if (ImGui::Checkbox("In##pe", &enIn)) g_portInEnabled.store(enIn, std::memory_order_relaxed);
                    ImGui::SameLine();
                    for (int p = 0; p < 4; ++p) {
                        bool en = g_portEnabled[p].load(std::memory_order_relaxed);
                        std::string lbl = "Out " + std::to_string(p+1) + "##pe";
                        if (ImGui::Checkbox(lbl.c_str(), &en)) g_portEnabled[p].store(en, std::memory_order_relaxed);
                        if (p < 3) ImGui::SameLine();
                    }
                }

                ImGui::SetNextItemWidth(150);
                if (ImGui::InputText("Virtual Port Name", baseNameBuf, 128)) {
                    SaveSettings("", baseNameBuf, autoStartVirtualMidi, startWithWindows, autoReconnect, startToTray, colSys, colIn, colOut, stayOnTop, startMinimized, lightUI);
                }
                
                if (portsRunning) ImGui::EndDisabled();

                ImGui::Spacing();
                
                bool disableComSelect = isConnected || comPorts.empty();
                if (disableComSelect) ImGui::BeginDisabled();
                ImGui::SetNextItemWidth(120);
                ImGui::Combo("COM Port", &selectedComPort, comPorts.empty() ? nullptr : comPorts.data(), (int)comPorts.size());
                if (disableComSelect) ImGui::EndDisabled();
                
                ImGui::Spacing();
                ImGui::BeginDisabled(isConnected);
                if (ImGui::Button(isConnected ? "Connected" : "Connect", ImVec2(120, 30)) && midiAvailable && !comPorts.empty()) {
                    if (!isConnected) ConnectPort();
                }
                ImGui::EndDisabled();
                
                // Disconnect button
                if (isConnected) {
                    ImGui::SameLine();
                    if (ImGui::Button("Disconnect", ImVec2(120, 30))) {
                        g_serialPort.reset();
                        g_lastSentPort.store(0, std::memory_order_relaxed);
                        isConnected = false;
                        connectionLost = false;
                        connectionStatus = "Disconnected (virtual ports still active).";
                        AddLog(LogSourceType::APP_INFO, "Disconnected from COM port.", colSys);
                    }
                }

                ImGui::Spacing();
                ImGui::Separator();
                ImGui::Spacing();
                if (connectionLost) {
                    ImGui::TextColored(ImVec4(1.0f, 0.2f, 0.2f, 1.0f), "COM port disconnected!");
                } else {
                    ImGui::TextWrapped("%s", connectionStatus.c_str());
                }
                if (isConnected) {
                    // Refined 'Hunter Green' for Light UI (darker for contrast), bright green for Dark UI
                    ImVec4 chargeCol = lightUI ? ImVec4(0.0f, 0.45f, 0.0f, 1.0f) : ImVec4(0.6f, 1.0f, 0.6f, 1.0f);
                    ImGui::TextColored(chargeCol, "Serial charge: %.0f%%", g_chargePercent);

                    ImGui::Spacing();
                    ImGui::Separator();
                    ImGui::Spacing();
                    
                    ImGui::PushStyleColor(ImGuiCol_Button, ImVec4(0.7f, 0.1f, 0.1f, 1.0f));
                    ImGui::PushStyleColor(ImGuiCol_ButtonHovered, ImVec4(0.9f, 0.2f, 0.2f, 1.0f));
                    ImGui::PushStyleColor(ImGuiCol_ButtonActive, ImVec4(0.5f, 0.0f, 0.0f, 1.0f));
                    ImGui::PushStyleColor(ImGuiCol_Text, ImVec4(1.0f, 1.0f, 1.0f, 1.0f)); // Force white text for readability
                    if (ImGui::Button("Panic! (All Notes Off)", ImVec2(ImGui::GetContentRegionAvail().x * 0.5f, 30))) {
                        for (int p = 1; p <= 4; ++p) {
                            std::vector<uint8_t> pData;
                            for (int ch = 0; ch < 16; ++ch) {
                                pData.insert(pData.end(), { (uint8_t)(0xB0 | ch), 120, 0 });
                                pData.insert(pData.end(), { (uint8_t)(0xB0 | ch), 123, 0 });
                                pData.insert(pData.end(), { (uint8_t)(0xB0 | ch), 121, 0 });
                            }
                            SendToSerial(p, pData);
                        }
                        AddLog(LogSourceType::APP_INFO, "Sent Panic (All Sound/Notes Off, Reset CC) to all ports/channels.", ImVec4(1.0f, 0.5f, 0.0f, 1.0f));
                    }
                    ImGui::PopStyleColor(4);

                    ImGui::SameLine();
                    ImGui::PushStyleColor(ImGuiCol_Button, ImVec4(0.1f, 0.3f, 0.6f, 1.0f));
                    ImGui::PushStyleColor(ImGuiCol_ButtonHovered, ImVec4(0.2f, 0.4f, 0.8f, 1.0f));
                    ImGui::PushStyleColor(ImGuiCol_ButtonActive, ImVec4(0.0f, 0.2f, 0.5f, 1.0f));
                    ImGui::PushStyleColor(ImGuiCol_Text, ImVec4(1.0f, 1.0f, 1.0f, 1.0f)); // Force white text for readability
                    if (ImGui::Button("XG Reset", ImVec2(-1, 30))) {
                        for (int p = 1; p <= 4; ++p) {
                            if (!g_portEnabled[p-1].load()) continue;
                            std::vector<uint8_t> resetMsg = { 0xF0, 0x43, 0x10, 0x4C, 0x00, 0x00, 0x7E, 0x00, 0xF7 };
                            SendToSerial(p, resetMsg);
                        }
                        AddLog(LogSourceType::APP_INFO, "Sent XG System On Reset to all active ports.", ImVec4(0.4f, 0.6f, 1.0f, 1.0f));
                    }
                    ImGui::PopStyleColor(4);
                }

                // Stop Virtual Ports button — show when not connected but ports are alive
                if (!isConnected && !g_virtualPortsOut.empty()) {
                    ImGui::Spacing();
                    ImGui::PushStyleColor(ImGuiCol_Button,        ImVec4(0.6f, 0.3f, 0.0f, 1.0f));
                    ImGui::PushStyleColor(ImGuiCol_ButtonHovered, ImVec4(0.8f, 0.4f, 0.0f, 1.0f));
                    ImGui::PushStyleColor(ImGuiCol_ButtonActive,  ImVec4(0.5f, 0.2f, 0.0f, 1.0f));
                    if (ImGui::Button("Stop Virtual Ports", ImVec2(-1, 28))) {
                        g_virtualPortsOut.clear();
                        g_virtualPortIn.reset();
                        connectionStatus = "Disconnected.";
                        AddLog(LogSourceType::APP_INFO, "Virtual MIDI ports stopped.", colSys);
                    }
                    ImGui::PopStyleColor(3);
                }

                ImGui::EndTabItem();
            }

            // ── Debug View Tab ─────────────────────────────────────────────
            if (ImGui::BeginTabItem("Debug View")) {
                ImGui::Spacing();
                // Filters Row 1
                ImGui::Checkbox("In", &filterIn); ImGui::SameLine();
                ImGui::Checkbox("Out 1", &filterOut[0]); ImGui::SameLine();
                ImGui::Checkbox("Out 2", &filterOut[1]); ImGui::SameLine();
                ImGui::Checkbox("Out 3", &filterOut[2]); ImGui::SameLine();
                ImGui::Checkbox("Out 4", &filterOut[3]); ImGui::SameLine();
                ImGui::Checkbox("Timestamp", &showTimestamp); 
                // Filters Row 2
                ImGui::Checkbox("System Info", &filterSys); ImGui::SameLine();
                ImGui::Checkbox("Filter Sync", &filterSync); ImGui::SameLine();
                bool stopScrollPrev = stopScroll;
                ImGui::Checkbox("Stop Scroll", &stopScroll); ImGui::SameLine();
                ImGui::Checkbox("Info", &showNames);
                ImGui::Separator();

                FilterState currentFilter = {filterSys, filterIn, filterSync, {filterOut[0],filterOut[1],filterOut[2],filterOut[3]}};
                size_t currentLogVersion;
                {
                    std::lock_guard<std::mutex> lock(g_logMutex);
                    currentLogVersion = g_logVersion;
                }

                bool filterChanged = (currentFilter != lastFilterState);
                bool logChanged    = (currentLogVersion != snapshotLogVersion);

                static float stopScrollCapturedRatio = 1.0f;
                static bool stopScrollFirstFrame = false;

                if (stopScroll) {
                    // Just toggled ON — capture frozen snapshot now
                    if (!stopScrollWasOn) {
                        if (filterChanged || logChanged || cachedSnapshot.empty()) {
                            RebuildSnapshot();
                        }
                        frozenSnapshot = cachedSnapshot;
                        stopScrollWasOn = true;
                        stopScrollFirstFrame = true;
                        snapshotLogVersion = SIZE_MAX; // force buffer rebuild once
                    }
                    
                    // Render frozen snapshot as selectable text area for copying
                    static std::string frozenBuffer;
                    if (logChanged || snapshotLogVersion == SIZE_MAX) {
                        frozenBuffer.clear();
                        for (auto& l : frozenSnapshot) {
                            frozenBuffer += (showTimestamp ? (l.timestamp + l.text) : l.text);
                            if (showNames && !l.cmdName.empty()) frozenBuffer += " " + l.cmdName;
                            frozenBuffer += "\r\n";
                        }
                        snapshotLogVersion = currentLogVersion;
                    }

                    ImGui::BeginChild("LogRegion", ImVec2(0, -ImGui::GetTextLineHeightWithSpacing() * 2.5f), true);
                    if (stopScrollFirstFrame) {
                        ImGui::SetScrollY(ImGui::GetScrollMaxY());
                        stopScrollFirstFrame = false;
                    }
                    ImGui::InputTextMultiline("##frozenView", (char*)frozenBuffer.c_str(), frozenBuffer.size(), ImVec2(-1, -1), ImGuiInputTextFlags_ReadOnly);
                    ImGui::EndChild();
                } else {
                    // Stop scroll turned OFF — resume live view
                    if (stopScrollWasOn) {
                        stopScrollWasOn = false;
                        snapshotLogVersion = SIZE_MAX; // force live rebuild
                    }

                    if (filterChanged || logChanged) {
                        RebuildSnapshot();
                    }

                    ImGui::BeginChild("LogRegion", ImVec2(0, -ImGui::GetTextLineHeightWithSpacing() * 1.8f), true, ImGuiWindowFlags_HorizontalScrollbar);
                    for (const auto& line : cachedSnapshot) {
                        ImVec4 drawCol = line.color;
                        std::string display = showTimestamp ? (line.timestamp + line.text) : line.text;
                        if (showNames && !line.cmdName.empty())
                            ImGui::TextColored(drawCol, "%s %s", display.c_str(), line.cmdName.c_str());
                        else
                            ImGui::TextColored(drawCol, "%s", display.c_str());
                    }

                    // Auto-scroll only on new items
                    bool hasNewItems = logChanged && !filterChanged;
                    if (hasNewItems) {
                        ImGui::SetScrollHereY(1.0f);
                        lastRenderedVersion = currentLogVersion;
                    }
                    ImGui::EndChild();
                }

                ImGui::Separator();
                if (ImGui::Button("Clear Logs")) {
                    std::lock_guard<std::mutex> lock(g_logMutex);
                    for (int i = 0; i < LOG_SOURCE_COUNT; ++i) g_logBySource[i].clear();
                    g_logVersion++;
                    cachedSnapshot.clear();
                    frozenSnapshot.clear();
                }

                ImGui::EndTabItem();
            }

            // ── Settings Tab ───────────────────────────────────────────────
            if (ImGui::BeginTabItem("Settings")) {
                ImGui::Spacing();
                
                bool settingsChanged = false;

                if (ImGui::Checkbox("Reduce to System Tray on minimize", &g_reduceToTray)) {
                    settingsChanged = true;
                }
                
                if (ImGui::Checkbox("Start with Windows", &startWithWindows)) {
                    SetStartWithWindowsRegistry(startWithWindows);
                    settingsChanged = true;
                }

                if (ImGui::Checkbox("Stay on top", &stayOnTop)) {
                    SetWindowPos(hwnd, stayOnTop ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
                    settingsChanged = true;
                }

                if (ImGui::Checkbox("Start minimized", &startMinimized)) {
                    settingsChanged = true;
                }

                if (ImGui::Checkbox("Light UI", &lightUI)) {
                    if (lightUI) {
                        ImGui::StyleColorsLight();
                        ImGui::GetStyle().Colors[ImGuiCol_Text]         = ImVec4(0.00f, 0.00f, 0.00f, 1.00f);
                        ImGui::GetStyle().Colors[ImGuiCol_WindowBg]     = ImVec4(0.95f, 0.95f, 0.95f, 1.00f);
                        ImGui::GetStyle().Colors[ImGuiCol_ChildBg]      = ImVec4(1.00f, 1.00f, 1.00f, 0.50f);
                        ImGui::GetStyle().Colors[ImGuiCol_Border]       = ImVec4(0.00f, 0.00f, 0.00f, 0.15f);
                        ImGui::GetStyle().Colors[ImGuiCol_Button]       = ImVec4(0.85f, 0.85f, 0.88f, 1.00f);
                        ImGui::GetStyle().Colors[ImGuiCol_ButtonHovered]= ImVec4(0.75f, 0.75f, 0.78f, 1.00f);
                        ImGui::GetStyle().Colors[ImGuiCol_ButtonActive] = ImVec4(0.65f, 0.65f, 0.68f, 1.00f);
                        
                        // Switch to High-Contrast Light Mode Defaults
                        colSys  = ImVec4(0.30f, 0.30f, 0.30f, 1.00f);
                        colIn   = ImVec4(0.00f, 0.45f, 0.65f, 1.00f);
                        colOut[0] = ImVec4(0.70f, 0.00f, 0.00f, 1.00f);
                        colOut[1] = ImVec4(0.60f, 0.35f, 0.00f, 1.00f);
                        colOut[2] = ImVec4(0.00f, 0.55f, 0.00f, 1.00f);
                        colOut[3] = ImVec4(0.45f, 0.00f, 0.65f, 1.00f);
                    } else {
                        ImGui::StyleColorsDark();
                        // Restore Vibrant Dark Mode Defaults
                        colSys  = ImVec4(0.60f, 0.60f, 0.60f, 1.00f);
                        colIn   = ImVec4(0.40f, 0.80f, 1.00f, 1.00f);
                        colOut[0] = ImVec4(1.00f, 0.40f, 0.40f, 1.00f);
                        colOut[1] = ImVec4(1.00f, 0.80f, 0.40f, 1.00f);
                        colOut[2] = ImVec4(0.40f, 1.00f, 0.40f, 1.00f);
                        colOut[3] = ImVec4(0.80f, 0.40f, 1.00f, 1.00f);
                    }
                    settingsChanged = true;
                }
                
                if (ImGui::Checkbox("Auto-start Virtual MIDI on launch", &autoStartVirtualMidi)) {
                    settingsChanged = true;
                }
                
                if (ImGui::Checkbox("Auto-reconnect dropped COM port", &autoReconnect)) {
                    settingsChanged = true;
                }
                
                if (ImGui::Checkbox("Transmit Sync Data (Clock/Start/Stop/ActiveSense)", &g_transmitSync)) {
                    settingsChanged = true;
                }

                if (ImGui::SliderInt("Debug View Max Lines", &g_maxLogLines, 100, 1000)) {
                    settingsChanged = true;
                }
                
                ImGui::Separator();
                ImGui::Text("Debug View Colors");
                ImGui::Spacing();
                
                ImGui::Columns(3, nullptr, false);
                ImGui::SetColumnWidth(0, 110 * main_scale);
                ImGui::SetColumnWidth(1, 110 * main_scale);
                ImGui::SetColumnWidth(2, 110 * main_scale);
                
                // Row 1
                if (ImGui::ColorEdit4("Out 1", (float*)&colOut[0], ImGuiColorEditFlags_NoInputs)) settingsChanged = true;
                ImGui::NextColumn();
                if (ImGui::ColorEdit4("Out 3", (float*)&colOut[2], ImGuiColorEditFlags_NoInputs)) settingsChanged = true;
                ImGui::NextColumn();
                if (ImGui::ColorEdit4("COM Input", (float*)&colIn, ImGuiColorEditFlags_NoInputs)) settingsChanged = true;
                ImGui::NextColumn();
                
                // Row 2
                if (ImGui::ColorEdit4("Out 2", (float*)&colOut[1], ImGuiColorEditFlags_NoInputs)) settingsChanged = true;
                ImGui::NextColumn();
                if (ImGui::ColorEdit4("Out 4", (float*)&colOut[3], ImGuiColorEditFlags_NoInputs)) settingsChanged = true;
                ImGui::NextColumn();
                if (ImGui::ColorEdit4("System Info", (float*)&colSys, ImGuiColorEditFlags_NoInputs)) settingsChanged = true;
                
                ImGui::Columns(1);

                if (settingsChanged) {
                    std::string cName = "";
                    if (!comPortsVec.empty() && selectedComPort >= 0 && selectedComPort < (int)comPortsVec.size()) {
                        cName = comPortsVec[selectedComPort];
                    }
                    SaveSettings(cName, baseNameBuf, autoStartVirtualMidi, startWithWindows, autoReconnect, startToTray, colSys, colIn, colOut, stayOnTop, startMinimized, lightUI);
                }
                
                ImGui::EndTabItem();
            }
            ImGui::EndTabBar();
        }

        ImGui::End();

        // Rendering
        ImGui::Render();
        const float clear_color_with_alpha[4] = { clear_color.x * clear_color.w, clear_color.y * clear_color.w, clear_color.z * clear_color.w, clear_color.w };
        g_pd3dDeviceContext->OMSetRenderTargets(1, &g_mainRenderTargetView, nullptr);
        g_pd3dDeviceContext->ClearRenderTargetView(g_mainRenderTargetView, clear_color_with_alpha);
        ImGui_ImplDX11_RenderDrawData(ImGui::GetDrawData());

        HRESULT hr = g_pSwapChain->Present(1, 0);
        g_SwapChainOccluded = (hr == DXGI_STATUS_OCCLUDED);
    }

    // Exit Cleanup: Silence all and stop playback
    if (isConnected && g_serialPort && g_serialPort->IsOpen()) {
        for (int p = 1; p <= 4; ++p) {
            std::vector<uint8_t> quiet;
            for (int ch = 0; ch < 16; ++ch) {
                quiet.insert(quiet.end(), { (uint8_t)(0xB0 | ch), 0x7B, 0x00 });
            }
            quiet.push_back(0xFC); // MIDI Stop
            SendToSerial(p, quiet);
        }
        
        // Also send MIDI Stop to virtual Input if active
        if (g_virtualPortIn) {
            g_virtualPortIn->SendMidi({ 0xFC });
        }
    }

    g_serialPort.reset();
    g_virtualPortsOut.clear();
    g_virtualPortIn.reset();

    // Persist imgui layout into settings.ini before destroying context
    SaveImguiIniToSettings();

    ImGui_ImplDX11_Shutdown();
    ImGui_ImplWin32_Shutdown();
    ImGui::DestroyContext();

    CleanupDeviceD3D();
    ::DestroyWindow(hwnd);
    ::UnregisterClassW(wc.lpszClassName, wc.hInstance);

    return 0;
}

bool CreateDeviceD3D(HWND hWnd)
{
    DXGI_SWAP_CHAIN_DESC sd;
    ZeroMemory(&sd, sizeof(sd));
    sd.BufferCount = 2;
    sd.BufferDesc.Width = 0;
    sd.BufferDesc.Height = 0;
    sd.BufferDesc.Format = DXGI_FORMAT_R8G8B8A8_UNORM;
    sd.BufferDesc.RefreshRate.Numerator = 60;
    sd.BufferDesc.RefreshRate.Denominator = 1;
    sd.Flags = DXGI_SWAP_CHAIN_FLAG_ALLOW_MODE_SWITCH;
    sd.BufferUsage = DXGI_USAGE_RENDER_TARGET_OUTPUT;
    sd.OutputWindow = hWnd;
    sd.SampleDesc.Count = 1;
    sd.SampleDesc.Quality = 0;
    sd.Windowed = TRUE;
    sd.SwapEffect = DXGI_SWAP_EFFECT_DISCARD;

    UINT createDeviceFlags = 0;
    D3D_FEATURE_LEVEL featureLevel;
    const D3D_FEATURE_LEVEL featureLevelArray[2] = { D3D_FEATURE_LEVEL_11_0, D3D_FEATURE_LEVEL_10_0, };
    HRESULT res = D3D11CreateDeviceAndSwapChain(nullptr, D3D_DRIVER_TYPE_HARDWARE, nullptr, createDeviceFlags, featureLevelArray, 2, D3D11_SDK_VERSION, &sd, &g_pSwapChain, &g_pd3dDevice, &featureLevel, &g_pd3dDeviceContext);
    if (res == DXGI_ERROR_UNSUPPORTED)
        res = D3D11CreateDeviceAndSwapChain(nullptr, D3D_DRIVER_TYPE_WARP, nullptr, createDeviceFlags, featureLevelArray, 2, D3D11_SDK_VERSION, &sd, &g_pSwapChain, &g_pd3dDevice, &featureLevel, &g_pd3dDeviceContext);
    if (res != S_OK) return false;

    CreateRenderTarget();
    return true;
}

void CleanupDeviceD3D()
{
    CleanupRenderTarget();
    if (g_pSwapChain) { g_pSwapChain->Release(); g_pSwapChain = nullptr; }
    if (g_pd3dDeviceContext) { g_pd3dDeviceContext->Release(); g_pd3dDeviceContext = nullptr; }
    if (g_pd3dDevice) { g_pd3dDevice->Release(); g_pd3dDevice = nullptr; }
}

void CreateRenderTarget()
{
    ID3D11Texture2D* pBackBuffer;
    g_pSwapChain->GetBuffer(0, IID_PPV_ARGS(&pBackBuffer));
    g_pd3dDevice->CreateRenderTargetView(pBackBuffer, nullptr, &g_mainRenderTargetView);
    pBackBuffer->Release();
}

void CleanupRenderTarget()
{
    if (g_mainRenderTargetView) { g_mainRenderTargetView->Release(); g_mainRenderTargetView = nullptr; }
}

extern IMGUI_IMPL_API LRESULT ImGui_ImplWin32_WndProcHandler(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);

LRESULT WINAPI WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    if (ImGui_ImplWin32_WndProcHandler(hWnd, msg, wParam, lParam))
        return true;

    switch (msg)
    {
    case WM_SIZE:
        if (wParam == SIZE_MINIMIZED) {
            if (g_reduceToTray) {
                Shell_NotifyIconW(NIM_ADD, &g_nid);
                ShowWindow(hWnd, SW_HIDE);
            }
            return 0;
        }
        g_ResizeWidth = (UINT)LOWORD(lParam);
        g_ResizeHeight = (UINT)HIWORD(lParam);
        return 0;
    case WM_TRAYICON:
        if (lParam == WM_LBUTTONUP || lParam == WM_RBUTTONUP || lParam == WM_LBUTTONDBLCLK) {
            ShowWindow(hWnd, SW_RESTORE);
            SetForegroundWindow(hWnd);
            Shell_NotifyIconW(NIM_DELETE, &g_nid);
        }
        return 0;
    case WM_SYSCOMMAND:
        if ((wParam & 0xfff0) == SC_KEYMENU) return 0;
        break;
    case WM_DESTROY:
        Shell_NotifyIconW(NIM_DELETE, &g_nid);
        ::PostQuitMessage(0);
        return 0;
    }
    return ::DefWindowProcW(hWnd, msg, wParam, lParam);
}

// ── Centralized Serial Writing with Multiplexing ──────────────────────────────────
void SendToSerial(int portIdx, const std::vector<uint8_t>& data) {
    if (!g_serialPort || !g_serialPort->IsOpen() || data.empty()) return;

    std::lock_guard<std::mutex> lock(g_serialWriteMutex);
    std::vector<uint8_t> outData;

    // Yamaha Multiplexing (Serial TO HOST protocol):
    // Prepend 0xF5 <port> if we are switching from a different virtual port.
    // 1-indexed ports: 1-4
    if (g_lastSentPort.load(std::memory_order_relaxed) != portIdx) {
        outData.push_back(0xF5);
        outData.push_back((uint8_t)portIdx);
        g_lastSentPort.store(portIdx, std::memory_order_relaxed);
    }

    outData.insert(outData.end(), data.begin(), data.end());
    g_serialPort->Write(outData);
    g_bytesSent.fetch_add(outData.size(), std::memory_order_relaxed);
}
