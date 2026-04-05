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
#include <winrt/Windows.ApplicationModel.h>
#include <winrt/Windows.Services.Store.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <shobjidl.h> // For IInitializeWithWindow
#include <winmidi/init/Microsoft.Windows.Devices.Midi2.Initialization.hpp>

#include "SerialPort.h"
#include "VirtualMidiPort.h"

void SendToSerial(int portIdx, const std::vector<uint8_t>& data);

#include <shellapi.h>
#include <shlobj.h>
#include <appmodel.h>

#define WM_TRAYICON (WM_APP + 1)
#define IDI_ICON1 101
NOTIFYICONDATAW g_nid = {};

// Global config
bool g_reduceToTray  = false;
bool g_transmitSync  = true;
std::atomic<bool> g_isShuttingDown{false};
std::atomic<bool> g_cleanupStarted{false};
std::atomic<bool> g_autoStartAttempted{false};
std::atomic<bool> g_readyToExit{false};
std::string g_shutdownSubStatus = "";
struct ShutdownTask { std::string name; std::string status; };
std::vector<ShutdownTask> g_shutdownTasks;
std::mutex g_shutdownMutex;

// Store Update State
bool g_updateAvailable = false;
bool g_isUpdating = false;
winrt::Windows::Foundation::Collections::IVectorView<winrt::Windows::Services::Store::StorePackageUpdate> g_availableUpdates{ nullptr };
bool g_mockUpdate = false;

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
std::shared_ptr<SerialPort> g_serialPort;
std::shared_ptr<VirtualMidiPort> g_virtualPortIn;
std::vector<std::shared_ptr<VirtualMidiPort>> g_virtualPortsOut;
std::mutex g_midiPtrMutex;

std::shared_ptr<SerialPort> GetSafeSerialPort() {
    std::lock_guard<std::mutex> lock(g_midiPtrMutex);
    return g_serialPort;
}

void ResetSafeSerialPort() {
    std::shared_ptr<SerialPort> sp;
    {
        std::lock_guard<std::mutex> lock(g_midiPtrMutex);
        sp = std::move(g_serialPort);
        g_serialPort = nullptr;
    } 
}

void SetSafeSerialPort(std::shared_ptr<SerialPort> sp) {
    std::lock_guard<std::mutex> lock(g_midiPtrMutex);
    g_serialPort = sp;
}

std::shared_ptr<VirtualMidiPort> GetSafeVirtualPortIn() {
    std::lock_guard<std::mutex> lock(g_midiPtrMutex);
    return g_virtualPortIn;
}

void SetSafeVirtualPortIn(std::shared_ptr<VirtualMidiPort> vp) {
    std::lock_guard<std::mutex> lock(g_midiPtrMutex);
    g_virtualPortIn = vp;
}

void ResetSafeVirtualPortIn() {
    std::shared_ptr<VirtualMidiPort> vpi;
    {
        std::lock_guard<std::mutex> lock(g_midiPtrMutex);
        vpi = std::move(g_virtualPortIn);
        g_virtualPortIn = nullptr;
    }
}

std::vector<std::shared_ptr<VirtualMidiPort>> GetSafeVirtualPortsOut() {
    std::lock_guard<std::mutex> lock(g_midiPtrMutex);
    return g_virtualPortsOut;
}

void ClearSafeVirtualPortsOut() {
    std::vector<std::shared_ptr<VirtualMidiPort>> toClean;
    {
        std::lock_guard<std::mutex> lock(g_midiPtrMutex);
        toClean = std::move(g_virtualPortsOut);
        g_virtualPortsOut.clear();
    }
}

void AddSafeVirtualPortOut(std::shared_ptr<VirtualMidiPort> vp) {
    std::lock_guard<std::mutex> lock(g_midiPtrMutex);
    g_virtualPortsOut.push_back(vp);
}

void ResetSafeVirtualPortsOutIndex(int i) {
    std::shared_ptr<VirtualMidiPort> vp;
    {
        std::lock_guard<std::mutex> lock(g_midiPtrMutex);
        if (i >= 0 && i < (int)g_virtualPortsOut.size()) {
            vp = std::move(g_virtualPortsOut[i]);
            g_virtualPortsOut[i] = nullptr;
        }
    }
}

// Serial write mutex + port running-status
// g_lastSentPort == 0 means "no port selected yet" (force prefix on first write)
std::mutex g_serialWriteMutex;
std::atomic<int> g_lastSentPort{0};

// Per-port enable flags (toggled via UI, read from callbacks)
std::atomic<bool> g_portEnabled[4];
// portInEnabled: controls COM -> virtual In (serial receive callback)
std::atomic<bool> g_portInEnabled{true};
std::string g_detectedSynth = "";
bool showNames = true;
std::atomic<bool> g_isConnecting{false};
std::atomic<bool> g_autoGetSynthInfo{true};

enum class MidiServiceStatus { INITIALIZING, AVAILABLE, UNAVAILABLE, NOT_RESPONDING };
std::atomic<MidiServiceStatus> g_midiStatus{MidiServiceStatus::INITIALIZING};
std::atomic<bool> midiAvailable{false};

std::mutex g_synthNameOverlayMutex;
bool g_waitingForIdentity = false;

struct YamahaSynthIdentity {
    uint8_t fMSB, fLSB, mMSB, mLSB;
    const char* name;
};

const YamahaSynthIdentity g_xgSynthTable[] = {
    {0x41, 0x00, 0x1B, 0x00, "Yamaha MU2000"},
    {0x41, 0x1B, 0x04, 0x0A, "Yamaha MU2000EX"},
    {0x41, 0x00, 0x1C, 0x00, "Yamaha MU1000"},
    {0x41, 0x00, 0x03, 0x00, "Yamaha MU128/MU100"},
    {0x41, 0x4B, 0x01, 0x00, "Yamaha MU100 (Legacy)"},
    {0x41, 0x49, 0x01, 0x00, "Yamaha MU90"},
    {0x11, 0x03, 0x00, 0x00, "Yamaha MU80"},
    {0x11, 0x04, 0x00, 0x00, "Yamaha MU50"},
    {0x11, 0x05, 0x00, 0x00, "Yamaha MU15/MU10"},
    {0x11, 0x06, 0x00, 0x00, "Yamaha MU5"},
    {0x41, 0x5C, 0x03, 0x00, "Yamaha CS6x/CS6R"},
    {0x41, 0x10, 0x02, 0x00, "Yamaha CS1x"},
    {0x41, 0x22, 0x02, 0x00, "Yamaha CS2x"},
    {0x41, 0x11, 0x02, 0x00, "Yamaha AN1x"},
    {0x41, 0x5E, 0x03, 0x00, "Yamaha S80"},
    {0x41, 0x23, 0x04, 0x00, "Yamaha S30"},
    {0x41, 0x7F, 0x00, 0x00, "Yamaha S90/S90ES"},
    {0x41, 0x04, 0x34, 0x00, "Yamaha QY100"},
    {0x41, 0x49, 0x0B, 0x00, "Yamaha QY70"},
    {0x41, 0x4C, 0x00, 0x00, "Yamaha CBX-K1/K2"},
    {0x41, 0x27, 0x01, 0x00, "Yamaha Clavinova CLP"},
    {0x41, 0x26, 0x01, 0x00, "Yamaha Clavinova CVP"},
    {0x41, 0x49, 0x04, 0x00, "Yamaha VL70-m"},
    {0x41, 0x00, 0x01, 0x00, "Yamaha SW1000XG"}
};
const int g_xgSynthTableSize = sizeof(g_xgSynthTable) / sizeof(YamahaSynthIdentity);
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
    bool isSync;
    uint64_t seqNum;
};

std::mutex g_logMutex;
// Per-source ring buffers ΓÇö each capped at g_maxLogLines independently
std::deque<LogEntry> g_logBySource[LOG_SOURCE_COUNT];
int g_maxLogLines = 500;
uint64_t g_logSeq = 0;    // global sequence counter
size_t g_logVersion = 0;  // incremented on every push

void AddLog(LogSourceType source, const std::string& msg, bool isSync = false, const std::string& cmdName = "") {
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
    g_logBySource[idx].push_back({source, timeStr, msg, cmdName, isSync, ++g_logSeq});
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
bool IsPackagedProcess() {
    UINT32 length = 0;
    LONG result = GetCurrentPackageFullName(&length, NULL);
    return result == ERROR_INSUFFICIENT_BUFFER;
}

std::string GetIniPathDir() {
    if (IsPackagedProcess()) {
        PWSTR path = NULL;
        if (SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, NULL, &path) == S_OK) {
            std::wstring ws(path);
            CoTaskMemFree(path);
            std::string s(ws.begin(), ws.end());
            s += "\\ToHostBridge";
            CreateDirectoryA(s.c_str(), NULL);
            return s;
        }
    }
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
    WritePrivateProfileStringA("General", "AutoGetSynthInfo", g_autoGetSynthInfo.load() ? "1" : "0", ini.c_str());
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
    g_autoGetSynthInfo.store(GetPrivateProfileIntA("General", "AutoGetSynthInfo", 1, ini.c_str()) != 0);
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
    if (IsPackagedProcess()) {
        try {
            // For Store/MSIX, we use the StartupTask API instead of the Registry
            auto task = winrt::Windows::ApplicationModel::StartupTask::GetAsync(L"ToHostBridgeStartup").get();
            auto state = task.State();

            if (enable) {
                if (state != winrt::Windows::ApplicationModel::StartupTaskState::Enabled) {
                    auto result = task.RequestEnableAsync().get();
                    if (result == winrt::Windows::ApplicationModel::StartupTaskState::DisabledByUser) {
                        AddLog(LogSourceType::APP_INFO, "[Startup] Task is DISABLED in Task Manager. Please enable it manually.");
                    } else if (result == winrt::Windows::ApplicationModel::StartupTaskState::Enabled) {
                        AddLog(LogSourceType::APP_INFO, "[Startup] Task enabled successfully.");
                    } else {
                        AddLog(LogSourceType::APP_INFO, "[Startup] Could not enable task. Status: " + std::to_string((int)result));
                    }
                }
            } else {
                task.Disable();
                AddLog(LogSourceType::APP_INFO, "[Startup] Task disabled.");
            }
        } catch (...) {
            // Log or ignore WinRT errors
            AddLog(LogSourceType::APP_INFO, "Error setting Store startup task.");
        }
        return;
    }

    HKEY hKey;
    if (RegOpenKeyExA(HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS) {
        if (enable) {
            char path[MAX_PATH];
            GetModuleFileNameA(NULL, path, MAX_PATH);
            RegSetValueExA(hKey, "ToHostBridge", 0, REG_SZ, (const BYTE*)path, (DWORD)(strlen(path) + 1));
        } else {
            RegDeleteValueA(hKey, "ToHostBridge");
        }
        RegCloseKey(hKey);
    }
}

// Main code explicitly for Windows Subsystem
bool IsPackaged() {
    UINT32 length = 0;
    LONG result = GetCurrentPackageFullName(&length, nullptr);
    return result != APPMODEL_ERROR_NO_PACKAGE;
}

winrt::fire_and_forget CheckForUpdatesAsync(HWND hwnd) {
    if (g_mockUpdate) {
        // Sleep a bit to simulate network delay
        std::this_thread::sleep_for(std::chrono::seconds(2));
        g_updateAvailable = true;
        co_return;
    }

    if (!IsPackaged()) co_return;

    try {
        using namespace winrt::Windows::Services::Store;
        StoreContext context = StoreContext::GetDefault();
        
        // Associate the StoreContext with our app window (required for desktop apps)
        auto initializeWithWindow = context.as<IInitializeWithWindow>();
        initializeWithWindow->Initialize(hwnd);

        auto updates = co_await context.GetAppAndOptionalStorePackageUpdatesAsync();
        if (updates.Size() > 0) {
            g_availableUpdates = updates;
            g_updateAvailable = true;
        }
    } catch (...) {
        // Silently fail to avoid crashing if Store service is unavailable or throwing
    }
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nCmdShow)
{
    // Parse command line for --mock-update
    int argc;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (argv) {
        for (int i = 0; i < argc; i++) {
            if (wcscmp(argv[i], L"--mock-update") == 0) {
                g_mockUpdate = true;
            }
        }
        LocalFree(argv);
    }
    // 0. Startup Virtual Port Cleanup (To handle prior crashes)
    VirtualMidiPort::RemoveStalePorts();
    

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
    
    // Asynchronous MIDI Service Initialization with 5-second timeout
    std::thread midiInitThread([]() {
        Microsoft::Windows::Devices::Midi2::Initialization::MidiDesktopAppSdkInitializer midiInit;
        bool sdkAvailable = midiInit.InitializeSdkRuntime() && midiInit.EnsureServiceAvailable();
        
        if (sdkAvailable) {
            midiAvailable.store(true);
            g_midiStatus.store(MidiServiceStatus::AVAILABLE);
        } else {
            midiAvailable.store(false);
            g_midiStatus.store(MidiServiceStatus::UNAVAILABLE);
        }
    });

    // Detach and wait for status change with timeout
    std::thread([midiInitThread = std::move(midiInitThread)]() mutable {
        auto start = std::chrono::steady_clock::now();
        while (g_midiStatus.load() == MidiServiceStatus::INITIALIZING) {
            if (std::chrono::duration_cast<std::chrono::seconds>(std::chrono::steady_clock::now() - start).count() >= 5) {
                g_midiStatus.store(MidiServiceStatus::NOT_RESPONDING);
                midiAvailable.store(false);
                break;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(100));
        }
        if (midiInitThread.joinable()) midiInitThread.detach(); // We let the init thread finish or hang on its own
    }).detach();


    ImGui_ImplWin32_EnableDpiAwareness();
    float main_scale = ImGui_ImplWin32_GetDpiScaleForMonitor(::MonitorFromPoint(POINT{ 0, 0 }, MONITOR_DEFAULTTOPRIMARY));

    WNDCLASSEXW wc = { sizeof(wc), CS_CLASSDC, WndProc, 0L, 0L, GetModuleHandle(nullptr), LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1)), nullptr, nullptr, nullptr, L"ToHostBridge", LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1)) };
    ::RegisterClassExW(&wc);
    // Adjusted window dimensions for bold headers and better spacing
    HWND hwnd = ::CreateWindowW(wc.lpszClassName, L"ToHost Bridge v1.2.2", WS_OVERLAPPEDWINDOW, 100, 100, (int)(530 * main_scale), (int)(430 * main_scale), nullptr, nullptr, wc.hInstance, nullptr);

    // Initial check for Microsoft Store updates
    CheckForUpdatesAsync(hwnd);

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
    if (startToTray || startMinimized) {
        ::Shell_NotifyIconW(NIM_ADD, &g_nid);
        ::ShowWindow(hwnd, SW_HIDE);
    } else {
        ::ShowWindow(hwnd, SW_SHOWDEFAULT);
    }
    ::UpdateWindow(hwnd);

    if (comPortsVec.size() > 0 && selectedComPort >= (int)comPortsVec.size()) selectedComPort = 0;

    bool isConnected = false;
    std::atomic<bool> connectionLost{false};
    std::string activeComName = savedComName;
    std::string connectionStatus = "";

    // Startup: if saved port not present yet, mark as waiting
    if (autoStartVirtualMidi && !savedComName.empty() && !foundSavedPort) {
        connectionLost.store(true);
        activeComName = savedComName;
        connectionStatus = "Waiting for " + savedComName + "...";
    }

    // Filters
    bool filterSys = true;
    bool filterIn = true;
    bool filterOut[4] = {true, true, true, true};
    bool filterSync = true;
    bool stopScroll = false;
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
        // Each source is individually capped ΓÇö no merged trim needed
        cachedSnapshot = std::move(merged);
        snapshotLogVersion = g_logVersion;
        lastFilterState = {filterSys, filterIn, filterSync, {filterOut[0],filterOut[1],filterOut[2],filterOut[3]}};
    };

    // Pre-declare lambda for Connect
    auto ConnectPort = [&]() -> bool {
        if (isConnected || g_midiStatus.load() != MidiServiceStatus::AVAILABLE) return false;
        if (comPorts.empty()) return false;
        std::string baseNameStr(baseNameBuf);
        ResetSafeSerialPort();

        std::string comName = comPorts[selectedComPort];
        activeComName = comName;

        // Restart virtual MIDI ports if they were stopped
        if (GetSafeVirtualPortsOut().empty() && !g_virtualPortIn) {
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
                        AddLog(srcType, "[App->COM] P" + std::to_string(i) + ": " + BytesToHex(data), isSync, GetMidiCommandName(data));
                    });
                AddSafeVirtualPortOut(std::move(vport));
                ::Sleep(500); // Wait for Windows MIDI Services to commit each endpoint before registering the next.
                              // 150ms was too short — the async registration raced and ports appeared out of order.
            }
            if (g_portInEnabled.load(std::memory_order_relaxed)) {
                std::wstring inPortName = wBaseName + L" In";
                SetSafeVirtualPortIn(std::make_shared<VirtualMidiPort>(inPortName, nullptr));
            }
            AddLog(LogSourceType::APP_INFO, "Virtual MIDI ports started.");
        }
        
        struct MidiParser {
            uint8_t runningStatus = 0;
            std::vector<uint8_t> buffer;
            bool isSysEx = false;
            std::vector<uint8_t> sysexBuffer;
        };
        auto parser = std::make_shared<MidiParser>();
        
        SetSafeSerialPort(std::make_shared<SerialPort>(comName, [colIn, parser](const std::vector<uint8_t>& data) {
            g_bytesReceived.fetch_add(data.size(), std::memory_order_relaxed);
            for (uint8_t b : data) {
                // Realtime bytes (High Priority)
                if (b >= 0xF8) {
                    std::vector<uint8_t> rt = { b };
                    if (g_transmitSync && g_portInEnabled.load(std::memory_order_relaxed)) {
                        try {
                            if (auto vpi = GetSafeVirtualPortIn()) vpi->SendMidi(rt);
                        } catch (...) {}
                    }
                    AddLog(LogSourceType::COM_IN_SYNC, "[COM->App] " + BytesToHex(rt), true, GetMidiCommandName(rt));
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
                        if (g_portInEnabled.load(std::memory_order_relaxed)) {
                            try {
                                if (auto vpiLocal = GetSafeVirtualPortIn()) vpiLocal->SendMidi(parser->sysexBuffer);
                            } catch (...) {}
                        }
                        AddLog(LogSourceType::COM_IN, "[COM->App] SysEx: " + BytesToHex(parser->sysexBuffer), false, "System Exclusive");
                        
                        // Identity Reply: F0 7E [dev] 06 02 43 00 [fMSB] [fLSB] [mMSB] [mLSB] ... F7
                        if (parser->sysexBuffer.size() >= 12 && parser->sysexBuffer[1] == 0x7E && parser->sysexBuffer[3] == 0x06 && parser->sysexBuffer[4] == 0x02 && parser->sysexBuffer[5] == 0x43) {
                            uint8_t fMSB = parser->sysexBuffer[7];
                            uint8_t fLSB = parser->sysexBuffer[8];
                            uint8_t mMSB = parser->sysexBuffer[9];
                            uint8_t mLSB = parser->sysexBuffer[10];

                            std::string foundName = "Unknown (" + BytesToHex({fMSB, fLSB, mMSB, mLSB}) + ")";
                            for (int i = 0; i < g_xgSynthTableSize; ++i) {
                                if (g_xgSynthTable[i].fMSB == fMSB && g_xgSynthTable[i].fLSB == fLSB &&
                                    g_xgSynthTable[i].mMSB == mMSB && g_xgSynthTable[i].mLSB == mLSB) {
                                    foundName = g_xgSynthTable[i].name;
                                    break;
                                }
                            }
                            {
                                std::lock_guard<std::mutex> lock(g_synthNameOverlayMutex);
                                g_detectedSynth = foundName;
                                g_waitingForIdentity = false;
                            }
                        }
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
                    if (g_portInEnabled.load(std::memory_order_relaxed)) {
                        try {
                            if (auto vpiLocal = GetSafeVirtualPortIn()) vpiLocal->SendMidi(parser->buffer);
                        } catch (...) {}
                    }
                    AddLog(LogSourceType::COM_IN, "[COM->App] " + BytesToHex(parser->buffer), false, GetMidiCommandName(parser->buffer));
                    parser->buffer.clear();
                }
            }
        }));

        // Reset port running-status so the receiver re-syncs on (re)connect
        g_lastSentPort.store(0, std::memory_order_relaxed);

        std::stringstream connMsg;
        if (GetSafeSerialPort() && GetSafeSerialPort()->IsOpen()) {
            connMsg << "Opened " << comName << " successfully.\n";
            connMsg << "Virtual ports:\n";
            for (int i = 1; i <= 4; ++i) {
                if (g_portEnabled[i-1].load(std::memory_order_relaxed))
                    connMsg << "  " << baseNameStr << " Out " << i << "\n";
            }
            if (g_portInEnabled.load(std::memory_order_relaxed))
                connMsg << "  " << baseNameStr << " In\n";
            AddLog(LogSourceType::APP_INFO, "Connected to " + comName + " (" + baseNameStr + ")");
            isConnected = true;
            connectionLost = false;

            // Auto-trigger Synth Identity Discovery on connect (if enabled)
            if (GetSafeSerialPort() && GetSafeSerialPort()->IsOpen() && g_autoGetSynthInfo.load()) {
                // Auto-trigger Synth Identity Discovery on connect (Immediate)
                std::vector<uint8_t> stdId = { 0xF0, 0x7E, 0x7F, 0x06, 0x01, 0xF7 };
                if(auto sp2 = GetSafeSerialPort()) sp2->Write(stdId);
                AddLog(LogSourceType::APP_INFO, "[System] " + BytesToHex(stdId), false, "(Identity Request)");
                
                {
                    std::lock_guard<std::mutex> lock(g_synthNameOverlayMutex);
                    g_detectedSynth = "";
                    g_waitingForIdentity = true;
                }
                // Timeout to clear "Querying..." if no reply
                std::thread([&]() {
                    Sleep(2000);
                    if (g_waitingForIdentity) {
                        std::lock_guard<std::mutex> lock(g_synthNameOverlayMutex);
                        if (g_detectedSynth == "Querying...") g_detectedSynth = "";
                        g_waitingForIdentity = false;
                    }
                }).detach();
            }

            // Persist the newly-connected COM port name to settings
            SaveSettings(comName, baseNameStr, autoStartVirtualMidi, startWithWindows, autoReconnect, startToTray, colSys, colIn, colOut, stayOnTop, startMinimized, lightUI);
            connectionStatus = connMsg.str();
            return true;
        } else {
            connMsg << "Failed to open " << comName << ". Is it in use?";
            AddLog(LogSourceType::APP_INFO, "Failed to open " + comName);
            ResetSafeSerialPort();
            connectionStatus = connMsg.str();
            return false;
        }
    };

    bool firstFrame = true;
    bool done = false;
    while (!done && !g_readyToExit.load())
    {
        MSG msg;
        while (::PeekMessage(&msg, nullptr, 0U, 0U, PM_REMOVE))
        {
            ::TranslateMessage(&msg);
            ::DispatchMessage(&msg);
            if (msg.message == WM_QUIT)
                done = true;
        }
        if (done || g_readyToExit.load()) break;

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
            // Removed captured midiOk - we'll check it live in the thread

            std::thread([&, savedNameCopy, baseNameCopy, autoStart, foundPort,
                         colOutCopy, colSysCopy, colInCopy, portsVecCopy]() mutable {
                // Wait for MIDI status to leave INITIALIZING
                while (g_midiStatus.load() == MidiServiceStatus::INITIALIZING) { ::Sleep(100); }
                bool midiOk = (g_midiStatus.load() == MidiServiceStatus::AVAILABLE);

                // ΓöÇΓöÇ Create virtual MIDI ports (if auto-start is enabled and MIDI is AVAILABLE) ΓöÇΓöÇ
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
                                AddLog(srcType, "[App->COM] P" + std::to_string(i) + ": " + BytesToHex(data), isSync, GetMidiCommandName(data));
                            });
                        AddSafeVirtualPortOut(std::move(vport));
                    }
                    if (g_portInEnabled.load(std::memory_order_relaxed)) {
                        std::wstring inPortName = wBaseName + L" In";
                        SetSafeVirtualPortIn(std::make_shared<VirtualMidiPort>(inPortName, nullptr));
                    }
                    AddLog(LogSourceType::APP_INFO, "Virtual MIDI ports started.");
                }

                // ΓöÇΓöÇ Auto-connect to COM port ΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇ
                if (autoStart && !portsVecCopy.empty()) {
                    if (!savedNameCopy.empty()) {
                        for (int i = 0; i < (int)portsVecCopy.size(); ++i) {
                            if (portsVecCopy[i] == savedNameCopy) {
                                selectedComPort = i;
                                ConnectPort();
                                return;
                            }
                        }
                        // Port not available yet ΓÇö connectionLost already set in main
                    } else {
                        ConnectPort(); // no saved name, connect to first
                    }
                }
                g_autoStartAttempted.store(true); // Ensure flag is set even if not started
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

        // ΓöÇΓöÇ Auto-Reconnect Polling + Bandwidth (every 1 second) ΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇ
        static auto lastCheck = std::chrono::steady_clock::now();
        static uint64_t prevSent = 0, prevRecv = 0;
        auto nowCheck = std::chrono::steady_clock::now();
        if (std::chrono::duration_cast<std::chrono::milliseconds>(nowCheck - lastCheck).count() > 1000) {
            lastCheck = nowCheck;

            // Bandwidth: 38400 baud, 8N1 ΓåÆ 10 bits/byte ΓåÆ 3840 bytes/sec max
            uint64_t curSent = g_bytesSent.load(std::memory_order_relaxed);
            uint64_t curRecv = g_bytesReceived.load(std::memory_order_relaxed);
            double sentRate = (double)(curSent - prevSent);
            double recvRate = (double)(curRecv - prevRecv);
            double maxRate = (sentRate > recvRate) ? sentRate : recvRate;
            // Calibration: While 38400 baud is 3840 B/s, practical synth processing 
            // and multiplexing overhead make the "struggle" point much lower.
            double saturationPoint = 1500.0; 
            double chargeRaw = (maxRate / saturationPoint) * 100.0;
            prevSent = curSent; prevRecv = curRecv;
            
            float newCharge = isConnected ? ((float)chargeRaw > 100.0f ? 100.0f : (float)chargeRaw) : 0.0f;
            // High alpha (0.7) for fast capture, and immediate reset on zero to avoid the "laggy" tail
            if (newCharge == 0.0f) g_chargePercent = 0.0f;
            else g_chargePercent = (g_chargePercent * 0.3f) + (newCharge * 0.7f);
            // Update Win32 title bar
            std::wstring wTitle = L"ToHost Bridge v1.0";

            if (!activeComName.empty()) {
                char dummyPath[1000];
                DWORD res = QueryDosDeviceA(activeComName.c_str(), dummyPath, sizeof(dummyPath));
                bool portExists = (res != 0);

                // Detect sudden disconnect
                if (isConnected && !portExists) {
                    isConnected = false;
                    connectionLost.store(true);
                    ResetSafeSerialPort();   // close the dead handle cleanly
                    AddLog(LogSourceType::APP_INFO, "COM port lost: " + activeComName);
                    connectionStatus = "COM port disconnected!";
                }

                // Auto-reconnect when port reappears (triggered by either drop recovery or initial auto-start wait)
                if (!isConnected && connectionLost.load() && portExists && (autoReconnect || autoStartVirtualMidi)) {
                    // Rebuild port list so dropdown + index are correct
                    comPortsVec = GetAvailableComPorts();
                    comPorts.clear();
                    for (const auto& p : comPortsVec) comPorts.push_back(p.c_str());
                    
                    // Match by name
                    for (int i = 0; i < (int)comPortsVec.size(); ++i) {
                        if (comPortsVec[i] == activeComName) { selectedComPort = i; break; }
                    }
                    
                    if (ConnectPort()) {
                        connectionLost.store(false);
                    } else {
                        // If it fails (e.g. port just appeared but not yet ready),
                        // we keep connectionLost=true so we retry next second.
                        connectionStatus = "Waiting for " + activeComName + " (retrying)...";
                    }
                }
            }
        }

        ImGui_ImplDX11_NewFrame();
        ImGui_ImplWin32_NewFrame();
        ImGui::NewFrame();

        // Late MIDI Init Watcher: If service became available and autostart is on, trigger it
        if (!g_autoStartAttempted.load() && autoStartVirtualMidi && g_midiStatus.load() == MidiServiceStatus::AVAILABLE && GetSafeVirtualPortsOut().empty() && !g_virtualPortIn) {
            g_autoStartAttempted.store(true);
            std::string baseNameStr(baseNameBuf);
            std::thread([&, baseNameStr]() {
                AddLog(LogSourceType::APP_INFO, "MIDI Service became available. Triggering auto-start...");
                // Note: ConnectPort() called by the main loop will handle creating 
                // the virtual ports once it sees connectionLost set to true below.
                if (autoStartVirtualMidi) {
                    connectionLost.store(true);
                }
                g_autoStartAttempted.store(true);
            }).detach();
        }

        if (g_isShuttingDown.load()) {
            // Full-screen "Closing" overlay
            ImGui::SetNextWindowPos(ImVec2(0, 0));
            ImGui::SetNextWindowSize(io.DisplaySize);
            ImGui::Begin("ShutdownOverlay", nullptr, ImGuiWindowFlags_NoDecoration | ImGuiWindowFlags_NoInputs | ImGuiWindowFlags_NoNav);
            
            float centerX = io.DisplaySize.x * 0.5f;
            float centerY = io.DisplaySize.y * 0.5f;
            
            // Move everything higher (Start higher than center)
            float currentY = centerY - 100.0f;

            const char* msg = "Finalizing ToHost Bridge...";
            ImVec2 txtSz = ImGui::CalcTextSize(msg);
            ImGui::SetCursorPos(ImVec2(centerX - txtSz.x * 0.5f, currentY));
            ImGui::TextColored(ImVec4(1.0f, 0.4f, 0.4f, 1.0f), "%s", msg);
            currentY += 40;

            if (!g_shutdownTasks.empty()) {
                std::lock_guard<std::mutex> lock(g_shutdownMutex);
                float listWidth = 280.0f * (io.DisplaySize.x / 530.0f); 
                float startX = centerX - listWidth * 0.5f;

                for (const auto& task : g_shutdownTasks) {
                    ImVec4 color = (task.status == "OK") ? ImVec4(0.4f, 1.0f, 0.4f, 1.0f) : (task.status == "..." ? ImVec4(1.0f, 1.0f, 0.0f, 1.0f) : ImVec4(0.7f, 0.7f, 0.7f, 1.0f));
                    
                    // Task Name (Left-column)
                    ImGui::SetCursorPos(ImVec2(startX, currentY));
                    ImGui::Text("%s", task.name.c_str());
                    
                    // Status (Right-column, attracted to the left of its column)
                    ImGui::SetCursorPos(ImVec2(startX + 190.0f, currentY));
                    ImGui::TextColored(color, "%s", task.status.c_str());
                    
                    currentY += 25;
                }
            }

            if (!g_cleanupStarted.load()) {
                g_cleanupStarted.store(true);
                std::thread([&]() {
                    // Initialize tasks
                    {
                        std::lock_guard<std::mutex> lock(g_shutdownMutex);
                        g_shutdownTasks.push_back({"Silence hardware", "..."});
                        for (int i = 0; i < (int)GetSafeVirtualPortsOut().size(); i++) 
                            g_shutdownTasks.push_back({"Close Virtual Out " + std::to_string(i+1), "Pending"});
                        if (auto vpi = GetSafeVirtualPortIn()) g_shutdownTasks.push_back({"Close Virtual In", "Pending"});
                        g_shutdownTasks.push_back({"Finalize connection", "Pending"});
                    }

                    // 1. Silence serial synths
                    if (isConnected && GetSafeSerialPort() && GetSafeSerialPort()->IsOpen()) {
                        for (int p = 1; p <= 4; ++p) {
                            std::vector<uint8_t> quiet;
                            for (int ch = 0; ch < 16; ++ch) quiet.insert(quiet.end(), { (uint8_t)(0xB0 | ch), 0x7B, 0x00 });
                            quiet.push_back(0xFC); // MIDI Stop
                            SendToSerial(p, quiet);
                        }
                    }
                    { std::lock_guard<std::mutex> lock(g_shutdownMutex); g_shutdownTasks[0].status = "OK"; }
                    
                    // 2. Send MIDI Stop to Virtual Input (if exists)
                    if (auto vpi = GetSafeVirtualPortIn()) {
                        vpi->SendMidi({ 0xFC });
                    }

                    // 3. Parallel-ish Close with 100ms stagger
                    std::vector<std::thread> cleanThreads;
                    
                    // Launch Output Port Closures
                    for (int i = 0; i < (int)GetSafeVirtualPortsOut().size(); ++i) {
                        int taskIdx = i + 1;
                        cleanThreads.emplace_back([i, taskIdx]() {
                            { std::lock_guard<std::mutex> lock(g_shutdownMutex); g_shutdownTasks[taskIdx].status = "..."; }
                            ResetSafeVirtualPortsOutIndex(i);
                            { std::lock_guard<std::mutex> lock(g_shutdownMutex); g_shutdownTasks[taskIdx].status = "OK"; }
                        });
                        ::Sleep(1000); // 1s stagger between STARTING each port closure
                    }
                    
                    // Launch Input Port Closure
                    if (auto vpi = GetSafeVirtualPortIn()) {
                        int taskIdx = (int)g_shutdownTasks.size() - 2;
                        cleanThreads.emplace_back([taskIdx]() {
                            { std::lock_guard<std::mutex> lock(g_shutdownMutex); g_shutdownTasks[taskIdx].status = "..."; }
                            ResetSafeVirtualPortIn();
                            { std::lock_guard<std::mutex> lock(g_shutdownMutex); g_shutdownTasks[taskIdx].status = "OK"; }
                        });
                        ::Sleep(1000);
                    }
                    
                    // Wait for all port driver cleanup to finish (the "Barrier")
                    for (auto& t : cleanThreads) {
                        if (t.joinable()) t.join();
                    }
                    ClearSafeVirtualPortsOut(); 
                    
                    // 4. Finalize
                    int finalIdx = (int)g_shutdownTasks.size() - 1;
                    { std::lock_guard<std::mutex> lock(g_shutdownMutex); g_shutdownTasks[finalIdx].status = "..."; }
                    ResetSafeSerialPort();
                    { std::lock_guard<std::mutex> lock(g_shutdownMutex); g_shutdownTasks[finalIdx].status = "OK"; }
                    
                    ::Sleep(200); 
                    g_readyToExit.store(true);
                }).detach();
            }
        } else {
            ImGui::SetNextWindowPos(ImVec2(0, 0));
            ImGui::SetNextWindowSize(io.DisplaySize);
            ImGui::Begin("Main", nullptr, ImGuiWindowFlags_NoTitleBar | ImGuiWindowFlags_NoResize | ImGuiWindowFlags_NoMove);

            // Microsoft Store Update Popup
            static bool updatePopupOpened = false;
            if (g_updateAvailable && !updatePopupOpened) {
                ImGui::OpenPopup("Update Available");
                updatePopupOpened = true;
            }

            if (ImGui::BeginPopupModal("Update Available", nullptr, ImGuiWindowFlags_AlwaysAutoResize)) {
                ImGui::Spacing();
                ImGui::Text("A new update is available in the Microsoft Store!");
                if (g_mockUpdate) ImGui::TextColored(ImVec4(1.0f, 1.0f, 0.0f, 1.0f), "[DEBUG] Mock Update Mode Active");
                ImGui::Spacing();
                ImGui::Text("Would you like to download and install it now?");
                ImGui::TextDisabled("The application will close to begin the update process.");
                ImGui::Spacing();
                ImGui::Separator();
                ImGui::Spacing();

                if (ImGui::Button("Update Now", ImVec2(140, 35))) {
                    if (g_mockUpdate) {
                        ::MessageBoxA(NULL, "Mock update triggered successfully!", "Debug", MB_OK);
                    } else {
                        // Standard fire-and-forget for the update request
                        []() -> winrt::fire_and_forget {
                            try {
                                auto context = winrt::Windows::Services::Store::StoreContext::GetDefault();
                                // HWND was already initialized in CheckForUpdatesAsync
                                co_await context.RequestDownloadAndInstallStorePackageUpdatesAsync(g_availableUpdates);
                            } catch (...) {}
                        }();
                    }
                    ImGui::CloseCurrentPopup();
                }
                ImGui::SameLine();
                if (ImGui::Button("Later", ImVec2(140, 35))) {
                    ImGui::CloseCurrentPopup();
                }
                ImGui::EndPopup();
            }

        // Header Error Messaging for MIDI Services (Moved to Connection Tab)
  
        // Top-right device overlay
        float preTabY = ImGui::GetCursorPosY();
        if (isConnected && !g_detectedSynth.empty()) {
            std::lock_guard<std::mutex> lock(g_synthNameOverlayMutex);
            std::string disp = g_detectedSynth;
            float textWidth = ImGui::CalcTextSize(disp.c_str()).x;
            ImGui::SetCursorPosX(ImGui::GetWindowContentRegionMax().x - textWidth - 5.0f);
            ImGui::TextColored(lightUI ? ImVec4(0.0f, 0.0f, 0.0f, 1.0f) : ImVec4(0.0f, 1.0f, 1.0f, 1.0f), "%s", disp.c_str());
            ImGui::SetCursorPosY(preTabY);
        }

        if (ImGui::BeginTabBar("Tabs")) {

            // ΓöÇΓöÇ Connection Tab ΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇ
            if (ImGui::BeginTabItem("Connection")) {
                ImGui::Spacing();
                bool portsRunning = !GetSafeVirtualPortsOut().empty();
                
                ImGui::BeginDisabled(isConnected);
                if (ImGui::Button("Refresh COM Ports")) {
                    comPortsVec = GetAvailableComPorts();
                    comPorts.clear();
                    for (const auto& p : comPortsVec) comPorts.push_back(p.c_str());
                    selectedComPort = 0;
                }
                ImGui::EndDisabled();
                ImGui::SameLine();
                
                ImGui::BeginDisabled(portsRunning || isConnected);
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
                
                ImGui::EndDisabled();

                ImGui::Spacing();
                
                bool disableComSelect = isConnected || comPorts.empty();
                if (disableComSelect) ImGui::BeginDisabled();
                ImGui::SetNextItemWidth(120);
                ImGui::Combo("COM Port", &selectedComPort, comPorts.empty() ? nullptr : comPorts.data(), (int)comPorts.size());
                if (disableComSelect) ImGui::EndDisabled();

                ImGui::Spacing();
                bool disableConnect = isConnected || g_isConnecting.load() || g_midiStatus.load() != MidiServiceStatus::AVAILABLE;
                ImGui::BeginDisabled(disableConnect);
                if (ImGui::Button(isConnected ? "Connected" : "Connect", ImVec2(120, 30)) && !comPorts.empty()) {
                    if (!isConnected && !g_isConnecting.load()) {
                        if (g_midiStatus.load() != MidiServiceStatus::AVAILABLE) {
                            connectionStatus = "Cannot connect: Windows MIDI Service is not responding.";
                            AddLog(LogSourceType::APP_INFO, "Connection aborted: MIDI Service not available.");
                        } else {
                            g_isConnecting.store(true);
                            connectionStatus = "Connecting...";
                            std::thread([&]() {
                                std::atomic<bool> portOpDone{false};
                                std::thread worker([&]() {
                                    ConnectPort();
                                    portOpDone.store(true);
                                });
                                
                                auto start = std::chrono::steady_clock::now();
                                while (!portOpDone.load()) {
                                    if (std::chrono::duration_cast<std::chrono::seconds>(std::chrono::steady_clock::now() - start).count() >= 10) {
                                        // TIMEOUT - assume MIDI service is hung/broken
                                        g_midiStatus.store(MidiServiceStatus::NOT_RESPONDING);
                                        connectionStatus = "Cannot connect: Windows MIDI Service is not responding.";
                                        AddLog(LogSourceType::APP_INFO, "Connection timed out after 10s (MIDI Service hang).");
                                        if (worker.joinable()) worker.detach();
                                        g_isConnecting.store(false);
                                        return;
                                    }
                                    std::this_thread::sleep_for(std::chrono::milliseconds(100));
                                }
                                if (worker.joinable()) worker.join();
                                g_isConnecting.store(false);
                            }).detach();
                        }
                    }
                }
                ImGui::EndDisabled();
                
                // Disconnect button
                if (isConnected) {
                    ImGui::SameLine();
                    if (ImGui::Button("Disconnect", ImVec2(120, 30))) {
                        ResetSafeSerialPort();
                        g_lastSentPort.store(0, std::memory_order_relaxed);
                        isConnected = false;
                        connectionLost.store(false);
                        connectionStatus = "Disconnected (virtual ports still active).";
                        AddLog(LogSourceType::APP_INFO, "Disconnected from COM port.");
                    }
                }

                ImGui::Spacing();
                ImGui::Separator();
                ImGui::Spacing();
                
                // Moved MIDI Service Errors to right under the separator
                MidiServiceStatus currentStatus = g_midiStatus.load();
                if (currentStatus == MidiServiceStatus::INITIALIZING) {
                    ImGui::TextWrapped("Checking MIDI Service...");
                } else if (currentStatus == MidiServiceStatus::NOT_RESPONDING || currentStatus == MidiServiceStatus::UNAVAILABLE) {
                    ImGui::PushStyleColor(ImGuiCol_ChildBg, ImVec4(0.4f, 0.1f, 0.1f, 1.0f));
                    ImGui::BeginChild("MidiError", ImVec2(0, ImGui::GetContentRegionAvail().y), true);
                    if (currentStatus == MidiServiceStatus::NOT_RESPONDING) {
                        ImGui::TextColored(ImVec4(1.0f, 0.8f, 0.0f, 1.0f), "CRITICAL ERROR: Windows MIDI Service is not responding.");
                        ImGui::TextWrapped("The Windows MIDI Services process (MidiSrv.exe) appears to be hung. This application will not be able to create virtual MIDI ports.");
                        ImGui::TextWrapped("RECOMMENDATION: Please restart your computer or stop 'MidiSrv.exe' in Task Manager.");
                        if (ImGui::Button("Open Task Manager")) ShellExecuteA(NULL, "open", "taskmgr.exe", NULL, NULL, SW_SHOWNORMAL);
                    } else {
                        ImGui::TextColored(ImVec4(1.0f, 0.2f, 0.2f, 1.0f), "ERROR: Windows MIDI Services Runtime is missing.");
                        ImGui::TextWrapped("ToHost Bridge requires the Windows MIDI Services (MIDI 2.0) runtime for virtual ports and reliable MIDI communication.");
                        if (ImGui::Button("Download Windows MIDI Services (x64 Installer)")) {
                            ShellExecuteA(NULL, "open", "https://microsoft.github.io/MIDI/get-latest/#:~:text=Download%20Latest%20x64%20Installer", NULL, NULL, SW_SHOWNORMAL);
                        }
                    }
                    ImGui::EndChild();
                    ImGui::PopStyleColor();
                } else if (g_isConnecting.load()) {
                    ImGui::TextWrapped("Connecting...");
                } else if (connectionLost.load()) {
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
                        AddLog(LogSourceType::APP_INFO, "Sent Panic (All Sound/Notes Off, Reset CC) to all ports/channels.");
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
                        AddLog(LogSourceType::APP_INFO, "Sent XG System On Reset to all active ports.");
                    }
                    ImGui::PopStyleColor(4);
                }

                // Stop Virtual Ports button ΓÇö show when not connected but ports are alive
                if (!isConnected && !GetSafeVirtualPortsOut().empty()) {
                    ImGui::Spacing();
                    ImGui::BeginDisabled(g_isConnecting.load());
                    ImGui::PushStyleColor(ImGuiCol_Button,        ImVec4(0.6f, 0.3f, 0.0f, 1.0f));
                    ImGui::PushStyleColor(ImGuiCol_ButtonHovered, ImVec4(0.8f, 0.4f, 0.0f, 1.0f));
                    ImGui::PushStyleColor(ImGuiCol_ButtonActive,  ImVec4(0.5f, 0.2f, 0.0f, 1.0f));
                    if (ImGui::Button("Stop Virtual Ports", ImVec2(-1, 28))) {
                        g_isConnecting.store(true);
                        std::thread([&]() {
                            ClearSafeVirtualPortsOut();
                            ResetSafeVirtualPortIn();
                            connectionStatus = "Disconnected.";
                            AddLog(LogSourceType::APP_INFO, "Virtual MIDI ports stopped.");
                            g_isConnecting.store(false);
                        }).detach();
                    }
                    ImGui::PopStyleColor(3);
                    ImGui::EndDisabled();
                }

                // MIDI Service error block moved up

                ImGui::EndTabItem();
            }

            // ΓöÇΓöÇ Debug View Tab ΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇ
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
                    // Just toggled ON ΓÇö capture frozen snapshot now
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
                    // Stop scroll turned OFF ΓÇö resume live view
                    if (stopScrollWasOn) {
                        stopScrollWasOn = false;
                        snapshotLogVersion = SIZE_MAX; // force live rebuild
                    }

                    if (filterChanged || logChanged) {
                        RebuildSnapshot();
                    }

                    ImGui::BeginChild("LogRegion", ImVec2(0, -ImGui::GetTextLineHeightWithSpacing() * 1.8f), true, ImGuiWindowFlags_HorizontalScrollbar);
                    for (const auto& line : cachedSnapshot) {
                        ImVec4 drawCol;
                        switch (line.source) {
                            case LogSourceType::APP_INFO:    drawCol = colSys; break;
                            case LogSourceType::COM_IN:      
                            case LogSourceType::COM_IN_SYNC: drawCol = colIn; break;
                            case LogSourceType::APP_OUT_1:   drawCol = colOut[0]; break;
                            case LogSourceType::APP_OUT_2:   drawCol = colOut[1]; break;
                            case LogSourceType::APP_OUT_3:   drawCol = colOut[2]; break;
                            case LogSourceType::APP_OUT_4:   drawCol = colOut[3]; break;
                            default: drawCol = ImVec4(1,1,1,1); break;
                        }

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
            ImGui::SameLine();
            if (ImGui::Button("Get synth info")) {
                if (GetSafeSerialPort() && GetSafeSerialPort()->IsOpen()) {
                    std::vector<uint8_t> stdId = { 0xF0, 0x7E, 0x7F, 0x06, 0x01, 0xF7 };
                    if(auto sp2 = GetSafeSerialPort()) sp2->Write(stdId);
                    AddLog(LogSourceType::APP_INFO, "[System] " + BytesToHex(stdId), false, "(Identity Request)");
                    {
                        std::lock_guard<std::mutex> lock(g_synthNameOverlayMutex);
                        g_detectedSynth = "";
                        g_waitingForIdentity = true;
                    }
                    std::thread([&]() {
                        Sleep(2000);
                        if (g_waitingForIdentity) {
                            std::lock_guard<std::mutex> lock(g_synthNameOverlayMutex);
                            if (g_detectedSynth == "") g_detectedSynth = "";
                            g_waitingForIdentity = false;
                        }
                    }).detach();
                } else {
                    std::lock_guard<std::mutex> lock(g_synthNameOverlayMutex);
                    g_detectedSynth = "Port Closed";
                }
            }
                ImGui::EndTabItem();
            }

            // ΓöÇΓöÇ Settings Tab ΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇ
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
                        colSys    = ImVec4(0.20f, 0.20f, 0.20f, 1.00f); // Darker gray
                        colIn     = ImVec4(0.00f, 0.30f, 0.60f, 1.00f); // Deep blue
                        colOut[0] = ImVec4(0.80f, 0.00f, 0.00f, 1.00f); // Bold red
                        colOut[1] = ImVec4(0.65f, 0.30f, 0.00f, 1.00f); // Deep orange
                        colOut[2] = ImVec4(0.00f, 0.50f, 0.00f, 1.00f); // Deep green
                        colOut[3] = ImVec4(0.50f, 0.00f, 0.50f, 1.00f); // Purple
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

                if (ImGui::Checkbox("Query device info at connection", (bool*)&g_autoGetSynthInfo)) {
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
    if (isConnected && GetSafeSerialPort() && GetSafeSerialPort()->IsOpen()) {
        for (int p = 1; p <= 4; ++p) {
            std::vector<uint8_t> quiet;
            for (int ch = 0; ch < 16; ++ch) {
                quiet.insert(quiet.end(), { (uint8_t)(0xB0 | ch), 0x7B, 0x00 });
            }
            quiet.push_back(0xFC); // MIDI Stop
            SendToSerial(p, quiet);
        }
        
        // Also send MIDI Stop to virtual Input if active
        if (auto vpi = GetSafeVirtualPortIn()) {
            vpi->SendMidi({ 0xFC });
        }
    }

    ResetSafeSerialPort();
    ClearSafeVirtualPortsOut();
    ResetSafeVirtualPortIn();

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
    case WM_CLOSE:
        if (g_isShuttingDown.load()) {
            // Second click: Hide the window while cleanup continues in background
            ::ShowWindow(hWnd, SW_HIDE);
        } else {
            g_isShuttingDown.store(true);
        }
        return 0; // Always return 0 to prevent DefWindowProc from closing the window immediately
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

// ΓöÇΓöÇ Centralized Serial Writing with Multiplexing ΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇ
void SendToSerial(int portIdx, const std::vector<uint8_t>& data) {
    auto sp = GetSafeSerialPort();
    if (!sp || !sp->IsOpen() || data.empty()) return;

    std::lock_guard<std::mutex> lock(g_serialWriteMutex);
    std::vector<uint8_t> mux;

    // Yamaha Multiplexing (Serial TO HOST protocol):
    // Prepend 0xF5 <port> if we are switching from a different virtual port.
    // 1-indexed ports: 1-4
    if (g_lastSentPort.load(std::memory_order_relaxed) != portIdx) {
        mux.push_back(0xF5);
        mux.push_back((uint8_t)portIdx);
        g_lastSentPort.store(portIdx, std::memory_order_relaxed);
        
        // Log the multiplexing byte for visibility (Port Select)
        std::string hex = "F5 " + BytesToHex({ (uint8_t)portIdx });
        std::string desc = "(Addressing Parts " + std::to_string((portIdx - 1) * 16 + 1) + "-" + std::to_string(portIdx * 16) + ")";
        AddLog(LogSourceType::APP_INFO, "[Port Select] " + hex, false, desc);
        
        sp->Write(mux);
        g_bytesSent.fetch_add(mux.size(), std::memory_order_relaxed);
    }

    sp->Write(data);
    g_bytesSent.fetch_add(data.size(), std::memory_order_relaxed);
}




