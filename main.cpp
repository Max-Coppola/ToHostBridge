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
#include <condition_variable>
#include <unordered_set>

#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.ApplicationModel.h>
#include <winrt/Windows.Services.Store.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <shobjidl.h> // For IInitializeWithWindow
#include <winmidi/init/Microsoft.Windows.Devices.Midi2.Initialization.hpp>
#include <winrt/Microsoft.Windows.Devices.Midi2.Endpoints.Virtual.h>

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

std::atomic<bool> g_isManualStopping{false};
std::vector<ShutdownTask> g_manualStopTasks;

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
// Per-port enable flags (toggled via UI, read from callbacks)
std::atomic<bool> g_portEnabled[5];
// portInEnabled: controls COM -> virtual In (serial receive callback)
std::atomic<bool> g_portInEnabled{true};
std::string g_detectedSynth = "";
bool showNames = true;
std::atomic<bool> g_isConnecting{false};
std::atomic<bool> g_autoGetSynthInfo{true};

enum class MidiServiceStatus { INITIALIZING, AVAILABLE, UNAVAILABLE, NOT_RESPONDING };
std::atomic<MidiServiceStatus> g_midiStatus{MidiServiceStatus::INITIALIZING};

// Active Device Watcher Synchronization
std::mutex g_creationMutex;
std::condition_variable g_creationCv;
std::unordered_set<std::wstring> g_arrivedDevices; // Stores DeviceId seen by watcher
winrt::Microsoft::Windows::Devices::Midi2::MidiEndpointDeviceWatcher g_watcher{ nullptr };
winrt::event_token g_watcherAddedToken;
winrt::Microsoft::Windows::Devices::Midi2::MidiSession g_sharedMidiSession{ nullptr };
std::unique_ptr<Microsoft::Windows::Devices::Midi2::Initialization::MidiDesktopAppSdkInitializer> g_midiSdkInitializer;
std::chrono::steady_clock::time_point g_wmsConnectStartTime;
std::atomic<bool> g_pendingMidiStart{ false };
std::atomic<bool> g_autoStartTriggered{ false };

// --- Pre-Registration State -------------------------------------------
// Virtual MIDI devices are registered eagerly at MIDI service availability
// time, so ConnectPort only needs to open connections (near-instant).
std::mutex g_preRegMutex;
std::vector<PreRegisteredVirtualDevice> g_preRegisteredPorts; // index 0-4 = Out 1-5, index 5 = In
std::atomic<bool> g_preRegistrationDone{ false };
// ----------------------------------------------------------------------

void WaitForDeviceArrival(winrt::hstring deviceId, int timeoutMs = 2000) {
    if (deviceId.empty()) return;
    std::wstring id = deviceId.c_str();
    std::unique_lock<std::mutex> lock(g_creationMutex);
    bool success = g_creationCv.wait_for(lock, std::chrono::milliseconds(timeoutMs), [&]() {
        return g_arrivedDevices.find(id) != g_arrivedDevices.end();
    });
    if (!success) {
        // Log timeout warning if needed
    }
}

// Parallel version: waits until ALL supplied device IDs have arrived (or timeout).
// Much faster than one sequential WaitForDeviceArrival per port because all
// PnP notifications fire concurrently.
void WaitForAllDevices(const std::vector<winrt::hstring>& ids, int timeoutMs = 2000) {
    if (ids.empty()) return;
    std::unique_lock<std::mutex> lock(g_creationMutex);
    g_creationCv.wait_for(lock, std::chrono::milliseconds(timeoutMs), [&]() {
        return std::all_of(ids.begin(), ids.end(), [](const winrt::hstring& id) {
            return !id.empty() && g_arrivedDevices.count(id.c_str()) > 0;
        });
    });
}
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
    APP_OUT_4,
    APP_OUT_5
};
static const int LOG_SOURCE_COUNT = 8;

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

void SaveSettings(const std::string& comName, const std::string& baseName, bool autoStart, bool startWin, bool autoReconn, bool startTray, ImVec4 cSys, ImVec4 cIn, ImVec4* cOut, bool stayOnTop, bool startMinimized, bool lightUI, bool quickInit) {
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
    WritePrivateProfileStringA("General", "QuickInitialization", quickInit ? "1" : "0", ini.c_str());
    WritePrivateProfileStringA("General", "AutoGetSynthInfo", g_autoGetSynthInfo.load() ? "1" : "0", ini.c_str());
    WritePrivateProfileStringA("Ports", "PortInEnabled", g_portInEnabled.load() ? "1" : "0", ini.c_str());
    for (int p = 0; p < 5; ++p)
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
    for (int i=0; i<5; i++) saveColor(("ColOut" + std::to_string(i+1)).c_str(), cOut[i]);
}

void LoadSettings(std::string& comName, char* baseName, bool& autoStart, bool& startWin, bool& autoReconn, bool& startTray, ImVec4& cSys, ImVec4& cIn, ImVec4* cOut, bool& stayOnTop, bool& startMinimized, bool& lightUI, bool& quickInit) {
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
    quickInit     = GetPrivateProfileIntA("General", "QuickInitialization", 1, ini.c_str()) != 0;
    g_autoGetSynthInfo.store(GetPrivateProfileIntA("General", "AutoGetSynthInfo", 1, ini.c_str()) != 0);
    g_portInEnabled.store(GetPrivateProfileIntA("Ports", "PortInEnabled", 1, ini.c_str()) != 0);
    for (int p = 0; p < 5; ++p)
        g_portEnabled[p].store(GetPrivateProfileIntA("Ports", ("Port" + std::to_string(p+1) + "Enabled").c_str(), (p < 4 ? 1 : 0), ini.c_str()) != 0);

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
    cOut[4] = loadColor("ColOut5", ImVec4(0.4f, 1.0f, 1.0f, 1.0f));
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
    auto g_appStartTime = std::chrono::steady_clock::now();
    g_wmsConnectStartTime = std::chrono::steady_clock::time_point{}; // Initialize to 'epoch' (not started)



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
        std::wcout << L"Error: Another instance of ToHost Bridge is already running (Mutex exists). Exiting." << std::endl;
        HWND hExisting = FindWindowW(L"ToHostBridge", nullptr);
        if (hExisting) {
            ShowWindow(hExisting, SW_RESTORE);
            SetForegroundWindow(hExisting);
        }
        return 0;
    }

    winrt::init_apartment();
    
    // Asynchronous MIDI Service Initialization with 15-second timeout
    std::thread midiInitThread([]() {
        winrt::init_apartment();
        g_midiSdkInitializer = std::make_unique<Microsoft::Windows::Devices::Midi2::Initialization::MidiDesktopAppSdkInitializer>();
        bool sdkAvailable = g_midiSdkInitializer->InitializeSdkRuntime() && g_midiSdkInitializer->EnsureServiceAvailable();
        
        if (sdkAvailable) {
            midiAvailable.store(true);
            g_midiStatus.store(MidiServiceStatus::AVAILABLE);

            try {
                g_sharedMidiSession = winrt::Microsoft::Windows::Devices::Midi2::MidiSession::Create(L"ToHost Bridge Session");
                g_wmsConnectStartTime = std::chrono::steady_clock::now();
                AddLog(LogSourceType::APP_INFO, "Establishing WMS service connection...");
            } catch (...) {
                AddLog(LogSourceType::APP_INFO, "Failed to create shared MIDI session.");
            }

            // Setup and start the active device watcher
            try {
                using namespace winrt::Microsoft::Windows::Devices::Midi2;
                g_watcher = MidiEndpointDeviceWatcher::Create();
                if (g_watcher) {
                    g_watcherAddedToken = g_watcher.Added([](MidiEndpointDeviceWatcher const&, MidiEndpointDeviceInformationAddedEventArgs const& args) {
                        if (!args || !args.AddedDevice()) return;
                        std::lock_guard<std::mutex> lock(g_creationMutex);
                        g_arrivedDevices.insert(args.AddedDevice().EndpointDeviceId().c_str());
                        g_creationCv.notify_all();
                    });
                    g_watcher.Start();
                }
            } catch (...) {
                // If watcher fails, we'll fall back to timeout logic
            }
        } else {
            midiAvailable.store(false);
            g_midiStatus.store(MidiServiceStatus::UNAVAILABLE);
        }
    });
    midiInitThread.detach();

    // Wait for status change with timeout (on a separate thread)
    std::thread([]() {
        auto start = std::chrono::steady_clock::now();
        while (g_midiStatus.load() == MidiServiceStatus::INITIALIZING) {
            if (std::chrono::duration_cast<std::chrono::seconds>(std::chrono::steady_clock::now() - start).count() >= 15) {
                g_midiStatus.store(MidiServiceStatus::NOT_RESPONDING);
                midiAvailable.store(false);
                break;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(100));
        }
    }).detach();


    ImGui_ImplWin32_EnableDpiAwareness();
    float main_scale = ImGui_ImplWin32_GetDpiScaleForMonitor(::MonitorFromPoint(POINT{ 0, 0 }, MONITOR_DEFAULTTOPRIMARY));

    WNDCLASSEXW wc = { sizeof(wc), CS_CLASSDC, WndProc, 0L, 0L, GetModuleHandle(nullptr), LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1)), nullptr, nullptr, nullptr, L"ToHostBridge", LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1)) };
    ::RegisterClassExW(&wc);
    // Adjusted window dimensions for bold headers and better spacing
    HWND hwnd = ::CreateWindowW(wc.lpszClassName, L"ToHost Bridge v1.3.0", WS_OVERLAPPEDWINDOW, 100, 100, (int)(530 * main_scale), (int)(430 * main_scale), nullptr, nullptr, wc.hInstance, nullptr);

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
    bool quickInitialization = true;
    ImVec4 colSys, colIn, colOut[5]; // Corrected size to 5 ports

    // Init port-enable flags to true for first 4 ports before LoadSettings (which may override)
    for (int i = 0; i < 5; ++i) g_portEnabled[i].store(i < 4);

    LoadSettings(savedComName, baseNameBuf, autoStartVirtualMidi, startWithWindows, autoReconnect, startToTray, colSys, colIn, colOut, stayOnTop, startMinimized, lightUI, quickInitialization);
    // Sync registry with loaded INI setting (especially important on first run after MSI install)
    SetStartWithWindowsRegistry(startWithWindows);

    // Load imgui layout AFTER LoadSettings so the INI file path is established
    LoadImguiIniFromSettings();
    if (lightUI)  ImGui::StyleColorsLight();
    if (stayOnTop) SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
    
    // Eagerly populate settings.ini on first boot
    if (GetFileAttributesA(GetIniPath().c_str()) == INVALID_FILE_ATTRIBUTES) {
        SaveSettings(savedComName, baseNameBuf, autoStartVirtualMidi, startWithWindows, autoReconnect, startToTray, colSys, colIn, colOut, stayOnTop, startMinimized, lightUI, quickInitialization);
    }

    // --- Eager Pre-Registration Thread -----------------------------------
    // Now that settings are loaded we know the base name and port enables.
    // Pre-register all virtual MIDI devices in the background so the
    // Windows MIDI Service can pipeline the PnP work while the UI loads.
    // ConnectPort() will use these pre-registered devices for near-instant open.
    {
        std::string  preRegBaseName(baseNameBuf);
        bool         preRegPortEnabled[5];
        for (int i = 0; i < 5; ++i) preRegPortEnabled[i] = g_portEnabled[i].load();
        bool         preRegInEnabled = g_portInEnabled.load();

        std::thread([preRegBaseName, preRegPortEnabled, preRegInEnabled, quickInitialization]() {
            winrt::init_apartment();

            // Wait until the MIDI service is available (or gives up)
            auto waitStart = std::chrono::steady_clock::now();
            while (g_midiStatus.load() == MidiServiceStatus::INITIALIZING) {
                if (std::chrono::duration_cast<std::chrono::seconds>(
                        std::chrono::steady_clock::now() - waitStart).count() >= 20) break;
                std::this_thread::sleep_for(std::chrono::milliseconds(50));
            }
            // Also wait for the shared session to exist (created slightly after AVAILABLE is set)
            waitStart = std::chrono::steady_clock::now();
            while (!g_sharedMidiSession) {
                if (std::chrono::duration_cast<std::chrono::seconds>(
                        std::chrono::steady_clock::now() - waitStart).count() >= 5) break;
                std::this_thread::sleep_for(std::chrono::milliseconds(20));
            }

            if (g_midiStatus.load() != MidiServiceStatus::AVAILABLE || !g_sharedMidiSession) {
                AddLog(LogSourceType::APP_INFO, "[PreReg] Skipped: MIDI service not available.");
                return;
            }

            if (!quickInitialization) {
                AddLog(LogSourceType::APP_INFO, "[PreReg] Waiting 20s for Windows MIDI Service to stabilize...");
                std::this_thread::sleep_for(std::chrono::seconds(20));
            }

            std::wstring wBase(preRegBaseName.begin(), preRegBaseName.end());
            std::vector<PreRegisteredVirtualDevice> regs(6); // 0-4 = Out 1-5, 5 = In

            for (int i = 0; i < 5; ++i) {
                if (!preRegPortEnabled[i]) { regs[i] = {}; continue; }
                
                std::wstring name = wBase + L" Out " + std::to_wstring(i + 1);
                regs[i] = VirtualMidiPort::PreRegister(name);
                AddLog(LogSourceType::APP_INFO,
                    std::string("[PreReg] Out ") + std::to_string(i+1) +
                    (regs[i].valid ? " OK" : " FAILED"));
                    
                if (!quickInitialization && regs[i].valid) {
                    std::wstring targetId(regs[i].deviceEndpointId.c_str());
                    std::unique_lock<std::mutex> lock(g_creationMutex);
                    g_creationCv.wait_for(lock, std::chrono::milliseconds(3000), [&]() {
                        return g_arrivedDevices.find(targetId) != g_arrivedDevices.end();
                    });
                    // Brief breather to let Windows serialize internally
                    std::this_thread::sleep_for(std::chrono::milliseconds(50));
                }
            }
            if (preRegInEnabled) {
                regs[5] = VirtualMidiPort::PreRegister(wBase + L" In");
                AddLog(LogSourceType::APP_INFO,
                    std::string("[PreReg] In ") + (regs[5].valid ? "OK" : "FAILED"));

                if (!quickInitialization && regs[5].valid) {
                    std::wstring targetId(regs[5].deviceEndpointId.c_str());
                    std::unique_lock<std::mutex> lock(g_creationMutex);
                    g_creationCv.wait_for(lock, std::chrono::milliseconds(3000), [&]() {
                        return g_arrivedDevices.find(targetId) != g_arrivedDevices.end();
                    });
                    std::this_thread::sleep_for(std::chrono::milliseconds(50));
                }
            }

            {
                std::lock_guard<std::mutex> lk(g_preRegMutex);
                g_preRegisteredPorts = std::move(regs);
            }
            g_preRegistrationDone.store(true);
            AddLog(LogSourceType::APP_INFO, "[PreReg] All endpoints pre-registered.");
        }).detach();
    }
    // ---------------------------------------------------------------------
    
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
    std::mutex connStatusMutex;
    std::string connectionStatus = "";
    
    auto SetConnStatus = [&](const std::string& s) {
        std::lock_guard<std::mutex> lock(connStatusMutex);
        connectionStatus = s;
    };
    auto GetConnStatus = [&]() -> std::string {
        std::lock_guard<std::mutex> lock(connStatusMutex);
        return connectionStatus;
    };

    // Startup: if saved port not present yet, mark as waiting
    if (autoStartVirtualMidi && !savedComName.empty() && !foundSavedPort) {
        connectionLost.store(true);
        activeComName = savedComName;
        SetConnStatus("Waiting for " + savedComName + "...");
    }

    // Filters
    bool filterSys = true;
    bool filterIn = true;
    bool filterOut[5] = {true, true, true, true, true};
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
        bool out[5];
        bool operator!=(const FilterState& o) const {
            return sys!=o.sys || in_!=o.in_ || sync!=o.sync ||
                   out[0]!=o.out[0] || out[1]!=o.out[1] || out[2]!=o.out[2] || out[3]!=o.out[3] ||
                   out[4]!=o.out[4];
        }
    };
    FilterState lastFilterState = {true, true, true, {true,true,true,true,true}};

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
                case LogSourceType::APP_OUT_5:   show = filterOut[4]; break;
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
        lastFilterState = {filterSys, filterIn, filterSync, {filterOut[0],filterOut[1],filterOut[2],filterOut[3],filterOut[4]}};
    };

    // std::function allows recursive/forward reference capture without 'own initializer' errors
    std::function<bool()> ConnectPort;
    ConnectPort = [&]() -> bool {
        if (isConnected || g_midiStatus.load() != MidiServiceStatus::AVAILABLE) return false;
        if (comPorts.empty()) return false;
        std::string baseNameStr(baseNameBuf);
        ResetSafeSerialPort();

        std::string comName = comPorts[selectedComPort];
        activeComName = comName;

        // --- Open virtual MIDI ports -----------------------------------------------
        // Fast path: use pre-registered loopback pairs (registration already done
        // in background, so CreateEndpointConnection+Open is near-instant).
        // Slow path fallback: full CreateTransientLoopbackEndpoints + serial PnP wait.
        if (GetSafeVirtualPortsOut().empty() && !g_virtualPortIn) {
            ClearSafeVirtualPortsOut();
            ResetSafeVirtualPortIn();
            std::wstring wBaseName(baseNameStr.begin(), baseNameStr.end());
            std::string  currentlyOpened;

            // Grab pre-registered data (if ready) under lock, then release
            std::vector<PreRegisteredVirtualDevice> preRegs;
            bool useFastPath = false;
            {
                std::lock_guard<std::mutex> lk(g_preRegMutex);
                if (g_preRegistrationDone.load() && !g_preRegisteredPorts.empty()) {
                    preRegs    = g_preRegisteredPorts;
                    useFastPath = true;
                }
            }
            AddLog(LogSourceType::APP_INFO,
                useFastPath ? "Using fast-path (pre-registered endpoints)." : "Using slow-path (registering now).");

            std::vector<winrt::hstring> allDeviceIds; // collected for single parallel PnP wait

            for (int i = 1; i <= 5; ++i) {
                if (!g_isConnecting.load()) {
                    AddLog(LogSourceType::APP_INFO, "Connection aborted by user.");
                    return false;
                }
                if (!g_portEnabled[i-1].load(std::memory_order_relaxed)) continue;

                std::wstring portName = wBaseName + L" Out " + std::to_wstring(i);
                LogSourceType srcType = (LogSourceType)((int)LogSourceType::APP_OUT_1 + (i-1));

                auto rxCb = [i, srcType](const std::vector<uint8_t>& data) {
                    if (!g_transmitSync && !data.empty() && data[0] >= 0xF8) return;
                    bool isSync = (data.size() == 1 && data[0] >= 0xF8);
                    SendToSerial(i, data);
                    AddLog(srcType, "[App->COM] P" + std::to_string(i) + ": " + BytesToHex(data), isSync, GetMidiCommandName(data));
                };

                if (!currentlyOpened.empty()) currentlyOpened += ", ";
                SetConnStatus(currentlyOpened + "Opening Out " + std::to_string(i) + "...");

                std::unique_ptr<VirtualMidiPort> vport;
                if (useFastPath && (i-1) < (int)preRegs.size() && preRegs[i-1].valid) {
                    // FAST PATH: endpoint already in service, just open connection
                    vport = std::make_unique<VirtualMidiPort>(g_sharedMidiSession, preRegs[i-1], rxCb);
                } else {
                    // SLOW PATH: full registration (fallback)
                    vport = std::make_unique<VirtualMidiPort>(g_sharedMidiSession, portName, rxCb);
                    int retry = 2;
                    while (!vport->IsValid() && retry <= 10) {
                        if (!g_isConnecting.load()) return false;
                        vport = std::make_unique<VirtualMidiPort>(g_sharedMidiSession,
                            portName + L" (" + std::to_wstring(retry++) + L")", rxCb);
                    }
                }

                if (vport->IsValid()) {
                    allDeviceIds.push_back(vport->GetDeviceId());
                    AddSafeVirtualPortOut(std::move(vport));
                } else {
                    AddSafeVirtualPortOut(nullptr); // preserve index
                }
                currentlyOpened += "Out " + std::to_string(i);
            }

            if (g_portInEnabled.load(std::memory_order_relaxed)) {
                if (!g_isConnecting.load()) {
                    AddLog(LogSourceType::APP_INFO, "Connection aborted by user.");
                    return false;
                }
                std::wstring inPortName = wBaseName + L" In";
                SetConnStatus(currentlyOpened + "|Opening MIDI In ...");

                std::shared_ptr<VirtualMidiPort> vportIn;
                if (useFastPath && (int)preRegs.size() >= 6 && preRegs[5].valid) {
                    vportIn = std::make_shared<VirtualMidiPort>(g_sharedMidiSession, preRegs[5], nullptr);
                } else {
                    vportIn = std::make_shared<VirtualMidiPort>(g_sharedMidiSession, inPortName, nullptr);
                    int retry = 2;
                    while (!vportIn->IsValid() && retry <= 10) {
                        if (!g_isConnecting.load()) return false;
                        vportIn = std::make_shared<VirtualMidiPort>(g_sharedMidiSession,
                            inPortName + L" (" + std::to_wstring(retry++) + L")", nullptr);
                    }
                }
                if (vportIn->IsValid()) {
                    allDeviceIds.push_back(vportIn->GetDeviceId());
                    SetSafeVirtualPortIn(vportIn);
                }
            }

            // Single parallel PnP wait — all arrivals fire concurrently.
            // For fast path, devices already arrived during pre-reg; this returns immediately.
            // For slow path, waits up to 2s for all ports at once (vs. 2s * N serially).
            WaitForAllDevices(allDeviceIds, 2000);

            // Consume pre-reg data — clear so reconnect uses fresh registration
            if (useFastPath) {
                std::lock_guard<std::mutex> lk(g_preRegMutex);
                g_preRegisteredPorts.clear();
                g_preRegistrationDone.store(false);
            }

            SetConnStatus("");
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

        if (GetSafeSerialPort() && GetSafeSerialPort()->IsOpen()) {
            AddLog(LogSourceType::APP_INFO, "Connected to " + comName + " (" + baseNameStr + ")");
            isConnected = true;
            connectionLost = false;
            g_midiStatus.store(MidiServiceStatus::AVAILABLE);

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
                    Sleep(15000);
                    if (g_waitingForIdentity) {
                        std::lock_guard<std::mutex> lock(g_synthNameOverlayMutex);
                        if (g_detectedSynth == "Querying...") g_detectedSynth = "";
                        g_waitingForIdentity = false;
                    }
                }).detach();
            }

            // Persist the newly-connected COM port name to settings
            SaveSettings(comName, baseNameStr, autoStartVirtualMidi, startWithWindows, autoReconnect, startToTray, colSys, colIn, colOut, stayOnTop, startMinimized, lightUI, quickInitialization);
            SetConnStatus("");
            return true;
        } else {
            AddLog(LogSourceType::APP_INFO, "Failed to open " + comName);
            ResetSafeSerialPort();
            SetConnStatus("Failed to open " + comName + ". Is it in use?");
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
            if (autoStartVirtualMidi) {
                g_pendingMidiStart.store(true);
            }
            g_autoStartTriggered.store(true);
        }

        // Optimized Transition Logic: Fast-stepped startup (Fake delay removed, replaced by visual-only fast sequence)
        if (g_pendingMidiStart.load() && !isConnected && !g_isConnecting.load() && g_wmsConnectStartTime != std::chrono::steady_clock::time_point{}) {
            auto now = std::chrono::steady_clock::now();
            auto elapsedMs = std::chrono::duration_cast<std::chrono::milliseconds>(now - g_wmsConnectStartTime).count();
            int wmsWaitDelayMs = quickInitialization ? 1500 : 26000;
            // We'll use a very fast visual sequence to satisfy the "near instant" requirement while keeping premium feel.
            if (elapsedMs >= wmsWaitDelayMs && g_preRegistrationDone.load()) { 
                g_pendingMidiStart.store(false);
                g_isConnecting.store(true);
                SetConnStatus("Connecting...");
                std::thread([&]() {
                    ConnectPort();
                    g_isConnecting.store(false);
                }).detach();
            }
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
            std::wstring wTitle = L"ToHost Bridge v1.3.0";

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
                    SetConnStatus("COM port disconnected!");
                }

                bool isBooting = false;
                // Loopback is instant, no boot period needed.

                // Auto-reconnect when port reappears (triggered by either drop recovery or initial auto-start wait)
                if (!isBooting && !isConnected && connectionLost.load() && portExists && (autoReconnect || autoStartVirtualMidi)) {
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
                        SetConnStatus("Waiting for " + activeComName + " (retrying)...");
                    }
                }
            }
        }

        ImGui_ImplDX11_NewFrame();
        ImGui_ImplWin32_NewFrame();
        ImGui::NewFrame();

        if (g_isShuttingDown.load()) {
            // Full-screen "Closing" overlay
            ImGui::SetNextWindowPos(ImVec2(0, 0));
            ImGui::SetNextWindowSize(io.DisplaySize);
            ImGui::Begin("ShutdownOverlay", nullptr, ImGuiWindowFlags_NoDecoration | ImGuiWindowFlags_NoInputs | ImGuiWindowFlags_NoNav);
            
            float centerX = io.DisplaySize.x * 0.5f;
            float centerY = io.DisplaySize.y * 0.5f;
            
            // Move everything higher (Start higher than center)
            float currentY = centerY - 160.0f;

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
                        for (int p = 1; p <= 5; ++p) {
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
            ImGui::End();
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

            ImVec2 center = ImGui::GetMainViewport()->GetCenter();
            ImGui::SetNextWindowPos(center, ImGuiCond_Appearing, ImVec2(0.5f, 0.5f));
            if (ImGui::BeginPopupModal("Update Available", nullptr, ImGuiWindowFlags_AlwaysAutoResize)) {
                ImGui::Spacing();
                
                auto centerText = [](const char* text, bool bold = false) {
                    float windowWidth = ImGui::GetWindowSize().x;
                    float textWidth = ImGui::CalcTextSize(text).x;
                    ImGui::SetCursorPosX((windowWidth - textWidth) * 0.5f);
                    if (bold) ImGui::TextColored(ImVec4(1.0f, 1.0f, 1.0f, 1.0f), "%s", text);
                    else ImGui::Text("%s", text);
                };

                centerText("Update Available via Microsoft Store", true);
                
                if (g_mockUpdate) {
                    ImGui::PushStyleColor(ImGuiCol_Text, ImVec4(1.0f, 1.0f, 0.0f, 1.0f));
                    centerText("[MOCK MODE: v1.2.1.0]");
                    ImGui::PopStyleColor();
                }
                
                ImGui::Spacing();
                ImGui::Separator();
                ImGui::Spacing();

                // Center buttons
                float b1Width = 180.0f;
                float b2Width = 140.0f;
                float totalWidth = b1Width + b2Width + ImGui::GetStyle().ItemSpacing.x;
                ImGui::SetCursorPosX((ImGui::GetWindowSize().x - totalWidth) * 0.5f);
                
                if (ImGui::Button("Close & update", ImVec2(b1Width, 40))) {
                    if (g_mockUpdate) {
                        ::MessageBoxA(NULL, "Mock update triggered. PFN: MaxCoppola.ToHostBridge_cmgt2fcrje4mm", "Debug", MB_OK);
                    } else {
                        ::ShellExecuteA(NULL, "open", "ms-windows-store://pdp/?PFN=MaxCoppola.ToHostBridge_cmgt2fcrje4mm", NULL, NULL, SW_SHOWNORMAL);
                    }
                    g_isShuttingDown.store(true);
                    ImGui::CloseCurrentPopup();
                }
                ImGui::SameLine();
                if (ImGui::Button("Maybe later", ImVec2(b2Width, 40))) {
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

                // --- TOP: MIDI Service Status / Errors (Highest Priority) ---
                MidiServiceStatus currentStatus = g_midiStatus.load();
                bool showMidiError = true;
                if (currentStatus == MidiServiceStatus::NOT_RESPONDING) {
                    auto elapsed = std::chrono::duration_cast<std::chrono::seconds>(std::chrono::steady_clock::now() - g_appStartTime).count();
                    if (elapsed < 90) showMidiError = false;
                }

                if (currentStatus == MidiServiceStatus::INITIALIZING) {
                    ImGui::TextWrapped("Checking MIDI Service...");
                    ImGui::Separator();
                } else if (showMidiError && (currentStatus == MidiServiceStatus::NOT_RESPONDING || currentStatus == MidiServiceStatus::UNAVAILABLE)) {
                    ImGui::PushStyleColor(ImGuiCol_ChildBg, ImVec4(0.4f, 0.1f, 0.1f, 1.0f));
                    ImGui::BeginChild("MidiError", ImVec2(0, 100 * main_scale), true); // Fixed height for top placement
                    if (currentStatus == MidiServiceStatus::NOT_RESPONDING) {
                        ImGui::TextColored(ImVec4(1.0f, 0.8f, 0.0f, 1.0f), "CRITICAL ERROR: Windows MIDI Service is not responding.");
                        ImGui::TextWrapped("Recommendation: Restart computer or stop 'Windows MIDI Service'.");
                        if (ImGui::Button("Open Task Manager")) ShellExecuteA(NULL, "open", "taskmgr.exe", NULL, NULL, SW_SHOWNORMAL);
                    } else {
                        ImGui::TextColored(ImVec4(1.0f, 0.2f, 0.2f, 1.0f), "ERROR: Windows MIDI Services Runtime is missing.");
                        if (ImGui::Button("Download Installer")) ShellExecuteA(NULL, "open", "https://microsoft.github.io/MIDI/get-latest/", NULL, NULL, SW_SHOWNORMAL);
                    }
                    ImGui::EndChild();
                    ImGui::PopStyleColor();
                    ImGui::Separator();
                }


                bool portsRunning = !GetSafeVirtualPortsOut().empty();

                ImGui::Columns(2, "ConnectionSplit", true);
                ImGui::SetColumnWidth(0, 263 * main_scale); 

                // --- LEFT COLUMN: Configuration & Connectivity ---
                ImGui::BeginDisabled(portsRunning || isConnected);
                ImGui::SetNextItemWidth(120 * main_scale); // Reduced from 150
                if (ImGui::InputText("Virtual Port Name", baseNameBuf, 128)) {
                    SaveSettings("", baseNameBuf, autoStartVirtualMidi, startWithWindows, autoReconnect, startToTray, colSys, colIn, colOut, stayOnTop, startMinimized, lightUI, quickInitialization);
                }
                ImGui::EndDisabled();

                ImGui::Spacing();
                bool disableComSelect = isConnected || comPorts.empty();
                if (disableComSelect) ImGui::BeginDisabled();
                ImGui::SetNextItemWidth(80 * main_scale); // 2/3 of previous 120
                ImGui::Combo("COM Port", &selectedComPort, comPorts.empty() ? nullptr : comPorts.data(), (int)comPorts.size());
                if (disableComSelect) ImGui::EndDisabled();

                ImGui::SameLine();
                ImGui::BeginDisabled(isConnected);
                if (ImGui::Button("Refresh", ImVec2(ImGui::GetContentRegionAvail().x, 0))) {
                    comPortsVec = GetAvailableComPorts();
                    comPorts.clear();
                    for (const auto& p : comPortsVec) comPorts.push_back(p.c_str());
                }
                ImGui::EndDisabled();

                ImGui::Spacing();
                bool disableConnect = isConnected || g_isConnecting.load() || g_midiStatus.load() != MidiServiceStatus::AVAILABLE;
                ImGui::BeginDisabled(disableConnect);
                
                float halfBtnWidth = (ImGui::GetContentRegionAvail().x - ImGui::GetStyle().ItemSpacing.x) / 2.0f;
                
                bool isBooting = false;
                if (g_wmsConnectStartTime != std::chrono::steady_clock::time_point{}) {
                    auto now = std::chrono::steady_clock::now();
                    auto elapsed = std::chrono::duration_cast<std::chrono::seconds>(now - g_wmsConnectStartTime).count();
                    int wmsWaitDelay = quickInitialization ? 2 : 26;
                    if (elapsed < wmsWaitDelay) isBooting = true;
                }

                if (!isConnected && !g_pendingMidiStart.load()) {
                    if (ImGui::Button("Connect", ImVec2(halfBtnWidth, 30 * main_scale))) {
                        if (isBooting) {
                            g_pendingMidiStart.store(true);
                        } else if (!g_isConnecting.load()) {
                            if (g_midiStatus.load() != MidiServiceStatus::AVAILABLE) {
                                connectionStatus = "Cannot connect: Windows MIDI Service is not responding.";
                                AddLog(LogSourceType::APP_INFO, "Connection aborted: MIDI Service not available.");
                            } else {
                                g_isConnecting.store(true);
                                SetConnStatus("Connecting...");
                                std::thread([&]() {
                                    std::atomic<bool> portOpDone{false};
                                    std::thread worker([&]() {
                                        ConnectPort();
                                        portOpDone.store(true);
                                    });
                                    auto start = std::chrono::steady_clock::now();
                                    while (!portOpDone.load()) {
                                        if (std::chrono::duration_cast<std::chrono::seconds>(std::chrono::steady_clock::now() - start).count() >= 15) {
                                            g_midiStatus.store(MidiServiceStatus::NOT_RESPONDING);
                                            SetConnStatus("Cannot connect: Windows MIDI Service is not responding.");
                                            AddLog(LogSourceType::APP_INFO, "Connection timed out after 15s (MIDI Service hang).");
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
                }
                ImGui::EndDisabled();
                
                if (isConnected || g_pendingMidiStart.load() || g_isConnecting.load()) {
                    const char* btnLabel = isConnected ? "Disconnect" : "Stop Connection";
                    if (ImGui::Button(btnLabel, ImVec2(halfBtnWidth, 30 * main_scale))) {
                        if (g_pendingMidiStart.load()) {
                            g_pendingMidiStart.store(false);
                            connectionLost.store(false);
                        } else if (g_isConnecting.load()) {
                             // Force abort flag if we can, otherwise just update state
                             g_isConnecting.store(false); 
                             SetConnStatus("Connection aborted.");
                        } else {
                            ResetSafeSerialPort();
                            isConnected = false;
                            connectionLost.store(false);
                            SetConnStatus("Disconnected (virtual ports still active).");
                            AddLog(LogSourceType::APP_INFO, "Disconnected from COM port.");
                        }
                    }
                }


                // --- Status Messages (Moved to Bottom) ---

                // --- RIGHT COLUMN: Active Ports ---
                ImGui::NextColumn();
                ImGui::Spacing();
                
                // Vertical centering for the checkboxes
                ImGui::Dummy(ImVec2(0, 10 * main_scale));

                ImGui::BeginDisabled(portsRunning || isConnected);
                {
                    // Row 1: Out 1, 2, 3
                    for (int i = 0; i < 3; ++i) {
                        bool en = g_portEnabled[i].load(std::memory_order_relaxed);
                        std::string lbl = "Out " + std::to_string(i+1) + "##pe";
                        if (ImGui::Checkbox(lbl.c_str(), &en)) g_portEnabled[i].store(en, std::memory_order_relaxed);
                        if (i < 2) ImGui::SameLine();
                    }
                    ImGui::Spacing();
                    // Row 2: Out 4, 5, In
                    for (int i = 3; i < 5; ++i) {
                        bool en = g_portEnabled[i].load(std::memory_order_relaxed);
                        std::string lbl = "Out " + std::to_string(i+1) + "##pe";
                        if (ImGui::Checkbox(lbl.c_str(), &en)) g_portEnabled[i].store(en, std::memory_order_relaxed);
                        ImGui::SameLine();
                    }
                    bool enIn = g_portInEnabled.load(std::memory_order_relaxed);
                    if (ImGui::Checkbox("In##pe", &enIn)) g_portInEnabled.store(enIn, std::memory_order_relaxed);
                }
                ImGui::EndDisabled();

                ImGui::Columns(1);
                ImGui::Separator();

                // --- BOTTOM: Connection Status & Extras ---
                
                float footerHeight = 0.0f;
                if (isConnected) footerHeight = 52.0f * main_scale;

                // Detailed Manual Stop Progress (checklist style)
                if (g_isManualStopping.load()) {
                    ImGui::BeginChild("ManualStopTasks", ImVec2(0, -(footerHeight + 4.0f * main_scale)), true);
                    
                    std::lock_guard<std::mutex> lock(g_shutdownMutex);
                    for (const auto& task : g_manualStopTasks) {
                        ImVec4 color = (task.status == "OK") ? ImVec4(0.4f, 1.0f, 0.4f, 1.0f) : (task.status == "..." ? ImVec4(1.0f, 1.0f, 0.0f, 1.0f) : ImVec4(0.7f, 0.7f, 0.7f, 1.0f));
                        ImGui::Text("%s", task.name.c_str());
                        ImGui::SameLine(220.0f * main_scale);
                        ImGui::TextColored(color, "[ %s ]", task.status.c_str());
                    }
                    ImGui::EndChild();
                } else if (!isConnected && !GetSafeVirtualPortsOut().empty()) {
                    ImGui::Spacing();
                    ImGui::BeginDisabled(g_isConnecting.load());
                    ImGui::PushStyleColor(ImGuiCol_Button,        ImVec4(0.6f, 0.3f, 0.0f, 1.0f));
                    ImGui::PushStyleColor(ImGuiCol_ButtonHovered, ImVec4(0.8f, 0.4f, 0.0f, 1.0f));
                    ImGui::PushStyleColor(ImGuiCol_ButtonActive,  ImVec4(0.5f, 0.2f, 0.0f, 1.0f));
                    if (ImGui::Button("Stop Virtual Ports", ImVec2(-1, 28 * main_scale))) {
                        std::thread([&]() {
                            g_isConnecting.store(true);
                            g_isManualStopping.store(true);
                            connectionLost.store(false); // Make sure it doesn't try to reconnect
                            g_pendingMidiStart.store(false); // Cancel any pending boot connection
                            
                            std::vector<std::thread> cleanThreads;
                            std::vector<int> validOutIndices;
                            bool hasVirtualIn = false;
                            {
                                std::lock_guard<std::mutex> lock(g_shutdownMutex);
                                g_manualStopTasks.clear();
                                auto outs = GetSafeVirtualPortsOut();
                                for (int i = 0; i < (int)outs.size(); ++i) {
                                    if (outs[i]) {
                                        g_manualStopTasks.push_back({"Close Virtual Out " + std::to_string(i+1), "Pending"});
                                        validOutIndices.push_back(i);
                                    }
                                }
                                if (GetSafeVirtualPortIn()) {
                                    g_manualStopTasks.push_back({"Close Virtual In", "Pending"});
                                    hasVirtualIn = true;
                                }
                            }

                            int tIdx = 0;
                            for (int i : validOutIndices) {
                                int curIdx = tIdx; // capture current task index
                                cleanThreads.emplace_back([i, curIdx]() {
                                    { std::lock_guard<std::mutex> lock(g_shutdownMutex); g_manualStopTasks[curIdx].status = "..."; }
                                    ResetSafeVirtualPortsOutIndex(i);
                                    { std::lock_guard<std::mutex> lock(g_shutdownMutex); g_manualStopTasks[curIdx].status = "OK"; }
                                });
                                ::Sleep(1000); // 1s stagger like in shutdown logic
                                tIdx++;
                            }
                            if (hasVirtualIn) {
                                int curIdx = tIdx;
                                cleanThreads.emplace_back([curIdx]() {
                                    { std::lock_guard<std::mutex> lock(g_shutdownMutex); g_manualStopTasks[curIdx].status = "..."; }
                                    ResetSafeVirtualPortIn();
                                    { std::lock_guard<std::mutex> lock(g_shutdownMutex); g_manualStopTasks[curIdx].status = "OK"; }
                                });
                                ::Sleep(1000);
                            }

                            for (auto& t : cleanThreads) if (t.joinable()) t.join();
                            
                            ClearSafeVirtualPortsOut();
                            SetConnStatus("Disconnected.");
                            AddLog(LogSourceType::APP_INFO, "Virtual MIDI ports stopped.");
                            g_isManualStopping.store(false);
                            g_isConnecting.store(false);
                        }).detach();
                    }
                    ImGui::PopStyleColor(3);
                    ImGui::EndDisabled();
                    ImGui::Spacing();
                }

                if (!g_isManualStopping.load()) {
                    ImGui::BeginChild("StatusScrollArea", ImVec2(0, -footerHeight), false, ImGuiWindowFlags_None);
                    
                    bool showWmsMessage = false;

                    if (showWmsMessage) {
                        // Already showed it
                    } else if (g_pendingMidiStart.load()) {
                         ImGui::TextWrapped("Queued for connection... (waiting for WMS)");
                         if (g_midiStatus.load() == MidiServiceStatus::NOT_RESPONDING) {
                             ImGui::PushStyleColor(ImGuiCol_Text, ImVec4(1.0f, 0.4f, 0.4f, 1.0f));
                             ImGui::TextWrapped("!! MIDI SERVICE NOT RESPONDING !! Check Task Manager.");
                             ImGui::PopStyleColor();
                         }
                    } else if (g_isConnecting.load()) {
                         std::string curStat = GetConnStatus();
                         ImGui::TextWrapped("%s", curStat.c_str());
                         if (g_midiStatus.load() == MidiServiceStatus::NOT_RESPONDING) {
                             ImGui::PushStyleColor(ImGuiCol_Text, ImVec4(1.0f, 0.5f, 0.0f, 1.0f));
                             ImGui::TextWrapped("Warning: MIDI Service hang detected. This may take a while.");
                             ImGui::PopStyleColor();
                         }
                    } else if (!connectionLost.load() && !isConnected && GetConnStatus().find("Opening") != std::string::npos) {
                        std::string curStat = GetConnStatus();
                        float alignOffset = ImGui::CalcTextSize("Virtual outputs: ").x + 10.0f;
                        
                        size_t pipePos = curStat.find('|');
                        std::string outPart = (pipePos != std::string::npos) ? curStat.substr(0, pipePos) : curStat;
                        std::string inPart = (pipePos != std::string::npos) ? curStat.substr(pipePos + 1) : "";

                        ImGui::TextColored(lightUI ? ImVec4(0.3f, 0.3f, 0.3f, 1.0f) : ImVec4(0.7f, 0.7f, 0.7f, 1.0f), "Virtual outputs:");
                        ImGui::SameLine(alignOffset);
                        ImGui::TextWrapped("%s", outPart.empty() ? "None" : outPart.c_str());

                        if (!inPart.empty()) {
                            ImGui::TextColored(lightUI ? ImVec4(0.3f, 0.3f, 0.3f, 1.0f) : ImVec4(0.7f, 0.7f, 0.7f, 1.0f), "Virtual inputs:");
                            ImGui::SameLine(alignOffset);
                            ImGui::TextWrapped("%s", inPart.c_str());
                        }
                    } else if (connectionLost.load()) {
                        ImGui::TextColored(ImVec4(1.0f, 0.2f, 0.2f, 1.0f), "COM port disconnected!");
                    } else {
                        if (isConnected) {
                            // Discover active ports strictly from successfully loaded physical instances
                            std::string outsDisp;
                            auto activeOuts = GetSafeVirtualPortsOut();
                            for (int i = 0; i < (int)activeOuts.size(); ++i) {
                                if (activeOuts[i]) {
                                    if (!outsDisp.empty()) outsDisp += ", ";
                                    outsDisp += "Out " + std::to_string(i + 1);
                                }
                            }
                            if (outsDisp.empty()) outsDisp = "None";
                            std::string inDisp = g_portInEnabled.load() ? "In" : "None";

                            // Row 1: Connected Message
                            std::string succMsg = "Opened " + activeComName + " successfully.";
                            ImGui::Text("%s", succMsg.c_str());

                            // Row 2: Serial Charge
                            char chargeBuf[64];
                            sprintf(chargeBuf, "Serial charge: %3.0f%%", g_chargePercent);
                            ImVec4 chargeCol = lightUI ? ImVec4(0.0f, 0.45f, 0.0f, 1.0f) : ImVec4(0.6f, 1.0f, 0.6f, 1.0f);
                            ImGui::TextColored(chargeCol, "%s", chargeBuf);

                            ImGui::Spacing();
                            ImGui::Separator();
                            ImGui::Spacing();

                            float alignOffset = ImGui::CalcTextSize("Virtual outputs: ").x + 10.0f;

                            // Row 3: Virtual Outputs
                            ImGui::TextColored(lightUI ? ImVec4(0.3f, 0.3f, 0.3f, 1.0f) : ImVec4(0.7f, 0.7f, 0.7f, 1.0f), "Virtual outputs:");
                            ImGui::SameLine(alignOffset);
                            {
                                std::string stat = GetConnStatus();
                                bool stillOpening = !stat.empty() && stat.find("Opening") != std::string::npos;
                                if (stillOpening) {
                                    ImGui::TextWrapped("%s", stat.c_str());
                                } else {
                                    ImGui::TextWrapped("%s", outsDisp.c_str());
                                }
                            }

                            // Row 4: Virtual Inputs
                            ImGui::TextColored(lightUI ? ImVec4(0.3f, 0.3f, 0.3f, 1.0f) : ImVec4(0.7f, 0.7f, 0.7f, 1.0f), "Virtual inputs:");
                            ImGui::SameLine(alignOffset); 
                            ImGui::TextWrapped("%s", inDisp.c_str());
                        } else {
                            ImGui::TextWrapped("%s", GetConnStatus().c_str());
                        }
                    }
                    ImGui::EndChild();
                } else {
                    ImGui::Dummy(ImVec2(0, footerHeight));
                }

                if (isConnected) {
                    ImGui::Spacing();
                    
                    ImGui::PushStyleColor(ImGuiCol_Button, ImVec4(0.7f, 0.1f, 0.1f, 1.0f));
                    ImGui::PushStyleColor(ImGuiCol_ButtonHovered, ImVec4(0.9f, 0.2f, 0.2f, 1.0f));
                    ImGui::PushStyleColor(ImGuiCol_ButtonActive, ImVec4(0.5f, 0.0f, 0.0f, 1.0f));
                    ImGui::PushStyleColor(ImGuiCol_Text, ImVec4(1.0f, 1.0f, 1.0f, 1.0f));
                    if (ImGui::Button("Panic! (All Notes Off)", ImVec2(ImGui::GetContentRegionAvail().x * 0.5f, 30 * main_scale))) {
                        for (int p = 1; p <= 5; ++p) {
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
                    ImGui::PushStyleColor(ImGuiCol_Text, ImVec4(1.0f, 1.0f, 1.0f, 1.0f));
                    if (ImGui::Button("XG Reset", ImVec2(-1, 30 * main_scale))) {
                        for (int p = 1; p <= 5; ++p) {
                            if (!g_portEnabled[p-1].load()) continue;
                            std::vector<uint8_t> resetMsg = { 0xF0, 0x43, 0x10, 0x4C, 0x00, 0x00, 0x7E, 0x00, 0xF7 };
                            SendToSerial(p, resetMsg);
                        }
                        AddLog(LogSourceType::APP_INFO, "Sent XG System On Reset to all active ports.");
                    }
                    ImGui::PopStyleColor(4);
                }

                // MIDI Service error block moved up

                ImGui::EndTabItem();
            }

            // ΓöÇΓöÇ Debug View Tab ΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇΓöÇ
            if (ImGui::BeginTabItem("Debug View")) {
                static bool showDebugFilters = false;
                ImGui::Spacing();
                
                if (showDebugFilters) {
                    ImGui::Columns(2, "DebugFiltersSplit", true);
                    ImGui::SetColumnWidth(0, 263 * main_scale); // Match connection tab

                    // LEFT COLUMN: Ports (3x3)
                    {
                        // Vertical centering
                        ImGui::Dummy(ImVec2(0, 10 * main_scale));
                        
                        // Horizontal centering offset
                        float hOffset = 20 * main_scale;
                        
                        // Row 1: Out 1, 2, 3
                        ImGui::SetCursorPosX(ImGui::GetCursorPosX() + hOffset);
                        for (int i = 0; i < 3; ++i) {
                            std::string lbl = "Out " + std::to_string(i+1) + "##df";
                            ImGui::Checkbox(lbl.c_str(), &filterOut[i]);
                            if (i < 2) ImGui::SameLine();
                        }
                        ImGui::Spacing();
                        // Row 2: Out 4, 5, In
                        ImGui::SetCursorPosX(ImGui::GetCursorPosX() + hOffset);
                        for (int i = 3; i < 5; ++i) {
                            std::string lbl = "Out " + std::to_string(i+1) + "##df";
                            ImGui::Checkbox(lbl.c_str(), &filterOut[i]);
                            ImGui::SameLine();
                        }
                        ImGui::Checkbox("In##df", &filterIn);
                    }

                    ImGui::NextColumn();

                    // RIGHT COLUMN: Controls (3 rows)
                    // Row 1
                    ImGui::Checkbox("System Info", &filterSys); ImGui::SameLine();
                    ImGui::Checkbox("Filter Sync", &filterSync);
                    
                    // Row 2
                    ImGui::Checkbox("Timestamp", &showTimestamp); ImGui::SameLine();
                    ImGui::Checkbox("Description", &showNames);
                    
                    // Row 3
                    ImGui::Checkbox("Stop Scroll", &stopScroll);

                    ImGui::Columns(1);
                    ImGui::Separator();
                }

                FilterState currentFilter = {filterSys, filterIn, filterSync, {filterOut[0],filterOut[1],filterOut[2],filterOut[3],filterOut[4]}};
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
                            case LogSourceType::APP_OUT_5:   drawCol = colOut[4]; break;
                            default: drawCol = ImVec4(1, 1, 1, 1); break;
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
                if (ImGui::Button(showDebugFilters ? "Hide filters" : "Show filters")) {
                    showDebugFilters = !showDebugFilters;
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

                ImGui::PushStyleVar(ImGuiStyleVar_ItemSpacing, ImVec2(ImGui::GetStyle().ItemSpacing.x, 3.0f * main_scale));

                if (ImGui::Checkbox("Reduce to System Tray on minimize", &g_reduceToTray)) settingsChanged = true;
                if (ImGui::Checkbox("Auto-start Virtual MIDI on launch", &autoStartVirtualMidi)) settingsChanged = true;
                if (ImGui::Checkbox("Auto-reconnect dropped COM port", &autoReconnect)) settingsChanged = true;
                if (ImGui::Checkbox("Quick start (ports may reorder)", &quickInitialization)) settingsChanged = true;
                if (ImGui::Checkbox("Query device info at connection", (bool*)&g_autoGetSynthInfo)) settingsChanged = true;

                if (ImGui::Checkbox("Start with Windows", &startWithWindows)) {
                    SetStartWithWindowsRegistry(startWithWindows);
                    settingsChanged = true;
                }
                if (ImGui::Checkbox("Stay on top", &stayOnTop)) {
                    SetWindowPos(hwnd, stayOnTop ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
                    settingsChanged = true;
                }
                if (ImGui::Checkbox("Start minimized", &startMinimized)) settingsChanged = true;
                if (ImGui::Checkbox("Transmit Sync Data", &g_transmitSync)) settingsChanged = true;

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
                        colOut[4] = ImVec4(0.00f, 0.40f, 0.40f, 1.00f); // Teal
                        colOut[5] = ImVec4(0.00f, 0.00f, 0.60f, 1.00f); // Dark blue
                        colOut[6] = ImVec4(0.60f, 0.00f, 0.30f, 1.00f); // Deep pink
                        colOut[7] = ImVec4(0.00f, 0.00f, 0.00f, 1.00f); // Black
                    } else {
                        ImGui::StyleColorsDark();
                        // Restore Vibrant Dark Mode Defaults
                        colSys  = ImVec4(0.60f, 0.60f, 0.60f, 1.00f);
                        colIn   = ImVec4(0.40f, 0.80f, 1.00f, 1.00f);
                        colOut[0] = ImVec4(1.00f, 0.40f, 0.40f, 1.00f);
                        colOut[1] = ImVec4(1.00f, 0.80f, 0.40f, 1.00f);
                        colOut[2] = ImVec4(0.40f, 1.00f, 0.40f, 1.00f);
                        colOut[3] = ImVec4(0.80f, 0.40f, 1.00f, 1.00f);
                        colOut[4] = ImVec4(0.40f, 1.00f, 1.00f, 1.00f);
                    }
                    settingsChanged = true;
                }
                
                ImGui::PopStyleVar();
                ImGui::Dummy(ImVec2(0, 5.0f));
                
                ImGui::Separator();
                ImGui::Text("Debug View Colors");
                ImGui::Spacing();
                
                ImGui::Columns(4, nullptr, false);
                float colWidth = 120 * main_scale;
                for (int i=0; i<4; i++) ImGui::SetColumnWidth(i, colWidth);
                
                // Row 1: Out 1, 2, 3, 4
                if (ImGui::ColorEdit4("Out 1", (float*)&colOut[0], ImGuiColorEditFlags_NoInputs)) settingsChanged = true;
                ImGui::NextColumn();
                if (ImGui::ColorEdit4("Out 2", (float*)&colOut[1], ImGuiColorEditFlags_NoInputs)) settingsChanged = true;
                ImGui::NextColumn();
                if (ImGui::ColorEdit4("Out 3", (float*)&colOut[2], ImGuiColorEditFlags_NoInputs)) settingsChanged = true;
                ImGui::NextColumn();
                if (ImGui::ColorEdit4("Out 4", (float*)&colOut[3], ImGuiColorEditFlags_NoInputs)) settingsChanged = true;
                ImGui::NextColumn();
                
                // Row 2: Out 5, In, System
                if (ImGui::ColorEdit4("Out 5", (float*)&colOut[4], ImGuiColorEditFlags_NoInputs)) settingsChanged = true;
                ImGui::NextColumn();
                if (ImGui::ColorEdit4("In", (float*)&colIn, ImGuiColorEditFlags_NoInputs)) settingsChanged = true;
                ImGui::NextColumn();
                if (ImGui::ColorEdit4("System", (float*)&colSys, ImGuiColorEditFlags_NoInputs)) settingsChanged = true;
                ImGui::NextColumn();
                
                ImGui::Columns(1);

                if (settingsChanged) {
                    std::string cName = "";
                    if (!comPortsVec.empty() && selectedComPort >= 0 && selectedComPort < (int)comPortsVec.size()) {
                        cName = comPortsVec[selectedComPort];
                    }
                    SaveSettings(cName, baseNameBuf, autoStartVirtualMidi, startWithWindows, autoReconnect, startToTray, colSys, colIn, colOut, stayOnTop, startMinimized, lightUI, quickInitialization);
                }
                
                ImGui::EndTabItem();
            }
            ImGui::EndTabBar();
        }
        ImGui::End();

        // --- DRAW FAKE STARTUP PROGRESS BAR (Premium Overlay) ---
        if ((g_pendingMidiStart.load() || g_isConnecting.load()) && g_wmsConnectStartTime != std::chrono::steady_clock::time_point{}) {
            auto now = std::chrono::steady_clock::now();
            auto ms = std::chrono::duration_cast<std::chrono::milliseconds>(now - g_wmsConnectStartTime).count();
            
            int wmsWaitDelayMs = quickInitialization ? 1500 : 26000;
            float timeScale = 1500.0f / (float)wmsWaitDelayMs;
            float scaledMs = ms * timeScale;

            float progress = 0.0f;
            if (g_isConnecting.load()) {
                progress = 100.0f; // Show full but maybe it's still busy
            } else {
                // User requested steps: 13, 33, 56, 68, 88, 99
                if (scaledMs < 150) progress = (scaledMs / 150.0f) * 13.0f;
                else if (scaledMs < 350) progress = 13.0f;
                else if (scaledMs < 500) progress = 13.0f + ((scaledMs - 350) / 150.0f) * (33.0f - 13.0f);
                else if (scaledMs < 650) progress = 33.0f;
                else if (scaledMs < 800) progress = 33.0f + ((scaledMs - 650) / 150.0f) * (56.0f - 33.0f);
                else if (scaledMs < 950) progress = 56.0f + ((scaledMs - 800) / 150.0f) * (68.0f - 56.0f);
                else if (scaledMs < 1150) progress = 68.0f + ((scaledMs - 950) / 200.0f) * (88.0f - 68.0f);
                else if (scaledMs < 1450) progress = 88.0f + ((scaledMs - 1150) / 300.0f) * (99.0f - 88.0f);
                else progress = 99.0f;
            }

            ImDrawList* drawList = ImGui::GetForegroundDrawList();
            float barHeight = 2.0f * main_scale;
            ImVec2 pMin = ImVec2(0, io.DisplaySize.y - barHeight);
            ImVec2 pMax = ImVec2(io.DisplaySize.x, io.DisplaySize.y);
            drawList->AddRectFilled(pMin, pMax, ImColor(20, 20, 20, 200)); 
            
            ImColor col = g_isConnecting.load() ? ImColor(255, 170, 0, 255) : ImColor(0, 180, 255, 255);
            if (g_isConnecting.load()) {
                float pulse = (sinf((float)ImGui::GetTime() * 8.0f) * 0.5f + 0.5f) * 0.4f + 0.6f;
                col.Value.x *= pulse; col.Value.y *= pulse; col.Value.z *= pulse;
            }
            drawList->AddRectFilled(pMin, ImVec2(io.DisplaySize.x * (progress / 100.0f), pMax.y), col);
        }
    } // Close the UI 'else' block

    // Rendering
    ImGui::Render();
    const float clear_color_with_alpha[4] = { clear_color.x * clear_color.w, clear_color.y * clear_color.w, clear_color.z * clear_color.w, clear_color.w };
    g_pd3dDeviceContext->OMSetRenderTargets(1, &g_mainRenderTargetView, nullptr);
    g_pd3dDeviceContext->ClearRenderTargetView(g_mainRenderTargetView, clear_color_with_alpha);
    ImGui_ImplDX11_RenderDrawData(ImGui::GetDrawData());

    HRESULT hr = g_pSwapChain->Present(1, 0);
    g_SwapChainOccluded = (hr == DXGI_STATUS_OCCLUDED);
} // Close the while loop (started at line 1042)

// Exit Cleanup: Silence all and stop playback
if (isConnected && GetSafeSerialPort() && GetSafeSerialPort()->IsOpen()) {
    for (int p = 1; p <= 5; ++p) {
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

if (g_sharedMidiSession) {
    try { g_sharedMidiSession.Close(); } catch (...) {}
    g_sharedMidiSession = nullptr;
}

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
    // 1-indexed ports: 1-8
    if (g_lastSentPort.load(std::memory_order_relaxed) != portIdx) {
        mux.push_back(0xF5);
        mux.push_back((uint8_t)portIdx);
        g_lastSentPort.store(portIdx, std::memory_order_relaxed);
        
        // Log the multiplexing byte for visibility (Port Select)
        std::string hex = "F5 " + BytesToHex({ (uint8_t)portIdx });
        std::string desc = (portIdx == 5) ? "(Addressing MIDI OUT)" : ("(Addressing Parts " + std::to_string((portIdx - 1) * 16 + 1) + "-" + std::to_string(portIdx * 16) + ")");
        AddLog(LogSourceType::APP_INFO, "[Port Select] " + hex, false, desc);
        
        sp->Write(mux);
        g_bytesSent.fetch_add(mux.size(), std::memory_order_relaxed);
    }

    sp->Write(data);
    g_bytesSent.fetch_add(data.size(), std::memory_order_relaxed);
}
