#pragma once
#include <windows.h>
#include <vector>
#include <string>
#include <functional>
#include <mutex>

#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Microsoft.Windows.Devices.Midi2.h>
#include <winrt/Microsoft.Windows.Devices.Midi2.Endpoints.Virtual.h>

using MidiRxCallback = std::function<void(const std::vector<uint8_t>&)>;

// Holds the result of eager device pre-registration.
// Created at MIDI service availability time; consumed by the fast constructor.
struct PreRegisteredVirtualDevice {
    winrt::Microsoft::Windows::Devices::Midi2::Endpoints::Virtual::MidiVirtualDevice device{ nullptr };
    winrt::hstring deviceEndpointId;
    std::wstring portName;
    bool valid = false;
};

class VirtualMidiPort {
public:
    // Original constructor — creates device + connection in one shot (slow on cold service)
    VirtualMidiPort(winrt::Microsoft::Windows::Devices::Midi2::MidiSession const& sharedSession, const std::wstring& portName, MidiRxCallback rxCallback = nullptr);
    // Fast-path constructor — device already pre-registered, only opens the connection
    VirtualMidiPort(winrt::Microsoft::Windows::Devices::Midi2::MidiSession const& sharedSession, const PreRegisteredVirtualDevice& preReg, MidiRxCallback rxCallback = nullptr);
    ~VirtualMidiPort();

    // Eagerly pre-register a virtual device WITHOUT opening a connection.
    // Call this right after MidiSession::Create so PnP settles while the app loads.
    static PreRegisteredVirtualDevice PreRegister(const std::wstring& portName);
    static void RemoveStalePorts();

    bool IsValid() const;
    winrt::hstring GetDeviceId() const;
    void SendMidi(const std::vector<uint8_t>& data);

private:
    void OnMessageReceived(
        winrt::Microsoft::Windows::Devices::Midi2::IMidiMessageReceivedEventSource const& sender,
        winrt::Microsoft::Windows::Devices::Midi2::MidiMessageReceivedEventArgs const& args);

    MidiRxCallback m_rxCallback;
    std::wstring m_portName;
    std::atomic<bool> m_isClosing{ false };
    mutable std::mutex m_portMutex;
    
    winrt::Microsoft::Windows::Devices::Midi2::Endpoints::Virtual::MidiVirtualDevice m_virtualDevice{ nullptr };
    winrt::Microsoft::Windows::Devices::Midi2::MidiSession m_session{ nullptr };
    winrt::Microsoft::Windows::Devices::Midi2::MidiEndpointConnection m_connection{ nullptr };
    winrt::event_token m_messageReceivedToken{};
    std::vector<uint8_t> m_sysExAccumulator;
};
