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

class VirtualMidiPort {
public:
    VirtualMidiPort(const std::wstring& portName, MidiRxCallback rxCallback = nullptr);
    ~VirtualMidiPort();

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
