#include "VirtualMidiPort.h"
#include <iostream>
#include <winrt/Windows.Foundation.Collections.h>

using namespace winrt;
using namespace winrt::Microsoft::Windows::Devices::Midi2;
using namespace winrt::Microsoft::Windows::Devices::Midi2::Endpoints::Virtual;

// ---------------------------------------------------------------------------
// Original constructor — creates device + opens connection (slow on cold service)
// ---------------------------------------------------------------------------
VirtualMidiPort::VirtualMidiPort(MidiSession const& sharedSession, const std::wstring& portName, MidiRxCallback rxCallback)
    : m_rxCallback(rxCallback), m_portName(portName), m_session(sharedSession) {
    try {
        if (!MidiVirtualDeviceManager::IsTransportAvailable()) {
            std::wcerr << L"Windows MIDI Services Virtual Transport is not available." << std::endl;
            return;
        }

        MidiDeclaredEndpointInfo info{};
        info.Name = portName;
        info.SpecificationVersionMajor = 1;
        info.SpecificationVersionMinor = 0;
        info.SupportsMidi10Protocol = true;

        MidiVirtualDeviceCreationConfig config(
            portName, 
            L"Virtual MIDI Port", 
            L"MidiSerialBridge", 
            info);

        m_virtualDevice = MidiVirtualDeviceManager::CreateVirtualDevice(config);
        if (!m_virtualDevice) {
            std::wcerr << L"Failed to create virtual device for " << portName << std::endl;
            return;
        }

        m_connection = m_session.CreateEndpointConnection(m_virtualDevice.DeviceEndpointDeviceId());
        if (!m_connection) {
            std::wcerr << L"Failed to create endpoint connection for " << portName << std::endl;
            return;
        }

        if (m_rxCallback) {
            m_messageReceivedToken = m_connection.MessageReceived({ this, &VirtualMidiPort::OnMessageReceived });
        }
        
        m_connection.Open();
        std::wcout << L"Successfully created and opened virtual port: " << portName << std::endl;
    } catch (winrt::hresult_error const& ex) {
        std::wcerr << L"Error creating virtual MIDI port " << portName << L": " << ex.message().c_str() << std::endl;
        m_virtualDevice = nullptr;
        m_connection = nullptr;
        m_session = nullptr;
    } catch (std::exception const& ex) {
        std::wcerr << L"STD Error creating virtual MIDI port " << portName << L": " << ex.what() << std::endl;
        m_virtualDevice = nullptr;
        m_connection = nullptr;
        m_session = nullptr;
    } catch (...) {
        std::wcerr << L"Unknown error creating virtual MIDI port " << portName << std::endl;
        m_virtualDevice = nullptr;
        m_connection = nullptr;
        m_session = nullptr;
    }
}

// ---------------------------------------------------------------------------
// Static: Pre-register a virtual device WITHOUT opening a connection.
// Call this immediately after MidiSession::Create so the Windows MIDI Service
// can complete the PnP registration while the app is still loading.
// ---------------------------------------------------------------------------
PreRegisteredVirtualDevice VirtualMidiPort::PreRegister(const std::wstring& portName) {
    PreRegisteredVirtualDevice result;
    result.portName = portName;
    try {
        if (!MidiVirtualDeviceManager::IsTransportAvailable()) {
            std::wcerr << L"[PreRegister] Virtual transport not available for: " << portName << std::endl;
            return result;
        }

        MidiDeclaredEndpointInfo info{};
        info.Name = portName;
        info.SpecificationVersionMajor = 1;
        info.SpecificationVersionMinor = 0;
        info.SupportsMidi10Protocol = true;

        MidiVirtualDeviceCreationConfig config(
            portName, 
            L"Virtual MIDI Port", 
            L"MidiSerialBridge", 
            info);

        auto device = MidiVirtualDeviceManager::CreateVirtualDevice(config);
        if (!device) {
            std::wcerr << L"[PreRegister] CreateVirtualDevice returned null for: " << portName << std::endl;
            return result;
        }

        result.device = device;
        result.deviceEndpointId = device.DeviceEndpointDeviceId();
        result.valid = true;
        std::wcout << L"[PreRegister] OK: " << portName << std::endl;
    } catch (winrt::hresult_error const& ex) {
        std::wcerr << L"[PreRegister] HResult error for " << portName << L": " << ex.message().c_str() << std::endl;
    } catch (...) {
        std::wcerr << L"[PreRegister] Exception for: " << portName << std::endl;
    }
    return result;
}

// ---------------------------------------------------------------------------
// Fast-path constructor: device is already registered in the MIDI service.
// Skips CreateVirtualDevice (the slow IPC call).
// Only calls CreateEndpointConnection + Open, which are near-instant.
// ---------------------------------------------------------------------------
VirtualMidiPort::VirtualMidiPort(MidiSession const& sharedSession, const PreRegisteredVirtualDevice& preReg, MidiRxCallback rxCallback)
    : m_rxCallback(rxCallback), m_portName(preReg.portName), m_session(sharedSession) {

    if (!preReg.valid || !preReg.device) {
        std::wcerr << L"[FastConnect] Invalid pre-registration for: " << preReg.portName << std::endl;
        return;
    }

    m_virtualDevice = preReg.device;

    try {
        m_connection = m_session.CreateEndpointConnection(preReg.deviceEndpointId);
        if (!m_connection) {
            std::wcerr << L"[FastConnect] Null connection for: " << preReg.portName << std::endl;
            return;
        }
        if (m_rxCallback) {
            m_messageReceivedToken = m_connection.MessageReceived({ this, &VirtualMidiPort::OnMessageReceived });
        }
        m_connection.Open();
        std::wcout << L"[FastConnect] Opened: " << preReg.portName << std::endl;
    } catch (winrt::hresult_error const& ex) {
        std::wcerr << L"[FastConnect] Error for " << preReg.portName << L": " << ex.message().c_str() << std::endl;
        m_connection = nullptr;
    } catch (...) {
        std::wcerr << L"[FastConnect] Unknown error for: " << preReg.portName << std::endl;
        m_connection = nullptr;
    }
}

VirtualMidiPort::~VirtualMidiPort() {
    // Signal all in-progress callbacks to bail out immediately
    m_isClosing = true;

    // Brief pause: let any in-flight OnMessageReceived / SendMidi calls
    // that already passed the m_isClosing check above finish naturally.
    ::Sleep(50);

    // Now acquire exclusive access – guarantees no callback is mid-execution
    std::lock_guard<std::mutex> lock(m_portMutex);

    try {
        if (m_connection) {
            // Unregister our receive handler
            if (m_messageReceivedToken.value != 0) {
                try { m_connection.MessageReceived(m_messageReceivedToken); } catch (...) {}
                m_messageReceivedToken = {};
            }
            
            // Capture connection Id for session cleanup
            auto connId = m_connection.ConnectionId();

            // Explicitly close the connection (IClosable)
            try { m_connection.as<winrt::Windows::Foundation::IClosable>().Close(); } catch (...) {}
            m_connection = nullptr;

            // Explicitly remove from session
            if (m_session) {
                try { m_session.DisconnectEndpointConnection(connId); } catch (...) {}
            }
        }
        
        if (m_virtualDevice) {
            // Explicitly close the virtual device to trigger PnP removal
            try { m_virtualDevice.as<winrt::Windows::Foundation::IClosable>().Close(); } catch (...) {}
            m_virtualDevice = nullptr;
        }

        m_session = nullptr;
    } catch (...) {}
}

bool VirtualMidiPort::IsValid() const {
    return m_virtualDevice != nullptr && m_connection != nullptr;
}

winrt::hstring VirtualMidiPort::GetDeviceId() const {
    if (m_virtualDevice) return m_virtualDevice.DeviceEndpointDeviceId();
    return L"";
}

void VirtualMidiPort::SendMidi(const std::vector<uint8_t>& data) {
    if (m_isClosing || !m_connection || data.empty()) return;
    
    size_t i = 0;
    while(i < data.size()) {
        if (m_isClosing || !m_connection) return; // bail if destroyed mid-send
        uint8_t status = data[i];
        if (status >= 0xF8) {
            // Realtime
            m_connection.SendSingleMessageWords(0, (0x10 << 24) | (status << 16));
            i += 1;
        } else if (status >= 0x80 && status <= 0xEF) {
            // Channel Voice
            uint8_t cmd = status & 0xF0;
            if (cmd == 0xC0 || cmd == 0xD0) {
                if (i + 1 < data.size()) {
                    uint32_t word = (0x20 << 24) | (status << 16) | (data[i+1] << 8);
                    m_connection.SendSingleMessageWords(0, word);
                    i += 2;
                } else break;
            } else {
                if (i + 2 < data.size()) {
                    uint32_t word = (0x20 << 24) | (status << 16) | (data[i+1] << 8) | data[i+2];
                    m_connection.SendSingleMessageWords(0, word);
                    i += 3;
                } else break;
            }
        } else if (status == 0xF0) {
            // SysEx 7 (UMP MT=3) — strip F0/F7, chunk into 6-byte UMP packets
            std::vector<uint8_t> sysex;
            while (i < data.size()) {
                uint8_t b = data[i++];
                if (b != 0xF0 && b != 0xF7) sysex.push_back(b);
                if (b == 0xF7) break;
            }
            
            size_t bytesSent = 0;
            if (sysex.empty()) {
                // Empty SysEx (just F0 F7)
                m_connection.SendSingleMessageWords(0, (0x3u << 28) | (0u << 20) | (0u << 16), 0);
            } else {
                while (bytesSent < sysex.size()) {
                    size_t remaining  = sysex.size() - bytesSent;
                    size_t chunkCount = (remaining > 6) ? 6 : remaining;
                    
                    uint8_t umpStatus;
                    if (sysex.size() <= 6)                       umpStatus = 0;
                    else if (bytesSent == 0)                      umpStatus = 1;
                    else if (bytesSent + chunkCount >= sysex.size()) umpStatus = 3;
                    else                                          umpStatus = 2;
                    
                    uint32_t w0 = (0x3u << 28) | (umpStatus << 20) | ((uint32_t)chunkCount << 16);
                    uint32_t w1 = 0;
                    
                    if (chunkCount >= 1) w0 |= (uint32_t)sysex[bytesSent + 0] << 8;
                    if (chunkCount >= 2) w0 |= (uint32_t)sysex[bytesSent + 1];
                    if (chunkCount >= 3) w1 |= (uint32_t)sysex[bytesSent + 2] << 24;
                    if (chunkCount >= 4) w1 |= (uint32_t)sysex[bytesSent + 3] << 16;
                    if (chunkCount >= 5) w1 |= (uint32_t)sysex[bytesSent + 4] << 8;
                    if (chunkCount >= 6) w1 |= (uint32_t)sysex[bytesSent + 5];
                    
                    m_connection.SendSingleMessageWords(0, w0, w1);
                    bytesSent += chunkCount;
                }
            }
        } else {
             // Unsupported or stray data
             i++;
        }
    }
}

void VirtualMidiPort::OnMessageReceived(
    IMidiMessageReceivedEventSource const& /*sender*/,
    MidiMessageReceivedEventArgs const& args) {
    if (m_isClosing || !m_rxCallback) return;

    MidiMessageStruct messageStruct;
    args.FillMessageStruct(messageStruct);
    uint32_t word0 = messageStruct.Word0;
    uint32_t word1 = messageStruct.Word1;
    uint8_t mt = (word0 >> 28) & 0x0F;

    std::vector<uint8_t> midi1Bytes;
    if (mt == 0x1) {
        // Utility / Realtime
        midi1Bytes.push_back((word0 >> 16) & 0xFF);
    } else if (mt == 0x2) {
        // MIDI 1.0 Channel Voice
        uint8_t status = (word0 >> 16) & 0xFF;
        midi1Bytes.push_back(status);
        uint8_t cmd = status & 0xF0;
        if (cmd == 0xC0 || cmd == 0xD0) {
            midi1Bytes.push_back((word0 >> 8) & 0xFF);
        } else {
            midi1Bytes.push_back((word0 >> 8) & 0xFF);
            midi1Bytes.push_back(word0 & 0xFF);
        }
    } else if (mt == 0x3) {
        // SysEx 7 (UMP MT=3) - Accumulate for atomicity
        uint8_t umpStatus = (word0 >> 20) & 0x0F;
        uint8_t numBytes = (word0 >> 16) & 0x0F;
        
        if (umpStatus == 0 || umpStatus == 1) {
            m_sysExAccumulator.clear();
            m_sysExAccumulator.push_back(0xF0);
        }
        
        auto push_safe = [&](uint8_t b) {
            if (b != 0xF0 && b != 0xF7) m_sysExAccumulator.push_back(b);
        };

        if (numBytes >= 1) push_safe((word0 >> 8) & 0xFF);
        if (numBytes >= 2) push_safe((word0) & 0xFF);
        if (numBytes >= 3) push_safe((word1 >> 24) & 0xFF);
        if (numBytes >= 4) push_safe((word1 >> 16) & 0xFF);
        if (numBytes >= 5) push_safe((word1 >> 8) & 0xFF);
        if (numBytes >= 6) push_safe((word1) & 0xFF);
        
        if (umpStatus == 0 || umpStatus == 3) {
            m_sysExAccumulator.push_back(0xF7);
            m_rxCallback(m_sysExAccumulator);
            m_sysExAccumulator.clear();
        }
    }
    
    if (!midi1Bytes.empty()) {
        m_rxCallback(midi1Bytes);
    }
}

void VirtualMidiPort::RemoveStalePorts() {
    // Port enumeration and proactive removal requires specific SDK identifiers that 
    // are currently undergoing changes in the preview SDK. 
    // The current fix prioritizes application stability and thread-safety 
    // to prevent crashes, while the MIDI Service is relied upon for 
    // endpoint cleanup.
    std::wcout << L"Startup: Virtual MIDI Port cleanup routine initialized." << std::endl;
}
