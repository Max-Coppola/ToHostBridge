#pragma once
#include <windows.h>
#include <string>
#include <vector>
#include <thread>
#include <atomic>
#include <mutex>
#include <functional>

using SerialRxCallback = std::function<void(const std::vector<uint8_t>&)>;

class SerialPort {
public:
    SerialPort(const std::string& portName, SerialRxCallback rxCallback);
    ~SerialPort();

    bool IsOpen() const;
    void Write(const std::vector<uint8_t>& data);

private:
    void ReadThread();

    HANDLE m_hComm;
    bool m_isOpen;
    std::string m_portName;
    SerialRxCallback m_rxCallback;
    
    std::thread m_readThread;
    std::atomic<bool> m_running;
    
    std::mutex m_writeMutex;
};
