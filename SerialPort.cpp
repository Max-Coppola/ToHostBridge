#include "SerialPort.h"
#include <iostream>

SerialPort::SerialPort(const std::string& portName, SerialRxCallback rxCallback)
    : m_portName(portName), m_rxCallback(rxCallback), m_isOpen(false), m_running(false), m_hComm(INVALID_HANDLE_VALUE)
{
    std::string fullPortName = "\\\\.\\" + portName;
    m_hComm = CreateFileA(fullPortName.c_str(),
        GENERIC_READ | GENERIC_WRITE,
        0,
        NULL,
        OPEN_EXISTING,
        FILE_FLAG_OVERLAPPED,
        NULL);

    if (m_hComm == INVALID_HANDLE_VALUE) {
        std::cerr << "COM Port Access Denied or Port Not Found: " << portName << std::endl;
        return;
    }

    DCB dcbSerialParams = { 0 };
    dcbSerialParams.DCBlength = sizeof(dcbSerialParams);
    
    if (!GetCommState(m_hComm, &dcbSerialParams)) {
        std::cerr << "Error getting COM state." << std::endl;
        CloseHandle(m_hComm);
        m_hComm = INVALID_HANDLE_VALUE;
        return;
    }

    dcbSerialParams.BaudRate = 38400;
    dcbSerialParams.ByteSize = 8;
    dcbSerialParams.StopBits = ONESTOPBIT;
    dcbSerialParams.Parity   = NOPARITY;

    if (!SetCommState(m_hComm, &dcbSerialParams)) {
        std::cerr << "Error setting COM state (38400 8N1)." << std::endl;
        CloseHandle(m_hComm);
        m_hComm = INVALID_HANDLE_VALUE;
        return;
    }

    COMMTIMEOUTS timeouts = { 0 };
    timeouts.ReadIntervalTimeout = MAXDWORD;
    timeouts.ReadTotalTimeoutConstant = 0;
    timeouts.ReadTotalTimeoutMultiplier = 0;
    timeouts.WriteTotalTimeoutConstant = 50;
    timeouts.WriteTotalTimeoutMultiplier = 10;
    
    if (!SetCommTimeouts(m_hComm, &timeouts)) {
        std::cerr << "Error setting timeouts." << std::endl;
        CloseHandle(m_hComm);
        m_hComm = INVALID_HANDLE_VALUE;
        return;
    }

    m_isOpen = true;
    m_running = true;
    m_readThread = std::thread(&SerialPort::ReadThread, this);
}

SerialPort::~SerialPort() {
    m_running = false;
    if (m_readThread.joinable()) {
        m_readThread.join();
    }
    if (m_hComm != INVALID_HANDLE_VALUE) {
        CloseHandle(m_hComm);
    }
}

bool SerialPort::IsOpen() const {
    return m_isOpen;
}

void SerialPort::Write(const std::vector<uint8_t>& data) {
    if (!m_isOpen || data.empty()) return;

    std::lock_guard<std::mutex> lock(m_writeMutex);

    OVERLAPPED osWrite = { 0 };
    osWrite.hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
    if (osWrite.hEvent == NULL) return;

    DWORD dwWritten = 0;
    if (!WriteFile(m_hComm, data.data(), (DWORD)data.size(), &dwWritten, &osWrite)) {
        if (GetLastError() != ERROR_IO_PENDING) {
            std::cerr << "WriteFile failed!" << std::endl;
        } else {
            GetOverlappedResult(m_hComm, &osWrite, &dwWritten, TRUE);
        }
    }
    CloseHandle(osWrite.hEvent);
}

void SerialPort::ReadThread() {
    OVERLAPPED osReader = { 0 };
    osReader.hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
    if (osReader.hEvent == NULL) return;

    std::vector<uint8_t> buffer(4096);
    DWORD dwRead;

    while (m_running) {
        if (!ReadFile(m_hComm, buffer.data(), (DWORD)buffer.size(), &dwRead, &osReader)) {
            if (GetLastError() == ERROR_IO_PENDING) {
                if (GetOverlappedResult(m_hComm, &osReader, &dwRead, TRUE)) {
                    if (dwRead > 0) {
                        std::vector<uint8_t> data(buffer.begin(), buffer.begin() + dwRead);
                        if (m_rxCallback) m_rxCallback(data);
                    }
                }
            } else {
                std::cerr << "Read error or port disconnected." << std::endl;
                break;
            }
        } else {
            if (dwRead > 0) {
                std::vector<uint8_t> data(buffer.begin(), buffer.begin() + dwRead);
                if (m_rxCallback) m_rxCallback(data);
            }
        }
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }
    CloseHandle(osReader.hEvent);
}
