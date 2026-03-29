#include <windows.h>
#include <mmsystem.h>
#include <iostream>
#include <vector>
#include <string>
#include <thread>
#include <atomic>
#include <cmath>
#include <algorithm>
#include <conio.h>

#pragma comment(lib, "winmm.lib")

std::atomic<bool> g_appRunning{true};
std::atomic<bool> g_portBusy[5]{false, false, false, false, false}; // Logical ports 1-4 + In(0)
std::atomic<bool> g_portEnabled[5]{false, false, false, false, false};
std::atomic<uint64_t> g_portCount[5]{0, 0, 0, 0, 0};

void SendSysEx(HMIDIOUT hmo, const std::vector<uint8_t>& data) {
    MIDIHDR hdr = {};
    hdr.lpData = (LPSTR)data.data();
    hdr.dwBufferLength = (DWORD)data.size();
    midiOutPrepareHeader(hmo, &hdr, sizeof(hdr));
    midiOutLongMsg(hmo, &hdr, sizeof(hdr));
    int timeout = 100;
    while (!(hdr.dwFlags & MHDR_DONE) && g_appRunning && timeout-- > 0) Sleep(1);
    midiOutUnprepareHeader(hmo, &hdr, sizeof(hdr));
}

void CALLBACK MidiInProc(HMIDIIN hMidiIn, UINT wMsg, DWORD_PTR dwInstance, DWORD_PTR dwParam1, DWORD_PTR dwParam2) {
    if (wMsg == MIM_DATA || wMsg == MIM_LONGDATA) {
        g_portCount[0].fetch_add(1, std::memory_order_relaxed);
    }
}

void InThread(std::string portName) {
    HMIDIIN hmi = nullptr;
    UINT numDevs = midiInGetNumDevs();
    UINT devID = (UINT)-1;
    for (UINT i = 0; i < numDevs; ++i) {
        MIDIINCAPSA caps;
        if (midiInGetDevCapsA(i, &caps, sizeof(caps)) == MMSYSERR_NOERROR) {
            if (std::string(caps.szPname).find(portName) != std::string::npos) { devID = i; break; }
        }
    }
    if (devID == (UINT)-1 || midiInOpen(&hmi, devID, (DWORD_PTR)MidiInProc, 0, CALLBACK_FUNCTION) != MMSYSERR_NOERROR) return;
    midiInStart(hmi);
    while (g_appRunning) {
        if (!g_portEnabled[0]) { midiInStop(hmi); while(!g_portEnabled[0] && g_appRunning) Sleep(100); midiInStart(hmi); }
        Sleep(100);
    }
    midiInStop(hmi);
    midiInClose(hmi);
}

void OutThread(int portIdx, std::string portName) {
    HMIDIOUT hmo = nullptr;
    UINT numDevs = midiOutGetNumDevs();
    UINT devID = (UINT)-1;
    for (UINT i = 0; i < numDevs; ++i) {
        MIDIOUTCAPSA caps;
        if (midiOutGetDevCapsA(i, &caps, sizeof(caps)) == MMSYSERR_NOERROR) {
            if (std::string(caps.szPname).find(portName) != std::string::npos) { devID = i; break; }
        }
    }
    if (devID == (UINT)-1 || midiOutOpen(&hmo, devID, 0, 0, CALLBACK_NULL) != MMSYSERR_NOERROR) return;
    
    SendSysEx(hmo, {0xF0, 0x43, 0x10, 0x4C, 0x00, 0x00, 0x7E, 0x00, 0xF7});
    int step = 0;
    while (g_appRunning) {
        if (!g_portEnabled[portIdx]) { Sleep(100); continue; }
        
        // Staggered updates: only update 4 channels per frame to be "reasonable"
        for (int i = 0; i < 4; ++i) {
            int ch = (step * 4 + i) % 16;
            
            // Note On
            uint8_t note = 48 + (step % 12) + (ch);
            midiOutShortMsg(hmo, 0x90 | ch | (note << 8) | (80 << 16));
            
            // Occasional Pitch Bend
            int pb = (int)(8192 + 2000 * sin(step * 0.2));
            midiOutShortMsg(hmo, 0xE0 | ch | ((pb & 0x7F) << 8) | (((pb >> 7) & 0x7F) << 16));
            g_portCount[portIdx].fetch_add(2, std::memory_order_relaxed);
            
            // Note Off (immediately or after a short delay in the loop)
            // For a "reasonable" test, let's just keep notes short
            std::thread([hmo, ch, note]() {
                Sleep(80);
                midiOutShortMsg(hmo, 0x80 | ch | (note << 8) | (0 << 16));
            }).detach();
        }
        
        step++;
        Sleep(100); // 10Hz total update rate
    }
    midiOutReset(hmo);
    midiOutClose(hmo);
}

void DrawUI() {
    system("cls");
    std::cout << "====================================================\n";
    std::cout << "ToHost Bridge INTERACTIVE Stress Tool\n";
    std::cout << "====================================================\n\n";
    std::cout << "PORT          STATUS      MSG COUNT\n";
    std::cout << "-----------------------------------\n";
    auto printPort = [](int i, std::string name) {
        std::cout << (i==0?"IN ":"OUT ") << (i==0?1:i) << " [" << name << "]  " 
                  << (g_portEnabled[i] ? "RUNNING" : "PAUSED ") << "   " << g_portCount[i].load() << "\n";
    };
    printPort(1, "CBX Out 1"); printPort(2, "CBX Out 2");
    printPort(3, "CBX Out 3"); printPort(4, "CBX Out 4");
    printPort(0, "CBX In   ");
    std::cout << "\n-----------------------------------\n";
    std::cout << "CONTROLS: [1-4] Toggle Out 1-4 | [I] Toggle In | [Q] Quit\n";
}

int main() {
    std::vector<std::string> outs = {"CBX Out 1", "CBX Out 2", "CBX Out 3", "CBX Out 4"};
    std::vector<std::thread> threads;
    for(int i=0; i<4; ++i) threads.emplace_back(OutThread, i+1, outs[i]);
    threads.emplace_back(InThread, "CBX In");

    while (g_appRunning) {
        DrawUI();
        if (_kbhit()) {
            int c = _getch();
            if (c >= '1' && c <= '4') g_portEnabled[c-'0'] = !g_portEnabled[c-'0'];
            else if (c == 'i' || c == 'I') g_portEnabled[0] = !g_portEnabled[0];
            else if (c == 'q' || c == 'Q') g_appRunning = false;
        }
        Sleep(200);
    }
    for (auto& t : threads) if (t.joinable()) t.join();
    return 0;
}
