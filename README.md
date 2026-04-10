# ToHost Bridge

A professional solution for controlling legacy Yamaha synthesizers via their "TO HOST" serial port.

## Overview
This software replaces the 32-bit Yamaha CBX (Serial MIDI) driver for Windows, allowing modern PCs to take full control of Yamaha synthesizers via the serial "TO HOST" port. It is fully compatible with both 16-part instruments and multi-port synths, including:

- **MU Series**: MU2000, MU1000, MU128, MU100, MU90, MU80, MU50, MU15, MU10, MU5
- **QY Series**: QY100, QY70
- **CS/AN Series**: CS6x, CS6R, CS2x, CS1x, AN1x
- **S/VL Series**: S90, S80, S30, VL70-m
- **Hardware**: CBX-K1, CBX-K2, and many TO HOST equipped Clavinova/PSR models.

By using the high-speed "TO HOST" port (38,400 baud), this bridge recreates **5 virtual MIDI ports**, providing up to **64 + 16 simultaneous MIDI instruments** (5 ports x 16 channels). This level of control is not possible with a standard single-port USB-to-MIDI interface.

## Visual Interface
![Connection Tab](assets/Connection%20tab.png)
![Debug Tab](assets/Debug%20tab.png)
![Settings Tab](assets/Settings%20tab.png)

## Key Features
- **128-Channel Control**: Maps 5 virtual MIDI ports to the Yamaha multi-port serial protocol.
- **Solid SysEx Support**: Engineered for high-fidelity System Exclusive communications, handling bulk dumps and complex edits that many standard MIDI interfaces struggle with.
- **Auto-Reconnect**: Automatically polls and reconnects to your COM port if it's disconnected or turned off.
- **Background Operation**: Minimizes to the system tray to stay out of the way while you work.
- **Low Latency**: Optimized for real-time performance with support for latency timer adjustments.
- **Modern UI**: Powered by Dear ImGui for easy configuration and monitoring.
- **SysEx Support**: Fully compatible with Yamaha XG and other complex SysEx commands.

## Hardware Requirements
To connect your PC to the synth's "TO HOST" port, you need a **USB-to-Serial Mini-DIN 8-pin cable**.
- **FTDI-based cables**: Highly recommended for the best update rate and minimum note drops.
- **Prolific cables**: Tested and supported.

## COM Port Configuration
The software configures the port with the following settings required by Yamaha synths:
- **Baud Rate**: 38,400
- **Data Bits**: 8
- **Parity**: None
- **Stop Bits**: 1
- **Flow Control**: None (8-N-1)

> [!TIP]
> For the best experience, set your USB Serial Port's **Latency Timer** to **1ms** in the Windows Device Manager (under Port Settings -> Advanced).


## Getting Started
1. Install **Windows MIDI Services**:
   - Visit the [Windows MIDI Services Download Page](https://microsoft.github.io/MIDI/get-latest/#:~:text=Download%20Latest%20x64%20Installer).
   - Navigate to "Get the latest Windows MIDI Services runtime and tools".
   - Download and run the latest x64 installer.
2. Download **ToHost Bridge** from the [Microsoft Store](https://apps.microsoft.com/detail/9plvb2ngp1xs) or get the latest installer from the [GitHub Releases](https://github.com/Max-Coppola/ToHostBridge/releases) page.
3. Select your COM port in the application and you're ready to go!


## Port Mapping (Multiplexing)
ToHost Bridge enables **5 virtual ports** (64 + 16 MIDI channels) over a single serial connection:
- **CBX Out 1**: Parts 1-16
- **CBX Out 2**: Parts 17-32
- **CBX Out 3**: Parts 33-48
- **CBX Out 4**: Parts 49-64
- **CBX Out 5**: Physical MIDI OUT A (Channels 1-16)

## License
MIT License - Open for all. Created by Max Coppola.

---
*Created with ❤️ for the Yamaha XG community.*
