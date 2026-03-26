# ToHost Bridge

A professional solution for controlling legacy Yamaha synthesizers via their "TO HOST" serial port.

## Overview
This software replaces the aged Yamaha CBX (Serial MIDI) driver for Windows, allowing modern PCs to take full control of Yamaha synthesizers via the serial "TO HOST" port. It is fully compatible with both 16-part instruments and multi-port synthetizers, including:

- **MU Series**: MU2000, MU1000, MU128, MU100, MU90, MU80, MU50, MU15, MU10, MU5
- **QY Series**: QY100, QY70
- **CS/AN Series**: CS6x, CS6R, CS2x, CS1x, AN1x
- **S/VL Series**: S90, S80, S30, VL70-m
- **Hardware**: CBX-K1, CBX-K2, and many TO HOST equipped Clavinova/PSR models.

By using the high-speed "TO HOST" port (38,400 baud), this bridge recreates **4 virtual MIDI ports**, providing up to **64 simultaneous MIDI instruments** (4 ports x 16 channels). This level of control is not possible with a standard single-port USB-to-MIDI interface.

## Key Features
- **64-Channel Control**: Maps 4 virtual MIDI ports to the Yamaha multi-port serial protocol.
- **Solid SysEx Support**: Engineered for high-fidelity System Exclusive communications, handling bulk dumps and complex edits that many standard MIDI interfaces struggle with.
- **Auto-Reconnect**: Automatically polls and reconnects to your COM port if it's disconnected or turned off.
- **Background Operation**: Minimizes to the system tray to stay out of the way while you work.
- **Low Latency**: Optimized for real-time performance with support for latency timer adjustments.
- **Modern UI**: Powered by Dear ImGui for easy configuration and monitoring.
- **SysEx Support**: Fully compatible with Yamaha XG and other complex SysEx commands.
- **Single-Cable Efficiency**: Eliminates the need for separate MIDI IN and OUT cables by providing high-speed bidirectional communication over a single serial connection.

## Hardware & System Requirements
- **Operating System**: Windows 10 / 11 64-bit (Necessarry for the use of Windows MIDI Services).
- **Connection**: A **USB-to-Serial Mini-DIN 8-pin cable**.
   - **FTDI-based cables**: Highly recommended for the best update rate and minimum note drops.
   - **Prolific cables**: Fully tested and supported.
- **HOST SELECT Switch**: Ensure the HOST SELECT switch is set to **PC-2**.

## COM Port Configuration
The software configures the port with the following settings required by Yamaha synths:
- **Baud Rate**: 38,400
- **Data Bits**: 8
- **Parity**: None
- **Stop Bits**: 1
- **Flow Control**: None (8-N-1)

> [!TIP]
> For the best experience, set your USB Serial Port's **Latency Timer** to **1ms** in the Windows Device Manager (under Port Settings -> Advanced).

## User Interface

| Connection Tab | Debug Tab | Settings Tab |
| :---: | :---: | :---: |
| <img src="assets/Connection%20tab.png" width="300" alt="Connection Tab"> | <img src="assets/Debug%20tab.png" width="300" alt="Debug Tab"> | <img src="assets/Settings%20tab.png" width="300" alt="Settings Tab"> |

## Getting Started
1. Install **Windows MIDI Services**:
   - [**Download Windows MIDI Services (x64 Installer)**](https://microsoft.github.io/MIDI/get-latest/#:~:text=Download%20Latest%20x64%20Installer)
   - Follow the on-screen instructions to complete the runtime installation.
2. Download the latest release of **ToHost Bridge** from the GitHub [Releases](https://github.com/Max-Coppola/ToHostBridge/releases) page.
3. Select your COM port in the application and you're ready to go!

---
## License & Distribution
This software is provided for **free** under the **MIT License**.
- It is and will always be open-source.
- **It cannot be sold.**
- You are free to use, modify, and distribute it, provided the original copyright and license notice are included.

---
*Created with ❤️ for the Yamaha XG community.*
