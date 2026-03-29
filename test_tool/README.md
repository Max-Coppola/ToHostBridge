# ToHost Bridge Stress Test Tool

This tool is designed to stress-test the robustness of the ToHost Bridge by flooding all 4 virtual MIDI ports with high-density traffic.

## 🚀 How to use
1. Start **ToHost Bridge** and connect to your COM port.
2. Ensure at least some Out ports are enabled in the "Connection" tab.
3. Run `target_stress_test.exe` in this folder.
4. Watch the **Debug View** in the main app — it should be flooded with Blue/Yellow/Green/Purple traffic!
5. Check your **MU2000** display — it should show "P1 STRESS TEST", "P2 STRESS TEST", etc.

## 🛠️ What it does
- **64-Part Density**: Concurrent threads for "ToHost Bridge Out 1" through "Out 4".
- **Hardware Safe**: Uses standard XG Display and Parameter changes. No flash writes.
- **Traffic Components**:
    - Polyphonic melodies (Note-On/Off).
    - Continuous Pitch Bend and Modulation.
    - Random Channel Pressure (Aftertouch).
    - Periodic XG Parameter updates.

## 🛑 How to stop
- Press **Ctrl+C** in the console window. It will send "All Notes Off" and close cleanly.

---
*Note: This tool and folder are ignored by Git and will never be pushed to GitHub.*
