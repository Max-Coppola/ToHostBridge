import os

file_path = 'main.cpp'
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Declarations
old_decl = "std::unique_ptr<SerialPort> g_serialPort;"
new_decl = """std::shared_ptr<SerialPort> g_serialPort;
std::mutex g_serialPortMutex;

std::shared_ptr<SerialPort> GetSafeSerialPort() {
    std::lock_guard<std::mutex> lock(g_serialPortMutex);
    return g_serialPort;
}

void SetSafeSerialPort(std::shared_ptr<SerialPort> sp) {
    std::lock_guard<std::mutex> lock(g_serialPortMutex);
    g_serialPort = sp;
}

void ResetSafeSerialPort() {
    std::lock_guard<std::mutex> lock(g_serialPortMutex);
    if (g_serialPort) g_serialPort.reset();
}"""
content = content.replace(old_decl, new_decl)

# 2. Resets
content = content.replace("g_serialPort.reset();", "ResetSafeSerialPort();")
content = content.replace("g_serialPort->IsOpen()", "GetSafeSerialPort()->IsOpen()")

# 3. Write usages - Need to be safe. We replace direct g_serialPort calls with safe variable.
# Find instances like: if (g_serialPort && g_serialPort->IsOpen())
content = content.replace("g_serialPort && GetSafeSerialPort()->IsOpen()", "GetSafeSerialPort() && GetSafeSerialPort()->IsOpen()")
content = content.replace("if (!g_serialPort || !GetSafeSerialPort()->IsOpen() || data.empty()) return;", "auto sp = GetSafeSerialPort();\n    if (!sp || !sp->IsOpen() || data.empty()) return;")
content = content.replace("g_serialPort->Write(mux);", "sp->Write(mux);")
content = content.replace("g_serialPort->Write(data);", "sp->Write(data);")

# 4. Specific Write usages in the rest of the code:
content = content.replace("g_serialPort->Write(stdId);", "if(auto sp2 = GetSafeSerialPort()) sp2->Write(stdId);")

# 5. Initialization
old_init = "g_serialPort = std::make_unique<SerialPort>(comName, [colIn, parser](const std::vector<uint8_t>& data) {"
new_init = "SetSafeSerialPort(std::make_shared<SerialPort>(comName, [colIn, parser](const std::vector<uint8_t>& data) {"
content = content.replace(old_init, new_init)

# We must close SetSafeSerialPort properly.
# The lambda ends with:
#         });
#         
#         std::stringstream connMsg;

old_end = """        });
        
        std::stringstream connMsg;"""
new_end = """        }));
        
        std::stringstream connMsg;"""
content = content.replace(old_end, new_end)


with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Done")
