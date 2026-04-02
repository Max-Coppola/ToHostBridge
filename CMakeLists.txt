cmake_minimum_required(VERSION 3.10)
project(ToHostBridge VERSION 1.2.2 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED True)

include(FetchContent)
FetchContent_Declare(
    imgui
    URL https://github.com/ocornut/imgui/archive/refs/heads/master.zip
)
FetchContent_MakeAvailable(imgui)

set(IMGUI_SRCS
    ${imgui_SOURCE_DIR}/imgui.cpp
    ${imgui_SOURCE_DIR}/imgui_draw.cpp
    ${imgui_SOURCE_DIR}/imgui_tables.cpp
    ${imgui_SOURCE_DIR}/imgui_widgets.cpp
    ${imgui_SOURCE_DIR}/backends/imgui_impl_win32.cpp
    ${imgui_SOURCE_DIR}/backends/imgui_impl_dx11.cpp
)

add_executable(ToHostBridge WIN32
    main.cpp
    SerialPort.h SerialPort.cpp
    VirtualMidiPort.h VirtualMidiPort.cpp
    resources.rc
    ${IMGUI_SRCS}
)

file(GLOB MIDI2_NATIVE_INC "${CMAKE_CURRENT_SOURCE_DIR}/packages/Microsoft.Windows.Devices.Midi2.*/build/native/include")

target_include_directories(ToHostBridge PRIVATE 
    ${CMAKE_CURRENT_SOURCE_DIR}/include 
    ${MIDI2_NATIVE_INC}
    ${imgui_SOURCE_DIR}
    ${imgui_SOURCE_DIR}/backends
)

target_link_libraries(ToHostBridge 
    PRIVATE winmm
    PRIVATE windowsapp
    PRIVATE d3d11
    PRIVATE d3dcompiler
    PRIVATE shell32
)
