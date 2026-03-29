cmake_minimum_required(VERSION 3.10)
project(ToHostStressTest)

set(CMAKE_CXX_STANDARD 17)

add_executable(stress_test stress_test.cpp)
target_link_libraries(stress_test winmm.lib)

# Ensure the executable is named stress_test.exe
set_target_properties(stress_test PROPERTIES OUTPUT_NAME "stress_test")
