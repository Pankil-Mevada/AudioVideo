cmake_minimum_required(VERSION 3.20)
project(systemc_tlm_virtual_ecu LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(SystemCLanguage CONFIG REQUIRED)

add_executable(virtual_ecu main.cpp)
target_link_libraries(virtual_ecu PRIVATE SystemC::systemc)

