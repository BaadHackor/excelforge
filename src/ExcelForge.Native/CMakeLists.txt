cmake_minimum_required(VERSION 3.20)
project(ExcelForge.Native CXX)

set(CMAKE_CXX_STANDARD 17)

add_library(ExcelForge.Native SHARED ExcelNative.cpp ExcelNative.h)

target_compile_definitions(ExcelForge.Native PRIVATE
    EXCELNATIVE_EXPORTS
    _UNICODE UNICODE
    WIN32_LEAN_AND_MEAN
    _CRT_SECURE_NO_WARNINGS
)

target_link_libraries(ExcelForge.Native PRIVATE
    ole32 oleaut32 shlwapi
)

set_target_properties(ExcelForge.Native PROPERTIES
    OUTPUT_NAME "ExcelForge.Native"
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin"
)
