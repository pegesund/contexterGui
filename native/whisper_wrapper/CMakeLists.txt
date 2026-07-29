cmake_minimum_required(VERSION 3.20)

project(spell_whisper LANGUAGES C CXX)

if(NOT DEFINED WHISPER_CPP_DIR)
    message(FATAL_ERROR "Set WHISPER_CPP_DIR to a checked-out whisper.cpp source directory")
endif()

set(BUILD_SHARED_LIBS OFF CACHE BOOL "" FORCE)
set(WHISPER_BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
set(WHISPER_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(WHISPER_BUILD_SERVER OFF CACHE BOOL "" FORCE)
set(WHISPER_BUILD_TOOLS OFF CACHE BOOL "" FORCE)
set(GGML_OPENMP OFF CACHE BOOL "" FORCE)

# CI ships a portable DLL (no -march=native): it must run on every user CPU.
# Local dev builds can opt in to native codegen (AVX2/FMA on modern x86) with
# -DSPELL_WHISPER_NATIVE=ON — medium-q5 transcription is several times faster.
option(SPELL_WHISPER_NATIVE "Build ggml with native CPU optimizations" OFF)
set(GGML_NATIVE ${SPELL_WHISPER_NATIVE} CACHE BOOL "" FORCE)

add_subdirectory("${WHISPER_CPP_DIR}" "${CMAKE_BINARY_DIR}/whisper.cpp" EXCLUDE_FROM_ALL)

add_library(spell_whisper SHARED spell_whisper.cpp)
target_compile_features(spell_whisper PRIVATE cxx_std_17)
target_include_directories(spell_whisper PRIVATE "${WHISPER_CPP_DIR}/include")
target_link_libraries(spell_whisper PRIVATE whisper)

if(MSVC)
    target_compile_options(spell_whisper PRIVATE /EHsc)
endif()

set_target_properties(spell_whisper PROPERTIES
    OUTPUT_NAME "whisper"
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin"
    LIBRARY_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin"
    ARCHIVE_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/lib"
)
