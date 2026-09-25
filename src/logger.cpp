/**
 * ============================================================================
 *  logger.cpp - 日志系统实现
 * ============================================================================
 */
#include "logger.h"

#include <cstdarg>
#include <cstdio>
#include <cstring>
#include <string>
#include <vector>
#include <shlobj.h>

namespace pptx_selector {

Logger::Logger() : enabled_(false) {
}

Logger::Logger(bool enableLog, const wchar_t* logFile) : enabled_(enableLog) {
    if (logFile) file_ = logFile;
}

void Logger::initDir() const {
    if (!enabled_ || file_.empty()) return;

    std::wstring::size_type last = file_.find_last_of(L"\\/");
    if (last != std::wstring::npos && last > 0) {
        std::wstring dir = file_.substr(0, last);
        SHCreateDirectoryExW(nullptr, dir.c_str(), nullptr);
    }
}

void Logger::write(const char* format, ...) const {
    if (!enabled_ || file_.empty() || !format) return;
    char message[8192] = {};
    va_list args;
    va_start(args, format);
    _vsnprintf_s(message, sizeof(message), _TRUNCATE, format, args);
    va_end(args);
    writeUtf8(message, strlen(message));
}

void Logger::writeW(const wchar_t* format, ...) const {
    if (!enabled_ || file_.empty() || !format) return;
    wchar_t message[8192] = {};
    va_list args;
    va_start(args, format);
    _vsnwprintf_s(message, sizeof(message) / sizeof(message[0]), _TRUNCATE,
                  format, args);
    va_end(args);
    int bytes = WideCharToMultiByte(CP_UTF8, 0, message, -1, nullptr, 0,
                                   nullptr, nullptr);
    if (bytes <= 1) return;
    std::vector<char> utf8(static_cast<std::size_t>(bytes));
    WideCharToMultiByte(CP_UTF8, 0, message, -1, utf8.data(), bytes,
                        nullptr, nullptr);
    writeUtf8(utf8.data(), static_cast<std::size_t>(bytes - 1));
}

void Logger::writeUtf8(const char* text, std::size_t length) const {
    if (!enabled_ || file_.empty() || !text) return;
    HANDLE file = CreateFileW(file_.c_str(), FILE_APPEND_DATA,
                              FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr,
                              OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) return;

    LARGE_INTEGER size{};
    if (GetFileSizeEx(file, &size) && size.QuadPart == 0) {
        static const char bom[] = "\xEF\xBB\xBF";
        DWORD written = 0;
        WriteFile(file, bom, 3, &written, nullptr);
    }

    SYSTEMTIME st;
    GetLocalTime(&st);
    char prefix[64] = {};
    int prefixLength = _snprintf_s(prefix, sizeof(prefix), _TRUNCATE,
        "[%04d-%02d-%02d %02d:%02d:%02d] ", st.wYear, st.wMonth,
        st.wDay, st.wHour, st.wMinute, st.wSecond);
    DWORD written = 0;
    if (prefixLength > 0)
        WriteFile(file, prefix, static_cast<DWORD>(prefixLength), &written, nullptr);
    if (length > 0)
        WriteFile(file, text, static_cast<DWORD>(length), &written, nullptr);
    static const char newline[] = "\r\n";
    WriteFile(file, newline, 2, &written, nullptr);
    CloseHandle(file);
}

}  // namespace pptx_selector
