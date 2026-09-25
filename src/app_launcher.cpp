/**
 * ============================================================================
 *  app_launcher.cpp - 外部程序启动实现
 * ============================================================================
 */
#include "app_launcher.h"

#include <windows.h>

#include <string>

namespace pptx_selector {

namespace {

// 将一个参数按 Windows 命令行规则加引号，避免空格、中文或特殊字符
// 被目标程序错误拆分。CreateProcessW 会原样接收这条命令行。
std::wstring quoteCommandLineArg(const wchar_t* arg) {
    std::wstring quoted(L"\"");
    if (arg) {
        unsigned int backslashes = 0;
        for (const wchar_t* p = arg; *p; ++p) {
            if (*p == L'\\') {
                ++backslashes;
                continue;
            }
            if (*p == L'\"') {
                quoted.append(backslashes * 2 + 1, L'\\');
                quoted.push_back(L'\"');
                backslashes = 0;
                continue;
            }
            quoted.append(backslashes, L'\\');
            backslashes = 0;
            quoted.push_back(*p);
        }
        // 引号结尾前的反斜杠必须翻倍。
        quoted.append(backslashes * 2, L'\\');
    }
    quoted.push_back(L'\"');
    return quoted;
}

}  // namespace

bool appExists(const wchar_t* exePath) {
    if (!exePath || !*exePath) return false;
    const DWORD attributes = GetFileAttributesW(exePath);
    return attributes != INVALID_FILE_ATTRIBUTES &&
           (attributes & FILE_ATTRIBUTE_DIRECTORY) == 0;
}

bool launchApp(const wchar_t* exePath, const wchar_t* filePath,
               DWORD* errorCode) {
    if (errorCode) *errorCode = ERROR_INVALID_PARAMETER;
    if (!exePath || !*exePath || !filePath || !*filePath) return false;

    std::wstring commandLine = quoteCommandLineArg(exePath);
    commandLine.push_back(L' ');
    commandLine += quoteCommandLineArg(filePath);

    STARTUPINFOW startup = {};
    startup.cb = sizeof(startup);
    PROCESS_INFORMATION process = {};
    // CreateProcessW 要求命令行缓冲区可写，std::wstring 的连续存储满足这一点。
    BOOL ok = CreateProcessW(exePath, &commandLine[0], nullptr, nullptr, FALSE,
                             CREATE_NEW_PROCESS_GROUP, nullptr, nullptr,
                             &startup, &process);
    if (!ok) {
        if (errorCode) *errorCode = GetLastError();
        return false;
    }

    CloseHandle(process.hThread);
    CloseHandle(process.hProcess);
    if (errorCode) *errorCode = ERROR_SUCCESS;
    return true;
}

}  // namespace pptx_selector
