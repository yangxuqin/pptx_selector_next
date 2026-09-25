/**
 * ============================================================================
 *  config.cpp - 配置管理实现
 * ============================================================================
 */
#include "config.h"

#include <vector>

namespace pptx_selector {

Config::Config() {
    // 默认值：与历史行为保持一致，便于无 config.ini 时直接运行
    pathPowerPoint = L"C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE";
    pathWPS = L"C:\\Program Files\\Kingsoft\\WPS Office\\12.1.0.23542\\office6\\wpp.exe";
    logFile = L"D:\\logs\\pptx_selector.log";
    enableLog = true;
    fallback  = 0;  // 0=WPS, 1=PowerPoint
}

void LoadConfig(Config& cfg, const wchar_t* iniPath) {
    // GetPrivateProfileStringW 第三个参数是默认值；传入当前字段值，
    // 使 INI 缺失该键时保留默认，与历史行为一致。
    auto readString = [iniPath](const wchar_t* section, const wchar_t* key,
                                const std::wstring& fallback) {
        std::vector<wchar_t> buffer(32768);
        DWORD length = GetPrivateProfileStringW(
            section, key, fallback.c_str(), buffer.data(),
            static_cast<DWORD>(buffer.size()), iniPath);
        return std::wstring(buffer.data(), length);
    };
    cfg.pathPowerPoint = readString(L"Paths", L"PowerPoint", cfg.pathPowerPoint);
    cfg.pathWPS = readString(L"Paths", L"WPS", cfg.pathWPS);
    cfg.logFile = readString(L"Paths", L"LogFile", cfg.logFile);

    cfg.enableLog = (GetPrivateProfileIntW(L"Options", L"EnableLog", 1, iniPath) != 0);
    cfg.fallback  =  GetPrivateProfileIntW(L"Options", L"Fallback",  0, iniPath);
}

std::wstring GetExeDir() {
    std::vector<wchar_t> path(32768);
    DWORD length = GetModuleFileNameW(nullptr, path.data(),
                                      static_cast<DWORD>(path.size()));
    if (length == 0 || length >= path.size()) return L".";
    std::wstring result(path.data(), length);
    const std::wstring::size_type last = result.find_last_of(L"\\/");
    if (last == std::wstring::npos) return L".";
    result.resize(last);
    return result;
}

}  // namespace pptx_selector
