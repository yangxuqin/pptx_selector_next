/**
 * ============================================================================
 *  config.cpp - 配置管理实现
 * ============================================================================
 */
#include "config.h"

#include <cstdio>

namespace pptx_selector {

Config::Config() {
    // 默认值：与历史行为保持一致，便于无 config.ini 时直接运行
    strcpy_s(pathPowerPoint,
        "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE");
    strcpy_s(pathWPS,
        "C:\\Program Files\\Kingsoft\\WPS Office\\12.1.0.23542\\office6\\wpp.exe");
    logFile[0] = '\0';
    enableLog = true;
    fallback  = 0;  // 0=WPS, 1=PowerPoint
}

void LoadConfig(Config& cfg, const char* iniPath) {
    // GetPrivateProfileStringA 第三个参数是默认值；传入当前字段值，
    // 使 INI 缺失该键时保留默认，与历史行为一致。
    GetPrivateProfileStringA("Paths", "PowerPoint",
        cfg.pathPowerPoint, cfg.pathPowerPoint, MAX_PATH, iniPath);
    GetPrivateProfileStringA("Paths", "WPS",
        cfg.pathWPS, cfg.pathWPS, MAX_PATH, iniPath);
    GetPrivateProfileStringA("Paths", "LogFile",
        cfg.logFile, cfg.logFile, MAX_PATH, iniPath);

    cfg.enableLog = (GetPrivateProfileIntA("Options", "EnableLog", 1, iniPath) != 0);
    cfg.fallback  =  GetPrivateProfileIntA("Options", "Fallback",  0, iniPath);
}

std::string GetExeDir() {
    char path[MAX_PATH] = {};
    GetModuleFileNameA(nullptr, path, MAX_PATH);
    char* last = strrchr(path, '\\');
    if (last) *last = '\0';
    return path;
}

}  // namespace pptx_selector
