/**
 * ============================================================================
 *  config.h - 配置管理
 *  PPTX 智能打开器
 * ============================================================================
 *  从 exe 同目录下的 config.ini 读取路径与选项，缺失项保留默认值。
 * ============================================================================
 */
#pragma once

#include <windows.h>
#include <string>

namespace pptx_selector {

/// 程序运行配置。所有字段在 LoadConfig 后生效。
struct Config {
    std::wstring pathPowerPoint;    ///< PowerPoint 可执行文件路径
    std::wstring pathWPS;           ///< WPS 演示可执行文件路径
    std::wstring logFile;           ///< 日志文件路径
    bool enableLog;                 ///< 是否启用日志
    int  fallback;                  ///< 兜底程序：0=WPS, 1=PowerPoint

    /// 以内置默认值初始化配置。
    Config();
};

/// 从 INI 文件加载配置，覆盖现有字段。缺失项保留传入 cfg 中的值。
void LoadConfig(Config& cfg, const wchar_t* iniPath);

/// 返回 exe 所在目录（不带末尾反斜杠）。
std::wstring GetExeDir();

}  // namespace pptx_selector
