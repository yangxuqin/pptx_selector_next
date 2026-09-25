/**
 * ============================================================================
 *  app_launcher.h - 外部程序启动
 *  PPTX 智能打开器
 * ============================================================================
 *  封装 CreateProcessW，负责以指定 exe 打开 PPTX 文件。
 * ============================================================================
 */
#pragma once

#include <string>
#include <windows.h>

namespace pptx_selector {

/// 检查指定路径的可执行文件是否存在。
bool appExists(const wchar_t* exePath);

/// 以 exe 打开 filePath。返回是否成功创建进程。
/// @param exePath   可执行文件路径
/// @param filePath  要打开的 PPTX 文件路径
/// @param errorCode 可选，失败时返回 CreateProcessW 的错误码
bool launchApp(const wchar_t* exePath, const wchar_t* filePath,
               DWORD* errorCode = nullptr);

}  // namespace pptx_selector
