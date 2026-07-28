/**
 * ============================================================================
 *  PPTX 智能打开器 - pptx_selector.exe
 * ============================================================================
 *  功能：根据 PPTX 文件的创建软件，自动选择 PowerPoint 或 WPS 打开
 *  场景：教室电脑，不同老师用不同软件制作 PPT，避免排版错乱
 *  原理：PPTX 是 ZIP 包，读取 docProps/app.xml 的 <Application> 判断创建者
 *
 *  配置：自动读取 exe 同目录下的 config.ini，无需重新编译即可修改路径
 *
 *  编译（详见 build.bat）：
 *    windres resources\resource.rc -o resource.o
 *    g++ src\*.cpp third_party\miniz.c resource.o -o pptx_selector.exe ^
 *        -Isrc -Ithird_party -lshell32 -O2 -mwindows -static -static-libgcc ^
 *        -static-libstdc++ -s
 * ============================================================================
 */

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <shellapi.h>

#include <string>
#include <vector>

#include "app_launcher.h"
#include "config.h"
#include "error_codes.h"
#include "logger.h"
#include "pptx_detector.h"

using namespace pptx_selector;

namespace {

/// 读取整个文件到内存。失败返回空 vector。
std::vector<char> readFileToMemory(const wchar_t* path, const Logger& log) {
    HANDLE hFile = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ,
        nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) {
        log.write("错误: 无法打开文件 (CreateFileW 失败)");
        return {};
    }

    LARGE_INTEGER fsize{};
    GetFileSizeEx(hFile, &fsize);

    std::vector<char> data(static_cast<std::size_t>(fsize.QuadPart));
    DWORD bytesRead = 0;
    BOOL ok = ReadFile(hFile, data.data(),
        static_cast<DWORD>(data.size()), &bytesRead, nullptr);
    CloseHandle(hFile);  // 立即释放文件句柄

    if (!ok || bytesRead != data.size()) {
        log.write("错误: 读取文件失败 (读取 %lu / 期望 %zu 字节)",
            bytesRead, data.size());
        return {};
    }
    return data;
}

/// 按检测结果与可用性选择启动器。返回 nullptr 表示无可用启动器。
const char* chooseLauncher(const DetectionResult& detection,
                           const Config& cfg,
                           bool wpsExists,
                           bool pptExists,
                           const Logger& log) {
    if (detection.creator == AppCreator::WPS) {
        if (wpsExists) {
            log.write("决策: 使用 WPS 打开");
            return cfg.pathWPS;
        }
        if (pptExists) {
            log.write("WPS 不存在，降级为 PowerPoint 打开");
            return cfg.pathPowerPoint;
        }
    } else if (detection.creator == AppCreator::MicrosoftOffice) {
        if (pptExists) {
            log.write("决策: 使用 PowerPoint 打开");
            return cfg.pathPowerPoint;
        }
        if (wpsExists) {
            log.write("PowerPoint 不存在，降级为 WPS 打开");
            return cfg.pathWPS;
        }
    }
    return nullptr;
}

/// 执行兜底启动。
bool launchFallback(const Config& cfg, const char* pptxPath, const Logger& log) {
    const char* fb = (cfg.fallback == 1) ? cfg.pathPowerPoint : cfg.pathWPS;
    log.write("兜底: 使用 %s 打开", cfg.fallback == 1 ? "PowerPoint" : "WPS");
    return launchApp(fb, pptxPath);
}

}  // namespace

int WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int) {
    DWORD t0 = GetTickCount();

    // 1. 加载配置
    std::string iniPath = GetExeDir() + "\\config.ini";
    Config cfg;
    LoadConfig(cfg, iniPath.c_str());

    // 日志文件默认保存在 exe 同目录
    if (cfg.logFile[0] == '\0') {
        std::string logPath = GetExeDir() + "\\pptx_selector.log";
        strcpy_s(cfg.logFile, logPath.c_str());
    }

    // 2. 初始化日志
    Logger logger(cfg.enableLog, cfg.logFile);
    logger.initDir();

    logger.write("========== 程序启动 ==========");
    logger.write("配置文件: %s", iniPath.c_str());
    logger.write("PowerPoint: %s", cfg.pathPowerPoint);
    logger.write("WPS: %s", cfg.pathWPS);
    logger.write("日志文件: %s", cfg.logFile);
    logger.write("兜底策略: %s", cfg.fallback == 0 ? "WPS" : "PowerPoint");

    // 3. 解析命令行参数
    int argc = 0;
    LPWSTR* argvW = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (!argvW || argc < 2) {
        logger.write("错误: 命令行参数不足 (argc=%d)", argc);
        if (argvW) LocalFree(argvW);
        return static_cast<int>(Err::InvalidArgs);
    }

    // ANSI 路径（用于日志和 ShellExecuteA）
    char pptxPathA[MAX_PATH] = {};
    WideCharToMultiByte(CP_ACP, 0, argvW[1], -1,
        pptxPathA, MAX_PATH, nullptr, nullptr);
    logger.write("目标文件: %s", pptxPathA);

    // 4. 检查软件可用性
    bool pptExists = appExists(cfg.pathPowerPoint);
    bool wpsExists = appExists(cfg.pathWPS);
    logger.write("PowerPoint存在: %s", pptExists ? "是" : "否");
    logger.write("WPS存在: %s",         wpsExists ? "是" : "否");

    // 5. 目标文件存在性检查
    if (GetFileAttributesW(argvW[1]) == INVALID_FILE_ATTRIBUTES) {
        logger.write("错误: 目标文件不存在，兜底启动");
        launchFallback(cfg, pptxPathA, logger);
        LocalFree(argvW);
        return static_cast<int>(Err::FileNotFound);
    }

    // 6. 读取文件到内存
    std::vector<char> fileData = readFileToMemory(argvW[1], logger);
    LocalFree(argvW);  // 参数已转 ANSI，释放宽字符副本

    if (fileData.empty()) {
        logger.write("错误: 读取文件失败，兜底启动");
        launchFallback(cfg, pptxPathA, logger);
        return static_cast<int>(Err::ReadFailed);
    }
    logger.write("文件已读入内存 (%zu 字节)，耗时 %lu ms",
        fileData.size(), GetTickCount() - t0);

    // 7. 检测创建软件
    DetectionResult detection = detectPptxCreator(fileData.data(), fileData.size());
    logger.write("检测到 Application: %s",
        detection.application.empty() ? "(空)" : detection.application.c_str());

    // 8. 决策并启动
    const char* launcher = chooseLauncher(detection, cfg, wpsExists, pptExists, logger);
    bool launched = false;
    if (launcher) {
        launched = launchApp(launcher, pptxPathA);
        if (!launched) {
            logger.write("错误: 启动失败，尝试兜底");
        }
    }

    if (!launched) {
        launched = launchFallback(cfg, pptxPathA, logger);
    }

    logger.write("总耗时: %lu ms", GetTickCount() - t0);
    logger.write("========== 程序结束 ==========");
    return launched ? static_cast<int>(Err::Ok)
                    : static_cast<int>(Err::LaunchFailed);
}
