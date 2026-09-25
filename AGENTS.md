# AGENTS.md

本文件为 AI 编程助手提供本项目的协作指引，帮助助手快速理解项目结构、约定与工作流。

## 项目概述

**PPTX 智能打开器**（`pptx_selector.exe`）：一个 Windows 原生 C++ 小工具，根据 PPTX 文件的创建软件（Microsoft PowerPoint 或 WPS）自动选择对应软件打开，避免排版错乱。

- **目标用户**：教室电脑等多人共享环境
- **核心机制**：PPTX 是 ZIP 包，读取 `docProps/app.xml` 中的 `<Application>` 标签判断创建者
- **运行方式**：作为 `.pptx` 的默认打开程序，双击触发

## 技术栈

- 语言：C++（兼容 C++11 及以上）
- 构建：MinGW-w64（`g++` + `windres`），全静态链接
- 第三方库：miniz（单文件 ZIP 解压，位于 `third_party/`）
- 平台 API：Win32 API（文件 IO、CreateProcessW、INI 读取）

## 目录结构

```
trae/
├── README.md                 面向用户的项目文档
├── AGENTS.md                 本文件
├── .gitignore                忽略 *.exe / *.o / *.log 等
├── config.ini                运行配置（与 exe 同目录生效）
├── build.bat                 批处理编译脚本
├── compile.ps1               PowerShell 编译脚本
├── src/                      项目源码（pptx_selector 命名空间）
│   ├── main.cpp              入口 + 主流程编排
│   ├── error_codes.h         Err 枚举（统一退出码）
│   ├── config.h / .cpp       Config 结构 + LoadConfig + GetExeDir
│   ├── logger.h / .cpp       Logger 类（文件日志）
│   ├── pptx_detector.h/.cpp  detectPptxCreator + AppCreator 枚举
│   └── app_launcher.h/.cpp   appExists + launchApp
├── third_party/              第三方库（勿修改）
│   ├── miniz.c
│   └── miniz.h
└── resources/                Windows 资源
    ├── icon.ico
    └── resource.rc
```

## 架构与模块职责

所有业务代码位于 `pptx_selector` 命名空间。模块间通过头文件暴露最小接口，内部实现细节放在匿名命名空间中。

| 模块 | 职责 | 关键符号 |
|------|------|----------|
| `error_codes.h` | 统一错误码 | `enum class Err` |
| `config` | 读取 `config.ini`、定位 exe 目录 | `struct Config`, `LoadConfig()`, `GetExeDir()` |
| `logger` | 文件日志，自动时间戳 | `class Logger` |
| `pptx_detector` | 解析 ZIP 提取 `<Application>` 并分类 | `enum class AppCreator`, `detectPptxCreator()` |
| `app_launcher` | 检查 exe 存在性、CreateProcessW 启动 | `appExists()`, `launchApp()` |
| `main` | 流程编排与降级/兜底决策 | `WinMain`, `chooseLauncher()`, `launchFallback()` |

### 主流程（main.cpp）

1. `GetExeDir()` + `LoadConfig()` → 得到 `Config`
2. 按 `Config` 构造 `Logger`，`initDir()` 建目录
3. `CommandLineToArgvW` 取 PPTX 路径，全程保留 Unicode
4. `appExists()` 检查 PowerPoint / WPS 是否可用
5. `readFileToMemory()`（`CreateFileW`）读取 PPTX 到内存
6. `detectPptxCreator()` 识别创建软件
7. `chooseLauncher()` 决策；失败走 `launchFallback()`

### 决策与降级逻辑

- 检测为 WPS：优先 WPS，WPS 不存在则降级 PowerPoint
- 检测为 Microsoft：优先 PowerPoint，不存在则降级 WPS
- 未知 / 无可用启动器 / 启动失败：按 `Config.fallback` 兜底（0=WPS, 1=PowerPoint）

> **重要**：本程序设计为"尽量打开文件"，多数失败路径会兜底启动而非直接报错退出。修改时请保留此行为。

## 编码约定

- **命名空间**：所有业务代码放入 `pptx_selector`；模块内部辅助函数放匿名命名空间
- **头文件**：使用 `#pragma once`，包含守卫不用 `#ifndef`
- **字符串**：文件路径、配置和进程启动统一使用 Unicode（`wstring`/`CreateFileW`/`CreateProcessW`/`GetPrivateProfileStringW`），日志写为 UTF-8；避免在启动前转成 ANSI 造成中文路径损坏
- **错误处理**：用 `enum class Err` 返回语义化退出码；运行失败尽量兜底而非崩溃
- **资源管理**：HANDLE 立即 `CloseHandle`；`mz_zip_archive` 用 RAII `ZipGuard` 守护；`LPWSTR*` 用 `LocalFree` 释放
- **注释**：文件头用块注释说明模块用途；复杂逻辑用行注释；公开 API 用 `///` Doxygen 风格
- **第三方代码**：`third_party/` 下的 miniz 保持原样，不修改；通过 `MINIZ_IMPLEMENTATION` 在 `pptx_detector.cpp` 中生成实现

## 构建

从项目根目录执行：

```bash
# 批处理
build.bat

# 或 PowerShell
powershell -ExecutionPolicy Bypass -File compile.ps1
```

关键点：
- `windres` 编译 `resources/resource.rc` → `resource.o`（图标）
- `g++` 编译 `src/*.cpp` + `third_party/miniz.c` + `resource.o`
- 头文件路径：`-Isrc -Ithird_party`
- 链接：`-lshell32`（`SHCreateDirectoryExW` 等 Shell 目录 API）
- 静态链接 + `-mwindows`（无控制台）+ `-s`（剥离符号）

构建产物 `pptx_selector.exe` 和 `resource.o` 已被 `.gitignore` 忽略。

## 修改指引

### 新增配置项

1. 在 `Config` 结构（`config.h`）添加字段，在构造函数设默认值
2. 在 `LoadConfig`（`config.cpp`）用 `GetPrivateProfileStringW` / `GetPrivateProfileIntW` 读取
3. 更新 `config.ini` 注释与 `README.md` 配置表

### 新增检测规则

- 修改 `pptx_detector.cpp` 中的 `classifyApp()`，扩展 `AppCreator` 枚举与匹配逻辑
- 若需新增启动器类型，同步更新 `chooseLauncher()`（main.cpp）

### 调整降级/兜底策略

- 集中在 `main.cpp` 的 `chooseLauncher()` 与 `launchFallback()`
- 修改时务必保持"尽量打开文件"的设计意图

## 测试与验证

- 无自动化测试。验证方式：构造分别由 PowerPoint 和 WPS 创建的 PPTX，配置真实安装路径后运行
- 查看日志（默认 `D:\logs\pptx_selector.log`）确认决策链路
- 边界场景：中文路径、文件不存在、ZIP 损坏、软件未安装、config.ini 缺失

## 常见陷阱

1. **中文路径**：文件读取和启动都必须使用 `CreateFileW`/`GetFileAttributesW`/`CreateProcessW`，禁止先转换为 ANSI
2. **miniz 实现位置**：`MINIZ_IMPLEMENTATION` 只能在**一个** `.cpp` 中定义（当前是 `pptx_detector.cpp`），否则重复定义链接错误
3. **INI 默认值**：`GetPrivateProfileStringW` 第三个参数为默认值，传入当前字段值以实现"INI 缺失则保留默认"
4. **资源路径**：`resource.rc` 中 `ICON "icon.ico"` 是相对路径，`windres` 从 rc 文件所在目录解析，故两者必须同在 `resources/`
5. **构建目录**：脚本使用相对路径，必须从项目根目录执行

## 当前未解决问题

- 本机未安装 MinGW-w64，无法本地验证编译（代码逻辑忠实保留自原单文件版本）
- 无自动化测试用例
