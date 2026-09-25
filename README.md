# PPTX 智能打开器

根据 PPTX 文件的创建软件，自动选择 PowerPoint 或 WPS 打开，避免排版错乱。

## 功能特点

- 自动检测 PPTX 文件是由 Microsoft PowerPoint 还是 WPS 创建
- 智能选择对应软件打开，确保排版正确
- 通过 `config.ini` 配置路径，无需重新编译
- 零控制台窗口，静默运行（`-mwindows`）
- 检测软件不可用时自动降级到另一软件，最终走兜底策略
- 完整的日志记录，便于问题排查
- 模块化代码结构，便于维护与扩展

## 使用场景

适用于教室电脑等多人共享环境：不同老师使用不同软件制作 PPT，用对应软件打开才能避免排版错乱。将 `.pptx` 关联到本程序，双击即可自动选择正确的软件打开。

## 工作原理

1. 读取 PPTX 文件（PPTX 本质是 ZIP 格式）到内存
2. 用 miniz 解析 ZIP，提取 `docProps/app.xml`
3. 解析其中的 `<Application>` 标签内容
4. 根据应用名称判断创建软件：
   - 包含 `WPS` 或 `Kingsoft` → 使用 WPS 打开
   - 包含 `Microsoft` 或 `PowerPoint` → 使用 PowerPoint 打开
   - 其他情况 → 按 `Fallback` 配置兜底
5. 若检测到的软件未安装，自动降级到另一软件；都不可用则按兜底配置启动

## 目录结构

```
trae/
├── README.md              本文档
├── AGENTS.md              AI 助手协作指引
├── .gitignore             忽略构建产物
├── config.ini             运行配置（路径、日志、兜底策略）
├── build.bat              Windows 批处理编译脚本
├── compile.ps1            PowerShell 编译脚本
├── src/                   项目源码
│   ├── main.cpp           程序入口与主流程编排
│   ├── error_codes.h      统一错误码定义
│   ├── config.h / .cpp    配置加载（读 config.ini）
│   ├── logger.h / .cpp    日志系统
│   ├── pptx_detector.h/.cpp  PPTX 创建者检测（ZIP 解析）
│   └── app_launcher.h/.cpp   外部程序启动（CreateProcessW）
├── third_party/           第三方库
│   ├── miniz.c            miniz ZIP 解压库实现
│   └── miniz.h            miniz 头文件
└── resources/             资源文件
    ├── icon.ico           应用图标
    └── resource.rc        Windows 资源文件
```

## 编译

### 环境要求

- Windows 系统
- MinGW-w64（GCC 8+，提供 `g++` 与 `windres`）

### 方式一：批处理脚本

双击运行 `build.bat`：

```bat
build.bat
```

### 方式二：PowerShell 脚本

```powershell
powershell -ExecutionPolicy Bypass -File compile.ps1
```

### 方式三：手动编译

```bash
# 1. 编译资源文件（图标）
windres resources\resource.rc -o resource.o

# 2. 编译主程序
g++ src\main.cpp src\config.cpp src\logger.cpp src\pptx_detector.cpp src\app_launcher.cpp \
    third_party\miniz.c resource.o \
    -o pptx_selector.exe \
    -Isrc -Ithird_party -lshell32 \
    -O2 -mwindows -static -static-libgcc -static-libstdc++ -s
```

编译参数说明：

| 参数 | 说明 |
|------|------|
| `-Isrc -Ithird_party` | 头文件搜索路径 |
| `-lshell32` | 链接 Win32 进程与目录 API 所需库 |
| `-O2` | 优化级别 |
| `-mwindows` | Windows 子系统，无控制台窗口 |
| `-static` | 全静态链接，无运行时依赖 |
| `-s` | 剥离调试符号，减小体积 |

## 使用方法

### 命令行调用

```bash
pptx_selector.exe "path\to\your\presentation.pptx"
```

### 文件关联（推荐）

将 `.pptx` 文件默认打开方式设置为本程序，即可双击自动选择软件打开。

## 配置

将 `config.ini` 放在 `pptx_selector.exe` 同目录下，程序启动时自动读取。修改后无需重新编译，重启程序生效。

```ini
[Paths]
PowerPoint=C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE
WPS=C:\Program Files\Kingsoft\WPS Office\12.1.0.23542\office6\wpp.exe
LogFile=D:\logs\pptx_selector.log

[Options]
EnableLog=1
Fallback=0
```

| 配置项 | 说明 |
|--------|------|
| `PowerPoint` | PowerPoint 可执行文件路径 |
| `WPS` | WPS 演示可执行文件路径 |
| `LogFile` | 日志文件路径（自动创建目录） |
| `EnableLog` | 是否启用日志：`1`=启用，`0`=禁用 |
| `Fallback` | 兜底程序：`0`=WPS，`1`=PowerPoint |

> 无 `config.ini` 时程序使用内置默认值运行，但路径可能不匹配实际安装位置，建议始终配置。

## 架构概览

程序采用分层模块化设计，所有模块位于 `pptx_selector` 命名空间：

```
main.cpp (流程编排)
   │
   ├── config      读取 config.ini，得到 Config
   ├── logger      按 Config 初始化，记录运行日志
   ├── app_launcher 检查软件可用性、启动外部程序
   └── pptx_detector 读取内存中的 PPTX，识别创建软件
```

**主流程**（`main.cpp`）：

1. 加载配置 → 初始化日志
2. 解析命令行参数，获取 PPTX 路径（Unicode，兼容中文路径）
3. 检查 PowerPoint / WPS 是否存在
4. 读取 PPTX 到内存
5. 检测创建软件
6. 按检测结果 + 软件可用性选择启动器，失败则兜底

**错误码**（`error_codes.h`）：定义了 `Err` 枚举，统一程序退出码语义。

## 技术栈

- **C++**：核心逻辑实现
- **miniz**：轻量级单文件 ZIP 解压缩库（位于 `third_party/`）
- **Win32 API**：文件 IO、路径处理、Unicode 进程启动
- **CreateProcessW**：使用完整 Unicode 路径和严格转义的命令行启动外部程序

## 故障排查

### 程序无法打开 PPTX

1. 查看日志文件（默认 `D:\logs\pptx_selector.log`）
2. 确认 `config.ini` 中 PowerPoint 和 WPS 的路径是否正确
3. 检查 PPTX 文件是否损坏（非有效 ZIP）

### 日志无法创建

1. 检查日志路径所在目录是否有写入权限
2. 修改 `config.ini` 中的 `LogFile` 到有权限的位置
3. 设置 `EnableLog=0` 可禁用日志

### 编译失败

1. 确认 `g++` 与 `windres` 在 PATH 中（`g++ --version`）
2. 确认从项目根目录执行编译脚本（脚本使用相对路径）
3. 检查 `third_party/` 与 `resources/` 目录是否完整

## 注意事项

1. 确保 PowerPoint 和 WPS 的安装路径与 `config.ini` 配置一致
2. 确保日志路径所在目录有写入权限
3. 程序仅在 Windows 系统运行
4. 文件读取、配置、日志和启动均使用 Unicode Win32 API，避免中文路径经过系统 ANSI 代码页后损坏
5. 日志以 UTF-8（含 BOM）写入，便于记事本等工具正确识别中文
