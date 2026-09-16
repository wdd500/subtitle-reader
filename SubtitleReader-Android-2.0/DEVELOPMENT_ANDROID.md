# 字幕阅读器 Android 版开发文档

版本：v2.0（versionCode 2）
平台：Android（Jetpack Compose 原生应用）
包名：`com.subtitle.reader`
归档包：`SubtitleReader-Android-2.0.zip`

---

## 1. 项目概述

字幕阅读器是一个面向 Android 的本地字幕文件阅读工具，支持 `.srt`、`.sub`、`.ass/.ssa` 三种字幕格式的解析、沉浸式阅读、样式定制、阅读进度记忆与导入导出。应用为**纯本地**离线应用，无网络权限，不收集任何数据。

### 1.1 核心能力

| 能力 | 说明 |
|---|---|
| 格式解析 | SRT / SUB（MicroDVD）/ ASS / SSA 三种格式 |
| 阅读模式 | 百分比进度或分页模式切换 |
| 沉浸阅读 | 隐藏系统栏、挖孔屏利用、音量键/媒体键翻页、手势翻页 |
| 样式定制 | 中英双语、字号、字体、文字/背景颜色、屏幕亮度 |
| 阅读记忆 | 按文件记忆上次阅读位置（索引 + 百分比） |
| 文件管理 | 系统文件夹浏览、收藏文件夹、最近打开、按类型筛选 |
| 导入 | 应用内打开目录、系统"打开方式"/"分享"直达 |
| 导出 | 文本 (.txt) / Word (.docx) |
| 安全 | 无网络、无存储危险权限、使用 SAF 持久化授权 |

### 1.2 技术栈

| 项 | 版本 |
|---|---|
| Kotlin | 1.9.20 |
| Android Gradle Plugin | 8.2.0 |
| Gradle Wrapper | 8.5 |
| JDK | 17 |
| compileSdk / targetSdk | 34 |
| minSdk | 24 (Android 7.0) |
| Jetpack Compose BOM | 2024.01.00 |
| Compose Compiler | 1.5.5 |
| material3 / material-icons-extended | BOM 管理 |
| activity-compose | 1.8.2 |
| lifecycle-runtime-ktx | 2.7.0 |
| core-ktx | 1.12.0 |

---

## 2. 工程结构

```
android/
├── build.gradle.kts              # 根构建：声明 AGP / Kotlin 插件版本
├── settings.gradle.kts           # 仓库配置与模块（:app）+ rootProject.name
├── gradle.properties             # JVM 参数、AndroidX、R 类非传递开关
├── gradlew / gradlew.bat         # Gradle Wrapper 启动脚本
├── gradle/wrapper/               # wrapper jar + 版本配置（Gradle 8.5）
└── app/
    ├── build.gradle.kts          # 应用模块构建配置（含依赖）
    ├── proguard-rules.pro        # 混淆规则
    └── src/main/
        ├── AndroidManifest.xml   # 清单：单 Activity + 意图过滤
        ├── java/com/subtitle/reader/
        │   ├── MainActivity.kt           # 入口：edge-to-edge、意图解析、音量键拦截
        │   ├── model/SubtitleItem.kt     # 字幕条目数据模型
        │   ├── parser/SubtitleParser.kt  # SRT/SUB/ASS 解析器
        │   ├── ui/screens/ReaderScreen.kt# 主界面（全部 UI 与交互逻辑）
        │   ├── ui/theme/Theme.kt         # Material3 主题
        │   ├── ui/theme/Color.kt         # 主题色板
        │   └── util/
        │       ├── Strings.kt            # 中英双语字符串资源
        │       └── DocxExporter.kt       # DOCX 导出（纯 XML 组装）
        └── res/
            ├── values/strings.xml        # 应用名
            ├── values/themes.xml         # 基础主题
            └── mipmap-*/ic_launcher.png  # 各密度启动图标
```

---

## 3. 模块说明

### 3.1 入口层：`MainActivity.kt`

**职责**：Edge-to-edge 沉浸、外部意图接收、实体键翻页拦截。

- `onCreate` 调用 `enableEdgeToEdge()` 让内容绘制到状态栏/导航栏之下，实现真正的全屏沉浸。
- `resolveIncomingUri(intent)`：
  - `ACTION_VIEW` → 取 `intent.data`（文件或 content URI）。
  - `ACTION_SEND` → 取 `EXTRA_STREAM`（分享出的文件），API 33+ 使用类型化重载。
  - `content://` 时调用 `contentResolver.takePersistableUriPermission` 获取持久读权限。
  - 结果保存到 Compose 状态 `incomingUri`，传入 `ReaderScreen(initialUri)`。
- `onNewIntent`：应用已运行时收到新的 VIEW/SEND 意图会被中转并触发自动打开。
- `dispatchKeyEvent`：仅在 `ReaderEvents.volumeKeyPaging == true`（阅读模式）时拦截并消费：
  - `VOLUME_UP` / `MEDIA_PREVIOUS` / `MEDIA_REWIND` → `onPagePrev()`
  - `VOLUME_DOWN` / `MEDIA_NEXT` / `MEDIA_FAST_FORWARD` → `onPageNext()`
  - 其余按键放行给系统。

**桥接对象**；`ReaderEvents`（顶层 singleton）：

| 字段 | 类型 | 说明 |
|---|---|---|
| `volumeKeyPaging` | `@Volatile Boolean` | 是否启用实体键翻页 |
| `onPagePrev` / `onPageNext` | `(() -> Unit)?` | 由 `ReaderScreen` 注册的翻页回调 |

> 设计要点：按键事件发生在 Activity 层，翻页逻辑在 Compose 层。用静态单例桥接二者，避免把 Compose 状态上提到 Activity，也不引入 ViewModel/DI 复杂度。

### 3.2 数据模型：`SubtitleItem.kt`

```kotlin
data class SubtitleItem(
    val index: Int = 0,     // 行号
    val start: String = "", // 开始时间码（"HH:MM:SS.mmm" 或 SUB 的帧号）
    val end: String = "",   // 结束时间码
    val text: String = "",  // 字幕文本（可多行 \n）
    val duration: String = "" // 时长（"1.50s" 或 SUB 的帧数）
)
```

### 3.3 解析层：`SubtitleParser.kt`

**入口**：`parseFile(path, content)` 根据文件名**后缀**（`uri.lastPathSegment` / 显示文件名）分发到对应解析器。这是「时间码、行号、时长」处理的核心。

#### SRT（`parseSrt`）
- 以空行分块，每块须 ≥3 行。
- 第 1 行索引号（非数字则跳过该块）。
- 第 2 行匹配时间码 `HH:MM:SS[,.]mmm --> HH:MM:SS[,.]mmm`，逗号统一替换为点。
- 其余行为文本（可能多行，保留换行）。
- 时长由 `calcDuration(start, end)` 计算并格式化为 `%.2fs`。

#### SUB / MicroDVD（`parseSub`）
- 逐行匹配 `{startFrame}{endFrame}text`。
- 文本中的 `|` 转为换行。
- 时长为 `帧数 frames`；index 为 1 起始的行序。

#### ASS / SSA（`parseAss`）
- 扫描 `[Events]` 段，读取 `Format:` 定义字段顺序。
- `Dialogue:` 行用 `splitAssValue` 解析（按逗号分割但**跳过 `{...}` 花括号内逗号**，保证标签文本不被误切）。
- `Text` 字段去除 `{...}` 内联样式标签，`\N` / `\n` 转真实换行。
- Start / End 为 `h:mm:ss.cc`，时长用 `calcDuration` 计算。

#### 时长计算 `calcDuration`
```kotlin
fun toSec(t: String): Double  // "HH:MM:SS(.ccc)" → 秒
format("%.2fs", endSec - startSec)
```

### 3.4 界面层：`ReaderScreen.kt`（约 1640 行，核心文件）

单文件承载全部 UI 与状态，采用「压缩状态驱动」模式，无 ViewModel（符合规模匹配原则）。

#### 主要状态（`remember`）

| 状态 | 说明 |
|---|---|
| `items` | 当前字幕条目列表（空 = 未加载） |
| `showTime/showIndex/showDuration` | 「时间码/行号/时长」显示开关 |
| `fontSize / fontFamily` | 字号 / 字体 |
| `textColor / bgColor` | 文字 / 背景颜色 |
| `settingsVisible` | 设置面板（顶栏+底栏）显隐 |
| `displayMode` | 0=百分比，1=分页 |
| `brightness` | 屏幕亮度（-1=跟随系统） |
| `fileUri / fileName` | 当前文件 |
| `recentFiles / folderBrowserUri` | 最近记录 / 文件夹浏览器 |
| `pendingRestoreIndex` | 待恢复阅读位置 |

#### 数据流（打开文件 → 展示）

```
UI 触发（点空态 / RecentFiles / FolderBrowser / 分享意图）
  → openFromFolder / openRecentFile / LaunchedEffect(initialUri)
  → loadSubtitle(context, uri, name)        // BufferedReader + parseFile
  → applyLoaded(uri, result, uriString)     // 写状态 + 记最近 + 记位置
  → LaunchedEffect(items) 恢复 pendingRestoreIndex
  → LazyColumn + SubtitleEntry 渲染
```

**打开链路要点**：
- `loadSubtitle` 用**真实文件名**（非 URI 字符串）判定扩展名，避免 content URI 末尾编码导致解析失效。
- `applyLoaded` 记住进度，关闭时 `closeFile` / `savePosition` 写入 `pos_<hash>` 与 `posf_<hash>`。
- 空解析或 IO 失败 → Toast `load_fail`（`Strings.get("load_fail")`）。

#### 沉浸阅读（Readest 风格）

| 场景 | 行为 |
|---|---|
| 阅读中且设置隐藏 | 隐藏系统栏（`BEHAVIOR_SHOW_TRANSIENT_BARS_BY_SWIPE` 滑动可临时唤出）、API 28+ 设 `LAYOUT_IN_DISPLAY_CUTOUT_MODE_SHORT_EDGES` 利用挖孔区 |
| 显示设置 | 恢复系统栏与 `LAYOUT_IN_DISPLAY_CUTOUT_MODE_DEFAULT` |

#### 交互设计

- **右侧手势翻页**：`pointerInput` 监听点击，右 1/3 区域中带点击 → 上/下半屏分别上一页/下一页（按可见条目数翻页）；左 2/3 中带点击 → 切换设置面板。
- **进度条**：顶部进度 Row——百分比模式显示 `NN%`；分页模式显示 `第 x 页/共 y 页`；点击文字弹出跳转对话框。
- **实体键翻页**：`DisposableEffect` 在阅读模式下注册 `ReaderEvents` 回调（复用 `scrollByPage`）。

#### 底部设置面板

三行结构：
1. 功能开关：时间码 / 行号 / 时长 / 分页模式（FilterChip）。
2. 样式：字体选择、字号 - / + / 现值 / +、文字颜色、背景颜色、导出。
3. 亮度：`BrightnessHigh` 图标 + Slider（0.05~1.0）+ 自动按钮（恢复跟随系统）。

亮度通过 `window.attributes.screenBrightness` 覆盖，持久化到 prefs 键 `brightness`。

#### 对话框组件

| 组件 | 功能 |
|---|---|
| `FolderBrowserDialog` | 全屏文件夹浏览器：SAF 树查询、导航栈、收藏、类型筛选（全部/SRT/SUB/ASS），`safeDrawingPadding()` 适配安全区 |
| `RecentFilesDialog` | 最近打开（含进度百分比前缀），可清空 |
| `FontPickerDialog` | 扫描 `/system/fonts` 列出系统字体 |
| `ColorPickerDialog` | 预设色 + RGB 自定义滑块 + 实时预览 |
| `ExportDialog` | 选择 TXT / DOCX 导出 |
| `LanguageDialog` | 中/英切换 |
| `JumpDialog` | 跳转百分比或指定页数 |
| `AboutDialog` | 版本 / 作者信息 |

#### 本地文件操作辅助函数

| 函数 | 说明 |
|---|---|
| `loadSubtitle(context, uri, name)` | 打开流并解析 |
| `resolveDisplayName(context, uri)` | 用 `OpenableColumns.DISPLAY_NAME` 取真实文件名（回退到解码的 lastPathSegment） |
| `savePosition / loadPosition / loadPositionFraction` | 位置持久化，键 `pos_`/`posf_` + `abs(uri.hashCode())` |
| `getRecentFiles / addToRecentFiles / clearRecentFiles` | 最近记录 JSON 数组（上限 10 条） |
| `loadFavoriteFolders / saveFavoriteFolders` | 收藏文件夹 JSON（root/docId/name） |
| `listFolderChildren / matchesFilter` | SAF 子项枚举 + 扩展名筛选 |
| `buildExportText / saveExportTxt / saveExportDocx` | 导出内容生成与写出 |

### 3.5 工具层

#### `Strings.kt`
- `ZH` / `EN` 两个 `Map<String, String>` 常量 + `get(key)` 按当前语言取文案（缺失回退为 key 本身）。
- `setLanguage` 运行时切换，`getLanguage` 返回当前语言。用于在**应用内**中英文切换（不依赖系统 Locale）。

#### `DocxExporter.kt`
- 纯内存组装 Office Open XML：`[Content_Types].xml`、`_rels/.rels`、`word/document.xml` 等 zip 条目，生成合法 `.docx`。
- `createDocx(text)` 入口，输出 `ByteArray` 供 `openOutputStream` 写出。无第三方依赖。

### 3.6 主题层：`Theme.kt` / `Color.kt`

- Material3 动态色彩方案（受支持设备），否则回退到预设浅色/深色色板。
- `SubtitleReaderTheme` 包裹主界面。

---

## 4. 数据持久化（SharedPreferences）

文件：`subtitle_reader`（`Context.MODE_PRIVATE`）

| 键 | 类型 | 说明 |
|---|---|---|
| `pos_<abs(uri.hashCode())>` | Int | 上次阅读的 `firstVisibleItemIndex` |
| `posf_<abs(uri.hashCode())>` | Float | 上次阅读位置百分比（0~1） |
| `recent_files` | JSON Array | 最近打开 URI 列表（≤10，最新在前） |
| `favorite_folders` | JSON Array | 收藏文件夹 `{root, doc, name}` |
| `last_folder` | String | 上次浏览的文件夹 URI |
| `brightness` | Float | 亮度覆盖值（-1 = 跟随系统） |

> 说明：URI 存储的是 `content://` 持久授权 URI，重建系统后权限可能失效，读取失败会静默降级（Toast 提示无法读取）。

---

## 5. 意图（Intent）配置

`AndroidManifest.xml` 中 `MainActivity` 的过滤器：

| 过滤器 | 触发方式 |
|---|---|
| MAIN + LAUNCHER | 桌面图标启动 |
| VIEW, scheme=`file`, mime=`text/*` | 文件管理器直接打开（老式 file 路径） |
| VIEW, scheme=`content`, mime=`text/*` / `application/x-subrip` / `application/octet-stream` | 文件管理器 / 下载管理"打开方式" |
| SEND, mime=`text/*` / `application/x-subrip` / `application/octet-stream` | 其他应用"分享到"本应用 |

均为隐式意图，无 `autoVerify`，不需后台声明。

---

## 6. 构建指南

### 6.1 环境要求

- JDK 17
- Android SDK（compileSdk 34）
- （首次构建）可访问 `services.gradle.org` 下载 Gradle 8.5

### 6.2 构建命令

```bash
# Windows（项目根为 android/）
cd android
gradlew.bat :app:assembleDebug --no-daemon

# Linux / macOS
./gradlew :app:assembleDebug

# Release（当前未配置签名，产物为未签名 APK）
gradlew :app:assembleRelease
```

如需指定本机 SDK：在 `android/local.properties` 写入 `sdk.dir=C:\Users\<user>\AppData\Local\Android\Sdk`。

### 6.3 构建产物

- Debug APK：`android/app/build/outputs/apk/debug/app-debug.apk`
- 归档包已内置可直接安装的 `SubtitleReader-2.0.apk`

### 6.4 常见问题

| 问题 | 解决 |
|---|---|
| `Directory 'X' does not contain a Gradle build` | 必须在 `android/` 目录内执行 Gradle |
| SDK 找不到 | 配置 `local.properties` 的 `sdk.dir` |
| 图标编译报错 | `material-icons-extended` 已引入，勿移除 |
| 网络受限 | 首次需下载 Gradle 8.5 与依赖，联网一次后即可离线构建 |

---

## 7. 本地化

追加文案两步：
1. 在 `Strings.kt` 的 `ZH`、`EN` 两处同时添加同 key。
2. 界面用 `Strings.get("key")` 读取。

语言默认中文；切换只影响应用内文案，不改变系统语言。

---

## 8. 已知限制与后续建议

- **无在线播放/音视频同步**：当前为纯文本阅读器。若需与播放器联动，可引入 ExoPlayer 并按时间码索引字幕。
- **无云同步**：阅读进度仅存本地。可接入轻量云同步或 WebDAV。
- **ASS 样式还原**：解析已去除标签，未渲染 `{\fn}` 等字体/位置样式。若要还原，可引入 LibreTTS/自定义 Canvas。
- **SMI/SRT 变种**：部分字幕含行内格式或 BOM/编码问题（如 GBK），当前按 UTF-8 读取。可增加编码探测（`CharsetDetector`）。
- **测试**：目前无 `test/` / `androidTest/` 目录。建议补充：
  - 单元测试：`SubtitleParser`（各格式样例 + 边界）。
  - UI 测试：Compose `createAndroidComposeRule`。

---

## 9. 版本记录

| 版本 | 内容 |
|---|---|
| 1.0 | 初版：SRT/SUB/ASS 解析、基础阅读、导出 TXT/DOCX |
| 2.0 | 沉浸式阅读（edge-to-edge、挖孔屏、音量/媒体键/手势翻页）、亮度调节、分享/打开方式直达、收藏文件夹、类型筛选、最近记录进度、中英切换、文件夹浏览器修复（真实文件名解析）、顶部图标排布修复 |

---

*归档路径：`H:\opencode\Subtitle-Reader\SubtitleReader-Android-2.0.zip`*