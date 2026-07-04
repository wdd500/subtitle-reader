# Subtitle Reader / 字幕阅读器

A bilingual (Chinese/English) subtitle reader for desktop (Windows) and Android. Supports SRT, SUB (MicroDVD), and ASS/SSA subtitle formats.
双语言（中文/英文）字幕阅读器，支持桌面版（Windows）和 Android 版。支持 SRT、SUB（MicroDVD）和 ASS/SSA 字幕格式。

<img width="1202" height="1288" alt="1" src="https://github.com/user-attachments/assets/219b653b-d9c6-4f14-ab47-266ba5f2f49e" />
<img width="1202" height="1288" alt="2" src="https://github.com/user-attachments/assets/7de7a750-525a-4bad-b379-3b52b0951f66" />


## Features
## Features / 功能

- Open subtitle files (.srt / .sub / .ass)
  打开字幕文件
- Toggle display: timecode, line number, duration
  切换显示：时间码、行号、时长
- Change font family, font size, text color, background color
  更换字体、字号、文字颜色、背景颜色
- Export visible content as .txt or .docx
  导出为文本或 Word 文档
- Switch between Chinese and English UI instantly
- Remembers last directory (desktop)
  一键切换中英文界面
- Reading progress bar with percentage/page mode
  阅读进度条（百分比/分页模式）
- Recent files with saved progress
  最近打开文件及进度记忆

## Desktop (Python / tkinter)
## Desktop / 桌面版 (Python / tkinter)

- **Source:** `desktop/subtitle_reader.py`
- **Pre-built:** `desktop/dist/SubtitleReader.exe` (~18.8 MB, PyInstaller --onefile --windowed)
- **Requirements:** Python 3.x, python-docx (for .docx export)
- **Run:** `python desktop/subtitle_reader.py` or double-click the .exe
- **Source / 源码:** `desktop/subtitle_reader.py`
- **Requirements / 运行依赖:** Python 3.x, python-docx (for .docx export)
- **Run / 运行:** `python desktop/subtitle_reader.py`

### Build from source
### Build EXE / 打包可执行文件
```
pip install pyinstaller python-docx

pyinstaller --onefile --windowed --name subtitle_reader --icon=icon.ico subtitle_reader.py

## Android (Kotlin / Jetpack Compose)

- **Source:** `android/` (standard Gradle project)
- **Requirements:** Android Studio, Android SDK (API 24+), JDK 17+
- **Source / 源码:** `android/` (standard Gradle project)
- **Requirements / 运行要求:** Android Studio, Android SDK (API 24+), JDK 17+

### Build APK
### Build APK / 构建安装包
```bash
cd android
./gradlew assembleDebug
# APK at: app/build/outputs/apk/debug/app-debug.apk
```

## Usage
## Usage / 使用方法

1. Launch the app on your platform
2. Use **File** (desktop) / **Open** button (Android) to load a .srt, .sub, or .ass file
3. Adjust display toggles and styling via the toolbar
4. Use **File → Export** (desktop) or **Export** button (Android) to save as .txt or .docx
5. Switch language via **Language** menu (desktop) or the language icon in the top bar (Android)
1. Launch the app on your platform / 在对应平台启动应用
2. Tap **Open** button to load a .srt, .sub, or .ass file / 点击打开按钮加载字幕文件
3. Adjust display toggles and styling via the bottom toolbar / 通过底部工具栏调整显示选项和样式
4. Use Export button to save as .txt or .docx / 点击导出按钮保存为文本或 Word 文档
5. Switch language via the language icon in the top bar / 点击顶栏语言图标切换中英文

## License / 许可证

MIT
