# Subtitle Reader

A bilingual (Chinese/English) subtitle reader for desktop (Windows) and Android. Supports SRT, SUB (MicroDVD), and ASS/SSA subtitle formats.

## Features

- Open subtitle files (.srt / .sub / .ass)
- Toggle display: timecode, line number, duration
- Change font family, font size, text color, background color
- Export visible content as .txt or .docx
- Switch between Chinese and English UI instantly
- Remembers last directory (desktop)

## Desktop (Python / tkinter)

- **Source:** `desktop/subtitle_reader.py`
- **Pre-built:** `desktop/dist/SubtitleReader.exe` (~18.8 MB, PyInstaller --onefile --windowed)
- **Requirements:** Python 3.x, python-docx (for .docx export)
- **Run:** `python desktop/subtitle_reader.py` or double-click the .exe

### Build from source
```
pip install pyinstaller python-docx
pyinstaller --onefile --windowed --name SubtitleReader desktop/subtitle_reader.py
```

## Android (Kotlin / Jetpack Compose)

- **Source:** `android/` (standard Gradle project)
- **Requirements:** Android Studio, Android SDK (API 24+), JDK 17+

### Build APK
```bash
cd android
./gradlew assembleDebug
# APK at: app/build/outputs/apk/debug/app-debug.apk
```

## Usage

1. Launch the app on your platform
2. Use **File** (desktop) / **Open** button (Android) to load a .srt, .sub, or .ass file
3. Adjust display toggles and styling via the toolbar
4. Use **File → Export** (desktop) or **Export** button (Android) to save as .txt or .docx
5. Switch language via **Language** menu (desktop) or the language icon in the top bar (Android)
