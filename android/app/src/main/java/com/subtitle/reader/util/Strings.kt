package com.subtitle.reader.util

object Strings {
    private val ZH = mapOf(
        "app_name" to "字幕阅读器",
        "open" to "打开",
        "export" to "导出",
        "timecode" to "时间码",
        "line_no" to "行号",
        "duration" to "时长",
        "font" to "字体",
        "text_color" to "文字颜色",
        "bg_color" to "背景颜色",
        "no_file" to "未加载字幕。\n打开 .srt/.sub/.ass 文件开始阅读。",
        "export_title" to "导出字幕",
        "export_txt" to "文本 (.txt)",
        "export_docx" to "Word (.docx)",
        "choose_format" to "选择格式：",
        "cancel" to "取消",
        "select_font" to "选择字体",
        "close" to "关闭",
        "text_color_title" to "文字颜色",
        "bg_color_title" to "背景颜色",
        "presets" to "预设颜色：",
        "current" to "当前：",
        "file" to "文件：",
        "export_hint" to "打开一个字幕文件后可使用此功能",
        "language" to "语言",
        "lang_zh" to "中文",
        "lang_en" to "English",
        "hint" to "打开字幕文件 (.srt / .sub / .ass)",
        "close_file" to "关闭文件",
        "recent_files" to "最近打开",
        "no_recent_files" to "没有最近打开的文件",
        "confirm_close" to "确定关闭当前文件？",
        "yes" to "确定",
        "no" to "取消",
    )

    private val EN = mapOf(
        "app_name" to "Subtitle Reader",
        "open" to "Open",
        "export" to "Export",
        "timecode" to "Timecode",
        "line_no" to "Line No.",
        "duration" to "Duration",
        "font" to "Font",
        "text_color" to "Text Color",
        "bg_color" to "Bg Color",
        "no_file" to "No subtitles loaded.\nOpen a .srt / .sub / .ass file to begin.",
        "export_title" to "Export Subtitle",
        "export_txt" to "Text (.txt)",
        "export_docx" to "Word (.docx)",
        "choose_format" to "Choose format:",
        "cancel" to "Cancel",
        "select_font" to "Select Font",
        "close" to "Close",
        "text_color_title" to "Text Color",
        "bg_color_title" to "Background Color",
        "presets" to "Presets:",
        "current" to "Current:",
        "file" to "File:",
        "export_hint" to "Open a subtitle file first",
        "language" to "Language",
        "lang_zh" to "中文",
        "lang_en" to "English",
        "hint" to "Open a subtitle file (.srt / .sub / .ass)",
        "close_file" to "Close File",
        "recent_files" to "Recent Files",
        "no_recent_files" to "No recent files",
        "confirm_close" to "Close current file?",
        "yes" to "Yes",
        "no" to "No",
    )

    private var currentLang = "zh"

    fun setLanguage(lang: String) {
        currentLang = lang
    }

    fun getLanguage(): String = currentLang

    fun get(key: String): String {
        val map = if (currentLang == "zh") ZH else EN
        return map[key] ?: key
    }
}
