package com.subtitle.reader.model

data class SubtitleItem(
    val index: Int = 0,
    val start: String = "",
    val end: String = "",
    val text: String = "",
    val duration: String = ""
)
