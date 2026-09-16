package com.subtitle.reader.parser

import com.subtitle.reader.model.SubtitleItem

object SubtitleParser {

    fun parseFile(path: String, content: String): List<SubtitleItem> {
        return when {
            path.endsWith(".srt", ignoreCase = true) -> parseSrt(content)
            path.endsWith(".sub", ignoreCase = true) -> parseSub(content)
            path.endsWith(".ass", ignoreCase = true) || path.endsWith(".ssa", ignoreCase = true) -> parseAss(content)
            else -> emptyList()
        }
    }

    fun parseSrt(content: String): List<SubtitleItem> {
        val items = mutableListOf<SubtitleItem>()
        val blocks = content.trim().split(Regex("\\n\\s*\\n"))
        for (block in blocks) {
            val lines = block.trim().split("\n")
            if (lines.size < 3) continue
            try {
                val idx = lines[0].toIntOrNull() ?: continue
                val timeMatch = Regex(
                    """(\d+:\d+:\d+[,.]\d+)\s*-->\s*(\d+:\d+:\d+[,.]\d+)"""
                ).find(lines[1]) ?: continue
                val start = timeMatch.groupValues[1].replace(',', '.')
                val end = timeMatch.groupValues[2].replace(',', '.')
                val text = lines.drop(2).joinToString("\n")
                items.add(SubtitleItem(idx, start, end, text, calcDuration(start, end)))
            } catch (_: Exception) {}
        }
        return items
    }

    fun parseSub(content: String): List<SubtitleItem> {
        val items = mutableListOf<SubtitleItem>()
        val lines = content.trim().split("\n")
        for ((i, line) in lines.withIndex()) {
            val match = Regex("""\{(\d+)\}\{(\d+)\}(.*)""").find(line.trim()) ?: continue
            val start = match.groupValues[1]
            val end = match.groupValues[2]
            val text = match.groupValues[3].replace("|", "\n")
            val dur = "${end.toInt() - start.toInt()} frames"
            items.add(SubtitleItem(i + 1, start, end, text, dur))
        }
        return items
    }

    fun parseAss(content: String): List<SubtitleItem> {
        val items = mutableListOf<SubtitleItem>()
        val lines = content.split("\n")
        var inEvents = false
        var formatKeys = emptyList<String>()

        for (line in lines) {
            val trimmed = line.trim()
            if (trimmed.startsWith("[Events]")) {
                inEvents = true
                continue
            }
            if (inEvents) {
                if (trimmed.startsWith("Format:")) {
                    formatKeys = trimmed.removePrefix("Format:").split(",").map { it.trim() }
                    continue
                }
                if (trimmed.startsWith("Dialogue:")) {
                    val data = splitAssValue(trimmed.removePrefix("Dialogue:"), formatKeys)
                    val text = data["Text"]?.replace(Regex("""\{[^}]*\}"""), "")
                        ?.replace("\\N", "\n")?.replace("\\n", "\n") ?: ""
                    items.add(SubtitleItem(
                        index = items.size + 1,
                        start = data["Start"] ?: "",
                        end = data["End"] ?: "",
                        text = text,
                        duration = calcDuration(data["Start"] ?: "", data["End"] ?: "")
                    ))
                }
            }
        }
        return items
    }

    private fun splitAssValue(raw: String, keys: List<String>): Map<String, String> {
        val parts = mutableListOf<String>()
        var cur = StringBuilder()
        var braces = 0
        for (ch in raw) {
            when (ch) {
                '{' -> { braces++; cur.append(ch) }
                '}' -> { braces--; cur.append(ch) }
                ',' -> if (braces == 0) { parts.add(cur.toString().trim()); cur = StringBuilder() }
                else -> cur.append(ch)
            }
        }
        parts.add(cur.toString().trim())
        val map = mutableMapOf<String, String>()
        for (i in keys.indices) {
            if (i < parts.size) map[keys[i]] = parts[i]
        }
        return map
    }

    private fun calcDuration(start: String, end: String): String {
        fun toSec(t: String): Double {
            val parts = t.replace(',', '.').split(":")
            return parts[0].toDouble() * 3600 + parts[1].toDouble() * 60 + parts[2].toDouble()
        }
        return try {
            String.format("%.2fs", toSec(end) - toSec(start))
        } catch (_: Exception) { "" }
    }
}
