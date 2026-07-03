package com.subtitle.reader.ui.screens

import android.content.Context
import android.content.Intent
import android.net.Uri
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.background
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.itemsIndexed
import androidx.compose.foundation.lazy.rememberLazyListState
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.TextStyle
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import com.subtitle.reader.model.SubtitleItem
import com.subtitle.reader.parser.SubtitleParser
import com.subtitle.reader.util.DocxExporter
import com.subtitle.reader.util.Strings
import java.io.BufferedReader
import java.io.InputStreamReader
import kotlin.math.abs

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun ReaderScreen() {
    var items by remember { mutableStateOf<List<SubtitleItem>>(emptyList()) }
    var showTime by remember { mutableStateOf(false) }
    var showIndex by remember { mutableStateOf(false) }
    var showDuration by remember { mutableStateOf(false) }
    var fontSize by remember { mutableStateOf(22) }
    var fontFamily by remember { mutableStateOf<FontFamily>(FontFamily.Default) }
    var textColor by remember { mutableStateOf(Color.Black) }
    var bgColor by remember { mutableStateOf(Color.White) }
    var showFontDialog by remember { mutableStateOf(false) }
    var showColorPicker by remember { mutableStateOf(false) }
    var colorPickerTarget by remember { mutableStateOf("text") }
    var showExportDialog by remember { mutableStateOf(false) }
    var showLangMenu by remember { mutableStateOf(false) }
    var showCloseConfirm by remember { mutableStateOf(false) }
    var showRecentFiles by remember { mutableStateOf(false) }
    var fileName by remember { mutableStateOf("") }
    var fileUri by remember { mutableStateOf<Uri?>(null) }
    var langKey by remember { mutableStateOf(Strings.getLanguage()) }
    var pendingRestoreIndex by remember { mutableStateOf(-1) }

    val context = LocalContext.current
    val prefs = context.getSharedPreferences("subtitle_reader", Context.MODE_PRIVATE)
    val lazyListState = rememberLazyListState()

    val filePickerLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.OpenDocument()
    ) { uri ->
        uri?.let {
            savePosition(prefs, fileUri, lazyListState)
            val result = loadSubtitle(context, it)
            if (result != null) {
                items = result.first
                fileName = result.second
                fileUri = it
                addToRecentFiles(prefs, it.toString())
                val savedIdx = loadPosition(prefs, it.toString())
                pendingRestoreIndex = if (savedIdx >= 0) savedIdx else -1
            }
        }
    }

    val exportTxtLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.CreateDocument("text/plain")
    ) { uri ->
        uri?.let { saveExportTxt(context, it, items, showIndex, showTime, showDuration) }
    }

    val exportDocxLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.CreateDocument(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    ) { uri ->
        uri?.let { saveExportDocx(context, it, items, showIndex, showTime, showDuration, fontSize) }
    }

    fun closeFile() {
        savePosition(prefs, fileUri, lazyListState)
        items = emptyList()
        fileName = ""
        fileUri = null
        pendingRestoreIndex = -1
    }

    fun openRecentFile(uriString: String) {
        val uri = Uri.parse(uriString)
        try {
            val takeFlags = Intent.FLAG_GRANT_READ_URI_PERMISSION
            context.contentResolver.takePersistableUriPermission(uri, takeFlags)
        } catch (_: Exception) {}
        savePosition(prefs, fileUri, lazyListState)
        val result = loadSubtitle(context, uri)
        if (result != null) {
            items = result.first
            fileName = result.second
            fileUri = uri
            addToRecentFiles(prefs, uriString)
            val savedIdx = loadPosition(prefs, uriString)
            pendingRestoreIndex = if (savedIdx >= 0) savedIdx else -1
        }
    }

    LaunchedEffect(items) {
        if (items.isNotEmpty() && pendingRestoreIndex >= 0) {
            val idx = pendingRestoreIndex.coerceIn(0, items.size - 1)
            lazyListState.animateScrollToItem(idx)
            pendingRestoreIndex = -1
        }
    }

    val progress = if (items.isEmpty()) 0f
        else (lazyListState.firstVisibleItemIndex.toFloat() / items.size.toFloat()).coerceIn(0f, 1f)

    Scaffold(
        topBar = {
            TopAppBar(
                title = {
                    Text(
                        text = if (fileName.isNotEmpty()) fileName else Strings.get("app_name"),
                        maxLines = 1
                    )
                },
                actions = {
                    Box {
                        IconButton(onClick = { showLangMenu = true }) {
                            Icon(Icons.Default.Language, contentDescription = Strings.get("language"))
                        }
                        if (items.isNotEmpty()) {
                            IconButton(onClick = { showCloseConfirm = true }) {
                                Icon(Icons.Default.Close, contentDescription = Strings.get("close_file"))
                            }
                        }
                        IconButton(onClick = { showRecentFiles = true }) {
                            Icon(Icons.Default.History, contentDescription = Strings.get("recent_files"))
                        }
                        IconButton(onClick = { filePickerLauncher.launch(arrayOf("text/*", "*/*")) }) {
                            Icon(Icons.Default.FolderOpen, contentDescription = Strings.get("open"))
                        }
                    }
                },
                colors = TopAppBarDefaults.topAppBarColors(
                    containerColor = MaterialTheme.colorScheme.primaryContainer
                )
            )
        },
        bottomBar = {
            Surface(
                tonalElevation = 3.dp,
                shadowElevation = 8.dp
            ) {
                Column(modifier = Modifier.fillMaxWidth()) {
                    Row(
                        modifier = Modifier
                            .fillMaxWidth()
                            .padding(horizontal = 8.dp, vertical = 4.dp),
                        horizontalArrangement = Arrangement.SpaceEvenly
                    ) {
                        FilterChip(
                            selected = showTime,
                            onClick = { showTime = !showTime },
                            label = { Text(Strings.get("timecode"), fontSize = 12.sp) },
                            leadingIcon = { Icon(Icons.Default.Schedule, null, modifier = Modifier.size(16.dp)) }
                        )
                        FilterChip(
                            selected = showIndex,
                            onClick = { showIndex = !showIndex },
                            label = { Text(Strings.get("line_no"), fontSize = 12.sp) },
                            leadingIcon = { Icon(Icons.Default.Tag, null, modifier = Modifier.size(16.dp)) }
                        )
                        FilterChip(
                            selected = showDuration,
                            onClick = { showDuration = !showDuration },
                            label = { Text(Strings.get("duration"), fontSize = 12.sp) },
                            leadingIcon = { Icon(Icons.Default.Timer, null, modifier = Modifier.size(16.dp)) }
                        )
                    }
                    Row(
                        modifier = Modifier
                            .fillMaxWidth()
                            .padding(horizontal = 8.dp, vertical = 4.dp),
                        horizontalArrangement = Arrangement.SpaceEvenly
                    ) {
                        TextButton(onClick = { showFontDialog = true }) {
                            Icon(Icons.Default.FontDownload, null, modifier = Modifier.size(18.dp))
                            Spacer(Modifier.width(2.dp))
                            Text(Strings.get("font"), fontSize = 11.sp)
                        }
                        TextButton(onClick = { fontSize = (fontSize - 2).coerceAtLeast(8) }) {
                            Icon(Icons.Default.TextDecrease, null, modifier = Modifier.size(18.dp))
                        }
                        Text("${fontSize}", style = MaterialTheme.typography.bodySmall)
                        TextButton(onClick = { fontSize = (fontSize + 2).coerceAtMost(48) }) {
                            Icon(Icons.Default.TextIncrease, null, modifier = Modifier.size(18.dp))
                        }
                        TextButton(onClick = {
                            colorPickerTarget = "text"
                            showColorPicker = true
                        }) {
                            Icon(Icons.Default.TextFields, null, modifier = Modifier.size(18.dp))
                        }
                        TextButton(onClick = {
                            colorPickerTarget = "bg"
                            showColorPicker = true
                        }) {
                            Icon(Icons.Default.FormatColorFill, null, modifier = Modifier.size(18.dp))
                        }
                        if (items.isNotEmpty()) {
                            TextButton(onClick = { showExportDialog = true }) {
                                Icon(Icons.Default.FileDownload, null, modifier = Modifier.size(18.dp))
                            }
                        }
                    }
                }
            }
        }
    ) { padding ->
        Column(
            modifier = Modifier
                .fillMaxSize()
                .padding(padding)
                .background(bgColor)
        ) {
            if (items.isNotEmpty()) {
                LinearProgressIndicator(
                    progress = progress,
                    modifier = Modifier.fillMaxWidth()
                )
            }
            if (items.isEmpty()) {
                Box(
                    modifier = Modifier.fillMaxSize(),
                    contentAlignment = Alignment.Center
                ) {
                    Column(horizontalAlignment = Alignment.CenterHorizontally) {
                        Icon(
                            Icons.Default.Subtitles,
                            contentDescription = null,
                            modifier = Modifier.size(64.dp),
                            tint = MaterialTheme.colorScheme.onSurfaceVariant
                        )
                        Spacer(modifier = Modifier.height(16.dp))
                        Text(
                            Strings.get("hint"),
                            style = MaterialTheme.typography.bodyLarge,
                            color = MaterialTheme.colorScheme.onSurfaceVariant
                        )
                    }
                }
            } else {
                LazyColumn(
                    state = lazyListState,
                    modifier = Modifier
                        .fillMaxSize()
                        .padding(horizontal = 12.dp, vertical = 8.dp)
                ) {
                    itemsIndexed(items) { _, item ->
                        SubtitleEntry(
                            item = item,
                            showIndex = showIndex,
                            showTime = showTime,
                            showDuration = showDuration,
                            fontSize = fontSize,
                            fontFamily = fontFamily,
                            textColor = textColor
                        )
                    }
                }
            }
        }
    }

    if (showFontDialog) {
        FontPickerDialog(
            current = fontFamily,
            onSelect = { fontFamily = it; showFontDialog = false },
            onDismiss = { showFontDialog = false }
        )
    }

    if (showColorPicker) {
        ColorPickerDialog(
            current = if (colorPickerTarget == "text") textColor else bgColor,
            title = if (colorPickerTarget == "text") Strings.get("text_color_title") else Strings.get("bg_color_title"),
            onSelect = { color ->
                if (colorPickerTarget == "text") textColor = color else bgColor = color
                showColorPicker = false
            },
            onDismiss = { showColorPicker = false }
        )
    }

    if (showExportDialog) {
        ExportDialog(
            fileName = fileName,
            onExportTxt = {
                showExportDialog = false
                if (items.isNotEmpty()) {
                    exportTxtLauncher.launch("${fileName.substringBeforeLast('.')}.txt")
                }
            },
            onExportDocx = {
                showExportDialog = false
                if (items.isNotEmpty()) {
                    exportDocxLauncher.launch("${fileName.substringBeforeLast('.')}.docx")
                }
            },
            onDismiss = { showExportDialog = false }
        )
    }

    if (showLangMenu) {
        LanguageDialog(
            current = langKey,
            onSelect = { lang ->
                Strings.setLanguage(lang)
                langKey = lang
                showLangMenu = false
            },
            onDismiss = { showLangMenu = false }
        )
    }

    if (showCloseConfirm) {
        AlertDialog(
            onDismissRequest = { showCloseConfirm = false },
            title = { Text(Strings.get("close_file")) },
            text = { Text(Strings.get("confirm_close")) },
            confirmButton = {
                TextButton(onClick = {
                    showCloseConfirm = false
                    closeFile()
                }) {
                    Text(Strings.get("yes"))
                }
            },
            dismissButton = {
                TextButton(onClick = { showCloseConfirm = false }) {
                    Text(Strings.get("no"))
                }
            }
        )
    }

    if (showRecentFiles) {
        RecentFilesDialog(
            prefs = prefs,
            onSelect = { uriString ->
                showRecentFiles = false
                openRecentFile(uriString)
            },
            onDismiss = { showRecentFiles = false }
        )
    }
}

@Composable
private fun RecentFilesDialog(
    prefs: android.content.SharedPreferences,
    onSelect: (String) -> Unit,
    onDismiss: () -> Unit
) {
    val recent = getRecentFiles(prefs)
    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text(Strings.get("recent_files")) },
        text = {
            if (recent.isEmpty()) {
                Text(Strings.get("no_recent_files"))
            } else {
                Column {
                    recent.forEach { uriString ->
                        val name = uriString.substringAfterLast('/').substringAfterLast(':')
                            .ifEmpty { uriString }
                        Row(
                            verticalAlignment = Alignment.CenterVertically,
                            modifier = Modifier
                                .fillMaxWidth()
                                .clickable { onSelect(uriString) }
                                .padding(vertical = 8.dp)
                        ) {
                            Icon(
                                Icons.Default.Description,
                                null,
                                modifier = Modifier.size(20.dp),
                                tint = MaterialTheme.colorScheme.onSurfaceVariant
                            )
                            Spacer(Modifier.width(8.dp))
                            Text(name, modifier = Modifier.weight(1f))
                        }
                        Divider()
                    }
                }
            }
        },
        confirmButton = {
            TextButton(onClick = onDismiss) {
                Text(Strings.get("close"))
            }
        }
    )
}

@Composable
private fun SubtitleEntry(
    item: SubtitleItem,
    showIndex: Boolean,
    showTime: Boolean,
    showDuration: Boolean,
    fontSize: Int,
    fontFamily: FontFamily,
    textColor: Color
) {
    Column(modifier = Modifier.padding(vertical = 4.dp)) {
        val metaParts = mutableListOf<String>()
        if (showIndex) metaParts.add("[${item.index}]")
        if (showTime) metaParts.add("${item.start} -> ${item.end}")
        if (showDuration && item.duration.isNotEmpty()) metaParts.add("(dur: ${item.duration})")
        val meta = metaParts.joinToString(" ")

        if (meta.isNotEmpty()) {
            Text(
                text = meta,
                style = TextStyle(
                    fontSize = (fontSize - 4).coerceAtLeast(10).sp,
                    color = textColor.copy(alpha = 0.6f),
                    fontWeight = FontWeight.Light
                )
            )
        }
        Text(
            text = item.text,
            style = TextStyle(
                fontSize = fontSize.sp,
                fontFamily = fontFamily,
                color = textColor
            )
        )
        Divider(
            modifier = Modifier.padding(vertical = 4.dp),
            color = textColor.copy(alpha = 0.1f)
        )
    }
}

private fun loadSubtitle(context: Context, uri: Uri): Pair<List<SubtitleItem>, String>? {
    return try {
        val takeFlags = Intent.FLAG_GRANT_READ_URI_PERMISSION
        try {
            context.contentResolver.takePersistableUriPermission(uri, takeFlags)
        } catch (_: Exception) {}
        val inputStream = context.contentResolver.openInputStream(uri) ?: return null
        val reader = BufferedReader(InputStreamReader(inputStream))
        val content = reader.readText()
        reader.close()
        val items = SubtitleParser.parseFile(uri.toString(), content)
        val name = uri.lastPathSegment?.substringAfterLast('/') ?: uri.lastPathSegment ?: "Unknown"
        Pair(items, name)
    } catch (_: Exception) { null }
}

private fun buildExportText(
    items: List<SubtitleItem>,
    showIndex: Boolean,
    showTime: Boolean,
    showDuration: Boolean
): String {
    val lines = mutableListOf<String>()
    for (item in items) {
        val parts = mutableListOf<String>()
        if (showIndex) parts.add("[${item.index}]")
        if (showTime) parts.add("${item.start} -> ${item.end}")
        if (showDuration && item.duration.isNotEmpty()) parts.add("(dur: ${item.duration})")
        val prefix = parts.joinToString(" ")
        if (prefix.isNotEmpty()) lines.add(prefix)
        lines.add(item.text)
        lines.add("")
    }
    return lines.joinToString("\n")
}

private fun saveExportTxt(
    context: Context,
    uri: Uri,
    items: List<SubtitleItem>,
    showIndex: Boolean,
    showTime: Boolean,
    showDuration: Boolean
) {
    try {
        val content = buildExportText(items, showIndex, showTime, showDuration)
        context.contentResolver.openOutputStream(uri)?.use { out ->
            out.write(content.toByteArray(Charsets.UTF_8))
        }
    } catch (_: Exception) {}
}

private fun saveExportDocx(
    context: Context,
    uri: Uri,
    items: List<SubtitleItem>,
    showIndex: Boolean,
    showTime: Boolean,
    showDuration: Boolean,
    fontSize: Int
) {
    try {
        val content = buildExportText(items, showIndex, showTime, showDuration)
        val bytes = DocxExporter.createDocx(content)
        context.contentResolver.openOutputStream(uri)?.use { out ->
            out.write(bytes)
        }
    } catch (_: Exception) {}
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
private fun FontPickerDialog(
    current: FontFamily,
    onSelect: (FontFamily) -> Unit,
    onDismiss: () -> Unit
) {
    val fonts = listOf(
        "Default" to FontFamily.Default,
        "Serif" to FontFamily.Serif,
        "SansSerif" to FontFamily.SansSerif,
        "Monospace" to FontFamily.Monospace,
        "Cursive" to FontFamily.Cursive
    )
    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text(Strings.get("select_font")) },
        text = {
            Column {
                fonts.forEach { (name, family) ->
                    Row(
                        verticalAlignment = Alignment.CenterVertically,
                        modifier = Modifier
                            .fillMaxWidth()
                            .padding(vertical = 4.dp)
                    ) {
                        RadioButton(
                            selected = current == family,
                            onClick = { onSelect(family) }
                        )
                        Spacer(modifier = Modifier.width(8.dp))
                        Text(name, fontFamily = family)
                    }
                }
            }
        },
        confirmButton = { TextButton(onClick = onDismiss) { Text(Strings.get("close")) } }
    )
}

@Composable
private fun ColorPickerDialog(
    current: Color,
    title: String,
    onSelect: (Color) -> Unit,
    onDismiss: () -> Unit
) {
    val presetColors = listOf(
        Color.Black, Color.White, Color.Red, Color(0xFF2196F3),
        Color(0xFF4CAF50), Color(0xFFFF9800), Color(0xFF9C27B0),
        Color(0xFF795548), Color(0xFF607D8B), Color(0xFFE91E63),
        Color(0xFF00BCD4), Color(0xFF8BC34A), Color(0xFFFFC107),
        Color.Gray, Color(0xFF3F51B5), Color(0xFF009688)
    )
    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text(title) },
        text = {
            Column {
                Text(Strings.get("current"), style = MaterialTheme.typography.bodySmall)
                Box(
                    modifier = Modifier
                        .fillMaxWidth()
                        .height(32.dp)
                        .padding(vertical = 4.dp)
                        .background(current)
                )
                Spacer(modifier = Modifier.height(12.dp))
                Text(Strings.get("presets"), style = MaterialTheme.typography.bodySmall)
                Spacer(modifier = Modifier.height(4.dp))
                Column {
                    presetColors.chunked(4).forEach { row ->
                        Row(
                            modifier = Modifier.fillMaxWidth(),
                            horizontalArrangement = Arrangement.SpaceEvenly
                        ) {
                            row.forEach { color ->
                                Box(
                                    modifier = Modifier
                                        .size(40.dp)
                                        .background(color, shape = MaterialTheme.shapes.small)
                                        .then(
                                            if (current == color) Modifier.padding(2.dp) else Modifier
                                        ),
                                    contentAlignment = Alignment.Center
                                ) {
                                    if (current == color) {
                                        Icon(
                                            Icons.Default.Check,
                                            null,
                                            tint = if (color == Color.Black) Color.White else Color.Black,
                                            modifier = Modifier.size(20.dp)
                                        )
                                    }
                                }
                            }
                        }
                        Spacer(modifier = Modifier.height(8.dp))
                    }
                }
            }
        },
        confirmButton = {},
        dismissButton = { TextButton(onClick = onDismiss) { Text(Strings.get("close")) } }
    )
}

@Composable
private fun ExportDialog(
    fileName: String,
    onExportTxt: () -> Unit,
    onExportDocx: () -> Unit,
    onDismiss: () -> Unit
) {
    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text(Strings.get("export_title")) },
        text = {
            Column {
                Text("${Strings.get("file")} ${fileName.orEmpty()}", style = MaterialTheme.typography.bodyMedium)
                Spacer(modifier = Modifier.height(16.dp))
                Text(Strings.get("choose_format"), style = MaterialTheme.typography.bodySmall)
                Spacer(modifier = Modifier.height(8.dp))
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.SpaceEvenly
                ) {
                    OutlinedButton(onClick = onExportTxt) {
                        Icon(Icons.Default.Description, null, modifier = Modifier.size(18.dp))
                        Spacer(Modifier.width(4.dp))
                        Text(Strings.get("export_txt"))
                    }
                    OutlinedButton(onClick = onExportDocx) {
                        Icon(Icons.Default.TextSnippet, null, modifier = Modifier.size(18.dp))
                        Spacer(Modifier.width(4.dp))
                        Text(Strings.get("export_docx"))
                    }
                }
            }
        },
        confirmButton = { TextButton(onClick = onDismiss) { Text(Strings.get("cancel")) } }
    )
}

@Composable
private fun LanguageDialog(
    current: String,
    onSelect: (String) -> Unit,
    onDismiss: () -> Unit
) {
    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text(Strings.get("language")) },
        text = {
            Column {
                Row(
                    verticalAlignment = Alignment.CenterVertically,
                    modifier = Modifier.fillMaxWidth().padding(vertical = 4.dp)
                ) {
                    RadioButton(
                        selected = current == "zh",
                        onClick = { onSelect("zh") }
                    )
                    Spacer(Modifier.width(8.dp))
                    Text(Strings.get("lang_zh"))
                }
                Row(
                    verticalAlignment = Alignment.CenterVertically,
                    modifier = Modifier.fillMaxWidth().padding(vertical = 4.dp)
                ) {
                    RadioButton(
                        selected = current == "en",
                        onClick = { onSelect("en") }
                    )
                    Spacer(Modifier.width(8.dp))
                    Text(Strings.get("lang_en"))
                }
            }
        },
        confirmButton = { TextButton(onClick = onDismiss) { Text(Strings.get("close")) } }
    )
}

private fun savePosition(prefs: android.content.SharedPreferences, uri: Uri?, state: androidx.compose.foundation.lazy.LazyListState) {
    if (uri == null) return
    val key = "pos_" + abs(uri.toString().hashCode())
    prefs.edit().putInt(key, state.firstVisibleItemIndex).apply()
}

private fun loadPosition(prefs: android.content.SharedPreferences, uriString: String): Int {
    val key = "pos_" + abs(uriString.hashCode())
    return prefs.getInt(key, -1)
}

private fun getRecentFiles(prefs: android.content.SharedPreferences): List<String> {
    val json = prefs.getString("recent_files", "[]") ?: "[]"
    return try {
        val arr = org.json.JSONArray(json)
        (0 until arr.length()).map { arr.getString(it) }
    } catch (_: Exception) { emptyList() }
}

private fun addToRecentFiles(prefs: android.content.SharedPreferences, uriString: String) {
    val list = getRecentFiles(prefs).toMutableList()
    list.remove(uriString)
    list.add(0, uriString)
    val trimmed = list.take(10)
    val arr = org.json.JSONArray(trimmed)
    prefs.edit().putString("recent_files", arr.toString()).apply()
}
