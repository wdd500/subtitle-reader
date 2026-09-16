package com.subtitle.reader.ui.screens

import android.app.Activity
import android.content.Context
import android.content.Intent
import android.net.Uri
import android.os.Build
import android.provider.DocumentsContract
import android.provider.OpenableColumns
import android.view.WindowManager
import androidx.activity.compose.BackHandler
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.background
import androidx.compose.foundation.clickable
import androidx.compose.foundation.gestures.detectTapGestures
import androidx.compose.foundation.horizontalScroll
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.itemsIndexed
import androidx.compose.foundation.lazy.rememberLazyListState
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.verticalScroll
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.input.pointer.pointerInput
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.platform.LocalView
import androidx.compose.ui.text.TextStyle
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.compose.ui.window.Dialog
import androidx.compose.ui.window.DialogProperties
import androidx.core.view.WindowCompat
import androidx.core.view.WindowInsetsCompat
import androidx.core.view.WindowInsetsControllerCompat
import com.subtitle.reader.ReaderEvents
import com.subtitle.reader.model.SubtitleItem
import com.subtitle.reader.parser.SubtitleParser
import com.subtitle.reader.util.DocxExporter
import com.subtitle.reader.util.Strings
import java.io.BufferedReader
import java.io.File
import java.io.InputStreamReader
import java.net.URLDecoder
import kotlinx.coroutines.launch
import kotlin.math.abs

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun ReaderScreen(initialUri: String? = null) {
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
    
    var showRecentFiles by remember { mutableStateOf(false) }
    var fileName by remember { mutableStateOf("") }
    var fileUri by remember { mutableStateOf<Uri?>(null) }
    var langKey by remember { mutableStateOf(Strings.getLanguage()) }
    var pendingRestoreIndex by remember { mutableStateOf(-1) }
    var showJumpDialog by remember { mutableStateOf(false) }
    var showAboutDialog by remember { mutableStateOf(false) }
    var showClearConfirm by remember { mutableStateOf(false) }
    var settingsVisible by remember { mutableStateOf(true) }
    var showFolderBrowser by remember { mutableStateOf(false) }
    var folderBrowserUri by remember { mutableStateOf<Uri?>(null) }
    val context = LocalContext.current
    val view = LocalView.current
    val prefs = context.getSharedPreferences("subtitle_reader", Context.MODE_PRIVATE)
    val scope = rememberCoroutineScope()
    val lazyListState = rememberLazyListState()
    var recentFiles by remember { mutableStateOf(getRecentFiles(prefs)) }
    var displayMode by remember { mutableStateOf(0) } // 0=percent, 1=page
    var brightness by remember { mutableFloatStateOf(prefs.getFloat("brightness", -1f)) }

    LaunchedEffect(brightness) {
        val activity = context as? Activity
        if (activity != null) {
            val window = activity.window
            window.attributes = window.attributes.apply {
                screenBrightness = if (brightness < 0f) {
                    WindowManager.LayoutParams.BRIGHTNESS_OVERRIDE_NONE
                } else {
                    brightness.coerceIn(0.05f, 1f)
                }
            }
            prefs.edit().putFloat("brightness", brightness).apply()
        }
    }

    fun scrollByPage(forward: Boolean) {
        scope.launch {
            val visible = lazyListState.layoutInfo.visibleItemsInfo
            val ipp = if (visible.isNotEmpty())
                (visible.last().index - visible.first().index + 1).coerceAtLeast(1) else 1
            val cur = lazyListState.firstVisibleItemIndex
            val target = if (forward) {
                (cur + ipp).coerceAtMost(items.size - 1)
            } else {
                (cur - ipp).coerceAtLeast(0)
            }
            lazyListState.animateScrollToItem(target)
        }
    }

    LaunchedEffect(settingsVisible, items.isNotEmpty()) {
        val activity = context as? Activity
        if (activity != null) {
            val window = activity.window
            val insetsController = WindowCompat.getInsetsController(window, view)
            if (items.isNotEmpty() && !settingsVisible) {
                insetsController.systemBarsBehavior = WindowInsetsControllerCompat.BEHAVIOR_SHOW_TRANSIENT_BARS_BY_SWIPE
                insetsController.hide(WindowInsetsCompat.Type.systemBars())
                if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.P) {
                    window.attributes = window.attributes.apply {
                        layoutInDisplayCutoutMode = WindowManager.LayoutParams.LAYOUT_IN_DISPLAY_CUTOUT_MODE_SHORT_EDGES
                    }
                }
            } else {
                insetsController.show(WindowInsetsCompat.Type.systemBars())
                if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.P) {
                    window.attributes = window.attributes.apply {
                        layoutInDisplayCutoutMode = WindowManager.LayoutParams.LAYOUT_IN_DISPLAY_CUTOUT_MODE_DEFAULT
                    }
                }
            }
        }
    }

    DisposableEffect(settingsVisible, items.isNotEmpty()) {
        val active = items.isNotEmpty() && !settingsVisible
        ReaderEvents.volumeKeyPaging = active
        if (active) {
            ReaderEvents.onPagePrev = { scrollByPage(false) }
            ReaderEvents.onPageNext = { scrollByPage(true) }
        } else {
            ReaderEvents.onPagePrev = null
            ReaderEvents.onPageNext = null
        }
        onDispose {
            ReaderEvents.volumeKeyPaging = false
            ReaderEvents.onPagePrev = null
            ReaderEvents.onPageNext = null
        }
    }

    fun applyLoaded(uri: Uri, result: Pair<List<SubtitleItem>, String>, uriString: String) {
        items = result.first
        fileName = result.second
        fileUri = uri
        addToRecentFiles(prefs, uriString)
        val savedIdx = loadPosition(prefs, uriString)
        pendingRestoreIndex = if (savedIdx >= 0) savedIdx else -1
        settingsVisible = false
    }

    val folderPickerLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.OpenDocumentTree()
    ) { uri ->
        uri?.let {
            try {
                context.contentResolver.takePersistableUriPermission(
                    it,
                    Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_GRANT_WRITE_URI_PERMISSION
                )
            } catch (_: Exception) {}
            saveLastFolder(prefs, it.toString())
            folderBrowserUri = it
            showFolderBrowser = true
        }
    }

    fun openFolder() {
        val last = loadLastFolder(prefs)
        if (last != null) {
            folderBrowserUri = Uri.parse(last)
            showFolderBrowser = true
        } else {
            folderPickerLauncher.launch(null)
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
        savePosition(prefs, fileUri, lazyListState, items.size)
        items = emptyList()
        fileName = ""
        fileUri = null
        pendingRestoreIndex = -1
        settingsVisible = true
    }

    fun openRecentFile(uriString: String) {
        val uri = Uri.parse(uriString)
        try {
            val takeFlags = Intent.FLAG_GRANT_READ_URI_PERMISSION
            context.contentResolver.takePersistableUriPermission(uri, takeFlags)
        } catch (_: Exception) {}
        savePosition(prefs, fileUri, lazyListState, items.size)
        val name = uri.lastPathSegment?.substringAfterLast('/') ?: "Unknown"
        val result = loadSubtitle(context, uri, name)
        if (result != null) {
            applyLoaded(uri, result, uriString)
        }
    }

    fun openFromFolder(uri: Uri, name: String) {
        savePosition(prefs, fileUri, lazyListState, items.size)
        val result = loadSubtitle(context, uri, name)
        if (result != null && result.first.isNotEmpty()) {
            applyLoaded(uri, result, uri.toString())
            showFolderBrowser = false
        } else {
            android.widget.Toast.makeText(context, Strings.get("load_fail"), android.widget.Toast.LENGTH_SHORT).show()
        }
    }

    LaunchedEffect(items) {
        if (items.isNotEmpty() && pendingRestoreIndex >= 0) {
            val idx = pendingRestoreIndex.coerceIn(0, items.size - 1)
            lazyListState.animateScrollToItem(idx)
            pendingRestoreIndex = -1
        }
    }

    LaunchedEffect(initialUri) {
        val uriStr = initialUri ?: return@LaunchedEffect
        if (uriStr == fileUri?.toString()) return@LaunchedEffect
        val uri = Uri.parse(uriStr)
        savePosition(prefs, fileUri, lazyListState, items.size)
        val result = loadSubtitle(context, uri, resolveDisplayName(context, uri))
        if (result != null && result.first.isNotEmpty()) {
            applyLoaded(uri, result, uriStr)
        } else {
            android.widget.Toast.makeText(context, Strings.get("load_fail"), android.widget.Toast.LENGTH_SHORT).show()
        }
    }

    val progress = if (items.isEmpty()) 0f
        else (lazyListState.firstVisibleItemIndex.toFloat() / items.size.toFloat()).coerceIn(0f, 1f)

    val pageInfo = if (items.isNotEmpty() && displayMode == 1) {
        val layoutInfo = lazyListState.layoutInfo
        val visibleItems = layoutInfo.visibleItemsInfo
        if (visibleItems.isNotEmpty()) {
            val itemsPerPage = (visibleItems.last().index - visibleItems.first().index + 1).coerceAtLeast(1)
            val totalPages = (items.size + itemsPerPage - 1) / itemsPerPage
            val currentPage = (visibleItems.first().index / itemsPerPage) + 1
            Pair(currentPage.coerceAtMost(totalPages), totalPages)
        } else null
    } else null

    Scaffold(
        topBar = {
            if (settingsVisible) {
                TopAppBar(
                    title = {
                        Text(
                            text = if (fileName.isNotEmpty()) fileName else Strings.get("app_name"),
                            maxLines = 1
                        )
                    },
                    actions = {
                        Row {
                            IconButton(onClick = { showAboutDialog = true }) {
                                Icon(Icons.Default.Info, contentDescription = Strings.get("menu_about"))
                            }
                            IconButton(onClick = { showLangMenu = true }) {
                                Icon(Icons.Default.Language, contentDescription = Strings.get("language"))
                            }
                            IconButton(onClick = { showRecentFiles = true }) {
                                Icon(Icons.Default.History, contentDescription = Strings.get("recent_files"))
                            }
                            IconButton(onClick = { openFolder() }) {
                                Icon(Icons.Default.Folder, contentDescription = Strings.get("open_folder"))
                            }
                            if (items.isNotEmpty()) {
                                IconButton(onClick = { closeFile() }) {
                                    Icon(Icons.Default.Close, contentDescription = Strings.get("close_file"))
                                }
                            }
                        }
                    },
                    colors = TopAppBarDefaults.topAppBarColors(
                        containerColor = MaterialTheme.colorScheme.primaryContainer
                    )
                )
            }
        },
        bottomBar = {
            if (settingsVisible) {
                Surface(
                    tonalElevation = 3.dp,
                    shadowElevation = 8.dp,
                    modifier = Modifier.navigationBarsPadding()
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
                            FilterChip(
                                selected = displayMode == 1,
                                onClick = { displayMode = if (displayMode == 1) 0 else 1 },
                                label = {
                                    Text(
                                        if (displayMode == 1) Strings.get("mode_page") else Strings.get("mode_percent"),
                                        fontSize = 12.sp
                                    )
                                },
                                leadingIcon = { Icon(Icons.Default.ViewCarousel, null, modifier = Modifier.size(16.dp)) }
                            )
                        }
                        Row(
                            modifier = Modifier
                                .fillMaxWidth()
                                .padding(horizontal = 8.dp, vertical = 4.dp),
                            horizontalArrangement = Arrangement.SpaceEvenly,
                            verticalAlignment = Alignment.CenterVertically
                        ) {
                            TextButton(onClick = { showFontDialog = true }) {
                                Icon(Icons.Default.FontDownload, null, modifier = Modifier.size(18.dp))
                                Spacer(Modifier.width(2.dp))
                                Text(Strings.get("font"), fontSize = 11.sp)
                            }
                            TextButton(onClick = { fontSize = (fontSize - 2).coerceAtLeast(8) }) {
                                Icon(Icons.Default.TextDecrease, null, modifier = Modifier.size(18.dp))
                            }
                            Box(
                                modifier = Modifier
                                    .width(36.dp)
                                    .height(40.dp),
                                contentAlignment = Alignment.Center
                            ) {
                                Text(
                                    "${fontSize}",
                                    style = MaterialTheme.typography.bodyMedium,
                                    fontWeight = FontWeight.Bold,
                                    textAlign = TextAlign.Center
                                )
                            }
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
                        Row(
                            modifier = Modifier
                                .fillMaxWidth()
                                .padding(horizontal = 8.dp, vertical = 2.dp),
                            verticalAlignment = Alignment.CenterVertically
                        ) {
                            Icon(
                                Icons.Default.BrightnessHigh,
                                null,
                                modifier = Modifier.size(18.dp),
                                tint = MaterialTheme.colorScheme.onSurfaceVariant
                            )
                            Spacer(Modifier.width(4.dp))
                            Text(Strings.get("brightness"), fontSize = 12.sp)
                            Spacer(Modifier.width(4.dp))
                            Slider(
                                value = if (brightness >= 0f) brightness else 1f,
                                onValueChange = { brightness = it.coerceIn(0.05f, 1f) },
                                valueRange = 0.05f..1f,
                                modifier = Modifier.weight(1f)
                            )
                            IconButton(onClick = { brightness = -1f }) {
                                Icon(
                                    Icons.Default.BrightnessAuto,
                                    null,
                                    modifier = Modifier.size(18.dp),
                                    tint = if (brightness < 0f) MaterialTheme.colorScheme.primary
                                        else MaterialTheme.colorScheme.onSurfaceVariant
                                )
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
                .pointerInput(items.isNotEmpty()) {
                    detectTapGestures { offset ->
                        if (items.isNotEmpty()) {
                            val h = size.height
                            val w = size.width
                            val isRightZone = offset.x > w * 2f / 3f
                            val inMiddle = offset.y in h / 3f..h * 2f / 3f
                            if (isRightZone && inMiddle) {
                                scope.launch {
                                    val visible = lazyListState.layoutInfo.visibleItemsInfo
                                    val ipp = if (visible.isNotEmpty())
                                        (visible.last().index - visible.first().index + 1).coerceAtLeast(1) else 1
                                    val cur = lazyListState.firstVisibleItemIndex
                                    val target = if (offset.y < h / 2f) {
                                        (cur - ipp).coerceAtLeast(0)
                                    } else {
                                        (cur + ipp).coerceAtMost(items.size - 1)
                                    }
                                    lazyListState.animateScrollToItem(target)
                                }
                            } else if (!isRightZone && inMiddle) {
                                settingsVisible = !settingsVisible
                            }
                        }
                    }
                }
        ) {
            if (items.isNotEmpty()) {
                Row(
                    modifier = Modifier
                        .fillMaxWidth()
                        .pointerInput(displayMode) {
                            detectTapGestures { offset ->
                                scope.launch {
                                    val fraction = (offset.x / size.width.toFloat()).coerceIn(0f, 1f)
                                    if (displayMode == 1) {
                                        val layoutInfo = lazyListState.layoutInfo
                                        val visibleItems = layoutInfo.visibleItemsInfo
                                        if (visibleItems.isNotEmpty()) {
                                            val itemsPerPage = (visibleItems.last().index - visibleItems.first().index + 1).coerceAtLeast(1)
                                            val targetPage = (fraction * (items.size / itemsPerPage)).toInt().coerceIn(0, (items.size / itemsPerPage).coerceAtLeast(0))
                                            lazyListState.animateScrollToItem((targetPage * itemsPerPage).coerceAtMost(items.size - 1))
                                        }
                                    } else {
                                        val targetIdx = (fraction * items.size).toInt().coerceIn(0, items.size - 1)
                                        lazyListState.animateScrollToItem(targetIdx)
                                    }
                                }
                            }
                        },
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    LinearProgressIndicator(
                        progress = if (pageInfo != null) pageInfo.first.toFloat() / pageInfo.second.toFloat() else progress,
                        modifier = Modifier.weight(1f)
                    )
                    Text(
                        text = if (pageInfo != null) {
                            Strings.get("page_fmt")
                                .replaceFirst("{}", pageInfo.first.toString())
                                .replaceFirst("{}", pageInfo.second.toString())
                        } else "${(progress * 100).toInt()}%",
                        modifier = Modifier
                            .padding(start = 8.dp, end = 4.dp)
                            .clickable { showJumpDialog = true },
                        style = MaterialTheme.typography.bodyMedium,
                        fontWeight = FontWeight.Bold
                    )
                }
            }
            if (items.isEmpty()) {
                Column(
                    modifier = Modifier
                        .fillMaxSize()
                        .clickable { openFolder() },
                    horizontalAlignment = Alignment.CenterHorizontally
                ) {
                    Spacer(modifier = Modifier.weight(1f))
                    Icon(
                        Icons.Default.Subtitles,
                        contentDescription = null,
                        modifier = Modifier.size(64.dp),
                        tint = MaterialTheme.colorScheme.onSurfaceVariant
                    )
                    Spacer(modifier = Modifier.height(16.dp))
                    Text(
                        Strings.get("open_folder_hint"),
                        style = MaterialTheme.typography.bodyLarge,
                        color = MaterialTheme.colorScheme.onSurfaceVariant
                    )
                    if (recentFiles.isNotEmpty()) {
                        Spacer(modifier = Modifier.weight(0.5f))
                        Divider(modifier = Modifier.padding(horizontal = 32.dp))
                        Spacer(modifier = Modifier.height(8.dp))
                        TextButton(onClick = { showClearConfirm = true }) {
                            Text(Strings.get("clear_recent"))
                        }
                        Spacer(modifier = Modifier.height(8.dp))
                        Text(
                            Strings.get("recent_files"),
                            style = MaterialTheme.typography.titleSmall,
                            color = MaterialTheme.colorScheme.onSurfaceVariant
                        )
                        Spacer(modifier = Modifier.height(8.dp))
                        recentFiles.forEach { uriString ->
                            val rawName = uriString.substringAfterLast('/').substringAfterLast(':')
                                .ifEmpty { uriString }
                            val name = try { URLDecoder.decode(rawName, "UTF-8") } catch (_: Exception) { rawName }
                            val pct = loadPositionFraction(prefs, uriString)
                            val prefix = if (pct >= 0f) "[${(pct * 100).toInt()}%] " else ""
                            Row(
                                verticalAlignment = Alignment.CenterVertically,
                                modifier = Modifier
                                    .fillMaxWidth()
                                    .clickable { openRecentFile(uriString) }
                                    .padding(horizontal = 32.dp, vertical = 6.dp)
                            ) {
                                Icon(
                                    Icons.Default.Description,
                                    null,
                                    modifier = Modifier.size(20.dp),
                                    tint = MaterialTheme.colorScheme.onSurfaceVariant
                                )
                                Spacer(Modifier.width(8.dp))
                                Text(prefix + name, modifier = Modifier.weight(1f))
                            }
                        }
                        Spacer(modifier = Modifier.weight(1f))
                    } else {
                        Spacer(modifier = Modifier.weight(1f))
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

    BackHandler(
        enabled = items.isNotEmpty() || showFolderBrowser
    ) {
        when {
            showFolderBrowser -> {
                showFolderBrowser = false
                settingsVisible = true
            }
            items.isNotEmpty() && !settingsVisible -> settingsVisible = true
            else -> closeFile()
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

    if (showFolderBrowser) {
        folderBrowserUri?.let { folderUri ->
            FolderBrowserDialog(
                initialUri = folderUri,
                prefs = prefs,
                onOpenFile = { uri, name -> openFromFolder(uri, name) },
                onDismiss = { showFolderBrowser = false }
            )
        }
    }

    if (showClearConfirm) {
        AlertDialog(
            onDismissRequest = { showClearConfirm = false },
            title = { Text(Strings.get("clear_recent")) },
            text = { Text(Strings.get("clear_recent_confirm")) },
            confirmButton = {
                TextButton(onClick = {
                    clearRecentFiles(prefs)
                    recentFiles = emptyList()
                    showClearConfirm = false
                }) {
                    Text(Strings.get("yes"))
                }
            },
            dismissButton = {
                TextButton(onClick = { showClearConfirm = false }) {
                    Text(Strings.get("no"))
                }
            }
        )
    }

    if (showAboutDialog) {
        AboutDialog(onDismiss = { showAboutDialog = false })
    }

    if (showJumpDialog) {
        if (displayMode == 1 && pageInfo != null) {
            JumpDialog(
                currentProgress = pageInfo.first,
                isPageMode = true,
                maxPage = pageInfo.second,
                onJump = { value ->
                    showJumpDialog = false
                    scope.launch {
                        val layoutInfo = lazyListState.layoutInfo
                        val visibleItems = layoutInfo.visibleItemsInfo
                        if (visibleItems.isNotEmpty()) {
                            val itemsPerPage = (visibleItems.last().index - visibleItems.first().index + 1).coerceAtLeast(1)
                            lazyListState.animateScrollToItem(((value - 1) * itemsPerPage).coerceIn(0, items.size - 1))
                        }
                    }
                },
                onDismiss = { showJumpDialog = false }
            )
        } else {
            JumpDialog(
                currentProgress = (progress * 100).toInt(),
                isPageMode = false,
                maxPage = 100,
                onJump = { pct ->
                    showJumpDialog = false
                    scope.launch {
                        val fraction = (pct.toFloat() / 100f).coerceIn(0f, 1f)
                        val targetIdx = (fraction * items.size).toInt().coerceIn(0, items.size - 1)
                        lazyListState.animateScrollToItem(targetIdx)
                    }
                },
                onDismiss = { showJumpDialog = false }
            )
        }
    }
}

@Composable
private fun RecentFilesDialog(
    prefs: android.content.SharedPreferences,
    onSelect: (String) -> Unit,
    onDismiss: () -> Unit
) {
    var recent by remember { mutableStateOf(getRecentFiles(prefs)) }
    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text(Strings.get("recent_files")) },
        text = {
            if (recent.isEmpty()) {
                Text(Strings.get("no_recent_files"))
            } else {
                Column {
                    recent.forEach { uriString ->
                        val rawName = uriString.substringAfterLast('/').substringAfterLast(':')
                            .ifEmpty { uriString }
                        val name = try { URLDecoder.decode(rawName, "UTF-8") } catch (_: Exception) { rawName }
                        val pct = loadPositionFraction(prefs, uriString)
                        val prefix = if (pct >= 0f) "[${(pct * 100).toInt()}%] " else ""
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
                            Text(prefix + name, modifier = Modifier.weight(1f))
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
        },
        dismissButton = {
            if (recent.isNotEmpty()) {
                TextButton(onClick = {
                    clearRecentFiles(prefs)
                    recent = emptyList()
                }) {
                    Text(Strings.get("clear_recent"))
                }
            }
        }
    )
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
private fun FolderBrowserDialog(
    initialUri: Uri,
    prefs: android.content.SharedPreferences,
    onOpenFile: (Uri, String) -> Unit,
    onDismiss: () -> Unit
) {
    val context = LocalContext.current
    var rootUri by remember { mutableStateOf(initialUri) }
    var currentDocId by remember { mutableStateOf<String?>(null) }
    var stack by remember { mutableStateOf(listOf<String?>()) }
    var filter by remember { mutableStateOf(0) }
    var entries by remember { mutableStateOf<List<FolderEntry>>(emptyList()) }
    var folderName by remember { mutableStateOf(folderDisplayName(initialUri)) }
    var favorites by remember { mutableStateOf(loadFavoriteFolders(prefs)) }

    BackHandler(enabled = true) {
        if (stack.isNotEmpty()) {
            currentDocId = stack.last()
            stack = stack.dropLast(1)
        } else {
            onDismiss()
        }
    }

    val pickOther = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.OpenDocumentTree()
    ) { uri ->
        uri?.let {
            try {
                context.contentResolver.takePersistableUriPermission(
                    it,
                    Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_GRANT_WRITE_URI_PERMISSION
                )
            } catch (_: Exception) {}
            saveLastFolder(prefs, it.toString())
            rootUri = it
            currentDocId = null
            stack = emptyList()
            folderName = folderDisplayName(it)
        }
    }

    val rootDocId = try { DocumentsContract.getTreeDocumentId(rootUri) } catch (_: Exception) { null }
    fun refresh() {
        entries = listFolderChildren(context, rootUri, currentDocId, filter, rootDocId)
    }

    LaunchedEffect(rootUri, currentDocId, filter) {
        refresh()
    }

    val isFavorite = favorites.any { it.rootUri == rootUri.toString() && it.docId == currentDocId }

    Dialog(
        onDismissRequest = onDismiss,
        properties = DialogProperties(usePlatformDefaultWidth = false)
    ) {
        Surface(
            modifier = Modifier.fillMaxSize(),
            color = MaterialTheme.colorScheme.background
        ) {
            Column(
                modifier = Modifier
                    .fillMaxSize()
                    .safeDrawingPadding()
            ) {
                Row(
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(horizontal = 4.dp, vertical = 4.dp),
                    verticalAlignment = Alignment.CenterVertically
                ) {
                    IconButton(onClick = {
                        if (stack.isNotEmpty()) {
                            currentDocId = stack.last()
                            stack = stack.dropLast(1)
                        } else {
                            onDismiss()
                        }
                    }) {
                        Icon(Icons.Default.ArrowBack, contentDescription = Strings.get("back"))
                    }
                    Text(
                        folderName,
                        modifier = Modifier.weight(1f),
                        maxLines = 1,
                        overflow = TextOverflow.Ellipsis,
                        fontWeight = FontWeight.Bold
                    )
                    IconButton(onClick = {
                        val list = favorites.toMutableList()
                        list.removeAll { it.rootUri == rootUri.toString() && it.docId == currentDocId }
                        if (!isFavorite) {
                            list.add(0, FavoriteFolder(rootUri.toString(), currentDocId, folderName))
                        }
                        favorites = list
                        saveFavoriteFolders(prefs, list)
                    }) {
                        Icon(
                            if (isFavorite) Icons.Default.Star else Icons.Default.StarBorder,
                            contentDescription = Strings.get("favorite_toggle"),
                            tint = if (isFavorite) Color(0xFFFFC107) else Color.Unspecified
                        )
                    }
                    IconButton(onClick = { pickOther.launch(null) }) {
                        Icon(Icons.Default.CreateNewFolder, contentDescription = Strings.get("browse_other"))
                    }
                }
                if (favorites.isNotEmpty()) {
                    Row(
                        modifier = Modifier
                            .fillMaxWidth()
                            .horizontalScroll(rememberScrollState())
                            .padding(horizontal = 8.dp),
                        verticalAlignment = Alignment.CenterVertically
                    ) {
                        favorites.forEach { fav ->
                            AssistChip(
                                onClick = {
                                    rootUri = Uri.parse(fav.rootUri)
                                    currentDocId = fav.docId
                                    stack = emptyList()
                                    folderName = if (fav.name.isNotEmpty()) fav.name else folderDisplayName(Uri.parse(fav.rootUri))
                                },
                                label = { Text(fav.name, maxLines = 1) },
                                leadingIcon = { Icon(Icons.Default.Star, null, modifier = Modifier.size(16.dp)) }
                            )
                            Spacer(Modifier.width(6.dp))
                        }
                    }
                    Spacer(Modifier.height(4.dp))
                }
                Row(
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(horizontal = 8.dp, vertical = 4.dp)
                ) {
                    listOf(0, 1, 2, 3).forEachIndexed { idx, f ->
                        FilterChip(
                            selected = filter == f,
                            onClick = { filter = f },
                            label = {
                                Text(
                                    when (idx) {
                                        0 -> Strings.get("filter_all")
                                        1 -> Strings.get("filter_srt")
                                        2 -> Strings.get("filter_sub")
                                        else -> Strings.get("filter_ass")
                                    },
                                    fontSize = 12.sp
                                )
                            }
                        )
                        Spacer(Modifier.width(6.dp))
                    }
                }
                if (entries.isEmpty()) {
                    Box(
                        modifier = Modifier
                            .fillMaxWidth()
                            .weight(1f),
                        contentAlignment = Alignment.Center
                    ) {
                        Text(
                            Strings.get("no_files"),
                            color = MaterialTheme.colorScheme.onSurfaceVariant
                        )
                    }
                } else {
                    LazyColumn(modifier = Modifier.fillMaxWidth().weight(1f)) {
                        itemsIndexed(entries) { _, entry ->
                            Row(
                                modifier = Modifier
                                    .fillMaxWidth()
                                    .clickable {
                                        if (entry.isDir) {
                                            stack = stack + currentDocId
                                            currentDocId = entry.docId
                                            folderName = entry.name
                                        } else {
                                            val fileUri = DocumentsContract.buildDocumentUriUsingTree(rootUri, entry.docId)
                                            try {
                                                context.contentResolver.takePersistableUriPermission(
                                                    fileUri,
                                                    Intent.FLAG_GRANT_READ_URI_PERMISSION
                                                )
                                            } catch (_: Exception) {}
                                            onOpenFile(fileUri, entry.name)
                                        }
                                    }
                                    .padding(horizontal = 16.dp, vertical = 12.dp),
                                verticalAlignment = Alignment.CenterVertically
                            ) {
                                Icon(
                                    if (entry.isDir) Icons.Default.Folder else Icons.Default.Description,
                                    null,
                                    modifier = Modifier.size(22.dp),
                                    tint = MaterialTheme.colorScheme.onSurfaceVariant
                                )
                                Spacer(Modifier.width(12.dp))
                                Text(
                                    entry.name,
                                    maxLines = 1,
                                    overflow = TextOverflow.Ellipsis
                                )
                            }
                            Divider(color = MaterialTheme.colorScheme.outlineVariant.copy(alpha = 0.4f))
                        }
                    }
                }
            }
        }
    }
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

private fun resolveDisplayName(context: Context, uri: Uri): String {
    val fallback = uri.lastPathSegment?.substringAfterLast('/')?.let {
        try { URLDecoder.decode(it, "UTF-8") } catch (_: Exception) { it }
    } ?: "Unknown"
    return try {
        context.contentResolver
            .query(uri, arrayOf(OpenableColumns.DISPLAY_NAME), null, null, null)
            ?.use { c -> if (c.moveToFirst()) c.getString(0) else null }
            ?: fallback
    } catch (_: Exception) { fallback }
}

private fun loadSubtitle(context: Context, uri: Uri, name: String): Pair<List<SubtitleItem>, String>? {
    return try {
        val takeFlags = Intent.FLAG_GRANT_READ_URI_PERMISSION
        try {
            context.contentResolver.takePersistableUriPermission(uri, takeFlags)
        } catch (_: Exception) {}
        val inputStream = context.contentResolver.openInputStream(uri) ?: return null
        val reader = BufferedReader(InputStreamReader(inputStream))
        val content = reader.readText()
        reader.close()
        val items = SubtitleParser.parseFile(name, content)
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
    val systemFonts = remember {
        val fontsDir = File("/system/fonts/")
        if (!fontsDir.exists()) return@remember emptyList()
        val result = mutableListOf<Pair<String, FontFamily>>()
        val seen = mutableSetOf<String>()
        val files = fontsDir.listFiles { f ->
            val ext = f.extension.lowercase()
            f.isFile && (ext == "ttf" || ext == "otf")
        }?.sortedBy { it.name } ?: emptyList()
        for (file in files) {
            try {
                if (Build.VERSION.SDK_INT < Build.VERSION_CODES.O) continue
                val baseName = file.nameWithoutExtension
                    .replaceFirst(Regex("-[BBIiRr].*$"), "")
                    .replace("_", " ")
                    .trim()
                if (baseName.isBlank() || baseName in seen) continue
                seen.add(baseName)
                val displayName = baseName.split(" ").joinToString(" ") { w ->
                    w.replaceFirstChar { if (it.isLowerCase()) it.uppercase() else it.toString() }
                }
                val family = FontFamily(androidx.compose.ui.text.font.Font(file))
                result.add(displayName to family)
            } catch (_: Exception) { }
            if (result.size >= 40) break
        }
        result
    }
    val allFonts = remember(systemFonts) {
        listOf("Default" to FontFamily.Default) + systemFonts
    }
    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text(Strings.get("select_font")) },
        text = {
            Column(Modifier.verticalScroll(rememberScrollState())) {
                allFonts.forEach { (name, family) ->
                    Row(
                        verticalAlignment = Alignment.CenterVertically,
                        modifier = Modifier
                            .fillMaxWidth()
                            .clickable { onSelect(family) }
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
    var customRed by remember { mutableFloatStateOf(current.red * 255f) }
    var customGreen by remember { mutableFloatStateOf(current.green * 255f) }
    var customBlue by remember { mutableFloatStateOf(current.blue * 255f) }
    val customColor = Color(customRed / 255f, customGreen / 255f, customBlue / 255f)
    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text(title) },
        text = {
            Column(Modifier.verticalScroll(rememberScrollState())) {
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
                                    .clickable { onSelect(color) }
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
                Divider(modifier = Modifier.padding(vertical = 8.dp))
                Text(Strings.get("custom_color"), style = MaterialTheme.typography.bodySmall)
                Spacer(modifier = Modifier.height(4.dp))
                Row(verticalAlignment = Alignment.CenterVertically) {
                    Text(Strings.get("red"), modifier = Modifier.width(20.dp))
                    Slider(
                        value = customRed,
                        onValueChange = { customRed = it },
                        valueRange = 0f..255f,
                        modifier = Modifier.weight(1f)
                    )
                    Text("${customRed.toInt()}", modifier = Modifier.width(28.dp), textAlign = TextAlign.End)
                }
                Row(verticalAlignment = Alignment.CenterVertically) {
                    Text(Strings.get("green"), modifier = Modifier.width(20.dp))
                    Slider(
                        value = customGreen,
                        onValueChange = { customGreen = it },
                        valueRange = 0f..255f,
                        modifier = Modifier.weight(1f)
                    )
                    Text("${customGreen.toInt()}", modifier = Modifier.width(28.dp), textAlign = TextAlign.End)
                }
                Row(verticalAlignment = Alignment.CenterVertically) {
                    Text(Strings.get("blue"), modifier = Modifier.width(20.dp))
                    Slider(
                        value = customBlue,
                        onValueChange = { customBlue = it },
                        valueRange = 0f..255f,
                        modifier = Modifier.weight(1f)
                    )
                    Text("${customBlue.toInt()}", modifier = Modifier.width(28.dp), textAlign = TextAlign.End)
                }
                Spacer(modifier = Modifier.height(4.dp))
                Box(
                    modifier = Modifier
                        .fillMaxWidth()
                        .height(32.dp)
                        .background(customColor)
                        .clickable { onSelect(customColor) }
                )
                Spacer(modifier = Modifier.height(8.dp))
                Button(
                    onClick = { onSelect(customColor) },
                    modifier = Modifier.fillMaxWidth()
                ) {
                    Text(Strings.get("confirm"))
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

private fun savePosition(prefs: android.content.SharedPreferences, uri: Uri?, state: androidx.compose.foundation.lazy.LazyListState, totalItems: Int = 0) {
    if (uri == null) return
    val uriStr = uri.toString()
    val key = "pos_" + abs(uriStr.hashCode())
    val fKey = "posf_" + abs(uriStr.hashCode())
    val fraction = if (totalItems > 0) state.firstVisibleItemIndex.toFloat() / totalItems.toFloat() else 0f
    prefs.edit()
        .putInt(key, state.firstVisibleItemIndex)
        .putFloat(fKey, fraction.coerceIn(0f, 1f))
        .apply()
}

private fun loadPosition(prefs: android.content.SharedPreferences, uriString: String): Int {
    val key = "pos_" + abs(uriString.hashCode())
    return prefs.getInt(key, -1)
}

private fun loadPositionFraction(prefs: android.content.SharedPreferences, uriString: String): Float {
    val key = "posf_" + abs(uriString.hashCode())
    return prefs.getFloat(key, -1f)
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

private fun clearRecentFiles(prefs: android.content.SharedPreferences) {
    val editor = prefs.edit()
    editor.remove("recent_files")
    val all = prefs.all
    for (key in all.keys) {
        if (key.startsWith("pos_") || key.startsWith("posf_")) {
            editor.remove(key)
        }
    }
    editor.apply()
}

private data class FolderEntry(
    val name: String,
    val docId: String,
    val isDir: Boolean
)

private data class FavoriteFolder(
    val rootUri: String,
    val docId: String?,
    val name: String
)

private fun loadLastFolder(prefs: android.content.SharedPreferences): String? {
    return prefs.getString("last_folder", null)
}

private fun saveLastFolder(prefs: android.content.SharedPreferences, uri: String) {
    prefs.edit().putString("last_folder", uri).apply()
}

private fun folderDisplayName(uri: Uri): String {
    val seg = uri.lastPathSegment ?: return uri.toString()
    val raw = seg.substringAfter(":").ifEmpty { seg }
    return try { URLDecoder.decode(raw, "UTF-8") } catch (_: Exception) { raw }
}

private fun listFolderChildren(
    context: Context,
    rootUri: Uri,
    currentDocId: String?,
    filter: Int,
    rootDocId: String?
): List<FolderEntry> {
    val result = mutableListOf<FolderEntry>()
    try {
        val parentId = currentDocId ?: rootDocId ?: return emptyList()
        val childrenUri = DocumentsContract.buildChildDocumentsUriUsingTree(rootUri, parentId)
        val projection = arrayOf(
            DocumentsContract.Document.COLUMN_DOCUMENT_ID,
            DocumentsContract.Document.COLUMN_DISPLAY_NAME,
            DocumentsContract.Document.COLUMN_MIME_TYPE
        )
        context.contentResolver.query(childrenUri, projection, null, null, null)?.use { cursor ->
            while (cursor.moveToNext()) {
                val id = cursor.getString(0)
                val name = cursor.getString(1) ?: ""
                val mime = cursor.getString(2) ?: ""
                val isDir = mime == DocumentsContract.Document.MIME_TYPE_DIR
                if (isDir) {
                    result.add(FolderEntry(name, id, true))
                } else {
                    if (matchesFilter(name, filter)) {
                        result.add(FolderEntry(name, id, false))
                    }
                }
            }
        }
    } catch (_: Exception) {}
    return result.sortedWith(compareBy({ !it.isDir }, { it.name.lowercase() }))
}

private fun matchesFilter(name: String, filter: Int): Boolean {
    if (filter == 0) return true
    val ext = name.substringAfterLast('.', "").lowercase()
    return when (filter) {
        1 -> ext == "srt"
        2 -> ext == "sub"
        else -> ext == "ass" || ext == "ssa"
    }
}

private fun loadFavoriteFolders(prefs: android.content.SharedPreferences): List<FavoriteFolder> {
    val json = prefs.getString("favorite_folders", "[]") ?: "[]"
    return try {
        val arr = org.json.JSONArray(json)
        (0 until arr.length()).mapNotNull { i ->
            val obj = arr.optJSONObject(i) ?: return@mapNotNull null
            val root = obj.optString("root", "")
            if (root.isEmpty()) null
            else FavoriteFolder(root, obj.optString("doc").ifEmpty { null }, obj.optString("name", ""))
        }
    } catch (_: Exception) { emptyList() }
}

private fun saveFavoriteFolders(prefs: android.content.SharedPreferences, list: List<FavoriteFolder>) {
    val arr = org.json.JSONArray()
    list.forEach { fav ->
        val obj = org.json.JSONObject()
        obj.put("root", fav.rootUri)
        fav.docId?.let { obj.put("doc", it) }
        obj.put("name", fav.name)
        arr.put(obj)
    }
    prefs.edit().putString("favorite_folders", arr.toString()).apply()
}

@Composable
private fun AboutDialog(onDismiss: () -> Unit) {
    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text(Strings.get("menu_about"), modifier = Modifier.fillMaxWidth(), textAlign = androidx.compose.ui.text.style.TextAlign.Center) },
        text = {
            Column(modifier = Modifier.fillMaxWidth(), horizontalAlignment = Alignment.CenterHorizontally) {
                Strings.get("about_info").split("\n").forEach { line ->
                    Text(
                        text = line,
                        modifier = Modifier.fillMaxWidth(),
                        textAlign = androidx.compose.ui.text.style.TextAlign.Center,
                        style = MaterialTheme.typography.bodyMedium
                    )
                    Spacer(modifier = Modifier.height(4.dp))
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
private fun JumpDialog(
    currentProgress: Int,
    isPageMode: Boolean,
    maxPage: Int,
    onJump: (Int) -> Unit,
    onDismiss: () -> Unit
) {
    var input by remember { mutableStateOf(currentProgress.toString()) }
    val hintKey = if (isPageMode) "page_jump_hint" else "jump_to_hint"
    val hintText = if (isPageMode) Strings.get("page_jump_hint").replace("{}", maxPage.toString())
                   else Strings.get("jump_to_hint")
    val range = if (isPageMode) 1..maxPage else 0..100
    val label = if (isPageMode) "1-$maxPage" else "0-100"
    AlertDialog(
        onDismissRequest = onDismiss,
        title = { Text(Strings.get("jump_to")) },
        text = {
            Column {
                Text(hintText, style = MaterialTheme.typography.bodySmall)
                Spacer(modifier = Modifier.height(8.dp))
                OutlinedTextField(
                    value = input,
                    onValueChange = { input = it.filter { c -> c.isDigit() } },
                    label = { Text(label) },
                    singleLine = true,
                    modifier = Modifier.fillMaxWidth()
                )
            }
        },
        confirmButton = {
            TextButton(onClick = {
                val pct = input.toIntOrNull()
                if (pct != null && pct in range) {
                    onJump(pct)
                }
            }) {
                Text(Strings.get("yes"))
            }
        },
        dismissButton = {
            TextButton(onClick = onDismiss) {
                Text(Strings.get("cancel"))
            }
        }
    )
}
