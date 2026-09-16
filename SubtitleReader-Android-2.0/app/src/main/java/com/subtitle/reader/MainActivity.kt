package com.subtitle.reader

import android.content.Intent
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.view.KeyEvent
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Surface
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.setValue
import androidx.compose.ui.Modifier
import com.subtitle.reader.ui.screens.ReaderScreen
import com.subtitle.reader.ui.theme.SubtitleReaderTheme

object ReaderEvents {
    @Volatile var volumeKeyPaging: Boolean = false
    @Volatile var onPagePrev: (() -> Unit)? = null
    @Volatile var onPageNext: (() -> Unit)? = null
}

class MainActivity : ComponentActivity() {

    private var incomingUri by mutableStateOf<String?>(null)

    override fun onCreate(savedInstanceState: Bundle?) {
        enableEdgeToEdge()
        super.onCreate(savedInstanceState)
        incomingUri = resolveIncomingUri(intent)
        setContent {
            SubtitleReaderTheme {
                Surface(
                    modifier = Modifier.fillMaxSize(),
                    color = MaterialTheme.colorScheme.background
                ) {
                    ReaderScreen(initialUri = incomingUri)
                }
            }
        }
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        incomingUri = resolveIncomingUri(intent)
    }

    private fun resolveIncomingUri(intent: Intent): String? {
        var uri: Uri? = when (intent.action) {
            Intent.ACTION_VIEW -> intent.data
            Intent.ACTION_SEND -> if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU) {
                intent.getParcelableExtra(Intent.EXTRA_STREAM, Uri::class.java)
            } else {
                @Suppress("DEPRECATION") intent.getParcelableExtra(Intent.EXTRA_STREAM)
            }
            else -> null
        }
        uri ?: return null
        if (uri.scheme == "content") {
            try {
                contentResolver.takePersistableUriPermission(
                    uri,
                    Intent.FLAG_GRANT_READ_URI_PERMISSION
                )
            } catch (_: Exception) {}
        }
        return uri.toString()
    }

    override fun dispatchKeyEvent(event: KeyEvent): Boolean {
        if (event.action == KeyEvent.ACTION_DOWN && ReaderEvents.volumeKeyPaging) {
            when (event.keyCode) {
                KeyEvent.KEYCODE_VOLUME_UP -> {
                    ReaderEvents.onPagePrev?.invoke()
                    return true
                }
                KeyEvent.KEYCODE_VOLUME_DOWN -> {
                    ReaderEvents.onPageNext?.invoke()
                    return true
                }
                KeyEvent.KEYCODE_MEDIA_PREVIOUS,
                KeyEvent.KEYCODE_MEDIA_REWIND -> {
                    ReaderEvents.onPagePrev?.invoke()
                    return true
                }
                KeyEvent.KEYCODE_MEDIA_NEXT,
                KeyEvent.KEYCODE_MEDIA_FAST_FORWARD -> {
                    ReaderEvents.onPageNext?.invoke()
                    return true
                }
            }
        }
        return super.dispatchKeyEvent(event)
    }
}