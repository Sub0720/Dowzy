# dowzy.py
"""
Dowzy - Video & Audio Downloader (PyQt6)

Latest changes:
- Robust double-click behavior:
  - Double-click Title (col 1) or Progress (col 4):
      * if file exists -> reveal the file in OS file manager
      * else -> open the URL in default browser
- In-app splash overlay (droplet -> splash -> spread) that hides the app contents until it finishes.
- Splash renders "Dowzy" with richer colors.
- Keeps all original features (yt-dlp, trimming, ffmpeg fallback, per-row cancel, queue, preview, etc).
"""

import sys
import os
import threading
import traceback
import subprocess
import json
import re
import math
import random
import webbrowser
from dataclasses import dataclass
from typing import Optional

import requests
from PyQt6 import QtGui, QtCore, QtWidgets
from yt_dlp import YoutubeDL, utils as ytdl_utils

# ------------------------------
# Config / utils
# ------------------------------
APP_NAME = "Dowzy"
DEFAULT_WINDOWS_DOWNLOAD = r"C:\Downloads"
DEFAULT_DOWNLOAD_FOLDER = DEFAULT_WINDOWS_DOWNLOAD if sys.platform.startswith("win") else os.path.join(os.path.expanduser("~"), "Downloads")

YDL_OPTS_TEMPLATE = {
    'outtmpl': '%(title)s.%(ext)s',
    'quiet': True,
    'no_warnings': True,
    'noplaylist': True,
    'format': 'bestvideo+bestaudio/best',
    'merge_output_format': 'mp4',
}

def _which(prog):
    from shutil import which
    return which(prog)

def _get_subprocess_creation_args():
    kwargs = {}
    if os.name == 'nt':
        kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW
    else:
        kwargs['start_new_session'] = True
        kwargs['close_fds'] = True
    return kwargs

def _reveal_file(path: str):
    """Open file location and select the file (platform-specific)."""
    try:
        if os.name == 'nt':
            # explorer requires comma form: explorer /select,"C:\path\to\file"
            subprocess.Popen(['explorer', '/select,{}'.format(os.path.normpath(path))], **_get_subprocess_creation_args())
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', '-R', path], **_get_subprocess_creation_args())
        else:
            folder = os.path.dirname(path)
            subprocess.Popen(['xdg-open', folder], **_get_subprocess_creation_args())
    except Exception:
        try:
            QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(os.path.dirname(path)))
        except Exception:
            pass

# ------------------------------
# Data classes
# ------------------------------
@dataclass
class QueueItem:
    url: str
    format_tag: str
    dest_folder: str
    filename: Optional[str] = None
    thumbnail_url: Optional[str] = None
    filesize: Optional[int] = None
    start_time: Optional[str] = None  # HH:MM:SS or None
    end_time: Optional[str] = None    # HH:MM:SS or None
    skipped: bool = False
    table_row: Optional[int] = None
    title: Optional[str] = None

# ------------------------------
# GlossyButton (unchanged)
# ------------------------------
class GlossyButton(QtWidgets.QPushButton):
    def __init__(self, text: str, color_a: str = '#7c3aed', color_b: str = '#06b6d4', parent=None):
        super().__init__(text, parent)
        self.color_a = QtGui.QColor(color_a)
        self.color_b = QtGui.QColor(color_b)
        self.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.setSizePolicy(QtWidgets.QSizePolicy.Policy.Minimum, QtWidgets.QSizePolicy.Policy.Fixed)
        self._sheen_pos = -1.0
        self._anim_timer = QtCore.QTimer(self)
        self._anim_timer.setInterval(14)
        self._anim_timer.timeout.connect(self._advance_sheen)
        self.setMouseTracking(True)
        self.setFlat(True)
        f = self.font()
        f.setBold(True)
        self.setFont(f)

    def enterEvent(self, event):
        self._sheen_pos = -0.4
        self._anim_timer.start()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._anim_timer.stop()
        self._sheen_pos = -1.0
        self.update()
        super().leaveEvent(event)

    def _advance_sheen(self):
        self._sheen_pos += 0.035
        if self._sheen_pos > 1.6:
            self._sheen_pos = -1.0
            self._anim_timer.stop()
        self.update()

    def paintEvent(self, event):
        w = self.width(); h = self.height()
        radius = min(12, h // 4)
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
        grad = QtGui.QLinearGradient(0, 0, w, 0)
        grad.setColorAt(0.0, self.color_a)
        grad.setColorAt(1.0, self.color_b)
        p.setBrush(QtGui.QBrush(grad))
        p.setPen(QtCore.Qt.PenStyle.NoPen)
        rect = QtCore.QRectF(0, 0, w, h)
        p.drawRoundedRect(rect, radius, radius)
        overlay = QtGui.QLinearGradient(0, 0, 0, h)
        overlay.setColorAt(0.0, QtGui.QColor(255,255,255,45))
        overlay.setColorAt(0.5, QtGui.QColor(255,255,255,18))
        overlay.setColorAt(1.0, QtGui.QColor(255,255,255,5))
        p.setBrush(QtGui.QBrush(overlay))
        p.drawRoundedRect(rect.adjusted(1,1,-1,-1), max(6, radius-2), max(6, radius-2))
        if self._sheen_pos >= -1.0:
            sheen_w = w * 0.25
            x_center = int(self._sheen_pos * w)
            sheen_rect = QtCore.QRectF(x_center - sheen_w/2, -h, sheen_w, h*3)
            sheen_grad = QtGui.QLinearGradient(sheen_rect.topLeft(), sheen_rect.topRight())
            sheen_grad.setColorAt(0.0, QtGui.QColor(255,255,255,0))
            sheen_grad.setColorAt(0.45, QtGui.QColor(255,255,255,120))
            sheen_grad.setColorAt(0.55, QtGui.QColor(255,255,255,90))
            sheen_grad.setColorAt(1.0, QtGui.QColor(255,255,255,0))
            p.save()
            p.translate(w/2, h/2)
            p.rotate(-20)
            p.translate(-w/2, -h/2)
            p.setBrush(QtGui.QBrush(sheen_grad))
            p.setPen(QtCore.Qt.PenStyle.NoPen)
            p.drawRoundedRect(sheen_rect, radius, radius)
            p.restore()
        if self.underMouse():
            glow_color = QtGui.QColor(255,255,255,30)
            pen = QtGui.QPen(glow_color)
            pen.setWidth(2)
            p.setPen(pen)
            p.setBrush(QtCore.Qt.BrushStyle.NoBrush)
            p.drawRoundedRect(rect.adjusted(1,1,-1,-1), radius, radius)
        p.setPen(QtGui.QPen(QtGui.QColor(255,255,255)))
        fm = QtGui.QFontMetrics(self.font())
        text = self.text()
        elided = fm.elidedText(text, QtCore.Qt.TextElideMode.ElideRight, w-16)
        p.drawText(rect, QtCore.Qt.AlignmentFlag.AlignCenter, elided)
        p.end()



# ------------------------------
# Custom TitleBar (frameless window integration)
# ------------------------------
class TitleBar(QtWidgets.QWidget):
    def __init__(self, parent=None, app_name=APP_NAME):
        super().__init__(parent)
        self._drag_pos = None
        self._is_maximized = False
        self.parent_window = parent.window() if parent is not None else None

        self.setFixedHeight(44)
        self.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Fixed)
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_StyledBackground, True)

        # layout
        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(6, 0, 6, 0)
        layout.setSpacing(8)

        # app icon (use a small colored square as placeholder)
        self.icon_label = QtWidgets.QLabel(self)
        self.icon_label.setFixedSize(28, 28)
        self.icon_label.setStyleSheet('QLabel { background: qlineargradient(x1:0 y1:0, x2:1 y2:1, stop:0 #7c3aed, stop:1 #06b6d4); border-radius:4px; }')
        layout.addWidget(self.icon_label)

        # title text
        self.title = QtWidgets.QLabel(f"{app_name}")
        f = self.title.font(); f.setPointSize(10); f.setBold(True)
        self.title.setFont(f)
        self.title.setAlignment(QtCore.Qt.AlignmentFlag.AlignVCenter | QtCore.Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(self.title)

        # subtitle / clickable buy-me link (kept small)
        self.subtitle = QtWidgets.QLabel('✨ <a href=\"https://buymeacoffee.com/skaax007\">Buy me a coffee ☕</a>')
        self.subtitle.setTextFormat(QtCore.Qt.TextFormat.RichText)
        self.subtitle.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.TextBrowserInteraction)
        self.subtitle.setOpenExternalLinks(True)
        sub_font = self.subtitle.font(); sub_font.setPointSize(9)
        self.subtitle.setFont(sub_font)
        layout.addWidget(self.subtitle)

        layout.addStretch()

        btn_style = 'QPushButton { border: none; color: #e6eef8; padding: 6px; min-width:26px; } QPushButton:hover { background: rgba(255,255,255,0.03); }'

        # minimize / maximize / close
        self.min_btn = QtWidgets.QPushButton('\u2013')  # en dash as minimize
        self.min_btn.setFixedSize(36, 28)
        self.min_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.min_btn.setStyleSheet(btn_style)
        self.min_btn.clicked.connect(self._on_minimize)
        layout.addWidget(self.min_btn)

        self.max_btn = QtWidgets.QPushButton('\u25A1')  # square for maximize
        self.max_btn.setFixedSize(36, 28)
        self.max_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.max_btn.setStyleSheet(btn_style)
        self.max_btn.clicked.connect(self._on_maximize_restore)
        layout.addWidget(self.max_btn)

        self.close_btn = QtWidgets.QPushButton('\u2715')  # multiplication X
        self.close_btn.setFixedSize(36, 28)
        self.close_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.close_btn.setStyleSheet('QPushButton { border: none; color: #ffd6d6; padding: 6px; } QPushButton:hover { background: rgba(255,80,80,0.12); }')
        self.close_btn.clicked.connect(self._on_close)
        layout.addWidget(self.close_btn)

        # subtle shadow for titlebar
        try:
            effect = QtWidgets.QGraphicsDropShadowEffect(self)
            effect.setBlurRadius(12)
            effect.setOffset(0, 1)
            effect.setColor(QtGui.QColor(0,0,0,160))
            self.setGraphicsEffect(effect)
        except Exception:
            pass

    def mousePressEvent(self, ev):
        if ev.button() == QtCore.Qt.MouseButton.LeftButton:
            self._drag_pos = ev.globalPosition().toPoint() - self.window().frameGeometry().topLeft()
        super().mousePressEvent(ev)

    def mouseMoveEvent(self, ev):
        if ev.buttons() & QtCore.Qt.MouseButton.LeftButton and self._drag_pos is not None:
            # Move the window by the delta between current global pos and saved drag offset
            try:
                new_pos = ev.globalPosition().toPoint() - self._drag_pos
                self.window().move(new_pos)
            except Exception:
                pass
        super().mouseMoveEvent(ev)

    def mouseDoubleClickEvent(self, ev):
        # double-click to maximize / restore
        if ev.button() == QtCore.Qt.MouseButton.LeftButton:
            self._on_maximize_restore()
        super().mouseDoubleClickEvent(ev)

    def _on_minimize(self):
        try:
            self.window().showMinimized()
        except Exception:
            pass

    def _on_maximize_restore(self):
        try:
            w = self.window()
            if w.isMaximized():
                w.showNormal()
                self._is_maximized = False
            else:
                w.showMaximized()
                self._is_maximized = True
        except Exception:
            pass

    def _on_close(self):
        try:
            self.window().close()
        except Exception:
            pass

# ------------------------------
# Splash overlay (in-app, hides app until it finishes)
# ------------------------------
class SplashOverlay(QtWidgets.QWidget):
    finished = QtCore.pyqtSignal()

    def __init__(self, main_window):
        super().__init__(main_window)
        self.main_window = main_window
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_NoSystemBackground, True)
        self.setWindowFlags(QtCore.Qt.WindowType.FramelessWindowHint)
        self.setGeometry(self.main_window.rect())

        # animation state
        self.phase = 'falling'   # 'falling' -> 'impact' -> 'spread' -> 'done'
        self.drop_x = self.width() // 2
        self.drop_y = -80
        self.drop_radius = 12
        self.drop_speed = 14.0
        self.impact_y = int(self.height() * 0.45)

        self.ripples = []
        self.particles = []
        self.spread_radius = 0
        self.spread_speed = max(self.width(), self.height()) / 18.0

        self._timer = QtCore.QTimer(self)
        self._timer.setInterval(26)
        self._timer.timeout.connect(self._tick)
        self._running = False

        # for nicer title gradient
        self.title_gradient_phase = 0.0

    def start(self):
        # ensure overlay covers main window
        self.setGeometry(self.main_window.rect())
        self.drop_x = int(self.width() * 0.52)
        self.drop_y = -40
        self.drop_radius = 12
        self.drop_speed = 14.0
        self.phase = 'falling'
        self.ripples.clear()
        self.particles.clear()
        self.spread_radius = 0
        self.title_gradient_phase = 0.0
        self._running = True
        self.show()
        self.raise_()
        self._timer.start()

    def resizeEvent(self, ev):
        try:
            self.setGeometry(self.main_window.rect())
            self.impact_y = int(self.height() * 0.45)
        except Exception:
            pass
        super().resizeEvent(ev)

    def _make_splash(self):
        for i in range(3):
            self.ripples.append({'r': 8 + i*6, 'alpha': 220 - i*40, 'speed': 2 + i*1.2})
        for _ in range(22):
            ang = random.uniform(0, math.pi*2)
            speed = random.uniform(4.0, 10.0)
            vx = math.cos(ang) * speed
            vy = math.sin(ang) * speed * -0.6
            life = random.uniform(14, 32)
            px = self.drop_x + random.uniform(-12, 12)
            py = self.impact_y + random.uniform(-8, 8)
            self.particles.append({'x': px, 'y': py, 'vx': vx, 'vy': vy, 'life': life})

    def _tick(self):
        if self.phase == 'falling':
            self.drop_y += self.drop_speed
            self.drop_speed = min(self.drop_speed + 0.7, 44.0)
            if self.drop_y >= self.impact_y:
                self.phase = 'impact'
                self._make_splash()
        elif self.phase == 'impact':
            for r in self.ripples:
                r['r'] += r['speed'] * 3.2
                r['alpha'] = max(0, r['alpha'] - r['speed'] * 2.4)
            new_particles = []
            for p in self.particles:
                p['x'] += p['vx']
                p['y'] += p['vy']
                p['vy'] += 0.6
                p['life'] -= 1
                if p['life'] > 0:
                    new_particles.append(p)
            self.particles = new_particles
            if all(r['alpha'] <= 0 for r in self.ripples):
                self.phase = 'spread'
        elif self.phase == 'spread':
            self.spread_radius += self.spread_speed
            if self.spread_radius > math.hypot(self.width(), self.height()) * 1.12:
                self.phase = 'done'
                self._timer.stop()
                QtCore.QTimer.singleShot(140, self._finish)
        self.title_gradient_phase += 0.04
        self.update()

    def _finish(self):
        self._running = False
        try:
            self.hide()
        except Exception:
            pass
        self.finished.emit()

    def paintEvent(self, ev):
        w = self.width(); h = self.height()
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)

        # lush multicolor background gradient
        grad = QtGui.QLinearGradient(0, 0, w, h)
        grad.setColorAt(0.0, QtGui.QColor(8, 12, 28))
        grad.setColorAt(0.35, QtGui.QColor(10, 28, 60))
        grad.setColorAt(0.75, QtGui.QColor(18, 48, 98))
        grad.setColorAt(1.0, QtGui.QColor(6, 14, 32))
        p.fillRect(0, 0, w, h, grad)

        # ambient soft glow circle behind splash
        glow = QtGui.QRadialGradient(self.drop_x, self.impact_y, max(w,h)/2)
        glow.setColorAt(0.0, QtGui.QColor(40, 140, 255, 60))
        glow.setColorAt(0.8, QtGui.QColor(20, 80, 180, 12))
        glow.setColorAt(1.0, QtGui.QColor(0,0,0,0))
        p.setBrush(QtGui.QBrush(glow))
        p.setPen(QtCore.Qt.PenStyle.NoPen)
        p.drawRect(0,0,w,h)

        # ripples
        for r in self.ripples:
            alpha = int(max(0, min(255, r['alpha'])))
            color = QtGui.QColor(120, 200, 255, alpha)
            pen = QtGui.QPen(color)
            pen.setWidth(2)
            p.setPen(pen)
            p.setBrush(QtCore.Qt.BrushStyle.NoBrush)
            rect = QtCore.QRectF(self.drop_x - r['r'], self.impact_y - r['r'], r['r']*2, r['r']*2)
            p.drawEllipse(rect)

        # splash particles
        for part in self.particles:
            age = max(0, int(part['life']))
            alpha = max(10, min(255, int(240 * (age / 32.0))))
            color = QtGui.QColor(180, 230, 255, alpha)
            p.setPen(QtCore.Qt.PenStyle.NoPen)
            p.setBrush(QtGui.QBrush(color))
            rx = int(part['x']); ry = int(part['y'])
            p.drawEllipse(rx-3, ry-3, 6, 6)

        # falling drop (teardrop)
        if self.phase == 'falling':
            color = QtGui.QColor(150, 220, 255, 240)
            p.setPen(QtCore.Qt.PenStyle.NoPen)
            p.setBrush(QtGui.QBrush(color))
            dx = int(self.drop_x); dy = int(self.drop_y)
            r = int(self.drop_radius)
            p.drawEllipse(dx - r, dy - r, 2*r, 2*r)
            tail = QtGui.QPainterPath()
            tail.moveTo(dx, dy + r)
            tail.cubicTo(dx + r*0.9, dy + r*2.4, dx - r*0.9, dy + r*2.4, dx, dy + r)
            p.drawPath(tail)

        # spread overlay
        if self.phase in ('spread','done'):
            sr = self.spread_radius
            if sr < 1:
                sr = 1
            # multiple rings for visual interest
            for i in range(3):
                alpha = 120 - i*30
                color = QtGui.QColor(40 + i*20, 100 + i*40, 190 + i*20, alpha)
                p.setPen(QtCore.Qt.PenStyle.NoPen)
                p.setBrush(QtGui.QBrush(color))
                rect = QtCore.QRectF(self.drop_x - sr*(0.9 + i*0.08), self.impact_y - sr*(0.7 + i*0.08), sr*(1.8 + i*0.16), sr*(1.4 + i*0.16))
                p.drawEllipse(rect)

        # title "Dowzy" centered with animated gradient
        title = APP_NAME
        font = QtGui.QFont("Segoe UI", 36, QtGui.QFont.Weight.Bold)
        p.setFont(font)
        fm = p.fontMetrics()
        tw = fm.horizontalAdvance(title)
        tx = int((w - tw) / 2)
        ty = int(h * 0.22 + fm.ascent())

        # gradient for title text
        grad_text = QtGui.QLinearGradient(tx, ty-fm.ascent(), tx+tw, ty+fm.ascent())
        phase = (math.sin(self.title_gradient_phase) + 1.0) * 0.5
        # mix colors based on phase
        grad_text.setColorAt(0.0, QtGui.QColor(210, 245, 255))
        grad_text.setColorAt(max(0.0, phase-0.1), QtGui.QColor(150, 210, 255))
        grad_text.setColorAt(min(1.0, phase+0.1), QtGui.QColor(240, 200, 255))
        grad_text.setColorAt(1.0, QtGui.QColor(255, 220, 180))
        painter_path = QtGui.QPainterPath()
        painter_path.addText(tx, ty, font, title)
        p.setPen(QtCore.Qt.PenStyle.NoPen)
        p.setBrush(QtGui.QBrush(grad_text))
        p.drawPath(painter_path)

# ------------------------------
# Initialization Dialog (unchanged)
# ------------------------------
class InitializationDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Initialization")
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowType.WindowContextHelpButtonHint)
        self.setModal(False)
        self.setFixedSize(360, 96)
        layout = QtWidgets.QVBoxLayout(self)
        self.label = QtWidgets.QLabel("Initialization")
        self.label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label)
        self.progress = QtWidgets.QProgressBar()
        self.progress.setRange(0,0)  # indeterminate
        layout.addWidget(self.progress)
        self._dots = 0
        self._timer = QtCore.QTimer(self)
        self._timer.setInterval(450)
        self._timer.timeout.connect(self._tick)
        self._timer.start()

    def _tick(self):
        self._dots = (self._dots + 1) % 5
        self.label.setText("Initialization" + "." * self._dots)

    def closeEvent(self, ev):
        self._timer.stop()
        super().closeEvent(ev)

    def set_progress_percent(self, p: int):
        if self.progress.maximum() == 0:
            self.progress.setRange(0,100)
        self.progress.setValue(max(0,min(100,int(p))))

# ------------------------------
# Download worker (unchanged logic)
# ------------------------------
class DownloadWorker(QtCore.QObject):
    progress_changed = QtCore.pyqtSignal(int)  # percent
    status_changed = QtCore.pyqtSignal(str)
    finished = QtCore.pyqtSignal(object)  # item or exception
    thumbnail_ready = QtCore.pyqtSignal(QtGui.QPixmap)
    size_known = QtCore.pyqtSignal(str)

    def __init__(self, item: QueueItem):
        super().__init__()
        self.item = item
        self._abort = False

    def abort(self):
        self._abort = True

    def _human_readable(self, size_bytes: int) -> str:
        try:
            for unit in ['B','KB','MB','GB','TB']:
                if size_bytes < 1024.0:
                    return f"{size_bytes:.2f} {unit}"
                size_bytes /= 1024.0
        except Exception:
            pass
        return '—'

    def _normalize_time(self, t: str) -> str:
        if not t:
            return ''
        t = t.strip()
        parts = t.split(':')
        try:
            parts = [int(p) for p in parts]
        except Exception:
            return t
        if len(parts) == 3:
            h,m,s = parts
            return f"{h:02d}:{m:02d}:{s:02d}"
        if len(parts) == 2:
            m,s = parts
            return f"00:{m:02d}:{s:02d}"
        if len(parts) == 1:
            secs = parts[0]
            h = secs // 3600; m = (secs%3600)//60; s = secs%60
            return f"{h:02d}:{m:02d}:{s:02d}"
        return t

    def _time_to_seconds(self, t: str) -> Optional[int]:
        if not t:
            return None
        try:
            parts = [int(p) for p in t.split(':')]
        except Exception:
            return None
        if len(parts) == 3:
            return parts[0]*3600+parts[1]*60+parts[2]
        if len(parts) == 2:
            return parts[0]*60+parts[1]
        return parts[0]

    def run(self):
        try:
            # prefer CLI trimmed-before-download if requested & binary available
            use_cli = False
            if self.item.start_time and self.item.end_time and _which('yt-dlp'):
                use_cli = True

            if use_cli:
                # cli trimmed-download
                start_n = self._normalize_time(self.item.start_time)
                end_n = self._normalize_time(self.item.end_time)
                section_spec = f"*{start_n}-{end_n}"
                outtmpl = os.path.join(self.item.dest_folder, '%(title)s.%(ext)s')
                cmd = [
                    'yt-dlp',
                    '-f', self.item.format_tag,
                    '--merge-output-format', 'mp4',
                    '--no-warnings',
                    '--no-call-home',
                    '--newline',
                    '--download-sections', section_spec,
                    '-o', outtmpl,
                    self.item.url
                ]
                self.status_changed.emit('Starting (trimmed) download')
                kwargs = _get_subprocess_creation_args()
                proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1, **kwargs)
                pct_last = 0
                try:
                    for raw_line in proc.stdout:
                        if self._abort:
                            try:
                                proc.terminate()
                            except Exception:
                                pass
                            raise ytdl_utils.DownloadError('Aborted by user')
                        line = raw_line.strip()
                        m = re.search(r'([0-9]{1,3}(?:\.[0-9]+)?)\s*%', line)
                        if m:
                            try:
                                pct = int(float(m.group(1)))
                                if pct != pct_last:
                                    pct_last = pct
                                    self.progress_changed.emit(pct)
                            except Exception:
                                pass
                        if line:
                            self.status_changed.emit(line[:110])
                    ret = proc.wait()
                    if ret != 0:
                        raise Exception(f'yt-dlp exited with code {ret}')
                    if not self.item.filename and self.item.title:
                        for fn in os.listdir(self.item.dest_folder):
                            if fn.startswith(self.item.title):
                                self.item.filename = os.path.join(self.item.dest_folder, fn)
                                break
                    self.status_changed.emit('Completed')
                    self.finished.emit(self.item)
                    return
                finally:
                    try:
                        proc.stdout.close()
                    except Exception:
                        pass

            # fallback: Python YoutubeDL then post-trim
            opts = dict(YDL_OPTS_TEMPLATE)
            opts['outtmpl'] = os.path.join(self.item.dest_folder, '%(title)s.%(ext)s')
            opts['format'] = self.item.format_tag
            opts['progress_hooks'] = [self._progress_hook]
            opts['merge_output_format'] = 'mp4'
            with YoutubeDL(opts) as ydl:
                info = ydl.extract_info(self.item.url, download=False)
                thumb = info.get('thumbnail')
                if thumb:
                    self.item.thumbnail_url = thumb
                    try:
                        r = requests.get(thumb, timeout=8)
                        if r.status_code == 200:
                            pix = QtGui.QPixmap()
                            pix.loadFromData(r.content)
                            self.thumbnail_ready.emit(pix)
                    except Exception:
                        pass
                formats = info.get('formats')
                fs = None
                if formats:
                    for f in formats[::-1]:
                        if f.get('format_id') and f.get('format_id') in self.item.format_tag:
                            fs = f.get('filesize') or f.get('filesize_approx')
                            if fs:
                                break
                if not fs:
                    fs = info.get('filesize') or info.get('filesize_approx')
                if fs:
                    self.size_known.emit(self._human_readable(fs))
                try:
                    prepared = ydl.prepare_filename(info)
                    mo = opts.get('merge_output_format')
                    if mo:
                        base, _ = os.path.splitext(prepared)
                        final_expected = base + '.' + mo
                    else:
                        final_expected = prepared
                    self.item.filename = final_expected
                except Exception:
                    self.item.filename = None
                self.status_changed.emit('Starting download')
                ydl.download([self.item.url])

            # If trimming requested, run ffmpeg trimming (post-download)
            if self.item.start_time and self.item.end_time and self.item.filename:
                start_n = self._normalize_time(self.item.start_time)
                end_n = self._normalize_time(self.item.end_time)
                s_s = self._time_to_seconds(start_n); e_s = self._time_to_seconds(end_n)
                if s_s is None or e_s is None or e_s <= s_s:
                    self.status_changed.emit('Trim skipped (invalid times)')
                else:
                    self.status_changed.emit('Trimming selection')
                    base, ext = os.path.splitext(self.item.filename)
                    trimmed_tmp = base + '_trim_tmp' + ext

                    # try stream copy first (fast)
                    cmd_copy = [
                        'ffmpeg', '-y', '-i', self.item.filename,
                        '-ss', start_n, '-to', end_n,
                        '-c', 'copy', trimmed_tmp
                    ]
                    kwargs = _get_subprocess_creation_args()
                    try:
                        subprocess.run(cmd_copy, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True, **kwargs)
                        try:
                            os.replace(trimmed_tmp, self.item.filename)
                        except Exception:
                            try:
                                os.remove(self.item.filename)
                            except Exception:
                                pass
                            os.rename(trimmed_tmp, self.item.filename)
                        self.status_changed.emit('Trimmed')
                    except subprocess.CalledProcessError:
                        duration = str(e_s - s_s)
                        cmd_encode = [
                            'ffmpeg', '-y',
                            '-ss', start_n,
                            '-i', self.item.filename,
                            '-t', duration,
                            '-c:v', 'libx264', '-c:a', 'aac', '-strict', '-2',
                            trimmed_tmp
                        ]
                        try:
                            subprocess.run(cmd_encode, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True, **kwargs)
                            try:
                                os.replace(trimmed_tmp, self.item.filename)
                            except Exception:
                                try:
                                    os.remove(self.item.filename)
                                except Exception:
                                    pass
                                os.rename(trimmed_tmp, self.item.filename)
                            self.status_changed.emit('Trimmed')
                        except subprocess.CalledProcessError:
                            self.status_changed.emit('Trim failed')
                            try:
                                if os.path.exists(trimmed_tmp):
                                    os.remove(trimmed_tmp)
                            except Exception:
                                pass

            self.status_changed.emit('Completed')
            self.finished.emit(self.item)
        except Exception as e:
            tb = traceback.format_exc()
            self.finished.emit(e)

    def _progress_hook(self, d):
        try:
            if self._abort:
                raise ytdl_utils.DownloadError('Aborted by user')
            status = d.get('status')
            if status == 'downloading':
                total_bytes = d.get('total_bytes') or d.get('total_bytes_estimate')
                downloaded = d.get('downloaded_bytes') or d.get('downloaded_bytes')
                if total_bytes and downloaded:
                    percent = int(downloaded * 100 / total_bytes)
                    self.progress_changed.emit(percent)
                elif d.get('progress'):
                    try:
                        percent = int(d.get('progress') * 100)
                        self.progress_changed.emit(percent)
                    except Exception:
                        pass
                if total_bytes:
                    human = self._human_readable(total_bytes)
                    self.size_known.emit(human)
            elif status == 'finished':
                self.progress_changed.emit(100)
                self.status_changed.emit('Merging / finalizing')
        except Exception:
            raise

# ------------------------------
# Main window
# ------------------------------
class DowzyWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        # use a frameless window so we can merge the title bar with the app UI
        try:
            self.setWindowFlag(QtCore.Qt.WindowType.FramelessWindowHint)
        except Exception:
            pass
        # remove any default window margins so custom titlebar sits flush
        try:
            self.setContentsMargins(0,0,0,0)
        except Exception:
            pass
        self.setWindowTitle(APP_NAME)
        self.resize(960, 680)
        self.download_folder = DEFAULT_DOWNLOAD_FOLDER
        self.queue = []
        self.current_worker_thread = None

        self._init_count = 0
        self._init_lock = threading.Lock()
        self.init_dialog = None

        self._build_ui()

        # create overlay (parented to main window) and hide central area until overlay finishes
        self.splash_overlay = SplashOverlay(self)
        self.splash_overlay.finished.connect(self._on_splash_finished)
        # hide central contents until splash finishes
        self.centralWidget().setVisible(False)

        self._start_clipboard_timer()

    def _build_ui(self):
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        try:
            central.setContentsMargins(0,0,0,0)
        except Exception:
            pass
        layout = QtWidgets.QVBoxLayout(central)
        layout.setContentsMargins(0,0,0,0)

        # custom titlebar (frameless)
        titlebar = TitleBar(self)
        layout.addWidget(titlebar)

        # URL row
        row = QtWidgets.QHBoxLayout()
        self.url_edit = QtWidgets.QLineEdit()
        self.url_edit.setPlaceholderText('Paste video link here (auto-detected from clipboard)')
        row.addWidget(self.url_edit)
        clipboard_btn = GlossyButton('Paste from clipboard 📋'); clipboard_btn.clicked.connect(self._paste_from_clipboard)
        row.addWidget(clipboard_btn)
        layout.addLayout(row)

        # Options: preview, format, folder
        options = QtWidgets.QHBoxLayout()
        thumb_box = QtWidgets.QGroupBox('Preview')
        tb_layout = QtWidgets.QVBoxLayout()
        self.thumb_label = QtWidgets.QLabel(); self.thumb_label.setFixedSize(180,100); self.thumb_label.setStyleSheet('background: #222; border-radius:6px'); self.thumb_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        tb_layout.addWidget(self.thumb_label)
        self.size_label = QtWidgets.QLabel('Size: —'); tb_layout.addWidget(self.size_label)
        thumb_box.setLayout(tb_layout); options.addWidget(thumb_box)

        fmt_box = QtWidgets.QGroupBox('Download type & quality')
        f_layout = QtWidgets.QFormLayout()
        self.type_combo = QtWidgets.QComboBox(); self.type_combo.addItems(['Video + Audio (best) 🎬','Audio only (best) 🎧','Custom format (advanced)']); self.type_combo.currentIndexChanged.connect(self._format_changed)
        f_layout.addRow('Mode:', self.type_combo)
        self.quality_combo = QtWidgets.QComboBox(); self.format_map = {'Video + Audio (best) 🎬': 'bestvideo+bestaudio/best','Audio only (best) 🎧': 'bestaudio/best'}
        self.quality_combo.addItems(['Default (best)','1080p','720p','480p','320k (audio)'])
        f_layout.addRow('Preset:', self.quality_combo)
        self.custom_format = QtWidgets.QLineEdit(); self.custom_format.setPlaceholderText('eg: bestvideo[height<=1080]+bestaudio/best'); self.custom_format.hide(); f_layout.addRow('Format:', self.custom_format)
        fmt_box.setLayout(f_layout); options.addWidget(fmt_box, stretch=1)

        folder_box = QtWidgets.QGroupBox('Save folder')
        fo_layout = QtWidgets.QVBoxLayout()
        self.folder_edit = QtWidgets.QLineEdit(self.download_folder); fo_layout.addWidget(self.folder_edit)
        folder_btns = QtWidgets.QHBoxLayout()
        choose_btn = GlossyButton('Choose... 📁'); choose_btn.clicked.connect(self._choose_folder); folder_btns.addWidget(choose_btn)
        open_btn = GlossyButton('Open folder'); open_btn.clicked.connect(self._open_folder); folder_btns.addWidget(open_btn)
        fo_layout.addLayout(folder_btns); folder_box.setLayout(fo_layout); options.addWidget(folder_box)

        layout.addLayout(options)

        # toolbar / actions (clean)
        toolbar = QtWidgets.QHBoxLayout()
        left_grp = QtWidgets.QHBoxLayout()
        add_btn = GlossyButton('Add to queue ➕'); add_btn.clicked.connect(self._add_to_queue); left_grp.addWidget(add_btn)
        self.start_btn = GlossyButton('Start download ▶️'); self.start_btn.clicked.connect(self._start_queue); left_grp.addWidget(self.start_btn)
        self.cancel_btn = GlossyButton('Cancel download ⛔'); self.cancel_btn.clicked.connect(self._cancel_current); self.cancel_btn.setEnabled(False); left_grp.addWidget(self.cancel_btn)
        toolbar.addLayout(left_grp)

        toolbar.addSpacing(14)

        mid_grp = QtWidgets.QHBoxLayout()
        trim_label = QtWidgets.QLabel('Trim (HH:MM:SS):'); mid_grp.addWidget(trim_label)
        self.trim_start = QtWidgets.QLineEdit(); self.trim_start.setFixedWidth(110); self.trim_start.setInputMask("00:00:00;_"); self.trim_start.setPlaceholderText('HH:MM:SS'); mid_grp.addWidget(self.trim_start)
        mid_grp.addWidget(QtWidgets.QLabel('→'))
        self.trim_end = QtWidgets.QLineEdit(); self.trim_end.setFixedWidth(110); self.trim_end.setInputMask("00:00:00;_"); self.trim_end.setPlaceholderText('HH:MM:SS'); mid_grp.addWidget(self.trim_end)
        self.apply_btn = GlossyButton('Apply trim ✔️'); self.apply_btn.clicked.connect(self._apply_trim); mid_grp.addWidget(self.apply_btn)
        self.trim_applied_label = QtWidgets.QLabel('No trim applied'); mid_grp.addWidget(self.trim_applied_label)
        toolbar.addLayout(mid_grp)

        toolbar.addStretch()

        # (duplicate Choose/Open removed near trim per your request)

        layout.addLayout(toolbar)

        # Trim progress bar
        self.trim_progress = QtWidgets.QProgressBar()
        self.trim_progress.setRange(0,0)
        self.trim_progress.setVisible(False)
        self.trim_progress.setFixedWidth(260)
        trim_row = QtWidgets.QHBoxLayout()
        trim_row.addWidget(self.trim_progress)
        trim_row.addStretch()
        layout.addLayout(trim_row)

        # queue table
        self.table = QtWidgets.QTableWidget(0,6)
        self.table.setHorizontalHeaderLabels(['Status','Title / URL','Type','Size','Progress','✖'])
        self.table.horizontalHeader().setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeMode.Stretch)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        # click/double-click handlers
        self.table.cellClicked.connect(self._on_table_clicked)
        self.table.cellDoubleClicked.connect(self._on_table_double_clicked)
        layout.addWidget(self.table)

        self.status = QtWidgets.QLabel('Ready'); layout.addWidget(self.status)

        self.setStyleSheet(r"""
            QMainWindow { background: qlineargradient(x1:0 y1:0, x2:1 y2:1, stop:0 #0b0c10, stop:0.5 #091022, stop:1 #0b0c10); }
            QLabel { color: #e6eef8; }
            QLineEdit, QComboBox, QTextEdit {
                background: rgba(15, 23, 32, 0.55);
                color: #dbeeff;
                border: 1px solid rgba(255, 255, 255, 0.08);
                padding: 6px;
                border-radius: 8px;
            }
            QProgressBar {
                background: rgba(8, 16, 24, 0.6);
                border-radius: 6px;
                height: 18px;
                color: #dbeeff;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #7c3aed, stop:1 #06b6d4);
                border-radius: 6px;
            }
        """)

        self._set_start_button_state()

    # ----------------
    # clipboard + preview
    # ----------------
    def _paste_from_clipboard(self):
        cb = QtWidgets.QApplication.clipboard(); txt = cb.text().strip()
        if txt:
            self.url_edit.setText(txt); self._fetch_preview(txt)

    def _start_clipboard_timer(self):
        self._last_clip = ''
        self.clip_timer = QtCore.QTimer(self); self.clip_timer.timeout.connect(self._check_clipboard); self.clip_timer.start(1200)

    def _check_clipboard(self):
        cb = QtWidgets.QApplication.clipboard(); txt = cb.text().strip()
        if txt and txt != self._last_clip:
            self._last_clip = txt
            if txt.startswith('http'):
                self.url_edit.setText(txt); self._fetch_preview(txt)

    def _fetch_preview(self, url):
        def job():
            try:
                ydl_opts = dict(YDL_OPTS_TEMPLATE); ydl_opts['quiet'] = True
                with YoutubeDL(ydl_opts) as ydl:
                    info = ydl.extract_info(url, download=False)
                    title = info.get('title') or url; thumb = info.get('thumbnail'); filesize = info.get('filesize') or info.get('filesize_approx')
                    QtCore.QMetaObject.invokeMethod(self, '_update_preview_ui', QtCore.Qt.ConnectionType.QueuedConnection,
                                                    QtCore.Q_ARG(str, title or ''), QtCore.Q_ARG(str, thumb or ''), QtCore.Q_ARG(object, filesize))
            except Exception:
                pass
        threading.Thread(target=job, daemon=True).start()

    @QtCore.pyqtSlot(str, str, object)
    def _update_preview_ui(self, title, thumb, filesize):
        self.status.setText(f'Ready — {title}')
        if thumb:
            try:
                r = requests.get(thumb, timeout=6)
                if r.status_code == 200:
                    pix = QtGui.QPixmap(); pix.loadFromData(r.content)
                    pix = pix.scaled(self.thumb_label.size(), QtCore.Qt.AspectRatioMode.KeepAspectRatio, QtCore.Qt.TransformationMode.SmoothTransformation)
                    self.thumb_label.setPixmap(pix)
            except Exception:
                pass
        if filesize:
            self.size_label.setText('Size: ' + self._human_readable(filesize))
        else:
            self.size_label.setText('Size: —')

    def _format_changed(self, idx):
        if self.type_combo.currentText().startswith('Custom'):
            self.custom_format.show()
        else:
            self.custom_format.hide()

    def _choose_folder(self):
        path = QtWidgets.QFileDialog.getExistingDirectory(self, 'Choose folder', self.download_folder)
        if path:
            self.download_folder = path; self.folder_edit.setText(path)

    def _open_folder(self):
        path = self.folder_edit.text().strip() or self.download_folder
        if os.path.exists(path):
            QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(path))

    # ----------------
    # trim apply
    # ----------------
    def _apply_trim(self):
        s = self.trim_start.text(); e = self.trim_end.text()
        def mask_blank(x):
            return (not x) or ('_' in x) or x.strip() in (':','::','','  :  :  ')
        if mask_blank(s) and mask_blank(e):
            self.trim_applied_label.setText('No trim applied'); self.status.setText('Trim cleared (none applied)'); return
        if mask_blank(s) or mask_blank(e):
            self.status.setText('Both start and end required to apply trim'); return
        try:
            ps = [int(p) for p in s.split(':')]; pe = [int(p) for p in e.split(':')]
            def secs(p):
                if len(p)==3: return p[0]*3600+p[1]*60+p[2]
                if len(p)==2: return p[0]*60+p[1]
                return p[0]
            if secs(pe) <= secs(ps):
                self.status.setText('End time must be greater than start time'); return
        except Exception:
            self.status.setText('Invalid time format — use HH:MM:SS (apply helps)'); return
        self.trim_applied_label.setText(f'Applied: {s} → {e}'); self.status.setText('Trim times applied (will be used on Add to queue)')

    # ----------------
    # add to queue + initialization
    # ----------------
    def _add_to_queue(self):
        url = self.url_edit.text().strip()
        if not url:
            self.status.setText('Paste a link first'); return
        mode = self.type_combo.currentText()
        if mode.startswith('Custom'):
            fmt = self.custom_format.text().strip() or 'bestvideo+bestaudio/best'
        else:
            preset = self.quality_combo.currentText()
            if preset == 'Default (best)': fmt = self.format_map.get(self.type_combo.currentText(), 'bestvideo+bestaudio/best')
            elif preset == '1080p': fmt = 'bestvideo[height<=1080]+bestaudio/best'
            elif preset == '720p': fmt = 'bestvideo[height<=720]+bestaudio/best'
            elif preset == '480p': fmt = 'bestvideo[height<=480]+bestaudio/best'
            elif preset == '320k (audio)': fmt = 'bestaudio[abr<=320]/bestaudio/best'
            else: fmt = self.format_map.get(self.type_combo.currentText(), 'bestvideo+bestaudio/best')
        dest = self.folder_edit.text().strip() or self.download_folder

        s_raw = self.trim_start.text(); e_raw = self.trim_end.text()
        def mask_blank(x):
            return (not x) or ('_' in x) or x.strip() in (':','::','','  :  :  ')
        start_t = None if mask_blank(s_raw) else s_raw
        end_t = None if mask_blank(e_raw) else e_raw
        if (start_t and not end_t) or (end_t and not start_t):
            self.status.setText('Both start and end times required for trimming, or leave both blank'); return

        item = QueueItem(url=url, format_tag=fmt, dest_folder=dest, start_time=start_t, end_time=end_t)
        self.queue.append(item); self._insert_table_row(item)
        self.status.setText('Added to queue — initializing metadata...')
        self._start_initialization(item)
        self._set_start_button_state()

    def _insert_table_row(self, item: QueueItem):
        row = self.table.rowCount(); self.table.insertRow(row)
        status_item = QtWidgets.QTableWidgetItem('Queued')
        status_item.setFlags(QtCore.Qt.ItemFlag.ItemIsSelectable | QtCore.Qt.ItemFlag.ItemIsEnabled)
        url_item = QtWidgets.QTableWidgetItem(item.url)
        url_item.setFlags(QtCore.Qt.ItemFlag.ItemIsSelectable | QtCore.Qt.ItemFlag.ItemIsEnabled)
        # store QueueItem object in user role (robust)
        url_item.setData(QtCore.Qt.ItemDataRole.UserRole, item)
        type_item = QtWidgets.QTableWidgetItem(item.format_tag)
        type_item.setFlags(QtCore.Qt.ItemFlag.ItemIsSelectable | QtCore.Qt.ItemFlag.ItemIsEnabled)
        size_item = QtWidgets.QTableWidgetItem('—'); size_item.setFlags(QtCore.Qt.ItemFlag.ItemIsSelectable | QtCore.Qt.ItemFlag.ItemIsEnabled)
        prog_bar = QtWidgets.QProgressBar(); prog_bar.setValue(0)
        prog_bar.setTextVisible(False)
        self.table.setItem(row,0,status_item); self.table.setItem(row,1,url_item); self.table.setItem(row,2,type_item); self.table.setItem(row,3,size_item); self.table.setCellWidget(row,4,prog_bar)
        cross_btn = QtWidgets.QPushButton('✖'); cross_btn.setFixedWidth(34); cross_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        cross_btn.setToolTip('Cancel / remove this item')
        cross_btn.setStyleSheet('QPushButton { color: #ffb4b4; background: rgba(255,255,255,0.03); border-radius:6px }')

        item.table_row = row
        cross_btn._queue_item = item

        def on_cross_clicked():
            it = getattr(cross_btn, '_queue_item', None)
            if it is None:
                return
            try:
                current_item = None
                try:
                    if self.current_worker_thread and hasattr(self.current_worker_thread, 'worker'):
                        current_item = getattr(self.current_worker_thread.worker, 'item', None)
                except Exception:
                    current_item = None

                if current_item is it:
                    try:
                        self.current_worker_thread.worker.abort()
                    except Exception:
                        pass
                    try:
                        row_found = self._find_row_for_item(it)
                        if row_found is not None:
                            self.table.item(row_found,0).setText('Cancelling (user)')
                            pb = self.table.cellWidget(row_found,4)
                            if pb: pb.setValue(0)
                    except Exception:
                        pass
                    self.status.setText('Cancelling current download...')
                else:
                    removed_from_queue = False
                    try:
                        for q in list(self.queue):
                            if q is it:
                                self.queue.remove(q)
                                removed_from_queue = True
                                break
                    except Exception:
                        pass
                    try:
                        row_found = self._find_row_for_item(it)
                        if row_found is not None:
                            self.table.item(row_found,0).setText('Removed')
                            pb = self.table.cellWidget(row_found,4)
                            if pb: pb.setValue(0)
                            cross_btn.setEnabled(False)
                            f = self.table.item(row_found,1).font(); f.setStrikeOut(True); self.table.item(row_found,1).setFont(f)
                    except Exception:
                        pass
                    self.status.setText('Item removed from queue' if removed_from_queue else 'Item removed')
                self._set_start_button_state()
            except Exception:
                pass

        cross_btn.clicked.connect(on_cross_clicked)
        self.table.setCellWidget(row,5,cross_btn)

    def _find_row_for_item(self, item: QueueItem):
        for r in range(self.table.rowCount()):
            try:
                cell = self.table.item(r,1)
                if cell is None:
                    continue
                stored = cell.data(QtCore.Qt.ItemDataRole.UserRole)
                if stored is item:
                    return r
            except Exception:
                pass
        return None

    # -------------------
    # initialization logic (shows dialog while inits running)
    # -------------------
    def _start_initialization(self, item: QueueItem):
        with self._init_lock:
            self._init_count += 1
        QtCore.QTimer.singleShot(0, self._maybe_show_init_dialog)

        def job():
            title = ''; thumb = ''; filesize = None; expected_filename = ''
            try:
                if item.start_time and item.end_time and _which('yt-dlp'):
                    start = item.start_time; end = item.end_time
                    section_spec = f"*{start}-{end}"
                    cmd = ['yt-dlp','-J','--no-warnings','--no-call-home','--no-download','--download-sections', section_spec, '-f', item.format_tag, item.url]
                    kwargs = _get_subprocess_creation_args()
                    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, **kwargs)
                    out = proc.stdout.strip()
                    if out:
                        try:
                            j = json.loads(out)
                            title = j.get('title') or ''
                            thumb = j.get('thumbnail') or ''
                            fs = j.get('filesize') or j.get('filesize_approx')
                            if not fs:
                                fmts = j.get('formats') or []
                                for f in fmts[::-1]:
                                    if f.get('format_id') and f.get('format_id') in item.format_tag:
                                        fs = f.get('filesize') or f.get('filesize_approx')
                                        if fs:
                                            break
                            if fs:
                                filesize = int(fs)
                            mo = YDL_OPTS_TEMPLATE.get('merge_output_format','mp4')
                            if title:
                                expected_filename = os.path.join(item.dest_folder, title + '.' + mo)
                        except Exception:
                            pass
                else:
                    ydl_opts = dict(YDL_OPTS_TEMPLATE); ydl_opts['quiet'] = True; ydl_opts['format'] = item.format_tag
                    with YoutubeDL(ydl_opts) as ydl:
                        info = ydl.extract_info(item.url, download=False)
                        title = info.get('title') or ''
                        thumb = info.get('thumbnail') or ''
                        fs = None
                        formats = info.get('formats')
                        if formats:
                            for f in formats[::-1]:
                                if f.get('format_id') and f.get('format_id') in item.format_tag:
                                    fs = f.get('filesize') or f.get('filesize_approx')
                                    if fs:
                                        break
                        if not fs:
                            fs = info.get('filesize') or info.get('filesize_approx')
                        if fs:
                            filesize = int(fs)
                        try:
                            prepared = ydl.prepare_filename(info)
                            mo = ydl_opts.get('merge_output_format')
                            if mo:
                                base,_ = os.path.splitext(prepared)
                                expected_filename = base + '.' + mo
                            else:
                                expected_filename = prepared
                        except Exception:
                            expected_filename = ''
            except Exception:
                pass

            QtCore.QMetaObject.invokeMethod(self, '_finish_initialization_ui', QtCore.Qt.ConnectionType.QueuedConnection,
                                            QtCore.Q_ARG(object, item),
                                            QtCore.Q_ARG(str, title or ''),
                                            QtCore.Q_ARG(str, thumb or ''),
                                            QtCore.Q_ARG(object, filesize),
                                            QtCore.Q_ARG(str, expected_filename or ''))
        threading.Thread(target=job, daemon=True).start()

    @QtCore.pyqtSlot(object, str, str, object, str)
    def _finish_initialization_ui(self, item: QueueItem, title: str, thumb: str, filesize: object, expected_filename: str):
        try:
            if title:
                item.title = title
                row = self._find_row_for_item(item)
                if row is not None:
                    try:
                        self.table.item(row,1).setText(title)
                    except Exception:
                        pass
            if filesize:
                try:
                    item.filesize = int(filesize)
                    row = self._find_row_for_item(item)
                    if row is not None:
                        self.table.item(row,3).setText(self._human_readable(item.filesize))
                except Exception:
                    pass
            if thumb:
                item.thumbnail_url = thumb
                try:
                    r = requests.get(thumb, timeout=6)
                    if r.status_code == 200:
                        pix = QtGui.QPixmap(); pix.loadFromData(r.content)
                        pix = pix.scaled(self.thumb_label.size(), QtCore.Qt.AspectRatioMode.KeepAspectRatio, QtCore.Qt.TransformationMode.SmoothTransformation)
                        self.thumb_label.setPixmap(pix)
                except Exception:
                    pass
            if expected_filename:
                item.filename = expected_filename
        finally:
            with self._init_lock:
                self._init_count = max(0, self._init_count - 1)
            self._maybe_close_init_dialog()
            self._set_start_button_state()
            self.status.setText('Ready' if self._init_count == 0 else 'Initializing...')

    def _maybe_show_init_dialog(self):
        with self._init_lock:
            if self._init_count > 0 and self.init_dialog is None:
                self.init_dialog = InitializationDialog(self)
                self.init_dialog.show()

    def _maybe_close_init_dialog(self):
        with self._init_lock:
            if self._init_count == 0 and self.init_dialog:
                try:
                    self.init_dialog.close()
                except Exception:
                    pass
                self.init_dialog = None

    def _set_start_button_state(self):
        with self._init_lock:
            inits = self._init_count
        queued_non_skipped = 0
        for it in self.queue:
            if not it.skipped:
                queued_non_skipped += 1
        if inits > 0:
            self.start_btn.setEnabled(False)
        else:
            if self.current_worker_thread is not None:
                self.start_btn.setEnabled(False)
            else:
                self.start_btn.setEnabled(True if queued_non_skipped > 0 else False)

    # ----------------
    # start / cancel / process
    # ----------------
    def _start_queue(self):
        if not self.queue:
            self.status.setText('Queue is empty — add something'); return
        if self.current_worker_thread is not None:
            self.status.setText('Already downloading'); return
        with self._init_lock:
            if self._init_count > 0:
                self.status.setText('Waiting for initialization to finish...'); return
        if not any(not it.skipped for it in self.queue):
            self.status.setText('No non-skipped items to download'); self._set_start_button_state(); return
        self._process_next()

    def _cancel_current(self):
        if self.current_worker_thread and hasattr(self.current_worker_thread, 'worker'):
            try:
                self.current_worker_thread.worker.abort(); self.status.setText('Cancelling...'); self.cancel_btn.setEnabled(False)
            except Exception:
                pass

    def _process_next(self):
        while self.queue and self.queue[0].skipped:
            skipped_item = self.queue.pop(0)
            try:
                row = self._find_row_for_item(skipped_item)
                if row is not None:
                    self.table.item(row,0).setText('Skipped')
                    pb = self.table.cellWidget(row,4)
                    if pb: pb.setValue(0)
            except Exception: pass
        if not self.queue:
            self.status.setText('All done'); self._set_start_button_state(); return
        item = self.queue.pop(0)
        r = None; prog = None
        row_found = self._find_row_for_item(item)
        if row_found is not None:
            try:
                if self.table.item(row_found,0).text() in ('Queued','Starting'):
                    self.table.item(row_found,0).setText('Starting'); prog = self.table.cellWidget(row_found,4)
                    if prog: prog.setValue(0)
                    r = row_found
            except Exception:
                pass

        # create thread/worker
        thread = QtCore.QThread()
        worker = DownloadWorker(item)
        worker.moveToThread(thread)
        thread.worker = worker
        worker.progress_changed.connect(lambda p, row_ref=r: self._on_progress(row_ref,p))
        worker.status_changed.connect(lambda s, row_ref=r: self._on_status(row_ref,s))
        worker.thumbnail_ready.connect(lambda pix: self._set_thumbnail_pixmap(pix))
        worker.size_known.connect(lambda hs, row_ref=r: self._set_size(row_ref,hs))
        worker.finished.connect(lambda result, row_ref=r: self._on_finished(row_ref,result,thread))
        thread.started.connect(worker.run)
        thread.start()
        self.current_worker_thread = thread
        self.cancel_btn.setEnabled(True)
        self._set_start_button_state()
        self.status.setText('Downloading...')

    def _on_progress(self, row, percent):
        try:
            if row is None: return
            pb = self.table.cellWidget(row,4)
            if pb: pb.setValue(percent)
        except Exception: pass

    def _on_status(self, row, text):
        try:
            if row is None: return
            short = text if len(text) <= 110 else text[:110] + '...'
            self.table.item(row,0).setText(short)
            if 'Trimming selection' in text or text.startswith('Trimming'):
                self.trim_progress.setRange(0,0)
                self.trim_progress.setVisible(True)
            elif 'Trimmed' in text or 'Trim failed' in text or 'Trim skipped' in text:
                try:
                    self.trim_progress.setRange(0,100)
                    self.trim_progress.setValue(100)
                except Exception:
                    pass
                QtCore.QTimer.singleShot(900, lambda: self.trim_progress.setVisible(False))
            self.status.setText(text if len(text) < 120 else text[:120] + '...')
        except Exception: pass

    def _set_thumbnail_pixmap(self, pix: QtGui.QPixmap):
        p = pix.scaled(self.thumb_label.size(), QtCore.Qt.AspectRatioMode.KeepAspectRatio, QtCore.Qt.TransformationMode.SmoothTransformation)
        self.thumb_label.setPixmap(p)

    def _set_size(self, row, human_size):
        try:
            if row is None: return
            self.table.item(row,3).setText(human_size)
        except Exception: pass

    def _human_readable(self, size_bytes: int) -> str:
        try:
            for unit in ['B','KB','MB','GB','TB']:
                if size_bytes < 1024.0:
                    return f"{size_bytes:.2f} {unit}"
                size_bytes /= 1024.0
        except Exception:
            pass
        return '—'

    def _locate_downloaded_file(self, item: QueueItem) -> Optional[str]:
        """
        Try to return an actual existing filename for this item:
        - if item.filename exists on disk -> return it
        - else try to find files in dest_folder that match item.title (prefix) or URL-stem
        """
        try:
            if item.filename and os.path.exists(item.filename):
                return item.filename
            # try title-based matching
            candidates = []
            if item.title:
                prefix = re.sub(r'[\\/:"*?<>|]+', '', item.title).strip()
                for fn in os.listdir(item.dest_folder):
                    if fn.lower().startswith(prefix.lower()):
                        candidates.append(fn)
            # fallback: try safe stem from URL
            if not candidates:
                url_stem = None
                try:
                    p = re.sub(r'https?://', '', item.url or '')
                    p = p.split('/')[0:3]
                    url_stem = ''.join(p)
                except Exception:
                    url_stem = None
                if url_stem:
                    for fn in os.listdir(item.dest_folder):
                        if url_stem.lower() in fn.lower():
                            candidates.append(fn)
            # as last resort, any file with same timestamp after download? can't rely on that, so use first candidate
            if candidates:
                # prefer mp4 / mkv / mp3 / m4a etc
                prefer_exts = ['.mp4', '.mkv', '.webm', '.mp3', '.m4a', '.aac', '.flv']
                for ext in prefer_exts:
                    for fn in candidates:
                        if fn.lower().endswith(ext):
                            return os.path.join(item.dest_folder, fn)
                # otherwise return first
                return os.path.join(item.dest_folder, candidates[0])
        except Exception:
            pass
        return None

    def _on_finished(self, row, result, thread: QtCore.QThread):
        try:
            if isinstance(result, Exception):
                msg = f'Error: {str(result)}'
                if row is not None and row < self.table.rowCount(): self.table.item(row,0).setText('Failed')
                self.status.setText(msg)
            else:
                # result is QueueItem object
                # try to ensure item.filename points to an actual existing file
                try:
                    item = result
                    found = self._locate_downloaded_file(item)
                    if found:
                        item.filename = found
                    # update UI row text & progress
                    row_found = self._find_row_for_item(item)
                    if row_found is not None:
                        self.table.item(row_found,0).setText('Completed')
                        try:
                            self.table.cellWidget(row_found,4).setValue(100)
                        except Exception:
                            pass
                        # update size cell if known
                        try:
                            if item.filesize:
                                self.table.item(row_found,3).setText(self._human_readable(item.filesize))
                        except Exception:
                            pass
                except Exception:
                    pass
                self.status.setText('Completed')
        finally:
            try:
                thread.quit(); thread.wait(2000)
            except Exception:
                pass
            self.current_worker_thread = None
            self.cancel_btn.setEnabled(False)
            self._set_start_button_state()
            if self.queue:
                QtCore.QTimer.singleShot(800, self._process_next)

    # -----------------------------
    # click behavior:
    # - single-click Title (col==1): open the URL in browser
    # - double-click Title or Progress (col==1 or col==4):
    #     if file exists -> reveal file
    #     else -> open URL
    # -----------------------------
    def _on_table_clicked(self, row, col):
        try:
            if col != 1:
                return
            cell = self.table.item(row,1)
            if cell is None:
                return
            item = cell.data(QtCore.Qt.ItemDataRole.UserRole)
            if item is None:
                url = cell.text() if cell else None
                if url:
                    webbrowser.open(url)
                return
            # single click should open url (if not desired change, let me know)
            try:
                webbrowser.open(item.url)
            except Exception:
                pass
        except Exception:
            pass

    def _on_table_double_clicked(self, row, col):
        try:
            if col not in (1,4):
                return
            cell = self.table.item(row,1)
            if cell is None:
                return
            item = cell.data(QtCore.Qt.ItemDataRole.UserRole)
            if item is None:
                url = cell.text() if cell else None
                if url:
                    webbrowser.open(url)
                return
            # if file exists, reveal file; else open url
            fname = None
            try:
                # ensure latest file path
                fname = self._locate_downloaded_file(item)
            except Exception:
                fname = item.filename
            if fname and os.path.exists(fname):
                _reveal_file(fname)
            else:
                try:
                    webbrowser.open(item.url)
                except Exception:
                    pass
        except Exception:
            pass

    # ------------------------------
    # splash finished -> reveal central widget
    # ------------------------------
    def _on_splash_finished(self):
        try:
            self.centralWidget().setVisible(True)
            self.status.setText('Ready')
        except Exception:
            pass

# ------------------------------
# Entrypoint
# ------------------------------
def main():
    app = QtWidgets.QApplication(sys.argv)
    win = DowzyWindow()
    win.show()
    # start overlay animation; central widget will be shown when overlay finishes
    try:
        win.splash_overlay.start()
    except Exception:
        # fallback: if overlay fails, make central visible
        win.centralWidget().setVisible(True)
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
