#!/usr/bin/env python3
"""JPEG metadata editor — Tags, Subject, Comments via Windows Property System."""

import sys
from pathlib import Path
from typing import Optional

from PyQt6.QtCore import Qt, QRunnable, QThreadPool, pyqtSignal, QObject, QTimer
from PyQt6.QtGui import QPixmap, QAction
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QSplitter, QScrollArea,
    QGridLayout, QVBoxLayout, QFormLayout, QLabel, QPushButton,
    QLineEdit, QTextEdit, QToolBar, QFileDialog, QMessageBox, QFrame,
)

# ── Windows Property System ───────────────────────────────────────────────────
try:
    import pywintypes
    from win32com.propsys import propsys, pscon
    import pythoncom

    _GPS_DEFAULT   = getattr(propsys, "GPS_DEFAULT",   0)
    _GPS_READWRITE = getattr(propsys, "GPS_READWRITE", 2)

    # FMTID_SummaryInformation — canonical PKEYs used by Windows shell
    _FMTID_SI = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}"

    def _fallback_pkey(pid: int):
        return (pywintypes.IID(_FMTID_SI), pid)

    PKEY_Keywords = getattr(pscon, "PKEY_Keywords", _fallback_pkey(5))
    PKEY_Subject  = getattr(pscon, "PKEY_Subject",  _fallback_pkey(3))
    PKEY_Comment  = getattr(pscon, "PKEY_Comment",  _fallback_pkey(6))

    HAS_PROPSYS = True

except ImportError:
    HAS_PROPSYS = False


def _open_store(path: str, flags: int):
    return propsys.SHGetPropertyStoreFromParsingName(
        path, None, flags, propsys.IPropertyStore
    )


def _read_prop(store, key) -> object:
    try:
        return store.GetValue(key).GetValue()
    except Exception:
        return None


def _write_prop(store, key, value) -> None:
    if not value and value != 0:
        # VT_EMPTY clears the property
        pv = propsys.PROPVARIANTType()
    elif isinstance(value, (list, tuple)):
        pv = propsys.PROPVARIANTType(
            tuple(value), pythoncom.VT_VECTOR | pythoncom.VT_BSTR
        )
    else:
        pv = propsys.PROPVARIANTType(str(value))
    store.SetValue(key, pv)


# ── Constants ─────────────────────────────────────────────────────────────────
THUMB_SIZE = 120
GRID_COLS  = 4
JPEG_EXTS  = frozenset({".jpg", ".jpeg", ".jpe", ".jfif"})


# ── Async thumbnail loader ────────────────────────────────────────────────────
class _LoadSignals(QObject):
    done = pyqtSignal(str, QPixmap)


class _ThumbLoader(QRunnable):
    def __init__(self, path: str):
        super().__init__()
        self.path = path
        self.signals = _LoadSignals()
        self.setAutoDelete(True)

    def run(self):
        px = QPixmap(self.path)
        if px.isNull():
            px = QPixmap(THUMB_SIZE, THUMB_SIZE)
            px.fill(Qt.GlobalColor.darkGray)
        else:
            px = px.scaled(
                THUMB_SIZE, THUMB_SIZE,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        self.signals.done.emit(self.path, px)


# ── Thumbnail cell ────────────────────────────────────────────────────────────
class ThumbnailWidget(QFrame):
    selected = pyqtSignal(str)

    _STYLE_NORMAL = (
        "ThumbnailWidget {"
        "  background: transparent;"
        "  border: 2px solid transparent;"
        "  border-radius: 3px;"
        "}"
        "ThumbnailWidget:hover {"
        "  background: #eef2ff;"
        "  border-color: #b0c4de;"
        "}"
    )
    _STYLE_SELECTED = (
        "ThumbnailWidget {"
        "  background: #cce4ff;"
        "  border: 2px solid #0078d4;"
        "  border-radius: 3px;"
        "}"
    )

    def __init__(self, path: str, parent=None):
        super().__init__(parent)
        self.path = path
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setFixedSize(THUMB_SIZE + 24, THUMB_SIZE + 38)

        vbox = QVBoxLayout(self)
        vbox.setContentsMargins(4, 4, 4, 2)
        vbox.setSpacing(3)

        self._img = QLabel(alignment=Qt.AlignmentFlag.AlignCenter)
        self._img.setFixedSize(THUMB_SIZE, THUMB_SIZE)
        placeholder = QPixmap(THUMB_SIZE, THUMB_SIZE)
        placeholder.fill(Qt.GlobalColor.lightGray)
        self._img.setPixmap(placeholder)

        name = Path(path).name
        display = name if len(name) <= 16 else name[:13] + "..."
        self._name = QLabel(display, alignment=Qt.AlignmentFlag.AlignCenter)
        self._name.setFixedWidth(THUMB_SIZE + 16)
        font = self._name.font()
        font.setPointSize(8)
        self._name.setFont(font)

        vbox.addWidget(self._img)
        vbox.addWidget(self._name)
        self.setStyleSheet(self._STYLE_NORMAL)

    def set_pixmap(self, px: QPixmap):
        self._img.setPixmap(px)

    def select(self):
        self.setStyleSheet(self._STYLE_SELECTED)

    def deselect(self):
        self.setStyleSheet(self._STYLE_NORMAL)

    def mousePressEvent(self, event):
        self.selected.emit(self.path)


# ── Left pane: thumbnail grid ─────────────────────────────────────────────────
class ThumbnailGrid(QWidget):
    image_selected = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._widgets: dict[str, ThumbnailWidget] = {}
        self._current: Optional[str] = None
        self._pool = QThreadPool.globalInstance()

        self._layout = QGridLayout(self)
        self._layout.setContentsMargins(8, 8, 8, 8)
        self._layout.setSpacing(6)
        self._layout.setAlignment(
            Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft
        )

    def load_folder(self, folder: str):
        # Remove and destroy all existing cells
        for tw in self._widgets.values():
            self._layout.removeWidget(tw)
            tw.deleteLater()
        self._widgets.clear()
        self._current = None

        files = sorted(
            str(p)
            for p in Path(folder).iterdir()
            if p.is_file() and p.suffix.lower() in JPEG_EXTS
        )

        for i, path in enumerate(files):
            tw = ThumbnailWidget(path)
            tw.selected.connect(self._on_select)
            row, col = divmod(i, GRID_COLS)
            self._layout.addWidget(tw, row, col)
            self._widgets[path] = tw

            loader = _ThumbLoader(path)
            loader.signals.done.connect(self._on_loaded)
            self._pool.start(loader)

    def _on_loaded(self, path: str, px: QPixmap):
        # Guard against cells deleted by a subsequent load_folder call
        if path in self._widgets:
            self._widgets[path].set_pixmap(px)

    def _on_select(self, path: str):
        if self._current and self._current in self._widgets:
            self._widgets[self._current].deselect()
        self._current = path
        self._widgets[path].select()
        self.image_selected.emit(path)


# ── Right pane: metadata editor ───────────────────────────────────────────────
class MetadataPanel(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._path: Optional[str] = None

        vbox = QVBoxLayout(self)
        vbox.setContentsMargins(14, 14, 14, 14)
        vbox.setSpacing(10)

        self._title = QLabel("No image selected")
        self._title.setStyleSheet("font-weight: bold; font-size: 13px;")
        vbox.addWidget(self._title)

        form = QFormLayout()
        form.setVerticalSpacing(8)
        form.setHorizontalSpacing(10)

        self._tags    = QLineEdit(placeholderText="comma-separated")
        self._subject = QLineEdit()
        self._comment = QTextEdit()
        self._comment.setFixedHeight(110)
        self._comment.setAcceptRichText(False)

        form.addRow("Tags:", self._tags)
        form.addRow("Subject:", self._subject)
        form.addRow("Comments:", self._comment)
        vbox.addLayout(form)

        self._save_btn = QPushButton("Save")
        self._save_btn.setFixedWidth(80)
        self._save_btn.setEnabled(False)
        self._save_btn.clicked.connect(self._save)
        vbox.addWidget(self._save_btn)

        vbox.addStretch()
        self._set_enabled(False)

    def _set_enabled(self, on: bool):
        self._tags.setEnabled(on)
        self._subject.setEnabled(on)
        self._comment.setEnabled(on)
        self._save_btn.setEnabled(on)

    def load(self, path: str):
        self._path = path
        self._title.setText(Path(path).name)

        if not HAS_PROPSYS:
            self._tags.clear()
            self._subject.clear()
            self._comment.clear()
            self._set_enabled(True)
            return

        try:
            store = _open_store(path, _GPS_DEFAULT)
        except Exception as exc:
            QMessageBox.warning(self, "Open Error", str(exc))
            return

        kw = _read_prop(store, PKEY_Keywords)
        if isinstance(kw, (list, tuple)):
            self._tags.setText(", ".join(str(k) for k in kw))
        elif kw:
            self._tags.setText(str(kw))
        else:
            self._tags.clear()

        subj = _read_prop(store, PKEY_Subject)
        self._subject.setText(str(subj) if subj else "")

        comm = _read_prop(store, PKEY_Comment)
        self._comment.setPlainText(str(comm) if comm else "")

        self._set_enabled(True)

    def _save(self):
        if not self._path or not HAS_PROPSYS:
            return

        try:
            store = _open_store(self._path, _GPS_READWRITE)
        except Exception as exc:
            QMessageBox.warning(
                self, "Save Error", f"Cannot open file for writing:\n{exc}"
            )
            return

        try:
            tags = [t.strip() for t in self._tags.text().split(",") if t.strip()]
            _write_prop(store, PKEY_Keywords, tuple(tags))

            subj = self._subject.text().strip()
            _write_prop(store, PKEY_Subject, subj)

            comm = self._comment.toPlainText().strip()
            _write_prop(store, PKEY_Comment, comm)

            store.Commit()
        except Exception as exc:
            QMessageBox.warning(self, "Save Error", f"Write failed:\n{exc}")
            return

        self._save_btn.setText("Saved!")
        QTimer.singleShot(1500, lambda: self._save_btn.setText("Save"))


# ── Main window ───────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("JPEG Metadata Editor")
        self.resize(1100, 720)

        toolbar = QToolBar()
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        open_act = QAction("Open Folder...", self)
        open_act.setShortcut("Ctrl+O")
        open_act.triggered.connect(self._open_folder)
        toolbar.addAction(open_act)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.setCentralWidget(splitter)

        self._grid = ThumbnailGrid()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self._grid)
        scroll.setMinimumWidth(320)
        splitter.addWidget(scroll)

        self._meta = MetadataPanel()
        self._meta.setMinimumWidth(280)
        splitter.addWidget(self._meta)

        splitter.setSizes([700, 400])

        self._grid.image_selected.connect(self._meta.load)

        if not HAS_PROPSYS:
            self.statusBar().showMessage(
                "win32com.propsys unavailable — metadata read/write disabled"
            )

    def _open_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            self._grid.load_folder(folder)
            self.setWindowTitle(f"JPEG Metadata Editor — {folder}")


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
