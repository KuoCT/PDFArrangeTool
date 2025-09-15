import sys
import os
from pypdfium2 import PdfDocument
from pathlib import Path
import webbrowser
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QApplication, QMainWindow, \
    QListWidget, QListWidgetItem, QMenu, QFileDialog, QToolBar, QMessageBox, QSizePolicy, \
    QSlider, QRubberBand
from PySide6.QtGui import QPixmap, QImage, QIcon, QAction
from PySide6.QtCore import Qt, QSize, Signal, QPoint, QRect
from qt_material import apply_stylesheet 

def extract_single_page(path: str, page_index: int, output_path: str):
    reader = PdfReader(path)
    writer = PdfWriter()
    writer.add_page(reader.pages[page_index])
    with open(output_path, "wb") as f:
        writer.write(f)

def render_pdf_page(pdf_path: str, page_index: int, scale: float = 0.7) -> QPixmap:
        """使用 pypdfium2 渲染單頁 PDF 成為 QPixmap"""
        pdf = PdfDocument(pdf_path)
        page = pdf[page_index]

        # 使用該頁原始的旋轉角度
        rotation = page.get_rotation()

        # 使用 RGB 格式渲染（可改用 RGBA）
        render = page.render(
            scale=scale,
            rotation=rotation,
            grayscale=False
        )
        
        pil_image = render.to_pil()
        pil_image = pil_image.convert('RGBA')

        # QImage 需知道正確格式
        qimage = QImage(
            pil_image.tobytes(),
            pil_image.width,
            pil_image.height,
            pil_image.width * 4,
            QImage.Format.Format_RGBA8888
        )
        return QPixmap.fromImage(qimage)

class PDFPageWidget(QWidget):
    """A widget that shows a page thumbnail on top and a label below."""
    def __init__(self, pixmap: QPixmap, label: str, base_grid_w: int = 150):
        super().__init__()
        self._orig_pixmap = pixmap  # keep original (high quality) pixmap
        self._icon_target_w = 100   # base icon width used at 100%
        self._icon_target_h = 140   # base icon height used at 100%

        layout = QVBoxLayout()
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        self.icon_label = QLabel()
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.text_label = QLabel(label)
        self.text_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.text_label.setWordWrap(True)

        layout.addWidget(self.icon_label)
        layout.addWidget(self.text_label)
        self.setLayout(layout)

        # initialize at 100%
        self.update_scale(1.0, base_grid_w)

    def update_scale(self, zoom: float, grid_w: int) -> None:
        """Scale icon and text width according to zoom and grid width."""
        # Scale icon with smooth transformation
        target_w = max(32, int(self._icon_target_w * zoom))
        target_h = max(45, int(self._icon_target_h * zoom))
        scaled = self._orig_pixmap.scaled(
            target_w,
            target_h,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        self.icon_label.setPixmap(scaled)

        # text width: fit inside grid width minus margins (8+8) and a small buffer
        inner_w = max(40, grid_w - 20)  # leave a little breathing room
        self.text_label.setFixedWidth(inner_w)

class PDFListWidget(QListWidget):
    zoomChanged = Signal(float)
    def __init__(self):
        super().__init__()

        # 外觀設定：圖文上下、格子排列、拖曳排序
        self.setViewMode(QListWidget.ViewMode.ListMode)
        self.setFlow(QListWidget.Flow.LeftToRight)
        self.setWrapping(True)
        self.setResizeMode(QListWidget.ResizeMode.Adjust)
        self._base_grid_w = 150
        self._base_grid_h = 200
        self.setGridSize(QSize(self._base_grid_w, self._base_grid_h))
        self.setSpacing(10)
        self.setContentsMargins(10, 10, 10, 10)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

        # 拖曳與選取
        self.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setDragDropOverwriteMode(False)

        # zoom state (1.0 = 100%)
        self._zoom = 1.0

        # --- rubber band (選取框) ---
        self._rubber_active = False
        self._rubber_origin = QPoint()
        # 用 viewport 畫，避免捲動座標錯亂
        self._band = QRubberBand(QRubberBand.Shape.Rectangle, self.viewport())
        # 半透明填色層
        self._band_fill = QWidget(self.viewport())
        self._band_fill.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        # self._band_fill.setStyleSheet(
        #     "background-color: rgba(0, 120, 215, 50);"
        #     "border: 1px solid rgba(0, 120, 215, 170);"
        # )
        self._band_fill.hide()
        self.setStyleSheet("""
            QListWidget { 
                padding: 5px; 
                background-color: #474747; 
                
            }
        """)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            # 只在「空白處」起手才開啟橡皮筋；按在 item 上交給原本拖曳排序
            if self.itemAt(event.position().toPoint()) is None:
                self._rubber_active = True
                self._rubber_origin = event.position().toPoint()
                self._band.setGeometry(QRect(self._rubber_origin, self._rubber_origin))
                self._band.show()
                self._band_fill.setGeometry(self._band.geometry())
                self._band_fill.show()
                # 若沒按 Ctrl，先清空舊選取；按 Ctrl 則累加選取
                if not (QApplication.keyboardModifiers() & Qt.KeyboardModifier.ControlModifier):
                    self.clearSelection()
                return  # 不給父類處理，避免觸發其他行為
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._rubber_active:
            rect = QRect(self._rubber_origin, event.position().toPoint()).normalized()
            self._band.setGeometry(rect)
            self._band_fill.setGeometry(rect)

            # 依矩形與每個 item 的可視區碰撞做選取
            for i in range(self.count()):
                it = self.item(i)
                it_rect = self.visualItemRect(it)
                if rect.intersects(it_rect):
                    it.setSelected(True)
                else:
                    # 若按著 Ctrl 則累加，不取消既有；否則維持僅矩形內
                    if not (QApplication.keyboardModifiers() & Qt.KeyboardModifier.ControlModifier):
                        it.setSelected(False)
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._rubber_active and event.button() == Qt.MouseButton.LeftButton:
            self._rubber_active = False
            self._band.hide()
            self._band_fill.hide()
            return
        super().mouseReleaseEvent(event)

    @staticmethod
    def _soft_wrap_long_tokens(filename: str, max_chars: int) -> str:
        """
        Insert zero-width space every `max_chars` in long continuous segments (stem part).
        Keep extension intact.
        """
        stem, ext = os.path.splitext(filename)

        # 對於沒有分隔的長字串，固定每 max_chars 插入 ZWSP
        out = []
        count = 0
        for ch in stem:
            out.append(ch)
            count += 1
            if count >= max_chars:
                out.append('\u200b')  # 可換行點
                count = 0
        return ''.join(out) + ext
    
    def set_zoom(self, zoom: float) -> None:
        """Update grid/item/widget scales together."""
        self._zoom = max(0.5, min(5.0, zoom))  # clamp

        # new grid size based on zoom (keep aspect)
        grid_w = int(self._base_grid_w * self._zoom)
        grid_h = int(self._base_grid_h * self._zoom)
        self.setGridSize(QSize(grid_w, grid_h))

        # update each item's size hint + widget scale
        for i in range(self.count()):
            item = self.item(i)
            item.setSizeHint(QSize(grid_w, grid_h))
            widget = self.itemWidget(item)
            if isinstance(widget, PDFPageWidget):
                widget.update_scale(self._zoom, grid_w)

        # labels may need width refresh for wrapping
        self._refresh_wrapped_labels()

        # 通知外界（MainWindow / slider）
        self.zoomChanged.emit(self._zoom)

    def wheelEvent(self, event):
        """Ctrl + wheel to zoom."""
        if QApplication.keyboardModifiers() & Qt.KeyboardModifier.ControlModifier:
            delta = event.angleDelta().y()
            step = 0.1 if delta > 0 else -0.1
            self.set_zoom(self._zoom + step)
            event.accept()
        else:
            super().wheelEvent(event)
    
    def _refresh_wrapped_labels(self):
        """Recompute wrapping based on current grid width."""
        grid_w = self.gridSize().width()
        fm = self.fontMetrics()
        avg_w = max(1, fm.averageCharWidth())
        # estimate how many chars per line (minus margins)
        max_chars = max(4, (grid_w - 24) // avg_w)

        for i in range(self.count()):
            item = self.item(i)
            path, page = item.data(Qt.ItemDataRole.UserRole)
            file_name = Path(path).name
            wrapped_name = self._soft_wrap_long_tokens(file_name, max_chars)
            label_text = f"{wrapped_name} - p{page + 1}"

            widget = self.itemWidget(item)
            if isinstance(widget, PDFPageWidget):
                # update text
                widget.text_label.setText(label_text)
                # ensure width matches grid
                widget.update_scale(self._zoom, grid_w)

    def add_pdf(self, path: str):
        doc = PdfDocument(path)
        path_obj = Path(path)

        grid_w = self.gridSize().width()
        fm = self.fontMetrics()
        avg_w = max(1, fm.averageCharWidth())
        max_chars = max(4, (grid_w - 24) // avg_w)

        for i in range(len(doc)):
            pixmap = render_pdf_page(path, page_index=i, scale=0.7)

            file_name = path_obj.name
            wrapped_name = self._soft_wrap_long_tokens(file_name, max_chars)
            label_text = f"{wrapped_name} - p{i + 1}"

            item = QListWidgetItem()
            item.setSizeHint(self.gridSize())
            item.setData(Qt.ItemDataRole.UserRole, (str(path_obj), i))
            item.setToolTip(file_name)
            self.addItem(item)

            widget = PDFPageWidget(pixmap, label_text, base_grid_w=grid_w)
            widget.setToolTip(f"{file_name} - p{i + 1}")
            self.setItemWidget(item, widget)

    def contextMenuEvent(self, event):
        selected_items = self.selectedItems()
        if not selected_items:
            return

        menu = QMenu(self)
        delete_action = QAction("Delete Selected Pages", self)
        menu.addAction(delete_action)

        action = menu.exec(self.mapToGlobal(event.position().toPoint()))
        if action == delete_action:
            for item in selected_items:
                self.takeItem(self.row(item))

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Delete:
            selected_items = self.selectedItems()
            for item in selected_items:
                self.takeItem(self.row(item))
            event.accept()
        else:
            super().keyPressEvent(event)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        if not event.mimeData().hasUrls():
            super().dropEvent(event)
            self.refresh_labels()
            return

        urls = event.mimeData().urls()
        paths_to_add = []
        existing = {(self.item(i).data(Qt.ItemDataRole.UserRole)[0].lower(),
                     self.item(i).data(Qt.ItemDataRole.UserRole)[1])
                    for i in range(self.count())}

        for url in urls:
            local_path = url.toLocalFile()
            if not local_path:
                continue

            if os.path.isdir(local_path):
                for root, _, files in os.walk(local_path):
                    for name in files:
                        if name.lower().endswith(".pdf"):
                            full_path = os.path.abspath(os.path.join(root, name))
                            if (full_path.lower(), -1) not in existing:
                                paths_to_add.append(full_path)
            else:
                if local_path.lower().endswith(".pdf"):
                    full_path = os.path.abspath(local_path)
                    if (full_path.lower(), -1) not in existing:
                        paths_to_add.append(full_path)

        for pdf_path in paths_to_add:
            self.add_pdf(pdf_path)

        if paths_to_add:
            event.acceptProposedAction()
        else:
            event.ignore()

    def refresh_labels(self):
        self._refresh_wrapped_labels()

    def export_pdf(self, output_path="output.pdf"):
        merger = PdfMerger()
        temp_files = []

        for i in range(self.count()):
            path, page = self.item(i).data(Qt.ItemDataRole.UserRole)
            tmp_path = f"tmp_{i}.pdf"
            extract_single_page(path, page, tmp_path)
            temp_files.append(tmp_path)
            merger.append(tmp_path)

        merger.write(output_path)
        merger.close()

        for tmp in temp_files:
            try:
                os.remove(tmp)
            except Exception:
                pass

        webbrowser.open(output_path)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDFArrangeTool")
        self.setAcceptDrops(True)

        self.pdf_list = PDFListWidget()
        self.setCentralWidget(self.pdf_list)

        # 工具列
        toolbar = QToolBar("Toolbar")
        self.addToolBar(toolbar)

        open_action = QAction("Import PDF", self)
        merge_export_action = QAction("Merge", self)
        split_export_action = QAction("Split", self)
        clear_action = QAction("Clear", self)

        # 縮放滑桿
        zoom_slider = QSlider(Qt.Orientation.Horizontal, self)
        zoom_slider.setRange(50, 500)     # 50% ~ 500%
        zoom_slider.setValue(100)         # default 100%
        zoom_slider.setFixedWidth(160)
        zoom_slider.setToolTip("Zoom (Ctrl + mouse wheel)")

        def on_zoom_from_slider(val: int):
            self.pdf_list.set_zoom(val / 100.0)

        zoom_slider.valueChanged.connect(on_zoom_from_slider)

        def on_zoom_changed_from_view(z: float):
            zoom_slider.blockSignals(True)
            zoom_slider.setValue(int(round(z * 100)))
            zoom_slider.blockSignals(False)

        # 連接 PDFListWidget 的訊號
        self.pdf_list.zoomChanged.connect(on_zoom_changed_from_view)
        zoom_label = QLabel("Zoom")

        toolbar.addAction(open_action)
        toolbar.addAction(merge_export_action)
        toolbar.addAction(split_export_action)
        toolbar.addAction(clear_action)
        toolbar.addSeparator()
        
        # 加一個彈性 spacer，把後面的東西推到右邊
        spacer = QWidget()
        spacer.setMinimumWidth(10)  # 基礎寬度
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        toolbar.addWidget(spacer)

        # 第一個固定寬度 spacer
        margin_spacer_1 = QWidget()
        margin_spacer_1.setFixedWidth(10)
        margin_spacer_1.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Preferred)

        # 第二個固定寬度 spacer
        margin_spacer_2 = QWidget()
        margin_spacer_2.setFixedWidth(10)
        margin_spacer_2.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Preferred)

        toolbar.addWidget(zoom_label)
        toolbar.addWidget(margin_spacer_1)
        toolbar.addWidget(zoom_slider)
        toolbar.addWidget(margin_spacer_2)

        open_action.triggered.connect(self.open_pdf)
        merge_export_action.triggered.connect(self.merge_export_pdf)
        split_export_action.triggered.connect(self.split_export_pdf)
        clear_action.triggered.connect(self.clear_list)

    def open_pdf(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select PDF Files", "", "PDF Files (*.pdf)"
        )
        for f in files:
            self.pdf_list.add_pdf(f)

    
    def show_done_message(self, export_dir: str):
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setWindowTitle("Done")
        msg.setText(f"Exported to: {export_dir}")

        # 加入按鈕，並設定回傳值
        open_btn = msg.addButton("Open Folder", QMessageBox.ButtonRole.AcceptRole)
        ok_btn = msg.addButton("OK", QMessageBox.ButtonRole.RejectRole)

        # 顯示訊息框
        msg.exec()

        # 檢查是哪個按鈕被按下
        if msg.clickedButton().text() == "Open Folder":
            webbrowser.open(export_dir)

    def merge_export_pdf(self):
        if self.pdf_list.count() == 0:
            QMessageBox.warning(self, "Error", "No pages to export!")
            return
        file, _ = QFileDialog.getSaveFileName(
            self, "Save Merged PDF", "output.pdf", "PDF Files (*.pdf)"
        )
        if not file:
            return
        
        export_dir = os.path.dirname(file)
        self.pdf_list.export_pdf(file)
        self.show_done_message(export_dir)

    def split_export_pdf(self):
        if self.pdf_list.count() == 0:
            QMessageBox.warning(self, "Error", "No pages to export!")
            return

        # 讓使用者選一個檔案位置，但我們實際只取路徑與前綴
        file, _ = QFileDialog.getSaveFileName(
            self, "Save Split PDFs", "output_split.pdf", "PDF Files (*.pdf)"
        )
        if not file:
            return

        export_dir = os.path.dirname(file)
        prefix = Path(file).stem  # 去掉副檔名，作為 prefix

        num_digits = len(str(self.pdf_list.count()))
        for i in range(self.pdf_list.count()):
            path, page = self.pdf_list.item(i).data(Qt.ItemDataRole.UserRole)

            output_name = f"{prefix}_p{str(i + 1).zfill(num_digits)}.pdf"
            output_path = os.path.join(export_dir, output_name)

            extract_single_page(path, page, output_path)

        self.show_done_message(export_dir)

    def clear_list(self):
        self.pdf_list.clear()

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        script_path_base = Path(sys.executable).resolve().parent
    else:
        script_path_base = Path(__file__).resolve().parent
    app = QApplication(sys.argv)
    icon_path = script_path_base / "icons" / "icon.png"
    app.setWindowIcon(QIcon(str(icon_path)))
    theme_path = script_path_base / "theme" / "custom.xml"
    apply_stylesheet(app, theme = str(theme_path), invert_secondary = False, extra = {'font_size': 16})
    window = MainWindow()
    window.resize(800, 600)
    window.show()
    sys.exit(app.exec())