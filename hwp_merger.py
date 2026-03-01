import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QListWidgetItem, QLabel, QFileDialog,
    QMessageBox, QProgressBar
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QDragEnterEvent, QDropEvent


class MergeWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, file_list, output_path):
        super().__init__()
        self.file_list = file_list
        self.output_path = output_path

    def run(self):
        try:
            import win32com.client
            import shutil

            hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            hwp.XHwpWindows.Active_XHwpWindow.Visible = False

            self.progress.emit(10, "첫 번째 파일 열기...")
            hwp.Open(self.file_list[0], "HWP", "forceopen:true")

            for i, filepath in enumerate(self.file_list[1:], 1):
                self.progress.emit(
                    int(10 + (i / len(self.file_list)) * 80),
                    f"파일 합치는 중... ({i+1}/{len(self.file_list)})"
                )

                # 문서 끝으로 이동
                hwp.Run("MoveDocEnd")

                # 현재 문서가 비어있지 않으면 구역 나누기 삽입
                hwp.Run("BreakPage")  # 쪽 나누기 (필요에 따라 BreakSection으로 변경)

                # 다른 파일 내용을 현재 커서 위치에 삽입
                hwp.Insert(filepath, "HWP", "forceopen:true")

            self.progress.emit(90, "저장 중...")
            hwp.SaveAs(self.output_path, "HWP", "")
            hwp.Quit()

            self.progress.emit(100, "완료!")
            self.finished.emit(self.output_path)

        except ImportError:
            self.error.emit("win32com 모듈이 필요합니다.\npip install pywin32")
        except Exception as e:
            self.error.emit(f"오류 발생: {str(e)}")


class DropListWidget(QListWidget):
    """드래그 앤 드롭 지원 리스트"""
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setDragDropMode(QListWidget.InternalMove)  # 내부 순서 변경
        self.setStyleSheet("""
            QListWidget {
                border: 2px dashed #aaa;
                border-radius: 8px;
                background: #fafafa;
                font-size: 13px;
            }
            QListWidget::item {
                padding: 6px;
                border-bottom: 1px solid #eee;
            }
            QListWidget::item:selected {
                background: #d0e8ff;
                color: #000;
            }
        """)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if path.lower().endswith(('.hwp', '.hwpx')):
                    self.addItem(path)
            event.acceptProposedAction()
        else:
            super().dropEvent(event)


class HwpMerger(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HWP / HWPX 파일 합치기")
        self.setMinimumSize(650, 500)
        self.init_ui()

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # 안내 라벨
        info = QLabel("📄 HWP / HWPX 파일을 드래그하거나 추가 버튼을 눌러 불러오세요.\n리스트 안에서 드래그하여 순서를 변경할 수 있습니다.")
        info.setStyleSheet("color: #555; font-size: 12px;")
        info.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(info)

        # 파일 목록
        self.list_widget = DropListWidget()
        main_layout.addWidget(self.list_widget)

        # 버튼 행
        btn_layout = QHBoxLayout()

        btn_add = QPushButton("📂 파일 추가")
        btn_add.clicked.connect(self.add_files)

        btn_remove = QPushButton("🗑 선택 삭제")
        btn_remove.clicked.connect(self.remove_selected)

        btn_up = QPushButton("⬆ 위로")
        btn_up.clicked.connect(self.move_up)

        btn_down = QPushButton("⬇ 아래로")
        btn_down.clicked.connect(self.move_down)

        btn_clear = QPushButton("✖ 전체 삭제")
        btn_clear.clicked.connect(self.list_widget.clear)

        for btn in [btn_add, btn_remove, btn_up, btn_down, btn_clear]:
            btn.setStyleSheet("""
                QPushButton {
                    padding: 6px 14px;
                    border-radius: 5px;
                    background: #4a90d9;
                    color: white;
                    font-size: 13px;
                }
                QPushButton:hover { background: #357abd; }
                QPushButton:pressed { background: #2a6099; }
            """)
            btn_layout.addWidget(btn)

        main_layout.addLayout(btn_layout)

        # 진행 바
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.status_label)

        # 합치기 버튼
        btn_merge = QPushButton("🔗 파일 합치기")
        btn_merge.setStyleSheet("""
            QPushButton {
                padding: 10px;
                border-radius: 6px;
                background: #27ae60;
                color: white;
                font-size: 15px;
                font-weight: bold;
            }
            QPushButton:hover { background: #1e8449; }
            QPushButton:pressed { background: #196f3d; }
        """)
        btn_merge.clicked.connect(self.merge_files)
        main_layout.addWidget(btn_merge)

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "HWP 파일 선택", "", "HWP 파일 (*.hwp *.hwpx)"
        )
        for f in files:
            self.list_widget.addItem(f)

    def remove_selected(self):
        for item in self.list_widget.selectedItems():
            self.list_widget.takeItem(self.list_widget.row(item))

    def move_up(self):
        row = self.list_widget.currentRow()
        if row > 0:
            item = self.list_widget.takeItem(row)
            self.list_widget.insertItem(row - 1, item)
            self.list_widget.setCurrentRow(row - 1)

    def move_down(self):
        row = self.list_widget.currentRow()
        if row < self.list_widget.count() - 1:
            item = self.list_widget.takeItem(row)
            self.list_widget.insertItem(row + 1, item)
            self.list_widget.setCurrentRow(row + 1)

    def get_file_list(self):
        return [self.list_widget.item(i).text() for i in range(self.list_widget.count())]

    def merge_files(self):
        files = self.get_file_list()
        if len(files) < 2:
            QMessageBox.warning(self, "경고", "합칠 파일이 2개 이상 필요합니다.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self, "저장 경로 선택", "merged_output.hwp", "HWP 파일 (*.hwp)"
        )
        if not output_path:
            return

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.worker = MergeWorker(files, output_path)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_progress(self, value, msg):
        self.progress_bar.setValue(value)
        self.status_label.setText(msg)

    def on_finished(self, path):
        self.status_label.setText("✅ 합치기 완료!")
        QMessageBox.information(self, "완료", f"저장 완료:\n{path}")

    def on_error(self, msg):
        self.progress_bar.setVisible(False)
        self.status_label.setText("❌ 오류 발생")
        QMessageBox.critical(self, "오류", msg)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = HwpMerger()
    win.show()
    sys.exit(app.exec_())
