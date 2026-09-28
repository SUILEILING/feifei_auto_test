
import sys
import os
import ast
import json
import time
import importlib
from datetime import datetime
from typing import List, Dict, Optional

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QTableWidget, QTableWidgetItem, QPushButton,
    QProgressBar, QLabel, QLineEdit, QCheckBox, QHeaderView,
    QMessageBox, QTabWidget, QAbstractItemView, QMenu,
    QAction, QToolBar, QFrame, QSplitter, QInputDialog, QSizePolicy, QSpinBox,
    QDoubleSpinBox,
    QComboBox, QStackedWidget, QListWidget, QListWidgetItem, QPlainTextEdit,
    QScrollArea, QStyledItemDelegate, QDialog, QDialogButtonBox, QGridLayout,
    QTreeWidget, QTreeWidgetItem
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QUrl, QFileSystemWatcher, QTimer, QItemSelectionModel
from PyQt5.QtGui import QPixmap, QFont, QColor, QIcon, QBrush, QKeySequence
from PyQt5.QtGui import QDesktopServices

import souren_config
from souren_manager import SourenManager
from adb_integration import ADBFlightModeController

def long_path(path):
    if not path:
        return path
    if sys.platform.startswith("win"):
        try:
            abspath = os.path.abspath(path)
        except Exception:
            return path
        if abspath.startswith("\\\\?\\"):
            return abspath
        if abspath.startswith("\\\\"):
            return "\\\\?\\UNC\\" + abspath[2:]
        return "\\\\?\\" + abspath
    return path

def project_root():
    candidates = []
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        candidates.append(exe_dir)
        candidates.append(os.path.dirname(exe_dir))
    candidates.append(os.getcwd())
    candidates.append(os.path.dirname(os.path.abspath(__file__)))
    for d in candidates:
        if d and os.path.exists(os.path.join(d, "souren_config.py")):
            return d
    return candidates[0] if candidates else os.getcwd()

class MultiLineParamDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QPlainTextEdit(parent)
        editor.setLineWrapMode(QPlainTextEdit.NoWrap)
        editor.setFont(QFont("Consolas", 11))
        editor.setStyleSheet(
            "QPlainTextEdit { background:#fffef5; border:2px solid #2e8b57; "
            "selection-background-color:#b3d9ff; }")
        editor.installEventFilter(self)
        return editor

    def updateEditorGeometry(self, editor, option, index):
        rect = option.rect
        view = editor.parent()
        max_w = view.width() - rect.x() - 4 if view else rect.width()
        w = max(rect.width(), 760)
        if view and w > max_w:
            w = max(max_w, 400)
        h = max(rect.height(), 160)
        editor.setGeometry(rect.x(), rect.y(), w, h)

    def setEditorData(self, editor, index):
        editor.setPlainText(index.data(Qt.EditRole) or index.data(Qt.DisplayRole) or "")

    def setModelData(self, editor, model, index):
        model.setData(index, editor.toPlainText(), Qt.EditRole)

    def eventFilter(self, editor, event):
        from PyQt5.QtCore import QEvent
        if event.type() == QEvent.KeyPress:
            key = event.key()
            if key in (Qt.Key_Return, Qt.Key_Enter):
                mods = event.modifiers()
                if mods & (Qt.AltModifier | Qt.ShiftModifier):
                    editor.insertPlainText("\n")
                    return True
                self.commitData.emit(editor)
                self.closeEditor.emit(editor)
                return True
        return super().eventFilter(editor, event)


class DraggableTable(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setWordWrap(False)

        # 第0列(选择列)拖拽涂选状态
        self._painting_checks = False
        self._paint_check_state = Qt.Checked
        self._paint_anchor_row = -1

        self.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.horizontalHeader().customContextMenuRequested.connect(self._header_menu)

    def _paint_row_range(self, to_row):
        """把 anchor 到 to_row 之间所有可勾选行刷成 _paint_check_state。"""
        anchor = self._paint_anchor_row
        if anchor < 0 or to_row < 0:
            return
        lo, hi = (to_row, anchor) if to_row < anchor else (anchor, to_row)
        for r in range(lo, hi + 1):
            it = self.item(r, 0)
            if (it is not None and (it.flags() & Qt.ItemIsUserCheckable)
                    and it.checkState() != self._paint_check_state):
                it.setCheckState(self._paint_check_state)  # 触发 itemChanged → 计数/状态列同步

    def mousePressEvent(self, event):
        # 左键点在选择列: 进入拖拽涂选模式(点一下也能单勾, 按住拖动可连续勾)
        if event.button() == Qt.LeftButton:
            idx = self.indexAt(event.pos())
            if idx.isValid() and idx.column() == 0:
                it = self.item(idx.row(), 0)
                if it is not None and (it.flags() & Qt.ItemIsUserCheckable):
                    self._paint_check_state = (
                        Qt.Unchecked if it.checkState() == Qt.Checked else Qt.Checked)
                    self._paint_anchor_row = idx.row()
                    self._painting_checks = True
                    it.setCheckState(self._paint_check_state)
                    event.accept()
                    return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._painting_checks:
            # 贴近上/下边缘时自动滚动,方便在很长的用例表里一次拖选到底
            vp_h = self.viewport().height()
            y = event.pos().y()
            bar = self.verticalScrollBar()
            if y < 20:
                bar.setValue(bar.value() - 1)
            elif y > vp_h - 20:
                bar.setValue(bar.value() + 1)
            idx = self.indexAt(event.pos())
            if idx.isValid():
                self._paint_row_range(idx.row())
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._painting_checks:
            self._painting_checks = False
            self._paint_anchor_row = -1
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def _header_menu(self, pos):
        col = self.horizontalHeader().logicalIndexAt(pos)
        menu = QMenu(self)
        act_col = menu.addAction("复制本列(含表头)")
        act_hdr = menu.addAction("复制整行表头")
        act = menu.exec_(self.horizontalHeader().mapToGlobal(pos))
        if act == act_col and col >= 0:
            hdr = self.horizontalHeaderItem(col)
            lines = [hdr.text() if hdr else ""]
            for r in range(self.rowCount()):
                it = self.item(r, col)
                lines.append(it.text() if it else "")
            QApplication.clipboard().setText("\n".join(lines))
        elif act == act_hdr:
            hdrs = [self.horizontalHeaderItem(c).text() if self.horizontalHeaderItem(c) else ""
                    for c in range(self.columnCount())]
            QApplication.clipboard().setText("\t".join(hdrs))

    def keyPressEvent(self, event):

        if event.matches(QKeySequence.Copy):
            self._copy_selection()
            return
        super().keyPressEvent(event)

    def _copy_selection(self):
        ranges = self.selectedRanges()
        if not ranges:
            return
        rng = ranges[0]
        lines = []
        for r in range(rng.topRow(), rng.bottomRow() + 1):
            cells = []
            for c in range(rng.leftColumn(), rng.rightColumn() + 1):
                it = self.item(r, c)
                cells.append(it.text() if it else "")
            lines.append("\t".join(cells))
        QApplication.clipboard().setText("\n".join(lines))

    def dropEvent(self, event):
        if event.source() is not self:
            return
        drag_row = self.currentRow()
        drop_row = self.indexAt(event.pos()).row()
        if drop_row == -1 or drop_row == drag_row:
            return
        self.blockSignals(True)
        col_count = self.columnCount()
        drag_data = [self.item(drag_row, c) for c in range(col_count)]
        drop_data = [self.item(drop_row, c) for c in range(col_count)]
        for c in range(col_count):
            self.setItem(drag_row, c, drop_data[c])
            self.setItem(drop_row, c, drag_data[c])
        self.blockSignals(False)
        self._update_row_numbers()
        event.accept()

    def _update_row_numbers(self):
        for row in range(self.rowCount()):
            item = self.item(row, 1)
            if item:
                item.setText(str(row + 1))

class CopyableTable(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.horizontalHeader().customContextMenuRequested.connect(self._header_menu)
        self.verticalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.verticalHeader().customContextMenuRequested.connect(self._row_menu)

    def _row_menu(self, pos):
        row = self.verticalHeader().logicalIndexAt(pos)
        if row < 0:
            return
        sel_rows = sorted({idx.row() for idx in self.selectedIndexes()}, reverse=True)
        if row not in sel_rows:
            sel_rows = [row]
        menu = QMenu(self)
        label = f"删除选中的 {len(sel_rows)} 行" if len(sel_rows) > 1 else "删除本行"
        act_del = menu.addAction(f"{label}(仅从表格移除, 刷新后恢复)")
        act = menu.exec_(self.verticalHeader().mapToGlobal(pos))
        if act == act_del:
            for r in sel_rows:
                self.removeRow(r)

    def _header_menu(self, pos):
        col = self.horizontalHeader().logicalIndexAt(pos)
        menu = QMenu(self)
        act_col = menu.addAction("复制本列(含表头)")
        act_hdr = menu.addAction("复制整行表头")
        act_del = menu.addAction("删除本列(仅从表格移除, 刷新后恢复)")
        act = menu.exec_(self.horizontalHeader().mapToGlobal(pos))
        if act == act_col and col >= 0:
            hdr = self.horizontalHeaderItem(col)
            lines = [hdr.text() if hdr else ""]
            for r in range(self.rowCount()):
                it = self.item(r, col)
                lines.append(it.text() if it else "")
            QApplication.clipboard().setText("\n".join(lines))
        elif act == act_hdr:
            hdr = self.horizontalHeader()
            hdrs = []
            for visual in range(self.columnCount()):
                logical = hdr.logicalIndex(visual)
                it = self.horizontalHeaderItem(logical)
                hdrs.append(it.text() if it else "")
            QApplication.clipboard().setText("\t".join(hdrs))
        elif act == act_del and col >= 0:
            self._delete_column(col)

    def _delete_column(self, col):
        self.removeColumn(col)
        # 汇总表的表头筛选按列号记忆(_rate_filter_cols/_rate_filters), 删列后列号左移一位, 同步修正
        if hasattr(self, '_rate_filter_cols'):
            self._rate_filter_cols = {c - 1 if c > col else c
                                      for c in self._rate_filter_cols if c != col}
        if hasattr(self, '_rate_filters'):
            self._rate_filters = {(c - 1 if c > col else c): v
                                  for c, v in self._rate_filters.items() if c != col}

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Copy):
            ranges = self.selectedRanges()
            if ranges:
                rng = ranges[0]
                hdr = self.horizontalHeader()
                visual_cols = sorted(range(rng.leftColumn(), rng.rightColumn() + 1),
                                     key=lambda c: hdr.visualIndex(c))
                lines = []
                for r in range(rng.topRow(), rng.bottomRow() + 1):
                    cells = []
                    for c in visual_cols:
                        it = self.item(r, c)
                        cells.append(it.text() if it else "")
                    lines.append("\t".join(cells))
                QApplication.clipboard().setText("\n".join(lines))
            return
        if event.key() == Qt.Key_Delete:
            rows = sorted({idx.row() for idx in self.selectedIndexes()}, reverse=True)
            for r in rows:
                self.removeRow(r)
            return
        super().keyPressEvent(event)

    def wheelEvent(self, event):
        if event.modifiers() & Qt.ControlModifier:
            delta = event.angleDelta().y()
            if delta != 0:
                f = self.font()
                size = f.pointSizeF()
                if size <= 0:
                    size = 10.0
                size = max(6.0, min(28.0, size + (1.0 if delta > 0 else -1.0)))
                f.setPointSizeF(size)
                self.setFont(f)
                hf = self.horizontalHeader().font()
                hf.setPointSizeF(size)
                self.horizontalHeader().setFont(hf)
                self.resizeColumnsToContents()
                self.resizeRowsToContents()
            event.accept()
            return
        super().wheelEvent(event)

def _make_rotated_axis(labels=None):
    import pyqtgraph as pg
    from PyQt5.QtGui import QFont as _QFont, QFontMetrics as _QFM

    class RotatedAxis(pg.AxisItem):
        """竖排 90° 刻度文字,字号随相邻刻度的像素间距自适应:
        鼠标缩小(每格像素变窄)时字体跟着缩小,保证标签不重叠;放大则恢复。
        轴高按最长标签的完整像素长度设定,标签必须显示全,不截断。"""
        MIN_PT, MAX_PT = 5, 10

        def drawPicture(self, p, axisSpec, tickSpecs, textSpecs):
            p.setRenderHint(p.Antialiasing, True)
            p.setRenderHint(p.TextAntialiasing, True)
            pen, p0, p1 = axisSpec
            p.setPen(pen)
            p.drawLine(p0, p1)
            for pen_, p1_, p2_ in tickSpecs:
                p.setPen(pen_)
                p.drawLine(p1_, p2_)

            if not textSpecs:
                return
            # 相邻刻度中心的最小像素间距 → 决定字号(竖排文字的横向占宽≈字体行高)
            xs = sorted(rect.center().x() for rect, _f, _t in textSpecs)
            if len(xs) >= 2 and any(b > a for a, b in zip(xs, xs[1:])):
                gap = min(b - a for a, b in zip(xs, xs[1:]) if b > a)
            else:
                gap = 9999
            pt = max(self.MIN_PT, min(self.MAX_PT, int(gap * 0.72)))
            font = _QFont(p.font())
            font.setPointSize(pt)
            p.setFont(font)

            for rect, flags, text in textSpecs:
                p.save()
                p.translate(rect.center().x(), rect.top())
                p.rotate(90)
                p.setPen(self.textPen())
                p.drawText(0, 0, text)
                p.restore()

    ax = RotatedAxis(orientation='bottom')
    # 轴高 = 最长标签在最大字号(10pt)下的完整像素长度 + 余量,保证任何缩放下都显示全
    height = 110
    if labels:
        fm = _QFM(_QFont("Microsoft YaHei", RotatedAxis.MAX_PT))
        longest = max((fm.horizontalAdvance(str(t)) for t in labels), default=0)
        height = max(110, longest + 24)
    ax.setHeight(height)
    return ax

class _EmittingStream:
    # 工作线程侧聚合: print 每行至少 write 两次(内容+换行), 长挂测每秒几百行
    # 若每次 write 都 emit 一次跨线程信号, UI 事件队列会被塞爆导致"未响应"。
    # 这里在工作线程里攒着, 每 ~50ms 或满 4KB 才 emit 一次, 信号量降两个数量级。
    FLUSH_INTERVAL = 0.05
    FLUSH_SIZE = 4096

    def __init__(self, emit_fn, original=None):
        self._emit = emit_fn
        self._original = original
        self._buf = []
        self._buf_len = 0
        self._last_emit = 0.0

    def write(self, text):
        if self._original:
            try:
                self._original.write(text)
            except Exception:
                pass
        if not text:
            return
        self._buf.append(text)
        self._buf_len += len(text)
        now = time.time()
        if self._buf_len >= self.FLUSH_SIZE or (now - self._last_emit) >= self.FLUSH_INTERVAL:
            self._emit_buf(now)

    def _emit_buf(self, now=None):
        if not self._buf:
            return
        chunk = "".join(self._buf)
        self._buf = []
        self._buf_len = 0
        self._last_emit = now if now is not None else time.time()
        try:
            self._emit(chunk)
        except Exception:
            pass

    def flush(self):
        self._emit_buf()
        if self._original:
            try:
                self._original.flush()
            except Exception:
                pass

class ExecuteThread(QThread):
    progress_update = pyqtSignal(int, int, dict)
    finished_signal = pyqtSignal(bool, str)
    log_line = pyqtSignal(str)
    retry_result = pyqtSignal(int, float)

    def __init__(self, cases: List[Dict], retry_rounds: int = 0, parent=None):
        super().__init__(parent)
        self.cases = cases
        self.retry_rounds = retry_rounds
        self.stop_requested = False
        self._remote_cleaned = False
        self._cleaning_logs = False
        self._py_tid = None

    def stop(self):
        self.stop_requested = True

        try:
            import souren_core
            souren_core.request_stop()
        except Exception as e:
            print(f"设置停止标志失败: {e}")

        if getattr(self, '_cleaning_logs', False):
            return
        self._raise_keyboard_interrupt()

    def _raise_keyboard_interrupt(self):
        import ctypes
        tid = self._py_tid
        if not tid:
            return
        try:
            res = ctypes.pythonapi.PyThreadState_SetAsyncExc(
                ctypes.c_long(tid), ctypes.py_object(KeyboardInterrupt)
            )
            if res > 1:

                ctypes.pythonapi.PyThreadState_SetAsyncExc(ctypes.c_long(tid), None)
        except Exception as e:
            print(f"注入中断失败: {e}")

    def run(self):
        import threading, sys as _sys
        self._py_tid = threading.get_ident()

        try:
            import souren_core
            souren_core.clear_stop()
        except Exception as e:
            print(f"清除停止标志失败: {e}")

        _old_out, _old_err = _sys.stdout, _sys.stderr
        stream = _EmittingStream(self.log_line.emit, _old_out)
        err_stream = _EmittingStream(self.log_line.emit, _old_err)
        _sys.stdout = stream
        _sys.stderr = err_stream
        try:
            self._run_impl()
        finally:
            # 把工作线程侧攒着的尾巴刷出去, 再还原 stdout/stderr
            try:
                stream.flush(); err_stream.flush()
            except Exception:
                pass
            _sys.stdout, _sys.stderr = _old_out, _old_err

    def _run_impl(self):
        total = len(self.cases)

        # 新一轮批跑: 复位"调试日志级别已下发"标志 → 本轮首个 case 会下发一次, 后续 case 跳过
        try:
            import common as _common
            _common.reset_debug_log_flag()
        except Exception:
            pass

        main_dir = self._get_main_execution_dir()
        all_results = []
        failed_cases = []
        for idx, case in enumerate(self.cases):
            if self.stop_requested:
                break
            rate = self._run_one_case(idx, total, case, main_dir, all_results)
            if self.stop_requested:
                break

            if self.retry_rounds > 0 and rate is not None and rate < 100:
                failed_cases.append(case)

        if self.retry_rounds > 0 and failed_cases and not self.stop_requested:
            pool = failed_cases
            for rnd in range(1, self.retry_rounds + 1):
                if self.stop_requested or not pool:
                    break
                self.log_line.emit(
                    f"\n===== 失败自动重测 第 {rnd}/{self.retry_rounds} 轮，"
                    f"待重测 {len(pool)} 个用例 =====\n")
                next_pool = []
                for idx, case in enumerate(pool):
                    if self.stop_requested:
                        break
                    rate = self._run_one_case(idx, len(pool), case, main_dir, all_results,
                                              is_retry=True, retry_round=rnd)

                    if rate is not None:
                        self.retry_result.emit(case['row'], rate)

                    if rate is not None and rate < 100:
                        next_pool.append(case)
                pool = next_pool

        while True:
            try:
                self._finalize_execution(main_dir, all_results)
                break
            except KeyboardInterrupt:
                continue

        self.finished_signal.emit(True, "执行完成" if not self.stop_requested else "已停止（已保留日志）")

    def _run_one_case(self, idx, total, case, main_dir, all_results,
                      is_retry=False, retry_round=0):
        row = case['row']
        self.progress_update.emit(idx, total, {row: {'status': 'running', 'progress': 0, 'success_rate': None}})
        manager = None
        sub_dir = None
        try:
            script_name = os.path.basename(case['script_path'])
            seq = None if is_retry else idx + 1
            sub_dir = self._create_script_subdirectory(main_dir, script_name, case['params'], seq=seq)

            manager = SourenManager(result_dir=sub_dir, params=case['params'])
            success, workflow_result = manager.run_complete_workflow(case['script_path'])
            manager.cleanup()

            if isinstance(workflow_result, dict):
                workflow_result['script_file'] = script_name
                workflow_result['script_name'] = script_name
                workflow_result['script_directory'] = sub_dir
                if case['params']:
                    workflow_result['parameters'] = case['params']
                if is_retry:
                    workflow_result['is_retry'] = True
                    workflow_result['retry_round'] = retry_round

                    try:
                        if sub_dir:
                            with open(long_path(os.path.join(sub_dir, ".retry")), "w", encoding="utf-8") as f:
                                f.write(str(retry_round))
                    except Exception:
                        pass
                all_results.append(workflow_result)

            success_rate = self._extract_success_rate(workflow_result)
            if isinstance(workflow_result, dict) and workflow_result.get('interrupted'):
                final_status = 'stopped'
                self.stop_requested = True
                self._cleanup_remote_logs(sub_dir)
            else:
                final_status = 'completed' if success else 'failed'
            err_msg = workflow_result.get('error') if isinstance(workflow_result, dict) else None
            self.progress_update.emit(idx, total, {
                row: {
                    'status': final_status,
                    'progress': 100,
                    'success_rate': success_rate,
                    'msg': err_msg or '',
                    'is_retry': is_retry,
                    'retry_round': retry_round,
                }
            })
            return success_rate
        except KeyboardInterrupt:
            self.stop_requested = True
            self._cleanup_remote_logs(sub_dir)
            try:
                if manager is not None:
                    manager.cleanup()
            except Exception:
                pass
            self.progress_update.emit(idx, total, {
                row: {'status': 'stopped', 'progress': 100, 'success_rate': None,
                      'msg': '已被用户停止（已回传远程日志）'}
            })
            return None
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            self.progress_update.emit(idx, total, {
                row: {'status': 'error', 'progress': 0, 'success_rate': None,
                      'msg': f"{e}\n{tb}"}
            })
            return None

    def _cleanup_remote_logs(self, sub_dir):
        if getattr(self, '_remote_cleaned', False):
            return
        self._remote_cleaned = True
        self._cleaning_logs = True
        try:

            try:
                time.sleep(0.05)
            except KeyboardInterrupt:
                pass
            from common import remote_cleanup_on_interrupt
            print("[*] Stop：正在停止远程日志抓取并回传 nr/mainctrl/scpi/signal 日志...")

            while True:
                try:
                    ok = remote_cleanup_on_interrupt(sub_dir)
                    break
                except KeyboardInterrupt:
                    continue
            if ok:
                print("[OK] 远程日志已回传(nr/mainctrl/scpi/signal)")
        except Exception as e:
            print(f"[WARN] 中断清理/取远程日志失败: {e}")
        finally:
            self._cleaning_logs = False

    def _finalize_execution(self, main_dir, all_results):
        if not main_dir or not all_results:
            return

        try:
            from souren_exporter import create_summary_excel
            summary_file = create_summary_excel(main_dir, all_results)
            if summary_file:
                print(f"✅ 汇总Excel已创建: {summary_file}")
            else:
                print("⚠️ 汇总Excel创建失败")
        except Exception as e:
            print(f"❌ 创建汇总Excel失败: {e}")

        try:
            if getattr(souren_config, 'COLLECT_CORE_NETWORK_LOGS', False):
                from common import remote_collect_core_logs
                core_dir = remote_collect_core_logs(main_dir)
                if core_dir:
                    print(f"✅ 核心网日志已保存: {core_dir}")
        except Exception as e:
            print(f"❌ 收集核心网日志失败: {e}")

    def _get_main_execution_dir(self):
        if souren_config.LOG_ENABLED:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            suffix = souren_config.VERSION_DIR_SUFFIX
            script_dir = project_root()
            main_dir = os.path.join(script_dir, "log", f"execution_{timestamp}{suffix}")
            os.makedirs(main_dir, exist_ok=True)
            try:
                souren_config.set_execution_dir(main_dir)
            except Exception:
                pass
            return main_dir
        return None

    def _create_script_subdirectory(self, main_dir, script_name, params, seq=None):
        if main_dir is None:
            return None
        script_base = os.path.splitext(os.path.basename(script_name))[0]
        if params:
            sub_dir_name = self._build_subdir_name(script_base, params)
        else:
            sub_dir_name = script_base
        sub_dir_name = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{sub_dir_name}"
        if seq is not None:
            sub_dir_name = f"{seq}_{sub_dir_name}"
        if len(sub_dir_name) > 70:
            sub_dir_name = sub_dir_name[:70]

        sub_dir_name = sub_dir_name.rstrip(' .')
        sub_dir = os.path.join(main_dir, sub_dir_name)
        os.makedirs(sub_dir, exist_ok=True)
        return sub_dir

    def _build_subdir_name(self, script_base, params):
        parts = [script_base]
        if 'power_class' in params:
            pc = params.get('power_class')
            parts[0] += f"_pc{pc}" if pc is not None else "_pcnone"
        for key in ['nr_band', 'lte_band', 'nr_bw', 'lte_bw', 'scs', 'range']:
            v = params.get(key)
            # 只拼标量; bw/scs 现在可能是 dict(如 {'tdd':[30],'fdd':[15]}),
            # 拼进目录名会含 {}[]:'非法路径字符,跳过
            if v is not None and not isinstance(v, (dict, list)):
                parts.append(f"{key}{v}")
        name = "_".join(parts) if len(parts) > 1 else parts[0]
        # 兜底: 去掉任何 Windows 路径非法字符
        import re
        return re.sub(r'[<>:"/\\|?*{}\[\]\',\s]+', '_', name).rstrip('_')

    def _extract_success_rate(self, workflow_result):
        loop_results = workflow_result.get('loop_results', [])
        if not loop_results:
            return None
        rates = []
        for lr in loop_results:
            res = lr.get('result', {})
            rate = res.get('success_rate')
            if rate is not None:
                rates.append(rate)
        if rates:
            return round(sum(rates)/len(rates), 2)
        return None

class ConnectionTestThread(QThread):
    result_signal = pyqtSignal(bool, str)

    def run(self):
        try:
            from souren_core import VisaInstrumentController
        except Exception as e:
            self.result_signal.emit(False, f"导入仪器模块失败: {e}")
            return
        controller = None
        try:
            controller = VisaInstrumentController()
            success, message = controller.connect()
            self.result_signal.emit(success, message)
        except Exception as e:
            import traceback
            self.result_signal.emit(False, f"{e}\n{traceback.format_exc()}")
        finally:
            try:
                if controller:
                    controller.disconnect()
            except Exception:
                pass

class ManualCommandThread(QThread):

    line_signal = pyqtSignal(str, str)
    finished_signal = pyqtSignal(bool)

    _shared_controller = None

    def __init__(self, commands, parent=None):
        super().__init__(parent)
        self.commands = commands

    @classmethod
    def reset_controller(cls):
        ctrl = cls._shared_controller
        cls._shared_controller = None
        if ctrl:
            try:
                ctrl.disconnect()
            except Exception:
                pass

    def _get_controller(self):
        from souren_core import VisaInstrumentController
        ctrl = ManualCommandThread._shared_controller
        if ctrl is None or not getattr(ctrl, 'connected', False):
            ctrl = VisaInstrumentController()
            ok, msg = ctrl.connect()
            if not ok:
                self.line_signal.emit('error', f"连接失败: {msg}")
                return None
            self.line_signal.emit('info', f"已连接: {msg}")
            ManualCommandThread._shared_controller = ctrl
        return ctrl

    def run(self):
        ctrl = self._get_controller()
        if ctrl is None:
            self.finished_signal.emit(False)
            return
        all_ok = True
        for raw in self.commands:
            cmd = (raw or '').strip()
            if not cmd:
                continue
            self.line_signal.emit('send', cmd)
            # 本地处理 SLEEP <毫秒>: 不下发给仪器, 直接睡眠(与脚本执行行为一致)
            import re
            sleep_match = re.match(r'^\s*SLEEP\s+(\d+)\s*$', cmd, re.IGNORECASE)
            if sleep_match:
                sleep_ms = int(sleep_match.group(1))
                time.sleep(sleep_ms / 1000.0)
                self.line_signal.emit('recv', f"睡眠完成 ({sleep_ms}毫秒)")
                continue
            try:
                success, result = ctrl.execute_scpi_command(cmd)
                if success:
                    self.line_signal.emit('recv', str(result))
                else:
                    all_ok = False
                    self.line_signal.emit('error', str(result))
            except Exception as e:
                all_ok = False
                self.line_signal.emit('error', str(e))
        self.finished_signal.emit(all_ok)

class HistoryLineEdit(QLineEdit):
    def __init__(self, max_history=20, parent=None):
        super().__init__(parent)
        self.max_history = max_history
        self._history = []
        self._pos = None
        self._draft = ""

    def add_history(self, cmd):
        cmd = (cmd or "").strip()
        if not cmd:
            return

        if cmd in self._history:
            self._history.remove(cmd)
        self._history.append(cmd)
        if len(self._history) > self.max_history:
            self._history = self._history[-self.max_history:]
        self._pos = None
        self._draft = ""

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Up:
            if not self._history:
                return
            if self._pos is None:
                self._draft = self.text()
                self._pos = len(self._history) - 1
            elif self._pos > 0:
                self._pos -= 1
            self.setText(self._history[self._pos])
            self.selectAll()
            return
        if event.key() == Qt.Key_Down:
            if self._pos is None:
                return
            if self._pos < len(self._history) - 1:
                self._pos += 1
                self.setText(self._history[self._pos])
            else:

                self._pos = None
                self.setText(self._draft)
            self.selectAll()
            return
        super().keyPressEvent(event)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("YC 自动化测试工具")
        self.setMinimumSize(1300, 800)

        self.thread = None
        self._suspend_case_sync = False
        self._suppress_watch_until = 0.0

        try:
            souren_config.BASE_DIR = project_root()
        except Exception:
            pass

        self._reload_external_config()
        self._init_ui()
        self._load_initial_tabs()
        self._setup_config_watcher()

    def _reload_external_config(self):
        path = os.path.join(project_root(), "souren_config.py")
        if not os.path.exists(path):
            return
        wanted = {"PHONE_SCRIPTS", "AT_SCRIPTS", "NR_SCRIPTS", "LTE_SCRIPTS", "LOG_LEVEL_PARAMS",
                  "COLLECT_CORE_NETWORK_LOGS", "LTE_MEASUREMENT_REPORT_CHECK_SKIP_SCRIPTS",
                  "DEBUG_LOG_ENABLED",
                  "DEFAULT_IP", "REMOTE_SERVER_PORT", "REMOTE_SUDO_PASSWORD",
                  "VERSION", "LOOP_COUNT"}
        try:
            with open(path, "r", encoding="utf-8-sig") as f:
                source = f.read()
            tree = ast.parse(source)
            for node in tree.body:
                if not isinstance(node, ast.Assign):
                    continue
                for t in node.targets:
                    if isinstance(t, ast.Name) and (t.id in wanted or t.id.endswith("_SCRIPTS")):
                        try:
                            val = ast.literal_eval(node.value)
                            setattr(souren_config, t.id, val)
                        except Exception:
                            pass
        except Exception as e:
            print(f"加载外部配置失败(用内置默认): {e}")

        try:
            souren_config.INSTRUMENT_ADDRESS = f"TCPIP0::{souren_config.DEFAULT_IP}::inst0::INSTR"
            souren_config.VERSION_DIR_SUFFIX = souren_config._version_dir_suffix()
        except Exception:
            pass

    def _init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        top_widget = QWidget()
        top_layout = QHBoxLayout(top_widget)
        top_layout.setContentsMargins(0, 0, 0, 0)

        self.page_selector = QComboBox()
        self.page_selector.addItems(["🤖 自动化测试", "📶 Config UE Capability", "📋 动态Case List", "🔧 手动调试", "📊 测试报告", "⚙ 调试配置"])
        self.page_selector.setFixedWidth(200)
        self.page_selector.setStyleSheet("""
            QComboBox {
                font-size: 14px; font-weight: bold; color: #1f3a5f;
                padding: 6px 10px; border: 1px solid #3498db; border-radius: 8px;
                background: #ffffff;
            }
            QComboBox:hover { border: 1px solid #2471a3; background: #eaf3fc; }
            QComboBox::drop-down { border: none; width: 22px; }
        """)
        self.page_selector.currentIndexChanged.connect(self._on_page_changed)
        top_layout.addWidget(self.page_selector, 0, Qt.AlignVCenter | Qt.AlignLeft)

        top_layout.addStretch(1)

        center_box = QHBoxLayout()
        center_box.setSpacing(10)
        logo_label = QLabel()
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "image.png")
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path)
            if not pixmap.isNull():
                pixmap = pixmap.scaledToHeight(48, Qt.SmoothTransformation)
                logo_label.setPixmap(pixmap)
        center_box.addWidget(logo_label, 0, Qt.AlignVCenter)

        title_label = QLabel("🤖 YC 自动化测试系统 ✨")
        title_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        title_label.setStyleSheet("font-size: 26px; font-weight: bold; color: #1f3a5f; letter-spacing: 2px;")
        center_box.addWidget(title_label, 0, Qt.AlignVCenter)
        top_layout.addLayout(center_box)

        top_layout.addStretch(1)

        self.phone_btn = QPushButton("📱 首次使用手机请执行")
        self.phone_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                font-weight: bold;
                border: none;
                border-radius: 8px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        self.phone_btn.clicked.connect(self.run_phone_setup)
        top_layout.addWidget(self.phone_btn, 0, Qt.AlignVCenter | Qt.AlignRight)

        main_layout.addWidget(top_widget)

        self.stack = QStackedWidget()
        main_layout.addWidget(self.stack, 1)

        auto_page = QWidget()
        auto_layout = QVBoxLayout(auto_page)
        auto_layout.setContentsMargins(0, 0, 0, 0)
        auto_layout.setSpacing(10)

        config_group = QGroupBox("⚙ 全局配置")

        config_group.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        config_layout = QVBoxLayout(config_group)

        row1 = QHBoxLayout()
        self.ip_edit = QLineEdit(souren_config.DEFAULT_IP)
        self.port_edit = QLineEdit(str(souren_config.REMOTE_SERVER_PORT))
        self.password_edit = QLineEdit(souren_config.REMOTE_SUDO_PASSWORD)
        self.password_edit.setEchoMode(QLineEdit.Password)
        row1.addWidget(QLabel("IP:"))
        row1.addWidget(self.ip_edit)
        row1.addWidget(QLabel("端口:"))
        row1.addWidget(self.port_edit)
        row1.addWidget(QLabel("密码:"))
        row1.addWidget(self.password_edit)
        row1.addStretch()
        self.test_conn_btn = QPushButton("🔌 测试连接")
        self.test_conn_btn.setObjectName("testConnBtn")
        self.test_conn_btn.clicked.connect(self.test_instrument_connection)
        row1.addWidget(self.test_conn_btn)
        config_layout.addLayout(row1)

        row2 = QHBoxLayout()
        self.version_edit = QLineEdit(souren_config.VERSION)
        self.log_collect_check = QCheckBox("收集核心网日志")
        self.log_collect_check.setChecked(souren_config.COLLECT_CORE_NETWORK_LOGS)
        row2.addWidget(QLabel("基站版本:"))
        row2.addWidget(self.version_edit)
        self.idn_btn = QPushButton("🔎 获取仪表版本/SN")
        self.idn_btn.setToolTip("向仪表下发 *IDN? 获取版本与序列号")
        self.idn_btn.clicked.connect(self.query_instrument_idn)
        row2.addWidget(self.idn_btn)
        self.idn_result = QLineEdit()
        self.idn_result.setReadOnly(True)
        self.idn_result.setPlaceholderText("*IDN? 回显显示在这里")
        self.idn_result.setMinimumWidth(360)
        row2.addWidget(self.idn_result, 1)
        row2.addStretch()
        config_layout.addLayout(row2)

        self.log_level_table = QTableWidget()
        self.log_level_table.setColumnCount(4)
        self.log_level_table.setHorizontalHeaderLabels(["参数", "值", "参数", "值"])
        self.log_level_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.log_level_table.verticalHeader().setVisible(False)
        self.log_level_table.setFont(QFont("Microsoft YaHei", 10))
        self.log_level_table.horizontalHeader().setFont(QFont("Microsoft YaHei", 10, QFont.Bold))

        self.log_level_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.log_level_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.log_level_table.setStyleSheet(
            "QTableWidget { gridline-color: #d0d0d0; font-size: 10pt; background: #ffffff; }"
            "QHeaderView::section { padding: 3px; font-size: 10pt; background: #eef3f9; }"
        )
        log_params = list(souren_config.LOG_LEVEL_PARAMS.items())
        half = (len(log_params) + 1) // 2
        self.log_level_table.setRowCount(half)
        row_h = 24
        self.log_level_table.verticalHeader().setDefaultSectionSize(row_h)
        for i in range(half):

            k, v = log_params[i]
            k_item = QTableWidgetItem(k)
            k_item.setFlags(k_item.flags() & ~Qt.ItemIsEditable)
            self.log_level_table.setItem(i, 0, k_item)
            self.log_level_table.setItem(i, 1, QTableWidgetItem(str(v)))

            j = i + half
            if j < len(log_params):
                k2, v2 = log_params[j]
                k2_item = QTableWidgetItem(k2)
                k2_item.setFlags(k2_item.flags() & ~Qt.ItemIsEditable)
                self.log_level_table.setItem(i, 2, k2_item)
                self.log_level_table.setItem(i, 3, QTableWidgetItem(str(v2)))

        self.log_level_table.setFixedHeight(row_h * half + 34)

        self.skip_scripts_table = QTableWidget()
        self.skip_scripts_table.setColumnCount(1)
        self.skip_scripts_table.setVisible(False)
        skip_scripts = souren_config.LTE_MEASUREMENT_REPORT_CHECK_SKIP_SCRIPTS
        self.skip_scripts_table.setRowCount(len(skip_scripts))
        for i, name in enumerate(skip_scripts):
            self.skip_scripts_table.setItem(i, 0, QTableWidgetItem(name))

        auto_layout.addWidget(config_group)

        cases_title = QLabel("🧪 测试用例列表")
        cases_title.setTextInteractionFlags(Qt.TextSelectableByMouse)
        cases_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #1f3a5f; padding: 4px 2px;")
        auto_layout.addWidget(cases_title)

        self.tab_widget = QTabWidget()
        self.tab_widget.setTabPosition(QTabWidget.South)
        self.tab_widget.currentChanged.connect(
            lambda _idx: self._on_search_changed(self.search_edit.text() if hasattr(self, 'search_edit') else "")
        )
        self.tab_widget.currentChanged.connect(lambda _idx: self._update_selected_count())
        self.tab_widget.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #bdc3c7; border-radius: 6px; background: #f8f9fa; top: -1px; }
            QTabBar::tab {
                background: #dfe7ef; color: #34495e;
                padding: 9px 26px; margin-right: 3px;
                border-bottom-left-radius: 6px; border-bottom-right-radius: 6px;
                font-weight: bold; font-size: 13px; min-width: 60px;
            }
            QTabBar::tab:selected { background: #3498db; color: #ffffff; }
            QTabBar::tab:hover:!selected { background: #cbd8e6; }
        """)
        # 上半区(用例表+工具条)与下半区(执行日志)放进垂直 QSplitter,
        # 执行日志框可通过拖动分隔条自由拉大/缩小
        self._upper_area = QWidget()
        upper_v = QVBoxLayout(self._upper_area)
        upper_v.setContentsMargins(0, 0, 0, 0)
        upper_v.setSpacing(6)
        upper_v.addWidget(self.tab_widget, 1)

        tab_toolbar = QHBoxLayout()
        tab_toolbar.setSpacing(8)

        self.check_all_btn = QPushButton("☑ 全部勾选")
        self.check_all_btn.setObjectName("checkAllBtn")
        self.check_all_btn.clicked.connect(lambda: self.check_all_rows(self._current_tab_name()))
        self.uncheck_all_btn = QPushButton("☐ 全部取消勾选")
        self.uncheck_all_btn.setObjectName("uncheckAllBtn")
        self.uncheck_all_btn.clicked.connect(self.refresh_checkboxes)
        self.add_case_btn = QPushButton("➕ 添加")
        self.add_case_btn.setObjectName("addBtn")
        self.add_case_btn.clicked.connect(lambda: self.add_case_row(self._current_tab_name()))
        self.del_case_btn = QPushButton("➖ 删除选中")
        self.del_case_btn.setObjectName("delBtn")
        self.del_case_btn.clicked.connect(lambda: self.delete_selected_rows(self._current_tab_name()))
        self.up_btn = QPushButton("⬆ 上移")
        self.up_btn.clicked.connect(lambda: self.move_row_up(self._current_tab_name()))
        self.down_btn = QPushButton("⬇ 下移")
        self.down_btn.clicked.connect(lambda: self.move_row_down(self._current_tab_name()))

        for b in (self.check_all_btn, self.uncheck_all_btn, self.add_case_btn,
                  self.del_case_btn, self.up_btn, self.down_btn):
            tab_toolbar.addWidget(b)

        # 已勾选计数(随勾选实时更新); 提示可拖拽涂选
        self.selected_count_label = QLabel("已勾选 0 / 共 0")
        self.selected_count_label.setStyleSheet(
            "font-weight: bold; color: #ffffff; background: #2e8b57;"
            " padding: 4px 12px; border-radius: 4px;")
        self.selected_count_label.setToolTip("在“选择”列按住鼠标左键上下拖动，可连续勾选/取消多个用例")
        tab_toolbar.addWidget(self.selected_count_label)

        search_label = QLabel("🔍")
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("搜索脚本名/参数…")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setMinimumWidth(120)
        self.search_edit.setMaximumWidth(220)
        self.search_edit.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.search_edit.textChanged.connect(self._on_search_changed)
        tab_toolbar.addWidget(search_label)
        tab_toolbar.addWidget(self.search_edit)
        tab_toolbar.addStretch()
        upper_v.addLayout(tab_toolbar)

        tab_toolbar2 = QHBoxLayout()
        tab_toolbar2.setSpacing(8)

        loop_label = QLabel("🔁 循环次数:")
        loop_label.setStyleSheet("font-weight: bold; color: #2471a3;")
        self.loop_count_spin = QSpinBox()
        self.loop_count_spin.setRange(1, 100000)
        self.loop_count_spin.setValue(int(getattr(souren_config, 'LOOP_COUNT', 1) or 1))
        self.loop_count_spin.setFixedWidth(80)
        self.loop_count_spin.setToolTip("每个勾选用例循环执行的次数")
        self.loop_count_spin.valueChanged.connect(self._on_loop_count_changed)
        tab_toolbar2.addWidget(loop_label)
        tab_toolbar2.addWidget(self.loop_count_spin)

        self.retry_check = QCheckBox("🔁 失败自动重测")
        self.retry_check.setToolTip("勾选后：每轮跑完，把通过率<100%的用例放入重测池，重跑下方次数")
        self.retry_spin = QSpinBox()
        self.retry_spin.setRange(1, 20)
        self.retry_spin.setValue(1)
        self.retry_spin.setFixedWidth(60)
        self.retry_spin.setToolTip("失败用例最多重测的轮数")

        self._load_retry_settings()
        self.retry_check.toggled.connect(lambda *_: self._save_retry_settings())
        self.retry_spin.valueChanged.connect(lambda *_: self._save_retry_settings())
        tab_toolbar2.addWidget(self.retry_check)
        tab_toolbar2.addWidget(QLabel("次数:"))
        tab_toolbar2.addWidget(self.retry_spin)

        tab_toolbar2.addStretch()

        add_sheet_btn = QPushButton("➕ 添加Sheet")
        add_sheet_btn.clicked.connect(self.add_sheet)
        del_sheet_btn = QPushButton("➖ 删除当前Sheet")
        del_sheet_btn.clicked.connect(self.delete_current_sheet)
        self.sheet_left_btn = QPushButton("◀ 左移")
        self.sheet_left_btn.setToolTip("把当前 Sheet 向左移动一位")
        self.sheet_left_btn.clicked.connect(self.move_sheet_left)
        self.sheet_right_btn = QPushButton("右移 ▶")
        self.sheet_right_btn.setToolTip("把当前 Sheet 向右移动一位")
        self.sheet_right_btn.clicked.connect(self.move_sheet_right)
        save_btn = QPushButton("💾 保存")
        save_btn.setToolTip("保存用例列表、勾选状态与列宽到本地，重启后仍生效")
        save_btn.setStyleSheet("QPushButton { background:#2e8b57; color:white; font-weight:bold; }")
        save_btn.clicked.connect(self.save_all)
        tab_toolbar2.addWidget(add_sheet_btn)
        tab_toolbar2.addWidget(del_sheet_btn)
        tab_toolbar2.addWidget(self.sheet_left_btn)
        tab_toolbar2.addWidget(self.sheet_right_btn)
        tab_toolbar2.addWidget(save_btn)
        upper_v.addLayout(tab_toolbar2)

        self.tab_action_buttons = [
            self.check_all_btn, self.uncheck_all_btn, self.add_case_btn,
            self.del_case_btn, self.up_btn, self.down_btn,
            add_sheet_btn, del_sheet_btn, self.sheet_left_btn, self.sheet_right_btn,
            save_btn, self.loop_count_spin,
            self.retry_check, self.retry_spin,
        ]

        # 日志区(标题行+日志框)打包成一个 widget, 作为 splitter 的下半区
        log_area = QWidget()
        log_area_v = QVBoxLayout(log_area)
        log_area_v.setContentsMargins(0, 0, 0, 0)
        log_area_v.setSpacing(4)

        log_header = QHBoxLayout()
        log_title = QLabel("📜 执行日志 (拖动上方分隔条可调大小)")
        log_title.setStyleSheet("font-size: 13px; font-weight: bold; color: #1f3a5f;")
        clear_log_btn = QPushButton("🧹 清空")
        clear_log_btn.setFixedWidth(70)
        clear_log_btn.clicked.connect(self._clear_exec_log)
        log_header.addWidget(log_title)
        log_header.addStretch()
        log_header.addWidget(clear_log_btn)
        log_area_v.addLayout(log_header)

        self.exec_log = QPlainTextEdit()
        self.exec_log.setReadOnly(True)
        self.exec_log.setMaximumBlockCount(5000)
        self.exec_log.setMinimumHeight(70)
        self.exec_log.setStyleSheet(
            "QPlainTextEdit { background: #1e1e1e; color: #d0d0d0; "
            "font-family: Consolas, 'Microsoft YaHei'; font-size: 9.5pt; border-radius: 6px; }"
        )
        log_area_v.addWidget(self.exec_log, 1)

        self._main_splitter = QSplitter(Qt.Vertical)
        self._main_splitter.addWidget(self._upper_area)
        self._main_splitter.addWidget(log_area)
        self._main_splitter.setCollapsible(0, False)
        self._main_splitter.setCollapsible(1, False)
        self._main_splitter.setStretchFactor(0, 1)
        self._main_splitter.setStretchFactor(1, 0)
        self._main_splitter.setSizes([600, 140])   # 初始: 日志约 140px, 和原来差不多
        self._main_splitter.setHandleWidth(8)
        self._main_splitter.setStyleSheet(
            "QSplitter::handle:vertical { background: #d5dee8; border-radius: 3px; margin: 1px 60px; }"
            "QSplitter::handle:vertical:hover { background: #3498db; }")
        auto_layout.addWidget(self._main_splitter, 1)

        self._log_buffer = []
        self._log_flush_timer = QTimer(self)
        self._log_flush_timer.setInterval(150)
        self._log_flush_timer.timeout.connect(self._flush_exec_log)
        self._log_flush_timer.start()

        bottom_widget = QWidget()
        bottom_layout = QHBoxLayout(bottom_widget)

        self.log_btn = QPushButton("📊 查看报告")
        self.log_btn.clicked.connect(self.open_log_folder)
        bottom_layout.addWidget(self.log_btn)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setFormat("总体进度: %p% (%v/%m)")
        bottom_layout.addWidget(self.progress)

        self.exec_btn = QPushButton("▶ Execute")
        self.exec_btn.clicked.connect(self.start_execution)
        bottom_layout.addWidget(self.exec_btn)

        self.stop_btn = QPushButton("⏹ Stop")
        self.stop_btn.clicked.connect(self.stop_execution)
        self.stop_btn.setEnabled(False)
        bottom_layout.addWidget(self.stop_btn)

        self.status_label = QLabel("🐣 就绪")
        self.status_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.status_label.setStyleSheet("font-weight: bold; color: #16a085;")
        bottom_layout.addWidget(self.status_label)

        auto_layout.addWidget(bottom_widget)

        self.stack.addWidget(auto_page)
        self.stack.addWidget(self._build_ue_capability_page())
        self.stack.addWidget(self._build_dyn_case_page())
        self.stack.addWidget(self._build_manual_page())
        self.stack.addWidget(self._build_report_page())
        self.stack.addWidget(self._build_debug_config_page())

        self.setStyleSheet("""
            QMainWindow { background-color: #eaf1f8; }
            QWidget { font-family: 'Microsoft YaHei'; }
            QLabel { color: #2c3e50; }
            QGroupBox {
                font-weight: bold; border: 1px solid #c3d0de; border-radius: 8px;
                margin-top: 0.6em; padding-top: 0.5em; background: #ffffff;
            }
            QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 6px; color: #2471a3; }
            QLineEdit {
                border: 1px solid #cdd5df; border-radius: 5px; padding: 5px 7px;
                background: #ffffff; selection-background-color: #b3d9ff;
            }
            QLineEdit:focus { border: 1px solid #3498db; }
            QCheckBox { color: #2c3e50; }
            QPushButton {
                padding: 6px 14px; border-radius: 6px;
                background-color: #ffffff; border: 1px solid #cdd5df; color: #2c3e50;
            }
            QPushButton:hover { background-color: #eaf3fc; border: 1px solid #3498db; }
            QPushButton:pressed { background-color: #d6e6f5; }
            QPushButton#exec_btn { background-color: #27ae60; color: white; border: none; font-weight: bold; }
            QPushButton#exec_btn:hover { background-color: #219a52; }
            QPushButton#stop_btn { background-color: #e74c3c; color: white; border: none; font-weight: bold; }
            QPushButton#stop_btn:hover { background-color: #c0392b; }
            QPushButton#checkAllBtn { background-color: #eafaf1; border: 1px solid #27ae60; color: #1e8449; font-weight: bold; }
            QPushButton#checkAllBtn:hover { background-color: #d5f5e3; border: 1px solid #1e8449; }
            QPushButton#uncheckAllBtn { background-color: #fdf2f2; border: 1px solid #e59866; color: #ba4a00; font-weight: bold; }
            QPushButton#uncheckAllBtn:hover { background-color: #fae5d3; border: 1px solid #ba4a00; }
            QPushButton#addBtn { background-color: #eaf4fd; border: 1px solid #3498db; color: #2471a3; }
            QPushButton#addBtn:hover { background-color: #d6eaf8; border: 1px solid #2471a3; }
            QPushButton#delBtn { background-color: #fdeded; border: 1px solid #e74c3c; color: #b03a2e; }
            QPushButton#delBtn:hover { background-color: #fadbd8; border: 1px solid #b03a2e; }
            QPushButton#testConnBtn { background-color: #3498db; color: white; border: none; font-weight: bold; }
            QPushButton#testConnBtn:hover { background-color: #2980b9; }
            QPushButton#testConnBtn:disabled { background-color: #aab7c4; color: #eef; }
            QPushButton#manualSendBtn { background-color: #27ae60; color: white; border: none; font-weight: bold; }
            QPushButton#manualSendBtn:hover { background-color: #219a52; }
            QPushButton#manualSendBtn:disabled { background-color: #aab7c4; color: #eef; }
            QProgressBar {
                border: 1px solid #cdd5df; border-radius: 6px; text-align: center;
                background: #ffffff; height: 22px; color: #2c3e50; font-weight: bold;
            }
            QProgressBar::chunk {
                border-radius: 5px;
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #56ccf2, stop:1 #2f80ed);
            }
            QTableWidget::item:selected {
                background-color: #b3d9ff;
                color: #000000;
            }
        """)
        self.exec_btn.setObjectName("exec_btn")
        self.stop_btn.setObjectName("stop_btn")

        for lbl in self.findChildren(QLabel):
            lbl.setTextInteractionFlags(lbl.textInteractionFlags() | Qt.TextSelectableByMouse)

    def _on_page_changed(self, idx):
        self.stack.setCurrentIndex(idx)

        if idx == 2 and hasattr(self, 'dyn_list_combo'):
            if self.dyn_list_combo.currentText() != self.DYN_DEFAULT_LIST:
                self.dyn_list_combo.setCurrentText(self.DYN_DEFAULT_LIST)
            else:
                self._dyn_refresh_table()

        if idx == 4 and hasattr(self, '_refresh_report_executions'):
            self._refresh_report_executions()

    def _build_debug_config_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(12)

        title = QLabel("⚙ 调试配置")
        title.setStyleSheet("font-size:16px; font-weight:bold; color:#1f3a5f; padding:4px 2px;")
        layout.addWidget(title)

        core_group = QGroupBox("📡 核心网日志")
        core_v = QVBoxLayout(core_group)
        core_v.addWidget(self.log_collect_check)
        layout.addWidget(core_group)

        log_group = QGroupBox("🐛 调试日志级别")
        log_v = QVBoxLayout(log_group)
        self.debug_log_check = QCheckBox("配置log")
        self.debug_log_check.setChecked(getattr(souren_config, 'DEBUG_LOG_ENABLED', True))
        self.debug_log_check.setToolTip(
            "勾选: 基站启动时下发下方调试日志级别(CONFigure:VERSion:LOG:STATe)与 *rst;\n"
            "不勾: 跳过这两条 SCPI(不抓调试 log), 基站仍正常重启。")
        log_v.addWidget(self.debug_log_check)
        log_v.addWidget(self.log_level_table)
        layout.addWidget(log_group)

        layout.addStretch()

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        save_btn = QPushButton("💾 保存配置")
        save_btn.setObjectName("checkAllBtn")
        save_btn.setFixedWidth(140)
        save_btn.clicked.connect(self._save_debug_config)
        btn_row.addWidget(save_btn)
        layout.addLayout(btn_row)

        tip = QLabel("提示：点击“保存配置”写回 souren_config.py；执行时也会自动同步。")
        tip.setStyleSheet("color:#888; padding:4px 2px;")
        layout.addWidget(tip)
        return page

    def _save_debug_config(self):
        try:
            self._update_global_config()

            self._persist_var_to_config(
                "LOG_LEVEL_PARAMS", self._format_log_levels_literal())
            self._persist_var_to_config(
                "COLLECT_CORE_NETWORK_LOGS", repr(bool(self.log_collect_check.isChecked())))
            self._persist_var_to_config(
                "DEBUG_LOG_ENABLED", repr(bool(self.debug_log_check.isChecked())))
            QMessageBox.information(self, "已保存", "调试配置已写回 souren_config.py")
            self.status_label.setText("调试配置已保存")
        except Exception as e:
            QMessageBox.warning(self, "保存失败", f"保存失败:\n{e}")

    def _format_log_levels_literal(self):
        def _parse_val(s):
            s = (s or "").strip()
            try:
                if s.lstrip('-').isdigit():
                    return int(s)
                return float(s)
            except Exception:
                return s
        lines = ["{"]
        for row in range(self.log_level_table.rowCount()):
            for kcol, vcol in ((0, 1), (2, 3)):
                key_item = self.log_level_table.item(row, kcol)
                val_item = self.log_level_table.item(row, vcol)
                if key_item and val_item and key_item.text().strip():
                    k = key_item.text().strip()
                    v = _parse_val(val_item.text())
                    lines.append(f"    {k!r}: {v!r},")
        lines.append("}")
        return "\n".join(lines)

    def _build_report_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        top = QHBoxLayout()
        top.addWidget(QLabel("📁 执行记录:"))
        self.report_exec_combo = QComboBox()
        self.report_exec_combo.setMinimumWidth(420)
        self.report_exec_combo.setMaxVisibleItems(20)   # 下拉一次可见 20 项, 其余滚动
        self.report_exec_combo.currentIndexChanged.connect(self._on_report_exec_changed)
        top.addWidget(self.report_exec_combo)
        refresh_btn = QPushButton("🔄 刷新")
        refresh_btn.clicked.connect(self._refresh_report_executions)
        top.addWidget(refresh_btn)
        edit_title_btn = QPushButton("✏ 修改标题")
        edit_title_btn.setToolTip("给当前执行记录起一个自定义标题(只改显示，不重命名文件夹)")
        edit_title_btn.clicked.connect(self._edit_report_title)
        top.addWidget(edit_title_btn)
        del_btn = QPushButton("🗑 删除此记录")
        del_btn.setObjectName("delBtn")
        del_btn.clicked.connect(self._delete_report_execution)
        top.addWidget(del_btn)
        top.addStretch()
        layout.addLayout(top)

        splitter = QSplitter(Qt.Horizontal)

        self.report_case_list = QListWidget()
        self.report_case_list.setMinimumWidth(300)
        self.report_case_list.setMaximumWidth(460)
        self.report_case_list.setStyleSheet(
            "QListWidget { background:#ffffff; border:1px solid #c3d0de; border-radius:6px; font-size:10pt; }"
            "QListWidget::item { padding:6px 8px; border-bottom:1px solid #eef2f6; }"
            "QListWidget::item:selected { background:#3498db; color:white; }"
        )
        self.report_case_list.currentRowChanged.connect(self._on_report_case_changed)
        splitter.addWidget(self.report_case_list)

        self.report_scroll = QScrollArea()
        self.report_scroll.setWidgetResizable(True)
        self.report_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.report_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.report_scroll.setStyleSheet("QScrollArea { border:1px solid #c3d0de; border-radius:6px; background:#f8fafc; }")
        self._report_content = QWidget()
        self._report_content.setMinimumWidth(900)
        self._report_content_layout = QVBoxLayout(self._report_content)
        self._report_content_layout.setAlignment(Qt.AlignTop)
        self.report_scroll.setWidget(self._report_content)
        splitter.addWidget(self.report_scroll)
        splitter.setSizes([340, 1200])
        layout.addWidget(splitter, 1)

        self._report_exec_dirs = []
        self._report_cases = []
        return page

    def _delete_report_execution(self):
        idx = self.report_exec_combo.currentIndex()
        if not (0 <= idx < len(self._report_exec_dirs)):
            return
        exec_dir = self._report_exec_dirs[idx]
        name = os.path.basename(exec_dir)
        reply = QMessageBox.question(
            self, "确认删除",
            f"确定删除执行记录\n{name}\n\n该文件夹及其下所有 case、日志、Excel 都会被永久删除，无法恢复！",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        ok, err = self._robust_rmtree(exec_dir)
        if not ok:
            QMessageBox.warning(
                self, "删除失败",
                f"删除失败，可能有文件正被占用(如日志/压缩包/Excel 还开着，或正在被抓取)。\n"
                f"请关闭相关文件后重试。\n\n{err}")
            return
        self._refresh_report_executions()

    def _robust_rmtree(self, path, retries=3, delay=0.4):
        import shutil, stat, time as _t
        def _on_error(func, p, exc_info):
            try:
                os.chmod(p, stat.S_IWRITE)
                func(p)
            except Exception:
                raise
        last_err = ""
        for i in range(retries):
            try:
                shutil.rmtree(long_path(path), onerror=_on_error)
                return True, ""
            except Exception as e:
                last_err = str(e)
                if i < retries - 1:
                    _t.sleep(delay)
        return (not os.path.exists(long_path(path))), last_err

    def _log_root(self):
        return os.path.join(project_root(), "log")

    def _edit_report_title(self):
        idx = self.report_exec_combo.currentIndex()
        if not (0 <= idx < len(self._report_exec_dirs)):
            QMessageBox.information(self, "提示", "没有可修改的执行记录")
            return
        exec_dir = self._report_exec_dirs[idx]
        old_name = os.path.basename(exec_dir)
        new_name, ok = QInputDialog.getText(
            self, "修改标题", "执行记录标题(会真正重命名文件夹):", text=old_name)
        if not ok:
            return
        new_name = new_name.strip()
        if not new_name or new_name == old_name:
            return
        import re
        safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', new_name).rstrip(' .')
        if not safe:
            QMessageBox.warning(self, "无效标题", "标题为空或全是非法字符")
            return
        new_path = os.path.join(os.path.dirname(exec_dir), safe)
        if os.path.exists(long_path(new_path)):
            QMessageBox.warning(self, "重名", f"已存在同名记录，请换一个:\n{safe}")
            return
        ok2, err = self._robust_rename_dir(exec_dir, new_path)
        if not ok2:
            QMessageBox.warning(
                self, "重命名失败",
                f"重命名失败:\n{err}\n\n"
                "可能原因：该记录里有文件正被打开/占用（如 Excel/日志/压缩包在预览），"
                "或程序工作目录还停在此文件夹内。请关闭相关文件后重试。")
            return
        self._refresh_report_executions()
        for i, d in enumerate(self._report_exec_dirs):
            if os.path.basename(d) == safe:
                self.report_exec_combo.setCurrentIndex(i)
                break

    def _release_handles_under(self, folder):
        """释放本进程对该目录内文件持有的句柄,否则 Windows 不允许重命名(WinError 5):
        - souren_execution.log 的 logging.FileHandler(整段运行都开着);
        - SCPI 流式落盘文件句柄;
        - 强制 GC 收掉未挂在 logger 上、但文件仍开着的孤儿 FileHandler。"""
        import logging as _lg, gc
        folder_n = os.path.normcase(os.path.abspath(folder))
        def _under(p):
            try:
                pn = os.path.normcase(os.path.abspath(p))
                return pn == folder_n or pn.startswith(folder_n + os.sep)
            except Exception:
                return False
        # 1) SCPI 流式落盘文件
        try:
            from souren_core import VisaInstrumentController
            sp = getattr(VisaInstrumentController, "_scpi_path", None)
            if sp and _under(sp):
                VisaInstrumentController.close_scpi_log()
        except Exception:
            pass
        # 2) 指向该目录内文件的 logging FileHandler(souren_execution.log 等)
        try:
            loggers = [_lg.getLogger()]
            loggers += [_lg.getLogger(n) for n in list(_lg.Logger.manager.loggerDict.keys())]
            for lg in loggers:
                for h in list(getattr(lg, 'handlers', [])):
                    fn = getattr(h, 'baseFilename', None)
                    if fn and _under(fn):
                        try:
                            lg.removeHandler(h)
                            h.close()
                        except Exception:
                            pass
        except Exception:
            pass
        # 3) 兜底: 强制 GC,让未挂 logger 的孤儿 FileHandler 析构关闭文件
        try:
            gc.collect()
        except Exception:
            pass

    def _robust_rename_dir(self, src, dst, retries=3, delay=0.4):
        """重命名目录,规避 Windows 常见坑:
        - 进程 CWD 停在 src 内部时,重命名其祖先会被拒(WinError 5)→ 先撤出 CWD;
        - 本进程对 src 内文件持有句柄(日志/SCPI)→ 先释放这些句柄;
        - 文件被瞬时占用(WinError 32)→ 短暂重试几次。
        返回 (是否成功, 错误信息)。"""
        import time as _t
        try:
            cwd = os.path.normcase(os.path.abspath(os.getcwd()))
            s = os.path.normcase(os.path.abspath(src))
            # CWD 等于 src 或在 src 之内 → 撤到项目根,否则 Windows 不允许重命名 src
            # (chdir 是进程级的,执行期间进过 case 子目录,可能停在该记录内)
            if cwd == s or cwd.startswith(s + os.sep):
                os.chdir(project_root())
        except Exception:
            pass
        # 释放本进程持有的、指向该记录内文件的句柄(日志/SCPI)
        self._release_handles_under(src)
        last_err = ""
        for i in range(retries):
            try:
                os.rename(long_path(src), long_path(dst))
                return True, ""
            except Exception as e:
                last_err = str(e)
                if i < retries - 1:
                    _t.sleep(delay)
        return (not os.path.exists(long_path(src))), last_err

    def _refresh_report_executions(self):
        root = self._log_root()
        dirs = []
        if os.path.isdir(root):
            dirs = [os.path.join(root, d) for d in os.listdir(root)
                    if os.path.isdir(os.path.join(root, d))]
            dirs.sort(key=lambda d: os.path.getmtime(d), reverse=True)
            dirs = dirs[:30]   # 下拉最多显示最近 30 个执行记录目录
        self._report_exec_dirs = dirs
        self.report_exec_combo.blockSignals(True)
        self.report_exec_combo.clear()
        if dirs:
            self.report_exec_combo.addItems([os.path.basename(d) for d in dirs])
        else:
            self.report_exec_combo.addItem("（暂无执行记录）")
        self.report_exec_combo.blockSignals(False)
        self.report_exec_combo.setCurrentIndex(0)
        self._on_report_exec_changed(0)

    def _on_report_exec_changed(self, idx):
        self.report_case_list.blockSignals(True)
        self.report_case_list.clear()
        self._report_cases = []
        if 0 <= idx < len(self._report_exec_dirs):
            exec_dir = self._report_exec_dirs[idx]

            self.report_case_list.addItem("📊 最终数据汇总")

            subs = [os.path.join(exec_dir, d) for d in os.listdir(exec_dir)
                    if os.path.isdir(os.path.join(exec_dir, d))]

            subs = [d for d in subs if os.path.basename(d) != "core_network_logs"]

            def _ts_key(d):
                n = os.path.basename(d)
                parts = n.split("_")
                if (len(parts) >= 3 and parts[0].isdigit() and len(parts[0]) <= 4
                        and len(parts[1]) == 8 and parts[1].isdigit() and parts[2].isdigit()):
                    return (f"{int(parts[0]):06d}", parts[1] + parts[2], os.path.getmtime(d))
                if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
                    return ("999999", parts[0] + parts[1], os.path.getmtime(d))
                return ("999999", "99999999999999", os.path.getmtime(d))
            subs.sort(key=_ts_key)   
            self._report_cases = subs
            for d in subs:
                rec = self._find_case_json(d)
                params = rec.get('parameters', {}) if isinstance(rec, dict) else {}
                if not isinstance(params, dict):
                    params = {}
                self.report_case_list.addItem(self._display_config_name(os.path.basename(d), params))
        self.report_case_list.blockSignals(False)
        if self.report_case_list.count() > 0:
            self.report_case_list.setCurrentRow(0)

    def _on_report_case_changed(self, row):

        self._clear_report_content()
        if row < 0:
            return
        try:
            if row == 0:
                self._render_final_summary()
            else:
                case_idx = row - 1
                if 0 <= case_idx < len(self._report_cases):
                    self._render_case_charts(self._report_cases[case_idx])
        except Exception as e:
            import traceback
            traceback.print_exc()
            self._clear_report_content()
            err = QLabel(f"⚠ 生成报告时出错：\n{e}")
            err.setStyleSheet("color:#c0392b; padding:16px;")
            err.setWordWrap(True)
            err.setTextInteractionFlags(Qt.TextSelectableByMouse)
            self._report_content_layout.addWidget(err)

    def _clear_report_content(self):
        while self._report_content_layout.count():
            item = self._report_content_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.setParent(None)
                w.deleteLater()

    def _find_case_json(self, case_dir):
        import glob
        js = glob.glob(os.path.join(case_dir, "*_results_*.json"))
        if not js:
            js = glob.glob(os.path.join(case_dir, "*.json"))
        if not js:
            return None
        js.sort(key=lambda f: os.path.getmtime(long_path(f)), reverse=True)
        try:
            with open(long_path(js[0]), "r", encoding="utf-8") as f:
                data = json.load(f)
            return data[0] if isinstance(data, list) and data else data
        except Exception as e:
            print(f"读取 case JSON 失败: {e}")
            return None

    def _iter_loop_results(self, rec):
        out = []
        if not isinstance(rec, dict):
            return out
        for lr in rec.get('loop_results', []):
            res = lr.get('result', {}) if isinstance(lr, dict) else {}
            out.append((lr.get('loop_index', res.get('loop_iteration', 1)), res))
        return out

    def _load_case_charts(self, case_dir):
        rec = self._find_case_json(case_dir)
        charts = {}
        if not rec:
            return charts, rec
        for loop_iter, res in self._iter_loop_results(rec):
            for e in res.get('extracted_data', []):
                title = e.get('chart_title')
                if not title:
                    continue
                if e.get('summary_only'):
                    continue  
                y = e.get('extracted_data')
                if y is None:
                    continue
                try:
                    y = float(y)
                except (TypeError, ValueError):
                    continue  
                x = str(e.get('x_label', ''))
                charts.setdefault(title, []).append((x, y, loop_iter))
        return charts, rec

    def _render_case_charts(self, case_dir):
        import pyqtgraph as pg
        charts, rec = self._load_case_charts(case_dir)
        params = rec.get('parameters', {}) if isinstance(rec, dict) else {}
        if not isinstance(params, dict):
            params = {}
        name = self._display_config_name(os.path.basename(case_dir), params)

        title_lbl = QLabel(f"🧪 {name}")
        title_lbl.setStyleSheet("font-size:14px; font-weight:bold; color:#1f3a5f; padding:6px 4px;")
        title_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self._report_content_layout.addWidget(title_lbl)

        if rec:
            loops = self._iter_loop_results(rec)
            rates = [res.get('success_rate') for _, res in loops if res.get('success_rate') is not None]
            if rates:
                avg = round(sum(rates) / len(rates), 2)
                info = QLabel(f"循环次数: {len(loops)}   平均通过率: {avg}%")
                info.setStyleSheet("color:#2471a3; padding:2px 6px;")
                self._report_content_layout.addWidget(info)

        if not charts:
            empty = QLabel("（该用例无图表数据）")
            empty.setStyleSheet("color:#888; padding:20px;")
            self._report_content_layout.addWidget(empty)
            return

        pg.setConfigOptions(antialias=True, background='w', foreground='k')
        for title, points in charts.items():
            self._add_bar_chart(title, points)

    def _add_bar_chart(self, title, points):
        import pyqtgraph as pg

        cap = QLabel(f"📊 {title}")
        cap.setStyleSheet("font-size:13px; font-weight:bold; color:#34495e; padding:8px 4px 2px 4px;")
        self._report_content_layout.addWidget(cap)

        loops = sorted({lp for _, _, lp in points})
        loop_seq = {lp: [] for lp in loops}   
        for x, y, lp in points:
            loop_seq[lp].append((x, y))

        n_loop = len(loops)
        n_x = max((len(seq) for seq in loop_seq.values()), default=0)   

        longest_lp = max(loops, key=lambda lp: len(loop_seq[lp])) if loops else None
        x_labels = [xl for xl, _ in loop_seq[longest_lp]] if longest_lp is not None else []

        # 长标签(如 NSA 的 "NR:78 range:Low OB:1 BW:20 (300)")横排必重叠,
        # 点数多或标签超过 8 字符就改用竖排自适应字号的轴(轴高按最长标签算,完整显示不截断)
        max_label_len = max((len(str(x)) for x in x_labels), default=0)
        if n_x > 8 or max_label_len > 8:
            plot = pg.PlotWidget(axisItems={'bottom': _make_rotated_axis(x_labels)})
        else:
            plot = pg.PlotWidget()
        plot.setMouseEnabled(x=True, y=True)
        plot.showGrid(x=False, y=True, alpha=0.3)
        vb = plot.getViewBox()
        vb.setMenuEnabled(True)
        if n_loop > 1:
            plot.addLegend(offset=(-10, 10))

        group_w = 0.9
        bar_w = group_w / max(n_loop, 1)
        colors = [(52, 152, 219), (231, 76, 60), (46, 204, 113), (155, 89, 182),
                  (241, 196, 15), (26, 188, 156)]

        loop_data = {lp: {} for lp in loops}
        all_ys = []
        FAIL_COLOR = (231, 76, 60)   
        for li, lp in enumerate(loops):
            xs, hs = [], []               
            fxs, fhs = [], []            
            for pos, (xl, y) in enumerate(loop_seq[lp]):
                bx = pos - group_w / 2 + bar_w * (li + 0.5)
                loop_data[lp][pos] = y
                all_ys.append(y)
                if "[失败]" in str(xl):
                    fxs.append(bx); fhs.append(y)
                else:
                    xs.append(bx); hs.append(y)
            c = colors[li % len(colors)]
            if xs:
                bars = pg.BarGraphItem(x=xs, height=hs, width=bar_w * 0.96,
                                       brush=(*c, 220), pen=pg.mkPen(c, width=2.5))
                plot.addItem(bars)
            if fxs:
                fbars = pg.BarGraphItem(x=fxs, height=fhs, width=bar_w * 0.96,
                                        brush=(*FAIL_COLOR, 230),
                                        pen=pg.mkPen(FAIL_COLOR, width=2.5))
                plot.addItem(fbars)
            if n_loop > 1:

                proxy = pg.ScatterPlotItem([0], [0], size=12, symbol='s',
                                           brush=(*c, 220), pen=pg.mkPen(c))
                proxy.setVisible(False)
                plot.plotItem.legend.addItem(proxy, f"循环{lp}")

        ax = plot.getAxis('bottom')
        ax.setTicks([[(i, x_labels[i]) for i in range(min(n_x, len(x_labels)))]])
        ax.setStyle(tickTextOffset=8)
        plot.setLabel('left', '值')

        if all_ys:
            ymin, ymax = min(all_ys), max(all_ys)
        else:
            ymin, ymax = 0.0, 1.0
        if abs(ymax) < 1e-6 and abs(ymin) < 1e-6:
            # 全 0(如 BLER 全 0)也给个可视范围,柱顶数值标签不会被顶边裁掉
            lo, hi = -0.15, 1.0
        else:
            posspan = (ymax - min(0.0, ymin)) or abs(ymax) or 1.0
            # 顶部留 28% 余量放数值标签,底部留 15%,避免标签贴边/被裁
            hi = (ymax if ymax > 0 else 0.0) + posspan * 0.28
            lo = (min(0.0, ymin)) - posspan * 0.15
        plot.setYRange(lo, hi, padding=0)
        plot.setXRange(-0.6, max(n_x - 0.4, 0.6), padding=0)
        vb.disableAutoRange()
        span = (hi - lo) or 1.0
        for li, lp in enumerate(loops):
            for pos, (xl, y) in enumerate(loop_seq[lp]):
                px = pos - group_w / 2 + bar_w * (li + 0.5)
                lbl_color = (200, 30, 30) if "[失败]" in str(xl) else (40, 40, 40)
                if y >= 0:
                    txt = pg.TextItem(f"{y:.3f}", color=lbl_color, anchor=(0.5, 1.0))
                    txt.setPos(px, y + span * 0.015)
                else:
                    txt = pg.TextItem(f"{y:.3f}", color=lbl_color, anchor=(0.5, 0.0))
                    txt.setPos(px, y - span * 0.015)
                txt.setFont(QFont("Microsoft YaHei", 8))
                plot.addItem(txt)

        if n_x > 8:
            per = 82
            inner_w = max(1000, n_x * per)
            plot.setMinimumWidth(inner_w)
            plot.setMinimumHeight(460)   # 绘图区更高,柱子/数值看得清
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
            scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            scroll.setFixedHeight(560)   # 容纳绘图区 + 竖排长标签轴
            scroll.setWidget(plot)
            self._report_content_layout.addWidget(scroll)
        else:
            plot.setMinimumHeight(420)
            plot.setFixedHeight(420)
            self._report_content_layout.addWidget(plot)

    def _render_final_summary(self):
        idx = self.report_exec_combo.currentIndex()
        if not (0 <= idx < len(self._report_exec_dirs)):
            return
        exec_dir = self._report_exec_dirs[idx]

        title_lbl = QLabel("📊 最终数据汇总")
        title_lbl.setStyleSheet("font-size:14px; font-weight:bold; color:#1f3a5f; padding:6px 4px;")
        self._report_content_layout.addWidget(title_lbl)

        rows = []
        metric_titles = []
        retry_by_round = {}            # cfg_key -> {轮次: 该轮通过率}
        retry_metrics = {}             # cfg_key -> {指标: (轮次, 值)} 重测里测到的指标, 用于回填原行空缺
        nonretry_seen = set()          # 已作为“原始行”出现过的配置 key
        fallback_round = {}
        for case_dir in self._report_cases:
            charts, rec = self._load_case_charts(case_dir)
            if not rec:
                continue
            name = os.path.basename(case_dir)

            cfg_key = self._config_key_from_rec(rec, name)

            retry_marker = os.path.join(case_dir, ".retry")
            is_retry = os.path.exists(long_path(retry_marker))
            retry_round = 0
            if is_retry:
                try:
                    with open(long_path(retry_marker), encoding="utf-8") as f:
                        retry_round = int((f.read() or "0").strip() or 0)
                except Exception:
                    retry_round = 1
            elif cfg_key in nonretry_seen:
                # 兼容旧数据/写标记失败: 同一配置的重复目录(按时间排在原始之后)按重测计入
                is_retry = True
                fallback_round[cfg_key] = fallback_round.get(cfg_key, 0) + 1
                retry_round = fallback_round[cfg_key]

            loops = self._iter_loop_results(rec)
            if is_retry:

                rates = [res.get('success_rate') for _, res in loops if res.get('success_rate') is not None]
                rnd = retry_round if retry_round > 0 else 1
                if rates:
                    avg = round(sum(rates) / len(rates), 2)
                    # 每一轮重测各存一份(fail_try_1/2/3...),不再只保留最后一轮
                    retry_by_round.setdefault(cfg_key, {})[rnd] = avg
                # 重测里测到的指标(MeasuredValue 等)也收集起来: 原始跑失败时指标是空的,
                # 重测通过后从这里回填, 否则汇总表里 fail_try 通过了指标却仍空白。
                rm = retry_metrics.setdefault(cfg_key, {})
                for _, res in loops:
                    for e in res.get('extracted_data', []):
                        t = e.get('chart_title')
                        y = e.get('extracted_data')
                        if not t or y is None:
                            continue
                        if e.get('summary_only'):
                            val = str(y)
                        else:
                            try:
                                val = float(y)
                            except (TypeError, ValueError):
                                val = str(y)
                        old = rm.get(t)
                        if old is None or rnd >= old[0]:   # 高轮次(更近一次重测)覆盖
                            rm[t] = (rnd, val)
                        if t not in metric_titles:
                            metric_titles.append(t)
                continue

            nonretry_seen.add(cfg_key)

            for t in charts.keys():
                if t not in metric_titles:
                    metric_titles.append(t)

            for loop_iter, res in loops:
                rate = res.get('success_rate')
                metrics = {}
                NR_TITLES = {"DL NR_BLER", "UL NR_BLER", "NR TXP AVG", "MeasuredValue"}
                LTE_TITLES = {"DL LTE_BLER", "UL LTE_BLER", "LTE TXP AVG"}
                STATUS_TEXT = {"normal": "正常", "abnormal": "回退后异常"}
                nr_power = ""; lte_power = ""; nr_st = None; lte_st = None
                for e in res.get('extracted_data', []):
                    t = e.get('chart_title')
                    y = e.get('extracted_data')
                    xl = e.get('x_label')
                    st = e.get('status')
                    if t and y is not None:
                        if e.get('summary_only'):
                            metrics[t] = str(y)   
                        else:
                            try:
                                metrics[t] = float(y)
                            except (TypeError, ValueError):
                                metrics[t] = str(y)
                        if t not in metric_titles:
                            metric_titles.append(t)
                    if t in NR_TITLES and xl is not None:
                        if st or not nr_power:
                            nr_power = str(xl)
                        if st and not nr_st:
                            nr_st = st
                    elif t in LTE_TITLES and xl is not None:
                        if st or not lte_power:
                            lte_power = str(xl)
                        if st and not lte_st:
                            lte_st = st

                nr_power = self._trim_cell_power(nr_power)
                lte_power = self._trim_cell_power(lte_power)
                if nr_power and nr_st in STATUS_TEXT:
                    nr_power = f"{nr_power}({STATUS_TEXT[nr_st]})"
                if lte_power and lte_st in STATUS_TEXT:
                    lte_power = f"{lte_power}({STATUS_TEXT[lte_st]})"

                params = rec.get('parameters', {}) if isinstance(rec, dict) else {}
                if not isinstance(params, dict):
                    params = {}
                # 多 band 扫描的 case(有 band_list/nr_band_list 且多于1个 band): NR/LTE Cell Power
                # 只会取到其中某一个 band 的值,当作整条记录的代表会误导客户 → 留空(宁可不显示也不显示错的)。
                # channel switch 顶层 parameters 可能为空, 回退到 loop 结果里的 parameters 判断。
                chk_params = params
                if not (chk_params.get('nr_band_list') or chk_params.get('band_list')):
                    rp = res.get('parameters') if isinstance(res, dict) else None
                    if isinstance(rp, dict):
                        chk_params = rp
                _bl = chk_params.get('nr_band_list') or chk_params.get('band_list')
                if isinstance(_bl, (list, tuple)) and len(_bl) > 1:
                    nr_power = ""
                    lte_power = ""
                display_name = self._display_config_name(name, params)
                rows.append({
                    'name': display_name,
                    'cfg_key': cfg_key,
                    'loop': loop_iter,
                    'rate': rate,
                    'metrics': metrics,
                    'nr_power': nr_power,
                    'lte_power': lte_power,
                    'params': params,
                })

        if not rows:
            empty = QLabel("（该执行记录暂无可汇总的数据）")
            empty.setStyleSheet("color:#888; padding:20px;")
            self._report_content_layout.addWidget(empty)
            return

        # 原始跑失败导致指标空缺的, 用重测(fail_try)里测到的值回填
        for row in rows:
            rm = retry_metrics.get(row['cfg_key'])
            if not rm:
                continue
            for t, (_rnd, val) in rm.items():
                if row['metrics'].get(t) is None:
                    row['metrics'][t] = val

        max_round = 0
        for rmap in retry_by_round.values():
            if rmap:
                max_round = max(max_round, max(rmap.keys()))
        has_retry = max_round > 0
        has_nr_power = any(row.get('nr_power') for row in rows)
        has_lte_power = any(row.get('lte_power') for row in rows)

        PRIORITY_METRICS = ["MeasuredValue", "LowLimit", "UpLimit", "Modulation"]
        metric_titles = ([t for t in PRIORITY_METRICS if t in metric_titles]
                         + [t for t in metric_titles if t not in PRIORITY_METRICS])

        PARAM_SKIP = {'case_dir', 'script', 'nr_slots', 'lte_slots',
                      'resource_allocation_type', 'power_class', 'ue_cap'}
        PARAM_LABELS = {
            'lineLoss1': '线损1(dBm)', 'lineLoss3': '线损3(dBm)',
            'nr_band': 'nr_band', 'lte_band': 'lte_band',
            'nr_bw': 'nr_bw(MHz)', 'lte_bw': 'lte_bw(MHz)',
            'scs': 'scs(kHz)', 'range': 'range',
            'power_class': 'power_class', 'rb_mode': 'rb_mode',
            'waveform': 'waveform', 'TD': 'TD(s)',
        }
        param_keys = []
        for row in rows:
            for k, v in (row.get('params') or {}).items():
                if k in PARAM_SKIP or k in param_keys:
                    continue
                if k == 'lineLoss1':   # 单独放到通过率后面, 不跟其它参数挤在末尾
                    continue
                if isinstance(v, (dict, list)):
                    continue
                param_keys.append(k)

        # 线损单独成列, 放在通过率/fail_try 之后(醒目位置)
        has_lineloss = any((row.get('params') or {}).get('lineLoss1') is not None for row in rows)

        headers = ["配置名称", "循环", "通过率(%)"]
        if has_retry:
            headers += [f"fail_try_{k}通过率(%)" for k in range(1, max_round + 1)]
        if has_lineloss:
            headers += ["线损1(dBm)"]
        if has_nr_power:
            headers += ["NR Cell Power(dBm)"]
        if has_lte_power:
            headers += ["LTE Cell Power(dBm)"]
        headers += metric_titles
        headers += [PARAM_LABELS.get(k, k) for k in param_keys]
        table = CopyableTable()
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setRowCount(len(rows))
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        table.setSelectionBehavior(QAbstractItemView.SelectItems)
        table.setAlternatingRowColors(True)
        _hdr = table.horizontalHeader()
        _hdr.setSectionsMovable(True)
        _hdr.sectionMoved.connect(lambda *_: self._save_summary_col_order(table))
        _tf = table.font()
        _tf.setPointSize(10)
        table.setFont(_tf)
        table.setStyleSheet(
            "QTableWidget { background:#fff; gridline-color:#e6ecf3; alternate-background-color:#f4f8fc; }"
            "QHeaderView::section { background:#2c3e50; color:white; font-weight:bold; padding:5px; }"
        )
        for r, row in enumerate(rows):
            table.setItem(r, 0, QTableWidgetItem(row['name']))
            table.setItem(r, 1, QTableWidgetItem(f"循环{row['loop']}"))
            rate = row['rate']
            rate_item = QTableWidgetItem("" if rate is None else f"{rate:.2f}")
            if rate is not None:
                if rate >= 100:
                    rate_item.setBackground(QColor(144, 238, 144))
                elif rate >= 90:
                    rate_item.setBackground(QColor(255, 255, 153))
                else:
                    rate_item.setBackground(QColor(255, 182, 193))
            table.setItem(r, 2, rate_item)
            col = 3
            if has_retry:
                rmap = retry_by_round.get(row['cfg_key'], {})
                for k in range(1, max_round + 1):
                    v = rmap.get(k)
                    rt_item = QTableWidgetItem("" if v is None else f"{v:.2f}")
                    if v is not None:
                        if v >= 100:
                            rt_item.setBackground(QColor(144, 238, 144))
                        elif v >= 90:
                            rt_item.setBackground(QColor(255, 255, 153))
                        else:
                            rt_item.setBackground(QColor(255, 182, 193))
                    table.setItem(r, col, rt_item)
                    col += 1
            if has_lineloss:
                ll = (row.get('params') or {}).get('lineLoss1')
                table.setItem(r, col, QTableWidgetItem("" if ll is None else str(ll)))
                col += 1
            if has_nr_power:
                table.setItem(r, col, QTableWidgetItem(row.get('nr_power', '')))
                col += 1
            if has_lte_power:
                table.setItem(r, col, QTableWidgetItem(row.get('lte_power', '')))
                col += 1
            for c, t in enumerate(metric_titles):
                v = row['metrics'].get(t)
                if v is None:
                    txt = ""
                elif isinstance(v, float):
                    txt = f"{v:.3f}"
                else:
                    txt = str(v)
                table.setItem(r, col + c, QTableWidgetItem(txt))
            pcol = col + len(metric_titles)
            rparams = row.get('params') or {}
            for pc, k in enumerate(param_keys):
                v = rparams.get(k)
                txt = "" if v is None else str(v)
                table.setItem(r, pcol + pc, QTableWidgetItem(txt))
        table.resizeColumnsToContents()
        table.horizontalHeader().setStretchLastSection(True)
        table.setMinimumHeight(min(600, 40 + 30 * len(rows)))
        self._restore_summary_col_order(table)

        # 列头点击下拉筛选: 「通过率(%)」和「fail_try_N通过率(%)」列点表头弹菜单(全部/只看100%/只看<100%)。
        # 表头加 ▼ 提示可筛选; 各列筛选条件用 AND 组合。
        import re as _re
        def _is_rate_col(title):
            return title == "通过率(%)" or bool(_re.match(r'fail_try_\d+通过率', title))
        table._rate_filter_cols = set()
        _fm = table.fontMetrics()
        for c in range(table.columnCount()):
            it = table.horizontalHeaderItem(c)
            if it and _is_rate_col(it.text()):
                table._rate_filter_cols.add(c)
                it.setText(it.text() + " ▼")
                # 表头加 ▼ 后变长, 且这两列是客户重点看的 → 保证宽度足够完整显示表头
                need = _fm.horizontalAdvance(it.text()) + 36
                if table.columnWidth(c) < need:
                    table.setColumnWidth(c, need)
        table._rate_filters = {}   # {col: 'pass'|'fail'|None}
        hdr = table.horizontalHeader()
        hdr.setSectionsClickable(True)
        hdr.sectionClicked.connect(
            lambda col, t=table: self._on_summary_header_clicked(t, col))
        self._report_content_layout.addWidget(table)

    def _on_summary_header_clicked(self, table, col):
        if col not in getattr(table, '_rate_filter_cols', set()):
            return
        from PyQt5.QtGui import QCursor
        menu = QMenu(table)
        cur = table._rate_filters.get(col)
        a_all = menu.addAction("✓ 全部" if cur is None else "全部")
        a_pass = menu.addAction("✓ 只看 100%" if cur == 'pass' else "只看 100%")
        a_fail = menu.addAction("✓ 只看 <100%" if cur == 'fail' else "只看 <100%")
        act = menu.exec_(QCursor.pos())
        if act is None:
            return
        if act == a_all:
            table._rate_filters.pop(col, None)
        elif act == a_pass:
            table._rate_filters[col] = 'pass'
        elif act == a_fail:
            table._rate_filters[col] = 'fail'
        self._apply_summary_col_filters(table)

    def _apply_summary_col_filters(self, table):
        filters = getattr(table, '_rate_filters', {})
        for r in range(table.rowCount()):
            hide = False
            for col, mode in filters.items():
                it = table.item(r, col)
                txt = (it.text().strip() if it else "")
                try:
                    val = float(txt)
                except (TypeError, ValueError):
                    val = None
                if mode == 'pass':
                    if not (val is not None and val >= 100):
                        hide = True; break
                elif mode == 'fail':
                    if not (val is not None and val < 100):
                        hide = True; break
            table.setRowHidden(r, hide)

    def _summary_col_order_path(self):
        return os.path.join(project_root(), "summary_col_order.json")

    def _save_summary_col_order(self, table):
        try:
            hdr = table.horizontalHeader()
            order = []
            for visual in range(table.columnCount()):
                logical = hdr.logicalIndex(visual)
                it = table.horizontalHeaderItem(logical)
                order.append(it.text() if it else "")
            with open(self._summary_col_order_path(), "w", encoding="utf-8") as f:
                json.dump(order, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存汇总列顺序失败: {e}")

    def _restore_summary_col_order(self, table):
        path = self._summary_col_order_path()
        try:
            import re
            hdr = table.horizontalHeader()
            titles = [table.horizontalHeaderItem(c).text() if table.horizontalHeaderItem(c) else ""
                      for c in range(table.columnCount())]

            def _is_fail_try(t):
                return bool(re.match(r'fail_try_\d+通过率', t))

            # 读取历史保存顺序(可能没有该文件/没有 fail_try 列)
            saved = []
            if os.path.exists(path):
                with open(path, encoding="utf-8") as f:
                    saved = json.load(f)
                if not isinstance(saved, list):
                    saved = []
            rank = {t: i for i, t in enumerate(saved)}

            # 前置块整体固定在最左: 配置名称 → 循环 → 通过率 → fail_try_1..N → 线损1,
            # 不受历史保存顺序影响(否则改名/新列会被 rank 甩到末尾); 其余列按保存顺序排。
            def _idx_of(title):
                return titles.index(title) if title in titles else None

            front = []
            for t in ("配置名称", "循环", "通过率(%)"):
                c = _idx_of(t)
                if c is not None:
                    front.append(c)
            fail_cols = [c for c in range(len(titles)) if _is_fail_try(titles[c])]
            fail_cols.sort(key=lambda c: int(re.match(r'fail_try_(\d+)', titles[c]).group(1)))
            front.extend(fail_cols)
            c = _idx_of("线损1(dBm)")
            if c is not None:
                front.append(c)

            pinned = set(front)
            other_cols = [c for c in range(len(titles)) if c not in pinned]
            other_cols.sort(key=lambda c: (rank.get(titles[c], len(saved) + c), c))
            target = front + other_cols

            hdr.blockSignals(True)
            for visual_pos, logical in enumerate(target):
                cur = hdr.visualIndex(logical)
                if cur != visual_pos:
                    hdr.moveSection(cur, visual_pos)
            hdr.blockSignals(False)
        except Exception as e:
            print(f"恢复汇总列顺序失败: {e}")

    def _config_key_from_name(self, case_dir_name):
        parts = case_dir_name.split("_")

        if (len(parts) > 3 and parts[0].isdigit() and len(parts[0]) <= 4
                and len(parts[1]) == 8 and parts[1].isdigit() and parts[2].isdigit()):
            return "_".join(parts[3:])

        if len(parts) > 2 and parts[0].isdigit() and parts[1].isdigit():
            return "_".join(parts[2:])
        return case_dir_name

    def _config_key_from_rec(self, rec, fallback_name):
        # cfg_key 用途: 区分不同 case(不同→不同 key) + 匹配同一 case 的重测目录(相同参数→相同 key)。
        # 目录名会被截断到 70 字符、且原始带 seq_ 前缀/重测不带,不可靠 → 改用 JSON 里的完整参数。
        # 关键: 必须用"全部参数的规范序列化",否则 channel switch 这类(nr_band_list/rb_mode/bw/scs
        # 都是 list/dict)只按标量参数会退化成同 key,把两个不同 case 误判成互为重测。
        base = self._config_key_from_name(fallback_name)
        params = None
        if isinstance(rec, dict):
            p = rec.get('parameters')
            if isinstance(p, dict):
                params = p
            else:
                for lr in rec.get('loop_results', []):
                    res = lr.get('result', {}) if isinstance(lr, dict) else {}
                    pp = res.get('parameters')
                    if isinstance(pp, dict):
                        params = pp
                        break
        if not isinstance(params, dict):
            return base
        import re, json
        m = re.search(r'_(?:pc|nr_band|lte_band|nr_bw|lte_bw|scs|range)', base)
        script_base = base[:m.start()] if m else base
        # 排除易变/非配置项, 其余全部参与 key(list/dict 一并规范序列化)
        volatile = {'case_dir', 'script', 'script_name', 'script_file', 'file'}
        key_params = {k: params[k] for k in params if k not in volatile}
        # range 运行时可能被追加 ARFCN(如 "LOW" → "LOW:424000"), 归一化只取前缀,
        # 否则同一 case 的原始目录与重测目录 key 对不上, 重测通过率(fail_try)挂不到原行。
        rv = key_params.get('range')
        if isinstance(rv, str) and ':' in rv:
            key_params['range'] = rv.split(':')[0]
        try:
            params_sig = json.dumps(key_params, sort_keys=True, ensure_ascii=False)
        except Exception:
            params_sig = str(sorted(key_params.items(), key=lambda kv: kv[0]))
        return f"{script_base}|{params_sig}"

    def _display_config_name(self, case_dir_name, params):
        base = self._config_key_from_name(case_dir_name)
        if base.lower().startswith("yc"):
            import re
            m = re.search(r'_(?:nr_band|lte_band|nr_bw|lte_bw|scs|range)', base)
            if m:
                base = base[:m.start()]
            if params and 'power_class' in params and not re.search(r'_pc(?:\d|none)', base, re.I):
                pc = params.get('power_class')
                base += f"_pc{pc}" if pc is not None else "_pcnone"
        return base

    @staticmethod
    def _trim_cell_power(text):
        if not text:
            return text
        import re
        m = re.match(r'^\s*(?:nr|lte)_cell_power:\s*([-+]?[\d.]+)', str(text), re.I)
        return m.group(1) if m else text

    # ================= Config UE Capability 页面 =================

    def _build_ue_capability_page(self):
        import ue_capability as ucap
        self._ucap = ucap
        self._ucap_case_opts = self._load_ucap_case_opts()

        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(8)

        top = QHBoxLayout()
        top.setSpacing(10)
        lbl_sys = QLabel("System:")
        lbl_sys.setStyleSheet("font-weight:bold; color:#1f3a5f;")
        self.ucap_system_combo = QComboBox()
        self.ucap_system_combo.addItems(["NR"])
        self.ucap_system_combo.setFixedWidth(120)
        lbl_cap = QLabel("Capability:")
        lbl_cap.setStyleSheet("font-weight:bold; color:#1f3a5f;")
        self.ucap_cap_combo = QComboBox()
        self.ucap_cap_combo.addItems(["Bandwidth"])
        self.ucap_cap_combo.setFixedWidth(150)
        top.addWidget(lbl_sys); top.addWidget(self.ucap_system_combo)
        top.addSpacing(20)
        top.addWidget(lbl_cap); top.addWidget(self.ucap_cap_combo)
        top.addStretch()
        save_matrix_btn = QPushButton("💾 保存能力表")
        save_matrix_btn.setToolTip("保存矩阵勾选状态，软件重启后保持")
        save_matrix_btn.setStyleSheet(
            "QPushButton { background:#2e8b57; color:white; font-weight:bold; padding:6px 16px; border-radius:6px; }"
            "QPushButton:hover { background:#256f46; }")
        save_matrix_btn.clicked.connect(self._ucap_save_matrix)
        top.addWidget(save_matrix_btn)
        layout.addLayout(top)

        splitter = QSplitter(Qt.Horizontal)

        saved_unchecked = self._ucap_load_matrix_unchecked()
        self.ucap_matrix = QTableWidget()
        bws = ucap.ALL_BWS
        self.ucap_matrix.setColumnCount(2 + len(bws))
        self.ucap_matrix.setHorizontalHeaderLabels(
            ["Band", "SCS"] + [f"{b}MHz" for b in bws])
        rows = [(band, scs) for band in ucap.ALL_BANDS for scs in ucap.ALL_SCS]
        self.ucap_matrix.setRowCount(len(rows))
        self._ucap_rows = rows
        for r, (band, scs) in enumerate(rows):
            band_item = QTableWidgetItem(f"n{band}")
            band_item.setFlags(Qt.ItemIsEnabled)
            band_item.setTextAlignment(Qt.AlignCenter)
            self.ucap_matrix.setItem(r, 0, band_item)
            scs_item = QTableWidgetItem(str(scs))
            scs_item.setFlags(Qt.ItemIsEnabled)
            scs_item.setTextAlignment(Qt.AlignCenter)
            self.ucap_matrix.setItem(r, 1, scs_item)
            for c, bw in enumerate(bws, start=2):
                it = QTableWidgetItem()
                it.setTextAlignment(Qt.AlignCenter)
                if ucap.is_supported(band, scs, bw):
                    it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
                    unchecked = f"{band},{scs},{bw}" in saved_unchecked
                    it.setCheckState(Qt.Unchecked if unchecked else Qt.Checked)
                else:
                    it.setFlags(Qt.NoItemFlags)
                    it.setBackground(QColor(210, 210, 210))
                self.ucap_matrix.setItem(r, c, it)
        self.ucap_matrix.verticalHeader().setVisible(False)
        self.ucap_matrix.setColumnWidth(0, 58)
        self.ucap_matrix.setColumnWidth(1, 46)
        for c in range(2, 2 + len(bws)):
            self.ucap_matrix.setColumnWidth(c, 62)
        self.ucap_matrix.setEditTriggers(QTableWidget.NoEditTriggers)
        self.ucap_matrix.setAlternatingRowColors(True)
        self.ucap_matrix.setStyleSheet(
            "QTableWidget { background:#ffffff; gridline-color:#e6ecf3; alternate-background-color:#f6f9fc; }"
            "QHeaderView::section { background:#2c3e50; color:white; font-weight:bold; padding:4px; }")
        self.ucap_matrix.itemChanged.connect(lambda *_: self._ucap_update_preview())
        splitter.addWidget(self.ucap_matrix)

        right = QWidget()
        right_v = QVBoxLayout(right)
        right_v.setContentsMargins(4, 0, 0, 0)
        right_v.setSpacing(8)

        qc_group = QGroupBox("⚡ Quick Configuration (实际要跑的 band)")
        qc_v = QVBoxLayout(qc_group)
        qc_hint = QLabel("勾选要测的频段；用例 = 所选 band 在矩阵里勾中的 bw/scs 组合")
        qc_hint.setStyleSheet("color:#7f8c8d;")
        qc_hint.setWordWrap(True)
        qc_v.addWidget(qc_hint)
        qc_scroll = QScrollArea()
        qc_scroll.setWidgetResizable(True)
        qc_scroll.setStyleSheet("QScrollArea { border:1px solid #c3d0de; border-radius:6px; background:#ffffff; }")
        qc_inner = QWidget()
        qc_grid = QGridLayout(qc_inner)
        qc_grid.setContentsMargins(8, 8, 8, 8)
        qc_grid.setSpacing(4)
        self.ucap_band_checks = {}
        n_cols = 4
        for i, band in enumerate(ucap.ALL_BANDS):
            cb = QCheckBox(f"n{band}")
            cb.toggled.connect(lambda *_: self._ucap_update_preview())
            self.ucap_band_checks[band] = cb
            qc_grid.addWidget(cb, i // n_cols, i % n_cols)
        qc_scroll.setWidget(qc_inner)
        qc_v.addWidget(qc_scroll)
        right_v.addWidget(qc_group, 5)

        case_group = QGroupBox("🧪 Test Item List (yc 开头 case, 勾选弹出 RB/Range 配置)")
        cg_v = QVBoxLayout(case_group)
        self.ucap_case_list = QListWidget()
        self.ucap_case_list.setStyleSheet(
            "QListWidget { background:#ffffff; border:1px solid #c3d0de; border-radius:6px; }"
            "QListWidget::item { padding:4px 6px; }")
        for name in self._ucap_scan_yc_cases():
            li = QListWidgetItem(name)
            li.setFlags(li.flags() | Qt.ItemIsUserCheckable)
            li.setCheckState(Qt.Unchecked)
            self.ucap_case_list.addItem(li)
        self.ucap_case_list.itemChanged.connect(self._ucap_on_case_toggled)
        cg_v.addWidget(self.ucap_case_list)
        right_v.addWidget(case_group, 4)

        order_group = QGroupBox("🔢 执行顺序 (case / band 均可上下移)")
        og_v = QVBoxLayout(order_group)
        self.ucap_order_tree = QTreeWidget()
        self.ucap_order_tree.setHeaderHidden(True)
        self.ucap_order_tree.setStyleSheet(
            "QTreeWidget { background:#ffffff; border:1px solid #c3d0de; border-radius:6px; }"
            "QTreeWidget::item { padding:3px 4px; }")
        og_v.addWidget(self.ucap_order_tree)
        order_btns = QHBoxLayout()
        up_btn = QPushButton("⬆ 上移")
        up_btn.clicked.connect(lambda: self._ucap_order_move(-1))
        down_btn = QPushButton("⬇ 下移")
        down_btn.clicked.connect(lambda: self._ucap_order_move(1))
        order_btns.addWidget(up_btn)
        order_btns.addWidget(down_btn)
        order_btns.addStretch()
        og_v.addLayout(order_btns)
        right_v.addWidget(order_group, 5)

        # 线损: 控制生成用例的 lineLoss1(所有从本页生成的用例统一用这个线损)
        ll_row = QHBoxLayout()
        ll_lbl = QLabel("📉 线损 lineLoss1(dBm):")
        ll_lbl.setStyleSheet("font-weight:bold; color:#2471a3;")
        self.ucap_lineloss_spin = QDoubleSpinBox()
        self.ucap_lineloss_spin.setRange(0.0, 200.0)
        self.ucap_lineloss_spin.setDecimals(2)
        self.ucap_lineloss_spin.setSingleStep(1.0)
        self.ucap_lineloss_spin.setFixedWidth(110)
        try:
            _ll0 = float(self._ucap_case_opts.get("__lineLoss1__", 25.0))
        except Exception:
            _ll0 = 25.0
        self.ucap_lineloss_spin.setValue(_ll0)
        self.ucap_lineloss_spin.setToolTip("从本页生成的所有用例都会带上这个 lineLoss1 线损值")
        self.ucap_lineloss_spin.valueChanged.connect(
            lambda *_: (self._ucap_save_lineloss(), self._ucap_update_preview()))
        ll_row.addWidget(ll_lbl)
        ll_row.addWidget(self.ucap_lineloss_spin)
        ll_row.addStretch()
        right_v.addLayout(ll_row)

        splitter.addWidget(right)
        splitter.setSizes([880, 380])
        layout.addWidget(splitter, 1)

        bottom = QHBoxLayout()
        self.ucap_preview_label = QLabel("将生成 0 条用例")
        self.ucap_preview_label.setStyleSheet("color:#2471a3; font-weight:bold; padding:4px;")
        bottom.addWidget(self.ucap_preview_label)
        bottom.addStretch()
        add_btn = QPushButton("➕ 加入到动态Case List")
        add_btn.setStyleSheet(
            "QPushButton { background:#2e8b57; color:white; font-weight:bold; padding:8px 20px; border-radius:6px; }"
            "QPushButton:hover { background:#256f46; }")
        add_btn.clicked.connect(self._ucap_add_to_dyn)
        bottom.addWidget(add_btn)
        layout.addLayout(bottom)
        return page

    def _ucap_matrix_store_path(self):
        return os.path.join(project_root(), "ue_capability_matrix.json")

    def _ucap_load_matrix_unchecked(self):
        try:
            with open(self._ucap_matrix_store_path(), encoding="utf-8") as f:
                d = json.load(f)
            return set(d.get("unchecked", []))
        except Exception:
            return set()

    def _ucap_save_matrix(self):
        unchecked = []
        bws = self._ucap.ALL_BWS
        for r, (band, scs) in enumerate(self._ucap_rows):
            for c, bw in enumerate(bws, start=2):
                it = self.ucap_matrix.item(r, c)
                if it and (it.flags() & Qt.ItemIsUserCheckable) and it.checkState() != Qt.Checked:
                    unchecked.append(f"{band},{scs},{bw}")
        try:
            with open(self._ucap_matrix_store_path(), "w", encoding="utf-8") as f:
                json.dump({"unchecked": unchecked}, f, ensure_ascii=False, indent=2)
            self.status_label.setText(f"能力表已保存 (取消勾选 {len(unchecked)} 项)")
        except Exception as e:
            QMessageBox.warning(self, "保存失败", str(e))

    def _ucap_scan_yc_cases(self):
        try:
            case_dir = souren_config.get_case_directory('yc1100')
            return sorted(os.path.splitext(f)[0] for f in os.listdir(case_dir)
                          if f.startswith('yc') and f.endswith('.py'))
        except Exception as e:
            print(f"扫描 yc case 失败: {e}")
            return []

    def _ucap_on_case_toggled(self, item):
        if item.checkState() != Qt.Checked:
            self._ucap_update_preview()
            return
        name = item.text()
        prev = self._ucap_case_opts.get(name, {})
        dlg = QDialog(self)
        dlg.setWindowTitle(f"配置 {name}")
        dlg.setMinimumWidth(420)
        v = QVBoxLayout(dlg)

        def _mk_orderable_group(title, all_options, prev_checked):
            group = QGroupBox(title)
            gv = QVBoxLayout(group)
            lw = QListWidget()
            ordered = [t for t in prev_checked if t in all_options]
            ordered += [t for t in all_options if t not in ordered]
            for text in ordered:
                li = QListWidgetItem(text)
                li.setFlags(li.flags() | Qt.ItemIsUserCheckable)
                li.setCheckState(Qt.Checked if text in prev_checked else Qt.Unchecked)
                lw.addItem(li)
            gv.addWidget(lw)
            btn_row = QHBoxLayout()

            def _move(delta):
                row = lw.currentRow()
                new_row = row + delta
                if row < 0 or not (0 <= new_row < lw.count()):
                    return
                it = lw.takeItem(row)
                lw.insertItem(new_row, it)
                lw.setCurrentRow(new_row)

            up = QPushButton("⬆ 上移")
            up.clicked.connect(lambda: _move(-1))
            down = QPushButton("⬇ 下移")
            down.clicked.connect(lambda: _move(1))
            btn_row.addWidget(up)
            btn_row.addWidget(down)
            btn_row.addStretch()
            gv.addLayout(btn_row)
            return group, lw

        wf_row = QHBoxLayout()
        wf_row.addWidget(QLabel("🌊 Waveform:"))
        wf_combo = QComboBox()
        wf_combo.addItems(self._ucap.WAVEFORM_OPTIONS)
        wf_combo.setCurrentText(prev.get("waveform", "DFTS"))
        wf_combo.setFixedWidth(120)
        wf_row.addWidget(wf_combo)
        wf_row.addStretch()
        v.addLayout(wf_row)

        # RMC 配置: 控制只配 UL / 只配 DL / 两个方向都配(满配 slot 由 case 默认表+FDD 过滤决定)
        rmc_row = QHBoxLayout()
        rmc_row.addWidget(QLabel("🧩 RMC 配置:"))
        rmc_combo = QComboBox()
        rmc_combo.addItems(["UL和DL都配置", "仅UL配置", "仅DL配置"])
        rmc_combo.setCurrentText(prev.get("rmc", "UL和DL都配置"))
        rmc_combo.setFixedWidth(140)
        rmc_row.addWidget(rmc_combo)
        rmc_row.addStretch()
        v.addLayout(rmc_row)

        # MCS 配置: 覆盖 slot 表的 MCS1, 默认 4
        mcs_row = QHBoxLayout()
        mcs_row.addWidget(QLabel("🎚 MCS:"))
        mcs_spin = QSpinBox()
        mcs_spin.setRange(0, 27)
        mcs_spin.setValue(int(prev.get("mcs", 4)))
        mcs_spin.setFixedWidth(120)
        mcs_row.addWidget(mcs_spin)
        mcs_row.addStretch()
        v.addLayout(mcs_row)

        rb_group, rb_list = _mk_orderable_group(
            "📶 NR UL RB Mode (可多选, 顺序=执行顺序)",
            self._ucap.RB_MODE_OPTIONS, prev.get("rb", ["Outer_Full"]))
        v.addWidget(rb_group)

        rng_group, rng_list = _mk_orderable_group(
            "📡 Channel Range (可多选, 顺序=执行顺序)",
            self._ucap.RANGE_OPTIONS, prev.get("range", ["LOW", "MID", "HIGH"]))
        v.addWidget(rng_group)

        # 仅DL配置无上行, RB Mode 不可选: 强制勾选“无”并禁用其余(相当只跑一条);
        # 其它 RMC 模式下“无”不可选。
        _RB_NONE = self._ucap.RB_MODE_NONE

        def _apply_rmc_to_rb():
            dl_only = (rmc_combo.currentText() == "仅DL配置")
            rb_list.blockSignals(True)
            for i in range(rb_list.count()):
                it = rb_list.item(i)
                is_none = (it.text() == _RB_NONE)
                if dl_only:
                    it.setCheckState(Qt.Checked if is_none else Qt.Unchecked)
                    it.setFlags(it.flags() & ~Qt.ItemIsEnabled)
                elif is_none:
                    it.setCheckState(Qt.Unchecked)
                    it.setFlags(it.flags() & ~Qt.ItemIsEnabled)
                else:
                    it.setFlags(it.flags() | Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            rb_list.blockSignals(False)

        rmc_combo.currentTextChanged.connect(lambda _=None: _apply_rmc_to_rb())
        _apply_rmc_to_rb()


        pc_combo = None
        if name.startswith(self._ucap.POWER_CLASS_CASE_PREFIX):
            pc_row = QHBoxLayout()
            pc_row.addWidget(QLabel("⚡ Power Class:"))
            pc_combo = QComboBox()
            pc_combo.addItems(self._ucap.POWER_CLASS_OPTIONS)
            pc_combo.setCurrentText(str(prev.get("power_class", "None")))
            pc_combo.setFixedWidth(120)
            pc_row.addWidget(pc_combo)
            pc_row.addStretch()
            v.addLayout(pc_row)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        v.addWidget(btns)

        def _checked(lw):
            return [lw.item(i).text() for i in range(lw.count())
                    if lw.item(i).checkState() == Qt.Checked]

        if dlg.exec_() == QDialog.Accepted:
            rb = _checked(rb_list) or ["Outer_Full"]
            rng = _checked(rng_list) or ["LOW"]
            opts = {"rb": rb, "range": rng, "waveform": wf_combo.currentText(),
                    "rmc": rmc_combo.currentText(), "mcs": mcs_spin.value()}
            if pc_combo is not None:
                opts["power_class"] = pc_combo.currentText()
            self._ucap_case_opts[name] = opts
            self._save_ucap_case_opts()
        else:
            self.ucap_case_list.blockSignals(True)
            item.setCheckState(Qt.Unchecked)
            self.ucap_case_list.blockSignals(False)
        self._ucap_update_preview()

    def _ucap_case_opts_path(self):
        return os.path.join(project_root(), "ucap_case_opts.json")

    def _load_ucap_case_opts(self):
        path = self._ucap_case_opts_path()
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
            except Exception:
                pass
        return {}

    def _save_ucap_case_opts(self):
        try:
            with open(self._ucap_case_opts_path(), "w", encoding="utf-8") as f:
                json.dump(self._ucap_case_opts, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存 UE Capability case 配置失败: {e}")

    def _ucap_save_lineloss(self):
        # 线损单独存在 ucap_case_opts 的保留 key(不是脚本名, 不影响用例生成的遍历)
        try:
            self._ucap_case_opts["__lineLoss1__"] = round(float(self.ucap_lineloss_spin.value()), 2)
            self._save_ucap_case_opts()
        except Exception:
            pass

    def _ucap_lineloss(self):
        try:
            return round(float(self.ucap_lineloss_spin.value()), 2)
        except Exception:
            return None

    def _ucap_selected_bands(self):
        return [b for b, cb in self.ucap_band_checks.items() if cb.isChecked()]
    def _ucap_matrix_combos(self, bands=None):
        combos = []
        bws = self._ucap.ALL_BWS
        band_filter = set(bands) if bands is not None else None
        for r, (band, scs) in enumerate(self._ucap_rows):
            if band_filter is not None and band not in band_filter:
                continue
            for c, bw in enumerate(bws, start=2):
                it = self.ucap_matrix.item(r, c)
                if it and (it.flags() & Qt.ItemIsUserCheckable) and it.checkState() == Qt.Checked:
                    combos.append((band, scs, bw))
        return combos

    def _ucap_checked_cases(self):
        return [self.ucap_case_list.item(i).text() for i in range(self.ucap_case_list.count())
                if self.ucap_case_list.item(i).checkState() == Qt.Checked]


    def _ucap_order_sync(self):
        tree = self.ucap_order_tree
        checked_cases = self._ucap_checked_cases()
        sel_bands = self._ucap_selected_bands()

        old_case_order = []
        old_band_order = {}
        for i in range(tree.topLevelItemCount()):
            top = tree.topLevelItem(i)
            cname = top.data(0, Qt.UserRole)
            old_case_order.append(cname)
            old_band_order[cname] = [top.child(j).data(0, Qt.UserRole)
                                     for j in range(top.childCount())]

        case_order = [c for c in old_case_order if c in checked_cases]
        case_order += [c for c in checked_cases if c not in case_order]

        tree.blockSignals(True)
        tree.clear()
        for cname in case_order:
            top = QTreeWidgetItem([f"🧪 {cname}"])
            top.setData(0, Qt.UserRole, cname)
            prev_bands = old_band_order.get(cname, [])
            band_order = [b for b in prev_bands if b in sel_bands]
            band_order += [b for b in sel_bands if b not in band_order]
            for b in band_order:
                child = QTreeWidgetItem([f"n{b}"])
                child.setData(0, Qt.UserRole, b)
                top.addChild(child)
            tree.addTopLevelItem(top)
            top.setExpanded(True)
        tree.blockSignals(False)

    def _ucap_order_move(self, delta):
        tree = self.ucap_order_tree
        item = tree.currentItem()
        if item is None:
            return
        parent = item.parent()
        if parent is None:
            idx = tree.indexOfTopLevelItem(item)
            new_idx = idx + delta
            if 0 <= new_idx < tree.topLevelItemCount():
                expanded = item.isExpanded()
                tree.takeTopLevelItem(idx)
                tree.insertTopLevelItem(new_idx, item)
                item.setExpanded(expanded)
                tree.setCurrentItem(item)
        else:
            idx = parent.indexOfChild(item)
            new_idx = idx + delta
            if 0 <= new_idx < parent.childCount():
                parent.takeChild(idx)
                parent.insertChild(new_idx, item)
                tree.setCurrentItem(item)

    def _ucap_order_from_tree(self):
        tree = self.ucap_order_tree
        out = []
        for i in range(tree.topLevelItemCount()):
            top = tree.topLevelItem(i)
            bands = [top.child(j).data(0, Qt.UserRole) for j in range(top.childCount())]
            out.append((top.data(0, Qt.UserRole), bands))
        if not out:
            sel = self._ucap_selected_bands()
            out = [(c, list(sel)) for c in self._ucap_checked_cases()]
        return out

    def _ucap_build_cases(self):
        cases = []
        lineloss = self._ucap_lineloss()
        for script, bands in self._ucap_order_from_tree():
            opts = self._ucap_case_opts.get(script, {"rb": ["Outer_Full"], "range": ["LOW"]})
            waveform = opts.get("waveform", "DFTS")
            pc_raw = opts.get("power_class")
            rmc = opts.get("rmc", "UL和DL都配置")
            mcs = opts.get("mcs", 4)
            for band in bands:
                for band2, scs, bw in self._ucap_matrix_combos([band]):
                    for rng in opts["range"]:
                        for rb in opts["rb"]:
                            case = {
                                "script": script, "nr_band": band2, "nr_bw": bw, "scs": scs,
                                "range": rng, "rb_mode": rb, "waveform": waveform,
                                "mcs": mcs,
                                "case_dir": "yc1100",
                            }
                            # RMC 配置: 仅UL/仅DL -> nr_slots 只带该方向(空值=满配默认);
                            # 都配置 -> 不传 nr_slots, 用 case 默认 DL+UL 满配
                            if rmc == "仅UL配置":
                                case["nr_slots"] = {"UL": {}}
                            elif rmc == "仅DL配置":
                                case["nr_slots"] = {"DL": {}}
                            if lineloss is not None:
                                case["lineLoss1"] = lineloss
                            if pc_raw is not None:
                                if str(pc_raw) == "None":
                                    case["power_class"] = None
                                else:
                                    v = float(pc_raw)
                                    case["power_class"] = int(v) if v == int(v) else v
                            cases.append(case)
        return cases

    def _ucap_update_preview(self):
        self._ucap_order_sync()
        cases = self._ucap_build_cases()
        n = len(cases)
        n_combo = len(self._ucap_matrix_combos(self._ucap_selected_bands()))
        self.ucap_preview_label.setText(
            f"将生成 {n} 条用例 (Quick Config {len(self._ucap_selected_bands())} 个 band, "
            f"矩阵命中 {n_combo} 组 bw/scs)")
        if hasattr(self, '_dyn_lists'):
            self._dyn_lists[self.DYN_DEFAULT_LIST] = cases
            self._dyn_save_store()
            if (hasattr(self, 'dyn_list_combo')
                    and self.dyn_list_combo.currentText() == self.DYN_DEFAULT_LIST):
                self._dyn_refresh_table()

    def _ucap_add_to_dyn(self):
        new_cases = self._ucap_build_cases()
        if not new_cases:
            QMessageBox.information(self, "无可加入用例",
                                    "请在 Quick Configuration 勾选频段，并勾选至少一个 case")
            return
        self._dyn_replace_cases(new_cases)
        self.page_selector.setCurrentIndex(2)   
        self.status_label.setText(
            f"『默认』动态Case List已更新为当前选择的 {len(new_cases)} 条用例(未改动已保存的列表)")

    # ================= 动态 Case List 页面 =================

    def _dyn_store_path(self):
        return os.path.join(project_root(), "dyn_case_lists.json")

    def _dyn_load_store(self):
        try:
            with open(self._dyn_store_path(), encoding="utf-8") as f:
                d = json.load(f)
            return d if isinstance(d, dict) else {}
        except Exception:
            return {}

    def _dyn_save_store(self):
        try:
            with open(self._dyn_store_path(), "w", encoding="utf-8") as f:
                json.dump(self._dyn_lists, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存动态Case List失败: {e}")

    def _build_dyn_case_page(self):
        self._dyn_lists = self._dyn_load_store()
        if self._dyn_lists.get(self.DYN_DEFAULT_LIST):
            print(f"🧹 启动清理: 动态Case List『{self.DYN_DEFAULT_LIST}』暂存表清空 "
                  f"{len(self._dyn_lists[self.DYN_DEFAULT_LIST])} 条上次会话的用例")
        self._dyn_lists[self.DYN_DEFAULT_LIST] = []
        self._dyn_save_store()

        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(8)

        title = QLabel("📋 动态Case List")
        title.setStyleSheet("font-size:16px; font-weight:bold; color:#1f3a5f;")
        layout.addWidget(title)

        bar = QHBoxLayout()
        bar.addWidget(QLabel("列表:"))
        self.dyn_list_combo = QComboBox()
        self.dyn_list_combo.setMinimumWidth(180)
        self.dyn_list_combo.addItems(list(self._dyn_lists.keys()))
        self.dyn_list_combo.currentTextChanged.connect(lambda *_: self._dyn_refresh_table())
        bar.addWidget(self.dyn_list_combo)
        add_list_btn = QPushButton("➕ add")
        add_list_btn.setToolTip("新建一个动态Case列表")
        add_list_btn.clicked.connect(self._dyn_add_list)
        del_list_btn = QPushButton("➖ del")
        del_list_btn.setToolTip("删除当前动态Case列表")
        del_list_btn.clicked.connect(self._dyn_del_list)
        rename_list_btn = QPushButton("✏ 重命名")
        rename_list_btn.setToolTip("重命名当前列表")
        rename_list_btn.clicked.connect(self._dyn_rename_list)
        save_list_btn = QPushButton("💾 保存")
        save_list_btn.setToolTip("保存所有动态Case列表到本地")
        save_list_btn.setStyleSheet("QPushButton { background:#2e8b57; color:white; font-weight:bold; }")
        save_list_btn.clicked.connect(self._dyn_save_clicked)
        bar.addWidget(add_list_btn)
        bar.addWidget(del_list_btn)
        bar.addWidget(rename_list_btn)
        bar.addWidget(save_list_btn)
        bar.addStretch()
        layout.addLayout(bar)

        self.dyn_table = QTableWidget()
        self.dyn_table.setColumnCount(8)
        self.dyn_table.setHorizontalHeaderLabels(
            ["选择", "脚本名", "band", "bw", "scs", "range", "rb_mode", "参数(JSON)"])
        self.dyn_table.setColumnWidth(0, 46)
        self.dyn_table.setColumnWidth(1, 260)
        for c, w in ((2, 60), (3, 60), (4, 60), (5, 70), (6, 130)):
            self.dyn_table.setColumnWidth(c, w)
        self.dyn_table.horizontalHeader().setStretchLastSection(True)
        self.dyn_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.dyn_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.dyn_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.dyn_table.setAlternatingRowColors(True)
        self.dyn_table.setStyleSheet(
            "QTableWidget { background:#ffffff; gridline-color:#e6ecf3; alternate-background-color:#f6f9fc; }"
            "QHeaderView::section { background:#2c3e50; color:white; font-weight:bold; padding:5px; }")
        layout.addWidget(self.dyn_table, 1)

        bottom = QHBoxLayout()
        self.dyn_count_label = QLabel("")
        self.dyn_count_label.setStyleSheet("color:#2471a3; font-weight:bold;")
        bottom.addWidget(self.dyn_count_label)
        bottom.addStretch()
        check_all_btn = QPushButton("☑ 全选")
        check_all_btn.clicked.connect(lambda: self._dyn_set_all_checked(True))
        uncheck_all_btn = QPushButton("☐ 全不选")
        uncheck_all_btn.clicked.connect(lambda: self._dyn_set_all_checked(False))
        move_up_btn = QPushButton("⬆ 上移")
        move_up_btn.clicked.connect(lambda: self._dyn_move_rows(-1))
        move_down_btn = QPushButton("⬇ 下移")
        move_down_btn.clicked.connect(lambda: self._dyn_move_rows(1))
        del_row_btn = QPushButton("➖ 删除选中行")
        del_row_btn.clicked.connect(self._dyn_del_rows)
        clear_btn = QPushButton("🧹 清空当前列表")
        clear_btn.clicked.connect(self._dyn_clear_list)
        bottom.addWidget(check_all_btn)
        bottom.addWidget(uncheck_all_btn)
        bottom.addWidget(move_up_btn)
        bottom.addWidget(move_down_btn)
        bottom.addWidget(del_row_btn)
        bottom.addWidget(clear_btn)
        bottom.addSpacing(16)
        bottom.addWidget(QLabel("目标:"))
        self.dyn_target_combo = QComboBox()
        self.dyn_target_combo.addItems(["NR", "LTE"])
        self.dyn_target_combo.setFixedWidth(80)
        bottom.addWidget(self.dyn_target_combo)
        to_list_btn = QPushButton("🚀 加入测试列表")
        to_list_btn.setToolTip("把勾选(首列)的用例加入自动化测试页的 NR/LTE 列表；重启软件后自动清除")
        to_list_btn.setStyleSheet(
            "QPushButton { background:#2e8b57; color:white; font-weight:bold; padding:8px 20px; border-radius:6px; }"
            "QPushButton:hover { background:#256f46; }")
        to_list_btn.clicked.connect(self._dyn_send_to_list)
        bottom.addWidget(to_list_btn)
        layout.addLayout(bottom)

        self._dyn_refresh_table()
        return page

    def _dyn_current_name(self):
        return self.dyn_list_combo.currentText()

    def _dyn_current_cases(self):
        return self._dyn_lists.setdefault(self._dyn_current_name() or self.DYN_DEFAULT_LIST, [])

    def _dyn_refresh_table(self):
        cases = self._dyn_current_cases()
        self.dyn_table.setRowCount(len(cases))
        for r, case in enumerate(cases):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            chk.setCheckState(Qt.Checked)
            self.dyn_table.setItem(r, 0, chk)
            extra = {k: v for k, v in case.items()
                     if k not in ('script', 'nr_band', 'nr_bw', 'scs', 'range', 'rb_mode')}
            vals = [case.get('script', ''), case.get('nr_band', ''), case.get('nr_bw', ''),
                    case.get('scs', ''), case.get('range', ''), case.get('rb_mode', ''),
                    json.dumps(extra, ensure_ascii=False)]
            for c, val in enumerate(vals, start=1):
                it = QTableWidgetItem(str(val))
                if c in (2, 3, 4):
                    it.setTextAlignment(Qt.AlignCenter)
                self.dyn_table.setItem(r, c, it)
        self.dyn_count_label.setText(f"共 {len(cases)} 条用例")

    def _dyn_set_all_checked(self, checked):
        state = Qt.Checked if checked else Qt.Unchecked
        for r in range(self.dyn_table.rowCount()):
            it = self.dyn_table.item(r, 0)
            if it:
                it.setCheckState(state)

    def _dyn_move_rows(self, delta):
        cases = self._dyn_current_cases()
        sel = sorted({i.row() for i in self.dyn_table.selectedIndexes()})
        if not sel:
            return
        if delta < 0 and sel[0] == 0:
            return
        if delta > 0 and sel[-1] >= len(cases) - 1:
            return
        checks = [self.dyn_table.item(r, 0).checkState() == Qt.Checked
                  for r in range(self.dyn_table.rowCount())]
        order = sel if delta < 0 else list(reversed(sel))
        for r in order:
            cases[r], cases[r + delta] = cases[r + delta], cases[r]
            checks[r], checks[r + delta] = checks[r + delta], checks[r]
        self._dyn_save_store()
        self._dyn_refresh_table()
        for r, ck in enumerate(checks):
            it = self.dyn_table.item(r, 0)
            if it:
                it.setCheckState(Qt.Checked if ck else Qt.Unchecked)
        self.dyn_table.clearSelection()
        for r in [s + delta for s in sel]:
            self.dyn_table.selectRow(r) if len(sel) == 1 else \
                self.dyn_table.selectionModel().select(
                    self.dyn_table.model().index(r, 0),
                    QItemSelectionModel.Select | QItemSelectionModel.Rows)

    def _dyn_rename_list(self):
        old = self._dyn_current_name()
        if not old:
            return
        name, ok = QInputDialog.getText(self, "重命名列表", "新名称:", text=old)
        name = (name or "").strip()
        if not ok or not name or name == old:
            return
        if name in self._dyn_lists:
            QMessageBox.warning(self, "重名", f"列表 '{name}' 已存在")
            return
        self._dyn_lists[name] = self._dyn_lists.pop(old)
        idx = self.dyn_list_combo.currentIndex()
        self.dyn_list_combo.setItemText(idx, name)
        self._dyn_save_store()
        self.status_label.setText(f"列表已重命名: {old} → {name}")

    def _dyn_save_clicked(self):
        self._dyn_save_store()
        self.status_label.setText(
            f"动态Case List已保存 ({len(self._dyn_lists)} 个列表, "
            f"当前列表 {len(self._dyn_current_cases())} 条用例)")

    def _dyn_append_cases(self, new_cases):
        self._dyn_current_cases().extend(new_cases)
        self._dyn_save_store()
        self._dyn_refresh_table()

    def _dyn_replace_cases(self, new_cases, list_name=None):
        name = list_name or self.DYN_DEFAULT_LIST
        self._dyn_lists[name] = list(new_cases)
        if self.dyn_list_combo.findText(name) < 0:
            self.dyn_list_combo.addItem(name)
        if self.dyn_list_combo.currentText() == name:
            self._dyn_refresh_table()  
        else:
            self.dyn_list_combo.setCurrentText(name)
        self._dyn_save_store()

    def _dyn_add_list(self):
        name, ok = QInputDialog.getText(self, "新建列表", "列表名称:")
        name = (name or "").strip()
        if not ok or not name:
            return
        if name in self._dyn_lists:
            QMessageBox.warning(self, "重名", f"列表 '{name}' 已存在")
            return
        self._dyn_lists[name] = []
        self.dyn_list_combo.addItem(name)
        self.dyn_list_combo.setCurrentText(name)
        self._dyn_save_store()

    def _dyn_del_list(self):
        name = self._dyn_current_name()
        if not name:
            return
        if len(self._dyn_lists) <= 1:
            QMessageBox.information(self, "无法删除", "至少保留一个列表")
            return
        if QMessageBox.question(self, "确认删除",
                                f"删除列表 '{name}' 及其 {len(self._dyn_lists.get(name, []))} 条用例？",
                                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return
        self._dyn_lists.pop(name, None)
        self.dyn_list_combo.removeItem(self.dyn_list_combo.currentIndex())
        self._dyn_save_store()
        self._dyn_refresh_table()

    def _dyn_del_rows(self):
        rows = sorted({i.row() for i in self.dyn_table.selectedIndexes()}, reverse=True)
        if not rows:
            return
        cases = self._dyn_current_cases()
        for r in rows:
            if 0 <= r < len(cases):
                cases.pop(r)
        self._dyn_save_store()
        self._dyn_refresh_table()

    def _dyn_clear_list(self):
        cases = self._dyn_current_cases()
        if not cases:
            return
        if QMessageBox.question(self, "确认清空",
                                f"清空当前列表的 {len(cases)} 条用例？",
                                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return
        cases.clear()
        self._dyn_save_store()
        self._dyn_refresh_table()

    def _dyn_send_to_list(self):
        cases = self._dyn_current_cases()
        chosen = [cases[r] for r in range(self.dyn_table.rowCount())
                  if r < len(cases) and self.dyn_table.item(r, 0)
                  and self.dyn_table.item(r, 0).checkState() == Qt.Checked]
        if not chosen:
            QMessageBox.information(self, "无用例", "请先在首列勾选要加入的用例")
            return
        target = self.dyn_target_combo.currentText()  
        new_cases = [dict(c, ue_cap=True) for c in chosen]
        table = self.tables.get(target)
        if table is None:
            QMessageBox.warning(self, "错误", f"未找到 {target} 用例列表")
            return
        existing = self._table_to_cases(table)
        self._populate_table(target, existing + new_cases)
        self._sync_cases_to_config()
        self.page_selector.setCurrentIndex(0)
        for i in range(self.tab_widget.count()):
            if self.tab_widget.tabText(i) == target:
                self.tab_widget.setCurrentIndex(i)
                break
        self.status_label.setText(
            f"已从动态Case List加入 {len(new_cases)} 条用例到 {target} 列表(重启后自动清除)")

    def _build_manual_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        splitter = QSplitter(Qt.Vertical)

        top_group = QGroupBox("① 单独下发命令")
        top_v = QVBoxLayout(top_group)
        send_row = QHBoxLayout()
        self.manual_cmd_edit = HistoryLineEdit(max_history=50)
        self.manual_cmd_edit.setPlaceholderText("输入 SCPI 命令，↑↓ 调取历史，例如: *IDN?  或  CALL:CELL1 ON")
        self.manual_cmd_edit.returnPressed.connect(self.manual_send_single)

        self.manual_history_combo = QComboBox()
        self.manual_history_combo.setFixedWidth(210)
        self.manual_history_combo.setToolTip("历史命令（点击选取到输入框）")
        self.manual_history_combo.activated.connect(self._on_history_combo_picked)
        self.manual_send_btn = QPushButton("▶ 下发")
        self.manual_send_btn.setObjectName("manualSendBtn")
        self.manual_send_btn.clicked.connect(self.manual_send_single)
        self.manual_disconnect_btn = QPushButton("🔌 断开重连")
        self.manual_disconnect_btn.clicked.connect(self.manual_reset_connection)
        send_row.addWidget(QLabel("命令:"))
        send_row.addWidget(self.manual_cmd_edit, 1)
        send_row.addWidget(QLabel("历史:"))
        send_row.addWidget(self.manual_history_combo)
        send_row.addWidget(self.manual_send_btn)
        send_row.addWidget(self.manual_disconnect_btn)
        top_v.addLayout(send_row)
        splitter.addWidget(top_group)

        cfg_group = QGroupBox("② 配置下发（保存多组命令，一键下发）")
        cfg_h = QHBoxLayout(cfg_group)

        left_v = QVBoxLayout()
        left_v.addWidget(QLabel("配置列表:"))
        self.manual_cfg_list = QListWidget()
        self.manual_cfg_list.setFixedWidth(200)
        self.manual_cfg_list.currentRowChanged.connect(self._on_manual_cfg_selected)
        left_v.addWidget(self.manual_cfg_list, 1)
        cfg_btn_row = QHBoxLayout()
        add_cfg_btn = QPushButton("➕ add")
        add_cfg_btn.setObjectName("addBtn")
        add_cfg_btn.clicked.connect(self.manual_cfg_add)
        del_cfg_btn = QPushButton("➖ del")
        del_cfg_btn.setObjectName("delBtn")
        del_cfg_btn.clicked.connect(self.manual_cfg_del)
        rename_cfg_btn = QPushButton("✎ 改名")
        rename_cfg_btn.clicked.connect(self.manual_cfg_rename)
        cfg_btn_row.addWidget(add_cfg_btn)
        cfg_btn_row.addWidget(del_cfg_btn)
        cfg_btn_row.addWidget(rename_cfg_btn)
        left_v.addLayout(cfg_btn_row)
        cfg_h.addLayout(left_v)

        right_v = QVBoxLayout()
        right_v.addWidget(QLabel("命令内容（每行一条）:"))
        self.manual_cfg_edit = QPlainTextEdit()
        self.manual_cfg_edit.setPlaceholderText("每行一条 SCPI 命令，例如:\nCONFigure:CELL1:NR:Sign:SLOT:Clear\nCONFigure:CELL1:NR:SIGN:SLOT3:CTYPe PDSCh")
        right_v.addWidget(self.manual_cfg_edit, 1)
        cfg_act_row = QHBoxLayout()
        save_cfg_btn = QPushButton("💾 保存")
        save_cfg_btn.clicked.connect(self.manual_cfg_save)
        send_cfg_btn = QPushButton("▶ 下发本组")
        send_cfg_btn.setObjectName("manualSendBtn")
        send_cfg_btn.clicked.connect(self.manual_send_config)
        cfg_act_row.addStretch()
        cfg_act_row.addWidget(save_cfg_btn)
        cfg_act_row.addWidget(send_cfg_btn)
        right_v.addLayout(cfg_act_row)
        cfg_h.addLayout(right_v, 1)
        splitter.addWidget(cfg_group)

        echo_group = QGroupBox("③ 回显（下发=蓝色，返回=绿色，错误=红色）")
        echo_v = QVBoxLayout(echo_group)
        self.manual_echo = QPlainTextEdit()
        self.manual_echo.setReadOnly(True)
        self.manual_echo.setStyleSheet(
            "QPlainTextEdit { background: #1e1e1e; color: #e0e0e0; font-family: Consolas, 'Microsoft YaHei'; font-size: 10pt; border-radius: 6px; }"
        )
        echo_v.addWidget(self.manual_echo)
        echo_row = QHBoxLayout()
        clear_echo_btn = QPushButton("🧹 清空回显")
        clear_echo_btn.clicked.connect(lambda: self.manual_echo.clear())
        echo_row.addStretch()
        echo_row.addWidget(clear_echo_btn)
        echo_v.addLayout(echo_row)
        splitter.addWidget(echo_group)

        splitter.setSizes([90, 260, 260])
        layout.addWidget(splitter)

        self._manual_configs = self._load_manual_configs()
        self._refresh_manual_cfg_list()

        self.manual_cmd_edit._history = self._load_cmd_history()
        self._refresh_history_combo()
        return page

    def _cmd_history_path(self):
        return os.path.join(project_root(), "manual_cmd_history.json")

    def _load_cmd_history(self):
        path = self._cmd_history_path()
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, list):

                    seen, out = set(), []
                    for c in data:
                        c = str(c).strip()
                        if c and c not in seen:
                            seen.add(c)
                            out.append(c)
                    return out[-self.manual_cmd_edit.max_history:]
            except Exception:
                pass
        return []

    def _save_cmd_history(self):
        try:
            with open(self._cmd_history_path(), "w", encoding="utf-8") as f:
                json.dump(self.manual_cmd_edit._history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存命令历史失败: {e}")

    def _refresh_history_combo(self):
        combo = self.manual_history_combo
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("— 历史命令 —")
        for cmd in reversed(self.manual_cmd_edit._history):
            combo.addItem(cmd)
        combo.setCurrentIndex(0)
        combo.blockSignals(False)

    def _on_history_combo_picked(self, idx):
        if idx <= 0:
            return
        self.manual_cmd_edit.setText(self.manual_history_combo.itemText(idx))
        self.manual_cmd_edit.setFocus()

    def _manual_configs_path(self):
        return os.path.join(project_root(), "manual_configs.json")

    def _load_manual_configs(self):
        path = self._manual_configs_path()
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, list):
                    return data
            except Exception as e:
                print(f"读取 manual_configs.json 失败: {e}")

        return [{"name": "示例配置", "commands": ["*IDN?"]}]

    def _save_manual_configs(self):
        path = self._manual_configs_path()
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self._manual_configs, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.warning(self, "保存失败", f"写入 manual_configs.json 失败:\n{e}")

    def _refresh_manual_cfg_list(self, select_row=None):
        self.manual_cfg_list.blockSignals(True)
        self.manual_cfg_list.clear()
        for cfg in self._manual_configs:
            self.manual_cfg_list.addItem(cfg.get("name", "未命名"))
        self.manual_cfg_list.blockSignals(False)
        if self._manual_configs:
            row = 0 if select_row is None else max(0, min(select_row, len(self._manual_configs) - 1))
            self.manual_cfg_list.setCurrentRow(row)

    def _on_manual_cfg_selected(self, row):
        if 0 <= row < len(self._manual_configs):
            cmds = self._manual_configs[row].get("commands", [])
            self.manual_cfg_edit.setPlainText("\n".join(cmds))

    def manual_cfg_add(self):
        name, ok = QInputDialog.getText(self, "新建配置", "请输入配置名称:")
        if ok and name.strip():
            name = name.strip()
            if any(c.get("name") == name for c in self._manual_configs):
                QMessageBox.warning(self, "重复", f"配置 '{name}' 已存在")
                return
            self._manual_configs.append({"name": name, "commands": []})
            self._save_manual_configs()
            self._refresh_manual_cfg_list(select_row=len(self._manual_configs) - 1)

    def manual_cfg_del(self):
        row = self.manual_cfg_list.currentRow()
        if row < 0 or row >= len(self._manual_configs):
            return
        name = self._manual_configs[row].get("name", "")
        if QMessageBox.question(self, "确认删除", f"确定删除配置 '{name}' 吗？",
                                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            del self._manual_configs[row]
            self._save_manual_configs()
            self._refresh_manual_cfg_list(select_row=row - 1)

    def manual_cfg_rename(self):
        row = self.manual_cfg_list.currentRow()
        if row < 0 or row >= len(self._manual_configs):
            return
        old = self._manual_configs[row].get("name", "")
        name, ok = QInputDialog.getText(self, "重命名配置", "新名称:", text=old)
        if ok and name.strip():
            name = name.strip()
            if any(i != row and c.get("name") == name for i, c in enumerate(self._manual_configs)):
                QMessageBox.warning(self, "重复", f"配置 '{name}' 已存在")
                return
            self._manual_configs[row]["name"] = name
            self._save_manual_configs()
            self._refresh_manual_cfg_list(select_row=row)

    def manual_cfg_save(self):
        row = self.manual_cfg_list.currentRow()
        if row < 0 or row >= len(self._manual_configs):
            QMessageBox.information(self, "提示", "请先在左侧选择或新建一个配置")
            return
        cmds = [line.strip() for line in self.manual_cfg_edit.toPlainText().splitlines() if line.strip()]
        self._manual_configs[row]["commands"] = cmds
        self._save_manual_configs()
        self.manual_echo_append('info', f"已保存配置 '{self._manual_configs[row].get('name')}'（{len(cmds)} 条命令）")

    def manual_echo_append(self, kind, text):
        colors = {'send': '#4fc3f7', 'recv': '#81c784', 'error': '#e57373', 'info': '#ffd54f'}
        prefix = {'send': '📤 Send &gt;', 'recv': '📥 Recv &lt;', 'error': '❌ Error', 'info': 'ℹ'}
        color = colors.get(kind, '#e0e0e0')
        label = prefix.get(kind, '')

        if kind == 'recv' and not (text or '').strip():
            text = '(空响应)'
        safe = (text or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

        from PyQt5.QtGui import QTextCursor
        cursor = self.manual_echo.textCursor()
        cursor.movePosition(QTextCursor.End)
        if not self.manual_echo.document().isEmpty():
            cursor.insertBlock()
        cursor.insertHtml(f'<span style="color:{color}; white-space:pre;">{label} {safe}</span>')
        self.manual_echo.setTextCursor(cursor)
        self.manual_echo.ensureCursorVisible()

    def _sync_manual_ip(self):
        try:
            souren_config.DEFAULT_IP = self.ip_edit.text().strip()
            souren_config.INSTRUMENT_ADDRESS = f"TCPIP0::{souren_config.DEFAULT_IP}::inst0::INSTR"
        except Exception:
            pass

    def _run_manual_commands(self, commands):
        if getattr(self, '_manual_thread', None) and self._manual_thread.isRunning():
            QMessageBox.information(self, "忙", "上一批命令还在下发中，请稍候")
            return
        self._sync_manual_ip()
        self.manual_send_btn.setEnabled(False)
        self._manual_thread = ManualCommandThread(commands)
        self._manual_thread.line_signal.connect(self.manual_echo_append)
        self._manual_thread.finished_signal.connect(self._on_manual_finished)
        self._manual_thread.start()

    def _on_manual_finished(self, all_ok):
        self.manual_send_btn.setEnabled(True)

    def manual_send_single(self):
        cmd = self.manual_cmd_edit.text().strip()
        if not cmd:
            return
        self.manual_cmd_edit.add_history(cmd)
        self._refresh_history_combo()
        self._save_cmd_history()
        self._run_manual_commands([cmd])

    def manual_send_config(self):
        cmds = [line.strip() for line in self.manual_cfg_edit.toPlainText().splitlines() if line.strip()]
        if not cmds:
            QMessageBox.information(self, "提示", "当前配置没有命令")
            return
        self._run_manual_commands(cmds)

    def manual_reset_connection(self):
        ManualCommandThread.reset_controller()
        self.manual_echo_append('info', "已断开连接，下次下发将重新连接")

    def _config_file_path(self):

        candidate = os.path.join(project_root(), "souren_config.py")
        if os.path.exists(candidate):
            return candidate
        module_file = getattr(souren_config, "__file__", "") or ""
        if module_file and os.path.exists(module_file):
            return module_file
        return candidate

    def _setup_config_watcher(self):
        self._config_watcher = QFileSystemWatcher(self)
        path = self._config_file_path()
        if os.path.exists(path):
            self._config_watcher.addPath(path)
        self._config_watcher.fileChanged.connect(self._on_config_file_changed)

        try:
            self._config_mtime = os.path.getmtime(path) if os.path.exists(path) else 0.0
        except Exception:
            self._config_mtime = 0.0
        self._config_poll_timer = QTimer(self)
        self._config_poll_timer.timeout.connect(self._poll_config_mtime)
        self._config_poll_timer.start(1000)

    def _poll_config_mtime(self):
        path = self._config_file_path()
        if not os.path.exists(path):
            return
        try:
            mtime = os.path.getmtime(path)
        except Exception:
            return
        if mtime == getattr(self, '_config_mtime', 0.0):
            return
        self._config_mtime = mtime
        if path not in self._config_watcher.files():
            self._config_watcher.addPath(path)
        if time.time() < self._suppress_watch_until:
            return
        if self.thread and self.thread.isRunning():
            return
        self._reload_config_from_file()

    def _on_config_file_changed(self, path):

        def _readd():
            if os.path.exists(path) and path not in self._config_watcher.files():
                self._config_watcher.addPath(path)
        QTimer.singleShot(200, _readd)

        if time.time() < self._suppress_watch_until:
            return

        if self.thread and self.thread.isRunning():
            return

        QTimer.singleShot(300, self._reload_config_from_file)

    def _reload_config_from_file(self):
        try:
            importlib.reload(souren_config)
        except Exception as e:
            self.status_label.setText(f"配置重载失败: {e}")
            return

        self._suspend_case_sync = True
        try:
            self.ip_edit.setText(str(souren_config.DEFAULT_IP))
            self.port_edit.setText(str(souren_config.REMOTE_SERVER_PORT))
            self.password_edit.setText(str(souren_config.REMOTE_SUDO_PASSWORD))
            self.version_edit.setText(str(souren_config.VERSION))
            self.log_collect_check.setChecked(bool(souren_config.COLLECT_CORE_NETWORK_LOGS))
            if hasattr(self, 'debug_log_check'):
                self.debug_log_check.setChecked(bool(getattr(souren_config, 'DEBUG_LOG_ENABLED', True)))
            if hasattr(self, 'loop_count_spin'):
                self.loop_count_spin.blockSignals(True)
                self.loop_count_spin.setValue(int(getattr(souren_config, 'LOOP_COUNT', 1) or 1))
                self.loop_count_spin.blockSignals(False)

            log_params = list(souren_config.LOG_LEVEL_PARAMS.items())
            half = (len(log_params) + 1) // 2
            for i in range(half):
                k, v = log_params[i]
                if self.log_level_table.item(i, 0):
                    self.log_level_table.item(i, 0).setText(k)
                if self.log_level_table.item(i, 1):
                    self.log_level_table.item(i, 1).setText(str(v))
                j = i + half
                if j < len(log_params):
                    k2, v2 = log_params[j]
                    if self.log_level_table.item(i, 2):
                        self.log_level_table.item(i, 2).setText(k2)
                    if self.log_level_table.item(i, 3):
                        self.log_level_table.item(i, 3).setText(str(v2))

            skip = souren_config.LTE_MEASUREMENT_REPORT_CHECK_SKIP_SCRIPTS
            self.skip_scripts_table.setRowCount(len(skip))
            for i, nm in enumerate(skip):
                self.skip_scripts_table.setItem(i, 0, QTableWidgetItem(nm))

            if "Phone" in self.tables:
                self._populate_table("Phone", getattr(souren_config, "PHONE_SCRIPTS", []))
            if "AT" in self.tables:
                self._populate_table("AT", getattr(souren_config, "AT_SCRIPTS", []))
            if "NR" in self.tables:
                self._populate_table("NR", getattr(souren_config, "NR_SCRIPTS", []))
            if "LTE" in self.tables:
                self._populate_table("LTE", getattr(souren_config, "LTE_SCRIPTS", []))
            for var_name, tab_name, value in self._discover_custom_script_vars():
                if tab_name in self.tables:
                    self._populate_table(tab_name, value)
                else:
                    self._add_tab(tab_name, value)
        finally:
            self._suspend_case_sync = False

        self.status_label.setText("⟳ 已从配置文件重新加载")

    def _persist_var_to_config(self, var_name, new_literal_text):
        path = self._config_file_path()
        try:
            with open(path, "r", encoding="utf-8-sig") as f:
                source = f.read()
            lines = source.splitlines(keepends=True)
            tree = ast.parse(source)
            target_node = None
            for node in tree.body:
                if isinstance(node, ast.Assign):
                    for t in node.targets:
                        if isinstance(t, ast.Name) and t.id == var_name:
                            target_node = node
                            break
                if target_node:
                    break
            new_block = f"{var_name} = {new_literal_text}\n"
            if target_node is not None:
                start = target_node.lineno - 1
                end = getattr(target_node, "end_lineno", target_node.lineno)
                lines[start:end] = [new_block]
            else:
                insert_at = None
                if var_name.endswith("_SCRIPTS"):
                    last_end = None
                    for node in tree.body:
                        if isinstance(node, ast.Assign):
                            for t in node.targets:
                                if (isinstance(t, ast.Name) and t.id.endswith("_SCRIPTS")
                                        and t.id not in self._RESERVED_SCRIPT_VARS):
                                    last_end = getattr(node, "end_lineno", node.lineno)
                    insert_at = last_end
                if insert_at is not None:
                    lines[insert_at:insert_at] = ["\n" + new_block]
                else:
                    if lines and not lines[-1].endswith("\n"):
                        lines[-1] += "\n"
                    lines.append("\n" + new_block)
            with open(path, "w", encoding="utf-8") as f:
                f.writelines(lines)

            self._suppress_watch_until = time.time() + 1.5
            return True
        except Exception as e:
            QMessageBox.warning(self, "同步失败", f"写入 souren_config.py 失败:\n{e}")
            return False

    def _config_has_var(self, var_name):
        try:
            path = self._config_file_path()
            with open(path, "r", encoding="utf-8-sig") as f:
                tree = ast.parse(f.read())
            for node in tree.body:
                if isinstance(node, ast.Assign):
                    for t in node.targets:
                        if isinstance(t, ast.Name) and t.id == var_name:
                            return True
        except Exception as e:
            print(f"回读 souren_config 校验 {var_name} 失败: {e}")
        return False

    def _ensure_var_persisted(self, var_name, literal_text, retries=2):
        for _ in range(max(1, retries)):
            self._persist_var_to_config(var_name, literal_text)
            if self._config_has_var(var_name):
                return True
        return False

    def _format_cases_literal(self, cases):
        if not cases:
            return "[]"
        lines = ["["]
        for case in cases:
            lines.append("    " + repr(case) + ",")
        lines.append("]")
        return "\n".join(lines)

    @staticmethod
    def _format_params_str(params):
        def _has_non_str_key(obj):
            if isinstance(obj, dict):
                for k, v in obj.items():
                    if not isinstance(k, str):
                        return True
                    if _has_non_str_key(v):
                        return True
            elif isinstance(obj, (list, tuple)):
                for it in obj:
                    if _has_non_str_key(it):
                        return True
            return False

        if _has_non_str_key(params):
            return repr(params)
        try:

            return json.dumps(params, ensure_ascii=False, separators=(', ', ': '))
        except Exception:
            return repr(params)

    @staticmethod
    def _parse_params(param_str):
        param_str = (param_str or "").strip()
        if not param_str:
            return {}
        try:
            return ast.literal_eval(param_str)
        except Exception:
            return json.loads(param_str)

    def _row_to_case(self, table, row):
        script_item = table.item(row, 2)
        param_item = table.item(row, 3)
        script = script_item.text().strip() if script_item else ""
        if not script:
            return None
        param_str = param_item.text().strip() if param_item else ""
        try:
            params = self._parse_params(param_str)
        except Exception:
            params = {}
        case = {"script": script}
        case.update(params)
        return case

    def _table_to_cases(self, table):
        cases = []
        for row in range(table.rowCount()):
            case = self._row_to_case(table, row)
            if case:
                cases.append(case)
        return cases

    _BUILTIN_TAB_VARS = {
        "Phone": "PHONE_SCRIPTS",
        "AT": "AT_SCRIPTS",
        "NR": "NR_SCRIPTS",
        "LTE": "LTE_SCRIPTS",
    }
    _RESERVED_SCRIPT_VARS = {"LTE_MEASUREMENT_REPORT_CHECK_SKIP_SCRIPTS"}

    DYN_DEFAULT_LIST = "默认"

    def _config_var_for_tab(self, name):
        if name in self._BUILTIN_TAB_VARS:
            return self._BUILTIN_TAB_VARS[name]
        import re as _re
        ident = _re.sub(r'\W', '_', (name or '').strip())
        if not ident:
            ident = "SHEET"
        if ident[0].isdigit():
            ident = "_" + ident
        return f"{ident}_SCRIPTS"

    def _tab_name_for_var(self, var_name):
        rev = {v: k for k, v in self._BUILTIN_TAB_VARS.items()}
        if var_name in rev:
            return rev[var_name]
        return var_name[:-len("_SCRIPTS")] if var_name.endswith("_SCRIPTS") else var_name

    def _sync_cases_to_config(self):
        for tab_name, table in getattr(self, 'tables', {}).items():
            var_name = self._config_var_for_tab(tab_name)
            if var_name in self._RESERVED_SCRIPT_VARS:
                continue
            cases = self._table_to_cases(table)
            setattr(souren_config, var_name, cases)
            self._persist_var_to_config(var_name, self._format_cases_literal(cases))

    def _sync_skip_scripts_to_config(self):
        skip = []
        for row in range(self.skip_scripts_table.rowCount()):
            item = self.skip_scripts_table.item(row, 0)
            if item and item.text().strip():
                skip.append(item.text().strip())
        souren_config.LTE_MEASUREMENT_REPORT_CHECK_SKIP_SCRIPTS = skip
        self._persist_var_to_config(
            "LTE_MEASUREMENT_REPORT_CHECK_SKIP_SCRIPTS", repr(skip)
        )

    def sync_all_to_config(self):
        self._update_global_config()
        self._sync_cases_to_config()
        self._sync_skip_scripts_to_config()
        self._persist_global_config()
        if hasattr(self, 'loop_count_spin'):
            self._persist_var_to_config("LOOP_COUNT", repr(int(self.loop_count_spin.value())))

    def _persist_global_config(self):
        try:
            self._persist_var_to_config("DEFAULT_IP", repr(self.ip_edit.text().strip()))
            try:
                port = int(self.port_edit.text().strip())
                self._persist_var_to_config("REMOTE_SERVER_PORT", repr(port))
            except Exception:
                pass
            self._persist_var_to_config("REMOTE_SUDO_PASSWORD", repr(self.password_edit.text()))
            self._persist_var_to_config("VERSION", repr(self.version_edit.text().strip()))
            self._persist_var_to_config(
                "COLLECT_CORE_NETWORK_LOGS", repr(bool(self.log_collect_check.isChecked())))
            self._persist_var_to_config(
                "LOG_LEVEL_PARAMS", self._format_log_levels_literal())
        except Exception as e:
            print(f"持久化全局配置失败: {e}")

    def _current_tab_name(self):
        return self.tab_widget.tabText(self.tab_widget.currentIndex())

    def _on_search_changed(self, text):
        table = self.tables.get(self._current_tab_name())
        if not table:
            return
        kw = (text or "").strip().lower()
        for row in range(table.rowCount()):
            if not kw:
                table.setRowHidden(row, False)
                continue
            script_item = table.item(row, 2)
            param_item = table.item(row, 3)
            hay = ((script_item.text() if script_item else "") + " " +
                   (param_item.text() if param_item else "")).lower()
            table.setRowHidden(row, kw not in hay)

    def _load_initial_tabs(self):

        self._checked_states = self._load_checked_states()
        self._column_widths = self._load_column_widths()
        self._row_heights = self._load_row_heights()
        self._param_texts = self._load_param_texts()
        def _clean(var_name):
            cases = getattr(souren_config, var_name, []) or []
            kept = [c for c in cases if not (isinstance(c, dict) and c.get('ue_cap'))]
            if len(kept) != len(cases):
                print(f"🧹 启动清理: {var_name} 移除 {len(cases) - len(kept)} 条临时用例")
                setattr(souren_config, var_name, kept)
                try:
                    self._persist_var_to_config(var_name, self._format_cases_literal(kept))
                except Exception as e:
                    print(f"清理 {var_name} 写回失败: {e}")
            return kept

        nr_clean = _clean('NR_SCRIPTS')
        lte_clean = _clean('LTE_SCRIPTS')
        self._add_tab("Phone", souren_config.PHONE_SCRIPTS if hasattr(souren_config, 'PHONE_SCRIPTS') else [])
        self._add_tab("AT", souren_config.AT_SCRIPTS if hasattr(souren_config, 'AT_SCRIPTS') else [])
        self._add_tab("NR", nr_clean)
        self._add_tab("LTE", lte_clean)
        self._load_custom_sheets()
        self._apply_saved_sheet_order()

    def _discover_custom_script_vars(self):
        builtin_vars = set(self._BUILTIN_TAB_VARS.values())
        out = []
        for var_name in sorted(vars(souren_config).keys()):
            if not var_name.endswith("_SCRIPTS"):
                continue
            if var_name in builtin_vars or var_name in self._RESERVED_SCRIPT_VARS:
                continue
            value = getattr(souren_config, var_name, None)
            if not isinstance(value, list):
                continue
            if value and not all(isinstance(x, dict) for x in value):
                continue
            out.append((var_name, self._tab_name_for_var(var_name), value))
        return out

    def _load_custom_sheets(self):
        for var_name, tab_name, value in self._discover_custom_script_vars():
            if tab_name in getattr(self, 'tables', {}):
                continue
            self._add_tab(tab_name, value)
            print(f"📄 载入自定义 Sheet: {tab_name} ({len(value)} 条用例, 变量 {var_name})")

    def _checked_states_path(self):
        return os.path.join(project_root(), "case_checked_states.json")

    def _retry_settings_path(self):
        return os.path.join(project_root(), "retry_settings.json")

    def _load_retry_settings(self):
        path = self._retry_settings_path()
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    d = json.load(f)
                self.retry_check.setChecked(bool(d.get("enabled", False)))
                n = int(d.get("rounds", 1))
                self.retry_spin.setValue(max(1, min(20, n)))
            except Exception:
                pass

    def _save_retry_settings(self):
        try:
            with open(self._retry_settings_path(), "w", encoding="utf-8") as f:
                json.dump({"enabled": self.retry_check.isChecked(),
                           "rounds": self.retry_spin.value()},
                          f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存重测设置失败: {e}")

    def _load_checked_states(self):
        path = self._checked_states_path()
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
            except Exception:
                pass
        return {}

    def _save_checked_states(self):
        states = {}
        for tab_name, table in getattr(self, 'tables', {}).items():
            col = []
            for row in range(table.rowCount()):
                item = table.item(row, 0)
                col.append(bool(item and item.checkState() == Qt.Checked))
            states[tab_name] = col
        try:
            with open(self._checked_states_path(), "w", encoding="utf-8") as f:
                json.dump(states, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存勾选状态失败: {e}")

    def _param_texts_path(self):
        return os.path.join(project_root(), "case_param_texts.json")

    def _load_param_texts(self):
        path = self._param_texts_path()
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
            except Exception:
                pass
        return {}

    def _save_param_texts(self):
        data = {}
        for tab_name, table in getattr(self, 'tables', {}).items():
            texts = []
            for row in range(table.rowCount()):
                item = table.item(row, 3)
                texts.append(item.text() if item else "")
            data[tab_name] = texts
        self._param_texts = data
        try:
            with open(self._param_texts_path(), "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存参数文本失败: {e}")

    def _saved_param_text_for(self, tab_name, row, params):
        texts = getattr(self, '_param_texts', {}).get(tab_name)
        if not texts or row >= len(texts):
            return None
        raw = texts[row]
        if not raw or '\n' not in raw:
            return None
        try:
            parsed = self._parse_params(raw)
            if isinstance(parsed, dict):
                parsed.pop('script', None)
            if parsed == params:
                return raw
        except Exception:
            pass
        return None

    def _column_widths_path(self):
        return os.path.join(project_root(), "case_column_widths.json")

    def _load_column_widths(self):
        path = self._column_widths_path()
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
            except Exception:
                pass
        return {}

    def _apply_column_widths(self, name, table):
        widths = getattr(self, '_column_widths', {}).get(name)
        if not widths:
            return
        for col_str, w in widths.items():
            try:
                col = int(col_str)
                if 0 <= col < table.columnCount() and int(w) > 0:
                    table.setColumnWidth(col, int(w))
            except Exception:
                pass

    def _save_column_widths(self):
        data = {}
        for tab_name, table in getattr(self, 'tables', {}).items():
            data[tab_name] = {str(c): table.columnWidth(c)
                              for c in range(table.columnCount())}
        self._column_widths = data
        try:
            with open(self._column_widths_path(), "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存列宽失败: {e}")

    def _on_column_resized(self, table, *_):
        self._fit_param_row_heights(table)
        if not hasattr(self, '_colw_save_timer'):
            self._colw_save_timer = QTimer(self)
            self._colw_save_timer.setSingleShot(True)
            self._colw_save_timer.timeout.connect(self._save_column_widths)
        self._colw_save_timer.start(400)

    # ---------- 单行行高持久化(用户手动拉长/缩小某行, 重启后保持) ----------
    def _row_heights_path(self):
        return os.path.join(project_root(), "case_row_heights.json")

    def _load_row_heights(self):
        path = self._row_heights_path()
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
            except Exception:
                pass
        return {}

    def _tab_name_of(self, table):
        for n, t in getattr(self, 'tables', {}).items():
            if t is table:
                return n
        return None

    def _apply_row_heights(self, name, table):
        saved = getattr(self, '_row_heights', {}).get(name)
        if not saved:
            return
        self._suspend_row_save = True
        try:
            for row_str, h in saved.items():
                try:
                    r = int(row_str)
                    if 0 <= r < table.rowCount() and int(h) > 0:
                        table.setRowHeight(r, int(h))
                except Exception:
                    pass
        finally:
            self._suspend_row_save = False

    def _save_row_heights(self):
        try:
            with open(self._row_heights_path(), "w", encoding="utf-8") as f:
                json.dump(getattr(self, '_row_heights', {}), f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存行高失败: {e}")

    def _on_row_resized(self, table, logical, old, new):
        # 只记录用户手动拖动的行高; 程序性 setRowHeight(自动折行/加载)时 _suspend_row_save=True 会跳过
        if getattr(self, '_suspend_row_save', False):
            return
        tab_name = self._tab_name_of(table)
        if not tab_name:
            return
        self._row_heights.setdefault(tab_name, {})[str(logical)] = int(new)
        if not hasattr(self, '_rowh_save_timer'):
            self._rowh_save_timer = QTimer(self)
            self._rowh_save_timer.setSingleShot(True)
            self._rowh_save_timer.timeout.connect(self._save_row_heights)
        self._rowh_save_timer.start(400)

    def save_all(self):
        try:
            self.sync_all_to_config()
            self._save_checked_states()
            self._save_column_widths()
            self._save_row_heights()
            self._save_param_texts()
            self.status_label.setText("💾 已保存用例列表 / 勾选状态 / 列宽 / 行高 / 换行")
        except Exception as e:
            QMessageBox.warning(self, "保存失败", f"保存失败:\n{e}")

    def _add_tab(self, name, cases):
        tab_widget = QWidget()
        tab_layout = QVBoxLayout(tab_widget)

        table = DraggableTable()
        table.setColumnCount(7)
        table.setHorizontalHeaderLabels(["选择", "序号", "脚本名", "参数(JSON)", "进度(%)", "状态", "通过率(%)"])

        hdr = table.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.Interactive)
        hdr.setStretchLastSection(False)

        hdr.setSectionResizeMode(3, QHeaderView.Interactive)
        table.setColumnWidth(0, 50)
        table.setColumnWidth(1, 50)
        table.setColumnWidth(2, 160)
        table.setColumnWidth(3, 620)
        table.setColumnWidth(4, 70)
        table.setColumnWidth(5, 70)
        table.setColumnWidth(6, 80)
        self._apply_column_widths(name, table)
        # 参数列使用多行编辑器：支持 Alt+Enter 换行
        table.setItemDelegateForColumn(3, MultiLineParamDelegate(table))
        table.setAlternatingRowColors(True)
        table.verticalHeader().setDefaultSectionSize(28)

        _row_h = 28
        table.setMinimumHeight(_row_h * 7 + 34)
        table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        table.setFont(QFont("Consolas", 10))
        table.setStyleSheet("""
            QTableWidget { background-color: #ffffff; gridline-color: #e6ecf3; alternate-background-color: #f4f8fc; font-size: 10.5pt; }
            QTableWidget::item { padding: 3px 6px; }
            QTableWidget::item:selected { background-color: #b3d9ff; color: #000000; }
            QHeaderView::section { background-color: #2c3e50; color: white; font-weight: bold; padding: 5px; font-size: 10pt; }
        """)
        tab_layout.addWidget(table)

        if not hasattr(self, 'tables'):
            self.tables = {}
        self.tables[name] = table

        self._populate_table(name, cases)

        table.itemChanged.connect(lambda item, n=name: self._on_case_item_changed(n, item))

        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(
            lambda pos, n=name: self._show_case_context_menu(n, pos))

        table.horizontalHeader().sectionResized.connect(
            lambda *_a, t=table: self._on_column_resized(t))

        # 允许并持久化单行行高: 用户拖动行边框拉长/缩小某行, 重启后保持
        table.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)
        table.verticalHeader().sectionResized.connect(
            lambda logical, old, new, t=table: self._on_row_resized(t, logical, old, new))
        self._apply_row_heights(name, table)

        self.tab_widget.addTab(tab_widget, name)

    def add_sheet(self):
        name, ok = QInputDialog.getText(self, "新建Sheet", "请输入Sheet名称:")
        if ok and name.strip():
            name = name.strip()
            if name in self.tables:
                QMessageBox.warning(self, "重复", f"Sheet '{name}' 已存在")
                return
            var_name = self._config_var_for_tab(name)
            existing_vars = {self._config_var_for_tab(t) for t in self.tables}
            if var_name in existing_vars or var_name in self._RESERVED_SCRIPT_VARS:
                QMessageBox.warning(self, "名称冲突",
                                    f"该 Sheet 名对应的配置变量 {var_name} 已被占用，请换个名字")
                return
            self._add_tab(name, [])
            setattr(souren_config, var_name, [])
            if self._ensure_var_persisted(var_name, "[]"):
                msg = f"已新增 Sheet '{name}'，已同步 souren_config.{var_name}"
            else:
                msg = f"⚠ Sheet '{name}' 已创建，但同步 souren_config.{var_name} 失败"
                QMessageBox.warning(
                    self, "同步失败",
                    f"Sheet '{name}' 已创建，但未能写入 souren_config.{var_name}。\n"
                    f"请确认 souren_config.py 未被占用/无语法错误，然后点『💾 保存』重试。")
            self.tab_widget.setCurrentIndex(self.tab_widget.count() - 1)
            self._save_sheet_order()
            if hasattr(self, 'status_label'):
                self.status_label.setText(msg)

    def delete_current_sheet(self):
        if self.tab_widget.count() <= 1:
            QMessageBox.information(self, "提示", "至少保留一个标签页")
            return
        current_idx = self.tab_widget.currentIndex()
        tab_name = self.tab_widget.tabText(current_idx)
        reply = QMessageBox.question(self, "确认删除", f"确定要删除 Sheet '{tab_name}' 吗？",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.tab_widget.removeTab(current_idx)
            del self.tables[tab_name]
            if tab_name not in self._BUILTIN_TAB_VARS:
                var_name = self._config_var_for_tab(tab_name)
                if var_name not in self._RESERVED_SCRIPT_VARS:
                    self._remove_var_from_config(var_name)
                    if hasattr(souren_config, var_name):
                        try:
                            delattr(souren_config, var_name)
                        except Exception:
                            pass
            self._save_sheet_order()

    def move_sheet_left(self):
        self._move_sheet(-1)

    def move_sheet_right(self):
        self._move_sheet(1)

    def _move_sheet(self, delta):
        cur = self.tab_widget.currentIndex()
        tgt = cur + delta
        if cur < 0 or tgt < 0 or tgt >= self.tab_widget.count():
            return
        self.tab_widget.tabBar().moveTab(cur, tgt)
        self.tab_widget.setCurrentIndex(tgt)
        self._save_sheet_order()

    def _sheet_order_path(self):
        return os.path.join(project_root(), "case_sheet_order.json")

    def _save_sheet_order(self):
        order = [self.tab_widget.tabText(i) for i in range(self.tab_widget.count())]
        try:
            with open(self._sheet_order_path(), "w", encoding="utf-8") as f:
                json.dump(order, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存 Sheet 顺序失败: {e}")

    def _load_sheet_order(self):
        path = self._sheet_order_path()
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, list):
                    return [str(x) for x in data]
            except Exception:
                pass
        return []

    def _apply_saved_sheet_order(self):
        order = self._load_sheet_order()
        if not order:
            return
        bar = self.tab_widget.tabBar()
        target_pos = 0
        for name in order:
            for i in range(self.tab_widget.count()):
                if self.tab_widget.tabText(i) == name:
                    if i != target_pos:
                        bar.moveTab(i, target_pos)
                    target_pos += 1
                    break

    def _remove_var_from_config(self, var_name):
        path = self._config_file_path()
        try:
            with open(path, "r", encoding="utf-8-sig") as f:
                source = f.read()
            lines = source.splitlines(keepends=True)
            tree = ast.parse(source)
            target_node = None
            for node in tree.body:
                if isinstance(node, ast.Assign):
                    for t in node.targets:
                        if isinstance(t, ast.Name) and t.id == var_name:
                            target_node = node
                            break
                if target_node:
                    break
            if target_node is None:
                return True
            start = target_node.lineno - 1
            end = getattr(target_node, "end_lineno", target_node.lineno)
            del lines[start:end]
            with open(path, "w", encoding="utf-8") as f:
                f.writelines(lines)
            self._suppress_watch_until = time.time() + 1.5
            return True
        except Exception as e:
            print(f"从 souren_config.py 删除 {var_name} 失败: {e}")
            return False

    def _on_case_item_changed(self, tab_name, item):
        if self._suspend_case_sync:
            return
        if item is None:
            return
        if item.column() == 0 and not (self.thread and self.thread.isRunning()):
            table = self.tables.get(tab_name)
            if table:
                st = table.item(item.row(), 5)
                if st and st.text() in ("", "等待"):
                    st.setText("等待" if item.checkState() == Qt.Checked else "")
            self._update_selected_count()
            return
        if item.column() not in (2, 3):
            return
        if item.column() == 3:
            item.setToolTip(item.text())
            table = self.tables.get(tab_name)
            if table:
                self._fit_param_row_heights(table)
        self._sync_cases_to_config()

    def _populate_table(self, tab_name, cases):
        table = self.tables[tab_name]
        self._suspend_case_sync = True
        table.setRowCount(len(cases))

        saved_states = getattr(self, '_checked_states', {}).get(tab_name)
        for row, case in enumerate(cases):

            check_item = QTableWidgetItem()
            if saved_states is not None and row < len(saved_states):
                check_item.setCheckState(Qt.Checked if saved_states[row] else Qt.Unchecked)
            else:
                check_item.setCheckState(Qt.Checked)
            table.setItem(row, 0, check_item)

            seq_item = QTableWidgetItem(str(row+1))
            seq_item.setTextAlignment(Qt.AlignCenter)
            table.setItem(row, 1, seq_item)

            script = case.get('script', '')
            script_item = QTableWidgetItem(script)
            script_item.setFlags(script_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(row, 2, script_item)

            params = {k:v for k,v in case.items() if k != 'script'}
            saved_text = self._saved_param_text_for(tab_name, row, params)
            param_str = saved_text if saved_text is not None else self._format_params_str(params)
            param_item = QTableWidgetItem(param_str)
            param_item.setToolTip(param_str)
            table.setItem(row, 3, param_item)

            table.setItem(row, 4, QTableWidgetItem("0"))
            wait_txt = "等待" if check_item.checkState() == Qt.Checked else ""
            table.setItem(row, 5, QTableWidgetItem(wait_txt))
            table.setItem(row, 6, QTableWidgetItem(""))

        self._fit_param_row_heights(table)
        self._suspend_case_sync = False
        self._update_selected_count()

    def _fit_param_row_heights(self, table, max_lines=3):
        tab_name = self._tab_name_of(table)
        saved = getattr(self, '_row_heights', {}).get(tab_name, {}) if tab_name else {}
        self._suspend_row_save = True
        try:
            table.setWordWrap(True)
            base = 28
            table.resizeRowsToContents()
            cap = base * max_lines
            for r in range(table.rowCount()):
                # 用户手动设过该行高度 → 保持,不被自动折行覆盖
                if str(r) in saved and int(saved[str(r)]) > 0:
                    table.setRowHeight(r, int(saved[str(r)]))
                    continue
                h = table.rowHeight(r)
                if h < base:
                    table.setRowHeight(r, base)
                elif h > cap:
                    table.setRowHeight(r, cap)
        finally:
            self._suspend_row_save = False

    def _show_case_context_menu(self, tab_name, pos):
        table = self.tables[tab_name]
        row = table.rowAt(pos.y())
        if row < 0:
            return
        menu = QMenu(table)
        act_dup = menu.addAction("📋 复制此case到下一行")
        act_copy = menu.addAction("⧉ 复制选中内容")
        action = menu.exec_(table.viewport().mapToGlobal(pos))
        if action == act_dup:
            self._duplicate_case_row(tab_name, row)
        elif action == act_copy:
            table._copy_selection()

    def _duplicate_case_row(self, tab_name, src_row):
        table = self.tables[tab_name]
        self._suspend_case_sync = True
        dst = src_row + 1
        table.insertRow(dst)

        chk = QTableWidgetItem()
        src_chk = table.item(src_row, 0)
        chk.setCheckState(src_chk.checkState() if src_chk else Qt.Checked)
        table.setItem(dst, 0, chk)

        seq = QTableWidgetItem("")
        seq.setTextAlignment(Qt.AlignCenter)
        table.setItem(dst, 1, seq)

        src_script = table.item(src_row, 2)
        sc = QTableWidgetItem(src_script.text() if src_script else "")
        sc.setFlags(sc.flags() & ~Qt.ItemIsEditable)
        table.setItem(dst, 2, sc)

        src_param = table.item(src_row, 3)
        p = QTableWidgetItem(src_param.text() if src_param else "{}")
        p.setToolTip(p.text())
        table.setItem(dst, 3, p)

        table.setItem(dst, 4, QTableWidgetItem("0"))
        table.setItem(dst, 5, QTableWidgetItem("等待"))
        table.setItem(dst, 6, QTableWidgetItem(""))
        self._fit_param_row_heights(table)
        self._suspend_case_sync = False
        self._update_all_row_numbers(tab_name)
        self._sync_cases_to_config()
        table.selectRow(dst)

    def add_case_row(self, tab_name):
        table = self.tables[tab_name]
        self._suspend_case_sync = True
        row = table.rowCount()
        table.insertRow(row)
        check_item = QTableWidgetItem()
        check_item.setCheckState(Qt.Checked)
        table.setItem(row, 0, check_item)
        seq_item = QTableWidgetItem(str(row+1))
        seq_item.setTextAlignment(Qt.AlignCenter)
        table.setItem(row, 1, seq_item)
        table.setItem(row, 2, QTableWidgetItem("new_script"))
        table.setItem(row, 3, QTableWidgetItem("{}"))
        table.setItem(row, 4, QTableWidgetItem("0"))
        table.setItem(row, 5, QTableWidgetItem("等待"))
        table.setItem(row, 6, QTableWidgetItem(""))
        self._suspend_case_sync = False
        self._fit_param_row_heights(table)
        self._update_all_row_numbers(tab_name)
        self._sync_cases_to_config()

    def delete_selected_rows(self, tab_name):
        table = self.tables[tab_name]
        self._suspend_case_sync = True
        selected_rows = sorted(set(index.row() for index in table.selectedIndexes()), reverse=True)
        for row in selected_rows:
            table.removeRow(row)
        self._suspend_case_sync = False
        self._update_all_row_numbers(tab_name)
        self._sync_cases_to_config()

    def move_row_up(self, tab_name):
        table = self.tables[tab_name]
        self._suspend_case_sync = True
        row = table.currentRow()
        if row > 0:
            table.insertRow(row-1)
            for col in range(table.columnCount()):
                item = table.takeItem(row+1, col)
                if item:
                    table.setItem(row-1, col, item)
            table.removeRow(row+1)
            table.selectRow(row-1)
            self._update_all_row_numbers(tab_name)
        self._suspend_case_sync = False
        self._sync_cases_to_config()

    def move_row_down(self, tab_name):
        table = self.tables[tab_name]
        self._suspend_case_sync = True
        row = table.currentRow()
        if row < table.rowCount()-1:
            table.insertRow(row+2)
            for col in range(table.columnCount()):
                item = table.takeItem(row, col)
                if item:
                    table.setItem(row+2, col, item)
            table.removeRow(row)
            table.selectRow(row+1)
            self._update_all_row_numbers(tab_name)
        self._suspend_case_sync = False
        self._sync_cases_to_config()

    def _update_all_row_numbers(self, tab_name):
        table = self.tables[tab_name]
        for row in range(table.rowCount()):
            item = table.item(row, 1)
            if item:
                item.setText(str(row+1))

    def check_all_rows(self, tab_name):
        table = self.tables[tab_name]

        all_checked = True
        for row in range(table.rowCount()):
            item = table.item(row, 0)
            if item and item.checkState() != Qt.Checked:
                all_checked = False
                break
        new_state = Qt.Unchecked if all_checked else Qt.Checked
        for row in range(table.rowCount()):
            item = table.item(row, 0)
            if item:
                item.setCheckState(new_state)
        self._update_selected_count()
        self.status_label.setText("已取消全部勾选" if new_state == Qt.Unchecked else "已全部勾选")

    def add_skip_script_row(self):
        name, ok = QInputDialog.getText(self, "添加跳过脚本", "请添加脚本名:")
        if ok and name.strip():
            name = name.strip()

            for r in range(self.skip_scripts_table.rowCount()):
                item = self.skip_scripts_table.item(r, 0)
                if item and item.text().strip() == name:
                    QMessageBox.warning(self, "重复", f"'{name}' 已存在")
                    return
            row = self.skip_scripts_table.rowCount()
            self.skip_scripts_table.insertRow(row)
            self.skip_scripts_table.setItem(row, 0, QTableWidgetItem(name))
            self._sync_skip_scripts_to_config()

    def delete_skip_script_row(self):
        rows = sorted(set(idx.row() for idx in self.skip_scripts_table.selectedIndexes()), reverse=True)
        if not rows:
            row = self.skip_scripts_table.currentRow()
            if row >= 0:
                rows = [row]
        for row in rows:
            self.skip_scripts_table.removeRow(row)
        self._sync_skip_scripts_to_config()

    def _on_loop_count_changed(self, value):
        try:
            souren_config.LOOP_COUNT = int(value)
            self._persist_var_to_config("LOOP_COUNT", repr(int(value)))
            self.status_label.setText(f"循环次数已设为 {value}")
        except Exception:
            pass

    def refresh_checkboxes(self):
        for tab_name, table in self.tables.items():
            for row in range(table.rowCount()):
                check_item = table.item(row, 0)
                if check_item:
                    check_item.setCheckState(Qt.Unchecked)
        self._update_selected_count()
        self.status_label.setText("已取消所有勾选")

    def _update_selected_count(self):
        """刷新“已勾选 X / 共 Y”标签(统计当前标签页)。"""
        if not hasattr(self, 'selected_count_label'):
            return
        table = self.tables.get(self._current_tab_name()) if hasattr(self, 'tables') else None
        if not table:
            self.selected_count_label.setText("已勾选 0 / 共 0")
            return
        total = table.rowCount()
        checked = sum(
            1 for r in range(total)
            if table.item(r, 0) and table.item(r, 0).checkState() == Qt.Checked
        )
        self.selected_count_label.setText(f"已勾选 {checked} / 共 {total}")

    def start_execution(self):

        self.sync_all_to_config()
        current_idx = self.tab_widget.currentIndex()
        tab_name = self.tab_widget.tabText(current_idx)
        table = self.tables[tab_name]

        cases = []
        for row in range(table.rowCount()):
            check_item = table.item(row, 0)
            if check_item and check_item.checkState() == Qt.Checked:
                script_item = table.item(row, 2)
                param_item = table.item(row, 3)
                if not script_item or not param_item:
                    continue
                script = script_item.text().strip()
                param_str = param_item.text().strip()
                try:
                    params = self._parse_params(param_str)
                except Exception:
                    QMessageBox.warning(self, "参数错误", f"第{row+1}行参数不是合法参数(JSON/Python字面量)")
                    return
                script_path = self._find_script_file(script, params.get('case_dir'))
                if not script_path:
                    QMessageBox.warning(self, "文件不存在", f"脚本 {script} 未找到")
                    return
                cases.append({
                    'script_path': script_path,
                    'params': params,
                    'row': row,
                    'tab': tab_name
                })

        if not cases:
            QMessageBox.information(self, "提示", "当前标签页没有勾选的用例")
            return

        ip = self.ip_edit.text().strip()
        if ip and not self._ping_host(ip):
            reply = QMessageBox.warning(
                self, "仪器无法连通",
                f"无法 ping 通仪器 IP: {ip}\n\n"
                "可能原因：仪器未开机 / 网线未插 / IP 填写错误 / 不在同一网段。\n\n"
                "仍要继续执行吗？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        self._set_ui_enabled(False)
        self.progress.setValue(0)
        self.status_label.setText("执行中...")
        self._exec_errors = []
        if hasattr(self, 'exec_log'):
            self._clear_exec_log()

        checked_rows = {c['row'] for c in cases}
        for r in range(table.rowCount()):
            chk = table.item(r, 0)
            is_checked = bool(chk and chk.checkState() == Qt.Checked)
            status_val = "等待" if is_checked else ""
            for col, val in ((4, "0"), (5, status_val), (6, "")):
                it = table.item(r, col)
                if it:
                    it.setText(val)
                    it.setBackground(QColor(255, 255, 255))
                    it.setToolTip("")

        retry_rounds = self.retry_spin.value() if self.retry_check.isChecked() else 0
        self._retry_results = {}
        self.thread = ExecuteThread(cases, retry_rounds=retry_rounds)
        self.thread.progress_update.connect(self._on_progress_update)
        self.thread.finished_signal.connect(self._on_execution_finished)
        self.thread.log_line.connect(self._on_exec_log_line)
        self.thread.retry_result.connect(self._on_retry_result)
        self.thread.start()

    def _on_retry_result(self, row, rate):
        if not hasattr(self, '_retry_results'):
            self._retry_results = {}
        self._retry_results[row] = rate

    def stop_execution(self):
        if self.thread and self.thread.isRunning():
            self.status_label.setText("正在停止（中断当前用例并保存日志）...")
            self.stop_btn.setEnabled(False)
            self.thread.stop()

    def _clear_exec_log(self):
        self._log_buffer = []
        if hasattr(self, 'exec_log'):
            self.exec_log.clear()

    def _on_exec_log_line(self, text):
        if not hasattr(self, '_log_buffer'):
            self._log_buffer = []
        self._log_buffer.append(text)
        # 现在每条是聚合后的大块(~4KB), 按总字符数封顶, 防止 UI 忙时积压过大
        if len(self._log_buffer) > 2000:
            del self._log_buffer[:-1000]

    def _flush_exec_log(self):
        if not getattr(self, '_log_buffer', None):
            return
        if not hasattr(self, 'exec_log'):
            self._log_buffer = []
            return
        chunk = "".join(self._log_buffer)
        self._log_buffer = []
        # 单次刷入过大(UI 卡顿后积压)会让 insertText 本身卡 UI:
        # 控件只留 5000 行, 超过 ~400KB 的积压直接丢头保尾
        MAX_CHUNK = 400 * 1024
        if len(chunk) > MAX_CHUNK:
            chunk = "...(日志过快,中间部分已省略,完整日志见执行目录)...\n" + chunk[-MAX_CHUNK:]

        from PyQt5.QtGui import QTextCursor
        sb = self.exec_log.verticalScrollBar()
        at_bottom = sb.value() >= sb.maximum() - 4
        cursor = self.exec_log.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertText(chunk)
        if at_bottom:
            self.exec_log.setTextCursor(cursor)
            sb.setValue(sb.maximum())

    def _on_progress_update(self, current, total, row_status_dict):
        self.progress.setMaximum(total)
        current_tab = self.tab_widget.tabText(self.tab_widget.currentIndex())
        table = self.tables[current_tab]
        for row, status in row_status_dict.items():
            s = status.get('status')

            if s in ('completed', 'failed', 'error'):
                self.progress.setValue(min(current + 1, total))

            prog_item = table.item(row, 4)
            if prog_item:
                prog_item.setText(str(status.get('progress', 0)))

            status_item = table.item(row, 5)
            if status_item:
                status_item.setText(status.get('status', ''))

            rate_item = table.item(row, 6)
            rate = status.get('success_rate')
            if rate is not None:
                rate_item.setText(f"{rate:.2f}")

                if rate >= 100:
                    rate_item.setBackground(QColor(144, 238, 144))
                elif rate >= 90:
                    rate_item.setBackground(QColor(255, 255, 153))
                else:
                    rate_item.setBackground(QColor(255, 182, 193))
            elif status.get('status') == 'completed':
                rate_item.setText("无数据")

            if s == 'completed':
                status_item.setBackground(QColor(144, 238, 144))
            elif s == 'failed':
                status_item.setBackground(QColor(255, 182, 193))
            elif s == 'running':
                status_item.setBackground(QColor(255, 255, 153))
            elif s == 'stopped':
                status_item.setBackground(QColor(255, 204, 128))

            msg = status.get('msg')
            if s in ('failed', 'error') and msg:
                if not hasattr(self, '_exec_errors'):
                    self._exec_errors = []
                self._exec_errors.append(f"第{row+1}行: {msg}")
                if status_item:
                    status_item.setToolTip(str(msg))

    def _on_execution_finished(self, success, message):
        self._set_ui_enabled(True)
        self.progress.setValue(self.progress.maximum())
        self.status_label.setText(message)
        errors = getattr(self, '_exec_errors', [])
        if errors:
            detail = "\n\n".join(errors)

            is_conn = any(("仪器连接失败" in e or "连接失败" in e or "Timeout" in e or "VI_ERROR" in e)
                          for e in errors)
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Warning)
            box.setWindowTitle("执行错误")
            if is_conn:
                cur_ip = self.ip_edit.text().strip()
                box.setText(
                    "❌ 仪器连接失败，用例未能执行。\n\n"
                    f"当前仪器 IP: {cur_ip}\n"
                    "请依次检查：\n"
                    "  1) 仪器/基站是否已开机\n"
                    "  2) 网线是否插好、与电脑是否同一网段\n"
                    "  3) 上方“全局配置”里的 IP 是否与仪器实际 IP 一致\n"
                    f"  4) 能否 ping 通该 IP（命令行执行 ping {cur_ip or '<IP>'}）\n"
                    "  5) 是否已安装 NI-VISA 运行库\n\n"
                    "点击“Show Details”查看原始错误。"
                )
            else:
                box.setText("部分用例执行失败，错误信息如下（点击 Show Details 查看详情）：")
            box.setDetailedText(detail)
            box.exec_()
        self._exec_errors = []

    def _set_ui_enabled(self, enabled):
        self.exec_btn.setEnabled(enabled)
        for btn in getattr(self, 'tab_action_buttons', []):
            btn.setEnabled(enabled)
        self.stop_btn.setEnabled(not enabled)
        for table in self.tables.values():
            table.setEditTriggers(QTableWidget.DoubleClicked if enabled else QTableWidget.NoEditTriggers)

    def _ping_host(self, ip, timeout_ms=1000):
        import subprocess
        try:

            if sys.platform.startswith("win"):
                cmd = ["ping", "-n", "1", "-w", str(timeout_ms), ip]
                creationflags = 0x08000000
            else:
                cmd = ["ping", "-c", "1", "-W", str(max(1, timeout_ms // 1000)), ip]
                creationflags = 0
            result = subprocess.run(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                timeout=max(2, timeout_ms / 1000 + 1), creationflags=creationflags
            )
            return result.returncode == 0
        except Exception:
            return False

    def _find_script_file(self, script_name, case_dir=None):
        if not script_name.endswith('.py'):
            script_name_py = script_name + '.py'
        else:
            script_name_py = script_name

        if case_dir:
            try:
                case_dir_path = souren_config.get_case_directory(case_dir)
                candidate = os.path.join(case_dir_path, script_name_py)
                if os.path.exists(candidate):
                    return candidate
            except Exception:
                pass

        search_dirs = [project_root(), os.getcwd(),
                       os.path.dirname(os.path.abspath(__file__))]
        for base in search_dirs:
            path = os.path.join(base, script_name_py)
            if os.path.exists(path):
                return path
            if case_dir:
                path2 = os.path.join(base, case_dir, script_name_py)
                if os.path.exists(path2):
                    return path2
        return None

    def _update_global_config(self):
        souren_config.DEFAULT_IP = self.ip_edit.text()
        souren_config.REMOTE_SERVER_PORT = int(self.port_edit.text())
        souren_config.REMOTE_SUDO_PASSWORD = self.password_edit.text()
        souren_config.VERSION = self.version_edit.text()
        souren_config.COLLECT_CORE_NETWORK_LOGS = self.log_collect_check.isChecked()
        if hasattr(self, 'debug_log_check'):
            souren_config.DEBUG_LOG_ENABLED = self.debug_log_check.isChecked()
        if hasattr(self, 'loop_count_spin'):
            souren_config.LOOP_COUNT = int(self.loop_count_spin.value())

        def _parse_val(val_str):
            val_str = val_str.strip()
            try:
                if val_str.lstrip('-').isdigit():
                    return int(val_str)
                return float(val_str)
            except:
                return val_str

        for row in range(self.log_level_table.rowCount()):
            for kcol, vcol in ((0, 1), (2, 3)):
                key_item = self.log_level_table.item(row, kcol)
                val_item = self.log_level_table.item(row, vcol)
                if key_item and val_item and key_item.text().strip():
                    souren_config.LOG_LEVEL_PARAMS[key_item.text().strip()] = _parse_val(val_item.text())
        skip_scripts = []
        for row in range(self.skip_scripts_table.rowCount()):
            item = self.skip_scripts_table.item(row, 0)
            if item and item.text().strip():
                skip_scripts.append(item.text().strip())
        souren_config.LTE_MEASUREMENT_REPORT_CHECK_SKIP_SCRIPTS = skip_scripts
        souren_config.INSTRUMENT_ADDRESS = f"TCPIP0::{souren_config.DEFAULT_IP}::inst0::INSTR"
        souren_config.VERSION_DIR_SUFFIX = souren_config._version_dir_suffix()

    def open_log_folder(self):
        app_dir = project_root()
        log_path = os.path.join(app_dir, "log")
        if not os.path.exists(log_path):
            os.makedirs(log_path, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(log_path))

    def test_instrument_connection(self):
        if getattr(self, '_conn_thread', None) and self._conn_thread.isRunning():
            return

        try:
            souren_config.DEFAULT_IP = self.ip_edit.text().strip()
            souren_config.INSTRUMENT_ADDRESS = f"TCPIP0::{souren_config.DEFAULT_IP}::inst0::INSTR"
        except Exception:
            pass
        ip = self.ip_edit.text().strip()

        if ip and not self._ping_host(ip):
            reply = QMessageBox.warning(
                self, "仪器无法连通",
                f"无法 ping 通仪器 IP: {ip}\n\n"
                "可能：仪器未开机 / 网线未插 / IP 错误 / 不在同一网段。\n\n"
                "仍要尝试 VISA 连接吗？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        self.test_conn_btn.setEnabled(False)
        self.test_conn_btn.setText("⏳ 连接中...")
        self.status_label.setText(f"正在测试连接 {ip} ...")

        self._conn_thread = ConnectionTestThread()
        self._conn_thread.result_signal.connect(self._on_conn_test_result)
        self._conn_thread.start()

    def query_instrument_idn(self):
        if getattr(self, '_idn_thread', None) and self._idn_thread.isRunning():
            return
        try:
            souren_config.DEFAULT_IP = self.ip_edit.text().strip()
            souren_config.INSTRUMENT_ADDRESS = f"TCPIP0::{souren_config.DEFAULT_IP}::inst0::INSTR"
        except Exception:
            pass
        self.idn_btn.setEnabled(False)
        self.idn_btn.setText("⏳ 获取中...")
        self.idn_result.setText("")
        self._idn_thread = ConnectionTestThread()
        self._idn_thread.result_signal.connect(self._on_idn_result)
        self._idn_thread.start()

    def _on_idn_result(self, success, message):
        self.idn_btn.setEnabled(True)
        self.idn_btn.setText("🔎 获取仪表版本/SN")
        if success:
            idn = message.split(":", 1)[1].strip() if ":" in message else message.strip()
            self.idn_result.setText(idn)
            self.idn_result.setToolTip(idn)
            self.status_label.setText("已获取仪表 *IDN?")
        else:
            self.idn_result.setText("获取失败")
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Warning)
            box.setWindowTitle("获取失败")
            box.setText("*IDN? 获取失败")
            box.setDetailedText(str(message))
            box.exec_()

    def _on_conn_test_result(self, success, message):
        self.test_conn_btn.setEnabled(True)
        self.test_conn_btn.setText("🔌 测试连接")
        if success:
            self.status_label.setText("✅ 仪器连接成功")
            QMessageBox.information(self, "连接成功", f"仪器连接成功！\n\n{message}")
        else:
            self.status_label.setText("❌ 仪器连接失败")
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Critical)
            box.setWindowTitle("连接失败")
            box.setText("❌ 仪器连接失败\n\n点击“Show Details”查看原始错误。")
            box.setDetailedText(str(message))
            box.exec_()

    def closeEvent(self, event):
        try:
            self.sync_all_to_config()
        except Exception:
            pass
        try:
            self._save_checked_states()
        except Exception:
            pass
        try:
            self._save_column_widths()
        except Exception:
            pass
        try:
            self._save_param_texts()
        except Exception:
            pass
        try:
            self._save_cmd_history()
        except Exception:
            pass
        try:
            self._save_retry_settings()
        except Exception:
            pass
        super().closeEvent(event)

    def run_phone_setup(self):
        try:
            controller = ADBFlightModeController()
            if controller.device_id:
                reply = QMessageBox.question(
                    self, "确认", "将开启飞行模式并等待5秒后关闭,是否继续？",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.Yes:
                    controller.timed_flight_mode_control(5)
                    QMessageBox.information(self, "完成", "手机飞行模式控制执行完毕")
            else:
                QMessageBox.warning(self, "错误", "未连接手机设备,请检查USB调试")
        except Exception as e:
            QMessageBox.critical(self, "异常", f"执行失败: {e}")

if __name__ == "__main__":

    from PyQt5.QtWidgets import QInputDialog
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
