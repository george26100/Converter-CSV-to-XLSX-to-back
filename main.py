import sys
import time
from pathlib import Path

import pandas as pd
from PyQt6.QtCore import (
    Qt, QAbstractTableModel, QModelIndex,
    QPropertyAnimation, QEasingCurve, QTimer, pyqtSignal,
)
from PyQt6.QtGui import QAction, QKeySequence, QFont, QPalette, QColor
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableView, QPushButton, QFileDialog, QComboBox, QLabel,
    QSplitter, QTabWidget, QMessageBox, QProgressDialog,
    QAbstractItemView, QGroupBox, QStatusBar, QFrame,
    QGraphicsOpacityEffect, QHeaderView, QListWidget,
    QScrollArea,
)


# =============================================================
#                     МОДЕЛЬ ТАБЛИЦЫ (pandas)
# =============================================================
class PandasModel(QAbstractTableModel):
    """Модель QTableView поверх pandas.DataFrame.

    - Сортировка по клику на заголовок делегируется в MainWindow
      (чтобы синхронизировать _raw_df + фильтры).
    - Редактирование ячеек тоже делегируется через сигнал
      cell_edit_requested — модель не пишет в _df напрямую, чтобы
      MainWindow успел запушить историю и применить фильтры.
    """

    header_sort_requested = pyqtSignal(int, bool)      # (column, ascending)
    cell_edit_requested = pyqtSignal(int, int, str)    # (row, col, text)

    def __init__(self, df: pd.DataFrame | None = None):
        super().__init__()
        self._df = df if df is not None else pd.DataFrame()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df.columns)

    # ячейки редактируемы
    def flags(self, index):
        base = super().flags(index)
        if not index.isValid():
            return base
        return base | Qt.ItemFlag.ItemIsEditable

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        if role == Qt.ItemDataRole.DisplayRole:
            value = self._df.iat[index.row(), index.column()]
            return "" if pd.isna(value) else str(value)
        if role == Qt.ItemDataRole.TextAlignmentRole:
            value = self._df.iat[index.row(), index.column()]
            if isinstance(value, (int, float)) and not pd.isna(value):
                return int(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            return int(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                return str(self._df.columns[section])
            return str(self._df.index[section])
        if role == Qt.ItemDataRole.FontRole:
            f = QFont()
            f.setBold(True)
            return f
        return None

    def sort(self, column, order):
        self.header_sort_requested.emit(
            column, order == Qt.SortOrder.AscendingOrder
        )

    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if not index.isValid():
            return False
        if role not in (Qt.ItemDataRole.EditRole, Qt.ItemDataRole.DisplayRole):
            return False
        text = "" if value is None else str(value)
        self.cell_edit_requested.emit(index.row(), index.column(), text)
        return True

    def set_dataframe(self, df: pd.DataFrame):
        self.beginResetModel()
        self._df = df
        self.endResetModel()

    def dataframe(self) -> pd.DataFrame:
        return self._df


# =============================================================
#                     ВСПЛЫВАЮЩИЙ ТОСТ
# =============================================================
def show_toast(parent: QWidget, message: str, duration: int = 2200, color: str = "#2ecc71"):
    toast = QLabel(message, parent)
    toast.setStyleSheet(
        f"""
        QLabel {{
            background: {color};
            color: white;
            padding: 10px 20px;
            border-radius: 10px;
            font-weight: bold;
            font-size: 13px;
        }}
        """
    )
    toast.adjustSize()
    toast.move(
        (parent.width() - toast.width()) // 2,
        parent.height() - toast.height() - 70,
    )

    effect = QGraphicsOpacityEffect(toast)
    toast.setGraphicsEffect(effect)
    toast.show()

    fade_in = QPropertyAnimation(effect, b"opacity", toast)
    fade_in.setDuration(280)
    fade_in.setStartValue(0.0)
    fade_in.setEndValue(1.0)
    fade_in.setEasingCurve(QEasingCurve.Type.OutCubic)
    fade_in.start()
    toast._fade_in = fade_in

    def _fade_out():
        fade_out = QPropertyAnimation(effect, b"opacity", toast)
        fade_out.setDuration(280)
        fade_out.setStartValue(1.0)
        fade_out.setEndValue(0.0)
        fade_out.finished.connect(toast.deleteLater)
        fade_out.start()
        toast._fade_out = fade_out

    QTimer.singleShot(duration, _fade_out)


# =============================================================
#                        ГЛАВНОЕ ОКНО
# =============================================================
class MainWindow(QMainWindow):
    OP_EQ = "eq"
    OP_GT = "gt"
    OP_LT = "lt"
    OP_GE = "ge"
    OP_LE = "le"
    OP_BETWEEN = "between"

    def __init__(self):
        super().__init__()
        self.setWindowTitle("📊 CSV ↔ XLSX Конвертор")
        self.resize(1360, 860)
        self.setMinimumSize(1100, 620)

        self._raw_df: pd.DataFrame | None = None
        self.df: pd.DataFrame | None = None
        self._filters: list[dict] = []
        self.history: list[dict] = []
        self.filename: str = "data"

        self._build_ui()
        self._apply_styles()
        self._fade_in_window()
        self._update_stats()

    # ---------- построение интерфейса ----------
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(10)

        title = QLabel("📊 CSV ↔ XLSX Конвертор")
        title.setObjectName("heroTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(title)

        subtitle = QLabel("Загрузите файл → отредактируйте → скачайте XLSX или CSV")
        subtitle.setObjectName("subtitle")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        root.addWidget(subtitle)

        # ---- Статистика ----
        stats_row = QHBoxLayout()
        self.stat_labels = {}
        for key, caption in [
            ("rows", "Строк"),
            ("cols", "Столбцов"),
            ("empty", "Пустых ячеек"),
            ("numeric", "Числовых столбцов"),
        ]:
            card = QFrame()
            card.setObjectName("statCard")
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(14, 10, 14, 10)
            card_layout.setSpacing(2)

            cap = QLabel(caption)
            cap.setObjectName("statCaption")
            val = QLabel("0")
            val.setObjectName("statValue")

            card_layout.addWidget(cap)
            card_layout.addWidget(val)
            stats_row.addWidget(card)
            self.stat_labels[key] = val

        root.addLayout(stats_row)

        # ---- Главная область ----
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # ===== ЛЕВАЯ ЧАСТЬ — ТАБЛИЦА =====
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)

        table_tab = QWidget()
        tlayout = QVBoxLayout(table_tab)
        tlayout.setContentsMargins(0, 6, 0, 0)

        self.table = QTableView()
        self.model = PandasModel()
        self.table.setModel(self.model)
        self.table.setSortingEnabled(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
            | QAbstractItemView.EditTrigger.AnyKeyPressed
        )

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)
        header.setSortIndicatorShown(True)
        header.setSectionsClickable(True)

        pal = header.palette()
        pal.setColor(QPalette.ColorRole.ButtonText, QColor("#ffffff"))
        pal.setColor(QPalette.ColorRole.Text, QColor("#ffffff"))
        header.setPalette(pal)

        self.table.verticalHeader().setDefaultSectionSize(26)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        tlayout.addWidget(self.table)

        self.model.header_sort_requested.connect(self._on_header_sort)
        self.model.cell_edit_requested.connect(self._on_cell_edit)

        hint = QLabel(
            "💡 Двойной клик по ячейке — редактирование. "
            "Клик по заголовку — сортировка по одному столбцу. "
            "Для диапазона — «↕️ Сортировка» справа. "
            "Delete — удалить выделенное."
        )
        hint.setObjectName("hint")
        tlayout.addWidget(hint)

        self.tabs.addTab(table_tab, "📋 Таблица")
        splitter.addWidget(self.tabs)

        # ===== ПРАВАЯ ЧАСТЬ — ПАНЕЛЬ ДЕЙСТВИЙ (в QScrollArea) =====
        side_content = QWidget()
        sl = QVBoxLayout(side_content)
        sl.setContentsMargins(10, 0, 10, 0)
        sl.setSpacing(6)

        self.btn_open = QPushButton("📂 Открыть файл")
        self.btn_open.clicked.connect(self.load_file)
        sl.addWidget(self.btn_open)

        self.btn_export = QPushButton("💾 Экспорт в XLSX")
        self.btn_export.setObjectName("primaryBtn")
        self.btn_export.clicked.connect(self.export_xlsx)
        self.btn_export.setEnabled(False)
        sl.addWidget(self.btn_export)

        self.btn_export_csv = QPushButton("📄 Экспорт в CSV")
        self.btn_export_csv.setObjectName("csvBtn")
        self.btn_export_csv.clicked.connect(self.export_csv)
        self.btn_export_csv.setEnabled(False)
        sl.addWidget(self.btn_export_csv)

        # --- Фильтр ---
        filter_box = QGroupBox("🔍 Фильтр")
        fbl = QVBoxLayout(filter_box)
        fbl.setSpacing(4)
        fbl.setContentsMargins(8, 14, 8, 8)

        fbl.addWidget(QLabel("Столбец:"))
        self.filter_col = QComboBox()
        self.filter_col.currentTextChanged.connect(self._on_filter_col_changed)
        fbl.addWidget(self.filter_col)

        fbl.addWidget(QLabel("Условие:"))
        self.filter_op = QComboBox()
        self.filter_op.addItems([
            "Равно (=)",
            "Больше (>)",
            "Меньше (<)",
            "Не менее (≥)",
            "Не более (≤)",
            "Интервал [от; до]",
        ])
        self.filter_op.currentIndexChanged.connect(self._on_filter_op_changed)
        fbl.addWidget(self.filter_op)

        self.filter_value_label = QLabel("Значение:")
        fbl.addWidget(self.filter_value_label)
        self.filter_value = QComboBox()
        self.filter_value.setEditable(True)
        fbl.addWidget(self.filter_value)

        self.filter_value2_label = QLabel("До:")
        self.filter_value2_label.setVisible(False)
        fbl.addWidget(self.filter_value2_label)
        self.filter_value2 = QComboBox()
        self.filter_value2.setEditable(True)
        self.filter_value2.setVisible(False)
        fbl.addWidget(self.filter_value2)

        btn_add_filter = QPushButton("➕ Применить фильтр")
        btn_add_filter.clicked.connect(self.add_filter)
        fbl.addWidget(btn_add_filter)

        fbl.addWidget(QLabel("Активные фильтры:"))
        self.filters_list = QListWidget()
        self.filters_list.setMinimumHeight(80)
        self.filters_list.setMaximumHeight(150)
        fbl.addWidget(self.filters_list)

        fbtn = QHBoxLayout()
        btn_remove_filter = QPushButton("− Убрать")
        btn_remove_filter.clicked.connect(self.remove_selected_filter)
        fbtn.addWidget(btn_remove_filter)
        btn_clear_filters = QPushButton("✖ Все")
        btn_clear_filters.clicked.connect(self.clear_filters)
        fbtn.addWidget(btn_clear_filters)
        fbl.addLayout(fbtn)

        sl.addWidget(filter_box)

        # --- Сортировка (диапазон) ---
        sort_box = QGroupBox("↕️ Сортировка")
        sbl = QVBoxLayout(sort_box)
        sbl.setSpacing(4)
        sbl.setContentsMargins(8, 14, 8, 8)

        sbl.addWidget(QLabel("От столбца:"))
        self.sort_col_from = QComboBox()
        sbl.addWidget(self.sort_col_from)

        sbl.addWidget(QLabel("До столбца (диапазон):"))
        self.sort_col_to = QComboBox()
        self.sort_col_to.addItem("—")
        sbl.addWidget(self.sort_col_to)

        sbl.addWidget(QLabel("Порядок:"))
        self.sort_order = QComboBox()
        self.sort_order.addItems(["По возрастанию", "По убыванию"])
        sbl.addWidget(self.sort_order)

        btn_sort = QPushButton("Применить сортировку")
        btn_sort.clicked.connect(self.apply_sort)
        sbl.addWidget(btn_sort)

        sl.addWidget(sort_box)

        # --- Удаление ---
        del_box = QGroupBox("🗑️ Удаление")
        dbl = QVBoxLayout(del_box)
        dbl.setSpacing(4)
        dbl.setContentsMargins(8, 14, 8, 8)
        btn_del_rows = QPushButton("Удалить выбранные строки")
        btn_del_rows.clicked.connect(self.delete_rows)
        dbl.addWidget(btn_del_rows)
        btn_del_cols = QPushButton("Удалить выбранные столбцы")
        btn_del_cols.clicked.connect(self.delete_cols)
        dbl.addWidget(btn_del_cols)
        sl.addWidget(del_box)

        # --- История ---
        hist_box = QGroupBox("↩️ История")
        hbl = QVBoxLayout(hist_box)
        hbl.setSpacing(4)
        hbl.setContentsMargins(8, 14, 8, 8)
        self.btn_undo = QPushButton("Отменить последнее действие")
        self.btn_undo.clicked.connect(self.undo)
        self.btn_undo.setEnabled(False)
        hbl.addWidget(self.btn_undo)
        sl.addWidget(hist_box)

        sl.addSpacing(6)

        self.side_scroll = QScrollArea()
        self.side_scroll.setWidgetResizable(True)
        self.side_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.side_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.side_scroll.setWidget(side_content)
        self.side_scroll.setMinimumWidth(300)
        self.side_scroll.setMaximumWidth(400)

        splitter.addWidget(self.side_scroll)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 0)
        splitter.setSizes([1000, 340])

        root.addWidget(splitter, stretch=1)

        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("Готово. Откройте CSV или XLSX файл.")

        for key, slot in [
            ("Ctrl+O", self.load_file),
            ("Ctrl+S", self.export_xlsx),
            ("Ctrl+E", self.export_csv),
            ("Ctrl+Z", self.undo),
        ]:
            act = QAction(self)
            act.setShortcut(QKeySequence(key))
            act.triggered.connect(slot)
            self.addAction(act)

        act_del = QAction(self)
        act_del.setShortcut(QKeySequence("Delete"))
        act_del.triggered.connect(self.delete_selected)
        self.addAction(act_del)

    # ---------- стили ----------
    def _apply_styles(self):
        self.setStyleSheet(
            """
            QMainWindow, QWidget {
                background: #f5f7fa; color: #2c3e50;
                font-family: 'Segoe UI', Arial; font-size: 13px;
            }
            QLabel#heroTitle {
                font-size: 24px; font-weight: 800; color: #000000;
                padding: 2px;
            }
            QLabel#subtitle { color: #7a8a99; margin-bottom: 6px; }
            QLabel#hint { color: #000000; font-style: italic; padding: 4px; }

            QFrame#statCard {
                background: white; border-radius: 10px; border: 1px solid #e0e6ed;
            }
            QLabel#statCaption { color: #7a8a99; font-size: 11px; }
            QLabel#statValue { color: #2c3e50; font-size: 22px; font-weight: bold; }

            QTableView {
                background: white; color: #2c3e50;
                alternate-background-color: #f9fbfd;
                gridline-color: #e6ecf2; border: 1px solid #d0d7de;
                border-radius: 8px;
                selection-background-color: #ffb08a;
                selection-color: #000000;
            }
            QTableView QTableCornerButton::section {
                background: #2c3e50;
                border: none;
            }

            QHeaderView { background-color: #2c3e50; }

            QHeaderView::section {
                background-color: #2c3e50;
                color: white;
                padding: 7px;
                border: none;
                border-right: 1px solid #1a2530;
                font-weight: bold;
                font-size: 12px;
            }
            QHeaderView::section:hover {
                background-color: #3d5068;
            }
            QHeaderView::down-arrow {
                image: none;
                width: 0; height: 0;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 6px solid #ffffff;
                margin-right: 8px;
            }
            QHeaderView::up-arrow {
                image: none;
                width: 0; height: 0;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-bottom: 6px solid #ffffff;
                margin-right: 8px;
            }

            QPushButton {
                background: #ff7f50; color: white; border: none;
                padding: 8px 12px; border-radius: 7px; font-weight: bold;
            }
            QPushButton:hover { background: #ff6b3d; }
            QPushButton:pressed { background: #e55a2b; }
            QPushButton:disabled { background: #c8d2dc; }
            QPushButton#primaryBtn { background: #2ecc71; }
            QPushButton#primaryBtn:hover { background: #27ae60; }
            QPushButton#csvBtn { background: #3498db; }
            QPushButton#csvBtn:hover { background: #2980b9; }
            QPushButton#csvBtn:pressed { background: #21618c; }
            QPushButton#csvBtn:disabled { background: #c8d2dc; }

            QGroupBox {
                font-weight: bold; border: 1px solid #d0d7de; border-radius: 8px;
                margin-top: 10px; padding: 10px 8px 8px 8px;
                background: white; color: #2c3e50;
            }
            QGroupBox::title {
                subcontrol-origin: margin; left: 10px; padding: 0 4px;
                color: #2c3e50;
            }

            QComboBox, QLineEdit {
                padding: 5px; border: 1px solid #d0d7de; border-radius: 5px;
                background: white; color: #2c3e50;
            }
            QComboBox QAbstractItemView {
                background: white; color: #2c3e50;
                selection-background-color: #ffb08a;
                selection-color: #000000;
            }

            QListWidget {
                background: #f9fbfd; color: #2c3e50;
                border: 1px solid #d0d7de; border-radius: 5px;
                padding: 4px;
            }
            QListWidget::item:selected {
                background: #ffb08a; color: #000000;
            }

            QScrollArea {
                background: #f5f7fa;
                border: none;
            }
            QScrollArea > QWidget > QWidget {
                background: #f5f7fa;
            }
            QScrollBar:vertical {
                background: #e0e6ed; width: 10px; margin: 0; border-radius: 5px;
            }
            QScrollBar::handle:vertical {
                background: #b8c4d0; min-height: 30px; border-radius: 5px;
            }
            QScrollBar::handle:vertical:hover { background: #97a7b5; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0; border: none; background: none;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }

            QTabWidget::pane { border: none; }
            QTabBar::tab {
                background: #e0e6ed; padding: 8px 18px; color: #2c3e50;
                border-top-left-radius: 6px; border-top-right-radius: 6px;
                margin-right: 2px;
            }
            QTabBar::tab:selected { background: #ff7f50; color: white; }

            QStatusBar { background: #2c3e50; color: #ecf0f1; }
            """
        )

    # ---------- анимации ----------
    def _fade_in_window(self):
        self.setWindowOpacity(0.0)
        self._win_anim = QPropertyAnimation(self, b"windowOpacity")
        self._win_anim.setDuration(500)
        self._win_anim.setStartValue(0.0)
        self._win_anim.setEndValue(1.0)
        self._win_anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self._win_anim.start()

    def _pulse_label(self, label: QLabel):
        effect = QGraphicsOpacityEffect(label)
        label.setGraphicsEffect(effect)
        a = QPropertyAnimation(effect, b"opacity", label)
        a.setDuration(350)
        a.setStartValue(0.2)
        a.setEndValue(1.0)
        a.setEasingCurve(QEasingCurve.Type.OutQuad)
        a.start()
        label._anim = a

    # ---------- статистика ----------
    def _update_stats(self):
        if self.df is None:
            for v in self.stat_labels.values():
                v.setText("—")
            return
        values = {
            "rows": (
                f"{len(self.df)} / {len(self._raw_df)}"
                if self._filters and self._raw_df is not None
                else f"{len(self.df)}"
            ),
            "cols": f"{len(self.df.columns)}",
            "empty": f"{int(self.df.isna().sum().sum())}",
            "numeric": f"{len(self.df.select_dtypes('number').columns)}",
        }
        for key, val in values.items():
            self.stat_labels[key].setText(val)
            self._pulse_label(self.stat_labels[key])

    def _refresh_column_combos(self):
        cols = list(self._raw_df.columns) if self._raw_df is not None else []

        self.sort_col_from.blockSignals(True)
        self.sort_col_from.clear()
        self.sort_col_from.addItems(cols)
        self.sort_col_from.blockSignals(False)

        self.sort_col_to.blockSignals(True)
        self.sort_col_to.clear()
        self.sort_col_to.addItem("—")
        self.sort_col_to.addItems(cols)
        self.sort_col_to.blockSignals(False)

    @staticmethod
    def _renumber_columns(df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df.columns = [str(i + 1) for i in range(len(df.columns))]
        return df

    # ---------- редактирование ячеек ----------
    @staticmethod
    def _coerce_value(original, text: str):
        """Сохраняет тип ячейки при правке:
        - оригинал int/float → пробуем int, потом float, иначе строка;
        - пустая строка в числовой → NaN;
        - иначе — строка.
        """
        s = text.strip()
        if s == "":
            if isinstance(original, (int, float)) and not isinstance(original, bool):
                return float("nan")
            return ""

        if isinstance(original, (int, float)) and not isinstance(original, bool):
            try:
                return int(s)
            except ValueError:
                try:
                    return float(s.replace(",", "."))
                except ValueError:
                    return s
        return s

    def _on_cell_edit(self, row: int, col: int, text: str):
        if self._raw_df is None or self.df is None:
            return
        if row < 0 or row >= len(self.df):
            return
        if col < 0 or col >= len(self.df.columns):
            return

        col_name = self.df.columns[col]
        orig_idx = self.df.index[row]

        if col_name not in self._raw_df.columns or orig_idx not in self._raw_df.index:
            return

        old_val = self._raw_df.at[orig_idx, col_name]
        new_val = self._coerce_value(old_val, text)

        # Если фактически ничего не изменилось — выходим
        try:
            old_is_na = pd.isna(old_val)
            new_is_na = (isinstance(new_val, float) and pd.isna(new_val))
            if old_is_na and new_is_na:
                return
            if not old_is_na and not new_is_na and str(old_val) == str(new_val):
                return
        except Exception:
            pass

        self._push_history()
        self._raw_df.at[orig_idx, col_name] = new_val
        self._apply_filters()
        show_toast(
            self,
            f"✏️ Изменено: «{col_name}» → {new_val}",
            color="#3498db",
        )
        self.status.showMessage(f"Изменено значение: {col_name} = {new_val}")

    # ---------- фильтры ----------
    def _refresh_filter_columns(self):
        if self._raw_df is None:
            self.filter_col.blockSignals(True)
            self.filter_col.clear()
            self.filter_col.blockSignals(False)
            for cb in (self.filter_value, self.filter_value2):
                cb.blockSignals(True)
                cb.clear()
                cb.blockSignals(False)
            return
        self.filter_col.blockSignals(True)
        self.filter_col.clear()
        self.filter_col.addItems(list(self._raw_df.columns))
        self.filter_col.blockSignals(False)
        self._on_filter_col_changed()

    def _on_filter_col_changed(self):
        if self._raw_df is None:
            for cb in (self.filter_value, self.filter_value2):
                cb.blockSignals(True)
                cb.clear()
                cb.blockSignals(False)
            return
        col = self.filter_col.currentText()
        if not col or col not in self._raw_df.columns:
            for cb in (self.filter_value, self.filter_value2):
                cb.blockSignals(True)
                cb.clear()
                cb.blockSignals(False)
            return
        prev1 = self.filter_value.currentText()
        prev2 = self.filter_value2.currentText()
        vals = self._raw_df[col].dropna().astype(str).unique().tolist()
        vals.sort()
        for cb, prev in ((self.filter_value, prev1), (self.filter_value2, prev2)):
            cb.blockSignals(True)
            cb.clear()
            cb.addItems(vals)
            if prev in vals:
                cb.setCurrentText(prev)
            cb.blockSignals(False)

    def _on_filter_op_changed(self):
        is_between = self.filter_op.currentIndex() == 5
        self.filter_value2_label.setVisible(is_between)
        self.filter_value2.setVisible(is_between)
        self.filter_value_label.setText("От:" if is_between else "Значение:")

    def add_filter(self):
        if self._raw_df is None:
            QMessageBox.information(self, "Внимание", "Сначала загрузите файл.")
            return
        col = self.filter_col.currentText()
        if not col:
            return

        op_map = {
            0: self.OP_EQ,
            1: self.OP_GT,
            2: self.OP_LT,
            3: self.OP_GE,
            4: self.OP_LE,
            5: self.OP_BETWEEN,
        }
        op = op_map.get(self.filter_op.currentIndex(), self.OP_EQ)

        val1 = self.filter_value.currentText().strip()
        val2 = self.filter_value2.currentText().strip()

        if not val1:
            QMessageBox.information(self, "Внимание", "Введите значение фильтра.")
            return

        if op == self.OP_EQ:
            new_filter = {"col": col, "op": op, "value": val1}

        elif op in (self.OP_GT, self.OP_LT, self.OP_GE, self.OP_LE):
            try:
                v = float(val1.replace(",", "."))
            except ValueError:
                QMessageBox.warning(
                    self, "Ошибка",
                    "Для числового условия значение должно быть числом."
                )
                return
            new_filter = {"col": col, "op": op, "value": v}

        elif op == self.OP_BETWEEN:
            if not val2:
                QMessageBox.information(
                    self, "Внимание",
                    "Для интервала укажите оба значения («От» и «До»)."
                )
                return
            try:
                v1 = float(val1.replace(",", "."))
                v2 = float(val2.replace(",", "."))
            except ValueError:
                QMessageBox.warning(
                    self, "Ошибка",
                    "Для интервала оба значения должны быть числами."
                )
                return
            if v1 > v2:
                v1, v2 = v2, v1
            new_filter = {"col": col, "op": op, "value": v1, "value2": v2}
        else:
            return

        for f in self._filters:
            if (
                f["col"] == new_filter["col"]
                and f.get("op") == new_filter["op"]
                and f["value"] == new_filter["value"]
                and f.get("value2") == new_filter.get("value2")
            ):
                show_toast(self, "Такой фильтр уже добавлен", color="#e67e22")
                return

        self._push_history()
        self._filters.append(new_filter)
        self._apply_filters()
        self._refresh_filters_list()
        show_toast(self, f"🔍 {self._format_filter(new_filter)}")

    def remove_selected_filter(self):
        row = self.filters_list.currentRow()
        if row < 0 or row >= len(self._filters):
            return
        self._push_history()
        del self._filters[row]
        self._apply_filters()
        self._refresh_filters_list()
        show_toast(self, "Фильтр убран", color="#e67e22")

    def clear_filters(self):
        if not self._filters:
            return
        self._push_history()
        self._filters.clear()
        self._apply_filters()
        self._refresh_filters_list()
        show_toast(self, "Все фильтры сброшены", color="#e67e22")

    @staticmethod
    def _format_filter(f: dict) -> str:
        op = f.get("op", "eq")
        if op == "eq":
            return f"{f['col']} = {f['value']}"
        if op == "gt":
            return f"{f['col']} > {f['value']}"
        if op == "lt":
            return f"{f['col']} < {f['value']}"
        if op == "ge":
            return f"{f['col']} ≥ {f['value']}"
        if op == "le":
            return f"{f['col']} ≤ {f['value']}"
        if op == "between":
            return f"{f['col']} ∈ [{f['value']}; {f['value2']}]"
        return f"{f['col']} ? {f['value']}"

    def _refresh_filters_list(self):
        self.filters_list.clear()
        for f in self._filters:
            self.filters_list.addItem(self._format_filter(f))

    def _apply_filters(self):
        if self._raw_df is None:
            self.df = None
            self.model.set_dataframe(pd.DataFrame())
            self._update_stats()
            return

        df = self._raw_df
        for f in self._filters:
            col = f["col"]
            if col not in df.columns:
                continue
            op = f.get("op", "eq")
            try:
                if op == "eq":
                    df = df[df[col].astype(str) == str(f["value"])]
                elif op in ("gt", "lt", "ge", "le", "between"):
                    numeric = pd.to_numeric(df[col], errors="coerce")
                    if op == "gt":
                        df = df[numeric > f["value"]]
                    elif op == "lt":
                        df = df[numeric < f["value"]]
                    elif op == "ge":
                        df = df[numeric >= f["value"]]
                    elif op == "le":
                        df = df[numeric <= f["value"]]
                    elif op == "between":
                        df = df[(numeric >= f["value"]) & (numeric <= f["value2"])]
            except Exception:
                continue

        self.df = df
        self.model.set_dataframe(df)
        self._update_stats()

    # ---------- загрузка файла ----------
    def load_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл", "", "Таблицы (*.csv *.xlsx *.xls)"
        )
        if not path:
            return

        progress = QProgressDialog("🔄 Чтение файла...", None, 0, 100, self)
        progress.setWindowTitle("Загрузка")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setCancelButton(None)
        progress.setMinimumWidth(360)
        progress.show()

        for i in range(0, 101, 4):
            progress.setValue(i)
            QApplication.processEvents()
            time.sleep(0.005)

        try:
            if path.lower().endswith(".csv"):
                try:
                    df = pd.read_csv(path, encoding="utf-8")
                except UnicodeDecodeError:
                    df = pd.read_csv(path, encoding="cp1251")
            else:
                df = pd.read_excel(path)
        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "Ошибка", f"Не удалось прочитать файл:\n{e}")
            return

        df = self._renumber_columns(df)
        progress.setValue(100)
        progress.close()

        self._raw_df = df.copy()
        self._filters = []
        self.df = self._raw_df
        self.filename = Path(path).stem
        self.history.clear()
        self.model.set_dataframe(self.df)
        self._refresh_column_combos()
        self._refresh_filter_columns()
        self._refresh_filters_list()
        self._update_stats()
        self.btn_export.setEnabled(True)
        self.btn_export_csv.setEnabled(True)
        self.btn_undo.setEnabled(False)
        self.status.showMessage(
            f"Загружено: {Path(path).name} — "
            f"{len(df)} строк × {len(df.columns)} столбцов"
        )
        show_toast(self, f"✅ Загружено: {len(df)} строк × {len(df.columns)}")

    # ---------- история изменений ----------
    def _push_history(self):
        if self._raw_df is not None:
            self.history.append({
                "df": self._raw_df.copy(),
                "filters": [dict(f) for f in self._filters],
            })
            if len(self.history) > 20:
                self.history.pop(0)
            self.btn_undo.setEnabled(True)

    def undo(self):
        if not self.history:
            return
        state = self.history.pop()
        self._raw_df = state["df"]
        self._filters = state["filters"]
        self._apply_filters()
        self._refresh_column_combos()
        self._refresh_filter_columns()
        self._refresh_filters_list()
        self.btn_undo.setEnabled(bool(self.history))
        show_toast(self, "↩️ Изменение отменено", color="#e67e22")

    # ---------- сортировка ----------
    def apply_sort(self):
        if self._raw_df is None or self._raw_df.empty:
            return

        col_from = self.sort_col_from.currentText()
        col_to = self.sort_col_to.currentText()
        ascending = self.sort_order.currentText() == "По возрастанию"

        if not col_from or col_from not in self._raw_df.columns:
            return

        cols = list(self._raw_df.columns)

        if col_to and col_to != "—" and col_to != col_from and col_to in cols:
            idx_from = cols.index(col_from)
            idx_to = cols.index(col_to)
            if idx_from > idx_to:
                idx_from, idx_to = idx_to, idx_from
            sort_cols = cols[idx_from:idx_to + 1]
        else:
            sort_cols = [col_from]

        self._push_history()
        self._raw_df = self._raw_df.sort_values(
            by=sort_cols, ascending=ascending, na_position="last"
        )
        self._apply_filters()

        try:
            idx = cols.index(col_from)
            order = (
                Qt.SortOrder.AscendingOrder if ascending
                else Qt.SortOrder.DescendingOrder
            )
            h = self.table.horizontalHeader()
            h.blockSignals(True)
            h.setSortIndicator(idx, order)
            h.blockSignals(False)
        except Exception:
            pass

        if len(sort_cols) > 1:
            show_toast(
                self,
                f"↕️ Отсортировано по {len(sort_cols)} столбцам: "
                f"«{sort_cols[0]}» … «{sort_cols[-1]}»",
            )
            self.status.showMessage(
                f"↕️ Сортировка по диапазону: {sort_cols[0]} … {sort_cols[-1]}"
            )
        else:
            show_toast(self, f"↕️ Отсортировано по «{col_from}»")

    def _on_header_sort(self, column, ascending):
        if self._raw_df is None or self._raw_df.empty:
            return
        if column < 0 or column >= len(self._raw_df.columns):
            return
        col_name = self._raw_df.columns[column]
        self._push_history()
        self._raw_df = self._raw_df.sort_values(
            by=col_name, ascending=ascending, na_position="last"
        )
        self._apply_filters()

        self.sort_col_from.blockSignals(True)
        self.sort_col_from.setCurrentText(col_name)
        self.sort_col_from.blockSignals(False)

        self.sort_col_to.blockSignals(True)
        self.sort_col_to.setCurrentIndex(0)
        self.sort_col_to.blockSignals(False)

        self.sort_order.setCurrentText(
            "По возрастанию" if ascending else "По убыванию"
        )
        self.status.showMessage(f"↕️ Отсортировано по «{col_name}»")

    # ---------- удаление ----------
    def delete_rows(self):
        if self.df is None:
            return
        rows = sorted({i.row() for i in self.table.selectionModel().selectedIndexes()})
        self._delete_rows_by_index(rows)

    def delete_cols(self):
        if self.df is None:
            return
        cols = sorted({i.column() for i in self.table.selectionModel().selectedIndexes()})
        self._delete_cols_by_index(cols)

    def delete_selected(self):
        if self.df is None:
            return
        indexes = self.table.selectionModel().selectedIndexes()
        if not indexes:
            return
        row_ids = {i.row() for i in indexes}
        col_ids = {i.column() for i in indexes}
        if len(row_ids) < len(col_ids):
            self._delete_rows_by_index(sorted(row_ids))
        else:
            self._delete_cols_by_index(sorted(col_ids))

    def _delete_rows_by_index(self, rows):
        if self._raw_df is None or not rows:
            if self.df is not None and not rows:
                QMessageBox.information(
                    self, "Внимание",
                    "Выберите строки (или ячейки) для удаления."
                )
            return
        if len(rows) >= len(self.df):
            QMessageBox.warning(self, "Ошибка", "Нельзя удалить все строки.")
            return
        self._push_history()
        visible_labels = self.df.index[rows]
        self._raw_df = self._raw_df.drop(visible_labels)
        self._apply_filters()
        show_toast(self, f"🗑️ Удалено строк: {len(rows)}", color="#e74c3c")

    def _delete_cols_by_index(self, cols):
        if self._raw_df is None or not cols:
            if self.df is not None and not cols:
                QMessageBox.information(
                    self, "Внимание",
                    "Выберите столбцы для удаления."
                )
            return
        if len(cols) >= len(self._raw_df.columns):
            QMessageBox.warning(self, "Ошибка", "Нельзя удалить все столбцы.")
            return

        names = [self._raw_df.columns[i] for i in cols]
        old_remaining = [c for c in self._raw_df.columns if c not in names]

        self._push_history()
        self._raw_df = self._raw_df.drop(columns=names)
        self._raw_df = self._renumber_columns(self._raw_df)

        mapping = {old: new for old, new in zip(old_remaining, self._raw_df.columns)}
        self._filters = [
            {**f, "col": mapping[f["col"]]}
            for f in self._filters if f["col"] in mapping
        ]

        self._apply_filters()
        self._refresh_column_combos()
        self._refresh_filter_columns()
        self._refresh_filters_list()
        show_toast(self, f"🗑️ Удалено столбцов: {len(names)}", color="#e74c3c")

    # ---------- экспорт в XLSX ----------
    def export_xlsx(self):
        if self.df is None or self._raw_df is None:
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить XLSX", f"{self.filename}.xlsx", "Excel (*.xlsx)"
        )
        if not path:
            return

        progress = QProgressDialog("🔄 Конвертация в XLSX...", None, 0, 100, self)
        progress.setWindowTitle("Экспорт")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setCancelButton(None)
        progress.setMinimumWidth(360)
        progress.show()

        try:
            for i in range(0, 60, 4):
                progress.setValue(i)
                QApplication.processEvents()
                time.sleep(0.005)

            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                self.df.to_excel(writer, index=False, sheet_name="Data")
                ws = writer.sheets["Data"]
                for col in ws.columns:
                    max_len = max(
                        (len(str(c.value)) for c in col if c.value is not None),
                        default=8,
                    )
                    ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)
                ws.freeze_panes = "A2"

            for i in range(60, 101, 4):
                progress.setValue(i)
                QApplication.processEvents()
                time.sleep(0.005)

            progress.setValue(100)
            progress.close()

            show_toast(self, f"✅ Сохранено: {Path(path).name}")
            self.status.showMessage(f"Экспортировано: {path}")
        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл:\n{e}")

    # ---------- экспорт в CSV ----------
    def export_csv(self):
        if self.df is None or self._raw_df is None:
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить CSV", f"{self.filename}.csv",
            "CSV (разделитель ;) (*.csv);;CSV (разделитель ,) (*.csv)"
        )
        if not path:
            return

        sep = ","
        if "разделитель ," not in path:
            sep = ";"

        for suffix in [";;", ",,", " ;", " ,"]:
            if path.endswith(suffix):
                path = path[: -len(suffix)]

        if not path.lower().endswith(".csv"):
            path += ".csv"

        progress = QProgressDialog("🔄 Конвертация в CSV...", None, 0, 100, self)
        progress.setWindowTitle("Экспорт CSV")
        progress.setWindowModality(Qt.WindowModality.WindowModal)
        progress.setCancelButton(None)
        progress.setMinimumWidth(360)
        progress.show()

        try:
            for i in range(0, 60, 4):
                progress.setValue(i)
                QApplication.processEvents()
                time.sleep(0.005)

            self.df.to_csv(path, index=False, encoding="utf-8-sig", sep=sep)

            for i in range(60, 101, 4):
                progress.setValue(i)
                QApplication.processEvents()
                time.sleep(0.005)

            progress.setValue(100)
            progress.close()

            show_toast(self, f"✅ CSV: {Path(path).name}")
            self.status.showMessage(f"CSV экспортирован: {path}")
        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить CSV:\n{e}")


# =============================================================
#                          ЗАПУСК
# =============================================================
def main():
    app = QApplication(sys.argv)
    app.setApplicationName("CSV ↔ XLSX")
    app.setStyle("Fusion")

    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()