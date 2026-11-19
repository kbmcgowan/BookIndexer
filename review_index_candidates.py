"""Review candidate index terms and save edits and Keep choices to a CSV."""

import argparse
import csv
import os
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from docx.table import Table
from docx.text.paragraph import Paragraph
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt, Signal
from PySide6.QtGui import QColor, QTextCharFormat, QTextCursor, QTextDocument
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableView,
    QLineEdit,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from index_preferences import (
    HIDEABLE_FIELDS,
    PREFERENCES_PATH,
    Preferences,
    load_preferences,
    resolve_preference_path,
    save_preferences,
)


def parse_keep(value: str | None, row_number: int) -> bool:
    normalized = (value or "").strip().casefold()
    if normalized in {"", "false", "0", "no"}:
        return False
    if normalized in {"true", "1", "yes"}:
        return True
    raise ValueError(f"Invalid Keep value on CSV row {row_number}: {value!r}")


def load_candidates(path: Path) -> tuple[list[str], list[dict[str, Any]], bool]:
    with path.open("r", newline="", encoding="utf-8-sig") as csvfile:
        reader = csv.DictReader(csvfile)
        if not reader.fieldnames:
            raise ValueError("The CSV is empty or has no header row.")

        fieldnames = list(reader.fieldnames)
        if any(name is None or not name.strip() for name in fieldnames):
            raise ValueError("The CSV has a blank column name.")
        if len(set(fieldnames)) != len(fieldnames):
            raise ValueError("The CSV has duplicate column names.")
        if "document" not in fieldnames or "candidate_term" not in fieldnames:
            raise ValueError("The CSV must include document and candidate_term columns.")

        keep_column = next(
            (name for name in fieldnames if name.casefold() == "keep"),
            None,
        )
        needs_save = keep_column is None
        if keep_column is None:
            fieldnames.append("Keep")
        elif keep_column != "Keep":
            fieldnames[fieldnames.index(keep_column)] = "Keep"

        source_term_column = next(
            (name for name in fieldnames if name.casefold() == "source_candidate_term"),
            None,
        )
        source_term_input_column = source_term_column
        if source_term_column is None:
            fieldnames.append("source_candidate_term")
            source_term_column = "source_candidate_term"
            needs_save = True
        elif source_term_column != "source_candidate_term":
            fieldnames[fieldnames.index(source_term_column)] = "source_candidate_term"
            source_term_column = "source_candidate_term"

        rows = []
        for row_number, row in enumerate(reader, start=2):
            if None in row:
                raise ValueError(f"CSV row {row_number} has extra values.")
            if keep_column is not None and keep_column != "Keep":
                row["Keep"] = row.pop(keep_column, None)
            if (
                source_term_input_column is not None
                and source_term_input_column in row
                and source_term_input_column != "source_candidate_term"
            ):
                row["source_candidate_term"] = row.pop(source_term_input_column, None)

            for name in fieldnames:
                row.setdefault(name, "")

            document = row.get("document") or ""
            term = row.get("candidate_term") or ""
            if not document.strip() or not term.strip():
                raise ValueError(
                    f"CSV row {row_number} needs both document and candidate_term."
                )
            row["document"] = document
            row["candidate_term"] = term
            source_term = row.get("source_candidate_term")
            if not isinstance(source_term, str) or not source_term.strip():
                row["source_candidate_term"] = term
                needs_save = True
            raw_keep = row.get("Keep")
            if raw_keep is None or not str(raw_keep).strip():
                needs_save = True
            row["Keep"] = parse_keep(raw_keep, row_number)
            rows.append(row)

    return fieldnames, rows, needs_save


def save_candidates(
    path: Path,
    fieldnames: list[str],
    rows: list[dict[str, Any]],
) -> None:
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            newline="",
            encoding="utf-8-sig",
            dir=path.parent,
            prefix=f".{path.name}.",
            suffix=".tmp",
            delete=False,
        ) as csvfile:
            temporary_path = Path(csvfile.name)
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
        os.replace(temporary_path, path)
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()


def extract_docx_text(path: Path) -> str:
    document = Document(path)
    parts = []
    for block in document.iter_inner_content():
        if isinstance(block, Paragraph):
            parts.append(block.text)
        elif isinstance(block, Table):
            parts.extend(
                "\t".join(cell.text.replace("\n", " ") for cell in row.cells)
                for row in block.rows
            )
    return "\n\n".join(parts)


class PreferencesDialog(QDialog):
    def __init__(self, preferences: Preferences, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Reviewer Preferences")
        self.preferences = preferences

        self.index_path_edit = QLineEdit(str(preferences.index_path))
        self.book_path_edit = QLineEdit(str(preferences.book_path))
        form = QFormLayout()
        form.addRow("Index folder:", self._path_control(self.index_path_edit))
        form.addRow("Book folder:", self._path_control(self.book_path_edit))

        self.hidden_checkboxes = {}
        hidden_fields_panel = QWidget()
        hidden_fields_layout = QVBoxLayout(hidden_fields_panel)
        hidden_fields_layout.setContentsMargins(0, 0, 0, 0)
        for field in HIDEABLE_FIELDS:
            checkbox = QCheckBox(field)
            checkbox.setChecked(field in preferences.hidden_fields)
            hidden_fields_layout.addWidget(checkbox)
            self.hidden_checkboxes[field] = checkbox
        form.addRow("Hide columns:", hidden_fields_panel)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._save_and_accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(buttons)

    def _path_control(self, edit: QLineEdit) -> QWidget:
        control = QWidget()
        layout = QHBoxLayout(control)
        layout.setContentsMargins(0, 0, 0, 0)
        browse_button = QPushButton("Browse...")
        browse_button.clicked.connect(lambda: self._browse_for_directory(edit))
        layout.addWidget(edit, 1)
        layout.addWidget(browse_button)
        return control

    def _browse_for_directory(self, edit: QLineEdit) -> None:
        initial_path = edit.text().strip() or str(PREFERENCES_PATH.parent)
        selected_path = QFileDialog.getExistingDirectory(
            self,
            "Choose folder",
            initial_path,
        )
        if selected_path:
            edit.setText(selected_path)

    def _save_and_accept(self) -> None:
        index_text = self.index_path_edit.text().strip()
        book_text = self.book_path_edit.text().strip()
        if not index_text or not book_text:
            QMessageBox.warning(self, "Invalid preferences", "Folder paths cannot be blank.")
            return

        self.preferences = Preferences(
            index_path=resolve_preference_path(index_text),
            book_path=resolve_preference_path(book_text),
            hidden_fields=tuple(
                field
                for field in HIDEABLE_FIELDS
                if self.hidden_checkboxes[field].isChecked()
            ),
        )
        self.accept()


class CandidateTableModel(QAbstractTableModel):
    keep_changed = Signal()
    candidate_changed = Signal(int)

    def __init__(
        self,
        fieldnames: list[str],
        rows: list[dict[str, Any]],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.fieldnames = fieldnames
        self.rows = rows
        self.keep_column = self.fieldnames.index("Keep")
        self.candidate_column = self.fieldnames.index("candidate_term")

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self.fieldnames)

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if not index.isValid():
            return None

        field = self.fieldnames[index.column()]
        value = self.rows[index.row()].get(field, "")

        if field == "Keep" and role == Qt.ItemDataRole.CheckStateRole:
            return (
                Qt.CheckState.Checked
                if value
                else Qt.CheckState.Unchecked
            )
        if role == Qt.ItemDataRole.DisplayRole and field != "Keep":
            return "" if value is None else str(value)
        if role == Qt.ItemDataRole.EditRole and field == "candidate_term":
            return "" if value is None else str(value)
        if role == Qt.ItemDataRole.TextAlignmentRole and field in {
            "rank",
            "occurrences",
            "tfidf_score",
        }:
            return int(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        if role == Qt.ItemDataRole.ToolTipRole and field == "candidate_term":
            return str(value)
        return None

    def headerData(
        self,
        section: int,
        orientation: Qt.Orientation,
        role: int = Qt.ItemDataRole.DisplayRole,
    ) -> Any:
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            return self.fieldnames[section]
        return section + 1

    def flags(self, index: QModelIndex) -> Qt.ItemFlag:
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        flags = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
        if index.column() == self.keep_column:
            flags |= Qt.ItemFlag.ItemIsUserCheckable
        if index.column() == self.candidate_column:
            flags |= Qt.ItemFlag.ItemIsEditable
        return flags

    def setData(
        self,
        index: QModelIndex,
        value: Any,
        role: int = Qt.ItemDataRole.EditRole,
    ) -> bool:
        if (
            index.isValid()
            and index.column() == self.candidate_column
            and role == Qt.ItemDataRole.EditRole
        ):
            term = str(value).strip()
            if not term:
                return False
            row = self.rows[index.row()]
            if row["candidate_term"] == term:
                return True
            row["candidate_term"] = term
            self.dataChanged.emit(
                index,
                index,
                [Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.ToolTipRole],
            )
            self.candidate_changed.emit(index.row())
            return True

        if (
            not index.isValid()
            or index.column() != self.keep_column
            or role != Qt.ItemDataRole.CheckStateRole
        ):
            return False

        try:
            check_state = Qt.CheckState(value)
        except (TypeError, ValueError):
            return False

        if check_state == Qt.CheckState.Checked:
            keep = True
        elif check_state == Qt.CheckState.Unchecked:
            keep = False
        else:
            return False

        row = self.rows[index.row()]
        if row["Keep"] == keep:
            return True
        row["Keep"] = keep
        self.dataChanged.emit(index, index, [Qt.ItemDataRole.CheckStateRole])
        self.keep_changed.emit()
        return True


class CandidateReviewWindow(QMainWindow):
    def __init__(self, csv_path: Path, preferences: Preferences) -> None:
        super().__init__()
        self.csv_path = csv_path.resolve()
        self.preferences = preferences
        self.document_cache: dict[Path, str] = {}
        self.model: CandidateTableModel | None = None
        self.fieldnames: list[str] = []
        self.rows: list[dict[str, Any]] = []
        self.dirty = False

        self.setWindowTitle("Index Candidate Review")
        self.resize(1440, 900)
        self._build_ui()
        self._load_csv(self.csv_path)

    def _build_ui(self) -> None:
        open_button = QPushButton("Open CSV...")
        open_button.clicked.connect(self._choose_csv)
        preferences_button = QPushButton("Preferences...")
        preferences_button.clicked.connect(self._edit_preferences)
        self.save_button = QPushButton("Save")
        self.save_button.clicked.connect(self._save)
        self.save_button.setEnabled(False)
        self.path_label = QLabel()
        self.path_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )

        toolbar = QWidget()
        toolbar_layout = QVBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(0, 0, 0, 0)
        button_row = QWidget()
        button_layout = QHBoxLayout(button_row)
        button_layout.addWidget(open_button)
        button_layout.addWidget(preferences_button)
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.path_label, 1)
        toolbar_layout.addWidget(button_row)

        self.table = QTableView()
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Interactive
        )
        self.table.horizontalHeader().setStretchLastSection(False)

        self.preview_title = QLabel("Select a candidate")
        self.preview_info = QLabel("")
        self.preview = QTextEdit()
        self.preview.setReadOnly(True)
        self.preview.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)

        preview_panel = QWidget()
        preview_layout = QVBoxLayout(preview_panel)
        preview_layout.addWidget(self.preview_title)
        preview_layout.addWidget(self.preview_info)
        preview_layout.addWidget(self.preview, 1)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self.table)
        splitter.addWidget(preview_panel)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)

        central = QWidget()
        layout = QVBoxLayout(central)
        layout.addWidget(toolbar)
        layout.addWidget(splitter, 1)
        self.setCentralWidget(central)
        self.statusBar().showMessage("Ready")

    def _load_csv(self, path: Path) -> None:
        fieldnames, rows, needs_save = load_candidates(path)
        model = CandidateTableModel(fieldnames, rows, self)
        model.keep_changed.connect(self._mark_dirty)

        self.csv_path = path.resolve()
        self.fieldnames = fieldnames
        self.rows = rows
        self.model = model
        self.document_cache.clear()
        self.table.setModel(model)
        self.path_label.setText(str(self.csv_path))
        self.setWindowTitle(f"Index Candidate Review - {self.csv_path.name}")
        self.dirty = needs_save
        self._update_save_state()

        for column, field in enumerate(fieldnames):
            widths = {
                "document": 190,
                "candidate_term": 180,
                "tfidf_score": 85,
                "occurrences": 85,
                "rank": 55,
                "source": 105,
                "entity_type": 135,
                "Keep": 55,
                "source_candidate_term": 0,
            }
            self.table.setColumnWidth(column, widths.get(field, 120))
        self._apply_hidden_columns()

        model.candidate_changed.connect(self._mark_dirty)
        model.candidate_changed.connect(self._candidate_term_changed)
        self.table.selectionModel().currentRowChanged.connect(
            self._current_row_changed
        )
        if rows:
            self.table.selectRow(0)
        else:
            self._clear_preview("This CSV has no candidate rows.")

        if needs_save:
            self.statusBar().showMessage(
                "The Keep column or some Keep values need saving."
            )
        else:
            self.statusBar().showMessage(f"Loaded {len(rows)} candidates.")

    def _candidate_term_changed(self, row_index: int) -> None:
        if self.table.currentIndex().row() == row_index and self.model is not None:
            index = self.model.index(row_index, self.model.candidate_column)
            self._current_row_changed(index, index)

    def _apply_hidden_columns(self) -> None:
        if self.model is None:
            return
        hidden_fields = set(self.preferences.hidden_fields)
        hidden_fields.add("source_candidate_term")
        for column, field in enumerate(self.fieldnames):
            self.table.setColumnHidden(column, field in hidden_fields)

    def _edit_preferences(self) -> None:
        dialog = PreferencesDialog(self.preferences, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        try:
            save_preferences(dialog.preferences)
        except (OSError, ValueError) as error:
            QMessageBox.critical(self, "Could not save preferences", str(error))
            return

        self.preferences = dialog.preferences
        self._apply_hidden_columns()
        current = self.table.currentIndex()
        if current.isValid():
            self._current_row_changed(current, current)
        self.statusBar().showMessage(f"Saved preferences to {PREFERENCES_PATH}")

    def _choose_csv(self) -> None:
        if not self._confirm_discard():
            return
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Open candidate CSV",
            str(
                self.preferences.index_path
                if self.preferences.index_path.is_dir()
                else self.csv_path.parent
            ),
            "CSV files (*.csv);;All files (*)",
        )
        if not file_name:
            return
        try:
            self._load_csv(Path(file_name))
        except (OSError, csv.Error, ValueError) as error:
            QMessageBox.critical(self, "Could not open CSV", str(error))

    def _current_row_changed(
        self,
        current: QModelIndex,
        _previous: QModelIndex,
    ) -> None:
        if not current.isValid() or self.model is None:
            self._clear_preview("Select a candidate")
            return

        row = self.rows[current.row()]
        document_name = str(row["document"]).strip()
        term = str(row["candidate_term"]).strip()
        document_path = Path(document_name)
        if not document_path.is_absolute():
            document_path = self.preferences.book_path / document_path
        document_path = document_path.resolve()

        try:
            if document_path not in self.document_cache:
                self.document_cache[document_path] = extract_docx_text(document_path)
        except (OSError, ValueError, KeyError, zipfile.BadZipFile, PackageNotFoundError) as error:
            self.preview_title.setText(document_name)
            self.preview_info.setText(f"Could not read the DOCX: {error}")
            self.preview_info.setStyleSheet("color: #b00020;")
            self.preview.setPlainText("")
            self.preview.setExtraSelections([])
            return

        self.preview_title.setText(document_name)
        self.preview_info.setStyleSheet("")
        self.preview.setPlainText(self.document_cache[document_path])
        match_count = self._highlight_term(term)
        if match_count:
            self.preview_info.setText(
                f"{match_count} case-insensitive match(es) highlighted for {term!r}"
            )
        else:
            self.preview_info.setText(f"No text match found for {term!r}")

    def _highlight_term(self, term: str) -> int:
        if not term:
            self.preview.setExtraSelections([])
            return 0

        document = self.preview.document()
        cursor = QTextCursor(document)
        selections = []
        highlight_format = QTextCharFormat()
        highlight_format.setBackground(QColor("#ffeb3b"))
        highlight_format.setForeground(QColor("#000000"))

        while True:
            match = document.find(term, cursor, QTextDocument.FindFlags())
            if match.isNull():
                break
            selection = QTextEdit.ExtraSelection()
            selection.cursor = match
            selection.format = highlight_format
            selections.append(selection)
            cursor = match

        self.preview.setExtraSelections(selections)
        return len(selections)

    def _clear_preview(self, message: str) -> None:
        self.preview_title.setText(message)
        self.preview_info.setText("")
        self.preview.setPlainText("")
        self.preview.setExtraSelections([])

    def _mark_dirty(self) -> None:
        self.dirty = True
        self._update_save_state()
        self.statusBar().showMessage("Unsaved candidate changes")

    def _update_save_state(self) -> None:
        self.save_button.setEnabled(self.dirty)
        marker = " *" if self.dirty else ""
        self.setWindowTitle(
            f"Index Candidate Review - {self.csv_path.name}{marker}"
        )

    def _save(self) -> bool:
        try:
            save_candidates(self.csv_path, self.fieldnames, self.rows)
        except (OSError, csv.Error) as error:
            QMessageBox.critical(self, "Could not save CSV", str(error))
            return False

        self.dirty = False
        self._update_save_state()
        self.statusBar().showMessage(f"Saved {self.csv_path}")
        return True

    def _confirm_discard(self) -> bool:
        if not self.dirty:
            return True
        answer = QMessageBox.question(
            self,
            "Unsaved candidate changes",
            "Save your candidate edits and Keep choices before continuing?",
            QMessageBox.StandardButton.Save
            | QMessageBox.StandardButton.Discard
            | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Save,
        )
        if answer == QMessageBox.StandardButton.Save:
            return self._save()
        return answer == QMessageBox.StandardButton.Discard

    def closeEvent(self, event: Any) -> None:
        if self._confirm_discard():
            event.accept()
        else:
            event.ignore()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Review CSV index candidates and preview their DOCX context."
    )
    parser.add_argument(
        "csv",
        nargs="?",
        type=Path,
        help=(
            "Candidate CSV (defaults to index_candidates.csv in index_path "
            f"from {PREFERENCES_PATH.name})."
        ),
    )
    args = parser.parse_args()

    app = QApplication(sys.argv)
    app.setApplicationName("Index Candidate Review")
    try:
        preferences = load_preferences()
    except (OSError, ValueError) as error:
        QMessageBox.critical(
            None,
            "Could not load preferences",
            f"{PREFERENCES_PATH}\n\n{error}",
        )
        return 1

    csv_path = (
        args.csv.expanduser().resolve()
        if args.csv
        else preferences.index_path / "index_candidates.csv"
    )
    if not csv_path.is_file():
        initial_directory = (
            preferences.index_path
            if preferences.index_path.is_dir()
            else PREFERENCES_PATH.parent
        )
        selected_path, _ = QFileDialog.getOpenFileName(
            None,
            "Open candidate CSV",
            str(initial_directory),
            "CSV files (*.csv);;All files (*)",
        )
        if not selected_path:
            return 0
        csv_path = Path(selected_path)

    try:
        window = CandidateReviewWindow(csv_path, preferences)
    except (OSError, csv.Error, ValueError) as error:
        QMessageBox.critical(None, "Could not open CSV", str(error))
        return 1

    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
