"""
pleasereviewwindow.py

PyQt user interface for creating a PleaseReview copy of a
tracked-changes Microsoft Word document.
"""

from __future__ import annotations

import threading
from pathlib import Path
from typing import Any, Dict, Optional

from PyQt5.QtCore import (
    QObject,
    QThread,
    Qt,
    QUrl,
    pyqtSignal,
    pyqtSlot,
)
from PyQt5.QtGui import QDesktopServices, QIcon
from PyQt5.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from macros.PreparePleaseReview import (
    PleaseReviewOptions,
    PleaseReviewResult,
    build_output_path,
    prepare_please_review_copy,
)


class PleaseReviewWorker(QObject):
    """Run PleaseReview processing in a dedicated Qt worker thread."""

    progress = pyqtSignal(dict)
    finished = pyqtSignal(object)

    def __init__(
        self,
        source_path: str,
        output_path: str,
        options: PleaseReviewOptions,
    ) -> None:
        super().__init__()

        self._source_path = source_path
        self._output_path = output_path
        self._options = options
        self._cancel_event = threading.Event()

    @pyqtSlot()
    def run(self) -> None:
        """Run the processing engine and emit the final result."""
        result = prepare_please_review_copy(
            source_path=self._source_path,
            output_path=self._output_path,
            options=self._options,
            progress_callback=self.progress.emit,
            cancel_event=self._cancel_event,
        )

        self.finished.emit(result)

    def request_cancel(self) -> None:
        """Request safe cancellation before the next revision begins."""
        self._cancel_event.set()


class PleaseReviewWindow(QMainWindow):
    """Configure and run hybrid PleaseReview conversion."""

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)

        self._thread: Optional[QThread] = None
        self._worker: Optional[PleaseReviewWorker] = None
        self._last_result: Optional[PleaseReviewResult] = None

        self._source_path: Optional[Path] = None
        self._output_path: Optional[Path] = None

        self.setWindowTitle("Prepare PleaseReview Copy")
        self.setMinimumSize(780, 700)
        self.setWindowIcon(QIcon(self._resource_path("leaf.ico")))

        self._build_ui()
        self.update_stay_on_top()

    def _resource_path(self, filename: str) -> str:
        """Return the local source-tree path for a UI resource."""
        return str(Path(__file__).resolve().parent / filename)

    def _build_ui(self) -> None:
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        title = QLabel("Prepare a PleaseReview Copy", self)
        title.setObjectName("SectionHeader")
        title.setStyleSheet("font-size: 18px; font-weight: 600;")
        layout.addWidget(title)

        description = QLabel(
            (
                "Creates a separate review copy. Tracked formatting and "
                "property changes are accepted in the output copy first. "
                "Microsoft Word then accepts text insertions with yellow "
                "highlight and restores text deletions as red strikethrough "
                "using the custom style "
                "'DC PleaseReview Deleted Text'."
            ),
            self,
        )
        description.setWordWrap(True)
        layout.addWidget(description)

        layout.addWidget(self._build_paths_group())
        layout.addWidget(self._build_options_group())
        layout.addWidget(self._build_safety_frame())
        layout.addLayout(self._build_action_layout())

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Ready")
        layout.addWidget(self.progress_bar)

        self.status_label = QLabel(
            "Choose a source protocol to begin.",
            self,
        )
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        self.eta_label = QLabel(
            "Elapsed: — | ETA: Calculating... | Rate: —",
            self,
        )
        self.eta_label.setObjectName("FilePathLabel")
        self.eta_label.setWordWrap(True)
        layout.addWidget(self.eta_label)

        self.log_text = QTextEdit(self)
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(190)
        self.log_text.setPlaceholderText(
            "Processing activity, warnings, ETA, and completion details "
            "will appear here."
        )
        layout.addWidget(self.log_text)

    def _build_paths_group(self) -> QGroupBox:
        group = QGroupBox("Document paths", self)
        layout = QGridLayout(group)

        self.source_path_label = QLabel(
            "No source document selected.",
            self,
        )
        self.source_path_label.setWordWrap(True)
        self.source_path_label.setTextInteractionFlags(
            Qt.TextSelectableByMouse
        )

        self.output_path_label = QLabel(
            "Output path will be generated automatically.",
            self,
        )
        self.output_path_label.setWordWrap(True)
        self.output_path_label.setTextInteractionFlags(
            Qt.TextSelectableByMouse
        )

        self.choose_source_button = QPushButton(
            "Choose Source Document",
            self,
        )
        self.choose_source_button.clicked.connect(
            self.choose_source_document
        )

        self.choose_output_button = QPushButton(
            "Choose Output Location",
            self,
        )
        self.choose_output_button.clicked.connect(
            self.choose_output_location
        )
        self.choose_output_button.setEnabled(False)

        layout.addWidget(QLabel("Source:", self), 0, 0)
        layout.addWidget(self.source_path_label, 0, 1)
        layout.addWidget(self.choose_source_button, 0, 2)

        layout.addWidget(QLabel("Output:", self), 1, 0)
        layout.addWidget(self.output_path_label, 1, 1)
        layout.addWidget(self.choose_output_button, 1, 2)

        return group

    def _build_options_group(self) -> QGroupBox:
        group = QGroupBox("Review formatting", self)
        layout = QFormLayout(group)

        self.redaction_highlight_checkbox = QCheckBox(
            "Also highlight restored deletions in yellow",
            self,
        )
        self.redaction_highlight_checkbox.setChecked(False)

        self.process_moves_checkbox = QCheckBox(
            "Treat moved text as insertion/deletion text",
            self,
        )
        self.process_moves_checkbox.setChecked(False)

        layout.addRow(
            "Deletion appearance:",
            QLabel(
                (
                    "Red strikethrough using custom style "
                    "'DC PleaseReview Deleted Text'."
                ),
                self,
            ),
        )

        layout.addRow("", self.redaction_highlight_checkbox)
        layout.addRow("", self.process_moves_checkbox)

        return group

    def _build_safety_frame(self) -> QFrame:
        frame = QFrame(self)
        frame.setFrameShape(QFrame.StyledPanel)

        layout = QVBoxLayout(frame)

        title = QLabel("Safety behavior", self)
        title.setStyleSheet("font-weight: 600;")
        layout.addWidget(title)

        message = QLabel(
            (
                "The original protocol is never modified. The output copy "
                "first accepts tracked formatting/property changes, then "
                "processes remaining text revisions in Microsoft Word. "
                "Safe cancellation saves a clearly marked partial output "
                "and JSON audit log."
            ),
            self,
        )
        message.setWordWrap(True)
        layout.addWidget(message)

        return frame

    def _build_action_layout(self) -> QHBoxLayout:
        layout = QHBoxLayout()

        self.start_button = QPushButton(
            "Create PleaseReview Copy",
            self,
        )
        self.start_button.setObjectName("PrimaryButton")
        self.start_button.clicked.connect(self.start_processing)
        self.start_button.setEnabled(False)

        self.cancel_button = QPushButton(
            "Cancel After Current Revision",
            self,
        )
        self.cancel_button.clicked.connect(self.request_cancel)
        self.cancel_button.setEnabled(False)

        self.open_output_button = QPushButton(
            "Open Output Folder",
            self,
        )
        self.open_output_button.clicked.connect(
            self.open_output_folder
        )
        self.open_output_button.setEnabled(False)

        layout.addWidget(self.start_button)
        layout.addWidget(self.cancel_button)
        layout.addStretch()
        layout.addWidget(self.open_output_button)

        return layout

    def choose_source_document(self) -> None:
        """Choose a DOCX or DOCM tracked-changes source document."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Choose Tracked-Changes Word Document",
            "",
            "Word Documents (*.docx *.docm)",
        )

        if not file_path:
            return

        self._source_path = Path(file_path).resolve()
        self._output_path = build_output_path(self._source_path)

        self.source_path_label.setText(str(self._source_path))
        self.output_path_label.setText(str(self._output_path))

        self.choose_output_button.setEnabled(True)
        self.start_button.setEnabled(True)
        self.open_output_button.setEnabled(False)

        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Ready")

        self.status_label.setText(
            "Ready. A separate PleaseReview copy will be created."
        )

        self.eta_label.setText(
            "Elapsed: — | ETA: Calculating... | Rate: —"
        )

        self.log_text.clear()

    def choose_output_location(self) -> None:
        """Choose a new output filename without allowing overwrite."""
        if self._source_path is None:
            return

        default_path = str(
            self._output_path or build_output_path(self._source_path)
        )

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Choose PleaseReview Output File",
            default_path,
            "Word Documents (*.docx *.docm)",
        )

        if not file_path:
            return

        selected_output = Path(file_path).resolve()

        if selected_output == self._source_path:
            QMessageBox.warning(
                self,
                "Invalid Output Path",
                "The output path must be different from the source document.",
            )
            return

        if selected_output.exists():
            QMessageBox.warning(
                self,
                "Output Already Exists",
                (
                    "The selected output file already exists. "
                    "Choose a new name to prevent accidental overwrite."
                ),
            )
            return

        self._output_path = selected_output
        self.output_path_label.setText(str(self._output_path))

    def start_processing(self) -> None:
        """Validate settings and start hybrid processing in a worker thread."""
        if self._source_path is None or self._output_path is None:
            QMessageBox.warning(
                self,
                "Source Document Required",
                (
                    "Choose a source document and output location "
                    "before starting."
                ),
            )
            return

        if self._output_path.exists():
            QMessageBox.warning(
                self,
                "Output Already Exists",
                (
                    "The output file already exists. "
                    "Choose another output location."
                ),
            )
            return

        options = PleaseReviewOptions(
            deletion_highlight_yellow=(
                self.redaction_highlight_checkbox.isChecked()
            ),
            process_moved_text=(
                self.process_moves_checkbox.isChecked()
            ),
            progress_interval=1,
        )

        self._set_running_state(True)
        self._last_result = None

        self.log_text.clear()

        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("Starting...")

        self.status_label.setText(
            "Creating output copy and starting processing."
        )

        self.eta_label.setText(
            "Elapsed: — | ETA: Calculating... | Rate: —"
        )

        self._thread = QThread(self)

        self._worker = PleaseReviewWorker(
            source_path=str(self._source_path),
            output_path=str(self._output_path),
            options=options,
        )

        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self.handle_progress)
        self._worker.finished.connect(self.handle_finished)

        self._worker.finished.connect(self._thread.quit)
        self._worker.finished.connect(self._worker.deleteLater)

        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.finished.connect(
            self._clear_thread_references
        )

        self._thread.start()

    def request_cancel(self) -> None:
        """Request safe cancellation after the current revision completes."""
        if self._worker is None:
            return

        self._worker.request_cancel()
        self.cancel_button.setEnabled(False)

        self.status_label.setText(
            (
                "Cancellation requested. Processing will stop before the "
                "next revision begins."
            )
        )

        self._append_log(
            (
                "Cancellation requested. Waiting for the current revision "
                "operation to complete safely."
            )
        )

    @staticmethod
    def _format_duration(seconds: Optional[float]) -> str:
        """Format elapsed or remaining seconds for display."""
        if seconds is None or seconds < 0:
            return "Calculating..."

        total_seconds = max(0, round(seconds))

        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds_part = divmod(remainder, 60)

        if hours:
            return f"{hours}h {minutes}m"

        if minutes:
            return f"{minutes}m {seconds_part}s"

        return f"{seconds_part}s"

    @pyqtSlot(dict)
    def handle_progress(self, payload: Dict[str, Any]) -> None:
        """Render activity, progress, ETA, warnings, and errors."""
        kind = payload.get("kind", "log")
        message = str(payload.get("message", ""))

        if kind == "activity":
            self.status_label.setText(message)
            self._append_log(message)
            return

        if kind == "progress":
            percent = float(payload.get("percent", 0.0))

            completed = int(
                payload.get("completed_revisions", 0)
            )

            total = int(
                payload.get("total_revisions", 0)
            )

            elapsed_seconds = payload.get("elapsed_seconds")

            estimated_remaining_seconds = payload.get(
                "estimated_remaining_seconds"
            )

            revisions_per_minute = float(
                payload.get("revisions_per_minute", 0.0)
            )

            eta_confidence = str(
                payload.get("eta_confidence", "Calculating")
            )

            elapsed_text = self._format_duration(elapsed_seconds)

            remaining_text = self._format_duration(
                estimated_remaining_seconds
            )

            self.progress_bar.setValue(round(percent))

            self.progress_bar.setFormat(
                (
                    f"{percent:.1f}% — {completed:,} "
                    f"of {total:,} revisions"
                )
            )

            self.status_label.setText(message)

            self.eta_label.setText(
                (
                    f"Elapsed: {elapsed_text} | "
                    f"ETA: {remaining_text} | "
                    f"Rate: {revisions_per_minute:.1f} "
                    f"revisions/minute | "
                    f"Confidence: {eta_confidence}"
                )
            )

            self._append_log(message)
            return

        if kind == "started":
            total = int(payload.get("total_revisions", 0))

            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(0)

            self.progress_bar.setFormat(
                f"0.0% — 0 of {total:,} revisions"
            )

        if message:
            self.status_label.setText(message)
            self._append_log(message)

    @pyqtSlot(object)
    def handle_finished(self, result: PleaseReviewResult) -> None:
        """Render final completion, cancellation, or error summary."""
        self._last_result = result
        self._set_running_state(False)

        if result.total_revisions:
            final_percent = (
                result.completed_revisions
                / result.total_revisions
            ) * 100
        else:
            final_percent = 0.0

        self.progress_bar.setValue(round(final_percent))

        self.progress_bar.setFormat(
            (
                f"{final_percent:.1f}% — "
                f"{result.completed_revisions:,} of "
                f"{result.total_revisions:,} revisions"
            )
        )

        self.eta_label.setText(
            (
                "Total elapsed time: "
                f"{result.elapsed_seconds or 0:.1f} seconds"
            )
        )

        summary = self._build_completion_summary(result)

        self._append_log("")
        self._append_log(summary)

        self.open_output_button.setEnabled(
            Path(result.output_path).exists()
        )

        if result.success:
            self.status_label.setText(
                "PleaseReview copy completed successfully."
            )

            QMessageBox.information(
                self,
                "PleaseReview Copy Completed",
                summary,
            )
            return

        if result.cancelled:
            self.status_label.setText(
                (
                    "Processing was cancelled. Review the partial "
                    "output and audit log."
                )
            )

            QMessageBox.warning(
                self,
                "PleaseReview Copy Cancelled",
                summary,
            )
            return

        self.status_label.setText(
            "Processing failed. Review the audit log and error details."
        )

        QMessageBox.critical(
            self,
            "PleaseReview Copy Failed",
            summary,
        )

    def _build_completion_summary(
        self,
        result: PleaseReviewResult,
    ) -> str:
        """Build a readable completion summary without protocol text."""
        skipped_lines = [
            f"  • {revision_type}: {count:,}"
            for revision_type, count in sorted(
                result.skipped_by_type.items()
            )
        ]

        skipped_detail = (
            "\n".join(skipped_lines)
            if skipped_lines
            else "  • None"
        )

        formatting_lines = [
            f"  • {revision_type}: {count:,}"
            for revision_type, count in sorted(
                result.formatting_revisions_flattened_by_type.items()
            )
        ]

        formatting_detail = (
            "\n".join(formatting_lines)
            if formatting_lines
            else "  • None"
        )

        error_lines = [
            f"  • {error}"
            for error in result.errors[:2]
        ]

        errors_detail = (
            "\n".join(error_lines)
            if error_lines
            else "  • None"
        )

        if result.success:
            status = "Completed successfully"
        elif result.cancelled:
            status = "Cancelled with partial output"
        else:
            status = "Failed"

        return (
            f"Status: {status}\n\n"
            f"Formatting/property revisions accepted in output copy: "
            f"{result.formatting_revisions_flattened:,}\n"
            f"Insertions converted: {result.insertions_converted:,}\n"
            f"Deletions converted: {result.deletions_converted:,}\n"
            f"Moved-to conversions: {result.moved_to_converted:,}\n"
            f"Moved-from conversions: {result.moved_from_converted:,}\n"
            f"Unsupported revisions retained: "
            f"{result.unsupported_revisions:,}\n"
            f"Deletion-style warnings: "
            f"{result.deletion_validation_warnings:,}\n\n"
            f"Formatting/property revisions accepted:\n"
            f"{formatting_detail}\n\n"
            f"Skipped revision types:\n"
            f"{skipped_detail}\n\n"
            f"Output document:\n"
            f"{result.output_path}\n\n"
            f"Audit log:\n"
            f"{result.audit_path}\n\n"
            f"Elapsed time: "
            f"{result.elapsed_seconds or 0:.1f} seconds\n\n"
            f"Errors:\n"
            f"{errors_detail}"
        )

    def open_output_folder(self) -> None:
        """Open the output folder in Windows Explorer."""
        if self._last_result is not None:
            output_directory = Path(
                self._last_result.output_path
            ).parent

        elif self._output_path is not None:
            output_directory = self._output_path.parent

        else:
            return

        QDesktopServices.openUrl(
            QUrl.fromLocalFile(str(output_directory))
        )

    def _append_log(self, message: str) -> None:
        """Append a message to the processing log."""
        self.log_text.append(message)

    def _set_running_state(self, is_running: bool) -> None:
        """Enable or disable controls while processing is active."""
        self.choose_source_button.setEnabled(not is_running)

        self.choose_output_button.setEnabled(
            not is_running and self._source_path is not None
        )

        self.start_button.setEnabled(
            not is_running and self._source_path is not None
        )

        self.cancel_button.setEnabled(is_running)

        self.redaction_highlight_checkbox.setEnabled(
            not is_running
        )

        self.process_moves_checkbox.setEnabled(
            not is_running
        )

    @pyqtSlot()
    def _clear_thread_references(self) -> None:
        """Release worker-thread references after processing finishes."""
        self._thread = None
        self._worker = None

    def update_stay_on_top(self) -> None:
        """Synchronize stay-on-top state with the main window."""
        parent = self.parent()

        if parent is None:
            return

        if not hasattr(parent, "stay_on_top_checkbox"):
            return

        flags = self.windowFlags()

        if parent.stay_on_top_checkbox.isChecked():
            self.setWindowFlags(
                flags | Qt.WindowStaysOnTopHint
            )
        else:
            self.setWindowFlags(
                flags & ~Qt.WindowStaysOnTopHint
            )

        self.show()

    def closeEvent(self, event) -> None:
        """Prevent accidental closing while Word processing is active."""
        if (
            self._thread is not None
            and self._thread.isRunning()
        ):
            QMessageBox.warning(
                self,
                "Processing in Progress",
                (
                    "Please wait for processing to finish or use "
                    "'Cancel After Current Revision'. The window will "
                    "remain open until Word safely stops."
                ),
            )

            event.ignore()
            return

        super().closeEvent(event)
