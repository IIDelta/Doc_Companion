"""PyQt action queue for reference reconciliation and Style Guide citations."""
from __future__ import annotations

from pathlib import Path
import traceback

from PyQt5.QtCore import (
    QObject,
    QThread,
    Qt,
    pyqtSignal,
)

from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from macros.referencecheck import (
    CheckResult,
    Finding,
    InlineSuggestion,
    ParseWarning,
    buildrequestemail,
    checkreferences,
    writereport,
)


class CheckWorker(QObject):
    progress = pyqtSignal(int, int, str)
    completed = pyqtSignal(object)
    failed = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(
        self,
        documentpath: str,
        librarypath: str,
    ):
        super().__init__()

        self.documentpath = documentpath
        self.librarypath = librarypath
        self.cancelled = False

    def cancel(self):
        self.cancelled = True

    def run(self):
        try:
            result = checkreferences(
                self.documentpath,
                self.librarypath,
                cancel=lambda: self.cancelled,
                progress=self.progress.emit,
            )

            self.completed.emit(result)

        except Exception:
            self.failed.emit(
                traceback.format_exc()
            )

        finally:
            self.finished.emit()


class ReferenceCheckWindow(QMainWindow):
    actionheaders = [
        "Action",
        "Citation/reference",
        "Library candidate",
        "Confidence",
        "Reason",
    ]

    inlineheaders = [
        "Detected source",
        "Category",
        "Suggested in-text citation",
        "Reviewer action",
    ]

    simpleheaders = [
        "Status",
        "Details",
        "Section",
        "Paragraph",
    ]

    def __init__(self, parent=None):
        super().__init__(parent)

        self.documentpath = ""
        self.librarypath = ""

        self.result: CheckResult | None = None
        self.thread: QThread | None = None
        self.worker: CheckWorker | None = None

        self.setWindowTitle(
            "Reference Library Check — Final Workflow"
        )

        self.resize(1320, 780)

        self.buildui()
        self.updatestayontop()

    def buildui(self):
        root = QWidget(self)

        layout = QVBoxLayout(root)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        guidance = QLabel(
            "Use the action queues to insert verified EndNote "
            "citations, request missing literature references, "
            "and apply Takeda Style Guide inline citations for "
            "websites, regulatory sources, personal communications, "
            "data on file, and material in preparation. "
            "All suggestions require reviewer confirmation."
        )

        guidance.setWordWrap(True)

        layout.addWidget(guidance)

        layout.addLayout(
            self.buildfileselector(
                "Word document",
                self.selectdocument,
                "document",
            )
        )

        layout.addLayout(
            self.buildfileselector(
                "Shared library export",
                self.selectlibrary,
                "library",
            )
        )

        controls = QHBoxLayout()

        self.runbutton = QPushButton(
            "Run Reference Library Check"
        )

        self.runbutton.setObjectName(
            "PrimaryButton"
        )

        self.runbutton.clicked.connect(
            self.startcheck
        )

        self.cancelbutton = QPushButton("Cancel")

        self.cancelbutton.setObjectName(
            "DangerButton"
        )

        self.cancelbutton.setEnabled(False)

        self.cancelbutton.clicked.connect(
            self.cancelcheck
        )

        self.endnotebutton = QPushButton(
            "Copy EndNote Search"
        )

        self.endnotebutton.setObjectName(
            "UtilityButton"
        )

        self.endnotebutton.setEnabled(False)

        self.endnotebutton.clicked.connect(
            self.copyendnotesearch
        )

        self.inlinebutton = QPushButton(
            "Copy Suggested Citation"
        )

        self.inlinebutton.setObjectName(
            "UtilityButton"
        )

        self.inlinebutton.setEnabled(False)

        self.inlinebutton.clicked.connect(
            self.copysuggestion
        )

        self.contextbutton = QPushButton(
            "Copy Context"
        )

        self.contextbutton.setObjectName(
            "UtilityButton"
        )

        self.contextbutton.setEnabled(False)

        self.contextbutton.clicked.connect(
            self.copycontext
        )

        self.emailbutton = QPushButton(
            "Prepare Request Email"
        )

        self.emailbutton.setObjectName(
            "SuccessButton"
        )

        self.emailbutton.setEnabled(False)

        self.emailbutton.clicked.connect(
            self.prepareemail
        )

        self.exportbutton = QPushButton(
            "Export CSV"
        )

        self.exportbutton.setEnabled(False)

        self.exportbutton.clicked.connect(
            self.exportreport
        )

        for button in (
            self.runbutton,
            self.cancelbutton,
            self.endnotebutton,
            self.inlinebutton,
            self.contextbutton,
            self.emailbutton,
        ):
            controls.addWidget(button)

        controls.addStretch()
        controls.addWidget(self.exportbutton)

        layout.addLayout(controls)

        self.progressbar = QProgressBar()

        self.progressbar.setRange(0, 1)
        self.progressbar.setVisible(False)

        self.statuslabel = QLabel(
            "Select both inputs to begin."
        )

        self.statuslabel.setObjectName(
            "StatusLabel"
        )

        self.statuslabel.setWordWrap(True)

        layout.addWidget(self.progressbar)
        layout.addWidget(self.statuslabel)

        self.tabs = QTabWidget()

        self.endnotetable = self.createactiontable()
        self.requesttable = self.createactiontable()
        self.inlinetable = self.createtable(
            self.inlineheaders
        )
        self.existingtable = self.createtable(
            self.simpleheaders
        )
        self.warningtable = self.createtable(
            self.simpleheaders
        )

        self.tabs.addTab(
            self.endnotetable,
            "Insert via EndNote (0)",
        )

        self.tabs.addTab(
            self.requesttable,
            "Request References (0)",
        )

        self.tabs.addTab(
            self.inlinetable,
            "Style Guide Inline Sources (0)",
        )

        self.tabs.addTab(
            self.existingtable,
            "Existing EndNote Fields (0)",
        )

        self.tabs.addTab(
            self.warningtable,
            "Parser Warnings (0)",
        )

        layout.addWidget(
            self.tabs,
            1,
        )

        self.detail = QTextEdit()

        self.detail.setReadOnly(True)

        self.detail.setPlaceholderText(
            "Select a row to review context, candidate data, "
            "request details, or the suggested style-guide citation."
        )

        self.detail.setMaximumHeight(190)

        layout.addWidget(self.detail)

        self.setCentralWidget(root)

        for table in (
            self.endnotetable,
            self.requesttable,
            self.inlinetable,
            self.existingtable,
            self.warningtable,
        ):
            table.itemSelectionChanged.connect(
                self.showdetail
            )

    def buildfileselector(
        self,
        labeltext,
        callback,
        which,
    ):
        row = QHBoxLayout()

        label = QLabel(
            f"{labeltext}: not selected"
        )

        label.setWordWrap(True)

        button = QPushButton(
            f"Select {labeltext.title()}"
        )

        button.clicked.connect(callback)

        if which == "document":
            self.documentlabel = label
        else:
            self.librarylabel = label

        row.addWidget(label, 1)
        row.addWidget(button)

        return row

    def createtable(self, headers):
        table = QTableWidget(
            0,
            len(headers),
        )

        table.setHorizontalHeaderLabels(headers)

        table.setSelectionBehavior(
            QAbstractItemView.SelectRows
        )

        table.setSelectionMode(
            QAbstractItemView.SingleSelection
        )

        table.setEditTriggers(
            QAbstractItemView.NoEditTriggers
        )

        table.setAlternatingRowColors(True)

        table.horizontalHeader().setStretchLastSection(
            True
        )

        return table

    def createactiontable(self):
        table = self.createtable(
            self.actionheaders
        )

        table.setColumnWidth(0, 215)
        table.setColumnWidth(1, 360)
        table.setColumnWidth(2, 360)
        table.setColumnWidth(3, 90)

        return table

    def selectdocument(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Word Document",
            "",
            "Word Documents (*.docx)",
        )

        if path:
            self.documentpath = path

            self.documentlabel.setText(
                f"Word document: {Path(path).name}"
            )

    def selectlibrary(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Shared Library Export",
            "",
            "Text Export (*.txt);;All Files (*.*)",
        )

        if path:
            self.librarypath = path

            self.librarylabel.setText(
                f"Shared library export: {Path(path).name}"
            )

    def startcheck(self):
        if not self.documentpath or not self.librarypath:
            QMessageBox.warning(
                self,
                "Inputs required",
                (
                    "Select both a .docx document and "
                    "a shared-library .txt export."
                ),
            )

            return

        self.result = None

        self.clearresults()

        self.runbutton.setEnabled(False)
        self.cancelbutton.setEnabled(True)

        for button in (
            self.endnotebutton,
            self.inlinebutton,
            self.contextbutton,
            self.emailbutton,
            self.exportbutton,
        ):
            button.setEnabled(False)

        self.progressbar.setRange(0, 0)
        self.progressbar.setVisible(True)

        self.statuslabel.setText(
            "Starting local reference reconciliation…"
        )

        self.thread = QThread(self)

        self.worker = CheckWorker(
            self.documentpath,
            self.librarypath,
        )

        self.worker.moveToThread(self.thread)

        self.thread.started.connect(
            self.worker.run
        )

        self.worker.progress.connect(
            self.updateprogress
        )

        self.worker.completed.connect(
            self.checkcomplete
        )

        self.worker.failed.connect(
            self.checkfailed
        )

        self.worker.finished.connect(
            self.thread.quit
        )

        self.worker.finished.connect(
            self.worker.deleteLater
        )

        self.thread.finished.connect(
            self.thread.deleteLater
        )

        self.thread.finished.connect(
            self.threadfinished
        )

        self.thread.start()

    def updateprogress(
        self,
        current,
        total,
        message,
    ):
        if total > 0:
            self.progressbar.setRange(0, total)

            self.progressbar.setValue(
                min(current, total)
            )

            self.progressbar.setFormat(
                f"{current:,} / {total:,} "
                "work units (%p%)"
            )

        else:
            self.progressbar.setRange(0, 0)

        self.statuslabel.setText(message)

    def cancelcheck(self):
        if self.worker:
            self.worker.cancel()

            self.cancelbutton.setEnabled(False)

            self.statuslabel.setText(
                "Cancelling after the current operation…"
            )

    def checkcomplete(
        self,
        result,
    ):
        self.result = result

        endnoteitems = (
            result.matched
            + result.review
        )

        requestitems = (
            result.missing
            + result.potentialmissing
        )

        self.fillactiontable(
            self.endnotetable,
            endnoteitems,
        )

        self.fillactiontable(
            self.requesttable,
            requestitems,
        )

        self.fillinlinetable(
            result.inlineonly
        )

        self.fillsimpletable(
            self.existingtable,
            [
                (
                    "Existing EndNote citation field",
                    item.raw,
                    item.section,
                    item.paragraph,
                    item,
                )
                for item in result.existingendnote
            ],
        )

        self.fillsimpletable(
            self.warningtable,
            [
                (
                    "Parser warning",
                    item.message,
                    item.section,
                    item.paragraph,
                    item,
                )
                for item in result.warnings
            ],
        )

        self.tabs.setTabText(
            0,
            (
                "Insert via EndNote "
                f"({len(endnoteitems)})"
            ),
        )

        self.tabs.setTabText(
            1,
            (
                "Request References "
                f"({len(requestitems)})"
            ),
        )

        self.tabs.setTabText(
            2,
            (
                "Style Guide Inline Sources "
                f"({len(result.inlineonly)})"
            ),
        )

        self.tabs.setTabText(
            3,
            (
                "Existing EndNote Fields "
                f"({len(result.existingendnote)})"
            ),
        )

        self.tabs.setTabText(
            4,
            (
                "Parser Warnings "
                f"({len(result.warnings)})"
            ),
        )

        summary = result.summary()

        suffix = (
            " (cancelled; results are partial)"
            if result.cancelled
            else ""
        )

        self.statuslabel.setText(
            f"Complete{suffix}: "
            f"{summary['referenceblocks']} reference blocks; "
            f"{summary['matched']} matched; "
            f"{summary['review']} review; "
            f"{summary['missing']} missing; "
            f"{summary['potentialmissing']} "
            f"need identification; "
            f"{summary['inlineonly']} inline sources."
        )

        self.endnotebutton.setEnabled(
            bool(endnoteitems)
        )

        self.inlinebutton.setEnabled(
            bool(result.inlineonly)
        )

        self.contextbutton.setEnabled(
            bool(
                endnoteitems
                or requestitems
                or result.inlineonly
            )
        )

        self.emailbutton.setEnabled(
            bool(requestitems)
        )

        self.exportbutton.setEnabled(True)

    def checkfailed(
        self,
        detail,
    ):
        self.statuslabel.setText(
            "Reference Library Check failed."
        )

        QMessageBox.critical(
            self,
            "Reference Library Check failed",
            detail,
        )

    def threadfinished(self):
        self.progressbar.setVisible(False)

        self.runbutton.setEnabled(True)
        self.cancelbutton.setEnabled(False)

        self.thread = None
        self.worker = None

    def clearresults(self):
        for table in (
            self.endnotetable,
            self.requesttable,
            self.inlinetable,
            self.existingtable,
            self.warningtable,
        ):
            table.setRowCount(0)

        self.detail.clear()

    def fillactiontable(
        self,
        table,
        findings,
    ):
        table.setRowCount(len(findings))

        for row, finding in enumerate(findings):
            values = [
                finding.action,
                finding.documentrecord.label(),
                (
                    finding.bestmatch.label()
                    if finding.bestmatch
                    else "No confident candidate"
                ),
                f"{finding.score:.0%}",
                finding.reason,
            ]

            for column, value in enumerate(values):
                item = QTableWidgetItem(value)

                if column == 0:
                    item.setData(
                        Qt.UserRole,
                        finding,
                    )

                table.setItem(
                    row,
                    column,
                    item,
                )

        table.resizeRowsToContents()

    def fillinlinetable(
        self,
        suggestions,
    ):
        self.inlinetable.setRowCount(
            len(suggestions)
        )

        for row, suggestion in enumerate(suggestions):
            values = [
                suggestion.detected,
                suggestion.category,
                suggestion.suggestion,
                suggestion.missingmetadata,
            ]

            for column, value in enumerate(values):
                item = QTableWidgetItem(value)

                if column == 0:
                    item.setData(
                        Qt.UserRole,
                        suggestion,
                    )

                self.inlinetable.setItem(
                    row,
                    column,
                    item,
                )

        self.inlinetable.resizeRowsToContents()

    def fillsimpletable(
        self,
        table,
        items,
    ):
        table.setRowCount(len(items))

        for row, (
            status,
            detail,
            section,
            paragraph,
            data,
        ) in enumerate(items):
            values = [
                status,
                detail,
                section or "",
                str(paragraph or ""),
            ]

            for column, value in enumerate(values):
                item = QTableWidgetItem(value)

                if column == 0:
                    item.setData(
                        Qt.UserRole,
                        data,
                    )

                table.setItem(
                    row,
                    column,
                    item,
                )

        table.resizeRowsToContents()

    def selecteddata(self):
        table = self.sender()

        if not isinstance(table, QTableWidget):
            table = self.tabs.currentWidget()

        selected = (
            table.selectedItems()
            if isinstance(table, QTableWidget)
            else []
        )

        if not selected:
            return None

        item = table.item(
            selected[0].row(),
            0,
        )

        return (
            item.data(Qt.UserRole)
            if item
            else None
        )

    def showdetail(self):
        data = self.selecteddata()

        if isinstance(data, Finding):
            lines = [
                f"Action: {data.action}",
                f"Status: {data.category}",
                (
                    "Section: "
                    f"{data.documentrecord.section or 'Not identified'}"
                ),
                (
                    "Paragraph: "
                    f"{data.documentrecord.paragraph or 'Not identified'}"
                ),
                (
                    "Citation/reference found: "
                    f"{data.documentrecord.raw}"
                ),
                (
                    "Context: "
                    f"{data.documentrecord.context or 'Not available'}"
                ),
            ]

            if data.bestmatch:
                lines.extend(
                    [
                        "",
                        (
                            "Best library candidate "
                            f"({data.score:.1%}):"
                        ),
                        data.bestmatch.raw,
                        "",
                        "EndNote search text:",
                        data.bestmatch.endnote_search(),
                    ]
                )

            if data.candidates:
                lines.append("\nRanked candidates:")

                for number, candidate in enumerate(
                    data.candidates,
                    start=1,
                ):
                    lines.append(
                        f"{number}. "
                        f"{candidate.score:.1%} — "
                        f"{candidate.record.label()} "
                        f"[{candidate.reason}]"
                    )

            self.detail.setPlainText(
                "\n".join(lines)
            )

            return

        if isinstance(data, InlineSuggestion):
            self.detail.setPlainText(
                (
                    "Action: Use Style Guide inline citation\n"
                    f"Classification: {data.category}\n"
                    f"Detected source: {data.detected}\n"
                    f"Section: {data.section or 'Not identified'}\n"
                    f"Paragraph: {data.paragraph or 'Not identified'}\n"
                    f"Context: {data.context or 'Not available'}\n"
                    f"Why flagged: {data.reason}\n\n"
                    "Suggested in-text citation:\n"
                    f"{data.suggestion}\n\n"
                    "Reviewer action required:\n"
                    f"{data.missingmetadata}\n\n"
                    f"{data.relatedreference or 'No corresponding bibliography entry detected.'}"
                )
            )

            return

        if isinstance(data, ParseWarning):
            self.detail.setPlainText(
                (
                    "Parser warning:\n"
                    f"{data.message}\n\n"
                    f"Section: {data.section}\n"
                    f"Paragraph: {data.paragraph}\n"
                    f"Context:\n{data.context}"
                )
            )

    def copyendnotesearch(self):
        data = self.selecteddata()

        if (
            not isinstance(data, Finding)
            or not data.bestmatch
        ):
            QMessageBox.information(
                self,
                "Select a library match",
                (
                    "Select a row in Insert via EndNote "
                    "with a library candidate first."
                ),
            )

            return

        QApplication.clipboard().setText(
            data.bestmatch.endnote_search()
        )

        self.statuslabel.setText(
            "EndNote search text copied to the clipboard."
        )

    def copysuggestion(self):
        data = self.selecteddata()

        if not isinstance(data, InlineSuggestion):
            QMessageBox.information(
                self,
                "Select an inline source",
                (
                    "Select a row in Style Guide "
                    "Inline Sources first."
                ),
            )

            return

        QApplication.clipboard().setText(
            data.suggestion
        )

        self.statuslabel.setText(
            "Suggested in-text citation copied to the clipboard."
        )

    def copycontext(self):
        data = self.selecteddata()

        if isinstance(data, Finding):
            text = (
                f"Action: {data.action}\n"
                f"Section: {data.documentrecord.section}\n"
                f"Paragraph: {data.documentrecord.paragraph}\n"
                f"Citation/reference: {data.documentrecord.raw}\n"
                f"Context: {data.documentrecord.context}"
            )

        elif isinstance(data, InlineSuggestion):
            text = (
                "Action: Use Style Guide inline citation\n"
                f"Section: {data.section}\n"
                f"Paragraph: {data.paragraph}\n"
                f"Detected source: {data.detected}\n"
                f"Context: {data.context}\n"
                f"Suggested citation: {data.suggestion}"
            )

        else:
            QMessageBox.information(
                self,
                "Select a result",
                (
                    "Select an actionable result row first."
                ),
            )

            return

        QApplication.clipboard().setText(text)

        self.statuslabel.setText(
            "Context copied to the clipboard."
        )

    def prepareemail(self):
        if not self.result:
            return

        email = buildrequestemail(
            self.result,
            Path(self.documentpath).name,
        )

        QApplication.clipboard().setText(email)

        self.detail.setPlainText(email)

        self.statuslabel.setText(
            (
                "Missing-reference request email copied "
                "to the clipboard and displayed below."
            )
        )

    def exportreport(self):
        if not self.result:
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Review Report",
            "referencecheckreport.csv",
            "CSV Files (*.csv)",
        )

        if not path:
            return

        try:
            writereport(
                self.result,
                path,
            )

            QMessageBox.information(
                self,
                "Report saved",
                f"Review report saved to:\n{path}",
            )

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Export failed",
                str(exc),
            )

    def updatestayontop(self):
        parent = self.parent()

        if (
            parent is not None
            and hasattr(
                parent,
                "stay_on_top_checkbox",
            )
        ):
            if parent.stay_on_top_checkbox.isChecked():
                self.setWindowFlags(
                    self.windowFlags()
                    | Qt.WindowStaysOnTopHint
                )

            else:
                self.setWindowFlags(
                    self.windowFlags()
                    & ~Qt.WindowStaysOnTopHint
                )

            self.show()
