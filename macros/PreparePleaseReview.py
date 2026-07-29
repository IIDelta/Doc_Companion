"""
PreparePleaseReview.py

Hybrid PleaseReview engine for complex Word tracked-changes documents.

Process:
1. Creates a separate output copy.
2. Accepts tracked formatting/property revisions in the output copy only.
3. Opens the output copy in Microsoft Word.
4. Converts remaining text insertions and deletions:
   - Insertions: accepted and highlighted yellow.
   - Deletions: rejected/restored and marked with the custom character style
     "DC PleaseReview Deleted Text".
5. Writes a JSON audit log without exporting protocol text.

Requirements:
- Windows
- Microsoft Word Desktop
- pywin32
- lxml
"""

from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import threading
import time
import traceback
import zipfile

from collections import Counter, deque
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, Iterator, List, Optional

from lxml import etree

try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None


# ---------------------------------------------------------------------------
# OOXML namespaces and tracked-formatting revision elements
# ---------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(local_name: str) -> str:
    """Return a qualified WordprocessingML element name."""
    return f"{{{W_NS}}}{local_name}"


W_DOCUMENT = qn("document")
W_DOCUMENT_PROTECTION = qn("documentProtection")

W_RPR_CHANGE = qn("rPrChange")
W_PPR_CHANGE = qn("pPrChange")
W_SECT_PR_CHANGE = qn("sectPrChange")
W_TBL_PR_CHANGE = qn("tblPrChange")
W_TR_PR_CHANGE = qn("trPrChange")
W_TC_PR_CHANGE = qn("tcPrChange")
W_TBL_GRID_CHANGE = qn("tblGridChange")

FORMAT_CHANGE_TAGS = {
    W_RPR_CHANGE: "run_property_change",
    W_PPR_CHANGE: "paragraph_property_change",
    W_SECT_PR_CHANGE: "section_property_change",
    W_TBL_PR_CHANGE: "table_property_change",
    W_TR_PR_CHANGE: "table_row_property_change",
    W_TC_PR_CHANGE: "table_cell_property_change",
    W_TBL_GRID_CHANGE: "table_grid_change",
}


# ---------------------------------------------------------------------------
# Microsoft Word constants
# ---------------------------------------------------------------------------

WD_NO_PROTECTION = -1
WD_STYLE_TYPE_CHARACTER = 2

WD_REVISION_INSERT = 1
WD_REVISION_DELETE = 2
WD_REVISION_MOVED_FROM = 14
WD_REVISION_MOVED_TO = 15

WD_YELLOW = 7
WD_COLOR_RED = 255
WD_ALERTS_NONE = 0


# ---------------------------------------------------------------------------
# Application constants
# ---------------------------------------------------------------------------

DEFAULT_REDACTION_STYLE_NAME = "DC PleaseReview Deleted Text"
DEFAULT_OUTPUT_MARKER = "_PleaseReview"
AUDIT_SCHEMA_VERSION = "4.1"

ProgressCallback = Callable[[Dict[str, Any]], None]


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------

@dataclass
class PleaseReviewOptions:
    """Options supplied by the PleaseReview PyQt window."""

    redaction_style_name: str = DEFAULT_REDACTION_STYLE_NAME
    deletion_highlight_yellow: bool = False
    process_moved_text: bool = False

    # Retained for UI compatibility. Progress is emitted every revision so that
    # the user receives live counts and ETA information.
    progress_interval: int = 1

    slow_revision_warning_seconds: float = 15.0

    def validate(self) -> None:
        """Validate user-provided options."""
        if not self.redaction_style_name.strip():
            raise ValueError("A redaction style name is required.")

        if self.progress_interval < 1:
            raise ValueError("progress_interval must be at least 1.")

        if self.slow_revision_warning_seconds <= 0:
            raise ValueError(
                "slow_revision_warning_seconds must be greater than zero."
            )


@dataclass
class PleaseReviewResult:
    """Structured processing result returned to the UI and audit JSON."""

    source_path: str
    output_path: str
    audit_path: str
    started_utc: str

    completed_utc: Optional[str] = None

    success: bool = False
    cancelled: bool = False
    output_saved: bool = False

    total_revisions: int = 0
    completed_revisions: int = 0

    insertions_converted: int = 0
    deletions_converted: int = 0
    moved_to_converted: int = 0
    moved_from_converted: int = 0

    unsupported_revisions: int = 0
    deletion_validation_warnings: int = 0

    formatting_revisions_flattened: int = 0

    formatting_revisions_flattened_by_type: Dict[str, int] = field(
        default_factory=dict
    )

    skipped_by_type: Dict[str, int] = field(default_factory=dict)

    warnings: List[Dict[str, Any]] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    elapsed_seconds: Optional[float] = None
    word_version: Optional[str] = None
    source_track_revisions_enabled: Optional[bool] = None

    def to_dict(self) -> Dict[str, Any]:
        """Return a JSON-serializable version of this result."""
        return asdict(self)


class PleaseReviewError(RuntimeError):
    """Raised for expected PleaseReview processing failures."""


# ---------------------------------------------------------------------------
# Public path helpers
# ---------------------------------------------------------------------------

def build_output_path(source_path: str | Path) -> Path:
    """
    Build a non-conflicting default output path.

    Example:
        Protocol.docx -> Protocol_PleaseReview.docx
    """
    source = Path(source_path).expanduser().resolve()

    candidate = source.with_name(
        f"{source.stem}{DEFAULT_OUTPUT_MARKER}{source.suffix}"
    )

    counter = 2

    while candidate.exists():
        candidate = source.with_name(
            f"{source.stem}{DEFAULT_OUTPUT_MARKER}_{counter}{source.suffix}"
        )
        counter += 1

    return candidate


def build_audit_path(output_path: str | Path) -> Path:
    """Build the JSON audit-log path for an output document."""
    output = Path(output_path).expanduser().resolve()
    return output.with_suffix(".audit.json")


# ---------------------------------------------------------------------------
# General helpers
# ---------------------------------------------------------------------------

def _emit_progress(
    callback: Optional[ProgressCallback],
    payload: Dict[str, Any],
) -> None:
    """Emit progress without allowing UI logging errors to stop processing."""
    if callback is None:
        return

    try:
        callback(payload)
    except Exception:
        pass


def _format_duration(seconds: Optional[float]) -> str:
    """Format a duration for display in progress and ETA messages."""
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


def _parse_xml(xml_bytes: bytes) -> etree._Element:
    """Parse OOXML safely without entity resolution or network access."""
    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        remove_blank_text=False,
        huge_tree=True,
    )

    return etree.fromstring(
        xml_bytes,
        parser=parser,
    )


def _serialize_xml(root: etree._Element) -> bytes:
    """Serialize modified OOXML XML."""
    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=None,
    )


def _is_content_part(part_name: str) -> bool:
    """Return True if an OOXML part can contain visible Word content."""
    normalized = part_name.lower()

    if normalized == "word/document.xml":
        return True

    if normalized.startswith("word/header") and normalized.endswith(".xml"):
        return True

    if normalized.startswith("word/footer") and normalized.endswith(".xml"):
        return True

    return normalized in {
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
        "word/commentsextended.xml",
    }


# ---------------------------------------------------------------------------
# OOXML formatting-revision normalization
# ---------------------------------------------------------------------------

def _normalise_format_revisions(
    output_path: Path,
    progress_callback: Optional[ProgressCallback],
    cancel_event: Optional[threading.Event],
) -> tuple[Dict[str, int], bool]:
    """
    Flatten tracked formatting/property revisions in the output copy only.

    This preserves the current visible formatting but removes historical
    formatting-revision metadata such as w:rPrChange. It does not modify
    text insertion/deletion wrappers such as w:ins or w:del.
    """
    _emit_progress(
        progress_callback,
        {
            "kind": "activity",
            "message": (
                "Normalizing tracked formatting revisions in the output copy."
            ),
        },
    )

    with zipfile.ZipFile(output_path, "r") as source_zip:
        zip_infos = source_zip.infolist()

        package_parts = {
            zip_info.filename: source_zip.read(zip_info.filename)
            for zip_info in zip_infos
        }

    if "word/document.xml" not in package_parts:
        raise PleaseReviewError(
            "Output copy does not contain word/document.xml."
        )

    settings_xml = package_parts.get("word/settings.xml")

    if settings_xml is not None:
        settings_root = _parse_xml(settings_xml)

        protection = settings_root.find(
            f".//{W_DOCUMENT_PROTECTION}"
        )

        if protection is not None:
            enforcement = (
                protection.get(qn("enforcement"))
                or ""
            ).strip().lower()

            if enforcement in {"1", "true", "on"}:
                raise PleaseReviewError(
                    "The selected document has enforced Word protection."
                )

    flattened_counts: Counter[str] = Counter()
    changed_parts: Dict[str, bytes] = {}

    for part_name, xml_bytes in package_parts.items():
        if (
            cancel_event is not None
            and cancel_event.is_set()
        ):
            return dict(flattened_counts), True

        if not _is_content_part(part_name):
            continue

        root = _parse_xml(xml_bytes)
        changed = False

        # Build a snapshot because the XML tree is mutated while iterating.
        for element in list(root.iter()):
            category = FORMAT_CHANGE_TAGS.get(element.tag)

            if category is None:
                continue

            parent = element.getparent()

            if parent is None:
                continue

            parent.remove(element)

            flattened_counts[category] += 1
            changed = True

        if changed:
            changed_parts[part_name] = _serialize_xml(root)

    package_parts.update(changed_parts)

    file_descriptor, temporary_file_name = tempfile.mkstemp(
        prefix="dcprnorm-",
        suffix=output_path.suffix,
        dir=str(output_path.parent),
    )

    os.close(file_descriptor)

    temporary_path = Path(temporary_file_name)

    try:
        with zipfile.ZipFile(
            temporary_path,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            allowZip64=True,
        ) as output_zip:
            for zip_info in zip_infos:
                output_zip.writestr(
                    zip_info,
                    package_parts[zip_info.filename],
                )

        with zipfile.ZipFile(temporary_path, "r") as validation_zip:
            document_root = _parse_xml(
                validation_zip.read("word/document.xml")
            )

            if document_root.tag != W_DOCUMENT:
                raise PleaseReviewError(
                    "Normalized output is not a WordprocessingML document."
                )

        shutil.move(
            str(temporary_path),
            str(output_path),
        )

    except Exception:
        if temporary_path.exists():
            temporary_path.unlink()

        raise

    return dict(flattened_counts), False


# ---------------------------------------------------------------------------
# Microsoft Word helpers
# ---------------------------------------------------------------------------

def _require_word() -> None:
    """Verify that Windows, pywin32, and Word automation are available."""
    if sys.platform != "win32":
        raise PleaseReviewError(
            "Prepare PleaseReview Copy requires Windows and Microsoft Word."
        )

    if pythoncom is None or win32com is None:
        raise PleaseReviewError(
            "pywin32 is unavailable. Reinstall Doc Companion dependencies."
        )


def _safe_close_document(document: Any) -> None:
    """Close a Word document without allowing cleanup errors to interrupt."""
    if document is None:
        return

    try:
        document.Close(SaveChanges=False)
    except Exception:
        pass


def _safe_quit_word(word_app: Any) -> None:
    """Quit the dedicated Word instance without allowing cleanup errors."""
    if word_app is None:
        return

    try:
        word_app.Quit()
    except Exception:
        pass


def _iter_story_ranges(document: Any) -> Iterator[Any]:
    """
    Yield all linked Word story ranges.

    Includes the main story and, where available, headers, footers, notes,
    comments, text boxes, and other linked story ranges.
    """
    for initial_story_range in document.StoryRanges:
        current_story_range = initial_story_range

        while current_story_range is not None:
            yield current_story_range

            try:
                current_story_range = current_story_range.NextStoryRange
            except Exception:
                current_story_range = None


def _story_label(story_range: Any) -> str:
    """Return a readable Word story-type label."""
    try:
        return f"StoryType {int(story_range.StoryType)}"
    except Exception:
        return "Unknown Story"


def _build_revision_plan(document: Any) -> List[Dict[str, Any]]:
    """
    Cache Word revision collections once.

    Avoids repeatedly accessing StoryRange.Revisions.Count during processing.
    """
    plan: List[Dict[str, Any]] = []

    for story_range in _iter_story_ranges(document):
        try:
            revisions = story_range.Revisions
            revision_count = int(revisions.Count)
        except Exception:
            continue

        if revision_count < 1:
            continue

        plan.append(
            {
                "story_name": _story_label(story_range),
                "revisions": revisions,
                "revision_count": revision_count,
            }
        )

    return plan


def _revision_type_name(revision_type: int) -> str:
    """Return a readable Word revision-type label."""
    names = {
        WD_REVISION_INSERT: "insert",
        WD_REVISION_DELETE: "delete",
        WD_REVISION_MOVED_FROM: "moved_from",
        WD_REVISION_MOVED_TO: "moved_to",
    }

    return names.get(
        revision_type,
        f"word_revision_type_{revision_type}",
    )


def _ensure_redaction_style(
    document: Any,
    style_name: str,
) -> str:
    """
    Create or reuse the character style applied to restored deletions.
    """
    try:
        style = document.Styles(style_name)
    except Exception:
        style = document.Styles.Add(
            Name=style_name,
            Type=WD_STYLE_TYPE_CHARACTER,
        )

    style.Font.StrikeThrough = True
    style.Font.Color = WD_COLOR_RED

    return style_name


def _convert_insertion(revision: Any) -> None:
    """Highlight an insertion and accept it."""
    insertion_range = revision.Range.Duplicate

    insertion_range.HighlightColorIndex = WD_YELLOW

    revision.Accept()


def _convert_deletion(
    revision: Any,
    style_name: str,
    highlight_yellow: bool,
) -> tuple[bool, Optional[Dict[str, Any]]]:
    """
    Reject a deletion, restoring original text, then apply redaction style.
    """
    deletion_range = revision.Range.Duplicate

    start = getattr(deletion_range, "Start", None)
    end = getattr(deletion_range, "End", None)
    original_text = getattr(deletion_range, "Text", "")

    revision.Reject()

    restored_text = getattr(deletion_range, "Text", "")

    if original_text and not restored_text:
        return False, {
            "start": start,
            "end": end,
            "message": (
                "Deletion was rejected, but Word returned an empty restored "
                "range. The text was not styled automatically."
            ),
        }

    deletion_range.Style = style_name

    if highlight_yellow:
        deletion_range.HighlightColorIndex = WD_YELLOW

    return True, None


def _process_revision(
    revision: Any,
    options: PleaseReviewOptions,
    result: PleaseReviewResult,
) -> None:
    """Process one native Word revision."""
    revision_type = int(revision.Type)

    if revision_type == WD_REVISION_INSERT:
        _convert_insertion(revision)
        result.insertions_converted += 1
        return

    if revision_type == WD_REVISION_DELETE:
        _, warning = _convert_deletion(
            revision,
            options.redaction_style_name,
            options.deletion_highlight_yellow,
        )

        result.deletions_converted += 1

        if warning:
            result.deletion_validation_warnings += 1

            result.warnings.append(
                {
                    "category": "deletion_style_warning",
                    **warning,
                }
            )

        return

    if (
        revision_type == WD_REVISION_MOVED_TO
        and options.process_moved_text
    ):
        _convert_insertion(revision)
        result.moved_to_converted += 1
        return

    if (
        revision_type == WD_REVISION_MOVED_FROM
        and options.process_moved_text
    ):
        _, warning = _convert_deletion(
            revision,
            options.redaction_style_name,
            options.deletion_highlight_yellow,
        )

        result.moved_from_converted += 1

        if warning:
            result.deletion_validation_warnings += 1

            result.warnings.append(
                {
                    "category": "moved_from_style_warning",
                    **warning,
                }
            )

        return

    category = _revision_type_name(revision_type)

    result.unsupported_revisions += 1

    result.skipped_by_type[category] = (
        result.skipped_by_type.get(category, 0) + 1
    )


# ---------------------------------------------------------------------------
# Audit logging
# ---------------------------------------------------------------------------

def _write_audit_log(result: PleaseReviewResult) -> None:
    """Write a content-free JSON audit report."""
    audit_path = Path(result.audit_path)

    audit_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    payload = {
        "schema_version": AUDIT_SCHEMA_VERSION,
        "tool": "Doc Companion - Prepare PleaseReview Copy",
        "engine": (
            "Hybrid: OOXML formatting normalization plus "
            "native Microsoft Word text-revision processing"
        ),
        "cleanup_instruction": (
            "After approval, use Word Find/Replace, choose More, select "
            "Format > Style, select "
            f"'{DEFAULT_REDACTION_STYLE_NAME}', leave Replace With blank, "
            "and use Replace All only after final review approval."
        ),
        "result": result.to_dict(),
    }

    with audit_path.open(
        "w",
        encoding="utf-8",
    ) as audit_file:
        json.dump(
            payload,
            audit_file,
            indent=2,
            ensure_ascii=False,
        )


# ---------------------------------------------------------------------------
# Main public processing function
# ---------------------------------------------------------------------------

def prepare_please_review_copy(
    source_path: str | Path,
    output_path: str | Path | None = None,
    options: Optional[PleaseReviewOptions] = None,
    progress_callback: Optional[ProgressCallback] = None,
    cancel_event: Optional[threading.Event] = None,
) -> PleaseReviewResult:
    """
    Create a PleaseReview output copy.

    The original document is never modified.

    Text revisions are processed through Microsoft Word. Formatting/property
    revisions are flattened first in the output copy only, preventing Word from
    hanging on complex tracked formatting history.
    """
    _require_word()

    source = Path(source_path).expanduser().resolve()

    if not source.exists():
        raise PleaseReviewError(
            f"Source document does not exist: {source}"
        )

    if source.suffix.lower() not in {".docx", ".docm"}:
        raise PleaseReviewError(
            "Hybrid PleaseReview supports .docx and .docm files only. "
            "Save legacy .doc files as .docx before processing."
        )

    selected_options = options or PleaseReviewOptions()
    selected_options.validate()

    destination = (
        Path(output_path).expanduser().resolve()
        if output_path
        else build_output_path(source)
    )

    if destination == source:
        raise PleaseReviewError(
            "Output path must be different from the source document."
        )

    if destination.exists():
        raise PleaseReviewError(
            "Output file already exists. Choose a new output filename."
        )

    started_at = datetime.now(timezone.utc)

    result = PleaseReviewResult(
        source_path=str(source),
        output_path=str(destination),
        audit_path=str(build_audit_path(destination)),
        started_utc=started_at.isoformat(),
    )

    word_app = None
    document = None
    com_initialized = False

    try:
        # -------------------------------------------------------------------
        # Stage 1: Create output copy and flatten formatting revisions.
        # -------------------------------------------------------------------

        _emit_progress(
            progress_callback,
            {
                "kind": "log",
                "message": (
                    "Creating a safe output copy. The original document "
                    "will not be modified."
                ),
            },
        )

        destination.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        shutil.copy2(
            source,
            destination,
        )

        flattened_counts, cancelled = _normalise_format_revisions(
            destination,
            progress_callback,
            cancel_event,
        )

        result.formatting_revisions_flattened = sum(
            flattened_counts.values()
        )

        result.formatting_revisions_flattened_by_type = (
            flattened_counts
        )

        if cancelled:
            result.cancelled = True
            result.output_saved = True
            return result

        _emit_progress(
            progress_callback,
            {
                "kind": "log",
                "message": (
                    f"Accepted {result.formatting_revisions_flattened:,} "
                    "tracked formatting/property revisions in the output "
                    "copy. Text insertions and deletions remain tracked."
                ),
            },
        )

        # -------------------------------------------------------------------
        # Stage 2: Open the output copy in Word and process text revisions.
        # -------------------------------------------------------------------

        pythoncom.CoInitialize()
        com_initialized = True

        _emit_progress(
            progress_callback,
            {
                "kind": "log",
                "message": (
                    "Starting a dedicated Microsoft Word processing session "
                    "for text revisions."
                ),
            },
        )

        word_app = win32com.client.DispatchEx(
            "Word.Application"
        )

        result.word_version = str(
            getattr(word_app, "Version", "Unknown")
        )

        word_app.Visible = False
        word_app.DisplayAlerts = WD_ALERTS_NONE
        word_app.ScreenUpdating = False

        try:
            word_app.Options.Pagination = False
        except Exception:
            pass

        document = word_app.Documents.Open(
            FileName=str(destination),
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
            Visible=False,
            OpenAndRepair=False,
            NoEncodingDialog=True,
        )

        if bool(document.ReadOnly):
            raise PleaseReviewError(
                "Output copy opened read-only."
            )

        if int(document.ProtectionType) != WD_NO_PROTECTION:
            raise PleaseReviewError(
                "Document is protected."
            )

        result.source_track_revisions_enabled = bool(
            document.TrackRevisions
        )

        # Do not generate new tracked changes while styling output text.
        document.TrackRevisions = False

        selected_options.redaction_style_name = _ensure_redaction_style(
            document,
            selected_options.redaction_style_name,
        )

        revision_plan = _build_revision_plan(document)

        result.total_revisions = sum(
            item["revision_count"]
            for item in revision_plan
        )

        if result.total_revisions == 0:
            raise PleaseReviewError(
                "No text revisions remain after formatting normalization."
            )

        _emit_progress(
            progress_callback,
            {
                "kind": "started",
                "total_revisions": result.total_revisions,
                "message": (
                    f"Found {result.total_revisions:,} remaining Word "
                    "revisions after formatting normalization."
                ),
            },
        )

        # -------------------------------------------------------------------
        # Stage 3: Process revisions and calculate live ETA.
        # -------------------------------------------------------------------

        revision_stage_started = time.perf_counter()

        recent_durations: deque[float] = deque(maxlen=20)

        rolling_seconds_per_revision: Optional[float] = None

        ema_alpha = 0.25

        for story_item in revision_plan:
            story_name = story_item["story_name"]
            revisions = story_item["revisions"]
            revision_count = story_item["revision_count"]

            _emit_progress(
                progress_callback,
                {
                    "kind": "activity",
                    "message": (
                        f"Using cached Word revision collection for "
                        f"{story_name}: {revision_count:,} revisions."
                    ),
                },
            )

            for revision_index in range(
                revision_count,
                0,
                -1,
            ):
                if (
                    cancel_event is not None
                    and cancel_event.is_set()
                ):
                    result.cancelled = True
                    break

                remaining_before = (
                    result.total_revisions
                    - result.completed_revisions
                )

                estimated_before: Optional[float] = None

                if (
                    rolling_seconds_per_revision is not None
                    and result.completed_revisions >= 5
                ):
                    estimated_before = (
                        rolling_seconds_per_revision
                        * remaining_before
                    )

                elapsed_before = (
                    time.perf_counter()
                    - revision_stage_started
                )

                _emit_progress(
                    progress_callback,
                    {
                        "kind": "activity",
                        "message": (
                            f"Starting {story_name}, revision "
                            f"{revision_index:,} of "
                            f"{revision_count:,}. "
                            f"Elapsed: {_format_duration(elapsed_before)}. "
                            f"ETA: {_format_duration(estimated_before)}."
                        ),
                    },
                )

                revision_started = time.perf_counter()

                try:
                    revision = revisions.Item(revision_index)

                    revision_type = int(revision.Type)

                    _emit_progress(
                        progress_callback,
                        {
                            "kind": "activity",
                            "message": (
                                f"Processing {story_name}, revision "
                                f"{revision_index:,} of "
                                f"{revision_count:,} "
                                f"({_revision_type_name(revision_type)}). "
                                f"Elapsed: "
                                f"{_format_duration(elapsed_before)}. "
                                f"ETA: "
                                f"{_format_duration(estimated_before)}."
                            ),
                        },
                    )

                    _process_revision(
                        revision,
                        selected_options,
                        result,
                    )

                except Exception as exc:
                    result.unsupported_revisions += 1

                    result.skipped_by_type[
                        "revision_processing_error"
                    ] = (
                        result.skipped_by_type.get(
                            "revision_processing_error",
                            0,
                        )
                        + 1
                    )

                    result.warnings.append(
                        {
                            "category": "revision_processing_error",
                            "story": story_name,
                            "revision_index": revision_index,
                            "message": str(exc),
                        }
                    )

                revision_elapsed = (
                    time.perf_counter()
                    - revision_started
                )

                recent_durations.append(revision_elapsed)

                if rolling_seconds_per_revision is None:
                    rolling_seconds_per_revision = revision_elapsed
                else:
                    rolling_seconds_per_revision = (
                        ema_alpha * revision_elapsed
                        + (1 - ema_alpha)
                        * rolling_seconds_per_revision
                    )

                if (
                    revision_elapsed
                    >= selected_options.slow_revision_warning_seconds
                ):
                    result.warnings.append(
                        {
                            "category": "slow_revision_operation",
                            "story": story_name,
                            "revision_index": revision_index,
                            "seconds": round(revision_elapsed, 3),
                        }
                    )

                result.completed_revisions += 1

                elapsed_stage_seconds = (
                    time.perf_counter()
                    - revision_stage_started
                )

                remaining_revisions = (
                    result.total_revisions
                    - result.completed_revisions
                )

                estimated_remaining: Optional[float] = None

                if result.completed_revisions >= 5:
                    estimated_remaining = (
                        rolling_seconds_per_revision
                        * remaining_revisions
                    )

                if result.completed_revisions < 5:
                    confidence = "Warming up"
                elif result.completed_revisions < 15:
                    confidence = "Low"
                elif len(recent_durations) < 20:
                    confidence = "Moderate"
                else:
                    confidence = "Good"

                if elapsed_stage_seconds > 0:
                    revisions_per_minute = (
                        result.completed_revisions
                        / elapsed_stage_seconds
                    ) * 60
                else:
                    revisions_per_minute = 0.0

                percent = (
                    result.completed_revisions
                    / result.total_revisions
                ) * 100

                # Emit progress after every revision so the bar and ETA remain
                # live even if the UI still has progress_interval=25 configured.
                _emit_progress(
                    progress_callback,
                    {
                        "kind": "progress",
                        "completed_revisions": (
                            result.completed_revisions
                        ),
                        "total_revisions": result.total_revisions,
                        "percent": percent,
                        "elapsed_seconds": elapsed_stage_seconds,
                        "estimated_remaining_seconds": (
                            estimated_remaining
                        ),
                        "revisions_per_minute": revisions_per_minute,
                        "eta_confidence": confidence,
                        "message": (
                            f"Processed "
                            f"{result.completed_revisions:,} of "
                            f"{result.total_revisions:,} text revisions "
                            f"({percent:.1f}%). "
                            f"Elapsed: "
                            f"{_format_duration(elapsed_stage_seconds)}. "
                            f"ETA: "
                            f"{_format_duration(estimated_remaining)} "
                            f"({confidence} confidence)."
                        ),
                    },
                )

            if result.cancelled:
                break

        # Save the native Word-generated output copy.
        document.TrackRevisions = False
        document.Save()

        result.output_saved = True
        result.success = not result.cancelled

    except Exception as exc:
        result.errors.extend(
            [
                str(exc),
                traceback.format_exc(),
            ]
        )

        _emit_progress(
            progress_callback,
            {
                "kind": "error",
                "message": (
                    f"PleaseReview preparation failed: {exc}"
                ),
            },
        )

    finally:
        _safe_close_document(document)
        _safe_quit_word(word_app)

        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

        completed_at = datetime.now(timezone.utc)

        result.completed_utc = completed_at.isoformat()

        result.elapsed_seconds = round(
            (completed_at - started_at).total_seconds(),
            3,
        )

        result.skipped_by_type = dict(
            Counter(result.skipped_by_type)
        )

        try:
            _write_audit_log(result)
        except Exception as exc:
            result.errors.append(
                f"Could not write audit file: {exc}"
            )

    return result
