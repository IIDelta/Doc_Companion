"""Reference reconciliation, EndNote action queue, email requests, and inline citations."""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Callable, Iterable, Optional
import csv
import json
import re
import time
import unicodedata
import xml.etree.ElementTree as ET
import zipfile

REFERENCECHECKVERSION = "4.0.0-final-workflow"

WORDNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": WORDNS}

DOI_PATTERN = re.compile(
    r"(?:https?://(?:dx\.)?doi\.org/|doi:\s*)?"
    r"(10\.\d{4,9}/[-._;()/:a-z0-9]+)",
    re.I,
)
URL_PATTERN = re.compile(
    r"(?<!@)(?:https?://|www\.)[^\s<>\]\[(){}\"']+",
    re.I,
)
YEAR_PATTERN = re.compile(r"\b(?:18|19|20)\d{2}\b")
VANCOUVER_YEAR_PATTERN = re.compile(
    r"\.\s*((?:18|19|20)\d{2})(?=\s*;\s*\d)"
)
ENDNOTE_YEAR_PATTERN = re.compile(
    r"^[^.]{3,}\.\s*((?:18|19|20)\d{2})\."
)
PAREN_YEAR_REFERENCE_PATTERN = re.compile(
    r"^(.+?)\s*,?\s*\(((?:18|19|20)\d{2}[a-z]?)\)\s*(.+)$",
    re.I,
)

NUMBERED_REFERENCE_PATTERN = re.compile(
    r"^\s*(\d+)\s*[.)]\s*(.+)$"
)
PAGES_PATTERN = re.compile(
    r"\b\d{1,5}\s*[–-]\s*\d{1,5}\b"
)
PARENTHETICAL_PATTERN = re.compile(
    r"\([^()]{1,300}\)"
)

AUTHOR_YEAR_PAREN_PATTERN = re.compile(
    r"\(([^()]{2,160}?"
    r"(?:et al\.?|and\s+[A-Z][A-Za-z'’\-]+|[A-Z][A-Za-z'’\-]+)"
    r"\s*,?\s*(?:18|19|20)\d{2})\)",
    re.I,
)

AUTHOR_YEAR_NARRATIVE_PATTERN = re.compile(
    r"\b("
    r"[A-Z][A-Za-z'’\-]+(?:\s+et\s+al\.)?"
    r"|"
    r"[A-Z][A-Za-z'’\-]+\s+and\s+[A-Z][A-Za-z'’\-]+"
    r")\s*\(((?:18|19|20)\d{2})\)"
)

ENDNOTEID_PATTERN = re.compile(
    r"\[\[ENDNOTEID:(\d+)\]\]",
    re.I,
)

REFERENCE_HEADINGS = {
    "references",
    "reference list",
    "bibliography",
    "literature cited",
}

INLINE_ZOTERO_TYPES = {
    "webpage",
    "blogpost",
    "forumpost",
    "post",
    "web",
}

SOURCE_ORGANIZATIONS = {
    "orphanet": (
        "Website/webpage",
        "Orphanet",
    ),
    "food and drug administration": (
        "Government/regulatory document",
        "US Food and Drug Administration",
    ),
    "fda": (
        "Government/regulatory document",
        "FDA",
    ),
    "european medicines agency": (
        "Government/regulatory document",
        "European Medicines Agency",
    ),
    "ema": (
        "Government/regulatory document",
        "EMA",
    ),
    "international council for harmonisation": (
        "Government/regulatory document",
        "ICH",
    ),
    "ich guidance": (
        "Government/regulatory document",
        "ICH",
    ),
    "medicines and healthcare products regulatory agency": (
        "Government/regulatory document",
        "MHRA",
    ),
    "mhra": (
        "Government/regulatory document",
        "MHRA",
    ),
    "world health organization": (
        "Government/regulatory document",
        "World Health Organization",
    ),
    "eu clinical trials register": (
        "Government/regulatory document",
        "EU Clinical Trials Register",
    ),
    "eudract": (
        "Government/regulatory document",
        "EU Clinical Trials Register",
    ),
}


@dataclass
class Paragraph:
    number: int
    text: str
    section: str = ""


@dataclass
class Record:
    raw: str
    source: str
    title: str = ""
    firstauthor: str = ""
    year: str = ""
    doi: str = ""
    journal: str = ""
    pages: str = ""
    nativeid: str = ""
    paragraph: int = 0
    section: str = ""
    context: str = ""
    referenceblock: int = 0
    referencenumber: str = ""
    kind: str = "reference"
    itemtype: str = ""
    url: str = ""
    accessed: str = ""
    endnoteid: str = ""


    def label(self) -> str:
        author = self.firstauthor or "Unknown author"
        year = self.year or "n.d."
        title = self.title or self.raw[:90]
        return f"{author} ({year}). {title}"

    def key(self) -> tuple[str, str, str, str]:
        return (
            self.doi,
            norm(self.title),
            norm(self.firstauthor),
            self.year,
        )

    def endnote_search(self) -> str:
        lines = []

        if self.firstauthor:
            lines.append(f"Author: {self.firstauthor}")

        if self.year:
            lines.append(f"Year: {self.year}")

        if self.title:
            lines.append(f"Title: {self.title}")

        if self.doi:
            lines.append(f"DOI: {self.doi}")

        if self.journal:
            lines.append(f"Journal: {self.journal}")

        if self.endnoteid:
            lines.append(f"EndNote Record Number: {self.endnoteid}")

        return "\n".join(lines) or self.raw


@dataclass
class Candidate:
    record: Record
    score: float
    reason: str


@dataclass
class Finding:
    category: str
    documentrecord: Record
    action: str
    bestmatch: Optional[Record] = None
    score: float = 0.0
    reason: str = ""
    candidates: list[Candidate] = field(default_factory=list)

    def row(self) -> dict:
        endnote_record = self.bestmatch or self.documentrecord

        return {
            "category": self.category,
            "action": self.action,
            "document": self.documentrecord.label(),
            "bestmatch": (
                self.bestmatch.label()
                if self.bestmatch
                else ""
            ),
            "score": round(self.score * 100, 1),
            "reason": self.reason,
            "suggestedcitation": "",
            "endnotesearch": endnote_record.endnote_search(),
            "section": self.documentrecord.section,
            "paragraph": self.documentrecord.paragraph,
            "referenceblock": self.documentrecord.referenceblock,
            "referencenumber": self.documentrecord.referencenumber,
            "context": self.documentrecord.context,
            "missingmetadata": "",
            "rawdocument": self.documentrecord.raw,
            "rawmatch": (
                self.bestmatch.raw
                if self.bestmatch
                else ""
            ),
        }


@dataclass
class InlineSuggestion:
    category: str
    detected: str
    suggestion: str
    reason: str
    missingmetadata: str
    paragraph: int = 0
    section: str = ""
    context: str = ""
    source: str = "document text"
    relatedreference: str = ""

    def key(self) -> tuple[str, str, int]:
        return (
            self.category,
            norm(self.detected),
            self.paragraph,
        )

    def row(self) -> dict:
        return {
            "category": self.category,
            "action": "Use Style Guide inline citation",
            "document": self.detected,
            "bestmatch": self.relatedreference,
            "score": "",
            "reason": self.reason,
            "suggestedcitation": self.suggestion,
            "endnotesearch": "",
            "section": self.section,
            "paragraph": self.paragraph,
            "referenceblock": "",
            "referencenumber": "",
            "context": self.context,
            "missingmetadata": self.missingmetadata,
            "rawdocument": self.detected,
            "rawmatch": self.relatedreference,
        }


@dataclass
class ParseWarning:
    message: str
    paragraph: int = 0
    section: str = ""
    context: str = ""

    def row(self) -> dict:
        return {
            "category": "Parser warning",
            "action": "Review parser result",
            "document": self.message,
            "bestmatch": "",
            "score": "",
            "reason": self.message,
            "suggestedcitation": "",
            "endnotesearch": "",
            "section": self.section,
            "paragraph": self.paragraph,
            "referenceblock": "",
            "referencenumber": "",
            "context": self.context,
            "missingmetadata": "",
            "rawdocument": self.context,
            "rawmatch": "",
        }


@dataclass
class CheckResult:
    matched: list[Finding] = field(default_factory=list)
    review: list[Finding] = field(default_factory=list)
    missing: list[Finding] = field(default_factory=list)
    potentialmissing: list[Finding] = field(
        default_factory=list
    )
    existingendnote: list[Record] = field(
        default_factory=list
    )
    inlineonly: list[InlineSuggestion] = field(
        default_factory=list
    )
    warnings: list[ParseWarning] = field(
        default_factory=list
    )
    documentrecords: int = 0
    libraryrecords: int = 0
    nativecitations: int = 0
    referenceblocks: int = 0
    cancelled: bool = False

    def summary(self) -> dict:
        return {
            "documentrecords": self.documentrecords,
            "libraryrecords": self.libraryrecords,
            "nativecitations": self.nativecitations,
            "referenceblocks": self.referenceblocks,
            "matched": len(self.matched),
            "review": len(self.review),
            "missing": len(self.missing),
            "potentialmissing": len(
                self.potentialmissing
            ),
            "existingendnote": len(
                self.existingendnote
            ),
            "inlineonly": len(self.inlineonly),
            "warnings": len(self.warnings),
            "cancelled": self.cancelled,
        }


@dataclass
class LibraryIndex:
    bydoi: dict[str, list[Record]] = field(
        default_factory=lambda: defaultdict(list)
    )
    bytitle: dict[str, list[Record]] = field(
        default_factory=lambda: defaultdict(list)
    )
    byauthoryear: dict[
        tuple[str, str],
        list[Record],
    ] = field(
        default_factory=lambda: defaultdict(list)
    )
    byauthor: dict[str, list[Record]] = field(
        default_factory=lambda: defaultdict(list)
    )
    bytitleprefix: dict[str, list[Record]] = field(
        default_factory=lambda: defaultdict(list)
    )


def norm(value: str) -> str:
    value = unicodedata.normalize(
        "NFKD",
        value or "",
    )

    value = "".join(
        character
        for character in value
        if not unicodedata.combining(character)
    )

    value = value.casefold()
    value = value.replace("–", "-")
    value = value.replace("—", "-")
    value = re.sub(r"[^a-z0-9\s]", " ", value)

    return re.sub(r"\s+", " ", value).strip()


def normalizeddoi(value: str) -> str:
    match = DOI_PATTERN.search(value or "")

    if not match:
        return ""

    return match.group(1).rstrip(
        ".,;)"
    ).lower()


def compactpages(value: str) -> str:
    return re.sub(
        r"[^0-9]",
        "",
        value or "",
    )


def firstauthor(value: str) -> str:
    value = re.sub(
        r"\s+et\s+al\.?$",
        "",
        value.strip(),
        flags=re.I,
    )

    if not value:
        return ""

    if "," in value:
        value = value.split(",", 1)[0].strip()

        value = re.sub(
            r"\s+[A-Z]{1,3}$",
            "",
            value,
        )

        return norm(value)

    return norm(value.split()[0])


def isheading(text: str) -> bool:
    text = text.strip()

    if (
        not text
        or len(text) > 150
        or text.endswith((".", ";", ":"))
    ):
        return False

    words = text.split()

    if len(words) > 15:
        return False

    letters = [
        character
        for character in text
        if character.isalpha()
    ]

    if not letters:
        return False

    uppercount = sum(
        character.isupper()
        for character in letters
    )

    return (
        uppercount / len(letters) > 0.60
        or len(words) <= 8
    )


def buildsectionmap(
    texts: list[str],
) -> list[Paragraph]:
    currentsection = ""
    output: list[Paragraph] = []

    for number, text in enumerate(
        texts,
        start=1,
    ):
        if (
            isheading(text)
            and norm(text) not in REFERENCE_HEADINGS
        ):
            currentsection = text

        output.append(
            Paragraph(
                number=number,
                text=text,
                section=currentsection,
            )
        )

    return output


def splitrecords(text: str) -> list[str]:
    chunks = [
        re.sub(r"\s+", " ", item).strip()
        for item in re.split(r"\n\s*\n+", text)
    ]

    chunks = [
        item
        for item in chunks
        if item
    ]

    if len(chunks) > 1:
        return chunks

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    entries: list[str] = []
    buffer: list[str] = []

    startpattern = re.compile(
        r"^[A-ZÀ-ÖØ-Þ]"
        r"[A-Za-zÀ-ÖØ-öø-ÿ'’\- ]{1,80},"
        r".*\b(?:18|19|20)\d{2}\."
    )

    for line in lines:
        if buffer and startpattern.match(line):
            entries.append(" ".join(buffer))
            buffer = [line]
        else:
            buffer.append(line)

    if buffer:
        entries.append(" ".join(buffer))

    return entries


def publicationyear(text: str) -> str:
    vancouver = VANCOUVER_YEAR_PATTERN.search(text)

    if vancouver:
        return vancouver.group(1)

    endnote = ENDNOTE_YEAR_PATTERN.search(text)

    if endnote:
        return endnote.group(1)

    years = YEAR_PATTERN.findall(text)

    return years[-1] if years else ""


def parsereference(
    raw: str,
    source: str,
    *,
    paragraph: int = 0,
    section: str = "",
    referenceblock: int = 0,
    referencenumber: str = "",
    context: str = "",
    kind: str = "reference",
) -> Record:
    endnoteid_match = ENDNOTEID_PATTERN.search(raw)
    endnoteid = (
        endnoteid_match.group(1)
        if endnoteid_match
        else ""
    )

    raw_without_id = ENDNOTEID_PATTERN.sub(
        "",
        raw,
    )

    cleaned = re.sub(
        r"\s+",
        " ",
        raw_without_id,
    ).strip()

    numbered = NUMBERED_REFERENCE_PATTERN.match(
        cleaned
    )

    if numbered:
        referencenumber = (
            referencenumber
            or numbered.group(1)
        )
        cleaned = numbered.group(2)

    doi = normalizeddoi(cleaned)
    year = publicationyear(cleaned)

    author = ""
    title = ""
    journal = ""

    pagesmatch = PAGES_PATTERN.search(cleaned)

    pages = (
        pagesmatch.group(0)
        if pagesmatch
        else ""
    )

    vancouver = VANCOUVER_YEAR_PATTERN.search(
        cleaned
    )

    endnote = ENDNOTE_YEAR_PATTERN.search(
        cleaned
    )

    parenreference = PAREN_YEAR_REFERENCE_PATTERN.match(
        cleaned
    )

    if parenreference:
        author = firstauthor(
            parenreference.group(1)
        )

        remaining = parenreference.group(3)

        pieces = re.split(
            r"\.\s+(?=[A-Z])",
            remaining,
            maxsplit=1,
        )

        title = pieces[0].strip(" .")

        if len(pieces) > 1:
            journal = re.split(
                r"[,;]\s*\d",
                pieces[1],
                maxsplit=1,
            )[0].strip(" .")

    elif vancouver:
        beforeyear = cleaned[
            :vancouver.start()
        ].strip()

        sentences = [
            part.strip()
            for part in re.split(
                r"\.\s+",
                beforeyear,
            )
            if part.strip()
        ]

        if sentences:
            author = firstauthor(sentences[0])

        if len(sentences) >= 2:
            title = sentences[1]

        if len(sentences) >= 3:
            journal = sentences[-1]

    elif endnote:
        authorpart = cleaned[
            :endnote.start()
        ].strip(" .")

        author = firstauthor(authorpart)

        remaining = cleaned[
            endnote.end():
        ].strip()

        pieces = re.split(
            r"\.\s+(?=[A-Z0-9])",
            remaining,
            maxsplit=1,
        )

        title = (
            pieces[0].strip(" .")
            if pieces
            else ""
        )

        if len(pieces) > 1:
            journal = re.split(
                r"[,;]\s*\d",
                pieces[1],
                maxsplit=1,
            )[0].strip(" .")

    else:
        yearmatch = YEAR_PATTERN.search(cleaned)

        beforeyear = (
            cleaned[:yearmatch.start()].strip(
                " .;,"
            )
            if yearmatch
            else cleaned
        )

        author = firstauthor(beforeyear)

        remaining = (
            cleaned[yearmatch.end():].lstrip(
                " .,:-"
            )
            if yearmatch
            else ""
        )

        if remaining:
            title = re.split(
                r"\.\s+(?=[A-Z0-9])",
                remaining,
                maxsplit=1,
            )[0].strip(" .")

    if not title:
        title = cleaned

    return Record(
        raw=raw_without_id.strip(),
        source=source,
        title=title,
        firstauthor=author,
        year=year,
        doi=doi,
        journal=journal,
        pages=pages,
        paragraph=paragraph,
        section=section,
        context=context,
        referenceblock=referenceblock,
        referencenumber=referencenumber,
        kind=kind,
        endnoteid=endnoteid,
    )


def parsereferencetext(
    text: str,
    source: str,
) -> list[Record]:
    return [
        parsereference(item, source)
        for item in splitrecords(text)
        if item
    ]


def isreferencelike(text: str) -> bool:
    if not text or len(text) < 20:
        return False

    if NUMBERED_REFERENCE_PATTERN.match(text):
        return True

    if not YEAR_PATTERN.search(text):
        return False

    beginsauthor = bool(
        re.match(
            r"^[A-ZÀ-ÖØ-Þ]"
            r"[A-Za-zÀ-ÖØ-öø-ÿ'’\-]+"
            r"(?:,|\s*\()",
            text,
        )
    )

    hasbibliographicpunctuation = (
        text.count(",") >= 1
        or text.count(".") >= 2
    )

    return beginsauthor and hasbibliographicpunctuation


def parsereferenceblocks(
    paragraphs: list[Paragraph],
) -> tuple[
    list[Record],
    list[ParseWarning],
    int,
]:
    records: list[Record] = []
    warnings: list[ParseWarning] = []
    blockcount = 0

    index = 0

    while index < len(paragraphs):
        heading = paragraphs[index]

        if norm(heading.text) not in REFERENCE_HEADINGS:
            index += 1
            continue

        blockcount += 1
        blocknumber = blockcount
        cursor = index + 1

        while (
            cursor < len(paragraphs)
            and not paragraphs[cursor].text.strip()
        ):
            cursor += 1

        numberedmode = (
            cursor < len(paragraphs)
            and bool(
                NUMBERED_REFERENCE_PATTERN.match(
                    paragraphs[cursor].text
                )
            )
        )

        added = 0

        while cursor < len(paragraphs):
            current = paragraphs[cursor]
            text = current.text.strip()

            if not text:
                cursor += 1
                continue

            if norm(text) in REFERENCE_HEADINGS:
                break

            if numberedmode:
                match = NUMBERED_REFERENCE_PATTERN.match(
                    text
                )

                if not match:
                    break

                referencenumber = match.group(1)

            else:
                if not isreferencelike(text):
                    break

                referencenumber = ""

            records.append(
                parsereference(
                    text,
                    "documentbibliography",
                    paragraph=current.number,
                    section=current.section,
                    referenceblock=blocknumber,
                    referencenumber=referencenumber,
                    context=text,
                )
            )

            added += 1
            cursor += 1

        if not added:
            warnings.append(
                ParseWarning(
                    message=(
                        "References heading found but no consecutive "
                        "bibliographic entries were detected."
                    ),
                    paragraph=heading.number,
                    section=heading.section,
                    context=heading.text,
                )
            )

        index = cursor

    return records, warnings, blockcount


def parsezotero(
    instruction: str,
    paragraph: int = 0,
    section: str = "",
    context: str = "",
) -> list[Record]:
    marker = "CSL_CITATION"

    if marker not in instruction:
        return []

    try:
        payload = json.loads(
            instruction.split(marker, 1)[1].strip()
        )

    except (
        json.JSONDecodeError,
        IndexError,
    ):
        return []

    records: list[Record] = []

    for item in payload.get(
        "citationItems",
        [],
    ):
        data = item.get("itemData", {}) or {}

        authors = (
            data.get("author", [])
            or data.get("editor", [])
            or []
        )

        author = (
            norm(authors[0].get("family", ""))
            if authors
            else ""
        )

        dates = data.get(
            "issued",
            {},
        ).get(
            "date-parts",
            [],
        )

        year = (
            str(dates[0][0])
            if dates and dates[0]
            else ""
        )

        accessed = data.get(
            "accessed",
            {},
        ).get(
            "raw",
            "",
        )

        raw = (
            data.get("title", "")
            or payload.get(
                "properties",
                {},
            ).get(
                "plainCitation",
                "",
            )
        )

        records.append(
            Record(
                raw=raw,
                source="zotero",
                title=data.get("title", ""),
                firstauthor=author,
                year=year,
                doi=normalizeddoi(
                    data.get("DOI", "")
                ),
                journal=data.get(
                    "container-title",
                    "",
                ),
                nativeid=str(
                    item.get("id", "")
                ),
                paragraph=paragraph,
                section=section,
                context=context,
                kind="nativecitation",
                itemtype=data.get("type", ""),
                url=data.get("URL", ""),
                accessed=str(accessed or ""),
            )
        )

    return records


def dedupe(
    records: Iterable[Record],
) -> list[Record]:
    seen: set[
        tuple[str, str, str, str]
    ] = set()

    output: list[Record] = []

    for record in records:
        key = record.key()

        if key == ("", "", "", ""):
            key = (
                "",
                norm(record.raw),
                "",
                "",
            )

        if key not in seen:
            seen.add(key)
            output.append(record)

    return output


def readword(
    path: str | Path,
) -> tuple[
    list[Record],
    list[Record],
    list[Record],
    list[Record],
    list[Paragraph],
    list[ParseWarning],
    int,
    int,
]:
    path = Path(path)

    if path.suffix.lower() != ".docx":
        raise ValueError(
            "Reference Library Check requires a .docx Word document."
        )

    try:
        with zipfile.ZipFile(path) as archive:
            documentxml = archive.read(
                "word/document.xml"
            )

    except (
        KeyError,
        zipfile.BadZipFile,
    ) as exc:
        raise ValueError(
            "The selected file is not a readable .docx document."
        ) from exc

    root = ET.fromstring(documentxml)

    visibleparagraphs: list[str] = []

    for paragraph in root.findall(
        ".//w:body/w:p",
        NS,
    ):
        # w:t contains visible text only.
        visibletext = "".join(
            node.text or ""
            for node in paragraph.findall(
                ".//w:t",
                NS,
            )
        ).strip()

        visibleparagraphs.append(visibletext)

    paragraphs = buildsectionmap(
        visibleparagraphs
    )

    (
        bibliography,
        warnings,
        referenceblockcount,
    ) = parsereferenceblocks(paragraphs)

    bibliographyparagraphs = {
        record.paragraph
        for record in bibliography
    }

    paragraphlookup: dict[
        str,
        Paragraph,
    ] = {}

    for xmlparagraph in root.findall(
        ".//w:p",
        NS,
    ):
        instruction = "".join(
            node.text or ""
            for node in xmlparagraph.findall(
                ".//w:instrText",
                NS,
            )
        )

        visibletext = "".join(
            node.text or ""
            for node in xmlparagraph.findall(
                ".//w:t",
                NS,
            )
        ).strip()

        if instruction:
            matched = next(
                (
                    paragraph
                    for paragraph in paragraphs
                    if paragraph.text == visibletext
                ),
                Paragraph(
                    number=0,
                    text=visibletext,
                    section="",
                ),
            )

            paragraphlookup[instruction] = matched

    native: list[Record] = []
    existingendnote: list[Record] = []
    fieldcount = 0

    for instructionnode in root.findall(
        ".//w:instrText",
        NS,
    ):
        instruction = instructionnode.text or ""

        if "CSL_CITATION" in instruction:
            fieldcount += 1

            matched = paragraphlookup.get(
                instruction,
                Paragraph(
                    number=0,
                    text="",
                    section="",
                ),
            )

            native.extend(
                parsezotero(
                    instruction,
                    paragraph=matched.number,
                    section=matched.section,
                    context=matched.text,
                )
            )

        elif "ADDIN EN.CITE" in instruction.upper():
            fieldcount += 1

            existingendnote.append(
                Record(
                    raw=(
                        "Existing EndNote citation field"
                    ),
                    source="endnote",
                    kind="existingendnote",
                )
            )

    plain: list[Record] = []

    for paragraph in paragraphs:
        if paragraph.number in bibliographyparagraphs:
            continue

        for match in AUTHOR_YEAR_PAREN_PATTERN.finditer(
            paragraph.text
        ):
            plain.append(
                parsereference(
                    match.group(1),
                    "plaintext",
                    paragraph=paragraph.number,
                    section=paragraph.section,
                    context=paragraph.text,
                    kind="plaintextcitation",
                )
            )

        for match in AUTHOR_YEAR_NARRATIVE_PATTERN.finditer(
            paragraph.text
        ):
            raw = (
                f"{match.group(1)} "
                f"({match.group(2)})"
            )

            plain.append(
                parsereference(
                    raw,
                    "plaintext",
                    paragraph=paragraph.number,
                    section=paragraph.section,
                    context=paragraph.text,
                    kind="plaintextcitation",
                )
            )

    return (
        dedupe(native),
        bibliography,
        dedupe(plain),
        existingendnote,
        paragraphs,
        warnings,
        fieldcount,
        referenceblockcount,
    )


def similarity(
    left: str,
    right: str,
) -> float:
    leftnorm = norm(left)
    rightnorm = norm(right)

    if not leftnorm or not rightnorm:
        return 0.0

    sequence = SequenceMatcher(
        None,
        leftnorm,
        rightnorm,
    ).ratio()

    lefttokens = set(leftnorm.split())
    righttokens = set(rightnorm.split())

    jaccard = len(
        lefttokens & righttokens
    ) / max(
        1,
        len(lefttokens | righttokens),
    )

    return (
        0.60 * sequence
        + 0.40 * jaccard
    )


def scorematch(
    document: Record,
    library: Record,
) -> Candidate:
    if (
        document.doi
        and library.doi
        and document.doi == library.doi
    ):
        return Candidate(
            library,
            1.0,
            "Exact DOI match",
        )

    titleexact = bool(
        document.title
        and norm(document.title)
        == norm(library.title)
    )

    authorsame = bool(
        document.firstauthor
        and document.firstauthor
        == library.firstauthor
    )

    yearsame = bool(
        document.year
        and document.year == library.year
    )

    if titleexact and authorsame and yearsame:
        return Candidate(
            library,
            1.0,
            (
                "Exact normalized title, "
                "first author, and year"
            ),
        )

    if document.kind == "plaintextcitation":
        score = (
            0.70 * float(authorsame)
            + 0.30 * float(yearsame)
        )

        return Candidate(
            library,
            score,
            (
                "first author and year"
                if score == 1.0
                else "partial author-year match"
            ),
        )

    titlescore = similarity(
        document.title,
        library.title,
    )

    journalscore = (
        similarity(
            document.journal,
            library.journal,
        )
        if (
            document.journal
            and library.journal
        )
        else 0.0
    )

    pagescore = float(
        bool(
            compactpages(document.pages)
            and compactpages(document.pages)
            == compactpages(library.pages)
        )
    )

    score = (
        0.65 * titlescore
        + 0.15 * float(authorsame)
        + 0.10 * float(yearsame)
        + 0.05 * journalscore
        + 0.05 * pagescore
    )

    reasons = [
        f"title {titlescore:.0%}",
    ]

    if authorsame:
        reasons.append("first author")

    if yearsame:
        reasons.append("year")

    if journalscore >= 0.70:
        reasons.append(
            f"journal {journalscore:.0%}"
        )

    if pagescore:
        reasons.append("pages")

    return Candidate(
        library,
        score,
        ", ".join(reasons),
    )


def titleprefix(value: str) -> str:
    terms = [
        term
        for term in norm(value).split()
        if len(term) >= 4
    ]

    return " ".join(terms)[:120]


def formateta(seconds: float) -> str:
    seconds = max(0, int(seconds))

    minutes, seconds = divmod(
        seconds,
        60,
    )

    hours, minutes = divmod(
        minutes,
        60,
    )

    if hours:
        return (
            f"{hours:d}:"
            f"{minutes:02d}:"
            f"{seconds:02d}"
        )

    return f"{minutes:d}:{seconds:02d}"


def buildindex(
    library: list[Record],
    cancel: Callable[[], bool],
    progress: Callable[[int, int, str], None],
    totalwork: int,
) -> LibraryIndex:
    index = LibraryIndex()
    started = time.monotonic()

    for number, item in enumerate(
        library,
        start=1,
    ):
        if cancel():
            raise InterruptedError(
                "Reference Library Check cancelled during indexing."
            )

        if item.doi:
            index.bydoi[item.doi].append(item)

        normaltitle = norm(item.title)

        if normaltitle:
            index.bytitle[normaltitle].append(
                item
            )

        if item.firstauthor:
            index.byauthor[
                item.firstauthor
            ].append(item)

            if item.year:
                index.byauthoryear[
                    (
                        item.firstauthor,
                        item.year,
                    )
                ].append(item)

        prefix = titleprefix(item.title)

        if prefix:
            index.bytitleprefix[prefix].append(
                item
            )

        if (
            number == 1
            or number % 500 == 0
            or number == len(library)
        ):
            elapsed = time.monotonic() - started

            eta = (
                (elapsed / number)
                * (len(library) - number)
                if number
                else 0
            )

            progress(
                number,
                totalwork,
                (
                    f"Indexing shared library: "
                    f"{number:,}/{len(library):,} records "
                    f"— ETA {formateta(eta)}"
                ),
            )

    return index


def indexedcandidates(
    document: Record,
    index: LibraryIndex,
    limit: int = 3,
) -> list[Candidate]:
    pool: list[Record] = []

    if (
        document.doi
        and document.doi in index.bydoi
    ):
        pool.extend(
            index.bydoi[document.doi]
        )

    if not pool:
        normaltitle = norm(document.title)

        if (
            normaltitle
            and normaltitle in index.bytitle
        ):
            pool.extend(
                index.bytitle[normaltitle]
            )

    if (
        not pool
        and document.firstauthor
        and document.year
    ):
        pool.extend(
            index.byauthoryear.get(
                (
                    document.firstauthor,
                    document.year,
                ),
                [],
            )
        )

    if not pool and document.firstauthor:
        pool.extend(
            index.byauthor.get(
                document.firstauthor,
                [],
            )
        )

    if not pool and titleprefix(document.title):
        pool.extend(
            index.bytitleprefix.get(
                titleprefix(document.title),
                [],
            )
        )

    candidates = dedupe(pool)[:2000]

    ranked = [
        scorematch(document, item)
        for item in candidates
    ]

    ranked.sort(
        key=lambda item: item.score,
        reverse=True,
    )

    return ranked[:limit]


def stripurl(url: str) -> str:
    return re.sub(
        r"^(?:https?://)?(?:www\.)?",
        "",
        url.strip(),
        flags=re.I,
    ).rstrip(".,;:)")


def websitecitation(
    url: str,
    title: str = "",
    accessed: str = "",
) -> tuple[str, str]:
    shownurl = stripurl(url)

    title = title.strip() or "[Webpage title]"
    accessed = accessed.strip() or "[DD Month YYYY]"

    missing = []

    if title.startswith("["):
        missing.append("Confirm webpage title")

    if accessed.startswith("["):
        missing.append("Enter access date")

    return (
        (
            f"({shownurl}, "
            f"{title}, "
            f"Accessed {accessed})"
        ),
        (
            "; ".join(missing)
            or "Ready for reviewer confirmation"
        ),
    )


def governmentcitation(
    agency: str,
    text: str,
) -> tuple[str, str]:
    year = publicationyear(text) or "[year]"

    if "eudract" in norm(text):
        match = re.search(
            r"\b\d{4}-\d{6}-\d{2}\b",
            text,
        )

        identifier = (
            f"EudraCT number: {match.group(0)}"
            if match
            else "EudraCT number: [number]"
        )

        return (
            (
                f"(EU Clinical Trials Register, "
                f"{identifier}, {year})"
            ),
            (
                "Confirm publication year"
                if year == "[year]"
                else "Ready for reviewer confirmation"
            ),
        )

    return (
        (
            f"({agency}, "
            f"[document title or number], "
            f"{year})"
        ),
        (
            "Confirm document title or number"
            + (
                "; confirm year"
                if year == "[year]"
                else ""
            )
        ),
    )


def nativeinline(record: Record) -> bool:
    return (
        record.itemtype.casefold()
        in INLINE_ZOTERO_TYPES
    )


def relatedbibliography(
    detected: str,
    bibliography: list[Record],
) -> str:
    detectednorm = norm(detected)

    for record in bibliography:
        if not detectednorm:
            continue

        if (
            detectednorm in norm(record.raw)
            or norm(record.title) == detectednorm
        ):
            label = (
                f"Existing bibliography entry: "
                f"Block {record.referenceblock}"
            )

            if record.referencenumber:
                label += (
                    f", reference "
                    f"{record.referencenumber}"
                )

            return (
                f"{label} — {record.raw}"
            )

    return ""


def detectinlinesources(
    paragraphs: list[Paragraph],
    native: list[Record],
    bibliography: list[Record],
) -> list[InlineSuggestion]:
    suggestions: list[InlineSuggestion] = []

    seen: set[
        tuple[str, str, int]
    ] = set()

    def add(item: InlineSuggestion) -> None:
        if item.key() not in seen:
            seen.add(item.key())
            suggestions.append(item)

    for record in native:
        if nativeinline(record):
            suggestion, missing = websitecitation(
                record.url or "[website address]",
                record.title,
                record.accessed,
            )

            add(
                InlineSuggestion(
                    category="Website/webpage",
                    detected=(
                        record.title
                        or record.url
                    ),
                    suggestion=suggestion,
                    reason=(
                        "Native Zotero item type is webpage; "
                        "Style Guide treatment is inline-only."
                    ),
                    missingmetadata=missing,
                    paragraph=record.paragraph,
                    section=record.section,
                    context=record.context,
                    source="Zotero field",
                    relatedreference=relatedbibliography(
                        record.title or record.url,
                        bibliography,
                    ),
                )
            )

    bibliographyparagraphs = {
        item.paragraph
        for item in bibliography
    }

    for paragraph in paragraphs:
        if paragraph.number in bibliographyparagraphs:
            continue

        text = paragraph.text
        lower = text.casefold()

        for parenthetical in PARENTHETICAL_PATTERN.findall(
            text
        ):
            parentheticallower = parenthetical.casefold()

            if (
                "personal communication"
                in parentheticallower
                or "pers. comm"
                in parentheticallower
            ):
                hasdate = bool(
                    re.search(
                        r"\b\d{1,2}\s+[A-Za-z]+\s+"
                        r"(?:18|19|20)\d{2}\b",
                        parenthetical,
                    )
                )

                add(
                    InlineSuggestion(
                        category=(
                            "Personal communication"
                        ),
                        detected=parenthetical,
                        suggestion=(
                            parenthetical
                            if hasdate
                            else (
                                "(personal communication, "
                                "[person name], "
                                "[DD Month YYYY])"
                            )
                        ),
                        reason=(
                            "Explicit personal communication "
                            "phrase detected."
                        ),
                        missingmetadata=(
                            "Ready for reviewer confirmation"
                            if hasdate
                            else "Enter person name and date"
                        ),
                        paragraph=paragraph.number,
                        section=paragraph.section,
                        context=text,
                    )
                )

            elif "data on file" in parentheticallower:
                add(
                    InlineSuggestion(
                        category=(
                            "Sponsor data on file"
                        ),
                        detected=parenthetical,
                        suggestion=(
                            "(sponsor data on file)"
                        ),
                        reason=(
                            "Explicit data-on-file phrase "
                            "detected."
                        ),
                        missingmetadata=(
                            "Confirm whether sponsor should "
                            "be named differently"
                        ),
                        paragraph=paragraph.number,
                        section=paragraph.section,
                        context=text,
                    )
                )

            elif (
                "in preparation"
                in parentheticallower
                or "in prep"
                in parentheticallower
            ):
                add(
                    InlineSuggestion(
                        category=(
                            "Material in preparation"
                        ),
                        detected=parenthetical,
                        suggestion=parenthetical,
                        reason=(
                            "Explicit in-preparation phrase "
                            "detected."
                        ),
                        missingmetadata=(
                            "Confirm author, manuscript title, "
                            "and date; in-press articles belong "
                            "in EndNote."
                        ),
                        paragraph=paragraph.number,
                        section=paragraph.section,
                        context=text,
                    )
                )

        for match in URL_PATTERN.finditer(text):
            url = match.group(0).rstrip(".,;:)")

            if (
                "doi.org" in url.casefold()
                or normalizeddoi(url)
            ):
                continue

            suggestion, missing = websitecitation(url)

            add(
                InlineSuggestion(
                    category="Website/webpage",
                    detected=url,
                    suggestion=suggestion,
                    reason=(
                        "Non-DOI URL detected; Style Guide "
                        "treatment is inline-only."
                    ),
                    missingmetadata=missing,
                    paragraph=paragraph.number,
                    section=paragraph.section,
                    context=text,
                    relatedreference=relatedbibliography(
                        url,
                        bibliography,
                    ),
                )
            )

        for token, (
            category,
            displayname,
        ) in SOURCE_ORGANIZATIONS.items():
            if re.search(
                rf"\b{re.escape(token)}\b",
                lower,
            ):
                if category == "Website/webpage":
                    suggestion = (
                        f"({displayname}, "
                        f"[specific webpage title], "
                        f"Accessed [DD Month YYYY])"
                    )

                    missing = (
                        "Confirm exact webpage title "
                        "and access date"
                    )

                else:
                    suggestion, missing = governmentcitation(
                        displayname,
                        text,
                    )

                add(
                    InlineSuggestion(
                        category=category,
                        detected=displayname,
                        suggestion=suggestion,
                        reason=(
                            "Recognized website, government, "
                            "or regulatory organization mentioned "
                            "in document text."
                        ),
                        missingmetadata=missing,
                        paragraph=paragraph.number,
                        section=paragraph.section,
                        context=text,
                        relatedreference=relatedbibliography(
                            displayname,
                            bibliography,
                        ),
                    )
                )

    return suggestions


def requestaction(record: Record) -> str:
    if record.title and record.year:
        return (
            "Request reference from library owner"
        )

    return "Request source identification"


def checkreferences(
    documentpath: str | Path,
    librarypath: str | Path,
    cancel: Optional[
        Callable[[], bool]
    ] = None,
    progress: Optional[
        Callable[[int, int, str], None]
    ] = None,
) -> CheckResult:
    emit = progress or (
        lambda current, total, message: None
    )

    shouldcancel = cancel or (
        lambda: False
    )

    emit(
        0,
        0,
        "Reading Word document and reference blocks…",
    )

    (
        native,
        bibliography,
        plain,
        existingendnote,
        paragraphs,
        warnings,
        fieldcount,
        referenceblockcount,
    ) = readword(documentpath)

    emit(
        0,
        0,
        "Reading shared-library export…",
    )

    librarytext = Path(
        librarypath
    ).read_text(
        encoding="utf-8-sig"
    )

    library = dedupe(
        parsereferencetext(
            librarytext,
            "library",
        )
    )

    inlineonly = detectinlinesources(
        paragraphs,
        native,
        bibliography,
    )

    literaturenative = [
        item
        for item in native
        if not nativeinline(item)
    ]

    nativeauthoryears = {
        (
            item.firstauthor,
            item.year,
        )
        for item in literaturenative
        if item.firstauthor and item.year
    }

    plainforlibrary = [
        item
        for item in plain
        if (
            item.firstauthor,
            item.year,
        )
        not in nativeauthoryears
    ]

    documents = dedupe(
        literaturenative
        + bibliography
        + plainforlibrary
    )

    result = CheckResult(
        existingendnote=existingendnote,
        inlineonly=inlineonly,
        warnings=warnings,
        documentrecords=len(documents),
        libraryrecords=len(library),
        nativecitations=fieldcount,
        referenceblocks=referenceblockcount,
    )

    totalwork = (
        len(library)
        + len(documents)
        + len(plain)
    )

    started = time.monotonic()

    try:
        libraryindex = buildindex(
            library,
            shouldcancel,
            emit,
            totalwork,
        )

    except InterruptedError:
        result.cancelled = True
        return result

    emit(
        len(library),
        totalwork,
        (
            f"Matching 0/{len(documents)} "
            f"document records against indexed library…"
        ),
    )

    for number, document in enumerate(
        documents,
        start=1,
    ):
        if shouldcancel():
            result.cancelled = True
            return result

        ranked = indexedcandidates(
            document,
            libraryindex,
        )

        best = ranked[0] if ranked else None

        if (
            best
            and best.score == 1.0
            and document.kind != "plaintextcitation"
        ):
            result.matched.append(
                Finding(
                    category="Matched",
                    documentrecord=document,
                    action="Insert via EndNote",
                    bestmatch=best.record,
                    score=best.score,
                    reason=best.reason,
                    candidates=ranked,
                )
            )

        elif best and best.score >= 0.85:
            result.review.append(
                Finding(
                    category="Needs review",
                    documentrecord=document,
                    action=(
                        "Verify candidate then "
                        "insert via EndNote"
                    ),
                    bestmatch=best.record,
                    score=best.score,
                    reason=best.reason,
                    candidates=ranked,
                )
            )

        else:
            result.missing.append(
                Finding(
                    category="Missing from library",
                    documentrecord=document,
                    action=requestaction(document),
                    bestmatch=(
                        best.record
                        if best
                        else None
                    ),
                    score=(
                        best.score
                        if best
                        else 0.0
                    ),
                    reason=(
                        "No indexed candidate reached "
                        "the 85% review threshold"
                    ),
                    candidates=ranked,
                )
            )

        elapsed = time.monotonic() - started

        eta = (
            (elapsed / number)
            * (len(documents) - number)
            if number
            else 0
        )

        emit(
            len(library) + number,
            totalwork,
            (
                f"Matching {number}/{len(documents)} "
                f"document records "
                f"— ETA {formateta(eta)}"
            ),
        )

    bibliographyindex = buildindex(
        bibliography,
        shouldcancel,
        lambda current, total, message: None,
        0,
    )

    emit(
        len(library) + len(documents),
        totalwork,
        (
            "Checking formatted citations against "
            "the relevant document bibliography block…"
        ),
    )

    for number, citation in enumerate(
        plain,
        start=1,
    ):
        if shouldcancel():
            result.cancelled = True
            return result

        ranked = indexedcandidates(
            citation,
            bibliographyindex,
        )

        best = ranked[0] if ranked else None

        if not best or best.score < 0.85:
            result.potentialmissing.append(
                Finding(
                    category=(
                        "Potential missing reference"
                    ),
                    documentrecord=citation,
                    action=(
                        "Request source identification"
                    ),
                    bestmatch=(
                        best.record
                        if best
                        else None
                    ),
                    score=(
                        best.score
                        if best
                        else 0.0
                    ),
                    reason=(
                        "Formatted author-year citation "
                        "has no confident match in the "
                        "document bibliography."
                    ),
                    candidates=ranked,
                )
            )

        emit(
            (
                len(library)
                + len(documents)
                + number
            ),
            totalwork,
            (
                f"Checking formatted citations: "
                f"{number}/{len(plain)}"
            ),
        )

    emit(
        totalwork,
        totalwork,
        "Reference Library Check complete.",
    )

    return result


def buildrequestemail(
    result: CheckResult,
    documentname: str,
) -> str:
    requests = (
        result.missing
        + result.potentialmissing
    )

    lines = [
        f"Subject: Reference requests — {documentname}",
        "",
        "Hello,",
        "",
        (
            "Could you please help locate, confirm, or add "
            "the following references to the shared EndNote library?"
        ),
        "",
        f"Document checked: {documentname}",
    ]

    if not requests:
        lines.extend(
            [
                "",
                (
                    "No unresolved literature-reference "
                    "requests were identified."
                ),
            ]
        )

        return "\n".join(lines)

    fullreferences = [
        item
        for item in requests
        if item.action
        == "Request reference from library owner"
    ]

    identificationneeded = [
        item
        for item in requests
        if item.action
        != "Request reference from library owner"
    ]

    def appenditems(
        heading: str,
        items: list[Finding],
    ) -> None:
        if not items:
            return

        lines.extend(
            [
                "",
                heading,
            ]
        )

        for index, item in enumerate(
            items,
            start=1,
        ):
            record = item.documentrecord

            lines.extend(
                [
                    "",
                    (
                        f"{index}. Citation/reference found: "
                        f"{record.raw}"
                    ),
                    (
                        f"   Section: "
                        f"{record.section or 'Not identified'}"
                    ),
                    (
                        f"   Paragraph: "
                        f"{record.paragraph or 'Not identified'}"
                    ),
                    (
                        f"   Context: "
                        f"{record.context or 'Not available'}"
                    ),
                    (
                        "   Available metadata: "
                        f"{record.endnote_search().replace(chr(10), '; ')}"
                    ),
                    (
                        f"   Requested action: "
                        f"{item.action}"
                    ),
                ]
            )

            if (
                item.bestmatch
                and item.score >= 0.50
            ):
                lines.append(
                    (
                        "   Possible alternate record "
                        f"({item.score:.0%}): "
                        f"{item.bestmatch.raw}"
                    )
                )

    appenditems(
        (
            "A. Full references absent from "
            "the shared library"
        ),
        fullreferences,
    )

    appenditems(
        (
            "B. Citations needing bibliographic "
            "identification"
        ),
        identificationneeded,
    )

    lines.extend(
        [
            "",
            "Thank you,",
            "[Your Name]",
        ]
    )

    return "\n".join(lines)


def exportrows(
    result: CheckResult,
) -> list[dict]:
    rows: list[dict] = []

    for collection in (
        result.matched,
        result.review,
        result.missing,
        result.potentialmissing,
        result.inlineonly,
    ):
        rows.extend(
            item.row()
            for item in collection
        )

    for record in result.existingendnote:
        rows.append(
            ParseWarning(
                message=(
                    "Existing EndNote citation field detected "
                    "— excluded from matching."
                ),
                paragraph=record.paragraph,
                section=record.section,
                context=record.context,
            ).row()
        )

    rows.extend(
        warning.row()
        for warning in result.warnings
    )

    return rows


def writereport(
    result: CheckResult,
    path: str | Path,
) -> None:
    fields = [
        "category",
        "action",
        "document",
        "bestmatch",
        "score",
        "reason",
        "suggestedcitation",
        "endnotesearch",
        "section",
        "paragraph",
        "referenceblock",
        "referencenumber",
        "context",
        "missingmetadata",
        "rawdocument",
        "rawmatch",
    ]

    with Path(path).open(
        "w",
        encoding="utf-8-sig",
        newline="",
    ) as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=fields,
        )

        writer.writeheader()
        writer.writerows(
            exportrows(result)
        )
