#!/usr/bin/env python3
"""Fail-closed pre-publication audit.

Audits the real Git work tree plus exact publication-staging text,
public images, and GitHub Release candidates. Repository metadata (the root
``.git`` directory created by ``git init`` or ``actions/checkout``) is pruned
from filesystem traversal; a ``.git`` path inside an uploaded ZIP remains a
release-blocking error.

The audit checks text and metadata *and* extractable body content in DOCX,
XLSX, PPTX, and PDF files. It is a release gate, not an anonymity guarantee.
Reports must be written outside every audited scope, and absolute audited-root
paths are never stored in reports.
"""
from __future__ import annotations

import argparse
import io
import json
import os
import re
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

TEXT_EXTS = {
    ".md", ".txt", ".csv", ".tsv", ".yml", ".yaml", ".json", ".toml",
    ".ini", ".py", ".ps1", ".html", ".css", ".js", ".xml", ".rels",
}
OFFICE_EXTS = {".docx", ".xlsx", ".pptx"}
IMAGE_EXTS = {".jpg", ".jpeg", ".tif", ".tiff", ".webp", ".png"}
PDF_EXTS = {".pdf"}
ARCHIVE_EXTS = {".zip"}
# .git is silently pruned only for the audited repository work tree. Any .git
# member in a release ZIP is forbidden by scan_archive().
FORBIDDEN_DIRS = {"_work", "_private", "_drafts", "review", "logs", ".git"}
FORBIDDEN_PATH_PATTERNS = [
    re.compile(r"(?i)\b[A-Z]:[\\/]+Users[\\/]+[^\\/\s<>\"']+"),
    re.compile(r"(?i)/Users/[^/\s<>\"']+"),
    re.compile(r"(?i)/home/[^/\s<>\"']+"),
    re.compile(r"(?i)/mnt/data/[^\s<>\"']*"),
    re.compile(r"(?i)AppData[\\/]"),
    re.compile(r"(?i)OneDrive[\\/]"),
]
EMAIL_RE = re.compile(r"(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b")
PLACEHOLDER_RE = re.compile(r"\{\{[A-Z0-9_]+\}\}")
ALLOWED_EMAIL_DOMAINS = {"users.noreply.github.com"}
OFFICE_META_FIELDS = {"creator", "lastModifiedBy", "company", "manager"}
SENSITIVE_PDF_FIELDS = {"/Author", "/Subject", "/Keywords"}
SAFE_IMAGE_INFO_KEYS = {
    "dpi", "icc_profile", "srgb", "gamma", "chromaticity", "transparency",
    "aspect", "duration", "loop", "background",
}
MAX_ZIP_ENTRIES = 5000
MAX_ZIP_UNCOMPRESSED = 300 * 1024 * 1024
MAX_ARCHIVE_DEPTH = 3


@dataclass(frozen=True)
class Scope:
    label: str
    root: Path


class Audit:
    def __init__(self, allow_placeholders: bool, allowed_emails: set[str]) -> None:
        self.allow_placeholders = allow_placeholders
        self.allowed_emails = {e.lower() for e in allowed_emails}
        self.errors: list[dict[str, str]] = []
        self.warnings: list[dict[str, str]] = []
        self.files = 0

    def error(self, file: str, issue: str) -> None:
        self.errors.append({"file": file, "issue": issue})

    def warn(self, file: str, issue: str) -> None:
        self.warnings.append({"file": file, "issue": issue})

    def scan_text(self, display: str, data: bytes, *, self_source: bool = False) -> None:
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError:
            self.error(display, "text content is not UTF-8")
            return
        if self_source:
            return
        if not self.allow_placeholders:
            for placeholder in sorted(set(PLACEHOLDER_RE.findall(text))):
                self.error(display, f"unresolved placeholder {placeholder}")
        for pattern in FORBIDDEN_PATH_PATTERNS:
            for match in pattern.findall(text):
                self.error(display, f"local path exposed: {match}")
        for email in sorted(set(EMAIL_RE.findall(text))):
            lower = email.lower()
            domain = lower.rsplit("@", 1)[-1]
            if lower in self.allowed_emails or domain in ALLOWED_EMAIL_DOMAINS:
                continue
            if email == "{{PUBLIC_CONTACT_EMAIL}}":
                continue
            self.warn(display, f"email address present; confirm it is the public contact: {email}")

    def _office_content_member(self, suffix: str, name: str) -> bool:
        posix = PurePosixPath(name).as_posix()
        if suffix == ".docx":
            return posix.startswith("word/") and posix.endswith((".xml", ".rels"))
        if suffix == ".xlsx":
            return posix.startswith("xl/") and posix.endswith((".xml", ".rels"))
        if suffix == ".pptx":
            return posix.startswith("ppt/") and posix.endswith((".xml", ".rels"))
        return False

    def scan_office(self, display: str, suffix: str, data: bytes) -> None:
        """Scan Office metadata, body XML, comments, relationships and media."""
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                names = set(archive.namelist())
                for name in ("docProps/core.xml", "docProps/app.xml"):
                    if name not in names:
                        continue
                    root = ET.fromstring(archive.read(name))
                    for elem in root.iter():
                        field = elem.tag.split("}")[-1]
                        value = (elem.text or "").strip()
                        if field in OFFICE_META_FIELDS and value:
                            self.error(display, f"Office metadata {field}={value}")
                for info in archive.infolist():
                    if info.is_dir():
                        continue
                    name = PurePosixPath(info.filename).as_posix()
                    try:
                        member = archive.read(info)
                    except Exception as exc:
                        self.error(f"{display}!{name}", f"cannot read Office member: {exc}")
                        continue
                    if self._office_content_member(suffix, name):
                        # Includes DOCX main text/headers/footers/comments/links,
                        # XLSX cell/shared-string/comments/external-link XML, and
                        # PPTX slide/notes/comments/relationship XML.
                        self.scan_text(f"{display}!{name}", member)
                    elif "/media/" in f"/{name}" and Path(name).suffix.lower() in IMAGE_EXTS:
                        self.scan_image(f"{display}!{name}", member)
        except Exception as exc:
            self.error(display, f"Office content/metadata inspection failed: {exc}")

    def scan_image(self, display: str, data: bytes) -> None:
        try:
            from PIL import Image
        except Exception as exc:
            self.error(display, f"image inspection backend unavailable: {exc}")
            return
        try:
            with Image.open(io.BytesIO(data)) as image:
                exif = image.getexif()
                if exif:
                    self.error(display, f"EXIF metadata present ({len(exif)} tags)")
                info = dict(getattr(image, "info", {}) or {})
                if info.get("exif"):
                    self.error(display, "embedded EXIF blob present")
                unsafe = []
                for key, value in info.items():
                    if key in SAFE_IMAGE_INFO_KEYS or value in (None, b"", ""):
                        continue
                    unsafe.append(key)
                if unsafe:
                    self.error(display, "image text/XMP metadata present: " + ", ".join(sorted(unsafe)))
        except Exception as exc:
            self.error(display, f"image inspection failed: {exc}")

    def scan_pdf(self, display: str, data: bytes) -> None:
        try:
            from pypdf import PdfReader
        except Exception as exc:
            self.error(display, f"PDF inspection backend unavailable: {exc}")
            return
        try:
            reader = PdfReader(io.BytesIO(data))
            metadata = reader.metadata or {}
            for field in SENSITIVE_PDF_FIELDS:
                value = metadata.get(field)
                if value:
                    self.error(display, f"PDF metadata {field}={value}")
            for field in ("/Creator", "/Producer", "/Title"):
                value = metadata.get(field)
                if not value:
                    continue
                value_text = str(value)
                for pattern in FORBIDDEN_PATH_PATTERNS:
                    if pattern.search(value_text):
                        self.error(display, f"PDF metadata {field} contains local path: {value_text}")
                if EMAIL_RE.search(value_text):
                    self.error(display, f"PDF metadata {field} contains email: {value_text}")
                if field in {"/Creator", "/Producer"}:
                    self.warn(display, f"PDF metadata {field}={value_text}; confirm it is generic software metadata")
            for index, page in enumerate(reader.pages, start=1):
                try:
                    text = page.extract_text() or ""
                except Exception as exc:
                    self.error(f"{display}!page-{index}", f"PDF text extraction failed: {exc}")
                    continue
                if text:
                    self.scan_text(f"{display}!page-{index}", text.encode("utf-8"))
                # Also inspect link/action URI strings when available.
                try:
                    annotations = page.get("/Annots") or []
                    for annotation_ref in annotations:
                        annotation = annotation_ref.get_object()
                        action = annotation.get("/A")
                        if action and action.get("/URI"):
                            uri = str(action.get("/URI"))
                            self.scan_text(f"{display}!page-{index}-uri", uri.encode("utf-8"))
                except Exception:
                    # Annotation shapes vary; body/metadata inspection remains authoritative.
                    pass
        except Exception as exc:
            self.error(display, f"PDF inspection failed: {exc}")

    def scan_issue_form(self, display: str, data: bytes) -> None:
        try:
            import yaml
        except Exception as exc:
            self.error(display, f"YAML inspection backend unavailable: {exc}")
            return
        try:
            parsed = yaml.safe_load(data.decode("utf-8"))
        except Exception as exc:
            self.error(display, f"invalid YAML: {exc}")
            return
        if not isinstance(parsed, dict):
            self.error(display, "Issue Form must be a YAML mapping")
            return
        if "about" in parsed:
            self.error(display, "Issue Form uses Markdown-template key 'about'; use 'description'")
        for required in ("name", "description", "body"):
            if not parsed.get(required):
                self.error(display, f"Issue Form missing required top-level key: {required}")
        if "labels" in parsed and parsed["labels"]:
            self.warn(display, "Issue Form declares labels; create them in the repository before launch")

    def scan_archive(self, display: str, data: bytes, depth: int) -> None:
        if depth >= MAX_ARCHIVE_DEPTH:
            self.error(display, f"archive nesting exceeds depth {MAX_ARCHIVE_DEPTH}")
            return
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                infos = archive.infolist()
                if len(infos) > MAX_ZIP_ENTRIES:
                    self.error(display, f"ZIP contains too many entries: {len(infos)}")
                    return
                total = sum(info.file_size for info in infos)
                if total > MAX_ZIP_UNCOMPRESSED:
                    self.error(display, f"ZIP uncompressed size exceeds limit: {total}")
                    return
                for info in infos:
                    name = info.filename
                    posix = PurePosixPath(name)
                    member_path = posix.as_posix()
                    if name.startswith(("/", "\\")) or ".." in posix.parts:
                        self.error(display, f"unsafe ZIP member path: {name}")
                        continue
                    if any(part in FORBIDDEN_DIRS for part in posix.parts):
                        self.error(f"{display}!{member_path}", "forbidden work/private/.git path inside release ZIP")
                        continue
                    if info.is_dir():
                        continue
                    try:
                        member = archive.read(info)
                    except Exception as exc:
                        self.error(f"{display}!{name}", f"cannot read ZIP member: {exc}")
                        continue
                    member_self_source = (
                        member_path.endswith("tools/audit_public_release.py")
                        or member_path.endswith("tools/test_publication_audit_negative.py")
                        or "/audit_reports/NEGATIVE_TEST_REPORT." in member_path
                        or "/audit_reports/PUBLICATION_GATE_TESTS_" in member_path
                    )
                    self.scan_bytes(
                        f"{display}!{name}", Path(name).suffix.lower(), member,
                        depth=depth + 1, self_source=member_self_source,
                    )
        except Exception as exc:
            self.error(display, f"ZIP inspection failed: {exc}")

    def scan_bytes(self, display: str, suffix: str, data: bytes, *, depth: int = 0, self_source: bool = False) -> None:
        if suffix in TEXT_EXTS:
            self.scan_text(display, data, self_source=self_source)
        elif suffix in OFFICE_EXTS:
            self.scan_office(display, suffix, data)
        elif suffix in IMAGE_EXTS:
            self.scan_image(display, data)
        elif suffix in PDF_EXTS:
            self.scan_pdf(display, data)
        elif suffix in ARCHIVE_EXTS:
            self.scan_archive(display, data, depth)

    def scan_scope(self, scope: Scope) -> None:
        if not scope.root.exists():
            self.error(scope.label, "audit scope does not exist")
            return
        # os.walk lets us prune .git before descending. Other forbidden work
        # directories are reported once and pruned.
        for dirpath, dirnames, filenames in os.walk(scope.root):
            current = Path(dirpath)
            rel_dir = current.relative_to(scope.root)
            retained: list[str] = []
            for dirname in sorted(dirnames):
                rel = rel_dir / dirname
                if dirname == ".git":
                    continue
                if dirname in FORBIDDEN_DIRS:
                    self.error(f"{scope.label}/{rel.as_posix()}", "forbidden work/private directory")
                    continue
                retained.append(dirname)
            dirnames[:] = retained
            for filename in sorted(filenames):
                path = current / filename
                rel = path.relative_to(scope.root)
                if filename == ".git":
                    # Worktree/submodule metadata file; repository mechanics,
                    # not publication content.
                    continue
                display = f"{scope.label}/{rel.as_posix()}"
                self.files += 1
                size = path.stat().st_size
                if size > 100 * 1024 * 1024:
                    self.error(display, "file exceeds GitHub 100 MiB limit")
                elif size > 25 * 1024 * 1024:
                    self.warn(display, "large binary; prefer a GitHub Release attachment")
                try:
                    data = path.read_bytes()
                except Exception as exc:
                    self.error(display, f"cannot read file: {exc}")
                    continue
                rel_posix = rel.as_posix()
                self_source = (
                    rel_posix == "tools/audit_public_release.py"
                    or rel_posix.endswith("tools/test_publication_audit_negative.py")
                    or rel_posix.startswith("audit_reports/NEGATIVE_TEST_REPORT.")
                )
                if rel.parts[:2] == (".github", "ISSUE_TEMPLATE") and path.suffix.lower() in {".yml", ".yaml"} and rel.name != "config.yml":
                    self.scan_issue_form(display, data)
                self.scan_bytes(display, path.suffix.lower(), data, self_source=self_source)


def parse_extra_root(value: str) -> Scope:
    if "=" not in value:
        raise argparse.ArgumentTypeError("--extra-root must be LABEL=PATH")
    label, path = value.split("=", 1)
    label = label.strip()
    if not label:
        raise argparse.ArgumentTypeError("extra-root label is empty")
    return Scope(label, Path(path).expanduser().resolve())


def path_is_within(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", default=".", help="GitHub release repository root")
    parser.add_argument("--extra-root", action="append", default=[], type=parse_extra_root, help="LABEL=PATH, repeatable")
    parser.add_argument("--allow-placeholders", action="store_true")
    parser.add_argument("--allow-email", action="append", default=[])
    parser.add_argument("--report", help="Report path. Must be outside every audited scope.")
    args = parser.parse_args()

    scopes = [Scope("repo", Path(args.root).expanduser().resolve()), *args.extra_root]
    if args.report:
        report = Path(args.report).expanduser().resolve()
    else:
        report = Path(tempfile.gettempdir()) / "kimikara_publication_audit.json"
    if any(path_is_within(report, scope.root) for scope in scopes):
        print("ERROR: audit report must be written outside audited scopes", file=sys.stderr)
        return 2

    audit = Audit(args.allow_placeholders, set(args.allow_email))
    for scope in scopes:
        audit.scan_scope(scope)

    repo = scopes[0].root
    required = [
        "README.md", "ABOUT.md", "EDITORIAL_POLICY.md", "AI_USE_POLICY.md",
        "CORRECTIONS.md", "CONTACT.md", "SOURCE_POLICY.md", "PRIVACY_AND_SAFETY.md",
        "LICENSE_POLICY.md", "ARTICLE_INDEX.md", "RESEARCH_STATUS.md",
        "RESEARCH_REVIEW_STATUS.md",
    ]
    for name in required:
        if not (repo / name).is_file():
            audit.error(f"repo/{name}", "required file missing")

    result = {
        "status": "FAIL" if audit.errors else "PASS",
        "scopes": [scope.label for scope in scopes],
        "files": audit.files,
        "errors": audit.errors,
        "warnings": audit.warnings,
        "inspection_scope": {
            "text": "UTF-8 text and exact publication-staging text",
            "images": "EXIF and text/XMP metadata",
            "office": "metadata, extractable XML body/comments/relationships, embedded images",
            "pdf": "metadata, extractable page text and link URIs",
            "zip": "all members recursively, including Office/PDF body checks; forbidden .git/work paths fail",
        },
    }
    report.parent.mkdir(parents=True, exist_ok=True)
    report.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 1 if audit.errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
