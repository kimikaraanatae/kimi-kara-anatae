#!/usr/bin/env python3
"""Create a GitHub archive snapshot of the exact article published on note.

The source is the managed Markdown article. The generated snapshot changes its
front matter to ``publication_status: published`` and records the actual note
URL and timestamp. ARTICLE_INDEX is updated by header name: existing values in
columns such as ``訂正`` and ``関連研究`` are preserved.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import re
from pathlib import Path

FRONTMATTER = re.compile(r"\A---\s*\n(.*?)\n---\s*\n", re.S)


def parse_source(text: str) -> tuple[dict, str]:
    try:
        import yaml
    except Exception as exc:
        raise SystemExit(f"PyYAML is required: {exc}")
    match = FRONTMATTER.match(text)
    if not match:
        return {}, text
    metadata = yaml.safe_load(match.group(1)) or {}
    if not isinstance(metadata, dict):
        raise SystemExit("front matter must be a mapping")
    return metadata, text[match.end():]


def yaml_text(data: dict) -> str:
    import yaml
    return yaml.safe_dump(data, allow_unicode=True, sort_keys=False).strip()


def parse_markdown_row(line: str) -> list[str]:
    stripped = line.strip()
    if not (stripped.startswith("|") and stripped.endswith("|")):
        raise ValueError("not a Markdown table row")
    return [cell.strip() for cell in stripped[1:-1].split("|")]


def format_markdown_row(cells: list[str]) -> str:
    return "| " + " | ".join(cells) + " |"


def update_article_index(
    index_path: Path,
    article_id: str,
    title: str,
    kind: str,
    primary_magazine: str,
    subtype: str,
    public_state: str,
    published_date: str,
    version: str,
    note_url: str,
    snapshot_path: str,
) -> None:
    if not index_path.is_file():
        return
    lines = index_path.read_text(encoding="utf-8").splitlines()
    header_index = next((i for i, line in enumerate(lines) if line.startswith("| ID |")), None)
    if header_index is None:
        raise SystemExit("ARTICLE_INDEX.md table header not found")
    headers = parse_markdown_row(lines[header_index])
    required = {"ID", "記事", "主記事種別", "主マガジン", "副種別", "公開状態", "公開日", "更新日", "版", "note", "保存版"}
    missing = required - set(headers)
    if missing:
        raise SystemExit("ARTICLE_INDEX.md missing columns: " + ", ".join(sorted(missing)))

    found = False
    for idx in range(header_index + 2, len(lines)):
        line = lines[idx]
        if not line.startswith("|"):
            continue
        cells = parse_markdown_row(line)
        if len(cells) != len(headers):
            continue
        record = dict(zip(headers, cells))
        if record.get("ID") != article_id:
            continue
        # Only publication-derived columns are updated. Human-managed columns,
        # including 訂正 and 関連研究, remain unchanged.
        record.update({
            "記事": title,
            "主記事種別": kind,
            "主マガジン": primary_magazine,
            "副種別": subtype,
            "公開状態": public_state,
            "公開日": published_date,
            "更新日": published_date,
            "版": version,
            "note": note_url,
            "保存版": snapshot_path,
        })
        lines[idx] = format_markdown_row([record.get(header, "") for header in headers])
        found = True
        break
    if not found:
        raise SystemExit(f"ARTICLE_INDEX.md row not found: {article_id}")
    index_path.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True)
    parser.add_argument("--repo", default=".")
    parser.add_argument("--article-id", required=True)
    parser.add_argument("--version", required=True)
    parser.add_argument("--note-url", required=True)
    parser.add_argument("--published-at", required=True, help="ISO 8601, e.g. 2026-08-01T09:00:00+09:00")
    args = parser.parse_args()

    source = Path(args.source).resolve()
    repo = Path(args.repo).resolve()
    if not source.is_file():
        raise SystemExit(f"not found: {source}")
    try:
        published = dt.datetime.fromisoformat(args.published_at)
    except ValueError as exc:
        raise SystemExit(f"invalid --published-at: {exc}")
    if published.tzinfo is None:
        raise SystemExit("--published-at must include a timezone offset")

    source_bytes = source.read_bytes()
    metadata, body = parse_source(source_bytes.decode("utf-8"))
    article_id = args.article_id
    if metadata.get("article_id") and str(metadata["article_id"]) != article_id:
        raise SystemExit("source article_id does not match --article-id")
    metadata.update({
        "article_id": article_id,
        "version": args.version,
        "publication_status": "published",
        "note_url": args.note_url,
        "published_at": published.isoformat(),
        "archived_at": dt.datetime.now(dt.timezone.utc).isoformat(),
    })

    target_dir = repo / "articles" / article_id
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / f"{args.version}.md"
    target.write_text(f"---\n{yaml_text(metadata)}\n---\n\n{body.lstrip()}", encoding="utf-8", newline="\n")
    meta = {
        "article_id": article_id,
        "version": args.version,
        "note_url": args.note_url,
        "published_at": published.isoformat(),
        "snapshot": target.name,
        "source_sha256": hashlib.sha256(source_bytes).hexdigest(),
        "article_type": metadata.get("article_type"),
        "article_type_label": metadata.get("article_type_label"),
        "primary_magazine": metadata.get("primary_magazine"),
        "subtype": metadata.get("subtype"),
        "public_state": metadata.get("public_state"),
    }
    (target_dir / "metadata.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    title_match = re.search(r"^#\s+(.+)$", body, re.M)
    title = title_match.group(1).strip() if title_match else article_id
    update_article_index(
        repo / "ARTICLE_INDEX.md", article_id, title, str(metadata.get("article_type_label") or metadata.get("status") or "Article"),
        str(metadata.get("primary_magazine") or "—"), str(metadata.get("subtype") or "—"),
        str(metadata.get("public_state") or "公開ベータ"),
        published.date().isoformat(), args.version, args.note_url,
        f"articles/{article_id}/{args.version}.md",
    )
    print(target)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
