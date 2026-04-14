#!/usr/bin/env python3
"""Apply stable Word track changes to a DOCX."""

from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
import zipfile
from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

try:
    from lxml import etree
except ImportError as exc:
    raise SystemExit("Missing dependency: lxml") from exc

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
XML_NS = "http://www.w3.org/XML/1998/namespace"


@dataclass
class InlineReplacementLog:
    old_text: str
    new_text: str
    paragraph_index: int
    replaced_count: int
    mode: str
    target: str
    cell_ref: str | None


@dataclass
class CommentLog:
    target: str
    author: str
    paragraph_index: int
    cell_ref: str | None


@dataclass
class HeaderFooterLog:
    old_text: str
    new_text: str
    replaced_count: int
    mode: str
    target: str
    ref: str


# ─── Header / Footer paragraph helpers ─────────────────────────────────────────

def _iter_header_footer_xmls(temp_path: Path) -> list[tuple[str, etree._Element]]:
    """Load all header/footer XML files. Returns list of (ref, root)."""
    word_dir = temp_path / "word"
    if not word_dir.is_dir():
        return []
    results: list[tuple[str, etree._Element]] = []
    parser = etree.XMLParser(remove_blank_text=False)
    for xml_file in sorted(word_dir.glob("header*.xml")):
        tree = etree.parse(str(xml_file), parser=parser)
        results.append((xml_file.stem, tree.getroot()))
    for xml_file in sorted(word_dir.glob("footer*.xml")):
        tree = etree.parse(str(xml_file), parser=parser)
        results.append((xml_file.stem, tree.getroot()))
    return results


def _all_hf_paragraphs(
    hf_list: list[tuple[str, etree._Element]],
) -> list[tuple[str, etree._Element]]:
    """Return (ref, paragraph) for all paragraphs in headers/footers."""
    results: list[tuple[str, etree._Element]] = []
    for ref, root in hf_list:
        for p in root.xpath(".//w:p", namespaces=NS):
            results.append((ref, p))
    return results


def _rewrite_hf_xmls(
    hf_list: list[tuple[str, etree._Element]],
    temp_path: Path,
) -> None:
    """Write modified header/footer XML trees back to disk."""
    for ref, root in hf_list:
        xml_file = temp_path / "word" / f"{ref}.xml"
        tree = etree.ElementTree(root)
        tree.write(str(xml_file), xml_declaration=True, encoding="UTF-8", standalone=True)


def all_document_paragraphs(document_root: etree._Element) -> list[tuple[int, etree._Element, str | None]]:
    """
    Yield all paragraphs in document order: body paragraphs first,
    then table cell paragraphs.  Yields (display_index, paragraph_element, cell_ref).
    cell_ref is None for body paragraphs, 'T{r}-R{r}-C{c}' for table cells.
    """
    results: list[tuple[int, etree._Element, str | None]] = []

    # Body paragraphs (displayed as positive integers)
    for idx, p in enumerate(document_root.xpath("/w:document/w:body/w:p", namespaces=NS), start=1):
        results.append((idx, p, None))

    # Table cell paragraphs (displayed as table cell reference)
    tables = document_root.xpath("/w:document/w:body/w:tbl", namespaces=NS)
    for t_idx, tbl in enumerate(tables, start=1):
        rows = tbl.xpath("./w:tr", namespaces=NS)
        for r_idx, row in enumerate(rows, start=1):
            cells = row.xpath("./w:tc", namespaces=NS)
            for c_idx, cell in enumerate(cells, start=1):
                cell_ref = f"T{t_idx}-R{r_idx}-C{c_idx}"
                for _p in cell.xpath("./w:p", namespaces=NS):
                    results.append((0, _p, cell_ref))

    return results


def w_tag(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def preserve_space(node: etree._Element, text: str) -> None:
    if text.startswith(" ") or text.endswith(" "):
        node.set(f"{{{XML_NS}}}space", "preserve")


def paragraph_text(paragraph: etree._Element) -> str:
    parts: list[str] = []
    for node in paragraph.xpath("./w:r/w:t", namespaces=NS):
        parts.append(node.text or "")
    return "".join(parts)


def run_text(run: etree._Element) -> str:
    return "".join(node.text or "" for node in run.xpath("./w:t", namespaces=NS))


def clone_run_properties(run: etree._Element | None) -> etree._Element | None:
    if run is None:
        return None
    rpr = run.find("w:rPr", namespaces=NS)
    return deepcopy(rpr) if rpr is not None else None


def first_run_properties(paragraph: etree._Element) -> etree._Element | None:
    return clone_run_properties(paragraph.find("w:r", namespaces=NS))


def build_run(text: str, rpr: etree._Element | None, text_tag: str) -> etree._Element:
    run = etree.Element(w_tag("r"))
    if rpr is not None:
        run.append(deepcopy(rpr))
    text_node = etree.SubElement(run, w_tag(text_tag))
    preserve_space(text_node, text)
    text_node.text = text
    return run


def append_plain_text(
    paragraph: etree._Element,
    text: str,
    rpr: etree._Element | None,
) -> None:
    if text:
        paragraph.append(build_run(text, rpr, "t"))


def revision_wrapper(
    kind: str,
    revision_id: int,
    author: str,
    timestamp: str,
) -> etree._Element:
    node = etree.Element(w_tag(kind))
    node.set(w_tag("id"), str(revision_id))
    node.set(w_tag("author"), author)
    node.set(w_tag("date"), timestamp)
    return node


def set_paragraph_rsid(paragraph: etree._Element, rsid: str | None) -> None:
    if rsid:
        paragraph.set(w_tag("rsidR"), rsid)
        paragraph.set(w_tag("rsidRDefault"), rsid)


def reset_paragraph_contents(paragraph: etree._Element) -> None:
    ppr = paragraph.find("w:pPr", namespaces=NS)
    for child in list(paragraph):
        paragraph.remove(child)
    if ppr is not None:
        paragraph.append(ppr)


def current_timestamp() -> str:
    return (
        datetime.now(timezone.utc)
        .replace(microsecond=0)
        .isoformat()
        .replace("+00:00", "Z")
    )


def replace_paragraph(
    paragraph: etree._Element,
    old_text: str,
    new_text: str,
    author: str,
    revision_id: int,
    rsid: str | None,
) -> bool:
    if paragraph_text(paragraph) != old_text:
        return False

    rpr = first_run_properties(paragraph)
    timestamp = current_timestamp()
    reset_paragraph_contents(paragraph)

    deletion = revision_wrapper("del", revision_id, author, timestamp)
    deletion.append(build_run(old_text, rpr, "delText"))
    paragraph.append(deletion)

    insertion = revision_wrapper("ins", revision_id + 1, author, timestamp)
    insertion.append(build_run(new_text, rpr, "t"))
    paragraph.append(insertion)
    set_paragraph_rsid(paragraph, rsid)
    return True


def replace_text_in_single_run(
    paragraph: etree._Element,
    run: etree._Element,
    old_text: str,
    new_text: str,
    author: str,
    revision_id: int,
    rsid: str | None,
    occurrence: int | None = None,
) -> int:
    original_text = run_text(run)
    if not old_text or old_text not in original_text:
        return 0

    rpr = clone_run_properties(run)
    timestamp = current_timestamp()
    parent = run.getparent()
    if parent is None:
        return 0

    insert_at = parent.index(run)
    cursor = 0
    current_revision_id = revision_id
    replaced = 0
    new_nodes: list[etree._Element] = []

    match_index = 0
    while True:
        start = original_text.find(old_text, cursor)
        if start == -1:
            break
        match_index += 1

        prefix = original_text[cursor:start]
        if prefix:
            new_nodes.append(build_run(prefix, rpr, "t"))

        should_replace = occurrence is None or match_index == occurrence
        if should_replace:
            deletion = revision_wrapper("del", current_revision_id, author, timestamp)
            deletion.append(build_run(old_text, rpr, "delText"))
            new_nodes.append(deletion)

            insertion = revision_wrapper("ins", current_revision_id + 1, author, timestamp)
            insertion.append(build_run(new_text, rpr, "t"))
            new_nodes.append(insertion)

            current_revision_id += 2
            replaced += 1
        else:
            new_nodes.append(build_run(old_text, rpr, "t"))

        cursor = start + len(old_text)

        if occurrence is not None and match_index >= occurrence:
            break

    suffix = original_text[cursor:]
    if suffix:
        new_nodes.append(build_run(suffix, rpr, "t"))

    parent.remove(run)
    for offset, node in enumerate(new_nodes):
        parent.insert(insert_at + offset, node)
    set_paragraph_rsid(paragraph, rsid)
    return replaced


def run_contains_field_code(run: etree._Element) -> bool:
    """Check if run contains field code elements (fldChar, instrText)."""
    return (
        run.find("w:fldChar", namespaces=NS) is not None
        or run.find("w:instrText", namespaces=NS) is not None
    )


def get_paragraph_runs(paragraph: etree._Element) -> list[etree._Element]:
    """Get all w:r elements that contain text content (excluding fldChar runs)."""
    runs = []
    for run in paragraph.xpath("./w:r", namespaces=NS):
        # Skip runs that only contain fldChar (field codes)
        has_text = run.xpath("./w:t", namespaces=NS)
        if has_text:
            runs.append(run)
    return runs


def get_paragraph_runs_with_fields(paragraph: etree._Element) -> list[etree._Element]:
    """Get all w:r elements including those with field codes."""
    return paragraph.xpath("./w:r", namespaces=NS)


def build_run_text_index(runs: list[etree._Element]) -> tuple[str, list[tuple[int, int, etree._Element, int]]]:
    """
    Build a concatenated text string from runs and an index mapping positions to runs.
    Returns: (full_text, [(start, end, run, text_offset_in_run), ...])
    """
    full_text_parts = []
    index_map = []
    current_pos = 0
    
    for run in runs:
        text_nodes = run.xpath("./w:t", namespaces=NS)
        run_text = "".join(t.text or "" for t in text_nodes)
        if run_text:
            full_text_parts.append(run_text)
            index_map.append((current_pos, current_pos + len(run_text), run, len(full_text_parts) - 1))
            current_pos += len(run_text)
    
    return "".join(full_text_parts), index_map


def replace_text_across_runs(
    paragraph: etree._Element,
    runs: list[etree._Element],
    index_map: list[tuple[int, int, etree._Element, int]],
    match_start: int,
    match_end: int,
    old_text: str,
    new_text: str,
    author: str,
    revision_id: int,
    rsid: str | None,
) -> int:
    """
    Replace text that spans across multiple runs.
    Returns the next available revision_id.
    """
    timestamp = current_timestamp()
    
    # Find which runs are involved in the match
    involved_runs = []
    for start, end, run, _ in index_map:
        if end > match_start and start < match_end:
            involved_runs.append((start, end, run))
    
    if not involved_runs:
        return revision_id
    
    # Check if any involved run contains field codes
    # If so, skip to avoid breaking field code structure
    for _, _, run in involved_runs:
        if run_contains_field_code(run):
            print(
                f"警告: 跳过跨 run 替换，目标文本跨越 field code: '{old_text[:30]}...'",
                file=sys.stderr,
            )
            return revision_id
    
    # Get the first run's properties as the base style
    base_rpr = clone_run_properties(involved_runs[0][2])
    
    # Find the insertion point (before the first involved run)
    first_run = involved_runs[0][2]
    parent = first_run.getparent()
    if parent is None:
        return revision_id
    
    insert_at = parent.index(first_run)
    new_nodes: list[etree._Element] = []
    
    # Process prefix (text before match in the first run)
    first_run_start, first_run_end, first_run_el = involved_runs[0]
    prefix_len = match_start - first_run_start
    if prefix_len > 0:
        # Extract prefix text from the first run
        text_nodes = first_run_el.xpath("./w:t", namespaces=NS)
        prefix_text = ""
        prefix_remaining = prefix_len
        for t_node in text_nodes:
            t_text = t_node.text or ""
            if prefix_remaining <= 0:
                break
            if len(t_text) <= prefix_remaining:
                prefix_text += t_text
                prefix_remaining -= len(t_text)
            else:
                prefix_text += t_text[:prefix_remaining]
                prefix_remaining = 0
        if prefix_text:
            new_nodes.append(build_run(prefix_text, base_rpr, "t"))
    
    # Add deletion revision for old_text
    deletion = revision_wrapper("del", revision_id, author, timestamp)
    # Use the style from the first involved run for deletion
    del_rpr = clone_run_properties(first_run_el)
    deletion.append(build_run(old_text, del_rpr, "delText"))
    new_nodes.append(deletion)
    
    # Add insertion revision for new_text
    insertion = revision_wrapper("ins", revision_id + 1, author, timestamp)
    insertion.append(build_run(new_text, base_rpr, "t"))
    new_nodes.append(insertion)
    
    # Process suffix (text after match in the last run)
    last_run_start, last_run_end, last_run_el = involved_runs[-1]
    suffix_start_in_run = match_end - last_run_start
    last_run_text = run_text(last_run_el)
    suffix_text = last_run_text[suffix_start_in_run:]
    if suffix_text:
        new_nodes.append(build_run(suffix_text, base_rpr, "t"))
    
    # Remove all involved runs
    for _, _, run in involved_runs:
        if run.getparent() is not None:
            run.getparent().remove(run)
    
    # Insert new nodes
    for offset, node in enumerate(new_nodes):
        parent.insert(insert_at + offset, node)
    
    set_paragraph_rsid(paragraph, rsid)
    return revision_id + 2


def replace_inline(
    paragraph: etree._Element,
    old_text: str,
    new_text: str,
    author: str,
    revision_id: int,
    rsid: str | None,
    occurrence: int | None = None,
) -> tuple[int, int, str | None]:
    if not old_text:
        return 0, revision_id, None

    total_replaced = 0
    current_revision_id = revision_id
    modes_used: list[str] = []

    target_occurrence = occurrence
    remaining_occurrence = occurrence

    while True:
        replaced_this_round = 0
        mode_used: str | None = None

        # First pass: replace within single runs.
        for run in list(paragraph.xpath("./w:r", namespaces=NS)):
            replaced = replace_text_in_single_run(
                paragraph,
                run,
                old_text,
                new_text,
                author,
                current_revision_id,
                rsid,
                occurrence=remaining_occurrence,
            )
            if replaced:
                replaced_this_round = replaced
                current_revision_id += replaced * 2
                mode_used = "single-run"
                break

        # Second pass: replace one cross-run match, then re-scan the paragraph.
        if not replaced_this_round:
            runs = get_paragraph_runs(paragraph)
            if len(runs) > 1:
                full_text, index_map = build_run_text_index(runs)
                search_start = 0
                match_index = 0
                while True:
                    match_pos = full_text.find(old_text, search_start)
                    if match_pos == -1:
                        break

                    match_end = match_pos + len(old_text)
                    match_index += 1
                    should_replace = remaining_occurrence is None or match_index == remaining_occurrence
                    spans_multiple = False
                    for start, end, run, _ in index_map:
                        if start <= match_pos < end:
                            if match_end > end:
                                spans_multiple = True
                            break

                    if spans_multiple and should_replace:
                        next_revision_id = replace_text_across_runs(
                            paragraph,
                            runs,
                            index_map,
                            match_pos,
                            match_end,
                            old_text,
                            new_text,
                            author,
                            current_revision_id,
                            rsid,
                        )
                        if next_revision_id != current_revision_id:
                            replaced_this_round = 1
                            current_revision_id = next_revision_id
                            mode_used = "cross-run"
                        break

                    if should_replace and not spans_multiple:
                        break

                    search_start = match_end

        if not replaced_this_round:
            break

        total_replaced += replaced_this_round
        if mode_used:
            modes_used.append(mode_used)

        if target_occurrence is not None:
            break

    summary_mode = None
    if modes_used:
        summary_mode = "+".join(sorted(set(modes_used)))

    return total_replaced, current_revision_id, summary_mode


def delete_paragraph(
    paragraph: etree._Element,
    old_text: str,
    author: str,
    revision_id: int,
    rsid: str | None,
) -> bool:
    if paragraph_text(paragraph) != old_text:
        return False

    rpr = first_run_properties(paragraph)
    timestamp = current_timestamp()
    reset_paragraph_contents(paragraph)

    deletion = revision_wrapper("del", revision_id, author, timestamp)
    deletion.append(build_run(old_text, rpr, "delText"))
    paragraph.append(deletion)
    set_paragraph_rsid(paragraph, rsid)
    return True


def insert_after_paragraph(
    paragraph: etree._Element,
    anchor_text: str,
    new_text: str,
    author: str,
    revision_id: int,
    rsid: str | None,
) -> bool:
    if paragraph_text(paragraph) != anchor_text:
        return False

    timestamp = current_timestamp()
    rpr = first_run_properties(paragraph)
    new_paragraph = etree.Element(w_tag("p"))
    ppr = paragraph.find("w:pPr", namespaces=NS)
    if ppr is not None:
        new_paragraph.append(deepcopy(ppr))

    insertion = revision_wrapper("ins", revision_id, author, timestamp)
    insertion.append(build_run(new_text, rpr, "t"))
    new_paragraph.append(insertion)
    set_paragraph_rsid(new_paragraph, rsid)
    paragraph.addnext(new_paragraph)
    return True


def ensure_track_revisions(settings_root: etree._Element) -> str | None:
    rsid_root = settings_root.find(".//w:rsidRoot", namespaces=NS)
    rsid = rsid_root.get(w_tag("val")) if rsid_root is not None else None
    if settings_root.find("w:trackRevisions", namespaces=NS) is None:
        settings_root.append(etree.Element(w_tag("trackRevisions")))
    return rsid


# ─── Comments (批注) support ─────────────────────────────────────────────────

COMMENTS_NS = W_NS


def _comment_id_max(comments_root: etree._Element) -> int:
    """Return the highest w:id used in existing comments, or -1."""
    ids = comments_root.xpath("//@w:id", namespaces=NS)
    return max((int(i) for i in ids), default=-1)


def _load_or_create_comments(temp_path: Path) -> etree._Element:
    """Return the comments XML root element (creating the file if absent)."""
    comments_file = temp_path / "word" / "comments.xml"
    if comments_file.exists():
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(str(comments_file), parser=parser)
        return tree.getroot()
    # Build a fresh comments.xml
    root = etree.Element(w_tag("comments"))
    tree = etree.ElementTree(root)
    tree.write(
        str(comments_file),
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
    return root


def _insert_comment_in_run(
    paragraph: etree._Element,
    run: etree._Element,
    comment_id: int,
    author: str,
    comment_text: str,
    timestamp: str,
    comments_root: etree._Element,
) -> None:
    """
    Wrap the content of `run` with comment range markers and add a comment entry.
    The run text is preserved; comment markers are inserted before/after it.
    """
    run_text = "".join(node.text or "" for node in run.xpath("./w:t", namespaces=NS))
    rpr = clone_run_properties(run)
    parent = run.getparent()
    if parent is None:
        return
    insert_at = parent.index(run)

    # Append the new comment entry to comments.xml
    comment_elem = etree.SubElement(comments_root, w_tag("comment"))
    comment_elem.set(w_tag("id"), str(comment_id))
    comment_elem.set(w_tag("author"), author)
    comment_elem.set(w_tag("date"), timestamp)
    comment_elem.set(w_tag("initials"), author[:2].upper() if author else "YY")
    # Word stores comment text inside a w:r > w:t
    r_elem = etree.SubElement(comment_elem, w_tag("r"))
    t_elem = etree.SubElement(r_elem, w_tag("t"))
    t_elem.set(f"{{{XML_NS}}}space", "preserve")
    t_elem.text = comment_text

    # Build new nodes: commentRangeStart + run with commentReference + commentRangeEnd
    new_nodes: list[etree._Element] = []

    range_start = etree.Element(w_tag("commentRangeStart"))
    range_start.set(w_tag("id"), str(comment_id))
    new_nodes.append(range_start)

    # Clone the original run and add commentReference inside it
    ref_run = etree.Element(w_tag("r"))
    if rpr is not None:
        ref_run.append(deepcopy(rpr))
    ref_elem = etree.SubElement(ref_run, w_tag("commentReference"))
    ref_elem.set(w_tag("id"), str(comment_id))
    # Also include the original text so it remains visible
    if run_text:
        t_inside = etree.SubElement(ref_run, w_tag("t"))
        t_inside.set(f"{{{XML_NS}}}space", "preserve")
        t_inside.text = run_text
    new_nodes.append(ref_run)

    range_end = etree.Element(w_tag("commentRangeEnd"))
    range_end.set(w_tag("id"), str(comment_id))
    new_nodes.append(range_end)

    parent.remove(run)
    for offset, node in enumerate(new_nodes):
        parent.insert(insert_at + offset, node)


def _add_comment_to_paragraph(
    paragraph: etree._Element,
    target_text: str,
    comment_text: str,
    author: str,
    comment_id: int,
    comments_root: etree._Element,
) -> bool:
    """
    Scan `paragraph` for `target_text`.  When found, wrap it with comment markers.
    Returns True if the target was found and annotated.
    """
    timestamp = current_timestamp()
    for run in list(paragraph.xpath("./w:r", namespaces=NS)):
        if run_contains_field_code(run):
            continue
        run_txt = "".join(node.text or "" for node in run.xpath("./w:t", namespaces=NS))
        if target_text in run_txt:
            _insert_comment_in_run(
                paragraph, run, comment_id, author,
                comment_text, timestamp, comments_root,
            )
            return True
    return False


def apply_revisions(
    docx_in: Path,
    docx_out: Path,
    author: str,
    replacements: list[tuple[str, str]],
    inline_replacements: list[tuple[str, str]],
    inline_replacements_nth: list[tuple[str, str, int]],
    deletions: list[str],
    insertions: list[tuple[str, str]],
    comments: list[tuple[str, str]],
    inline_replacements_hf: list[tuple[str, str]] | None = None,
) -> tuple[int, list[InlineReplacementLog], list[CommentLog], list[HeaderFooterLog]]:
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        with zipfile.ZipFile(docx_in) as archive:
            archive.extractall(temp_path)

        document_path = temp_path / "word" / "document.xml"
        settings_path = temp_path / "word" / "settings.xml"
        parser = etree.XMLParser(remove_blank_text=False)
        document_tree = etree.parse(str(document_path), parser=parser)
        settings_tree = etree.parse(str(settings_path), parser=parser)
        document_root = document_tree.getroot()
        settings_root = settings_tree.getroot()
        rsid = ensure_track_revisions(settings_root)
        comments_root = _load_or_create_comments(temp_path)
        comment_id = _comment_id_max(comments_root) + 1
        comment_logs: list[CommentLog] = []
        hf_list = _iter_header_footer_xmls(temp_path)
        hf_logs: list[HeaderFooterLog] = []
        hf_replacements = inline_replacements_hf or []

        applied = 0
        revision_id = 1
        inline_logs: list[InlineReplacementLog] = []

        for old_text, new_text in replacements:
            matched = False
            for paragraph_index, paragraph, cell_ref in all_document_paragraphs(document_root):
                if replace_paragraph(
                    paragraph,
                    old_text,
                    new_text,
                    author,
                    revision_id,
                    rsid,
                ):
                    applied += 1
                    revision_id += 2
                    matched = True
                    break
            if not matched:
                print(f"未找到可替换段落: {old_text}", file=sys.stderr)

        for old_text, new_text in inline_replacements:
            matched = False
            for paragraph_index, paragraph, cell_ref in all_document_paragraphs(document_root):
                replaced, revision_id, mode = replace_inline(
                    paragraph,
                    old_text,
                    new_text,
                    author,
                    revision_id,
                    rsid,
                )
                if replaced:
                    applied += replaced
                    matched = True
                    location = cell_ref if cell_ref else f"第{paragraph_index}段"
                    inline_logs.append(
                        InlineReplacementLog(
                            old_text=old_text,
                            new_text=new_text,
                            paragraph_index=paragraph_index,
                            replaced_count=replaced,
                            mode=mode or "unknown",
                            target="all",
                            cell_ref=cell_ref,
                        )
                    )
            if not matched:
                print(f"未找到可局部替换文本: {old_text}", file=sys.stderr)

        for old_text, new_text, occurrence in inline_replacements_nth:
            matched = False
            for paragraph_index, paragraph, cell_ref in all_document_paragraphs(document_root):
                replaced, revision_id, mode = replace_inline(
                    paragraph,
                    old_text,
                    new_text,
                    author,
                    revision_id,
                    rsid,
                    occurrence=occurrence,
                )
                if replaced:
                    applied += replaced
                    matched = True
                    inline_logs.append(
                        InlineReplacementLog(
                            old_text=old_text,
                            new_text=new_text,
                            paragraph_index=paragraph_index,
                            replaced_count=replaced,
                            mode=mode or "unknown",
                            target=f"occurrence={occurrence}",
                            cell_ref=cell_ref,
                        )
                    )
                    break
            if not matched:
                print(f"未找到第 {occurrence} 处可局部替换文本: {old_text}", file=sys.stderr)

        for old_text in deletions:
            matched = False
            for _idx, paragraph, _cell_ref in all_document_paragraphs(document_root):
                if delete_paragraph(paragraph, old_text, author, revision_id, rsid):
                    applied += 1
                    revision_id += 1
                    matched = True
                    break
            if not matched:
                print(f"未找到可删除段落: {old_text}", file=sys.stderr)

        for anchor_text, new_text in insertions:
            matched = False
            for _idx, paragraph, _cell_ref in all_document_paragraphs(document_root):
                if insert_after_paragraph(
                    paragraph,
                    anchor_text,
                    new_text,
                    author,
                    revision_id,
                    rsid,
                ):
                    applied += 1
                    revision_id += 1
                    matched = True
                    break
            if not matched:
                print(f"未找到插入锚点段落: {anchor_text}", file=sys.stderr)

        for target_text, comment_text in comments:
            matched = False
            for paragraph_index, paragraph, cell_ref in all_document_paragraphs(document_root):
                if _add_comment_to_paragraph(
                    paragraph,
                    target_text,
                    comment_text,
                    author,
                    comment_id,
                    comments_root,
                ):
                    applied += 1
                    matched = True
                    comment_logs.append(
                        CommentLog(
                            target=target_text,
                            author=author,
                            paragraph_index=paragraph_index,
                            cell_ref=cell_ref,
                        )
                    )
                    comment_id += 1
                    break
            if not matched:
                print(f"未找到可添加批注的文本: {target_text}", file=sys.stderr)

        # ── Header / Footer inline replacements ───────────────────────────────
        for old_text, new_text in hf_replacements:
            for ref, hf_root in hf_list:
                for p_idx, paragraph in enumerate(hf_root.xpath(".//w:p", namespaces=NS), start=1):
                    replaced, revision_id, mode = replace_inline(
                        paragraph, old_text, new_text,
                        author, revision_id, rsid,
                    )
                    if replaced:
                        applied += replaced
                        hf_logs.append(
                            HeaderFooterLog(
                                old_text=old_text,
                                new_text=new_text,
                                replaced_count=replaced,
                                mode=mode or "unknown",
                                target="all",
                                ref=ref,
                            )
                        )

        document_tree.write(
            str(document_path),
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )
        settings_tree.write(
            str(settings_path),
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

        # Write comments.xml back if it was created or modified
        comments_file = temp_path / "word" / "comments.xml"
        if comments_file.exists():
            comments_tree = etree.ElementTree(comments_root)
            comments_tree.write(
                str(comments_file),
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True,
            )

        # Write modified header/footer XMLs back
        if hf_list:
            _rewrite_hf_xmls(hf_list, temp_path)

        with zipfile.ZipFile(docx_out, "w", zipfile.ZIP_DEFLATED) as archive:
            for file_path in temp_path.rglob("*"):
                if file_path.is_file():
                    archive.write(file_path, file_path.relative_to(temp_path))

    return applied, inline_logs, comment_logs, hf_logs


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Apply stable Word track changes to a DOCX.",
        epilog=(
            "Examples:\n"
            "  python3 scripts/track_changes.py in.docx out.docx --replace 'Old paragraph' 'New paragraph'\n"
            "  python3 scripts/track_changes.py in.docx out.docx --replace-inline 'old phrase' 'new phrase'\n"
            "  python3 scripts/track_changes.py in.docx out.docx --replace-inline-nth 'old phrase' 'new phrase' 2\n"
            "  python3 scripts/track_changes.py in.docx out.docx --delete 'Paragraph to remove'\n"
            "  python3 scripts/track_changes.py in.docx out.docx --insert-after 'Anchor paragraph' 'Inserted paragraph'\n"
            "  python3 scripts/track_changes.py in.docx out.docx --comment 'target phrase' 'Reviewer comment here'\n"
            "  python3 scripts/track_changes.py in.docx out.docx --replace-hf 'Company Name' 'New Company'"
        ),
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument("input_file")
    parser.add_argument("output_file")
    parser.add_argument("--author", default="Yachiyo")
    parser.add_argument(
        "--replace",
        nargs=2,
        action="append",
        metavar=("OLD", "NEW"),
        default=[],
    )
    parser.add_argument(
        "--replace-inline",
        nargs=2,
        action="append",
        metavar=("OLD", "NEW"),
        default=[],
        help="替换所有命中的局部文本（可跨多个段落、同段多处）",
    )
    parser.add_argument(
        "--replace-inline-nth",
        nargs=3,
        action="append",
        metavar=("OLD", "NEW", "N"),
        default=[],
        help="只替换每次扫描命中的第 N 处局部文本（当前按段落从前到后匹配，命中后停止）",
    )
    parser.add_argument("--delete", action="append", metavar="OLD", default=[])
    parser.add_argument(
        "--insert-after",
        nargs=2,
        action="append",
        metavar=("ANCHOR", "TEXT"),
        default=[],
    )
    parser.add_argument(
        "--comment",
        nargs=2,
        action="append",
        metavar=("TARGET", "TEXT"),
        default=[],
        help="在包含 TARGET 的文本位置添加一条批注（TEXT 为批注内容）",
    )
    parser.add_argument(
        "--replace-hf",
        nargs=2,
        action="append",
        metavar=("OLD", "NEW"),
        default=[],
        help="在页眉/页脚中替换所有命中的局部文本（--replace-inline 的页眉页脚版本）",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    inline_replacements_nth: list[tuple[str, str, int]] = []
    for old_text, new_text, occurrence_text in args.replace_inline_nth:
        try:
            occurrence = int(occurrence_text)
        except ValueError as exc:
            raise SystemExit(f"--replace-inline-nth 的 N 必须是正整数: {occurrence_text}") from exc
        if occurrence <= 0:
            raise SystemExit(f"--replace-inline-nth 的 N 必须是正整数: {occurrence_text}")
        inline_replacements_nth.append((old_text, new_text, occurrence))

    input_file = Path(args.input_file)
    output_file = Path(args.output_file)
    if not input_file.exists():
        raise SystemExit(f"输入文件不存在: {input_file}")

    if input_file.resolve() != output_file.resolve():
        shutil.copyfile(input_file, output_file)

    applied, inline_logs, comment_logs, hf_logs = apply_revisions(
        docx_in=input_file,
        docx_out=output_file,
        author=args.author,
        replacements=args.replace,
        inline_replacements=args.replace_inline,
        inline_replacements_nth=inline_replacements_nth,
        deletions=args.delete,
        insertions=args.insert_after,
        comments=args.comment,
        inline_replacements_hf=args.replace_hf,
    )
    print(f"已保存到: {output_file}")
    print(f"成功添加修订: {applied} 处")
    if inline_logs:
        print("局部替换日志:")
        for log in inline_logs:
            location = log.cell_ref if log.cell_ref else f"第{log.paragraph_index}段"
            print(
                f"- {location}: '{log.old_text}' -> '{log.new_text}' "
                f"命中 {log.replaced_count} 处 [{log.target}] ({log.mode})"
            )
    if comment_logs:
        print("批注日志:")
        for log in comment_logs:
            location = log.cell_ref if log.cell_ref else f"第{log.paragraph_index}段"
            print(f"- {location}: → 批注「{log.target}」({log.author})")
    if hf_logs:
        print("页眉/页脚替换日志:")
        done: set[str] = set()
        for log in hf_logs:
            key = f"{log.ref}:{log.old_text}"
            if key not in done:
                count = sum(1 for l in hf_logs if l.ref == log.ref and l.old_text == log.old_text)
                print(f"- {log.ref}: '{log.old_text}' -> '{log.new_text}' 命中 {count} 处")
                done.add(key)


if __name__ == "__main__":
    main()
