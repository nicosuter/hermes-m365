# pyright: reportMissingImports=false

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from html import unescape

from bs4 import BeautifulSoup, Tag


HIDDEN_STYLE_PATTERNS = (
    re.compile(r"(?:^|;)\s*display\s*:\s*none\b", re.IGNORECASE),
    re.compile(r"(?:^|;)\s*visibility\s*:\s*hidden\b", re.IGNORECASE),
    re.compile(r"(?:^|;)\s*opacity\s*:\s*0(?:\.0+)?\s*(?:;|$)", re.IGNORECASE),
    re.compile(r"(?:^|;)\s*font-size\s*:\s*0(?:px|em|rem|pt|%)?\s*(?:;|$)", re.IGNORECASE),
    re.compile(r"(?:^|;)\s*max-height\s*:\s*0(?:px|em|rem|pt|%)?\s*(?:;|$)", re.IGNORECASE),
)
_CSS_RULE_RE = re.compile(r"([^{}]+)\{([^{}]*)\}")


class _HiddenSelectors:
    classes: set[str]
    ids: set[str]
    tag_classes: dict[str, set[str]]
    tag_ids: dict[str, set[str]]

    def __init__(self) -> None:
        self.classes = set()
        self.ids = set()
        self.tag_classes = {}
        self.tag_ids = {}


def _collect_hidden_selectors(soup: BeautifulSoup) -> _HiddenSelectors:
    selectors = _HiddenSelectors()
    for style_tag in soup.find_all("style"):
        style_text = style_tag.get_text()
        for selector_str, declarations in _CSS_RULE_RE.findall(style_text):
            if not any(pattern.search(declarations) for pattern in HIDDEN_STYLE_PATTERNS):
                continue
            for selector in selector_str.split(","):
                selector = selector.strip()
                if not selector or " " in selector or ":" in selector or "[" in selector:
                    continue
                if selector.startswith("."):
                    selectors.classes.add(selector[1:])
                elif selector.startswith("#"):
                    selectors.ids.add(selector[1:])
                elif "." in selector:
                    parts = selector.split(".", 1)
                    selectors.tag_classes.setdefault(parts[0], set()).add(parts[1])
                elif "#" in selector:
                    parts = selector.split("#", 1)
                    selectors.tag_ids.setdefault(parts[0], set()).add(parts[1])
    return selectors


def _remove_css_hidden_elements(soup: BeautifulSoup, selectors: _HiddenSelectors) -> None:
    for cls_name in selectors.classes:
        for element in list(soup.find_all(class_=lambda c, name=cls_name: _class_match(c, name))):
            element.decompose()
    for id_name in selectors.ids:
        element = soup.find(id=id_name)
        if element:
            element.decompose()
    for tag_name, classes in selectors.tag_classes.items():
        for cls_name in classes:
            for element in list(soup.find_all(tag_name, class_=lambda c, name=cls_name: _class_match(c, name))):
                element.decompose()
    for tag_name, ids in selectors.tag_ids.items():
        for id_name in ids:
            element = soup.find(tag_name, id=id_name)
            if element:
                element.decompose()


def _class_match(raw_class: object | None, target: str) -> bool:
    if raw_class is None:
        return False
    if isinstance(raw_class, list):
        return target in raw_class
    if isinstance(raw_class, str):
        return target in raw_class.split()
    return False


CID_REFERENCE_RE = re.compile(r"(?i)cid:([^\s\"'<>)]*)")


def sanitize_html_body(html: str) -> str:
    """Sanitize HTML email body to plain text.

    Removes scripts, styles, hidden elements (inline styles and CSS class-based),
    and unsafe tags. Converts HTML entities and normalizes whitespace.
    """
    if not html:
        return ""

    soup = BeautifulSoup(html, "html.parser")
    hidden_selectors = _collect_hidden_selectors(soup)
    _remove_css_hidden_elements(soup, hidden_selectors)

    for element in list(soup.find_all(_is_hidden_or_unsafe)):
        element.decompose()

    for tag in soup.find_all(("br", "p", "div", "li", "tr", "table", "section", "article", "h1", "h2", "h3", "h4", "h5", "h6")):
        _ = tag.append("\n")

    text = soup.get_text(separator=" ")
    return _normalize_text(unescape(text))


def insert_inline_attachment_markers(body: str, inline_attachments: Sequence[Mapping[str, object]]) -> str:
    if not inline_attachments:
        return body

    attachments_by_cid = {
        _normalize_content_id(attachment.get("contentId")): attachment
        for attachment in inline_attachments
        if _normalize_content_id(attachment.get("contentId"))
    }
    inserted_cids: set[str] = set()

    def replace_reference(match: re.Match[str]) -> str:
        raw_cid = match.group(1).rstrip(".,;:")
        suffix = match.group(1)[len(raw_cid) :]
        normalized_cid = _normalize_content_id(raw_cid)
        attachment = attachments_by_cid.get(normalized_cid)
        if not attachment:
            return match.group(0)
        inserted_cids.add(normalized_cid)
        return f"{_inline_marker(attachment)}{suffix}"

    marked_body = CID_REFERENCE_RE.sub(replace_reference, body)
    fallback_markers = [
        _inline_marker(attachment)
        for attachment in inline_attachments
        if _normalize_content_id(attachment.get("contentId")) not in inserted_cids
    ]
    if not fallback_markers:
        return marked_body

    separator = "\n\n" if marked_body.strip() else ""
    return f"{marked_body.rstrip()}{separator}Inline attachments:\n" + "\n".join(fallback_markers)


def _is_hidden_or_unsafe(element: Tag) -> bool:
    if element.name in {"script", "style", "noscript", "template", "head", "meta", "link", "title"}:
        return True
    if element.has_attr("hidden") or element.get("aria-hidden") == "true":
        return True
    style = str(element.get("style", ""))
    return any(pattern.search(style) for pattern in HIDDEN_STYLE_PATTERNS)


def _normalize_text(text: str) -> str:
    lines = [re.sub(r"[ \t\r\f\v\xa0]+", " ", line).strip() for line in text.splitlines()]
    collapsed: list[str] = []
    previous_blank = False
    for line in lines:
        if not line:
            if not previous_blank and collapsed:
                collapsed.append("")
            previous_blank = True
            continue
        collapsed.append(line)
        previous_blank = False
    return "\n".join(collapsed).strip()


def _normalize_content_id(content_id: object) -> str:
    if not isinstance(content_id, str):
        return ""
    return content_id.strip().strip("<>").lower()


def _inline_marker(attachment: Mapping[str, object]) -> str:
    filename = _attachment_text(attachment.get("name")) or _attachment_text(attachment.get("filename")) or "unnamed attachment"
    mime_type = _attachment_text(attachment.get("contentType")) or _attachment_text(attachment.get("mimeType")) or "application/octet-stream"
    return f'[Inline attachment called "{filename}" ({mime_type}), use get_attachment to fetch]'


def _attachment_text(value: object) -> str:
    return value.strip() if isinstance(value, str) else ""
