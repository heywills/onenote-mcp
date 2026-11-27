"""Convert simple HTML into OneNote page XML elements.

OneNote's <one:T> holds inline HTML only. Block elements become OneNote structure:
- <p>, <div>, <br>: one OE per paragraph or line
- <h1>-<h6>: OE with a heading QuickStyleDef
- <ul>, <ol>, <li>: OE with <one:List>, nested lists in OEChildren
- <table>: <one:Table>
- <pre>: one monospace OE per line
"""

import html
import re
import xml.etree.ElementTree as ET
from html.parser import HTMLParser

NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
ET.register_namespace("one", NS)

# OneNote's built-in heading styles: (font size, color, italic). h2-h4 match pages made in OneNote.
HEADING_STYLES = {
    1: ("16.0", "#1E4E79", False),
    2: ("14.0", "#2E75B5", False),
    3: ("12.0", "#5B9BD5", False),
    4: ("12.0", "#5B9BD5", True),
    5: ("11.0", "#2E75B5", False),
    6: ("11.0", "#2E75B5", True),
}
INLINE = {"b", "strong", "i", "em", "u", "s", "strike", "del", "sup", "sub", "a", "span", "code"}
VOID = {"br", "hr", "img", "input", "meta", "link"}


def _q(tag: str) -> str:
    return f"{{{NS}}}{tag}"


class _Node:
    def __init__(self, tag: str, attrs: dict | None = None, parent=None):
        self.tag, self.attrs, self.parent, self.children = tag, attrs or {}, parent, []


class _TreeBuilder(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = _Node("root")
        self.cur = self.root

    def handle_starttag(self, tag, attrs):
        node = _Node(tag, dict(attrs), self.cur)
        self.cur.children.append(node)
        if tag not in VOID:
            self.cur = node

    def handle_endtag(self, tag):
        n = self.cur
        while n is not self.root and n.tag != tag:
            n = n.parent
        if n is not self.root:
            self.cur = n.parent

    def handle_data(self, data):
        self.cur.children.append(data)


def _inline_html(nodes) -> str:
    """Render inline nodes as the HTML OneNote accepts inside <one:T>."""
    out = []
    for n in nodes:
        if isinstance(n, str):
            out.append(html.escape(re.sub(r"\s+", " ", n), quote=False))
        elif n.tag in ("b", "strong"):
            out.append(f"<b>{_inline_html(n.children)}</b>")
        elif n.tag in ("i", "em"):
            out.append(f"<i>{_inline_html(n.children)}</i>")
        elif n.tag == "u":
            out.append(f"<u>{_inline_html(n.children)}</u>")
        elif n.tag in ("s", "strike", "del"):
            out.append(f"<span style='text-decoration:line-through'>{_inline_html(n.children)}</span>")
        elif n.tag in ("sup", "sub"):
            out.append(f"<{n.tag}>{_inline_html(n.children)}</{n.tag}>")
        elif n.tag == "code":
            out.append(f"<span style='font-family:Consolas'>{_inline_html(n.children)}</span>")
        elif n.tag == "a":
            href = html.escape(n.attrs.get("href", ""), quote=True)
            out.append(f'<a href="{href}">{_inline_html(n.children)}</a>')
        else:
            out.append(_inline_html(n.children))
    return "".join(out)


class _Builder:
    def __init__(self):
        self.headings: set[int] = set()

    def oe(self, parent, inner_html: str = "", style: int | None = None, list_kind: str | None = None):
        oe = ET.SubElement(parent, _q("OE"))
        if style is not None:
            oe.set("quickStyleIndex", str(style))
        if list_kind:
            lst = ET.SubElement(oe, _q("List"))
            if list_kind == "ol":
                ET.SubElement(lst, _q("Number"), numberSequence="0", numberFormat="##.")
            else:
                ET.SubElement(lst, _q("Bullet"), bullet="2")
        ET.SubElement(oe, _q("T")).text = inner_html.strip()
        return oe

    def blocks(self, nodes, parent):
        """Emit block content of `nodes` as OEs under `parent` (an OEChildren)."""
        inline_run = []

        def flush():
            text = _inline_html(inline_run).strip()
            if text:
                self.oe(parent, text)
            inline_run.clear()

        for n in nodes:
            if isinstance(n, str) or n.tag in INLINE:
                inline_run.append(n)
                continue
            flush()
            if n.tag == "br":
                continue
            if n.tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
                level = int(n.tag[1])
                self.headings.add(level)
                if len(parent):
                    self.oe(parent, "")  # blank line before a heading, like hand-made OneNote pages
                self.oe(parent, _inline_html(n.children), style=level)
            elif n.tag in ("ul", "ol"):
                self.list_items(n, parent)
            elif n.tag == "table":
                self.table(n, parent)
            elif n.tag == "pre":
                text = "".join(x if isinstance(x, str) else _text(x) for x in n.children).strip("\n")
                for line in text.split("\n"):
                    esc = html.escape(line, quote=False).replace(" ", "&nbsp;") or "&nbsp;"
                    self.oe(parent, f"<span style='font-family:Consolas'>{esc}</span>")
            elif n.tag == "hr":
                self.oe(parent, "―" * 20)
            else:  # p, div, blockquote, section, li outside a list, etc.
                self.blocks(n.children, parent)
        flush()

    def list_items(self, list_node, parent):
        for li in list_node.children:
            if isinstance(li, str) or li.tag != "li":
                continue
            inline = [c for c in li.children if isinstance(c, str) or c.tag in INLINE]
            nested = [c for c in li.children if not isinstance(c, str) and c.tag not in INLINE]
            oe = self.oe(parent, _inline_html(inline), list_kind=list_node.tag)
            if nested:
                self.blocks(nested, ET.SubElement(oe, _q("OEChildren")))

    def table(self, table_node, parent):
        rows = [r for r in _iter(table_node) if r.tag == "tr"]
        if not rows:
            return
        cells = [[c for c in r.children if not isinstance(c, str) and c.tag in ("td", "th")] for r in rows]
        ncols = max(len(r) for r in cells) or 1
        oe = ET.SubElement(parent, _q("OE"))
        tbl = ET.SubElement(oe, _q("Table"), bordersVisible="true")
        cols = ET.SubElement(tbl, _q("Columns"))
        for i in range(ncols):
            ET.SubElement(cols, _q("Column"), index=str(i), width="120")
        for row in cells:
            r = ET.SubElement(tbl, _q("Row"))
            for i in range(ncols):
                cell = ET.SubElement(r, _q("Cell"))
                children = ET.SubElement(cell, _q("OEChildren"))
                if i < len(row):
                    content = row[i].children
                    if row[i].tag == "th":
                        content = [_Node("b", parent=None)]
                        content[0].children = row[i].children
                    self.blocks(content, children)
                if not len(children):
                    self.oe(children, "")


def _iter(node):
    for c in node.children:
        if not isinstance(c, str):
            yield c
            yield from _iter(c)


def _text(node) -> str:
    return "".join(c if isinstance(c, str) else _text(c) for c in node.children)


def add_html_to_page(page_root: ET.Element, content_html: str) -> None:
    """Append `content_html` to a OneNote page element as an Outline, adding heading styles it needs."""
    tb = _TreeBuilder()
    tb.feed(content_html)
    tb.close()

    builder = _Builder()
    outline = ET.Element(_q("Outline"))
    builder.blocks(tb.root.children, ET.SubElement(outline, _q("OEChildren")))

    # Heading styles: QuickStyleDef index = heading level. Must sit with the other QuickStyleDefs.
    existing = {q.get("index") for q in page_root.findall(_q("QuickStyleDef"))}
    defs = page_root.findall(_q("QuickStyleDef"))
    insert_at = (list(page_root).index(defs[-1]) + 1) if defs else 0
    for level in sorted(builder.headings):
        if str(level) in existing:
            continue
        size, color, italic = HEADING_STYLES[level]
        q = ET.Element(
            _q("QuickStyleDef"), index=str(level), name=f"h{level}", fontColor=color,
            highlightColor="automatic", font="Calibri", fontSize=size,
            spaceBefore="0.0", spaceAfter="0.0",
        )
        if italic:
            q.set("italic", "true")
        page_root.insert(insert_at, q)
        insert_at += 1

    page_root.append(outline)
