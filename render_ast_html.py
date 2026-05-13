#!/usr/bin/env python3
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
import html
from pathlib import Path

SRC = Path("AST_MANUAL.md")
OUT = Path("AST_MANUAL.html")


@dataclass
class Node:
    text: str
    children: list["Node"] = field(default_factory=list)


def parse_lines(lines: list[str]) -> Node:
    root = Node("AST")
    stack: list[tuple[int, Node]] = [(-1, root)]
    prev_indent = -1
    prev_was_bullet = False

    for raw in lines:
        if not raw.strip():
            continue
        if raw.lstrip().startswith("#"):
            continue

        indent = len(raw) - len(raw.lstrip(" "))
        stripped = raw.strip()
        is_bullet = stripped.startswith("- ")
        text = stripped[2:] if is_bullet else stripped

        # If we have a non-bullet followed by a bullet at same indent,
        # treat the bullet as a child of the previous node.
        if is_bullet and (not prev_was_bullet) and indent == prev_indent:
            indent = prev_indent + 2

        node = Node(text)
        while stack and indent <= stack[-1][0]:
            stack.pop()
        stack[-1][1].children.append(node)
        stack.append((indent, node))

        prev_indent = indent
        prev_was_bullet = is_bullet

    return root


def to_mermaid(tree: Node) -> str:
    lines = ["graph TD"]
    counter = 0

    def esc(label: str) -> str:
        safe = label.replace("\"", "'")
        return html.escape(safe)

    def walk(node: Node, parent_id: str | None = None) -> None:
        nonlocal counter
        node_id = f"n{counter}"
        counter += 1
        lines.append(f"{node_id}[\"{esc(node.text)}\"]")
        if parent_id:
            lines.append(f"{parent_id} --> {node_id}")
        for child in node.children:
            walk(child, node_id)

    for child in tree.children:
        walk(child)

    return "\n".join(lines)


def main() -> None:
    lines = SRC.read_text(encoding="utf-8").splitlines()
    tree = parse_lines(lines)
    mermaid = to_mermaid(tree)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    html_doc = f"""<!doctype html>
<html lang=\"es\">
<head>
  <meta charset=\"utf-8\">
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
  <title>AST Manual</title>
  <style>
    :root {{ color-scheme: light; }}
    body {{ font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace; margin: 24px; background: #f6f5f2; color: #1e1e1e; }}
    header {{ margin-bottom: 16px; }}
    h1 {{ font-size: 20px; margin: 0 0 6px 0; }}
    .meta {{ font-size: 12px; color: #5b5b5b; }}
    .card {{ background: #ffffff; border: 1px solid #e6e2db; border-radius: 8px; padding: 16px; }}
    .mermaid {{ background: #ffffff; padding: 12px; border-radius: 8px; }}
  </style>
  <script type=\"module\">
    import mermaid from "https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs";
    mermaid.initialize({{ startOnLoad: true, theme: "default" }});
  </script>
</head>
<body>
  <header>
    <h1>AST Manual (grafo)</h1>
    <div class=\"meta\">Fuente: {SRC.name} · Generado: {timestamp}</div>
  </header>
  <div class=\"card\">
    <pre class=\"mermaid\">\n{mermaid}\n</pre>
  </div>
</body>
</html>"""

    OUT.write_text(html_doc, encoding="utf-8")
    print(f"HTML generado: {OUT}")


if __name__ == "__main__":
    main()
