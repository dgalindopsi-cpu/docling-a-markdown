"""Extract editable text and SmartArt hierarchy from PowerPoint into Markdown.

This reads the OOXML package directly because many presentation converters omit
SmartArt diagram data, even when PowerPoint displays it as editable content.
"""

from __future__ import annotations

import argparse
from collections import defaultdict
from pathlib import Path
import posixpath
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def tag(prefix: str, name: str) -> str:
    return f"{{{NS[prefix]}}}{name}"


def paragraphs(container: ET.Element | None) -> list[str]:
    if container is None:
        return []
    result = []
    for paragraph in container.iter(tag("a", "p")):
        pieces = []
        for child in paragraph:
            if child.tag == tag("a", "br"):
                value = re.sub(r"\s+", " ", "".join(pieces)).strip()
                if value:
                    result.append(value)
                pieces = []
            else:
                pieces.extend((node.text or "") for node in child.iter(tag("a", "t")))
        value = re.sub(r"\s+", " ", "".join(pieces)).strip()
        if value:
            result.append(value)
    return result


def slide_relationships(archive: ZipFile, slide_number: int) -> dict[str, str]:
    rel_path = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
    if rel_path not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read(rel_path))
    result = {}
    for relation in root:
        target = relation.get("Target")
        if target and not target.startswith(("http:", "https:")):
            result[relation.get("Id", "")] = posixpath.normpath(
                posixpath.join("ppt/slides", target)
            )
    return result


def smartart_nodes(archive: ZipFile, data_path: str) -> tuple[list[tuple[int, list[str]]], int]:
    root = ET.fromstring(archive.read(data_path))
    points: dict[str, tuple[str | None, list[str]]] = {}
    point_order = []
    for point in root.findall(".//dgm:ptLst/dgm:pt", NS):
        ident = point.get("modelId")
        if not ident:
            continue
        points[ident] = (point.get("type"), paragraphs(point.find("dgm:t", NS)))
        point_order.append(ident)

    children: dict[str, list[tuple[int, str]]] = defaultdict(list)
    parents = set()
    for connection in root.findall(".//dgm:cxnLst/dgm:cxn", NS):
        if connection.get("type") not in (None, "parOf"):
            continue
        source, destination = connection.get("srcId"), connection.get("destId")
        if source in points and destination in points:
            children[source].append((int(connection.get("srcOrd", "0")), destination))
            parents.add(destination)

    roots = [ident for ident in point_order if points[ident][0] == "doc"]
    if not roots:
        roots = [ident for ident in point_order if ident not in parents]

    ordered = []
    seen = set()

    def visit(ident: str, depth: int) -> None:
        if ident in seen:
            return
        seen.add(ident)
        kind, content = points[ident]
        child_depth = depth
        if content and kind not in ("pres", "parTrans", "sibTrans"):
            ordered.append((depth, content))
            child_depth += 1
        for _, child in sorted(children.get(ident, [])):
            visit(child, child_depth)

    for ident in roots:
        visit(ident, 0)
    for ident in point_order:
        if ident not in seen and points[ident][1] and points[ident][0] not in (
            "pres", "parTrans", "sibTrans"
        ):
            visit(ident, 0)

    expected = sum(
        bool(content) and kind not in ("pres", "parTrans", "sibTrans")
        for kind, content in points.values()
    )
    return ordered, expected


def format_node(content: list[str]) -> str:
    if len(content) >= 2 and len(content[0]) <= 65:
        lead = content[0].rstrip(":")
        return f"**{lead}:** {' '.join(content[1:])}"
    return " ".join(content)


def markdown_table(table: ET.Element) -> tuple[list[str], int]:
    rows = []
    for row in table.findall("a:tr", NS):
        cells = []
        for cell in row.findall("a:tc", NS):
            value = " ".join(paragraphs(cell.find("a:txBody", NS)))
            cells.append(value.replace("|", "\\|"))
        rows.append(cells)
    if not rows:
        return [], 0
    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]
    result = ["| " + " | ".join(rows[0]) + " |"]
    result.append("| " + " | ".join(["---"] * width) + " |")
    result.extend("| " + " | ".join(row) + " |" for row in rows[1:])
    return result, sum(len(row) for row in rows)


def convert_file(input_path: Path) -> tuple[str, dict[str, int]]:
    with ZipFile(input_path) as archive:
        slide_numbers = sorted(
            int(match.group(1))
            for name in archive.namelist()
            if (match := re.fullmatch(r"ppt/slides/slide(\d+)\.xml", name))
        )
        if not slide_numbers:
            raise ValueError(f"El PowerPoint no contiene diapositivas: {input_path}")
        first_slide = ET.fromstring(archive.read(f"ppt/slides/slide{slide_numbers[0]}.xml"))
        cover_text = [
            value
            for shape in first_slide.findall(".//p:spTree/p:sp", NS)
            for value in paragraphs(shape.find("p:txBody", NS))
            if value.casefold() not in ("subtítulo", "subtitulo")
        ]
        deck_title = cover_text[-1] if cover_text else input_path.stem
        lines = [f"# {deck_title}", ""]
        metrics = {
            "slides": len(slide_numbers),
            "smartart_diagrams": 0,
            "smartart_nodes": 0,
            "emitted_nodes": 0,
            "tables": 0,
            "table_cells": 0,
            "unsupported_graphics": 0,
        }
        for number in slide_numbers:
            slide = ET.fromstring(archive.read(f"ppt/slides/slide{number}.xml"))
            relationships = slide_relationships(archive, number)
            slide_lines = []
            for shape in slide.findall(".//p:spTree/*", NS):
                if shape.tag == tag("p", "sp"):
                    for value in paragraphs(shape.find("p:txBody", NS)):
                        if value == deck_title or value.casefold() in ("subtítulo", "subtitulo"):
                            continue
                        slide_lines.extend([f"### {value}", ""])
                elif shape.tag == tag("p", "graphicFrame"):
                    refs = shape.find(".//dgm:relIds", NS)
                    if refs is None:
                        table = shape.find(".//a:tbl", NS)
                        if table is not None:
                            table_lines, cell_count = markdown_table(table)
                            metrics["tables"] += 1
                            metrics["table_cells"] += cell_count
                            slide_lines.extend([*table_lines, ""])
                        else:
                            metrics["unsupported_graphics"] += 1
                        continue
                    metrics["smartart_diagrams"] += 1
                    data_path = relationships.get(refs.get(tag("r", "dm"), ""))
                    if not data_path or data_path not in archive.namelist():
                        raise ValueError(
                            f"No se encontró el SmartArt de la diapositiva {number}: {data_path}"
                        )
                    nodes, expected = smartart_nodes(archive, data_path)
                    if len(nodes) != expected:
                        raise ValueError(
                            f"SmartArt incompleto en la diapositiva {number}: "
                            f"{len(nodes)}/{expected} nodos"
                        )
                    metrics["smartart_nodes"] += expected
                    metrics["emitted_nodes"] += len(nodes)
                    for depth, content in nodes:
                        slide_lines.append(f"{'  ' * depth}- {format_node(content)}")
                    if nodes:
                        slide_lines.append("")
            if slide_lines:
                lines.extend([f"## Diapositiva {number}", "", *slide_lines])
    return "\n".join(lines).rstrip() + "\n", metrics


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="A .pptx file or a folder of them")
    parser.add_argument("output", type=Path, help="Output folder for .md files")
    args = parser.parse_args()
    sources = sorted(args.input.glob("*.pptx")) if args.input.is_dir() else [args.input]
    args.output.mkdir(parents=True, exist_ok=True)
    for source in sources:
        markdown, metrics = convert_file(source)
        target = args.output / f"{source.stem}.md"
        target.write_text(markdown, encoding="utf-8")
        print(f"{source.name}: {metrics}, {len(markdown)} characters -> {target}")


if __name__ == "__main__":
    main()
