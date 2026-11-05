import argparse
import re
import sys
import zipfile
from pathlib import Path

from defusedxml import ElementTree as ET
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BLOCK_TAGS = {"p", "div", "br", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "table"}
PREFERRED_FIRST = ["ReqIF.ForeignID", "ReqIF.ChapterName", "ReqIF.Name", "ReqIF.Text"]
TEXT_HINTS = ["ReqIF.ChapterName", "ReqIF.Name", "ReqIF.Text"]
MAX_CELL = 32000
HEADER_LABELS = {
    "TITLE": "Title",
    "CREATION-TIME": "Created",
    "REQ-IF-TOOL-ID": "Exporting tool",
    "SOURCE-TOOL-ID": "Source tool",
    "COMMENT": "Comment",
}

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADING_FILL = PatternFill("solid", fgColor="DDEBF7")


def strip_namespaces(root):
    for el in root.iter():
        if isinstance(el.tag, str) and "}" in el.tag:
            el.tag = el.tag.split("}", 1)[1]
    return root


def load_roots(path):
    if path.suffix.lower() == ".reqifz":
        roots = []
        with zipfile.ZipFile(path) as archive:
            for name in archive.namelist():
                if name.lower().endswith(".reqif"):
                    roots.append((name, strip_namespaces(ET.fromstring(archive.read(name)))))
        if not roots:
            sys.exit(f"No .reqif file found inside {path.name}")
        return roots
    return [(path.name, strip_namespaces(ET.parse(path).getroot()))]


def xhtml_to_text(element):
    parts = []

    def walk(node):
        tag = node.tag if isinstance(node.tag, str) else ""
        if tag == "li":
            parts.append("\n• ")
        elif tag in BLOCK_TAGS:
            parts.append("\n")
        if node.text:
            parts.append(node.text)
        for child in node:
            walk(child)
            if child.tail:
                parts.append(child.tail)
        if tag in BLOCK_TAGS and tag != "br":
            parts.append("\n")

    for child in element:
        walk(child)
    if element.text:
        parts.insert(0, element.text)
    text = "".join(parts)
    text = re.sub(r"[ \t\r\f\v]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n+(• )", r"\n\1", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def ref_text(node, path):
    found = node.find(path)
    if found is None:
        return None
    if len(found):
        return (found[0].text or "").strip()
    return (found.text or "").strip()


class ReqIFDocument:
    def __init__(self, name, root):
        self.name = name
        self.enum_values = {}
        self.attr_names = {}
        self.attr_order = []
        self.type_names = {}
        self.objects = {}
        self.relations = []
        self.specifications = []
        self.header = {}
        self._parse(root)

    def _parse(self, root):
        header = root.find("THE-HEADER/REQ-IF-HEADER")
        if header is not None:
            for key in ("TITLE", "CREATION-TIME", "REQ-IF-TOOL-ID", "SOURCE-TOOL-ID", "COMMENT"):
                value = header.findtext(key)
                if value:
                    self.header[key] = value.strip()

        content = root.find("CORE-CONTENT/REQ-IF-CONTENT")
        if content is None:
            sys.exit(f"{self.name}: not a valid ReqIF file (REQ-IF-CONTENT missing)")

        for datatype in content.findall("DATATYPES/*"):
            for enum in datatype.findall("SPECIFIED-VALUES/ENUM-VALUE"):
                label = enum.get("LONG-NAME")
                if not label:
                    embedded = enum.find("PROPERTIES/EMBEDDED-VALUE")
                    label = embedded.get("OTHER-CONTENT") if embedded is not None else enum.get("IDENTIFIER")
                self.enum_values[enum.get("IDENTIFIER")] = label

        for spec_type in content.findall("SPEC-TYPES/*"):
            self.type_names[spec_type.get("IDENTIFIER")] = spec_type.get("LONG-NAME") or spec_type.get("IDENTIFIER")
            for attr in spec_type.findall("SPEC-ATTRIBUTES/*"):
                attr_id = attr.get("IDENTIFIER")
                attr_name = attr.get("LONG-NAME") or attr_id
                self.attr_names[attr_id] = attr_name
                if attr_name not in self.attr_order:
                    self.attr_order.append(attr_name)

        for obj in content.findall("SPEC-OBJECTS/SPEC-OBJECT"):
            self.objects[obj.get("IDENTIFIER")] = self._parse_object(obj)

        for rel in content.findall("SPEC-RELATIONS/SPEC-RELATION"):
            type_id = ref_text(rel, "TYPE")
            self.relations.append({
                "type": self.type_names.get(type_id, type_id or ""),
                "source": ref_text(rel, "SOURCE"),
                "target": ref_text(rel, "TARGET"),
            })

        for spec in content.findall("SPECIFICATIONS/SPECIFICATION"):
            rows = []
            self._walk_hierarchy(spec.find("CHILDREN"), [], rows)
            self.specifications.append({
                "name": spec.get("LONG-NAME") or spec.get("IDENTIFIER"),
                "rows": rows,
            })

    def _parse_object(self, obj):
        type_id = ref_text(obj, "TYPE")
        values = {}
        for value in obj.findall("VALUES/*"):
            definition = ref_text(value, "DEFINITION")
            name = self.attr_names.get(definition, definition)
            kind = value.tag.replace("ATTRIBUTE-VALUE-", "")
            if kind == "XHTML":
                the_value = value.find("THE-VALUE")
                text = xhtml_to_text(the_value) if the_value is not None else ""
            elif kind == "ENUMERATION":
                refs = [(r.text or "").strip() for r in value.findall("VALUES/ENUM-VALUE-REF")]
                text = ", ".join(self.enum_values.get(r, r) for r in refs)
            else:
                text = value.get("THE-VALUE", "")
            values[name] = text
        return {
            "id": obj.get("IDENTIFIER"),
            "type": self.type_names.get(type_id, type_id or ""),
            "last_change": obj.get("LAST-CHANGE", ""),
            "values": values,
        }

    def _walk_hierarchy(self, children, numbering, rows):
        if children is None:
            return
        position = 0
        for node in children.findall("SPEC-HIERARCHY"):
            position += 1
            number = numbering + [position]
            obj_id = ref_text(node, "OBJECT")
            sub = node.find("CHILDREN")
            has_children = sub is not None and len(sub.findall("SPEC-HIERARCHY")) > 0
            rows.append({
                "number": ".".join(map(str, number)),
                "level": len(number),
                "object_id": obj_id,
                "is_heading": has_children,
            })
            self._walk_hierarchy(sub, number, rows)

    def display_id(self, obj_id):
        obj = self.objects.get(obj_id)
        if obj and obj["values"].get("ReqIF.ForeignID"):
            return obj["values"]["ReqIF.ForeignID"]
        return obj_id

    def summary_text(self, obj_id, limit=80):
        obj = self.objects.get(obj_id)
        if not obj:
            return ""
        for key in TEXT_HINTS:
            if obj["values"].get(key):
                text = obj["values"][key]
                break
        else:
            text = next((v for v in obj["values"].values() if v), "")
        text = " ".join(text.split())
        return text if len(text) <= limit else text[: limit - 1] + "…"


def clean_cell(value):
    if value is None or isinstance(value, (int, float)):
        return value
    text = ILLEGAL_CHARACTERS_RE.sub("", str(value))
    return text if len(text) <= MAX_CELL else text[:MAX_CELL] + " …[truncated]"


def sheet_title(name, used):
    base = re.sub(r"[\[\]\:\*\?\/\\]", "_", name).strip() or "Sheet"
    base = base[:31]
    title, counter = base, 2
    while title.lower() in used:
        suffix = f" ({counter})"
        title = base[: 31 - len(suffix)] + suffix
        counter += 1
    used.add(title.lower())
    return title


def ordered_columns(doc, object_ids):
    present = set()
    for obj_id in object_ids:
        obj = doc.objects.get(obj_id)
        if obj:
            present.update(k for k, v in obj["values"].items() if v not in (None, ""))
    first = [c for c in PREFERRED_FIRST if c in present]
    rest = [c for c in doc.attr_order if c in present and c not in first]
    extra = sorted(present - set(first) - set(rest))
    return first + rest + extra


def style_header(ws, widths):
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width


def guess_width(values, minimum=8, maximum=70):
    longest = max((len(line) for v in values if v for line in str(v).split("\n")), default=minimum)
    return min(max(longest + 2, minimum), maximum)


def write_specification(ws, doc, spec):
    object_ids = [r["object_id"] for r in spec["rows"]]
    attrs = ordered_columns(doc, object_ids)
    indent_col = next((attrs.index(c) for c in TEXT_HINTS if c in attrs), 0 if attrs else None)
    headers = ["No", "Level", "Type", "Identifier"] + attrs
    ws.append(headers)
    columns = [[h] for h in headers]

    for row in spec["rows"]:
        obj = doc.objects.get(row["object_id"], {"type": "", "values": {}})
        values = [row["number"], row["level"], obj["type"], row["object_id"]]
        values += [obj["values"].get(a, "") for a in attrs]
        ws.append([clean_cell(v) for v in values])
        for col, v in zip(columns, values):
            col.append(v)
        excel_row = ws.max_row
        for cell in ws[excel_row]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(row=excel_row, column=2).alignment = Alignment(vertical="top", horizontal="center")
        text_col = next((attrs.index(c) for c in TEXT_HINTS if c in attrs and obj["values"].get(c)), indent_col)
        if text_col is not None:
            target = ws.cell(row=excel_row, column=5 + text_col)
            target.alignment = Alignment(vertical="top", wrap_text=True, indent=min(row["level"] - 1, 15))
        if row["is_heading"]:
            for cell in ws[excel_row]:
                cell.font = Font(bold=True)
                cell.fill = HEADING_FILL

    widths = [guess_width(c, maximum=12 if i < 2 else 70) for i, c in enumerate(columns)]
    style_header(ws, widths)
    return len(spec["rows"])


def write_relations(ws, doc, locations):
    headers = ["Relation Type", "Source No", "Source ID", "Source", "Target No", "Target ID", "Target"]
    ws.append(headers)
    columns = [[h] for h in headers]
    for rel in doc.relations:
        values = [
            rel["type"],
            locations.get(rel["source"], ""), doc.display_id(rel["source"]), doc.summary_text(rel["source"]),
            locations.get(rel["target"], ""), doc.display_id(rel["target"]), doc.summary_text(rel["target"]),
        ]
        ws.append([clean_cell(v) for v in values])
        for cell in ws[ws.max_row]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        for col, v in zip(columns, values):
            col.append(v)
    style_header(ws, [guess_width(c, maximum=50) for c in columns])


def write_summary(ws, docs, stats):
    ws.append(["Item", "Value"])
    for doc in docs:
        ws.append(["Source file", doc.name])
        for key, value in doc.header.items():
            ws.append([HEADER_LABELS.get(key, key), clean_cell(value)])
        ws.append(["Requirements (objects)", len(doc.objects)])
        ws.append(["Relations", len(doc.relations)])
        ws.append(["Specifications", len(doc.specifications)])
        ws.append(["", ""])
    ws.append(["Sheet", "Rows"])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
    ws.cell(row=ws.max_row, column=2).font = Font(bold=True)
    for title, count in stats:
        ws.append([title, count])
    style_header(ws, [28, 60])


def convert(input_path, output_path):
    docs = [ReqIFDocument(name, root) for name, root in load_roots(input_path)]
    wb = Workbook()
    summary_ws = wb.active
    summary_ws.title = "Summary"
    used = {"summary"}
    stats = []
    multi = len(docs) > 1

    for doc in docs:
        locations = {}
        placed = set()
        for spec in doc.specifications:
            label = f"{Path(doc.name).stem} - {spec['name']}" if multi else spec["name"]
            ws = wb.create_sheet(sheet_title(label, used))
            count = write_specification(ws, doc, spec)
            stats.append((ws.title, count))
            for row in spec["rows"]:
                placed.add(row["object_id"])
                locations.setdefault(row["object_id"], f"{spec['name']} / {row['number']}")

        loose = [oid for oid in doc.objects if oid not in placed]
        if loose:
            ws = wb.create_sheet(sheet_title("Unstructured", used))
            spec = {"rows": [{"number": str(i), "level": 1, "object_id": oid, "is_heading": False}
                             for i, oid in enumerate(loose, start=1)]}
            stats.append((ws.title, write_specification(ws, doc, spec)))

        if doc.relations:
            ws = wb.create_sheet(sheet_title("Relations", used))
            write_relations(ws, doc, locations)
            stats.append((ws.title, len(doc.relations)))

    write_summary(summary_ws, docs, stats)
    wb.save(output_path)
    return docs, stats


def main():
    parser = argparse.ArgumentParser(description="Convert ReqIF / ReqIFz requirement files into a structured Excel workbook.")
    parser.add_argument("input", help="Path to a .reqif or .reqifz file")
    parser.add_argument("-o", "--output", help="Output .xlsx file (default: <input name>.xlsx)")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        sys.exit(f"File not found: {input_path}")
    output_path = Path(args.output) if args.output else input_path.with_suffix(".xlsx")

    docs, stats = convert(input_path, output_path)
    objects = sum(len(d.objects) for d in docs)
    relations = sum(len(d.relations) for d in docs)
    specs = sum(len(d.specifications) for d in docs)
    print(f"Specifications: {specs} | Requirements: {objects} | Relations: {relations}")
    for title, count in stats:
        print(f"  - {title}: {count} rows")
    print(f"Saved -> {output_path}")


if __name__ == "__main__":
    main()
