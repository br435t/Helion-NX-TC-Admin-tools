"""Excel editor for the Helion MatML material library.

    python material_xlsx.py export [xml] [xlsx] [--props A,B | --all]
                                                  XML  -> Excel (one row per material; columns: Name, Class, Subclass, MassDensityConstant_110)
    python material_xlsx.py import [xlsx] [xml]   Excel -> XML (only changed cells are written; backup is made first)

Standard library only. Untouched parts of the XML are preserved byte-for-byte.
"""
import argparse
import datetime
import re
import shutil
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent
DEFAULT_XML = HERE / "Materials" / "helion_material_library.xml"
DEFAULT_XLSX = HERE / "helion_material_library.xlsx"

FIXED_COLS = ["#", "Name", "Class", "Subclass"]
DEFAULT_PROPS = ["MassDensityConstant_110"]
VALUES_SUFFIX = " [values]"
DATA_INDENT = "\n" + " " * 16


# ---------------------------------------------------------------- XML helpers

def load_xml(path):
    raw = Path(path).read_bytes()
    return ET.fromstring(raw), raw


def serialize(root, original_raw):
    """Serialize `root` in the same style as the source file (CRLF, '<X/>', blank line after prolog)."""
    text = ET.tostring(root, encoding="unicode").replace(" />", "/>")
    out = '<?xml version="1.0" encoding="UTF-8"?>\n\n' + text
    if original_raw.rstrip(b"\r\n") != original_raw:
        out += "\n"
    return out.replace("\n", "\r\n").encode("utf-8")


def parse_numbers(text, fmt):
    """'1.0e+00 , 2.0e+00' -> [1.0, 2.0]; raises ValueError on bad input."""
    parts = [p.strip() for p in (text or "").split(",") if p.strip()]
    if fmt != "integer":
        return [float(p) for p in parts]
    out = []
    for p in parts:
        f = float(p)
        if f != int(f):
            raise ValueError(p)
        out.append(int(f))
    return out


def format_numbers(vals, fmt):
    if fmt == "integer":
        return " , ".join(str(int(v)) for v in vals)
    return " , ".join("%.6e" % v for v in vals)


def cell_text(data):
    """Display text of a <Data> element."""
    t = (data.text or "").strip()
    return re.sub(r"\s+", " ", t) if data.get("format") != "string" else t


def split_ws(text):
    text = text or ""
    lead = text[: len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()):]
    return lead, trail


def set_data_text(data, core):
    lead, trail = split_ws(data.text)
    if not lead and not trail and not (data.text or "").strip():
        lead, trail = DATA_INDENT, DATA_INDENT[:-4]
    data.text = lead + core + trail


def qualifier_map(root):
    """(property, integer value text) -> Qualifier text, harvested from the existing library."""
    qm = {}
    for pd in root.iter("PropertyData"):
        d, q = pd.find("Data"), pd.find("Qualifier")
        if d is not None and q is not None and d.get("format") == "integer" and (q.text or "").strip():
            qm.setdefault((pd.get("property"), (d.text or "").strip()), q.text)
    return qm


def collect_schema(root):
    """Ordered property columns: [(prop_id, units_text, has_values_column, values_units_text)]."""
    units = {}
    for pdet in root.iter("PropertyDetails"):
        u = pdet.find("Units")
        units[pdet.get("id")] = u.get("name") if u is not None else ""
    param_units = {}
    for pdet in root.iter("ParameterDetails"):
        u = pdet.find("Units")
        param_units[pdet.get("id")] = (pdet.findtext("Name") or "", u.get("name") if u is not None else "")
    order, has_pv = [], {}
    for pd in root.iter("PropertyData"):
        pid = pd.get("property")
        if pid not in has_pv:
            order.append(pid)
            has_pv[pid] = None
        pv = pd.find("ParameterValue")
        if pv is not None and has_pv[pid] is None:
            has_pv[pid] = pv.get("parameter")
    # prefer the declaration order from Metadata when available
    meta_order = [d.get("id") for d in root.iter("PropertyDetails")]
    order = [p for p in meta_order if p in has_pv] + [p for p in order if p not in set(meta_order)]
    schema = []
    for pid in order:
        param = has_pv[pid]
        pname, punits = param_units.get(param, ("", ""))
        schema.append((pid, units.get(pid, ""), param, (pname + (" (" + punits + ")" if punits else "")) if param else ""))
    return schema


# ---------------------------------------------------------------- xlsx writer

def col_letter(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _cell(ref, value, style=0):
    if value is None or value == "":
        return f'<c r="{ref}" s="{style}"/>' if style else ""
    if isinstance(value, (int, float)):
        return f'<c r="{ref}" s="{style}"><v>{value!r}</v></c>'
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{escape(_BAD.sub("", value))}</t></is></c>'


def write_xlsx(path, labels, header, rows, help_lines):
    ncols = len(header)
    sheet = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
             f'<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="2" topLeftCell="C3" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>'
             '<sheetFormatPr defaultRowHeight="15"/><cols>'
             '<col min="1" max="1" width="5" customWidth="1"/><col min="2" max="2" width="32" customWidth="1"/>'
             f'<col min="3" max="{ncols}" width="22" customWidth="1"/></cols><sheetData>']
    for r, (vals, style) in enumerate([(labels, 2), (header, 0)] + [(x, 0) for x in rows], start=1):
        cells = "".join(_cell(f"{col_letter(c)}{r}", v, style) for c, v in enumerate(vals))
        sheet.append(f'<row r="{r}">{cells}</row>')
    sheet.append('</sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>')
    ref = f"A2:{col_letter(ncols - 1)}{len(rows) + 2}"
    table = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Materials" displayName="Materials" '
             f'ref="{ref}" totalsRowShown="0"><autoFilter ref="{ref}"/><tableColumns count="{ncols}">'
             + "".join(f'<tableColumn id="{i}" name="{escape(h, {chr(34): "&quot;"})}"/>' for i, h in enumerate(header, start=1))
             + '</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>')

    help_rows = "".join(
        f'<row r="{i}">{_cell(f"A{i}", line)}</row>' for i, line in enumerate(help_lines, start=1))
    help_sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                  '<cols><col min="1" max="1" width="120" customWidth="1"/></cols>'
                  f'<sheetData>{help_rows}</sheetData></worksheet>')

    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
              '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font>'
              '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
              '<font><i/><sz val="10"/><color rgb="FF595959"/><name val="Calibri"/></font></fonts>'
              '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
              '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill>'
              '<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill></fills>'
              '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
              '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
              '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
              '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
              '<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>'
              '</styleSheet>')
    files = {
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
        "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Materials" sheetId="1" r:id="rId1"/><sheet name="Help" sheetId="2" r:id="rId2"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        "xl/styles.xml": styles,
        "xl/worksheets/sheet1.xml": "".join(sheet),
        "xl/worksheets/_rels/sheet1.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
        "xl/tables/table1.xml": table,
        "xl/worksheets/sheet2.xml": help_sheet,
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in files.items():
            z.writestr(name, content)


# ---------------------------------------------------------------- xlsx reader

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def col_index(ref):
    n = 0
    for ch in re.match(r"[A-Z]+", ref).group():
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_xlsx(path):
    """Return the first sheet as a list of rows (lists of str, '' for empty)."""
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", NS):
                shared.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid = wb.find("m:sheets/m:sheet", NS).get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = next(r.get("Target") for r in rels if r.get("Id") == rid).lstrip("/")
        sheet = ET.fromstring(z.read(target if target.startswith("xl/") else "xl/" + target))
    rows = []
    for row in sheet.find("m:sheetData", NS):
        r = int(row.get("r")) - 1
        while len(rows) <= r:
            rows.append([])
        for c in row:
            i = col_index(c.get("r"))
            t = c.get("t")
            if t == "inlineStr":
                v = "".join(x.text or "" for x in c.iter("{%s}t" % NS["m"]))
            else:
                ve = c.find("m:v", NS)
                v = "" if ve is None or ve.text is None else ve.text
                if t == "s" and v != "":
                    v = shared[int(v)]
                elif t == "b":
                    v = "TRUE" if v == "1" else "FALSE"
            while len(rows[r]) <= i:
                rows[r].append("")
            rows[r][i] = v
    return rows


# ---------------------------------------------------------------- export

HELP = [
    "Helion material library editor",
    "",
    "Sheet 'Materials': one row per material, one column per property (row 1 = labels/units, row 2 = table header: MatML property id).",
    "Columns named '<property> [values]' hold the dependent table of a temperature-dependent property;",
    "the base column then holds the temperatures. Both are comma-separated lists and must have the same length.",
    "",
    "Editing rules:",
    "  - Do not edit column A (#), the label row 1 or the table header row 2. Do not reorder or delete rows/columns.",
    "  - Column B (Name), C (Class) and D (Subclass) can be renamed.",
    "  - 'integer' properties take whole numbers; the matching description (e.g. 'von Mises') is updated automatically",
    "    when the value is already used elsewhere in the library. Otherwise it is kept and a warning is printed.",
    "  - 'exponential' properties take any number (1800 or 1.8e3); they are written back as 1.800000e+03.",
    "  - A blank cell on a property the material does not have means 'not present'. Fill it in to add the property.",
    "  - Deleting a material or removing a property is not supported.",
    "",
    "Save the workbook, close Excel, then run import_to_xml.bat. A timestamped backup of the XML is made first.",
]


def do_export(xml_path, xlsx_path, props=None):
    """props: property ids to expose as columns; None = every property in the library."""
    root, _ = load_xml(xml_path)
    schema = collect_schema(root)
    if props is not None:
        known = {s[0] for s in schema}
        unknown = [p for p in props if p not in known]
        if unknown:
            sys.exit("Unknown property id(s): " + ", ".join(unknown))
        schema = [s for s in schema if s[0] in props]
    header = list(FIXED_COLS)
    sub = ["", "Material name", "Class name", "Subclass name"]
    colmap = []  # (header, pid, is_values)
    for pid, units, param, pinfo in schema:
        header.append(pid)
        sub.append(units or "")
        colmap.append((pid, False))
        if param:
            header.append(pid + VALUES_SUFFIX)
            sub.append(pinfo)
            colmap.append((pid, True))
    rows = []
    mats = root.findall("Material")
    for i, m in enumerate(mats, start=1):
        bd = m.find("BulkDetails")
        row = [i, bd.findtext("Name") or "", bd.findtext("Class/Name") or "", bd.findtext("Subclass/Name") or ""]
        props = {p.get("property"): p for p in bd.findall("PropertyData")}
        for pid, is_vals in colmap:
            p = props.get(pid)
            if p is None:
                row.append("")
                continue
            d = p.find("ParameterValue/Data") if is_vals else p.find("Data")
            if d is None:
                row.append("")
                continue
            txt = cell_text(d)
            fmt = d.get("format")
            if fmt in ("integer", "exponential") and "," not in txt and txt:
                try:
                    row.append(int(txt) if fmt == "integer" else float(txt))
                    continue
                except ValueError:
                    pass
            row.append(txt)
        rows.append(row)
    write_xlsx(xlsx_path, sub, header, rows, HELP)
    print(f"Exported {len(mats)} materials x {len(header) - len(FIXED_COLS)} property columns -> {xlsx_path}")


# ---------------------------------------------------------------- import

def values_equal(a_text, b_text, fmt):
    if fmt == "string":
        return a_text.strip() == b_text.strip()
    try:
        return parse_numbers(a_text, fmt) == parse_numbers(b_text, fmt)
    except ValueError:
        return False


def do_import(xlsx_path, xml_path):
    root, raw = load_xml(xml_path)
    rows = read_xlsx(xlsx_path)
    if len(rows) < 3:
        sys.exit("Workbook has no data rows.")
    header = [re.sub(r" \([^()]*\)$", "", h.strip()) for h in rows[1]]
    mats = root.findall("Material")
    qm = qualifier_map(root)
    fmt_of, delim_of = {}, {}
    for pd in root.iter("PropertyData"):
        fmt_of.setdefault(pd.get("property"), pd.find("Data").get("format"))
        if pd.get("delimiter"):
            delim_of[pd.get("property")] = pd.get("delimiter")
    has_values = {pd.get("property") for pd in root.iter("PropertyData") if pd.find("ParameterValue") is not None}

    errors, warnings, changes = [], [], []
    ignored = set()
    seen_idx = set()
    for rno, row in enumerate(rows[2:], start=3):
        row = row + [""] * (len(header) - len(row))
        if not any(c.strip() for c in row):
            continue
        try:
            idx = int(float(row[0])) - 1
            mat = mats[idx]
        except (ValueError, IndexError):
            errors.append(f"Row {rno}: invalid or missing '#' ({row[0]!r}); rows must not be inserted.")
            continue
        if idx in seen_idx:
            errors.append(f"Row {rno}: duplicate '#' {idx + 1}.")
            continue
        seen_idx.add(idx)
        bd = mat.find("BulkDetails")
        label = f"Row {rno} ({bd.findtext('Name')})"
        props = {p.get("property"): p for p in bd.findall("PropertyData")}
        pending = []  # (apply-callable, description)

        # Name / Class / Subclass
        for col, path in ((1, "Name"), (2, "Class/Name"), (3, "Subclass/Name")):
            new = row[col].strip()
            el = bd.find(path)
            old = (el.text or "").strip() if el is not None else ""
            if new != old:
                if el is None:
                    if not new:
                        continue
                    if path == "Name":
                        errors.append(f"{label}: Name cannot be blank.")
                        continue
                    pending.append((lambda b_=bd, t=path.split("/")[0], n=new: add_group(b_, t, n), f"{path}: added {new!r}"))
                    continue
                pending.append((lambda e=el, n=new: setattr(e, "text", n), f"{path}: {old!r} -> {new!r}"))

        # Properties
        for c in range(len(FIXED_COLS), len(header)):
            h = header[c]
            if not h:
                continue
            is_vals = h.endswith(VALUES_SUFFIX)
            pid = h[: -len(VALUES_SUFFIX)] if is_vals else h
            if pid not in fmt_of or (is_vals and pid not in has_values):
                if h:
                    ignored.add(h)
                continue
            new = row[c].strip()
            p = props.get(pid)
            if p is None:
                if not new:
                    continue
                if is_vals or pid in has_values:
                    warnings.append(f"{label}: '{h}' cannot be added to a material lacking it (dependent table). Skipped.")
                    continue
                fmt = fmt_of[pid]
                try:
                    core = new if fmt == "string" else format_numbers(parse_numbers(new, fmt), fmt)
                except ValueError:
                    errors.append(f"{label}, {col_letter(c)}{rno} ({pid}): {new!r} is not a valid {fmt}.")
                    continue
                pending.append((lambda b=bd, i=pid, f=fmt, k=core: add_property(b, i, f, k, delim_of.get(i), qm),
                                f"{pid}: added {core!r}"))
                continue
            d = p.find("ParameterValue/Data") if is_vals else p.find("Data")
            fmt = d.get("format")
            if values_equal(cell_text(d), new, fmt):
                continue
            try:
                core = new if fmt == "string" else format_numbers(parse_numbers(new, fmt), fmt)
            except ValueError:
                errors.append(f"{label}, {col_letter(c)}{rno} ({h}): {new!r} is not a valid {fmt}.")
                continue
            pending.append((lambda dd=d, k=core, pp=p, ii=pid, ff=fmt, isv=is_vals: update_data(dd, k, pp, ii, ff, isv, qm, warnings, label),
                            f"{h}: {cell_text(d)[:40]!r} -> {core[:40]!r}"))
        changes.append((mat, label, pending))

    if errors:
        print("Import aborted - nothing was written. Fix these cells:")
        for e in errors:
            print("  " + e)
        sys.exit(1)

    total = sum(len(p) for _, _, p in changes)
    for mat, label, pending in changes:
        for fn, desc in pending:
            fn()
            print(f"  {label}: {desc}")
    for mat in mats:
        for p in mat.iter("PropertyData"):
            pv = p.find("ParameterValue")
            if pv is not None:
                n1 = len([x for x in (p.find("Data").text or "").split(",") if x.strip()])
                n2 = len([x for x in (pv.find("Data").text or "").split(",") if x.strip()])
                if n1 != n2:
                    warnings.append(f"{mat.findtext('BulkDetails/Name')}: {p.get('property')} has {n1} temperatures but {n2} values.")
    if ignored:
        print("Ignored non-property columns: " + ", ".join(sorted(ignored)))
    for w in warnings:
        print("WARNING: " + w)
    if total == 0:
        print("No changes found; XML left untouched.")
        return
    backup_dir = Path(xml_path).parent / "backup"
    backup_dir.mkdir(exist_ok=True)
    backup = backup_dir / f"{Path(xml_path).stem}_{datetime.datetime.now():%Y%m%d_%H%M%S}.xml"
    shutil.copy2(xml_path, backup)
    Path(xml_path).write_bytes(serialize(root, raw))
    print(f"Applied {total} change(s) to {xml_path}\nBackup: {backup}")


def update_data(d, core, p, pid, fmt, is_vals, qm, warnings, label):
    set_data_text(d, core)
    if fmt == "integer" and not is_vals:
        q = p.find("Qualifier")
        if q is not None:
            new_q = qm.get((pid, core))
            if new_q is not None:
                q.text = new_q
            else:
                warnings.append(f"{label}: {pid}={core} has no known description; old Qualifier {q.text!r} kept.")


def add_group(bd, tag, name):
    """Add <Class>/<Subclass> with a <Name> child, placed after Name (Class) or after Class (Subclass)."""
    after = bd.find("Class") if tag == "Subclass" and bd.find("Class") is not None else bd.find("Name")
    g = ET.Element(tag)
    g.text = "\n" + " " * 16
    g.tail = "\n" + " " * 12
    n = ET.SubElement(g, "Name")
    n.text = name
    n.tail = "\n" + " " * 12
    bd.insert(list(bd).index(after) + 1, g)


def add_property(bd, pid, fmt, core, delim, qm):
    prev = list(bd)[-1]
    prev.tail = "\n" + " " * 12
    p = ET.SubElement(bd, "PropertyData", {"property": pid})
    if delim:
        p.set("delimiter", delim)
    p.text = "\n" + " " * 16
    p.tail = "\n" + " " * 8
    d = ET.SubElement(p, "Data", {"format": fmt})
    d.text = DATA_INDENT + core
    d.tail = "\n" + " " * 16
    if fmt == "integer" and (pid, core) in qm:
        q = ET.SubElement(p, "Qualifier")
        q.text = qm[(pid, core)]
        q.tail = "\n" + " " * 16
    list(p)[-1].tail = "\n" + " " * 12


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    e = sub.add_parser("export")
    e.add_argument("xml", nargs="?", default=str(DEFAULT_XML))
    e.add_argument("xlsx", nargs="?", default=str(DEFAULT_XLSX))
    e.add_argument("--props", default=",".join(DEFAULT_PROPS), help="comma-separated property ids to export (default: %(default)s)")
    e.add_argument("--all", action="store_true", help="export every property in the library")
    i = sub.add_parser("import")
    i.add_argument("xlsx", nargs="?", default=str(DEFAULT_XLSX))
    i.add_argument("xml", nargs="?", default=str(DEFAULT_XML))
    a = ap.parse_args()
    if a.cmd == "export":
        do_export(a.xml, a.xlsx, None if a.all else [p.strip() for p in a.props.split(",") if p.strip()])
    else:
        do_import(a.xlsx, a.xml)


if __name__ == "__main__":
    main()
