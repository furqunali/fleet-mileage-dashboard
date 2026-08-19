#!/usr/bin/env python3
r"""
protect_workbook_v7.py — lock the reference tabs of the v7 workbook so HR can
only fill inputs / make dropdown selections, never alter the reference data.

Dashboard, Monthly Plan and Directory are ALREADY sheet-protected (input cells
unlocked / formulas locked). This adds the same protection to the three tabs
that were still open:
    Matrix        (pure reference distance grid)
    Map of Site   (map image + note)
    How To Use    (instructions)

Sheets are identified BY CONTENT (not filename), because an Excel re-save can
renumber the worksheet files. Surgical zip edit: only inserts
`<sheetProtection sheet="1" objects="1" scenarios="1"/>` right after each
</sheetData>; every other byte (slicer/pivotCache/data-model/table/image) is
preserved. No password (matches the existing protected tabs) — HR can still do
Review > Unprotect if they ever truly need to; this just stops accidental edits.

Base  : Latest Version (v7)/...v7.xlsx  (the CLEAN delivered template)
Output: <temp>  (caller COM-verifies, then deploys)
"""
import zipfile, os, re

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(HERE, "Latest Version (v7)", "Fleet_Distance_Dashboard_v7.xlsx")
TMP  = os.path.join(HERE, "_v7_protected.tmp.xlsx")

PROT = '<sheetProtection sheet="1" objects="1" scenarios="1"/>'

def main():
    zin = zipfile.ZipFile(BASE, "r")
    names = zin.namelist()
    parts = {n: zin.read(n) for n in names}
    infos = {i.filename: i for i in zin.infolist()}
    zin.close()

    changed = []
    for n in names:
        if not re.match(r"xl/worksheets/sheet\d+\.xml$", n):
            continue
        xml = parts[n].decode("utf-8")
        if "<sheetProtection" in xml:
            continue                                   # already protected
        assert "</sheetData>" in xml, f"{n}: no </sheetData>"
        # identify for the log
        if 'ref="A1:AF' in xml:      label = "Matrix"
        elif "<drawing " in xml:     label = "Map of Site"
        else:                        label = "How To Use / other reference"
        xml = xml.replace("</sheetData>", "</sheetData>" + PROT, 1)
        parts[n] = xml.encode("utf-8")
        changed.append((n, label))

    zout = zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED)
    for n in names:
        zout.writestr(infos[n], parts[n]) if n in infos else zout.writestr(n, parts[n])
    zout.close()
    print("Protected tabs:", changed)
    print("WROTE TEMP:", TMP)

if __name__ == "__main__":
    main()
