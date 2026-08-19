"""
patch_workbook.py
-----------------
Produces a corrected, versioned copy of the Fleet distance workbook WITHOUT
re-serializing it through a spreadsheet library (which would drop the slicer,
table, pivot cache and cached formula values). It edits the raw XML inside the
.xlsx zip so everything else is preserved byte-for-byte.

Two changes only:
  1. Data fix  : Matrix!E27  (site 0077 -> site 0017)
                 197 -> 184. The 197 was a copy error inherited from an earlier
                 working file (it duplicated the 0075 row value and
                 broke proximity ordering); 184 matches the symmetric
                 0017->0077 cell and neighbouring sites.
  2. Cleanup   : replace non-breaking spaces (U+00A0) with normal spaces in
                 shared strings and the slicer cache (Democity address).

Input : Fleet_Distance_Dashboard.xlsx
Output: Fleet_Distance_Dashboard_v2.xlsx
"""
import os
import re
import shutil
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "Fleet_Distance_Dashboard.xlsx")
OUT = os.path.join(HERE, "Fleet_Distance_Dashboard_v2.xlsx")

MATRIX_SHEET = "xl/worksheets/sheet2.xml"   # <sheet name="Matrix" .../> -> rId2 -> sheet2.xml
NBSP = b"\xc2\xa0"                            # UTF-8 encoding of U+00A0


def patch():
    if not os.path.exists(SRC):
        raise SystemExit("Source workbook not found: " + SRC)

    zin = zipfile.ZipFile(SRC, "r")
    members = zin.namelist()

    changed = {}

    # --- 1. Matrix!E27: 197 -> 184 ---
    matrix_xml = zin.read(MATRIX_SHEET).decode("utf-8")
    new_matrix, n = re.subn(
        r'(<c r="E27"[^>]*><v>)197(</v></c>)',
        r"\g<1>184\g<2>",
        matrix_xml,
        count=1,
    )
    if n != 1:
        zin.close()
        raise SystemExit("ERROR: expected exactly one Matrix!E27=197 cell, found %d" % n)
    changed[MATRIX_SHEET] = new_matrix.encode("utf-8")

    # --- 2. Non-breaking spaces -> normal spaces ---
    for name in ("xl/sharedStrings.xml", "xl/slicerCaches/slicerCache1.xml"):
        if name in members:
            raw = zin.read(name)
            if NBSP in raw:
                changed[name] = raw.replace(NBSP, b" ")

    # --- write the new zip, copying every member, substituting the changed ones ---
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = changed.get(info.filename, zin.read(info.filename))
            # preserve original metadata (name, date, compression flags)
            zout.writestr(info, data)
    zin.close()

    print("Wrote", os.path.basename(OUT))
    print("  Matrix!E27: 197 -> 184  (cells changed: %d)" % 1)
    print("  NBSP cleaned in:", ", ".join(sorted(changed)) )


if __name__ == "__main__":
    patch()
