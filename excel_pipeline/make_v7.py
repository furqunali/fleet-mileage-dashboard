#!/usr/bin/env python3
r"""
make_v7.py — mint Fleet_Distance_Dashboard_v7.xlsx from the restyled v6.

v7 is a label-only bump of v6 (which already carries the dark Monthly Plan
restyle). No cell data / formulas / layout change — only visible version strings:
  * Dashboard subtitle  (sharedStrings si 160): "Version v6"  -> "Version v7"
  * Map-of-Site footer  (sharedStrings si 202): "Version v4 ... 21 Jul 2026"
                                                -> "Version v7 ... 28 Jul 2026"
  * Monthly Plan subtitle (sheet6.xml inline) : "Version v6"  -> "Version v7"

Surgical zip edit so slicer/pivotCache/data-model/Table1/image are preserved.
v6 is left intact as the prior snapshot.
"""
import zipfile, os

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "Fleet_Distance_Dashboard_v6.xlsx")
DST = os.path.join(HERE, "Fleet_Distance_Dashboard_v7.xlsx")

def main():
    zin = zipfile.ZipFile(SRC, "r")
    names = zin.namelist()
    parts = {n: zin.read(n) for n in names}
    infos = {i.filename: i for i in zin.infolist()}
    zin.close()

    ss = parts["xl/sharedStrings.xml"].decode("utf-8")
    before = ss
    ss = ss.replace("Version v6", "Version v7")                  # si 160 (Dashboard)
    ss = ss.replace("Version v4", "Version v7")                  # si 202 (Map tab)
    ss = ss.replace("Data as of 21 Jul 2026", "Data as of 28 Jul 2026")
    assert ss != before, "sharedStrings: no version string replaced"
    parts["xl/sharedStrings.xml"] = ss.encode("utf-8")

    s6 = parts["xl/worksheets/sheet6.xml"].decode("utf-8")
    assert "Version v6" in s6, "sheet6: 'Version v6' not found"
    parts["xl/worksheets/sheet6.xml"] = s6.replace("Version v6", "Version v7").encode("utf-8")

    tmp = DST + ".tmp"
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    for n in names:
        zout.writestr(infos[n], parts[n]) if n in infos else zout.writestr(n, parts[n])
    zout.close()
    os.replace(tmp, DST)
    print("WROTE", DST)

if __name__ == "__main__":
    main()
