"""
polish_workbook.py  (P1-3, P1-4, P1-5)
--------------------------------------
Builds a production-polished workbook from v3 via surgical XML edits (so the
slicer / pivot / data model / image / formulas / cached values are preserved):

  P1-3  Sheet protection on Dashboard & Directory (the only sheets with
        formulas). Formula cells are locked; INPUT cells are unlocked so normal
        use is unchanged:
          Dashboard : B17, F17 (calculator dropdowns), C27:C34 (route stops)
          Directory : B5:E33 (code/name/address/type), K5:L33 (lat/lon)
        No password — adding a new site (copying formulas down) just needs
        Review > Unprotect Sheet first. Matrix (source of truth) and Map of Site
        (picture + slicer) are left fully editable/interactive on purpose.

  P1-4  How To Use: add a "Web dashboard & coordinates" note; extend the
        add-a-site named-range list with dir_lat / dir_lon.

  P1-5  Visible version/date: Dashboard subtitle + a How To Use VERSION line.

Input : Fleet_Distance_Dashboard_v3.xlsx
Output: Fleet_Distance_Dashboard_v4.xlsx
No data or formulas are changed.
"""
import os
import re
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "Fleet_Distance_Dashboard_v3.xlsx")
OUT = os.path.join(HERE, "Fleet_Distance_Dashboard_v4.xlsx")

VERSION_LINE = "Version v4  ·  Data as of 21 Jul 2026  ·  Source of truth = Matrix + Directory tabs."


# ---- style (cellXfs) twinning: create unlocked copies on demand ----
class Styles:
    def __init__(self, styles_xml):
        self.xml = styles_xml
        m = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', styles_xml, re.S)
        self.pre, body, self.count = styles_xml[:m.start()], m.group(2), int(m.group(1))
        self.post = styles_xml[m.end():]
        self.xfs = re.findall(r'<xf\b[^>]*?/>|<xf\b[^>]*?>.*?</xf>', body, re.S)
        assert len(self.xfs) == self.count, (len(self.xfs), self.count)
        self.twins = {}   # old index -> new index

    def unlocked_twin(self, old):
        old = int(old)
        if old in self.twins:
            return self.twins[old]
        xf = self.xfs[old]
        # normalise: drop any existing applyProtection attr and <protection/> child
        if xf.endswith("/>"):
            attrs = xf[3:-2]
            children = ""
        else:
            mm = re.match(r'<xf\b(.*?)>(.*)</xf>', xf, re.S)
            attrs, children = mm.group(1), mm.group(2)
        attrs = re.sub(r'\s*applyProtection="[^"]*"', '', attrs)
        children = re.sub(r'<protection[^>]*/>', '', children)
        new_xf = '<xf' + attrs + ' applyProtection="1">' + children + '<protection locked="0"/></xf>'
        new_idx = self.count
        self.xfs.append(new_xf)
        self.count += 1
        self.twins[old] = new_idx
        return new_idx

    def serialize(self):
        body = ''.join(self.xfs)
        return self.pre + '<cellXfs count="%d">' % self.count + body + '</cellXfs>' + self.post


def remap_cell_style(xml, ref, styles):
    """Point one cell's style at an unlocked twin. Returns new xml."""
    m = re.search(r'<c r="%s"(?:\s+s="(\d+)")?' % re.escape(ref), xml)
    if not m:
        raise SystemExit("cell %s not found" % ref)
    old = m.group(1) if m.group(1) is not None else "0"
    new = styles.unlocked_twin(old)
    if m.group(1) is None:
        repl = '<c r="%s" s="%d"' % (ref, new)
    else:
        repl = '<c r="%s" s="%d"' % (ref, new)
    return xml[:m.start()] + repl + xml[m.end():]


def add_protection(xml):
    prot = '<sheetProtection sheet="1" objects="1" scenarios="1"/>'
    return xml.replace('</sheetData>', '</sheetData>' + prot, 1)


def write_cell_text(xml, ref, text):
    """Fill an existing empty <c r=REF s=.. /> with an inline string."""
    esc = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    def repl(m):
        return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, m.group(1), esc)
    new, n = re.subn(r'<c r="%s"( s="\d+")?\s*/>' % re.escape(ref), repl, xml, count=1)
    if n != 1:
        raise SystemExit("empty cell %s not found for text write" % ref)
    return new


def main():
    zin = zipfile.ZipFile(SRC, "r")
    changed = {}

    styles = Styles(zin.read("xl/styles.xml").decode("utf-8"))

    # ---- Dashboard (sheet1): unlock inputs, protect ----
    d = zin.read("xl/worksheets/sheet1.xml").decode("utf-8")
    dash_inputs = ["B17", "F17"] + ["C%d" % r for r in range(27, 35)]
    for ref in dash_inputs:
        d = remap_cell_style(d, ref, styles)
    d = add_protection(d)
    changed["xl/worksheets/sheet1.xml"] = d.encode("utf-8")

    # ---- Directory (sheet3): unlock inputs, protect ----
    dr = zin.read("xl/worksheets/sheet3.xml").decode("utf-8")
    dir_inputs = [c + str(r) for r in range(5, 34) for c in ("B", "C", "D", "E", "K", "L")]
    for ref in dir_inputs:
        dr = remap_cell_style(dr, ref, styles)
    dr = add_protection(dr)
    changed["xl/worksheets/sheet3.xml"] = dr.encode("utf-8")

    # ---- How To Use (sheet5): provenance / web notes into empty cells ----
    h = zin.read("xl/worksheets/sheet5.xml").decode("utf-8")
    notes = {
        "B24": "WEB DASHBOARD (HTML) & COORDINATES",
        "B25": "The HTML dashboard is GENERATED from this workbook (Matrix = distances; Directory = names/addresses/types + LAT/LON in columns K/L). Do not edit the HTML by hand.",
        "B26": "Regenerate after edits: run build_html_from_workbook.py (see README.md). It rebuilds Fleet_Distance_Calculator_Generate.html; run it with --verify to confirm the HTML matches this workbook.",
        "B27": "The web map draws real road routes via OSRM with a straight-line fallback. Recorded Matrix miles remain authoritative; the routed distance is a visual estimate. Some pins are ZIP/city-level approximate.",
        "B28": "Coordinates were geocoded once via geocode_sites.py into Directory K/L. To refine a pin, edit K/L (or coord_overrides.json) and regenerate.",
        "B29": "Protection: Dashboard & Directory are protected (no password) so formulas can't be overwritten. To add a new site, use Review > Unprotect Sheet first, follow the 5 steps above, then re-protect.",
        "B31": VERSION_LINE,
    }
    for ref, txt in notes.items():
        h = write_cell_text(h, ref, txt)
    changed["xl/worksheets/sheet5.xml"] = h.encode("utf-8")

    # ---- sharedStrings: extend named-range list + Dashboard subtitle version ----
    ss = zin.read("xl/sharedStrings.xml").decode("utf-8")
    ss2 = ss.replace("dir_miles, dir_label.", "dir_miles, dir_label, dir_lat, dir_lon.", 1)
    if ss2 == ss:
        raise SystemExit("named-range string not found")
    ss3 = ss2.replace("pull live from the Matrix tab.",
                      "pull live from the Matrix tab.  " + VERSION_LINE, 1)
    if ss3 == ss2:
        raise SystemExit("Dashboard subtitle string not found")
    changed["xl/sharedStrings.xml"] = ss3.encode("utf-8")

    # ---- styles.xml with appended unlocked twins ----
    changed["xl/styles.xml"] = styles.serialize().encode("utf-8")

    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info, changed.get(info.filename, zin.read(info.filename)))
    zin.close()

    print("Wrote", os.path.basename(OUT))
    print("  unlocked-twin styles created:", len(styles.twins), "-> cellXfs now", styles.count)
    print("  protected: Dashboard, Directory | inputs unlocked | Matrix & Map left editable")


if __name__ == "__main__":
    main()
