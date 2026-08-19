"""
visual_polish_workbook.py  (Excel visual/formatting pass, items 1-7)
--------------------------------------------------------------------
Builds a visually polished workbook from v4 via surgical XML edits. Preserves
ALL data, formulas, cached values, coordinates, sheet protection, the slicer,
pivot cache, Power-Pivot data model, the embedded map image, and named ranges.

  1. Number formats  : numFmt 164 "0.#" -> "General" (kills trailing-dot 29./173.
                       across KPIs, calculator, route legs, Directory miles,
                       Matrix grid). LAT/LON (Directory K/L) -> 0.00000.
  2. Print-ready     : landscape + fit-to-1-wide + print areas + repeating
                       titles on Dashboard, Matrix, Directory.
  3. Frozen panes    : freeze Directory header row 4 (Matrix already frozen);
                       reset Dashboard/Directory to open at the top.
  4. Branding        : corporate brand-yellow band across the Dashboard top row
                       with the company wordmark (styled cells, no image asset).
  5. Tab colors      : color the "Map of Site" tab (green) for consistency.
  6. Typo            : "Suit 904" -> "Suite 904" (head-office address).
  7. Matrix HQ hi-lite: bold the HQ row (row 2) + HQ column (col B) distances and
                       highlight the HQ header cells (A1/A2/B1) in brand yellow.

Input : Fleet_Distance_Dashboard_v4.xlsx
Output: Fleet_Distance_Dashboard_v5.xlsx
"""
import os, re, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "Fleet_Distance_Dashboard_v4.xlsx")
OUT = os.path.join(HERE, "Fleet_Distance_Dashboard_v5.xlsx")

GOLD = "FFFFC72C"; INK = "FF10151C"; MAPGREEN = "FF57C785"
BRAND_TEXT = "DEMO FLEET CO     ·     FLEET DISTANCE OPERATIONS"

def col_letters(n):
    s=""
    while n: n,r=divmod(n-1,26); s=chr(65+r)+s
    return s

# ---------------- styles.xml manager ----------------
class Styles:
    def __init__(self, xml):
        self.xml = xml
        self.fonts, self.fo_open = self._section("fonts")
        self.fills, self.fi_open = self._section("fills")
        self.borders, self.bo_open = self._section("borders")
        self.xfs, self.xf_open = self._cellxfs()

    def _section(self, tag):
        m = re.search(r'(<%s[^>]*>)(.*?)</%s>' % (tag, tag), self.xml, re.S)
        item = {"fonts":r'<font>.*?</font>|<font/>',
                "fills":r'<fill>.*?</fill>',
                "borders":r'<border[^>]*>.*?</border>|<border[^>]*/>'}[tag]
        return re.findall(item, m.group(2), re.S), m.group(1)

    def _cellxfs(self):
        m = re.search(r'(<cellXfs[^>]*>)(.*?)</cellXfs>', self.xml, re.S)
        return re.findall(r'<xf\b[^>]*?/>|<xf\b[^>]*?>.*?</xf>', m.group(2), re.S), m.group(1)

    # --- adders ---
    def add_font(self, s): self.fonts.append(s); return len(self.fonts)-1
    def add_fill(self, s): self.fills.append(s); return len(self.fills)-1
    def add_border(self, s): self.borders.append(s); return len(self.borders)-1
    def add_xf(self, s): self.xfs.append(s); return len(self.xfs)-1

    def font_id_of(self, xf):
        m = re.search(r'fontId="(\d+)"', self.xfs[int(xf)]); return int(m.group(1)) if m else 0
    def border_id_of(self, xf):
        m = re.search(r'borderId="(\d+)"', self.xfs[int(xf)]); return int(m.group(1)) if m else 0

    def bold_font_from(self, font_xml):
        return font_xml if "<b/>" in font_xml else font_xml.replace("<font>", "<font><b/>", 1)

    # copy an xf, override attrs, keep children
    def variant(self, base_idx, **over):
        xf = self.xfs[int(base_idx)]
        if xf.endswith("/>"):
            head, kids = xf[3:-2], ""
        else:
            mm = re.match(r'<xf\b(.*?)>(.*)</xf>', xf, re.S); head, kids = mm.group(1), mm.group(2)
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', head))
        attrs.update({k: str(v) for k, v in over.items()})
        a = " ".join('%s="%s"' % (k, attrs[k]) for k in attrs)
        return ("<xf %s/>" % a) if not kids else ("<xf %s>%s</xf>" % (a, kids))

    def serialize(self):
        out = self.xml
        out = re.sub(r'<fonts[^>]*>.*?</fonts>',
                     re.sub(r'count="\d+"', 'count="%d"' % len(self.fonts), self.fo_open) + "".join(self.fonts) + "</fonts>",
                     out, count=1, flags=re.S)
        out = re.sub(r'<fills[^>]*>.*?</fills>',
                     re.sub(r'count="\d+"', 'count="%d"' % len(self.fills), self.fi_open) + "".join(self.fills) + "</fills>",
                     out, count=1, flags=re.S)
        out = re.sub(r'<borders[^>]*>.*?</borders>',
                     re.sub(r'count="\d+"', 'count="%d"' % len(self.borders), self.bo_open) + "".join(self.borders) + "</borders>",
                     out, count=1, flags=re.S)
        out = re.sub(r'<cellXfs[^>]*>.*?</cellXfs>',
                     re.sub(r'count="\d+"', 'count="%d"' % len(self.xfs), self.xf_open) + "".join(self.xfs) + "</cellXfs>",
                     out, count=1, flags=re.S)
        return out


def set_cell_style(xml, ref, new_s):
    m = re.search(r'<c r="%s"(?:\s+s="\d+")?' % re.escape(ref), xml)
    if not m: raise SystemExit("cell %s not found" % ref)
    return xml[:m.start()] + ('<c r="%s" s="%d"' % (ref, new_s)) + xml[m.end():]

def cur_style(xml, ref):
    m = re.search(r'<c r="%s"(?:\s+s="(\d+)")?' % re.escape(ref), xml)
    return m.group(1) if (m and m.group(1)) else "0"

def set_cell_text(xml, ref, new_s, text):
    esc = text.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
    m = re.search(r'<c r="%s"(?:\s+s="\d+")?\s*/>' % re.escape(ref), xml)
    if not m: raise SystemExit("empty cell %s not found" % ref)
    return xml[:m.start()] + ('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, new_s, esc)) + xml[m.end():]


def main():
    zin = zipfile.ZipFile(SRC, "r")
    changed = {}
    st = Styles(zin.read("xl/styles.xml").decode("utf-8"))

    # ---- item 1: numFmt 164 -> General; add 0.00000 (165) ----
    st.xml = st.xml.replace('<numFmt numFmtId="164" formatCode="0.#"/>',
                            '<numFmt numFmtId="164" formatCode="General"/>', 1)
    if '<numFmts' in st.xml:
        st.xml = re.sub(r'(<numFmts count=")(\d+)(">)',
                        lambda m: m.group(1)+str(int(m.group(2))+1)+m.group(3)+'<numFmt numFmtId="165" formatCode="0.00000"/>',
                        st.xml, count=1)
    # reload sections after editing numFmts/xml (fonts/fills/borders/xfs lists already parsed; keep)
    # NOTE: st.serialize rebuilds from lists + st.xml, so numFmts edit persists via st.xml.

    # shared new style objects
    gold_fill = st.add_fill('<fill><patternFill patternType="solid"><fgColor rgb="%s"/><bgColor indexed="64"/></patternFill></fill>' % GOLD)
    brand_font = st.add_font('<font><b/><sz val="11"/><color rgb="%s"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>' % INK)
    grid_font = st.font_id_of(19)
    bold_grid_font = st.add_font(st.bold_font_from(st.fonts[grid_font]))

    # ---- item 1 (cont): LAT/LON 0.00000, remap K5:L33 ----
    latlon_xf = st.add_xf(st.variant(95, numFmtId=165, applyNumberFormat="1"))

    # ---- item 4: branding xfs ----
    brand_text_xf = st.add_xf('<xf numFmtId="0" fontId="%d" fillId="%d" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' % (brand_font, gold_fill))
    brand_fill_xf = st.add_xf('<xf numFmtId="0" fontId="0" fillId="%d" borderId="0" xfId="0" applyFill="1"/>' % gold_fill)

    # ---- item 7: Matrix HQ highlight xfs ----
    hdr_hi = {}   # base style -> gold-fill+brand-font twin (for A1/A2/B1)
    for base in ("15", "16", "18"):
        hdr_hi[base] = st.add_xf(st.variant(base, fillId=gold_fill, fontId=brand_font, applyFill="1", applyFont="1"))
    bold_twin = {}  # base style -> bold-font twin (HQ row/col distances)
    for base in ("19", "20", "22"):
        bold_twin[base] = st.add_xf(st.variant(base, fontId=bold_grid_font, applyFont="1"))

    changed["xl/styles.xml"] = None  # placeholder; serialized at end

    # ================= sheet edits =================
    # ---- Dashboard (sheet1): branding band, scroll reset, print ----
    d = zin.read("xl/worksheets/sheet1.xml").decode("utf-8")
    d = set_cell_text(d, "B1", brand_text_xf, BRAND_TEXT)
    for c in range(3, 10):  # C1..I1 fill band
        d = set_cell_style(d, col_letters(c) + "1", brand_fill_xf)
    d = d.replace('<row r="1" spans="1:14" ht="15" customHeight="1"',
                  '<row r="1" spans="1:14" ht="22" customHeight="1"', 1)
    d = d.replace('<sheetView showGridLines="0" topLeftCell="A10" zoomScale="115" zoomScaleNormal="115" workbookViewId="0"><selection activeCell="B10" sqref="B10:C10"/></sheetView>',
                  '<sheetView showGridLines="0" zoomScale="115" zoomScaleNormal="115" workbookViewId="0"><selection activeCell="B2" sqref="B2"/></sheetView>', 1)
    d = d.replace('</sheetPr>', '<pageSetUpPr fitToPage="1"/></sheetPr>', 1)
    # Dashboard is a one-screen dashboard -> fit to a single page (width AND height)
    d = re.sub(r'<pageSetup[^>]*/>', '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="1" horizontalDpi="300" verticalDpi="300"/>', d, count=1)
    d = re.sub(r'<pageMargins[^>]*/>', '<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>', d, count=1)
    changed["xl/worksheets/sheet1.xml"] = d.encode("utf-8")

    # ---- Matrix (sheet2): HQ highlight + print (already frozen) ----
    mx = zin.read("xl/worksheets/sheet2.xml").decode("utf-8")
    for ref in ("A1", "A2", "B1"):
        mx = set_cell_style(mx, ref, hdr_hi[cur_style(mx, ref)])
    hq_cells = ["%s2" % col_letters(c) for c in range(2, 31)] + ["B%d" % r for r in range(2, 31)]
    for ref in dict.fromkeys(hq_cells):
        base = cur_style(mx, ref)
        if base in bold_twin:
            mx = set_cell_style(mx, ref, bold_twin[base])
    mx = mx.replace('</sheetPr>', '<pageSetUpPr fitToPage="1"/></sheetPr>', 1)
    mx = re.sub(r'<pageSetup[^>]*/>', '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0" horizontalDpi="300" verticalDpi="300"/>', mx, count=1)
    mx = re.sub(r'<pageMargins[^>]*/>', '<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>', mx, count=1)
    changed["xl/worksheets/sheet2.xml"] = mx.encode("utf-8")

    # ---- Directory (sheet3): freeze header, print ----
    dr = zin.read("xl/worksheets/sheet3.xml").decode("utf-8")
    for r in range(5, 34):
        dr = set_cell_style(dr, "K%d" % r, latlon_xf)
        dr = set_cell_style(dr, "L%d" % r, latlon_xf)
    dr = dr.replace('<sheetView showGridLines="0" topLeftCell="A16" zoomScaleNormal="100" workbookViewId="0"><selection activeCell="B1" sqref="B1:H35"/></sheetView>',
                    '<sheetView showGridLines="0" zoomScaleNormal="100" workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="B5" sqref="B5"/></sheetView>', 1)
    dr = dr.replace('</sheetPr>', '<pageSetUpPr fitToPage="1"/></sheetPr>', 1)
    if '<pageSetup' in dr:
        dr = re.sub(r'<pageSetup[^>]*/>', '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0" horizontalDpi="300" verticalDpi="300"/>', dr, count=1)
    else:
        dr = dr.replace('<pageMargins', '<pageMargins', 1)  # keep; add pageSetup after margins
        dr = re.sub(r'(<pageMargins[^>]*/>)', r'\1<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0" horizontalDpi="300" verticalDpi="300"/>', dr, count=1)
    dr = re.sub(r'<pageMargins[^>]*/>', '<pageMargins left="0.3" right="0.3" top="0.4" bottom="0.4" header="0.2" footer="0.2"/>', dr, count=1)
    changed["xl/worksheets/sheet3.xml"] = dr.encode("utf-8")

    # ---- Map of Site (sheet4): tab color ----
    s4 = zin.read("xl/worksheets/sheet4.xml").decode("utf-8")
    s4 = re.sub(r'(<worksheet\b[^>]*>)', r'\1<sheetPr><tabColor rgb="%s"/></sheetPr>' % MAPGREEN, s4, count=1)
    changed["xl/worksheets/sheet4.xml"] = s4.encode("utf-8")

    # ---- item 6: typo in sharedStrings ----
    ss = zin.read("xl/sharedStrings.xml").decode("utf-8")
    ss2 = ss.replace("Suit 904", "Suite 904", 1)
    if ss2 == ss: raise SystemExit("'Suit 904' not found")
    changed["xl/sharedStrings.xml"] = ss2.encode("utf-8")

    # ---- item 2: print areas + titles in workbook.xml ----
    wbx = zin.read("xl/workbook.xml").decode("utf-8")
    names = ('<definedName name="_xlnm.Print_Area" localSheetId="0">Dashboard!$B$1:$I$35</definedName>'
             '<definedName name="_xlnm.Print_Area" localSheetId="1">Matrix!$A$1:$AD$30</definedName>'
             '<definedName name="_xlnm.Print_Titles" localSheetId="1">Matrix!$A:$A,Matrix!$1:$1</definedName>'
             '<definedName name="_xlnm.Print_Area" localSheetId="2">Directory!$B$1:$F$33</definedName>'
             '<definedName name="_xlnm.Print_Titles" localSheetId="2">Directory!$4:$4</definedName>')
    wbx2 = wbx.replace('<definedNames>', '<definedNames>' + names, 1)
    if wbx2 == wbx: raise SystemExit("definedNames not found")
    changed["xl/workbook.xml"] = wbx2.encode("utf-8")

    # ---- finalize styles ----
    changed["xl/styles.xml"] = st.serialize().encode("utf-8")

    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info, changed.get(info.filename) if changed.get(info.filename) is not None else zin.read(info.filename))
    zin.close()
    print("Wrote", os.path.basename(OUT))
    print("  fonts:%d fills:%d borders:%d cellXfs:%d" % (len(st.fonts), len(st.fills), len(st.borders), len(st.xfs)))


if __name__ == "__main__":
    main()
