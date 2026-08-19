#!/usr/bin/env python3
r"""
build_v9_buttons.py  ->  Fleet_Distance_Dashboard_v9.xlsm

Adds the HR-web-style action buttons to the Excel "Monthly Plan" tab, via Excel
COM (so the slicer / pivotCache / data-model / Table1 / image are all preserved,
which an openpyxl re-save would drop). Source = the clean protected v8 .xlsx.

Buttons (top-right strip, OUTSIDE the B:G print area so they never print):
  * Add 5 trip rows   -> AddTripRows    (grows the trip log 16..55, SUM auto-expands)
  * Add 5 extra rows  -> AddExtraRows    (grows the temp-entries block 62..71)
  * Fill Down (Ctrl+D)-> FillDownSelection (guarded: input cells only)
  * Print             -> PrintPlan       (opens the print dialog)

Each macro unprotects -> edits -> re-protects with the SAME flags
(sheet/objects/scenarios), so the workbook stays locked for HR.

Prereqs (already handled this session): AccessVBOM=1, Excel fully closed.
Run with the Python that has pywin32:
  python build_v9_buttons.py
"""
import os, sys, zipfile, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
SRC  = os.path.join(HERE, "Fleet_Distance_Dashboard_v8.xlsx")
OUT  = os.path.join(HERE, "Fleet_Distance_Dashboard_v9.xlsm")
WORK = os.path.join(HERE, "_v9_build_unlocked.xlsx")   # temp: v8 with sheet6 protection stripped


def make_unlocked_copy(src, dst):
    """Copy the workbook byte-for-byte but drop <sheetProtection.../> from the
    Monthly Plan sheet (sheet6.xml). Lets COM add buttons without the flaky
    Worksheet.Unprotect call; we re-Protect via COM at the end."""
    zin = zipfile.ZipFile(src, "r")
    names = zin.namelist()
    infos = {i.filename: i for i in zin.infolist()}
    parts = {n: zin.read(n) for n in names}
    zin.close()
    s6 = parts["xl/worksheets/sheet6.xml"].decode("utf-8")
    import re
    new = re.sub(r'<sheetProtection[^>]*/>', '', s6, count=1)
    if new == s6:
        raise SystemExit("sheetProtection not found in sheet6.xml")
    parts["xl/worksheets/sheet6.xml"] = new.encode("utf-8")
    if os.path.exists(dst):
        os.remove(dst)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for n in names:
        zout.writestr(infos[n], parts[n])
    zout.close()


def reprotect_sheet(xlsm_path, sheet_name, protection):
    """Re-insert <sheetProtection.../> into the given sheet's XML inside a saved
    .xlsm. COM's Worksheet.Protect throws 0x800A03EC on this particular workbook,
    so we do it at the zip/XML level. Placed right before <mergeCells> (schema
    order: sheetProtection precedes mergeCells)."""
    import re
    z = zipfile.ZipFile(xlsm_path, "r")
    names = z.namelist()
    infos = {i.filename: i for i in z.infolist()}
    parts = {n: z.read(n) for n in names}
    z.close()

    wbxml = parts["xl/workbook.xml"].decode("utf-8")
    m = re.search(r'<sheet\b[^>]*\bname="%s"[^>]*\br:id="([^"]+)"'
                  % re.escape(sheet_name), wbxml)
    if not m:
        raise SystemExit("reprotect: sheet '%s' not found in workbook.xml" % sheet_name)
    rid = m.group(1)
    rels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rm = re.search(r'<Relationship\b[^>]*\bId="%s"[^>]*\bTarget="([^"]+)"'
                   % re.escape(rid), rels)
    if not rm:
        raise SystemExit("reprotect: rel '%s' not found" % rid)
    target = rm.group(1)
    part = target[1:] if target.startswith("/") else "xl/" + target
    if part not in parts:
        raise SystemExit("reprotect: part '%s' missing" % part)

    sx = parts[part].decode("utf-8")
    if "<sheetProtection" in sx:
        return part  # already protected
    for anchor in ("<mergeCells", "<dataValidations", "<pageMargins", "<pageSetup"):
        if anchor in sx:
            sx = sx.replace(anchor, protection + anchor, 1)
            break
    else:
        sx = sx.replace("</sheetData>", "</sheetData>" + protection, 1)
    parts[part] = sx.encode("utf-8")

    tmp = xlsm_path + ".tmp"
    zo = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    for n in names:
        zo.writestr(infos[n], parts[n])
    zo.close()
    os.replace(tmp, xlsm_path)
    return part

# COM numeric constants (avoid gencache so we can use an isolated instance)
XL_DOWN               = -4121
XL_PASTE_ALL          = -4104
XL_FREE_FLOATING      = 3
XL_XLSM               = 52     # xlOpenXMLWorkbookMacroEnabled
VBEXT_CT_STDMODULE    = 1

VBA = r'''Option Explicit

Private Const SH As String = "Monthly Plan"

Private Sub LockSheet(ws As Worksheet)
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Private Function FindLabelRow(ws As Worksheet, ByVal txt As String) As Long
    Dim c As Range
    Set c = ws.Range("B1:B400").Find(What:=txt, LookAt:=xlWhole, _
                                     MatchCase:=False, LookIn:=xlValues)
    If c Is Nothing Then FindLabelRow = 0 Else FindLabelRow = c.Row
End Function

Private Sub AddRowsBlock(ByVal markerText As String, _
                         ByVal clearMiles As Boolean, ByVal mergeDE As Boolean)
    Const N As Long = 5
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH)
    Dim mk As Long
    mk = FindLabelRow(ws, markerText)
    If mk = 0 Then
        MsgBox "Could not find the '" & markerText & "' row.", vbExclamation
        Exit Sub
    End If

    Dim lastData As Long, tmpl As Long
    lastData = mk - 1          ' bottom (last-styled) data row
    tmpl = lastData - 1        ' a middle data row to clone from

    Application.ScreenUpdating = False
    On Error Resume Next
    ws.Unprotect ""              ' no password; swallow if the call complains
    On Error GoTo failed

    ' Insert N rows ABOVE the last data row so SUM(...:lastData) auto-expands.
    ws.Rows(lastData).Resize(N).Insert Shift:=xlDown, _
        CopyOrigin:=xlFormatFromLeftOrAbove

    ' Clone the template row (styles, Miles formula, From/To dropdowns) into new rows.
    Dim src As Range, dst As Range
    Set src = ws.Range(ws.Cells(tmpl, "B"), ws.Cells(tmpl, "G"))
    Set dst = ws.Range(ws.Cells(lastData, "B"), ws.Cells(lastData + N - 1, "G"))
    src.Copy Destination:=dst    ' direct copy tiles src across dst; no clipboard needed

    ' Blank the input cells (keep the Miles FORMULA on trip rows).
    Dim r As Long
    For r = lastData To lastData + N - 1
        ws.Cells(r, "C").ClearContents                                  ' Date
        ws.Range(ws.Cells(r, "D"), ws.Cells(r, "E")).ClearContents      ' From/To or Site
        ws.Cells(r, "G").ClearContents                                  ' Notes
        If clearMiles Then ws.Cells(r, "F").ClearContents               ' Miles (extra block)
        If mergeDE Then
            On Error Resume Next
            ws.Range(ws.Cells(r, "D"), ws.Cells(r, "E")).Merge
            On Error GoTo failed
        End If
    Next r

    ' Renumber the "#" column for the whole block.
    Dim hdr As Long
    hdr = lastData
    Do While hdr > 1 And CStr(ws.Cells(hdr, "B").Value) <> "#"
        hdr = hdr - 1
    Loop
    Dim newMarker As Long, k As Long
    newMarker = mk + N
    k = 0
    For r = hdr + 1 To newMarker - 1
        k = k + 1
        ws.Cells(r, "B").Value = k
    Next r

    On Error Resume Next
    LockSheet ws
    On Error GoTo 0
    Application.ScreenUpdating = True
    Application.Goto ws.Cells(lastData, "C"), True
    Exit Sub

failed:
    On Error Resume Next
    LockSheet ws
    Application.ScreenUpdating = True
    MsgBox "Could not add rows: " & Err.Description, vbExclamation
End Sub

Public Sub AddTripRows()
    AddRowsBlock "PLAN TRIP MILES", False, False
End Sub

Public Sub AddExtraRows()
    AddRowsBlock "EXTRA / TEMPORARY MILES", True, True
End Sub

Public Sub FillDownSelection()
    If ActiveSheet.Name <> SH Then
        MsgBox "Go to the Monthly Plan tab first.", vbInformation, "Fill Down"
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rng As Range
    Set rng = Selection
    If rng.Rows.Count < 2 Then
        MsgBox "Select the cell you want to copy AND the cells below it, " & _
               "then click Fill Down.", vbInformation, "Fill Down"
        Exit Sub
    End If
    Dim c As Range
    For Each c In rng.Cells
        If c.Row > rng.Row And c.Locked Then
            MsgBox "Fill Down only works on the input cells " & _
                   "(Date, From, To, Site, Miles, Notes).", vbExclamation, "Fill Down"
            Exit Sub
        End If
    Next c
    On Error Resume Next
    rng.FillDown
End Sub

Public Sub PrintPlan()
    ThisWorkbook.Worksheets(SH).Activate
    On Error Resume Next
    Application.Dialogs(xlDialogPrint).Show
End Sub
'''


def main():
    import win32com, pythoncom, pywintypes
    from win32com.client import dynamic, gencache

    # Force PURE late binding everywhere: purge any early-bound (gen_py) wrappers
    # and stop pywin32 regenerating them, so child objects (Worksheet, etc.) stay
    # dynamic. Early binding mis-marshals Protect/Unprotect's optional args ->
    # 0x800A03EC. gen_py is just a cache; other tools regenerate it on demand.
    gencache.is_readonly = True
    try:
        gp = win32com.__gen_path__
        if os.path.isdir(gp):
            shutil.rmtree(gp, ignore_errors=True)
    except Exception:
        pass

    if not os.path.exists(SRC):
        sys.exit("SRC not found: " + SRC)

    make_unlocked_copy(SRC, WORK)
    print("built unlocked working copy:", os.path.basename(WORK))

    # Fresh, isolated, LATE-BOUND Excel instance.
    # CoCreateInstance(CLSCTX_LOCAL_SERVER) => a brand-new process (won't attach to
    # the user's running Excel), and dynamic.Dispatch => late binding, which — unlike
    # pywin32's gen_py early binding — marshals Protect/Unprotect's many optional
    # args correctly (the gen_py path throws 0x800A03EC on those).
    clsid = pywintypes.IID("Excel.Application")   # resolves ProgID -> CLSID
    disp = pythoncom.CoCreateInstance(
        clsid, None,
        pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.IID_IDispatch)
    xl = dynamic.Dispatch(disp)
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    if int(xl.Workbooks.Count) != 0:
        xl.Quit()
        raise SystemExit("Attached to an Excel that already has workbooks open; aborting.")

    wb = None
    try:
        # positional: FileName, UpdateLinks=0
        wb = xl.Workbooks.Open(WORK, 0)
        print("Excel version:", xl.Version, "| ReadOnly:", wb.ReadOnly)
        if wb.ReadOnly:
            raise SystemExit(
                "working copy opened READ-ONLY -> another Excel has it open. "
                "Please close ALL Excel windows and re-run.")

        # --- inject the VBA module ---
        try:
            proj = wb.VBProject
        except Exception as e:
            raise SystemExit(
                "Cannot access the VBA project (AccessVBOM). Make sure Excel is "
                "fully closed and 'Trust access to the VBA project object model' "
                "is enabled. Underlying error: %r" % (e,))
        # drop a prior copy if re-running
        for i in range(proj.VBComponents.Count, 0, -1):
            comp = proj.VBComponents.Item(i)
            if comp.Name == "mFleet":
                proj.VBComponents.Remove(comp)
        comp = proj.VBComponents.Add(VBEXT_CT_STDMODULE)
        comp.Name = "mFleet"
        comp.CodeModule.AddFromString(VBA.replace("\n", "\r\n"))

        # --- add the buttons on the Monthly Plan sheet ---
        ws = wb.Worksheets("Monthly Plan")
        print("Monthly Plan ProtectContents (should be False now):", ws.ProtectContents)

        # remove any prior FB_ buttons (idempotent re-run)
        for shp in list(ws.Shapes):
            try:
                if shp.Name.startswith("FB_"):
                    shp.Delete()
            except Exception:
                pass

        base = ws.Range("I3")           # to the right of the B:G print area
        left = float(base.Left)
        top = float(base.Top)
        W, H, GAP = 155.0, 26.0, 6.0
        defs = [
            ("FB_addtrip",  "Add 5 trip rows",     "AddTripRows"),
            ("FB_addextra", "Add 5 extra rows",    "AddExtraRows"),
            ("FB_filldown", "Fill Down (Ctrl+D)",  "FillDownSelection"),
            ("FB_print",    "Print",               "PrintPlan"),
        ]
        y = top
        for nm, cap, macro in defs:
            btn = ws.Buttons().Add(left, y, W, H)
            btn.Name = nm
            btn.Caption = cap
            btn.OnAction = macro
            btn.Placement = XL_FREE_FLOATING
            y += H + GAP

        # NOTE: Monthly Plan stays UNPROTECTED here; COM Protect fails on this
        # workbook (0x800A03EC). We re-apply protection via XML after saving.

        # --- save as macro-enabled .xlsm ---
        if os.path.exists(OUT):
            os.remove(OUT)
        wb.SaveAs(OUT, XL_XLSM)   # positional: FileName, FileFormat

        # --- verify ---
        n_btn = sum(1 for s in ws.Shapes if s.Name.startswith("FB_"))
        has_mod = any(proj.VBComponents.Item(i).Name == "mFleet"
                      for i in range(1, proj.VBComponents.Count + 1))
        print("OK  wrote:", OUT)
        print("    buttons on Monthly Plan:", n_btn)
        print("    mFleet module present  :", has_mod)
        print("    dimension              :", ws.UsedRange.Address)
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        xl.Quit()
        if os.path.exists(WORK):
            try:
                os.remove(WORK)
            except Exception:
                pass

    # Only reached if the build succeeded (exceptions propagate past finally).
    part = reprotect_sheet(OUT, "Monthly Plan",
                           '<sheetProtection sheet="1" objects="1" scenarios="1"/>')
    print("re-applied Monthly Plan protection via XML into", part)


if __name__ == "__main__":
    main()
