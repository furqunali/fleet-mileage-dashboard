#!/usr/bin/env python3
r"""
fix_plan_merged_inputs.py — fix the locked merged input cells on Monthly Plan.

Employee (B10:C10), Month (D10:E10) and Rate (F10:G10) are MERGED. Their anchor
cells (B10/D10/F10) are unlocked (s77) but the second half (C10/E10/G10) were
left locked (s78). On a protected sheet Excel only lets you TYPE into a merged
cell when EVERY cell in the merge is unlocked — so Employee/Rate couldn't be
typed into. Flip C10/E10/G10 from s78 -> s77 (identical look, just unlocked).

Base : Latest Version (v7)/...v7.xlsx (current protected build)
Output: <temp>  (caller COM-verifies, then deploys to master + folder)
"""
import zipfile, os, re
HERE=r"."
BASE=os.path.join(HERE,"Latest Version (v7)","Fleet_Distance_Dashboard_v7.xlsx")
TMP =os.path.join(HERE,"_v7_fix.tmp.xlsx")

def main():
    zin=zipfile.ZipFile(BASE,"r"); names=zin.namelist()
    parts={n:zin.read(n) for n in names}; infos={i.filename:i for i in zin.infolist()}
    zin.close()
    # locate Monthly Plan sheet by content
    mpf=next(n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$",n)
             and b"MONTHLY TRAVEL PLAN" in parts[n])
    xml=parts[mpf].decode("utf-8")
    for ref in ("C10","E10","G10"):
        old=f'<c r="{ref}" s="78"/>'; new=f'<c r="{ref}" s="77"/>'
        assert old in xml, f"{ref}: expected {old} not found"
        xml=xml.replace(old,new,1)
    parts[mpf]=xml.encode("utf-8")
    zo=zipfile.ZipFile(TMP,"w",zipfile.ZIP_DEFLATED)
    for n in names:
        zo.writestr(infos[n],parts[n]) if n in infos else zo.writestr(n,parts[n])
    zo.close()
    print("fixed merged input cells in",mpf,"-> WROTE",TMP)

if __name__=="__main__":
    main()
