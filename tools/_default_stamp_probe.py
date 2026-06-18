import pymupdf as fitz
import pikepdf
from pathlib import Path


def view_block(path: Path, kinds=None):
    kinds = kinds or {"Square", "Stamp", "FreeText"}
    doc = fitz.open(path)
    pg = doc[0]
    rm = pg.derotation_matrix
    anns = []
    for a in pg.annots() or []:
        if a.type[1] not in kinds:
            continue
        try:
            r = fitz.Rect(a.rect)
        except Exception:
            continue
        pts = [fitz.Point(r.x0, r.y0) * rm, fitz.Point(r.x1, r.y1) * rm]
        vr = fitz.Rect(min(p.x for p in pts), min(p.y for p in pts), max(p.x for p in pts), max(p.y for p in pts))
        anns.append((a.type[1], (a.get_text() or "")[:20], list(vr), a.rotation))
    doc.close()
    if not anns:
        return []
    xs0 = [a[2][0] for a in anns]
    ys0 = [a[2][1] for a in anns]
    xs1 = [a[2][2] for a in anns]
    ys1 = [a[2][3] for a in anns]
    u = [min(xs0), min(ys0), max(xs1), max(ys1)]
    uw, uh = u[2] - u[0], u[3] - u[1]
    out = []
    for kind, c, r, rot in anns:
        cx = ((r[0] + r[2]) / 2 - u[0]) / uw
        cy = ((r[1] + r[3]) / 2 - u[1]) / uh
        out.append({"kind": kind, "text": c, "rot": rot, "cx": cx, "cy": cy})
    return out


def compare(stamp: Path, out: Path):
    print(f"\n=== {stamp.name} -> {out.name} ===")
    src = view_block(stamp, {"Square", "FreeText"})
    gen = view_block(out, {"Square", "FreeText"})
    src_sq = next(a for a in src if a["kind"] == "Square")
    gen_sq = next(a for a in gen if a["kind"] == "Square")
    print(
        f"Square src cx={src_sq['cx']:.3f} cy={src_sq['cy']:.3f} | "
        f"gen cx={gen_sq['cx']:.3f} cy={gen_sq['cy']:.3f}"
    )
    src_ft = [a for a in src if a["kind"] == "FreeText"]
    gen_ft = {a["text"].strip(): a for a in gen if a["kind"] == "FreeText"}
    max_d = 0.0
    for s in src_ft:
        key = s["text"].strip()
        g = gen_ft.get(key)
        if not g:
            print(f"  missing in output: {key[:30]!r}")
            continue
        d = ((s["cx"] - g["cx"]) ** 2 + (s["cy"] - g["cy"]) ** 2) ** 0.5
        max_d = max(max_d, d)
        if d > 0.05:
            print(
                f"  mismatch {key[:20]!r}: src ({s['cx']:.3f},{s['cy']:.3f}) "
                f"gen ({g['cx']:.3f},{g['cy']:.3f}) d={d:.4f}"
            )
    print(f"FreeText pairs={len(src_ft)} max normalized center delta={max_d:.4f}")
    with pikepdf.open(out) as doc:
        sub = {}
        for ref in doc.pages[0].Annots or []:
            st = str(ref.get("/Subtype", ""))
            sub[st] = sub.get(st, 0) + 1
        ft_rots = {
            int(ref.get("/Rotate", 0) or 0)
            for ref in doc.pages[0].Annots or []
            if str(ref.get("/Subtype", "")) == "/FreeText"
        }
    print(f"subtypes {sub} FreeText /Rotate values {sorted(ft_rots)}")


for stamp, out in [
    (Path("stamps/Peer_Review_Stamp.pdf"), Path("test/output/default-review-stamp-test.pdf")),
    (Path("stamps/IC_Stamp.pdf"), Path("test/output/default-check-stamp-test.pdf")),
]:
    compare(stamp, out)
