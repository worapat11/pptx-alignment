#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
align_like_reference.py
=======================
จัดตำแหน่ง "รูปอะไหล่ / รหัสอะไหล่ / ชื่ออะไหล่ / ช่องสี" ของทุกหน้า
ให้ตรงกับหน้าต้นแบบ (reference) ที่คุณจัดไว้ถูกต้องแล้ว

วิธีใช้ (ตัวอย่าง):
    python align_like_reference.py deck.pptx --ref 1
    python align_like_reference.py deck.pptx --ref 1 --targets 2-5
    python align_like_reference.py deck.pptx --ref 4 --targets 6 --dry-run

หลักการ:
  1. หา "ชิ้นอะไหล่" ในแต่ละหน้า = กล่องรหัส (เช่น 53207-K12-V00) + ชื่อใต้รหัส
     + รูปเหนือรหัส + ช่องสีสี่เหลี่ยมข้างรหัส
  2. จับคู่ชิ้นระหว่างหน้าต้นแบบกับหน้าปลายทางด้วยเลขอะไหล่ 5 หลักแรก
     (ต่อท้ายสี/รุ่น เช่น -ZM -ZA ต่างกันได้)
  3. ย้ายชิ้นของหน้าปลายทางไปอยู่ตำแหน่งเดียวกับต้นแบบ
  4. ชิ้นที่จับคู่ไม่ได้ จะไม่ถูกแตะ และแจ้งไว้ในรายงาน

ต้องติดตั้ง:  pip install python-pptx
ไฟล์ต้นฉบับจะไม่ถูกแก้ไข โปรแกรมบันทึกเป็นไฟล์ใหม่เสมอ
"""
import argparse
import csv
import os
import re
import statistics
import sys
import unicodedata

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# ----------------------------------------------------------------------------
# ค่าตั้งต้น (ถ้าไฟล์รุ่นอื่นโครงสร้างต่างไป แก้ตรงนี้ได้)
# ----------------------------------------------------------------------------
CODE_REGEX = re.compile(r"^[0-9A-Z]{5}-[0-9A-Z]{2,4}-[0-9A-Z]{2,}$")   # รูปแบบรหัส เช่น 53207-K12-V00 (ปรับได้ด้วย --code-regex)
PART_KEY_LEN = 5                                   # ใช้กี่หลักแรกในการจับคู่ชิ้น
CHIP_MAX_W, CHIP_MAX_H = 130_000, 120_000          # ขนาดสูงสุดของ "ช่องสีสี่เหลี่ยม" (EMU)
BIKE_MIN_AREA = 2.0e12                             # รูปที่ใหญ่กว่านี้ = รูปรถ (ไม่ย้าย)
NAME_MAX_BELOW = 420_000                           # ชื่อต้องอยู่ใต้รหัสไม่เกินเท่านี้
NAME_MAX_DX = 650_000                              # และห่างจากกึ่งกลางรหัสในแนวนอนไม่เกินเท่านี้
IMG_MAX_DX = 650_000                               # รูปต้องห่างจากรหัสในแนวนอนไม่เกินเท่านี้
THAI = re.compile(r"[\u0E00-\u0E7F]")


# ----------------------------------------------------------------------------
# ส่วนอ่านโครงสร้างหน้า
# ----------------------------------------------------------------------------
def txt(shape):
    return re.sub(r"\s+", " ", shape.text_frame.text).strip() if shape.has_text_frame else ""


def same_text(a, b):
    """เทียบข้อความโดยไม่สนช่องว่าง/การขึ้นบรรทัดใหม่"""
    return re.sub(r"\s+", "", a.text_frame.text) == re.sub(r"\s+", "", b.text_frame.text)


def est_line_width(text, per_glyph=60_000):
    """ประมาณความกว้างข้อความ 1 บรรทัดของฟอนต์ไทย (ไม่นับสระ/วรรณยุกต์บน-ล่าง) + ขอบกล่อง"""
    w = 0
    for ch in text:
        if unicodedata.category(ch) == "Mn":
            continue
        w += 30_000 if ch == " " else (68_000 if ord(ch) < 128 else per_glyph)
    return int(w + 182_880)


def cx_(s):  # กึ่งกลางแนวนอน
    return s.left + s.width / 2


def cy_(s):  # กึ่งกลางแนวตั้ง
    return s.top + s.height / 2


def parse_slide(slide):
    """คืนค่า (items, bike, frames, notes)"""
    shapes = list(slide.shapes)
    notes = []
    if any(s.shape_type == MSO_SHAPE_TYPE.GROUP for s in shapes):
        notes.append("มี Group อยู่ในหน้า: กลุ่มจะไม่ถูกนำมาจัด (ให้ Ungroup ก่อน)")

    codes = [s for s in shapes if s.has_text_frame and CODE_REGEX.match(txt(s).replace(" ", ""))]
    code_ids = {id(c) for c in codes}
    names = [s for s in shapes if s.has_text_frame and id(s) not in code_ids
             and THAI.search(txt(s))]
    pics = [s for s in shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]

    bike = None
    if pics:
        biggest = max(pics, key=lambda p: p.width * p.height)
        if biggest.width * biggest.height > BIKE_MIN_AREA:
            bike = biggest
    chips = [p for p in pics if p is not bike and p.width <= CHIP_MAX_W and p.height <= CHIP_MAX_H]
    mains = [p for p in pics if p is not bike and p not in chips]

    items = [{"code": c, "name": None, "mains": [], "chips": []} for c in codes]

    # รูปหลัก -> รหัสที่อยู่ใต้รูปและใกล้ที่สุด
    for p in mains:
        bottom, best = p.top + p.height, None
        for i, it in enumerate(items):
            c = it["code"]
            dy, dx = c.top - bottom, abs(cx_(p) - cx_(c))
            if -150_000 < dy < 900_000 and dx < IMG_MAX_DX:
                score = dx + abs(dy) * 0.6
                if best is None or score < best[0]:
                    best = (score, i)
        if best:
            items[best[1]]["mains"].append(p)
        else:
            notes.append(f"รูปที่ไม่รู้ว่าเป็นของชิ้นไหน (ไม่ถูกย้าย): {p.name}")

    # ช่องสี -> รหัสที่ใกล้ที่สุด
    for p in chips:
        best = None
        for i, it in enumerate(items):
            c = it["code"]
            dyc, dx = abs(cy_(p) - cy_(c)), abs(cx_(p) - cx_(c))
            if dyc < 420_000 and dx < 700_000:
                score = dx + dyc * 0.5
                if best is None or score < best[0]:
                    best = (score, i)
        if best:
            items[best[1]]["chips"].append(p)
        else:
            notes.append(f"ช่องสีที่ไม่รู้ว่าเป็นของชิ้นไหน (ไม่ถูกย้าย): {p.name}")

    # ชื่อ -> รหัสที่อยู่เหนือชื่อและใกล้ที่สุด (จับคู่แบบ 1 ต่อ 1)
    cands = []
    for n in names:
        for i, it in enumerate(items):
            c = it["code"]
            dy, dx = n.top - c.top, abs(cx_(n) - cx_(c))
            if 0 < dy < NAME_MAX_BELOW and dx < NAME_MAX_DX:
                cands.append((dx + dy * 0.5, i, n))
    cands.sort(key=lambda t: t[0])
    used_i, used_n = set(), set()
    for _, i, n in cands:
        if i in used_i or id(n) in used_n:
            continue
        items[i]["name"] = n
        used_i.add(i)
        used_n.add(id(n))

    # กรอบใหญ่ที่ไม่มีข้อความ (กรอบนอกของตาราง)
    sw, sh = slide.part.package.presentation_part.presentation.slide_width, \
        slide.part.package.presentation_part.presentation.slide_height
    frames = [s for s in shapes if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
              and not txt(s) and s.width * s.height > 0.4 * sw * sh]
    return items, bike, frames, notes


def add_keys(items):
    """ตั้งคีย์จับคู่ = เลขอะไหล่ 5 หลัก (ถ้าซ้ำในหน้าเดียวกัน เรียงตามตำแหน่งแล้วใส่ลำดับ)"""
    seen = {}
    for it in sorted(items, key=lambda i: (i["code"].top, i["code"].left)):
        num = txt(it["code"])[:PART_KEY_LEN]
        k = seen.get(num, 0)
        seen[num] = k + 1
        it["key"] = (num, k)


# ----------------------------------------------------------------------------
# ส่วนย้ายตำแหน่ง
# ----------------------------------------------------------------------------
def center_to(shape, ref):
    """ย้าย shape ให้กึ่งกลางตรงกับ ref (ขนาดของ shape ไม่เปลี่ยน)"""
    shape.left = int(round(cx_(ref) - shape.width / 2))
    shape.top = int(round(cy_(ref) - shape.height / 2))


def copy_box(shape, ref):
    shape.left, shape.top, shape.width, shape.height = ref.left, ref.top, ref.width, ref.height


def by_pos(shapes):
    return sorted(shapes, key=lambda p: (p.left, p.top))


def align_slide(target_items, ref_by_key, ref_chip_offset, alias, args_keep_wrap=False, ref_by_name=None):
    stats = dict(matched=0, unmatched=[], warn=[], moved=0)
    for it in target_items:
        num, k = it["key"]
        key = (alias.get(num, num), k)
        ref = ref_by_key.get(key)
        if ref is None and ref_by_name and it["name"] is not None:      # สำรอง: ชื่อเหมือนกันทุกตัวอักษร
            ref = ref_by_name.get(re.sub(r"\s+", "", it["name"].text_frame.text))
            if ref is not None:
                stats["warn"].append(f"{txt(it['code'])}: เลขอะไหล่ไม่ตรง จับคู่ด้วยชื่อแทน")
        if ref is None:
            stats["unmatched"].append(f"{txt(it['code'])}  {txt(it['name']) if it['name'] else ''}")
            continue
        stats["matched"] += 1
        label = txt(it["code"])
        old_cx, old_cy = cx_(it["code"]), cy_(it["code"])   # ตำแหน่งรหัสเดิม (ใช้พาช่องสีตามไป)

        center_to(it["code"], ref["code"]); stats["moved"] += 1

        if it["name"] is not None and ref["name"] is not None:
            if same_text(it["name"], ref["name"]):
                copy_box(it["name"], ref["name"])         # ชื่อเหมือนกัน: ใช้กล่องเดียวกัน (รวมขนาด)
            else:
                # ชื่อต่างกัน: จัดกึ่งกลางแนวนอน และให้บรรทัดแรกตรงกับต้นแบบ
                nm, rn = it["name"], ref["name"]
                if nm.height > rn.height * 1.3 and not args_keep_wrap:   # ยาวจนขึ้น 2 บรรทัด -> ทำเป็นบรรทัดเดียว
                    nm.height = rn.height
                    nm.width = max(nm.width, min(int(est_line_width(txt(nm)) * 1.08), 1_600_000))
                nm.left = int(round(cx_(rn) - nm.width / 2))
                nm.top = rn.top
                stats["warn"].append(f"{label}: ชื่อต่างจากต้นแบบ ('{txt(nm)}') จัดกึ่งกลางให้ ตรวจว่าไม่ชนข้างเคียง")
            stats["moved"] += 1

        tm, rm = by_pos(it["mains"]), by_pos(ref["mains"])
        old_main = [(cx_(a), cy_(a)) for a in tm]            # ตำแหน่งเดิมของรูป (ใช้พารูปส่วนเกินตามไป)
        if len(tm) != len(rm):
            stats["warn"].append(f"{label}: จำนวนรูปไม่เท่าต้นแบบ ({len(tm)} vs {len(rm)})"
                                 + (" รูปส่วนเกินจะตามรูปแรกไปโดยคงระยะเดิม" if len(tm) > len(rm) and rm else ""))
        for a, b in zip(tm, rm):
            center_to(a, b); stats["moved"] += 1
        if rm:
            for a, (ox, oy) in list(zip(tm, old_main))[len(rm):]:
                a.left = int(round(cx_(rm[0]) + (ox - old_main[0][0]) - a.width / 2))
                a.top = int(round(cy_(rm[0]) + (oy - old_main[0][1]) - a.height / 2))
                stats["moved"] += 1

        tc, rc = by_pos(it["chips"]), by_pos(ref["chips"])
        for a, b in zip(tc, rc):
            center_to(a, b); stats["moved"] += 1
        if len(tc) > len(rc):                              # ต้นแบบไม่มีช่องสีนี้ -> ใช้ระยะเฉลี่ยจากรหัส
            if ref_chip_offset is None:                     # ต้นแบบไม่มีช่องสีเลย -> พาช่องสีตามรหัสไปโดยคงระยะเดิม
                for a in tc[len(rc):]:
                    a.left = int(round(cx_(ref["code"]) + (cx_(a) - old_cx) - a.width / 2))
                    a.top = int(round(cy_(ref["code"]) + (cy_(a) - old_cy) - a.height / 2))
                    stats["moved"] += 1
                stats["chips_follow"] = stats.get("chips_follow", 0) + 1
            else:
                for a in tc[len(rc):]:
                    a.left = int(round(cx_(ref["code"]) + ref_chip_offset[0] - a.width / 2))
                    a.top = int(round(cy_(ref["code"]) + ref_chip_offset[1] - a.height / 2))
                    stats["moved"] += 1
                stats["warn"].append(f"{label}: ต้นแบบไม่มีช่องสีนี้ วางตามระยะเฉลี่ยของต้นแบบ")
        elif len(tc) < len(rc):
            stats["warn"].append(f"{label}: ต้นแบบมีช่องสี {len(rc)} ช่อง แต่หน้านี้มี {len(tc)} ช่อง")
    return stats


def parse_range(text, total):
    out = set()
    for part in text.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-")
            out.update(range(int(a), int(b) + 1))
        elif part:
            out.add(int(part))
    bad = [n for n in out if not 1 <= n <= total]
    if bad:
        sys.exit(f"เลขหน้าไม่ถูกต้อง: {sorted(bad)} (ไฟล์นี้มี {total} หน้า)")
    return sorted(out)


def load_alias(path):
    """ไฟล์ CSV 2 คอลัมน์: เลขอะไหล่หน้าปลายทาง,เลขอะไหล่ในหน้าต้นแบบ  (สำหรับอะไหล่ที่เปลี่ยนเลข)"""
    alias = {}
    if path:
        with open(path, encoding="utf-8-sig", newline="") as f:
            for row in csv.reader(f):
                if len(row) >= 2 and row[0].strip() and not row[0].startswith("#"):
                    alias[row[0].strip()[:PART_KEY_LEN]] = row[1].strip()[:PART_KEY_LEN]
    return alias


# ----------------------------------------------------------------------------
def run_file(path, args, out_path, code_re=None):
    """จัดไฟล์เดียว คืนค่า (สรุป dict, exit_code)"""
    prs = Presentation(path)
    total = len(prs.slides)
    if not 1 <= args.ref <= total:
        raise ValueError(f"หน้าต้นแบบต้องอยู่ระหว่าง 1-{total} (ไฟล์นี้มี {total} หน้า)")
    targets = [n for n in (parse_range(args.targets, total) if args.targets
                           else range(1, total + 1)) if n != args.ref]

    ref_items, _, ref_frames, ref_notes = parse_slide(prs.slides[args.ref - 1])
    add_keys(ref_items)
    if not ref_items:
        cands = []
        for shp in prs.slides[args.ref - 1].shapes:
            t = txt(shp) if shp.has_text_frame else ""
            if "-" in t and any(ch.isdigit() for ch in t) and len(t) < 30:
                cands.append(t)
        hint = ("ข้อความที่หน้าตาคล้ายรหัสในหน้านั้น: " + " | ".join(cands[:6])) if cands else "ไม่พบข้อความคล้ายรหัส"
        raise ValueError("ไม่พบชิ้นอะไหล่ในหน้าต้นแบบ ลองใช้ --code-regex ปรับรูปแบบรหัส. " + hint)
    ref_by_key = {it["key"]: it for it in ref_items}
    names_cnt = {}
    for it in ref_items:
        if it["name"] is not None:
            k = re.sub(r"\s+", "", it["name"].text_frame.text)
            names_cnt[k] = names_cnt.get(k, 0) + 1
    ref_by_name = {re.sub(r"\s+", "", it["name"].text_frame.text): it for it in ref_items
                   if it["name"] is not None and names_cnt[re.sub(r"\s+", "", it["name"].text_frame.text)] == 1}
    offs = [(cx_(c) - cx_(it["code"]), cy_(c) - cy_(it["code"])) for it in ref_items for c in it["chips"]]
    ref_chip_offset = (statistics.median(o[0] for o in offs), statistics.median(o[1] for o in offs)) if offs else None
    alias = load_alias(args.alias)
    print(f"หน้าต้นแบบ {args.ref}: พบ {len(ref_items)} ชิ้น, ช่องสี {len(offs)} ช่อง")
    for n in ref_notes:
        print("  หมายเหตุ:", n)

    summary = dict(pages=0, matched=0, found=0, unmatched=0)
    for n in targets:
        items, bike, frames, notes = parse_slide(prs.slides[n - 1])
        add_keys(items)
        st = align_slide(items, ref_by_key, ref_chip_offset, alias, args.keep_name_wrap, ref_by_name)
        summary["pages"] += 1
        summary["found"] += len(items)
        summary["matched"] += st["matched"]
        summary["unmatched"] += len(st["unmatched"])
        print(f"\nหน้า {n}: พบ {len(items)} ชิ้น, จับคู่ได้ {st['matched']}, ย้าย {st['moved']} ชิ้นส่วน")
        for msg in notes:
            print("  หมายเหตุ:", msg)
        if st.get("chips_follow"):
            print(f"  หมายเหตุ: หน้าต้นแบบไม่มีช่องสี จึงพาช่องสี {st['chips_follow']} ชิ้นตามรหัสไปโดยคงระยะเดิม "
                  f"(ถ้าต้องการวางแบบเฉพาะ ให้จัดหน้าที่มีช่องสีเป็นต้นแบบอีกชุด)")
        for msg in st["warn"]:
            print("  เตือน:", msg)
        for u in st["unmatched"]:
            print("  จับคู่ไม่ได้ (ไม่ถูกแตะ):", u)

        if frames and ref_frames and not args.no_frame:
            main_ref = max(ref_frames, key=lambda s: s.width * s.height)
            main_t = max(frames, key=lambda s: s.width * s.height)
            copy_box(main_t, main_ref)
            for f in [f for f in frames if f is not main_t]:
                if args.remove_extra_frames:
                    f._element.getparent().remove(f._element)
                    print(f"  ลบกรอบเกิน: {f.name}")
                else:
                    print(f"  พบกรอบเกิน '{f.name}' (ใส่ --remove-extra-frames เพื่อลบ)")

    if args.dry_run:
        print("\n[dry-run] ไม่ได้บันทึกไฟล์")
        return summary
    if os.path.abspath(out_path) == os.path.abspath(path):
        raise ValueError("ไฟล์ผลลัพธ์ต้องไม่ทับไฟล์ต้นฉบับ")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    prs.save(out_path)
    print(f"\nบันทึกแล้ว: {out_path}")
    return summary


def main():
    global CODE_REGEX
    ap = argparse.ArgumentParser(description="จัดทุกหน้าให้ตรงกับหน้าต้นแบบ (ไฟล์เดียว หรือทั้งโฟลเดอร์)")
    ap.add_argument("path", help="ไฟล์ .pptx หรือโฟลเดอร์ที่มีไฟล์ .pptx หลายไฟล์")
    ap.add_argument("--ref", type=int, required=True, help="เลขหน้าต้นแบบ (นับจาก 1) ใช้กับทุกไฟล์")
    ap.add_argument("--targets", help="หน้าที่จะจัด เช่น 2-5 หรือ 2,4,6 (ค่าเริ่มต้น: ทุกหน้าที่เหลือ)")
    ap.add_argument("--out", help="ชื่อไฟล์ผลลัพธ์ (ใช้กับไฟล์เดียว) ค่าเริ่มต้น: <ชื่อไฟล์>_aligned.pptx "
                                  "ถ้าเป็นโฟลเดอร์ จะบันทึกไว้ในโฟลเดอร์ย่อย aligned")
    ap.add_argument("--alias", help="ไฟล์ CSV จับคู่เลขอะไหล่ที่เปลี่ยนเลข (ปลายทาง,ต้นแบบ)")
    ap.add_argument("--code-regex", help="รูปแบบรหัสอะไหล่ (regex) ถ้ารุ่นไหนรหัสหน้าตาต่างจากปกติ")
    ap.add_argument("--no-frame", action="store_true", help="ไม่ปรับกรอบนอกให้เท่าต้นแบบ")
    ap.add_argument("--remove-extra-frames", action="store_true",
                    help="ลบกรอบใหญ่ที่เกินมา (เช่น กรอบเก่าแคบ ๆ ที่ต้นแบบไม่มี)")
    ap.add_argument("--keep-name-wrap", action="store_true",
                    help="ชื่อที่ต่างจากต้นแบบและยาวกว่า ให้คงขึ้น 2 บรรทัดตามเดิม (ค่าเริ่มต้น: ทำเป็นบรรทัดเดียว)")
    ap.add_argument("--dry-run", action="store_true", help="แสดงรายงานอย่างเดียว ไม่บันทึกไฟล์")
    args = ap.parse_args()
    if args.code_regex:
        CODE_REGEX = re.compile(args.code_regex)

    if os.path.isdir(args.path):
        files = sorted(f for f in os.listdir(args.path)
                       if f.lower().endswith(".pptx") and not f.startswith("~$") and not f.lower().endswith("_aligned.pptx"))
        if not files:
            sys.exit("ไม่พบไฟล์ .pptx ในโฟลเดอร์นี้")
        jobs = [(os.path.join(args.path, f), os.path.join(args.path, "aligned", f)) for f in files]
    else:
        out = args.out or os.path.splitext(args.path)[0] + "_aligned.pptx"
        jobs = [(args.path, out)]

    results, exit_code = [], 0
    for path, out in jobs:
        print("=" * 70)
        print("ไฟล์:", os.path.basename(path))
        try:
            results.append((os.path.basename(path), run_file(path, args, out), None))
        except Exception as e:                                # ไฟล์หนึ่งพัง ไม่ให้กระทบไฟล์อื่น
            print("  ผิดพลาด:", e)
            results.append((os.path.basename(path), None, str(e)))
            exit_code = 1
    if len(jobs) > 1 or exit_code:
        print("=" * 70)
        print("สรุป")
        for name, sm, err in results:
            if err:
                print(f"  ✗ {name}: {err}")
            else:
                flag = "✓" if sm["unmatched"] == 0 else "!"
                print(f"  {flag} {name}: {sm['pages']} หน้า, พบ {sm['found']} ชิ้น, จับคู่ได้ {sm['matched']}, "
                      f"จับคู่ไม่ได้ {sm['unmatched']}")
                if sm["unmatched"]:
                    exit_code = exit_code or 2
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
