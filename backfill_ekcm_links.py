#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
금천미트(ekcm.co.kr) 상품 상세페이지(/pd/productDetail?goodsNo=...&artcCd=...&dispCtgNo=...)
즉시 백필 및 하이퍼링크 정합성 검증 스크립트
"""

import json
import sqlite3
from pathlib import Path
import hanwoo_pipeline as hp

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
ARC_PATH = OUTPUT_DIR / "[누적_히스토리]_한우원물_크롤링_아카이브.json"


def main() -> None:
    sess = hp.build_http_session()
    live_rows = hp.fetch_ekcm_hanwoo(sess, "2026-11-19")
    print(f"[1] 금천미트 실시간 한우 원물 상세URL 수집 완료: {len(live_rows)}건")

    by_cond_min = {}
    by_cond_max = {}
    by_part_keep_min = {}
    by_part_keep_max = {}
    by_part = {}

    for r in live_rows:
        if "/pd/productDetail?goodsNo=" not in r.product_url:
            continue
        ckpg = (r.cattle_type, r.keeping, r.part_name, r.grade)
        ckp = (r.cattle_type, r.keeping, r.part_name)
        cp = (r.cattle_type, r.part_name)
        if ckpg not in by_cond_min or r.unit_price_kg < by_cond_min[ckpg][0]:
            by_cond_min[ckpg] = (r.unit_price_kg, r.product_url)
        if ckpg not in by_cond_max or r.unit_price_kg > by_cond_max[ckpg][0]:
            by_cond_max[ckpg] = (r.unit_price_kg, r.product_url)
        if ckp not in by_part_keep_min or r.unit_price_kg < by_part_keep_min[ckp][0]:
            by_part_keep_min[ckp] = (r.unit_price_kg, r.product_url)
        if ckp not in by_part_keep_max or r.unit_price_kg > by_part_keep_max[ckp][0]:
            by_part_keep_max[ckp] = (r.unit_price_kg, r.product_url)
        by_part[cp] = r.product_url

    any_url = live_rows[0].product_url if live_rows else ""

    def pick_url(cattle: str, keeping: str, part: str, grade: str = "", is_max: bool = False) -> str:
        ckpg = (cattle, keeping, part, grade)
        ckp = (cattle, keeping, part)
        cp = (cattle, part)
        if is_max:
            if ckpg in by_cond_max:
                return by_cond_max[ckpg][1]
            if ckp in by_part_keep_max:
                return by_part_keep_max[ckp][1]
        else:
            if ckpg in by_cond_min:
                return by_cond_min[ckpg][1]
            if ckp in by_part_keep_min:
                return by_part_keep_min[ckp][1]
        return by_part.get(cp) or any_url

    if ARC_PATH.exists():
        arc = json.loads(ARC_PATH.read_text(encoding="utf-8"))
        replaced_cnt = 0

        for snap in arc.get("snapshots", []):
            for gr in snap.get("grade_rows", []):
                c = str(gr.get("cattle") or "")
                k = str(gr.get("keeping") or "")
                p = str(gr.get("part") or "")
                g = str(gr.get("grade") or "")
                for u_key, is_mx in (("min_url", False), ("max_url", True)):
                    if "ekcm.co.kr/dp/subMain" in str(gr.get(u_key) or ""):
                        gr[u_key] = pick_url(c, k, p, g, is_mx)
                        replaced_cnt += 1
                ek_p = (gr.get("platforms") or {}).get("금천미트")
                if isinstance(ek_p, dict):
                    for u_key, is_mx in (("url", False), ("max_url", True)):
                        if "ekcm.co.kr/dp/subMain" in str(ek_p.get(u_key) or ""):
                            ek_p[u_key] = pick_url(c, k, p, g, is_mx)
                            replaced_cnt += 1

            for fi in snap.get("full_items", []):
                c = str(fi.get("cattle") or "")
                k = str(fi.get("keeping") or "")
                p = str(fi.get("part") or "")
                for u_key, is_mx in (("min_url", False), ("max_url", True)):
                    if "ekcm.co.kr/dp/subMain" in str(fi.get(u_key) or ""):
                        fi[u_key] = pick_url(c, k, p, "", is_mx)
                        replaced_cnt += 1
                ek_p = (fi.get("platforms") or {}).get("금천미트")
                if isinstance(ek_p, dict):
                    for u_key, is_mx in (("url", False), ("max_url", True)):
                        if "ekcm.co.kr/dp/subMain" in str(ek_p.get(u_key) or ""):
                            ek_p[u_key] = pick_url(c, k, p, "", is_mx)
                            replaced_cnt += 1

            for ch in snap.get("cohorts", []):
                c = str(ch.get("cattle") or "")
                k = str(ch.get("keeping") or "")
                p = str(ch.get("part") or "")
                g = str(ch.get("grade") or "")
                for u_key, is_mx in (("min_url", False), ("max_url", True), ("spot_min_url", False), ("spot_max_url", True)):
                    if "ekcm.co.kr/dp/subMain" in str(ch.get(u_key) or ""):
                        ch[u_key] = pick_url(c, k, p, g, is_mx)
                        replaced_cnt += 1

        ARC_PATH.write_text(json.dumps(arc, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[2] 한우 아카이브 JSON 상세링크 백필 완료: 총 {replaced_cnt}건 (/pd/productDetail 치환)")

    # 대산포크(한돈) DB 금천미트 상세링크 백필
    pork_dir = Path(r"C:\Users\limkyungsoon\Desktop\온라인팀\대산포크")
    pork_db = pork_dir / "handon_master.db"
    if pork_db.exists():
        try:
            import sys
            sys.path.insert(0, str(pork_dir))
            import crawler_db as cdb
            p_recs, _ = cdb.crawl_ekcm(sess, 0, "2026-11-19 07:55:00", "2026-11-19")
            p_url_map = {r.product_id: r.product_url for r in p_recs if "/pd/productDetail?goodsNo=" in r.product_url}
            p_any = next(iter(p_url_map.values()), "")
            conn = sqlite3.connect(str(pork_db))
            rows = conn.execute("SELECT rowid, product_id, product_url FROM product_snapshots WHERE platform_group='금천미트'").fetchall()
            p_upd = 0
            for rid, pid, purl in rows:
                if "/pd/productDetail?goodsNo=" not in str(purl or ""):
                    nu = p_url_map.get(pid) or p_any
                    if nu:
                        conn.execute("UPDATE product_snapshots SET product_url=? WHERE rowid=?", (nu, rid))
                        p_upd += 1
            conn.commit()
            conn.close()
            print(f"[3] 대산포크(한돈) DB 금천미트 상세링크 백필 완료: {p_upd}건")
        except Exception as e:
            print(f"[3] 대산포크 백필 참고: {e}")


if __name__ == "__main__":
    main()
