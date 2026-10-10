# compare_tsv.py : diff the ALV dumps produced by S01_compare_query.vbs
#   usage: python compare_tsv.py [raw_dir] [out_md]
# Old dump : results/raw/<CASE>_ZCRMR052.tsv
# New dump : results/raw/<CASE>_ZCRMR052S.tsv
# Columns excluded from the comparison (per scenario): SUBS_TYPE
import sys, os, glob, io, datetime

EXCLUDE_COLS = {"SUBS_TYPE"}

RAW = sys.argv[1] if len(sys.argv) > 1 else r"C:\sap-test\results\raw"
OUT = sys.argv[2] if len(sys.argv) > 2 else None


def read_dump(path):
    """Return dict with meta, cols, coltitles, rows."""
    with io.open(path, encoding="utf-16") as f:
        lines = f.read().split("\n")
    lines = [l.rstrip("\r") for l in lines]
    meta, cols, coltitles, rows = {}, [], {}, []
    in_data = False
    header = None
    for l in lines:
        if not l:
            continue
        if in_data:
            if header is None:
                header = l.split("\t")
                continue
            rows.append(l.split("\t"))
            continue
        if l.startswith("#DATA"):
            in_data = True
            continue
        p = l.split("\t")
        k = p[0]
        if k == "#COL":
            cols.append(p[2])
            coltitles[p[2]] = p[3] if len(p) > 3 else ""
        else:
            meta[k.lstrip("#")] = p[1] if len(p) > 1 else ""
    return {"meta": meta, "cols": cols, "coltitles": coltitles,
            "header": header or [], "rows": rows, "path": path}


def keep_cols(cols):
    return [c for c in cols if c not in EXCLUDE_COLS]


def project(dump, cmpcols):
    """rows as list of tuples over cmpcols, looked up by column NAME."""
    hdr = dump["header"]
    idx = {c: i for i, c in enumerate(hdr)}
    out = []
    for r in dump["rows"]:
        out.append(tuple(r[idx[c]] if c in idx and idx[c] < len(r) else "<MISSING>"
                         for c in cmpcols))
    return out


def compare(case, old, new):
    res = {"case": case, "verdict": "PASS", "notes": []}
    n = res["notes"]

    if old["meta"].get("GRID") == "NOT_FOUND" or new["meta"].get("GRID") == "NOT_FOUND":
        res["verdict"] = "ERROR"
        n.append("ALV 그리드를 찾지 못함 (old=%s new=%s)" %
                 (old["meta"].get("GRID", "ok"), new["meta"].get("GRID", "ok")))
        return res

    # selection screen field presence
    so, sn = old["meta"].get("SELSTATE", ""), new["meta"].get("SELSTATE", "")
    if so != sn:
        res["verdict"] = "FAIL"
        n.append("선택화면 필드 구성 불일치: old[%s] new[%s]" % (so, sn))

    # row counts
    ro = int(old["meta"].get("ROWCOUNT", -1))
    rn = int(new["meta"].get("ROWCOUNT", -1))
    res["rows_old"], res["rows_new"] = ro, rn
    if ro != rn:
        res["verdict"] = "FAIL"
        n.append("조회 건수 불일치: 기존 %d / 신규 %d" % (ro, rn))

    # title counts
    to = old["meta"].get("TITLECOUNT", "")
    tn = new["meta"].get("TITLECOUNT", "")
    res["title_old"], res["title_new"] = to, tn
    if to != tn:
        res["verdict"] = "FAIL"
        n.append("제목 건수 불일치: 기존 (%s) / 신규 (%s)" % (to, tn))

    # columns
    co, cn = old["cols"], new["cols"]
    only_old = [c for c in co if c not in cn]
    only_new = [c for c in cn if c not in co and c not in EXCLUDE_COLS]
    excluded_present = [c for c in cn if c in EXCLUDE_COLS]
    res["cols_old"], res["cols_new"] = co, cn
    res["excluded_present"] = excluded_present
    if only_old or only_new:
        res["verdict"] = "FAIL"
        if only_old:
            n.append("신규에 없는 컬럼: %s" % ", ".join(only_old))
        if only_new:
            n.append("신규에만 있는 컬럼(제외항목 아님): %s" % ", ".join(only_new))

    # column order over comparable columns
    ko, kn = keep_cols(co), keep_cols(cn)
    if ko == kn:
        res["colorder"] = "동일"
    else:
        res["colorder"] = "상이"
        res["verdict"] = "FAIL"
        n.append("컬럼 순서 불일치")

    # column titles
    tdiff = []
    for c in ko:
        if c in kn:
            a = old["coltitles"].get(c, "")
            b = new["coltitles"].get(c, "")
            if a != b:
                tdiff.append("%s: old[%s] new[%s]" % (c, a, b))
    if tdiff:
        res["verdict"] = "FAIL"
        n.append("컬럼명 불일치 %d건: %s" % (len(tdiff), "; ".join(tdiff[:5])))

    # how many rows were actually dumped (row cap)
    do = int(old["meta"].get("DUMPEDROWS", len(old["rows"])))
    dn = int(new["meta"].get("DUMPEDROWS", len(new["rows"])))
    res["dumped_old"], res["dumped_new"] = do, dn
    if do == 0 and dn == 0 and ro > 0:
        res["cell_diffs"] = None
        n.append("건수·컬럼만 비교 (셀 덤프 0행, rowcap=0 실행)")
        if ro != rn:
            res["verdict"] = "FAIL"
        return res
    if do < ro or dn < rn:
        n.append("셀 비교는 상위 %d행만 수행 (전체 %d행 중)" % (min(do, dn), ro))

    # cell values : compare the columns present in BOTH, matched by name
    common = [c for c in ko if c in kn]
    res["common_cols"] = len(common)
    if ro == rn and common:
        ro_rows = project(old, common)
        rn_rows = project(new, common)
        diffs = []
        for i, (a, b) in enumerate(zip(ro_rows, rn_rows)):
            if a != b:
                for j, c in enumerate(common):
                    if a[j] != b[j]:
                        diffs.append("row%d %s: old[%s] new[%s]" % (i, c, a[j], b[j]))
        res["cell_diffs"] = len(diffs)
        if diffs:
            # order may differ (ALV layout sort) - compare as multisets too
            same_set = sorted(ro_rows) == sorted(rn_rows)
            res["same_set"] = same_set
            if same_set:
                n.append("공통 %d개 컬럼 기준 **행 집합은 완전히 동일**하나 정렬 순서가 다름 "
                         "(%d행 중 위치가 다른 행 존재). ALV 레이아웃의 정렬 기준 차이로 보인다."
                         % (len(common), len(ro_rows)))
            else:
                res["verdict"] = "FAIL"
                only_o = [r for r in ro_rows if r not in rn_rows]
                only_n = [r for r in rn_rows if r not in ro_rows]
                n.append("공통 %d개 컬럼 셀 값 불일치. 정렬 무시 비교에서도 불일치: "
                         "기존에만 %d행, 신규에만 %d행" % (len(common), len(only_o), len(only_n)))
                n.append("예: " + "; ".join(diffs[:5]))
        else:
            res["same_set"] = True
    else:
        res["cell_diffs"] = None
        n.append("건수가 달라 셀 비교 생략")

    if ro == 0 and rn == 0:
        n.append("양쪽 모두 0건 — 데이터가 없어 실질 비교 불가")
        if res["verdict"] == "PASS":
            res["verdict"] = "NO_DATA"
    return res


def main():
    pairs = {}
    for p in glob.glob(os.path.join(RAW, "*_ZCRMR052.tsv")):
        case = os.path.basename(p)[: -len("_ZCRMR052.tsv")]
        newp = os.path.join(RAW, case + "_ZCRMR052S.tsv")
        if os.path.exists(newp):
            pairs[case] = (p, newp)
        else:
            pairs[case] = (p, None)

    results = []
    for case in sorted(pairs):
        op, np_ = pairs[case]
        if np_ is None:
            results.append({"case": case, "verdict": "ERROR",
                            "notes": ["신규 프로그램 덤프 없음"]})
            continue
        results.append(compare(case, read_dump(op), read_dump(np_)))

    lines = []
    lines.append("# ZCRM052 vs ZCRMR052S 비교 결과")
    lines.append("")
    lines.append("생성: %s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
    lines.append("")
    lines.append("| 케이스 | 판정 | 기존 건수 | 신규 건수 | 컬럼(기존/신규) | 컬럼순서 | 비교행 | 셀 불일치 |")
    lines.append("|---|---|---|---|---|---|---|---|")
    for r in results:
        lines.append("| %s | **%s** | %s | %s | %s / %s | %s | %s | %s |" % (
            r["case"], r["verdict"],
            r.get("rows_old", "-"), r.get("rows_new", "-"),
            len(r.get("cols_old", [])) or "-", len(r.get("cols_new", [])) or "-",
            r.get("colorder", "-"),
            r.get("dumped_old", "-"),
            "-" if r.get("cell_diffs") is None else r.get("cell_diffs")))
    lines.append("")
    for r in results:
        if r["notes"]:
            lines.append("## %s" % r["case"])
            for x in r["notes"]:
                lines.append("- %s" % x)
            if r.get("excluded_present"):
                lines.append("- 비교 제외 컬럼이 신규에 존재: %s" % ", ".join(r["excluded_present"]))
            lines.append("")
    out = "\n".join(lines)
    print(out)
    if OUT:
        with io.open(OUT, "w", encoding="utf-8") as f:
            f.write(out)
        print("\nwritten: %s" % OUT)


if __name__ == "__main__":
    main()
