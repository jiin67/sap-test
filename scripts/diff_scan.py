# -*- coding: utf-8 -*-
# diff_scan.py : 두 S18 소스 스캔 로그를 비교해 오브젝트별로 추가·삭제된 줄을 보여 준다(줄 번호는 무시).
#   usage: python diff_scan.py <old S18 log> <new S18 log>
import io, re, sys, collections
sys.stdout.reconfigure(encoding="utf-8")


def parse(p):
    txt = io.open(p, encoding="utf-16").read().splitlines()
    obj, out = None, collections.defaultdict(list)
    for l in txt:
        m = re.match(r"\s+a\d+: \[1\]5 \[4\](\S+)", l)
        if m:
            obj = m.group(1)
            continue
        m = re.match(r"\s+a\d+: \[4\](\d+) \[11\](.*)$", l)
        if m and obj:
            out[obj].append((int(m.group(1)), m.group(2).rstrip()))
    # 같은 줄이 여러 번 찍히므로 (줄번호, 내용) 로 중복 제거
    return {k: sorted(set(v)) for k, v in out.items()}


old, new = parse(sys.argv[1]), parse(sys.argv[2])
for obj in sorted(set(old) | set(new)):
    o = collections.Counter(t.strip() for _, t in old.get(obj, []))
    n = collections.Counter(t.strip() for _, t in new.get(obj, []))
    add, rem = n - o, o - n
    if not add and not rem:
        continue
    print("=" * 8, obj, "(old %d lines / new %d lines)" % (sum(o.values()), sum(n.values())))
    for ln, t in new.get(obj, []):
        if add[t.strip()] > 0:
            print("  + %5d  %s" % (ln, t))
            add[t.strip()] -= 1
    for ln, t in old.get(obj, []):
        if rem[t.strip()] > 0:
            print("  - %5d  %s" % (ln, t))
            rem[t.strip()] -= 1
